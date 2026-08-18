// @ts-check
/**
 * dsh-office-tools — model-facing tools for reading and generating common
 * office / document files.
 *
 * Read (office_read_*):    PDF, DOCX, XLSX/XLS/CSV, PPTX → text
 * Generate (office_write_*): DOCX, XLSX, PPTX, PDF, CSV
 *
 * Tool conventions follow the official DeepSeek Harness tool spec
 * (docs/user/develop/basic/tool.md and docs/cookbook/adding-a-tool.md):
 * - plain `defineTool` definitions registered on `ctx.tools`;
 * - typed `parameters` (nested object nodes declare `additionalProperties`);
 * - `execute` returns the canonical JSON value declared by `output.schema`;
 * - `output.render` owns the model-facing text;
 * - `presentCall` is a pure function of args only.
 *
 * Filesystem behavior mirrors the native tool-fs suite: paths resolve against
 * the calling session's workspace, reads emit `fs/observed`, and generated
 * files honor the sandbox policy (workspace-write containment + read-first
 * observation intents + one-shot escalation through the approval channel).
 */
import { defineTool } from '@deepseek-ai/dsh-tools'
import { realpathSync } from 'node:fs'
import { rename, stat, unlink, writeFile } from 'node:fs/promises'
import { dirname, sep } from 'node:path'
import { tmpdir } from 'node:os'
import { buildCsvBytes, buildDocxBytes, buildPdfBytes, buildPptxBytes, buildXlsxBytes } from './generators.js'
import { parseDocx, parsePdf, parsePptx, parseXlsx } from './parsers.js'

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

/** Default and max characters of extracted text returned by one read call. */
const READ_MAX_CHARS = 40000
const READ_MAX_CHARS_MAX = 200000
/** Byte cap for reading one binary document into memory. */
const READ_MAX_BYTES = 256 * 1024 * 1024
/** Default and max rows returned per sheet by `office_read_xlsx`. */
const XLSX_MAX_ROWS = 200
const XLSX_MAX_ROWS_MAX = 5000
/** Sandbox escalation vocabulary, mirroring @deepseek-ai/dsh-sandbox. */
const ESCALATION_MODES = ['workspace-write', 'danger-full-access']
const WIDER_MODES = {
  'read-only': ['workspace-write', 'danger-full-access'],
  'workspace-write': ['danger-full-access'],
}
const SANDOX_DENIAL_HINT = '\n[sandbox: escalation available — retry this exact operation once with sandbox_permissions (the narrowest wider mode that suffices) + justification; the approval prompt asks the user]'

// ---------------------------------------------------------------------------
// Small validators
// ---------------------------------------------------------------------------

function requiredPath(args) {
  if (typeof args.file_path !== 'string' || args.file_path.trim().length === 0) throw new Error('file_path must be a non-empty string')
  return args.file_path.trim()
}

function positiveInt(value, name, fallback, max) {
  if (value === undefined) return fallback
  if (!Number.isInteger(value) || value < 1) throw new Error(`${name} must be a positive integer`)
  if (max !== undefined && value > max) throw new Error(`${name} must be at most ${max}`)
  return value
}

function validateEscalationArgs(sandboxPermissions, justification) {
  if (sandboxPermissions !== undefined && justification === undefined) throw new Error('invalid escalation: sandbox_permissions requires a justification')
  if (justification !== undefined && sandboxPermissions === undefined) throw new Error('invalid escalation: justification is only valid together with sandbox_permissions')
  if (justification !== undefined && justification.trim().length === 0) throw new Error('invalid justification: expected a non-empty sentence')
}

// ---------------------------------------------------------------------------
// Path resolution (session workspace), mirroring dsh-tool-fs
// ---------------------------------------------------------------------------

const PARENT_PATH_SEGMENT = /(?:^|[\\/])\.\.(?:[\\/]|$)/

function canonicalPath(path) {
  try {
    return realpathSync.native(path)
  } catch {
    return path
  }
}

function sessionCwd(exec, requestedPath) {
  const cwd = exec.agent?.session.header.cwd
  if (cwd === undefined || (!PARENT_PATH_SEGMENT.test(cwd) && !PARENT_PATH_SEGMENT.test(requestedPath))) return cwd
  return canonicalPath(cwd)
}

function sessionResolveOptions(exec, requestedPath, policyWorkspaceRoot) {
  const cwd = policyWorkspaceRoot ?? sessionCwd(exec, requestedPath)
  return { ...(cwd !== undefined ? { cwd } : {}), signal: exec.signal }
}

/** Resolve a model-supplied path and require a regular file, like tool-fs. */
async function resolveRegularReadTarget(ctx, exec, requestedPath) {
  const target = await ctx.fs.resolve(requestedPath, sessionResolveOptions(exec, requestedPath))
  const info = await ctx.fs.stat(target, exec.signal)
  if (info === undefined) {
    ctx.emit('fs/observed', target, { kind: 'absent' }, exec)
    throw new Error(`cannot read "${target.displayPath}": not found`)
  }
  if (info.type !== 'file') throw new Error(`cannot read "${target.displayPath}": not a regular file`)
  return { target, info }
}

// ---------------------------------------------------------------------------
// Sandbox controller for generated files (mirrors dsh-fs-sandbox semantics)
// ---------------------------------------------------------------------------

const MISSING_CODES = new Set(['ENOENT', 'ENOTDIR'])

function isMissing(error) {
  return MISSING_CODES.has(error.code)
}

function comparablePath(path, caseSensitive) {
  return caseSensitive ? path : path.toLowerCase()
}

function isLexicallyUnder(path, root, caseSensitive) {
  const comparableTarget = comparablePath(path, caseSensitive)
  const comparableRoot = comparablePath(root, caseSensitive)
  if (comparableTarget === comparableRoot) return true
  const prefix = comparableRoot.endsWith(sep) ? comparableRoot : comparableRoot + sep
  return comparableTarget.startsWith(prefix)
}

async function statIfPresent(path) {
  try {
    return await stat(path, { bigint: true })
  } catch (error) {
    if (isMissing(error)) return undefined
    throw error
  }
}

function sameIdentity(left, right) {
  return left.dev === right.dev && left.ino === right.ino
}

async function isPathUnder(path, root, caseSensitive = process.platform !== 'win32') {
  if (isLexicallyUnder(path, root, caseSensitive)) return true
  const rootInfo = await statIfPresent(root)
  if (!rootInfo) return false
  let ancestor = path
  while (true) {
    const ancestorInfo = await statIfPresent(ancestor)
    if (ancestorInfo && sameIdentity(ancestorInfo, rootInfo)) return true
    const parent = dirname(ancestor)
    if (parent === ancestor) return false
    ancestor = parent
  }
}

/**
 * Per-call sandbox policy for generated files: resolves the standing policy,
 * handles one-shot escalation through the approval channel, and enforces
 * workspace-write containment with the same writable-root set as the native
 * filesystem fence (workspace root + host /tmp + os.tmpdir()).
 */
class OfficeSandbox {
  constructor(ctx) {
    this.ctx = ctx
    const defaultMode = ctx.fs.sandboxMode
    this.escalationModes = defaultMode === undefined ? [] : ESCALATION_MODES
    this.policyService = defaultMode === undefined ? undefined : ctx.get('sandboxPolicy')
    if (defaultMode !== undefined && this.policyService === undefined) {
      throw new Error('office-tools: the mounted filesystem confines but ctx.sandboxPolicy is missing')
    }
  }

  schemaFields() {
    return {
      sandbox_permissions: {
        type: 'string',
        enum: [...this.escalationModes],
        description: 'The wider sandbox mode this file operation needs. Only valid as a one-shot retry of an operation the sandbox just denied; requires justification and user approval.',
      },
      justification: {
        type: 'string',
        description: 'Required with sandbox_permissions: one sentence for the user explaining why this exact file operation needs the wider access.',
      },
    }
  }

  async resolvePolicy(toolName, args, exec) {
    validateEscalationArgs(args.sandbox_permissions, args.justification)
    const standing = this.policyService?.resolve({ ...(exec.agent !== undefined ? { session: exec.agent.session } : {}) })
    if (args.sandbox_permissions === undefined || args.justification === undefined) return standing
    if (this.escalationModes.length === 0) throw new Error('sandbox_permissions is not available in this composition (no sandboxing filesystem to escalate)')
    if (standing === undefined) throw new Error('sandbox_permissions requires a resolved sandbox policy, but none is available for this call')
    const requested = args.sandbox_permissions
    const wider = WIDER_MODES[standing.mode]
    if (wider === undefined || !wider.includes(requested)) {
      throw new Error(`sandbox escalation to "${requested}" is not strictly wider than this call's current "${standing.mode}" mode`)
    }
    const approver = this.ctx.get('approval')
    if (approver === undefined) throw new Error(`sandbox escalation to "${requested}" requires approval, but no approval service is composed`)
    if (exec.agent === undefined) throw new Error(`sandbox escalation to "${requested}" requires approval, but the call has no agent to route it through`)
    const outcome = await approver.request({
      agent: exec.agent,
      toolName,
      callId: exec.callId,
      reason: `escalate sandbox to ${requested}: ${args.justification}`,
      ...(exec.signal !== undefined ? { signal: exec.signal } : {}),
    })
    switch (outcome) {
      case 'allowed-once':
        return { ...standing, mode: requested }
      case 'rejected':
        throw new Error(`the user rejected escalating this operation to "${requested}"`)
      case 'cancelled':
        throw new Error(`approval for escalating to "${requested}" was cancelled`)
      case 'unavailable':
        throw new Error(`sandbox escalation to "${requested}" requires approval, but no approval channel is available`)
      default:
        throw new Error(`unexpected approval outcome: ${String(outcome)}`)
    }
  }

  /** Enforce the per-call policy against a resolved target; returns the fresh target to write. */
  async checkedTarget(target, policy) {
    const mode = policy.mode
    if (mode === 'danger-full-access') return target
    if (mode === 'read-only') throw new Error(`[sandbox: file access denied under read-only mode]${SANDOX_DENIAL_HINT}`)
    const fresh = await this.ctx.fs.resolve(target.displayPath)
    const roots = [...new Set([policy.workspaceRoot, '/tmp', tmpdir()].map(canonicalPath))].filter((root) => root !== undefined)
    for (const root of roots) {
      if (await isPathUnder(fresh.targetKey, root)) return fresh
    }
    throw new Error(`[sandbox: file access denied under workspace-write mode]${SANDOX_DENIAL_HINT}`)
  }
}

// ---------------------------------------------------------------------------
// Binary write path
// ---------------------------------------------------------------------------

/** Honor the read-first observation intent, mirroring dsh-fs-observation-policy. */
async function enforceWriteIntent(ctx, target, intent) {
  const info = await ctx.fs.stat(target)
  if (intent.kind === 'createIfAbsent') {
    if (info !== undefined) {
      throw new Error(`cannot write "${target.displayPath}": the file already exists and this session has not read it; read the file first`)
    }
    return
  }
  if (info === undefined) {
    throw new Error(`cannot write "${target.displayPath}": the file was removed since this session's last read; re-read to confirm`)
  }
  if (info.version !== intent.version) {
    throw new Error(`cannot write "${target.displayPath}": the file changed since this session's last read; re-read the file, then retry`)
  }
}

/**
 * Resolve + sandbox-check + intent-check the target, then atomically publish
 * the generated bytes (temp file + rename) and record the observation.
 */
async function commitBinary(ctx, sandbox, toolName, args, exec, bytes, format) {
  const policy = await sandbox.resolvePolicy(toolName, args, exec)
  const resolved = await ctx.fs.resolve(requiredPath(args), sessionResolveOptions(exec, args.file_path, policy?.workspaceRoot))
  const target = policy === undefined ? resolved : await sandbox.checkedTarget(resolved, policy)
  const intent = await ctx.waterfall('fs/write-intent', target, exec, () => undefined)
  if (intent !== undefined) await enforceWriteIntent(ctx, target, intent)
  const path = ctx.fs.processPath(target)
  const temp = `${path}.${process.pid}.${Date.now()}.office-tmp`
  try {
    await writeFile(temp, bytes, { flag: 'wx' })
    await rename(temp, path)
  } catch (error) {
    await unlink(temp).catch(() => {})
    throw new Error(`cannot write "${target.displayPath}": ${error instanceof Error ? error.message : String(error)}`, { cause: error })
  }
  const info = await ctx.fs.stat(target)
  if (info === undefined) throw new Error(`cannot write "${target.displayPath}": the file cannot be re-stat-ed after the write`)
  ctx.emit('fs/observed', target, { kind: 'present', version: info.version }, exec)
  return { path: target.displayPath, format, bytes: bytes.length }
}

// ---------------------------------------------------------------------------
// Presentation helpers
// ---------------------------------------------------------------------------

function readEnvelope(value, label) {
  return `<path>${value.path}</path>\n<type>${label}</type>\n<content>\n${value.text}\n</content>`
}

function readCall(args, title) {
  return { card: 'generic', title, kind: 'read', locations: [{ path: args.file_path }] }
}

function writeCall(args, format) {
  return { card: 'generic', title: `Generate ${format} ${args.file_path}`, kind: 'write', locations: [{ path: args.file_path }] }
}

function writeOutcomeSchema() {
  return {
    type: 'object',
    additionalProperties: false,
    properties: {
      path: { type: 'string', required: true },
      format: { type: 'string', required: true },
      bytes: { type: 'integer', required: true },
    },
  }
}

function writeOutcomeRender(_args, value) {
  return [{ type: 'text', text: `Generated ${value.format.toUpperCase()} file at ${value.path} (${value.bytes} bytes)` }]
}

// ---------------------------------------------------------------------------
// Read tools
// ---------------------------------------------------------------------------

function applyReadPdf(ctx) {
  ctx.tools.register(defineTool({
    name: 'office_read_pdf',
    description: 'Read a PDF file and return its extracted text, page by page (covers text layers; scanned images without OCR are not covered).',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to the PDF file, resolved by the filesystem backend.' },
      page_start: { type: 'number', description: '1-based first page to extract. Defaults to 1.' },
      page_end: { type: 'number', description: 'Last page to extract, inclusive. Defaults to the last page.' },
      max_chars: { type: 'number', description: `Maximum characters of extracted text to return. Defaults to ${READ_MAX_CHARS}.` },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          path: { type: 'string', required: true },
          totalPages: { type: 'integer', required: true },
          text: { type: 'string', required: true },
        },
      },
      render: (_args, value) => [{ type: 'text', text: readEnvelope(value, 'pdf') }],
    },
    isConcurrencySafe: () => true,
    async execute(args, exec) {
      const filePath = requiredPath(args)
      const pageStart = args.page_start === undefined ? undefined : positiveInt(args.page_start, 'page_start')
      const pageEnd = args.page_end === undefined ? undefined : positiveInt(args.page_end, 'page_end')
      if (pageStart !== undefined && pageEnd !== undefined && pageEnd < pageStart) throw new Error('page_end must be greater than or equal to page_start')
      const maxChars = positiveInt(args.max_chars, 'max_chars', READ_MAX_CHARS, READ_MAX_CHARS_MAX)
      const { target, info } = await resolveRegularReadTarget(ctx, exec, filePath)
      const bytes = await ctx.fs.readBytes(target, exec.signal, READ_MAX_BYTES)
      const parsed = await parsePdf(bytes, { pageStart, pageEnd, maxChars })
      ctx.emit('fs/observed', target, { kind: 'present', version: info.version }, exec)
      return { path: target.displayPath, totalPages: parsed.totalPages, text: parsed.text }
    },
    presentCall(args) {
      return readCall(args, `Read PDF ${args.file_path}`)
    },
  }))
}

function applyReadDocx(ctx) {
  ctx.tools.register(defineTool({
    name: 'office_read_docx',
    description: 'Read a Word document (.docx) and return its text content (headings, paragraphs, and table cell text; formatting is not preserved).',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to the .docx file, resolved by the filesystem backend.' },
      max_chars: { type: 'number', description: `Maximum characters of extracted text to return. Defaults to ${READ_MAX_CHARS}.` },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          path: { type: 'string', required: true },
          text: { type: 'string', required: true },
        },
      },
      render: (_args, value) => [{ type: 'text', text: readEnvelope(value, 'docx') }],
    },
    isConcurrencySafe: () => true,
    async execute(args, exec) {
      const filePath = requiredPath(args)
      const maxChars = positiveInt(args.max_chars, 'max_chars', READ_MAX_CHARS, READ_MAX_CHARS_MAX)
      const { target, info } = await resolveRegularReadTarget(ctx, exec, filePath)
      const bytes = await ctx.fs.readBytes(target, exec.signal, READ_MAX_BYTES)
      const text = await parseDocx(bytes, { maxChars })
      ctx.emit('fs/observed', target, { kind: 'present', version: info.version }, exec)
      return { path: target.displayPath, text }
    },
    presentCall(args) {
      return readCall(args, `Read DOCX ${args.file_path}`)
    },
  }))
}

function applyReadXlsx(ctx) {
  ctx.tools.register(defineTool({
    name: 'office_read_xlsx',
    description: 'Read an Excel workbook (.xlsx / .xls) or CSV file and return its cell values sheet by sheet, as tab-separated rows.',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to the spreadsheet file, resolved by the filesystem backend.' },
      sheet: { type: 'string', description: 'Sheet name or 1-based sheet index to read. Defaults to all sheets.' },
      range: { type: 'string', description: 'Cell range to read, e.g. "A1:C10". Defaults to the used range.' },
      max_rows: { type: 'number', description: `Maximum rows returned per sheet. Defaults to ${XLSX_MAX_ROWS}.` },
      max_chars: { type: 'number', description: `Maximum characters of extracted text to return. Defaults to ${READ_MAX_CHARS}.` },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          path: { type: 'string', required: true },
          text: { type: 'string', required: true },
        },
      },
      render: (_args, value) => [{ type: 'text', text: readEnvelope(value, 'spreadsheet') }],
    },
    isConcurrencySafe: () => true,
    async execute(args, exec) {
      const filePath = requiredPath(args)
      const maxRows = positiveInt(args.max_rows, 'max_rows', XLSX_MAX_ROWS, XLSX_MAX_ROWS_MAX)
      const maxChars = positiveInt(args.max_chars, 'max_chars', READ_MAX_CHARS, READ_MAX_CHARS_MAX)
      const { target, info } = await resolveRegularReadTarget(ctx, exec, filePath)
      const bytes = await ctx.fs.readBytes(target, exec.signal, READ_MAX_BYTES)
      const parsed = await parseXlsx(bytes, { sheet: args.sheet, range: args.range, maxRows, maxChars })
      ctx.emit('fs/observed', target, { kind: 'present', version: info.version }, exec)
      return { path: target.displayPath, text: parsed.text }
    },
    presentCall(args) {
      return readCall(args, `Read spreadsheet ${args.file_path}`)
    },
  }))
}

function applyReadPptx(ctx) {
  ctx.tools.register(defineTool({
    name: 'office_read_pptx',
    description: 'Read a PowerPoint file (.pptx) and return its text slide by slide (titles, body text, and table cell text; images and drawings are not covered).',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to the .pptx file, resolved by the filesystem backend.' },
      max_chars: { type: 'number', description: `Maximum characters of extracted text to return. Defaults to ${READ_MAX_CHARS}.` },
    },
    output: {
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          path: { type: 'string', required: true },
          slides: { type: 'integer', required: true },
          text: { type: 'string', required: true },
        },
      },
      render: (_args, value) => [{ type: 'text', text: readEnvelope(value, 'pptx') }],
    },
    isConcurrencySafe: () => true,
    async execute(args, exec) {
      const filePath = requiredPath(args)
      const maxChars = positiveInt(args.max_chars, 'max_chars', READ_MAX_CHARS, READ_MAX_CHARS_MAX)
      const { target, info } = await resolveRegularReadTarget(ctx, exec, filePath)
      const bytes = await ctx.fs.readBytes(target, exec.signal, READ_MAX_BYTES)
      const parsed = await parsePptx(bytes, { maxChars })
      ctx.emit('fs/observed', target, { kind: 'present', version: info.version }, exec)
      return { path: target.displayPath, slides: parsed.slides, text: parsed.text }
    },
    presentCall(args) {
      return readCall(args, `Read PPTX ${args.file_path}`)
    },
  }))
}

// ---------------------------------------------------------------------------
// Write tools
// ---------------------------------------------------------------------------

function applyWriteDocx(ctx, sandbox) {
  const escalation = sandbox.escalationModes.length > 0 ? sandbox.schemaFields() : {}
  ctx.tools.register(defineTool({
    name: 'office_write_docx',
    description: 'Generate a Word document (.docx) with an optional centered title, styled paragraphs, and captioned tables.',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to write the .docx file to, resolved by the filesystem backend.' },
      title: { type: 'string', description: 'Document title, rendered centered and large.' },
      paragraphs: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            text: { type: 'string', required: true, description: 'The paragraph text.' },
            style: { type: 'string', enum: ['normal', 'heading1', 'heading2', 'heading3', 'quote'], description: 'Paragraph style. Defaults to normal.' },
            bold: { type: 'boolean', description: 'Render the paragraph bold.' },
            italic: { type: 'boolean', description: 'Render the paragraph italic.' },
            align: { type: 'string', enum: ['left', 'center', 'right', 'justify'], description: 'Paragraph alignment. Defaults to left.' },
            color: { type: 'string', description: 'Text color as a hex value, e.g. "FF0000".' },
          },
        },
        description: 'Body paragraphs, in document order.',
      },
      tables: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            caption: { type: 'string', description: 'Optional table caption above the table.' },
            headers: { type: 'array', items: { type: 'string' }, description: 'Header row; rendered with a shaded background.' },
            rows: { type: 'array', items: { type: 'array', items: { type: 'string' } }, description: 'Table body rows.' },
          },
        },
        description: 'Tables to include, in document order.',
      },
      ...escalation,
    },
    output: { schema: writeOutcomeSchema(), render: writeOutcomeRender },
    async execute(args, exec) {
      const bytes = await buildDocxBytes({ title: args.title, paragraphs: args.paragraphs, tables: args.tables })
      return commitBinary(ctx, sandbox, 'office_write_docx', args, exec, bytes, 'docx')
    },
    presentCall(args) {
      return writeCall(args, 'DOCX')
    },
  }))
}

function applyWriteXlsx(ctx, sandbox) {
  const escalation = sandbox.escalationModes.length > 0 ? sandbox.schemaFields() : {}
  ctx.tools.register(defineTool({
    name: 'office_write_xlsx',
    description: 'Generate an Excel workbook (.xlsx) with one or more sheets of cell values. Cell values may be strings, numbers, booleans, or null.',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to write the .xlsx file to, resolved by the filesystem backend.' },
      sheets: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            name: { type: 'string', description: 'Sheet name, at most 31 characters. Defaults to Sheet1, Sheet2, ...' },
            rows: {
              type: 'array',
              items: {
                type: 'array',
                items: { oneOf: [{ type: 'string' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }] },
              },
              required: true,
              description: 'Cell values, row by row. Put the header row first when the sheet has one.',
            },
          },
        },
        required: true,
        description: 'Sheets to include, in workbook order.',
      },
      ...escalation,
    },
    output: { schema: writeOutcomeSchema(), render: writeOutcomeRender },
    async execute(args, exec) {
      const bytes = buildXlsxBytes({ sheets: args.sheets })
      return commitBinary(ctx, sandbox, 'office_write_xlsx', args, exec, bytes, 'xlsx')
    },
    presentCall(args) {
      return writeCall(args, 'XLSX')
    },
  }))
}

function applyWritePptx(ctx, sandbox) {
  const escalation = sandbox.escalationModes.length > 0 ? sandbox.schemaFields() : {}
  ctx.tools.register(defineTool({
    name: 'office_write_pptx',
    description: 'Generate a PowerPoint deck (.pptx) with slides that carry a title, subtitle, bullet list, and/or a table.',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to write the .pptx file to, resolved by the filesystem backend.' },
      slides: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            title: { type: 'string', description: 'Slide title.' },
            subtitle: { type: 'string', description: 'Slide subtitle under the title.' },
            bullets: {
              type: 'array',
              items: {
                oneOf: [
                  { type: 'string' },
                  {
                    type: 'object',
                    additionalProperties: false,
                    properties: {
                      text: { type: 'string', required: true },
                      level: { type: 'integer', description: 'Indentation level, 0 for top level.' },
                    },
                  },
                ],
              },
              description: 'Bullet points for this slide.',
            },
            table: {
              type: 'object',
              additionalProperties: false,
              properties: {
                headers: { type: 'array', items: { type: 'string' }, description: 'Table header row; rendered with a shaded background.' },
                rows: { type: 'array', items: { type: 'array', items: { type: 'string' } }, description: 'Table body rows.' },
              },
              description: 'Optional table rendered under the bullets.',
            },
          },
        },
        required: true,
        description: 'Slides to include, in deck order.',
      },
      ...escalation,
    },
    output: { schema: writeOutcomeSchema(), render: writeOutcomeRender },
    async execute(args, exec) {
      const bytes = await buildPptxBytes({ slides: args.slides })
      return commitBinary(ctx, sandbox, 'office_write_pptx', args, exec, bytes, 'pptx')
    },
    presentCall(args) {
      return writeCall(args, 'PPTX')
    },
  }))
}

function applyWritePdf(ctx, sandbox) {
  const escalation = sandbox.escalationModes.length > 0 ? sandbox.schemaFields() : {}
  ctx.tools.register(defineTool({
    name: 'office_write_pdf',
    description: 'Generate a PDF document with titled pages, paragraphs, bullets, and optional tables. Non-ASCII text (e.g. Chinese) renders via a Windows system font when available.',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to write the .pdf file to, resolved by the filesystem backend.' },
      page_size: { type: 'string', enum: ['A4', 'A3', 'A5', 'Letter', 'Legal', 'Tabloid'], description: 'Page size. Defaults to A4.' },
      orientation: { type: 'string', enum: ['portrait', 'landscape'], description: 'Page orientation. Defaults to portrait.' },
      pages: {
        type: 'array',
        items: {
          type: 'object',
          additionalProperties: false,
          properties: {
            title: { type: 'string', description: 'Page title, rendered centered at the top.' },
            paragraphs: { type: 'array', items: { type: 'string' }, description: 'Body paragraphs for this page.' },
            bullets: { type: 'array', items: { type: 'string' }, description: 'Bullet points for this page.' },
            table: {
              type: 'object',
              additionalProperties: false,
              properties: {
                headers: { type: 'array', items: { type: 'string' }, description: 'Table header row.' },
                rows: { type: 'array', items: { type: 'array', items: { type: 'string' } }, description: 'Table body rows.' },
              },
              description: 'Optional table for this page.',
            },
          },
        },
        required: true,
        description: 'Pages to include, in document order.',
      },
      ...escalation,
    },
    output: { schema: writeOutcomeSchema(), render: writeOutcomeRender },
    async execute(args, exec) {
      const bytes = await buildPdfBytes({ pages: args.pages }, { page_size: args.page_size, orientation: args.orientation })
      return commitBinary(ctx, sandbox, 'office_write_pdf', args, exec, bytes, 'pdf')
    },
    presentCall(args) {
      return writeCall(args, 'PDF')
    },
  }))
}

function applyWriteCsv(ctx, sandbox) {
  const escalation = sandbox.escalationModes.length > 0 ? sandbox.schemaFields() : {}
  ctx.tools.register(defineTool({
    name: 'office_write_csv',
    description: 'Generate a CSV file with an optional header row and a configurable delimiter. A UTF-8 BOM is included by default so Excel opens non-ASCII text correctly.',
    parameters: {
      file_path: { type: 'string', required: true, description: 'Path to write the .csv file to, resolved by the filesystem backend.' },
      headers: { type: 'array', items: { type: 'string' }, description: 'Optional header row, written first.' },
      rows: {
        type: 'array',
        items: {
          type: 'array',
          items: { oneOf: [{ type: 'string' }, { type: 'number' }, { type: 'boolean' }, { type: 'null' }] },
        },
        required: true,
        description: 'Data rows, cell by cell.',
      },
      delimiter: { type: 'string', enum: ['comma', 'tab', 'semicolon'], description: 'Field delimiter. Defaults to comma.' },
      bom: { type: 'boolean', description: 'Whether to include the UTF-8 BOM. Defaults to true.' },
      ...escalation,
    },
    output: { schema: writeOutcomeSchema(), render: writeOutcomeRender },
    async execute(args, exec) {
      const bytes = buildCsvBytes({ headers: args.headers, rows: args.rows }, { delimiter: args.delimiter, bom: args.bom })
      return commitBinary(ctx, sandbox, 'office_write_csv', args, exec, bytes, 'csv')
    },
    presentCall(args) {
      return writeCall(args, 'CSV')
    },
  }))
}

// ---------------------------------------------------------------------------
// Plugin entry
// ---------------------------------------------------------------------------

/** Cordis plugin name used by loader diagnostics. */
export const name = 'office-tools'

/** Services required by this plugin. */
export const inject = ['tools', 'fs', 'systemPrompt']

/**
 * Register the office read/generate tool suite and one prompt-guidance
 * section. Registrations are effects scoped to the plugin fiber: they are
 * removed automatically when the plugin is stopped, updated, or removed.
 */
export function apply(ctx) {
  const sandbox = new OfficeSandbox(ctx)
  applyReadPdf(ctx)
  applyReadDocx(ctx)
  applyReadXlsx(ctx)
  applyReadPptx(ctx)
  applyWriteDocx(ctx, sandbox)
  applyWriteXlsx(ctx, sandbox)
  applyWritePptx(ctx, sandbox)
  applyWritePdf(ctx, sandbox)
  applyWriteCsv(ctx, sandbox)
  ctx.systemPrompt.section({
    name: 'tool:office',
    order: 140,
    text: 'Use the office_* tools to read binary office/document files (PDF, DOCX, XLSX, PPTX, CSV — the native read tool only handles UTF-8 text) and to generate office files (DOCX, XLSX, PPTX, PDF, CSV).',
  })
}
