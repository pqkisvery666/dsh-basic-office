// @ts-check
/**
 * dsh-office-tools smoke test — run with `node test/smoke.mjs` after
 * `pnpm install` inside this package.
 *
 * Covers, without a running DSH host:
 * 1. every generator → parser round trip (docx / xlsx / pptx / pdf / csv);
 * 2. the full tool surface through `apply()` with a mock `ctx` (registration,
 *    argument validation, sandbox-free execution, `fs/observed` emission);
 * 3. tool-name isolation from the native DSH tool set.
 */
import assert from 'node:assert/strict'
import { existsSync, mkdirSync, readFileSync, realpathSync, rmSync, writeFileSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, join, resolve as resolvePath, isAbsolute } from 'node:path'
import { apply as applyPlugin, inject, name } from '../lib/index.js'
import { buildCsvBytes, buildDocxBytes, buildPdfBytes, buildPptxBytes, buildXlsxBytes } from '../lib/generators.js'
import { parseDocx, parsePdf, parsePptx, parseXlsx } from '../lib/parsers.js'

const workDir = join(tmpdir(), `dsh-office-smoke-${process.pid}`)
rmSync(workDir, { recursive: true, force: true })

/** Canonicalize a path: realpath the deepest existing ancestor, keep the tail. */
function canonicalize(abs) {
  let current = abs
  const tail = []
  while (!existsSync(current) && dirname(current) !== current) {
    tail.unshift(basename(current))
    current = dirname(current)
  }
  let real
  try {
    real = realpathSync(current)
  } catch {
    real = current
  }
  return join(real, ...tail)
}

/** A bare local-like `ctx.fs` (no sandboxing), backed by node:fs. */
function makeMockFs() {
  return {
    sandboxMode: undefined,
    async resolve(path, opts) {
      const abs = isAbsolute(path) ? path : resolvePath(opts?.cwd ?? process.cwd(), path)
      const canon = canonicalize(abs)
      return { targetKey: canon, displayPath: canon }
    },
    processPath(target) {
      return target.targetKey
    },
    async stat(target) {
      try {
        const info = await import('node:fs/promises').then((m) => m.stat(target.targetKey))
        return { version: `v:${info.mtimeMs}:${info.size}`, type: info.isDirectory() ? 'directory' : 'file', size: info.size }
      } catch (error) {
        if (error.code === 'ENOENT') return undefined
        throw error
      }
    },
    async readBytes(target, _signal, _maxBytes) {
      return new Uint8Array(await readFile(target.targetKey))
    },
    fileUrl() {
      throw new Error('not used in smoke test')
    },
    contains() {
      throw new Error('not used in smoke test')
    },
    lstat() {
      throw new Error('not used in smoke test')
    },
    readText() {
      throw new Error('not used in smoke test')
    },
    streamText() {
      throw new Error('not used in smoke test')
    },
    listDir() {
      throw new Error('not used in smoke test')
    },
    writeText() {
      throw new Error('not used in smoke test')
    },
    editText() {
      throw new Error('not used in smoke test')
    },
  }
}

/** A minimal tool-execution context. */
function makeExec() {
  return { signal: undefined, callId: 'smoke-call', agent: undefined }
}

// ---------------------------------------------------------------------------
// 1. Generator → parser round trips
// ---------------------------------------------------------------------------

const roundTripDir = join(workDir, 'roundtrip')
mkdirSync(roundTripDir, { recursive: true })

async function roundTripDocx() {
  const bytes = await buildDocxBytes({
    title: '冒烟测试报告',
    paragraphs: [
      { text: '第一章 概述', style: 'heading1' },
      { text: '这是正文，包含 hello world 与中文。', bold: true },
      { text: '引用行', style: 'quote' },
    ],
    tables: [{ caption: '表1 数据', headers: ['列A', '列B'], rows: [['1', '中文'], ['2', '3']] }],
  })
  assert.ok(bytes.length > 0, 'docx bytes non-empty')
  const text = await parseDocx(bytes, { maxChars: 100000 })
  for (const expected of ['冒烟测试报告', '第一章 概述', 'hello world', '表1 数据', '列A']) {
    assert.ok(text.includes(expected), `docx round trip missing: ${expected}`)
  }
}

async function roundTripXlsx() {
  const bytes = buildXlsxBytes({
    sheets: [
      { name: '数据', rows: [['名称', '数量'], ['苹果', 3], ['香蕉', 7.5]] },
      { name: '第二页', rows: [[true, null, 'x']] },
    ],
  })
  assert.ok(bytes.length > 0, 'xlsx bytes non-empty')
  const { text } = await parseXlsx(bytes, { maxRows: 200, maxChars: 100000 })
  for (const expected of ['数据', '苹果', '3', '第二页', 'TRUE']) {
    assert.ok(text.includes(expected), `xlsx round trip missing: ${expected}`)
  }
}

async function roundTripPptx() {
  const bytes = await buildPptxBytes({
    slides: [
      { title: '标题一', subtitle: '副标题', bullets: ['要点A', { text: '要点B', level: 1 }] },
      { title: '标题二', table: { headers: ['H1', 'H2'], rows: [['a', 'b'], ['c', 'd']] } },
    ],
  })
  assert.ok(bytes.length > 0, 'pptx bytes non-empty')
  const { slides, text } = await parsePptx(bytes, { maxChars: 100000 })
  assert.equal(slides, 2)
  for (const expected of ['标题一', '副标题', '要点A', '要点B', '标题二', 'H1', 'a']) {
    assert.ok(text.includes(expected), `pptx round trip missing: ${expected}`)
  }
}

async function roundTripPdf() {
  const bytes = await buildPdfBytes({
    pages: [
      { title: 'PDF标题', paragraphs: ['中文段落 Chinese paragraph 123', '第二段'], bullets: ['要点一', '要点二'] },
      { title: '第二页', table: { headers: ['A', 'B'], rows: [['1', '2']] } },
    ],
  }, { page_size: 'A4' })
  assert.ok(bytes.length > 0, 'pdf bytes non-empty')
  const { totalPages, text } = await parsePdf(bytes, { maxChars: 100000 })
  assert.ok(totalPages >= 2, `pdf has ${totalPages} pages`)
  for (const expected of ['PDF标题', '中文段落', 'Chinese paragraph', '要点一', '第二页', 'A', '1']) {
    assert.ok(text.includes(expected), `pdf round trip missing: ${expected}`)
  }
}

async function roundTripCsv() {
  const bytes = buildCsvBytes({ headers: ['城市', '人口'], rows: [['北京', 2189], ['上海', 2487]] })
  assert.ok(bytes.length > 0, 'csv bytes non-empty')
  const { text } = await parseXlsx(bytes, { maxRows: 200, maxChars: 100000 })
  for (const expected of ['城市', '人口', '北京', '2189']) {
    assert.ok(text.includes(expected), `csv round trip missing: ${expected}`)
  }
}

// ---------------------------------------------------------------------------
// 2. Full tool surface through apply() with a mock ctx
// ---------------------------------------------------------------------------

async function toolSurface() {
  const registered = []
  const sections = []
  const observed = []
  const ctx = {
    fs: makeMockFs(),
    tools: {
      register(definition) {
        registered.push(definition)
        return () => {}
      },
    },
    systemPrompt: {
      section(section) {
        sections.push(section)
      },
    },
    get() {
      return undefined
    },
    emit(event, target, observation) {
      observed.push({ event, target, observation })
    },
    async waterfall() {
      return undefined
    },
  }
  applyPlugin(ctx)

  assert.equal(name, 'office-tools')
  assert.deepEqual(inject, ['tools', 'fs', 'systemPrompt'])
  assert.ok(sections.some((s) => s.name === 'tool:office'), 'prompt section registered')

  const toolNames = registered.map((tool) => tool.name)
  const nativeNames = ['read', 'write', 'edit', 'read_image', 'bash', 'pwsh', 'web_search', 'glob', 'grep', 'skill']
  for (const native of nativeNames) assert.ok(!toolNames.includes(native), `collides with native tool: ${native}`)
  assert.equal(new Set(toolNames).size, toolNames.length, 'no duplicate tool names')

  const byName = Object.fromEntries(registered.map((tool) => [tool.name, tool]))
  const exec = makeExec()

  const filePath = (name) => join(roundTripDir, name)

  // office_write_docx → office_read_docx
  const docxOut = await byName.office_write_docx.execute({
    file_path: filePath('report.docx'),
    title: '工具测试文档',
    paragraphs: [{ text: '标题一', style: 'heading1' }, { text: '正文 hello' }],
    tables: [{ headers: ['X'], rows: [['1']] }],
  }, exec)
  assert.equal(docxOut.format, 'docx')
  assert.ok(docxOut.bytes > 0)
  assert.ok(existsSync(docxOut.path))
  assert.ok(observed.some((o) => o.event === 'fs/observed' && o.observation.kind === 'present'), 'write emits fs/observed')
  const docxRead = await byName.office_read_docx.execute({ file_path: filePath('report.docx') }, exec)
  assert.ok(docxRead.text.includes('工具测试文档'))

  // office_write_xlsx → office_read_xlsx
  await byName.office_write_xlsx.execute({
    file_path: filePath('data.xlsx'),
    sheets: [{ name: 'Sheet1', rows: [['名称', '值'], ['苹果', 3]] }],
  }, exec)
  const xlsxRead = await byName.office_read_xlsx.execute({ file_path: filePath('data.xlsx') }, exec)
  assert.ok(xlsxRead.text.includes('苹果') && xlsxRead.text.includes('3'))

  // office_write_pptx → office_read_pptx
  await byName.office_write_pptx.execute({
    file_path: filePath('deck.pptx'),
    slides: [{ title: '演示标题', bullets: ['要点1', '要点2'] }],
  }, exec)
  const pptxRead = await byName.office_read_pptx.execute({ file_path: filePath('deck.pptx') }, exec)
  assert.ok(pptxRead.text.includes('演示标题') && pptxRead.text.includes('要点1'))

  // office_write_pdf → office_read_pdf (with page range)
  await byName.office_write_pdf.execute({
    file_path: filePath('doc.pdf'),
    pages: [{ title: '第一页', paragraphs: ['第一页内容 alpha'] }, { title: '第二页', paragraphs: ['第二页内容 beta'] }],
  }, exec)
  const pdfRead = await byName.office_read_pdf.execute({ file_path: filePath('doc.pdf'), page_start: 2 }, exec)
  assert.ok(pdfRead.text.includes('第二页') && pdfRead.text.includes('beta'))
  assert.ok(!pdfRead.text.includes('alpha'), 'page range respected')

  // office_write_csv → office_read_xlsx (csv input)
  await byName.office_write_csv.execute({
    file_path: filePath('data.csv'),
    headers: ['A', 'B'],
    rows: [['x', 'y']],
    delimiter: 'tab',
  }, exec)
  const csvRead = await byName.office_read_xlsx.execute({ file_path: filePath('data.csv') }, exec)
  assert.ok(csvRead.text.includes('x') && csvRead.text.includes('y'))

  // missing file read fails cleanly
  await assert.rejects(() => byName.office_read_pdf.execute({ file_path: filePath('nope.pdf') }, exec), /not found/)

  // argument validation runs before execute (ToolArgsError)
  await assert.rejects(
    () => byName.office_write_xlsx.execute({ file_path: filePath('bad.xlsx'), sheets: [{ name: 'x' }] }, exec),
    /invalid arguments/,
  )

  // presentCall is a pure projection of args
  const call = byName.office_read_docx.presentCall({ file_path: 'a/b.docx' })
  assert.equal(call.card, 'generic')
  assert.equal(call.kind, 'read')
  assert.deepEqual(call.locations, [{ path: 'a/b.docx' }])

  // output.render produces the model-facing text
  const blocks = byName.office_read_pdf.output.render({}, { path: 'p.pdf', totalPages: 1, text: 'TEXT' })
  assert.equal(blocks[0].type, 'text')
  assert.ok(blocks[0].text.includes('TEXT'))
}

// ---------------------------------------------------------------------------

const rounds = [roundTripDocx, roundTripXlsx, roundTripPptx, roundTripPdf, roundTripCsv, toolSurface]
let failed = 0
for (const round of rounds) {
  try {
    await round()
    console.log(`ok - ${round.name}`)
  } catch (error) {
    failed += 1
    console.error(`FAIL - ${round.name}: ${error.message}`)
  }
}
rmSync(workDir, { recursive: true, force: true })
if (failed > 0) process.exit(1)
console.log('all smoke tests passed')
