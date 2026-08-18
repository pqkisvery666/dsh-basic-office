// @ts-check
/**
 * Parse-side pure functions for the office-tools plugin: turn raw bytes of
 * common binary document formats into bounded, model-facing text.
 *
 * Supported inputs:
 * - PDF             via `unpdf` (pdfjs-dist wrapper, Node-safe)
 * - DOCX            via `mammoth`
 * - XLSX / XLS / CSV via SheetJS (`xlsx`)
 * - PPTX            via `jszip` + slide XML text extraction
 */
import { extractText, getDocumentProxy } from 'unpdf'
import mammoth from 'mammoth'
import XLSX from 'xlsx'
import JSZip from 'jszip'

/** Safety cap on the number of pages extracted in one PDF call. */
export const MAX_EXTRACT_PAGES = 1000

/**
 * Bound a text output to `maxChars`, appending an explicit truncation footer
 * so the model never mistakes a truncated result for the full document.
 * @param {string} text - the extracted text.
 * @param {number} maxChars - inclusive character cap.
 * @returns {string} the bounded text.
 */
export function capChars(text, maxChars) {
  if (typeof text !== 'string' || text.length <= maxChars) return text
  return `${text.slice(0, maxChars)}\n\n... (output truncated at ${maxChars} chars; narrow the range or raise max_chars to see more)`
}

/** @param {Uint8Array | Buffer} bytes */
function toBuffer(bytes) {
  return Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes)
}

/** unpdf requires a plain Uint8Array; Buffers are rejected explicitly. */
function toUint8(bytes) {
  return Buffer.isBuffer(bytes) ? new Uint8Array(bytes) : bytes
}

/**
 * Extract per-page text from a PDF, honoring a 1-based page range and a
 * character cap. Throws a model-facing error when the bytes are not a PDF.
 * @param {Uint8Array | ArrayBuffer} bytes - the PDF bytes.
 * @param {{ pageStart?: number, pageEnd?: number, maxChars: number }} options
 * @returns {Promise<{ totalPages: number, text: string }>}
 */
export async function parsePdf(bytes, { pageStart, pageEnd, maxChars }) {
  let proxy
  try {
    proxy = await getDocumentProxy(toUint8(bytes))
  } catch (error) {
    throw new Error(`cannot read as PDF: ${error instanceof Error ? error.message : String(error)}`, { cause: error })
  }
  const totalPages = proxy.numPages
  const start = pageStart === undefined ? 1 : pageStart
  const end = pageEnd === undefined ? totalPages : Math.min(pageEnd, totalPages)
  if (start > totalPages) throw new Error(`page_start ${start} is out of range (the PDF has ${totalPages} page(s))`)
  if (end - start + 1 > MAX_EXTRACT_PAGES) throw new Error(`page range ${start}-${end} spans ${end - start + 1} page(s); the limit is ${MAX_EXTRACT_PAGES}`)
  let pageTexts
  try {
    const result = await extractText(proxy, { mergePages: false })
    pageTexts = Array.isArray(result.text) ? result.text : [result.text]
  } catch (error) {
    throw new Error(`cannot extract PDF text: ${error instanceof Error ? error.message : String(error)}`, { cause: error })
  }
  const parts = []
  let chars = 0
  for (let page = start; page <= end; page++) {
    const body = `--- Page ${page} of ${totalPages} ---\n${(pageTexts[page - 1] ?? '').trim()}`
    if (chars + body.length > maxChars) {
      parts.push(capChars(body, maxChars - chars))
      break
    }
    parts.push(body)
    chars += body.length
  }
  return { totalPages, text: parts.join('\n\n') }
}

/**
 * Extract plain text from a .docx file.
 * @param {Uint8Array | Buffer} bytes - the DOCX bytes.
 * @param {{ maxChars: number }} options
 * @returns {Promise<string>}
 */
export async function parseDocx(bytes, { maxChars }) {
  let value
  try {
    const result = await mammoth.extractRawText({ buffer: toBuffer(bytes) })
    value = result.value
  } catch (error) {
    throw new Error(`cannot read as DOCX: ${error instanceof Error ? error.message : String(error)}`, { cause: error })
  }
  return capChars(value.trim(), maxChars)
}

/**
 * Read a spreadsheet (.xlsx / .xls / .csv) sheet by sheet as tab-separated
 * rows of formatted cell values.
 * @param {Uint8Array | Buffer} bytes - the workbook bytes.
 * @param {{ sheet?: string, range?: string, maxRows: number, maxChars: number }} options
 * @returns {Promise<{ text: string }>}
 */
export async function parseXlsx(bytes, { sheet, range, maxRows, maxChars }) {
  let workbook
  try {
    workbook = XLSX.read(toBuffer(bytes), { type: 'buffer', cellDates: true })
  } catch (error) {
    throw new Error(`cannot read as spreadsheet: ${error instanceof Error ? error.message : String(error)}`, { cause: error })
  }
  let names = workbook.SheetNames
  if (sheet !== undefined) {
    const numeric = /^\d+$/.test(String(sheet)) ? Number(sheet) : undefined
    if (numeric !== undefined) {
      if (numeric < 1 || numeric > names.length) {
        throw new Error(`sheet index ${numeric} is out of range; the workbook has ${names.length} sheet(s): ${names.join(', ')}`)
      }
      names = [names[numeric - 1]]
    } else if (!names.includes(String(sheet))) {
      throw new Error(`sheet "${sheet}" not found; available sheets: ${names.join(', ')}`)
    } else {
      names = [String(sheet)]
    }
  }
  const blocks = []
  let chars = 0
  for (const name of names) {
    const ws = workbook.Sheets[name]
    const all = ws && ws['!ref']
      ? XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '', ...(range !== undefined ? { range } : {}) })
      : []
    const limited = all.slice(0, maxRows)
    const lines = [
      `### Sheet: ${name} (${all.length} row(s))`,
      ...limited.map((row, index) => `${index + 1}:\t${(row ?? []).map((cell) => cell === null || cell === undefined ? '' : String(cell)).join('\t')}`),
    ]
    if (all.length > limited.length) lines.push(`... (${all.length - limited.length} more row(s); raise max_rows to see them)`)
    const block = lines.join('\n')
    if (chars + block.length > maxChars) {
      blocks.push(capChars(block, maxChars - chars))
      break
    }
    blocks.push(block)
    chars += block.length
  }
  return { text: blocks.join('\n\n') }
}

/** Decode the XML entity escapes found inside a PPTX text run. */
function decodeXmlEntities(text) {
  return text
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCodePoint(parseInt(dec, 10)))
    .replace(/&amp;/g, '&')
}

/** Collect the text of every `<a:t>` run in one slide XML document, in order. */
function extractSlideText(xml) {
  const texts = []
  const re = /<a:t(?:\s[^>]*)?>([\s\S]*?)<\/a:t>/g
  let match
  while ((match = re.exec(xml)) !== null) {
    const text = decodeXmlEntities(match[1]).replace(/\s+/g, ' ').trim()
    if (text.length > 0) texts.push(text)
  }
  return texts
}

/**
 * Extract per-slide text from a .pptx file (covers body text, titles, and
 * table cell text; not images or drawings).
 * @param {Uint8Array | Buffer} bytes - the PPTX bytes.
 * @param {{ maxChars: number }} options
 * @returns {Promise<{ slides: number, text: string }>}
 */
export async function parsePptx(bytes, { maxChars }) {
  let zip
  try {
    zip = await JSZip.loadAsync(bytes)
  } catch (error) {
    throw new Error(`cannot read as PPTX: ${error instanceof Error ? error.message : String(error)}`, { cause: error })
  }
  const slides = Object.keys(zip.files)
    .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .map((name) => ({ num: Number(/slide(\d+)\.xml$/.exec(name)[1]), name }))
    .sort((a, b) => a.num - b.num)
  const parts = []
  let chars = 0
  for (const slide of slides) {
    let xml
    try {
      xml = await zip.files[slide.name].async('string')
    } catch (error) {
      throw new Error(`cannot read slide ${slide.num}: ${error instanceof Error ? error.message : String(error)}`, { cause: error })
    }
    const body = `--- Slide ${slide.num} ---\n${extractSlideText(xml).join('\n')}`
    if (chars + body.length > maxChars) {
      parts.push(capChars(body, maxChars - chars))
      break
    }
    parts.push(body)
    chars += body.length
  }
  return { slides: slides.length, text: parts.join('\n\n') }
}
