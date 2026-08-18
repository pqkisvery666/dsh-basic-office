// @ts-check
/**
 * Generate-side pure functions for the office-tools plugin: build the bytes
 * of common binary document formats from a small, model-friendly content
 * structure.
 *
 * Supported outputs:
 * - DOCX via `docx` (titles, styled paragraphs, tables)
 * - XLSX via SheetJS (`xlsx`)
 * - PPTX via `pptxgenjs` (titles, subtitles, bullet lists, tables)
 * - PDF  via `pdfkit` (CJK text via a Windows system font when available)
 * - CSV  hand-rolled with RFC-4180 escaping and an optional UTF-8 BOM
 */
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'
import XLSX from 'xlsx'
import pptxgen from 'pptxgenjs'
import PDFDocument from 'pdfkit'
import { existsSync } from 'node:fs'

// ---------------------------------------------------------------------------
// DOCX
// ---------------------------------------------------------------------------

const ALIGNMENT = {
  left: AlignmentType.LEFT,
  center: AlignmentType.CENTER,
  right: AlignmentType.RIGHT,
  justify: AlignmentType.JUSTIFIED,
}

function tableBorders() {
  const spec = { style: BorderStyle.SINGLE, size: 4, color: 'A6A6A6' }
  return {
    top: spec,
    bottom: spec,
    left: spec,
    right: spec,
    insideHorizontal: spec,
    insideVertical: spec,
  }
}

/**
 * Build a .docx document: optional centered title, styled paragraphs
 * (`normal` / `heading1..3` / `quote`, plus bold/italic/align/color), and
 * captioned tables with a shaded header row.
 * @param {{ title?: string, paragraphs?: Array<{ text: string, style?: string, bold?: boolean, italic?: boolean, align?: string, color?: string } | string>, tables?: Array<{ caption?: string, headers?: string[], rows?: (string | number)[][] }> }} content
 * @returns {Promise<Buffer>}
 */
export async function buildDocxBytes(content) {
  const children = []
  if (content.title !== undefined && content.title !== '') {
    children.push(new Paragraph({ text: String(content.title), heading: HeadingLevel.TITLE, alignment: AlignmentType.CENTER }))
  }
  for (const item of content.paragraphs ?? []) {
    const paragraph = typeof item === 'string' ? { text: item } : item
    const text = String(paragraph.text ?? '')
    switch (paragraph.style) {
      case 'heading1':
        children.push(new Paragraph({ text, heading: HeadingLevel.HEADING_1 }))
        break
      case 'heading2':
        children.push(new Paragraph({ text, heading: HeadingLevel.HEADING_2 }))
        break
      case 'heading3':
        children.push(new Paragraph({ text, heading: HeadingLevel.HEADING_3 }))
        break
      case 'quote':
        children.push(new Paragraph({
          children: [new TextRun({ text, italics: true, color: '595959' })],
          indent: { left: 720 },
        }))
        break
      default:
        children.push(new Paragraph({
          children: [new TextRun({
            text,
            bold: paragraph.bold === true,
            italics: paragraph.italic === true,
            ...(paragraph.color !== undefined ? { color: paragraph.color } : {}),
          })],
          ...(paragraph.align !== undefined && ALIGNMENT[paragraph.align] !== undefined ? { alignment: ALIGNMENT[paragraph.align] } : {}),
        }))
    }
  }
  for (const table of content.tables ?? []) {
    if (table.caption !== undefined && table.caption !== '') {
      children.push(new Paragraph({ text: String(table.caption), heading: HeadingLevel.HEADING_4 }))
    }
    const headerCells = (table.headers ?? []).map((header) => new TableCell({
      children: [new Paragraph({
        children: [new TextRun({ text: String(header), bold: true, color: 'FFFFFF' })],
        alignment: AlignmentType.CENTER,
      })],
      shading: { fill: '4472C4' },
    }))
    const rows = (table.rows ?? []).map((row) => new TableRow({
      children: (row ?? []).map((cell) => new TableCell({ children: [new Paragraph(String(cell ?? ''))] })),
    }))
    if (headerCells.length > 0) rows.unshift(new TableRow({ children: headerCells }))
    if (rows.length > 0) {
      children.push(new Table({
        rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: tableBorders(),
      }))
    }
  }
  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: 'Microsoft YaHei', size: 22 },
        },
      },
    },
    sections: [{ children }],
  })
  return Buffer.from(await Packer.toBuffer(doc))
}

// ---------------------------------------------------------------------------
// XLSX / CSV
// ---------------------------------------------------------------------------

/**
 * Build an .xlsx workbook with multiple sheets, auto-computed column widths,
 * and an optional header row (the first row of `rows`).
 * @param {{ sheets: Array<{ name?: string, rows: Array<Array<string | number | boolean | null>> }> }} content
 * @returns {Buffer}
 */
export function buildXlsxBytes(content) {
  const sheets = content.sheets ?? []
  if (sheets.length === 0) throw new Error('at least one sheet with rows is required')
  const workbook = XLSX.utils.book_new()
  sheets.forEach((sheet, index) => {
    const rows = (sheet.rows ?? []).map((row) => (row ?? []).map((cell) => (cell === undefined ? null : cell)))
    const ws = XLSX.utils.aoa_to_sheet(rows)
    const widths = []
    for (const row of rows) {
      for (let column = 0; column < row.length; column++) {
        const length = String(row[column] ?? '').length
        widths[column] = Math.max(widths[column] ?? 0, length)
      }
    }
    ws['!cols'] = widths.map((width) => ({ wch: Math.min(Math.max(width + 2, 8), 50) }))
    XLSX.utils.book_append_sheet(workbook, ws, (sheet.name || `Sheet${index + 1}`).slice(0, 31))
  })
  return XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })
}

const CSV_DELIMITERS = { comma: ',', tab: '\t', semicolon: ';' }

/** RFC-4180 single-cell escaping. */
function csvEscape(value) {
  const text = value === null || value === undefined ? '' : String(value)
  return /[",\r\n]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text
}

/**
 * Build a .csv file: optional header row, `rows` of cell values, a chosen
 * delimiter, and an optional UTF-8 BOM (on by default so Excel opens
 * non-ASCII text correctly).
 * @param {{ headers?: string[], rows: Array<Array<string | number | boolean | null>> }} content
 * @param {{ delimiter?: string, bom?: boolean }} [options]
 * @returns {Buffer}
 */
export function buildCsvBytes(content, options = {}) {
  const delimiter = CSV_DELIMITERS[options.delimiter ?? 'comma']
  if (delimiter === undefined) throw new Error(`unsupported delimiter "${options.delimiter}" (supported: comma, tab, semicolon)`)
  const rows = [...(content.headers !== undefined && content.headers.length > 0 ? [content.headers] : []), ...(content.rows ?? [])]
  if (rows.length === 0) throw new Error('at least one row is required')
  const text = rows.map((row) => (row ?? []).map(csvEscape).join(delimiter)).join('\r\n') + '\r\n'
  const bom = options.bom !== false
  return Buffer.from((bom ? '\uFEFF' : '') + text, 'utf8')
}

// ---------------------------------------------------------------------------
// PPTX
// ---------------------------------------------------------------------------

const PPTX_FONT = 'Microsoft YaHei'

/**
 * Build a .pptx deck: each slide may carry a title, a subtitle, a bullet
 * list (items may declare `level` for indentation), and a table with an
 * optional shaded header row.
 * @param {{ slides: Array<{ title?: string, subtitle?: string, bullets?: Array<{ text: string, level?: number } | string>, table?: { headers?: string[], rows?: (string | number)[][] } }> }} content
 * @returns {Promise<Buffer>}
 */
export async function buildPptxBytes(content) {
  const slides = content.slides ?? []
  if (slides.length === 0) throw new Error('at least one slide is required')
  const pptx = new pptxgen()
  pptx.layout = 'LAYOUT_WIDE'
  for (const def of slides) {
    const slide = pptx.addSlide()
    slide.background = { color: 'FFFFFF' }
    if (def.title !== undefined && def.title !== '') {
      slide.addText(String(def.title), {
        x: 0.6, y: 0.35, w: 12.1, h: 0.85, fontSize: 30, bold: true, color: '1F3864', fontFace: PPTX_FONT,
      })
    }
    if (def.subtitle !== undefined && def.subtitle !== '') {
      slide.addText(String(def.subtitle), {
        x: 0.6, y: 1.15, w: 12.1, h: 0.55, fontSize: 16, color: '595959', fontFace: PPTX_FONT,
      })
    }
    let y = def.subtitle !== undefined && def.subtitle !== '' ? 1.8 : def.title !== undefined && def.title !== '' ? 1.3 : 0.4
    if (def.bullets !== undefined && def.bullets.length > 0) {
      const items = def.bullets.map((bullet) => (
        typeof bullet === 'string'
          ? { text: bullet, options: {} }
          : { text: String(bullet.text ?? ''), options: bullet.level !== undefined ? { indentLevel: bullet.level } : {} }
      ))
      slide.addText(items, {
        x: 0.6, y, w: 12.1, h: Math.max(5.7 - (y - 1.3), 1.5), fontSize: 18, color: '262626', fontFace: PPTX_FONT,
        bullet: { characterCode: '2022', indent: 18 }, paraSpaceAfter: 10, breakLine: false,
      })
      y += Math.min(items.length * 0.55 + 0.3, 5.4)
    }
    if (def.table !== undefined) {
      const table = def.table
      const rows = [...(table.headers !== undefined && table.headers.length > 0 ? [table.headers] : []), ...(table.rows ?? [])]
        .map((row) => (row ?? []).map((cell) => String(cell ?? '')))
      const columns = Math.max(...rows.map((row) => row.length), 1)
      if (rows.length > 0) {
        const tableCells = rows.map((row, rowIndex) => row.map((cell) => ({
          text: cell,
          ...(rowIndex === 0 && table.headers !== undefined && table.headers.length > 0
            ? { options: { fill: { color: '4472C4' }, color: 'FFFFFF', bold: true } }
            : {}),
        })))
        slide.addTable(tableCells, {
          x: 0.6, y, w: 12.1, fontSize: 12, color: '262626', fontFace: PPTX_FONT,
          border: { type: 'solid', pt: 1, color: 'BFBFBF' },
          fill: { color: 'F7FAFF' },
          colW: Array.from({ length: columns }, () => 12.1 / columns),
        })
      }
    }
  }
  const out = await pptx.write({ outputType: 'nodebuffer' })
  return Buffer.from(out)
}

// ---------------------------------------------------------------------------
// PDF
// ---------------------------------------------------------------------------

const PDF_SIZES = new Set(['A4', 'A3', 'A5', 'Letter', 'Legal', 'Tabloid'])

/** Windows CJK system fonts, tried in order, for non-ASCII PDF text. */
const CJK_FONT_CANDIDATES = [
  { path: 'C:\\Windows\\Fonts\\msyh.ttc', family: 'MicrosoftYaHei' },
  { path: 'C:\\Windows\\Fonts\\simhei.ttf', family: 'SimHei' },
  { path: 'C:\\Windows\\Fonts\\simsun.ttc', family: 'SimSun' },
]

/**
 * Register the first available Windows CJK font on the document. Returns
 * whether a CJK font is active; without one, text renders in the standard
 * Helvetica (ASCII only).
 */
function registerCjkFont(doc) {
  for (const candidate of CJK_FONT_CANDIDATES) {
    try {
      if (!existsSync(candidate.path)) continue
      doc.registerFont('cjk', candidate.path, candidate.family)
      return true
    } catch {
      // try the next candidate
    }
  }
  return false
}

/** Draw one table with wrapped cells and row-spanning page breaks. */
function renderPdfTable(doc, table, cjk) {
  const headers = table.headers ?? []
  const rows = table.rows ?? []
  const all = [...(headers.length > 0 ? [headers] : []), ...rows]
  if (all.length === 0) return
  const columnCount = Math.max(...all.map((row) => row.length), 1)
  const margin = doc.page.margins.left
  const usable = doc.page.width - margin - doc.page.margins.right
  const columnWidth = usable / columnCount
  const padding = 6
  const fontSize = 9
  const font = cjk ? 'cjk' : 'Helvetica'
  const drawRow = (row, header) => {
    const cells = Array.from({ length: columnCount }, (_, index) => String(row[index] ?? ''))
    const heights = cells.map((cell) => doc.heightOfString(cell, { width: columnWidth - padding * 2, fontSize }))
    const height = Math.max(...heights, 14) + padding * 2
    if (doc.y + height > doc.page.height - doc.page.margins.bottom) doc.addPage()
    const y0 = doc.y
    for (let index = 0; index < columnCount; index++) {
      const x = margin + index * columnWidth
      doc.rect(x, y0, columnWidth, height).fillAndStroke(header ? 'E8EEF7' : 'FFFFFF', 'B8B8B8').lineWidth(0.5)
      doc.font(font).fontSize(fontSize).fillColor('1A1A1A').text(cells[index], x + padding, y0 + padding, {
        width: columnWidth - padding * 2,
        height: height - padding * 2,
        ellipsis: true,
        lineBreak: true,
      })
    }
    doc.y = y0 + height
  }
  if (headers.length > 0) drawRow(headers, true)
  for (const row of rows) drawRow(row, false)
  doc.moveDown(0.3)
}

/**
 * Build a PDF document: `pages` of titled sections with paragraphs, bullets,
 * and an optional table each. Non-ASCII text uses a Windows CJK font when
 * one is present; otherwise only ASCII renders correctly.
 * @param {{ pages: Array<{ title?: string, paragraphs?: string[], bullets?: string[], table?: { headers?: string[], rows?: (string | number)[][] } }> }} content
 * @param {{ page_size?: string, orientation?: string }} [options]
 * @returns {Promise<Buffer>}
 */
export function buildPdfBytes(content, options = {}) {
  const size = options.page_size ?? 'A4'
  if (!PDF_SIZES.has(size)) throw new Error(`unsupported page_size "${size}" (supported: ${[...PDF_SIZES].join(', ')})`)
  const layout = options.orientation === 'landscape' ? 'landscape' : 'portrait'
  const pages = content.pages ?? []
  if (pages.length === 0) throw new Error('at least one page is required')
  return new Promise((resolve, reject) => {
    try {
      const doc = new PDFDocument({ size, layout, margin: 50 })
      const chunks = []
      doc.on('data', (chunk) => chunks.push(chunk))
      doc.on('end', () => resolve(Buffer.concat(chunks)))
      doc.on('error', reject)
      const cjk = registerCjkFont(doc)
      const body = cjk ? 'cjk' : 'Helvetica'
      pages.forEach((page, index) => {
        if (index > 0) doc.addPage()
        if (page.title !== undefined && page.title !== '') {
          doc.font(cjk ? 'cjk' : 'Helvetica-Bold').fontSize(20).fillColor('1A1A1A').text(String(page.title), { align: 'center' })
          doc.moveDown(0.5)
        }
        doc.font(body).fontSize(11)
        for (const paragraph of page.paragraphs ?? []) {
          doc.text(String(paragraph), { align: 'justify', lineGap: 2 })
          doc.moveDown(0.35)
        }
        for (const bullet of page.bullets ?? []) {
          doc.text(`\u2022  ${String(bullet)}`, { indent: 12, lineGap: 2 })
          doc.moveDown(0.2)
        }
        if (page.table !== undefined) renderPdfTable(doc, page.table, cjk)
      })
      doc.end()
    } catch (error) {
      reject(error)
    }
  })
}
