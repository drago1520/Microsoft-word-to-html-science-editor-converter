import { strFromU8, unzipSync } from 'fflate'

/** Compact-shape children come back as one element or an array of them when repeated; normalize to an array. */
function list(value: Bun.XML.Value | Bun.XML.Value[] | undefined): Bun.XML.Element[] {
  if (value === undefined) return []
  const arr = Array.isArray(value) ? value : [value]
  return arr.filter((v): v is Bun.XML.Element => typeof v === 'object')
}

/** styleId -> its own w:tblBorders, walking w:basedOn when a style has none of its own. */
function tableStyleBorders(
  zip: ReturnType<typeof unzipSync>
): (styleId: string | undefined) => Bun.XML.Element | undefined {
  const file = zip['word/styles.xml']
  const stylesById = new Map<string, Bun.XML.Element>()
  if (file) {
    const root = Bun.XML.parse(strFromU8(file))['w:styles'] as Bun.XML.Element
    for (const style of list(root['w:style'])) {
      const id = style['@w:styleId'] as string | undefined
      if (id) stylesById.set(id, style)
    }
  }
  return function resolve(styleId: string | undefined, seen = new Set<string>()): Bun.XML.Element | undefined {
    if (!styleId || seen.has(styleId)) return undefined
    seen.add(styleId)
    const style = stylesById.get(styleId)
    const borders = (style?.['w:tblPr'] as Bun.XML.Element | undefined)?.['w:tblBorders'] as Bun.XML.Element | undefined
    return (
      borders ?? resolve((style?.['w:basedOn'] as Bun.XML.Element | undefined)?.['@w:val'] as string | undefined, seen)
    )
  }
}

export function wordCellBorders(docxBytes: Uint8Array) {
  const zip = unzipSync(docxBytes)
  const xml = strFromU8(zip['word/document.xml']!)
  const resolveTableStyleBorders = tableStyleBorders(zip)
  const body = (Bun.XML.parse(xml)['w:document'] as Bun.XML.Element)['w:body'] as Bun.XML.Element
  const tables = list(body['w:tbl'])
  // ponytail: flat tables only; nested tables and vertical merges need structural mapping.
  const hasNested = tables.some(t =>
    list(t['w:tr']).some(row => list(row['w:tc']).some(cell => list(cell['w:tbl']).length > 0))
  )
  const hasVMerge = JSON.stringify(tables).includes('"w:vMerge"')
  if (hasVMerge || hasNested) return []

  return tables.flatMap(table => {
    const tblPr = table['w:tblPr'] as Bun.XML.Element | undefined
    // Explicit table borders override the named style's; only fall back to the style when none are set inline.
    const tableBorders =
      (tblPr?.['w:tblBorders'] as Bun.XML.Element | undefined) ??
      resolveTableStyleBorders((tblPr?.['w:tblStyle'] as Bun.XML.Element | undefined)?.['@w:val'] as string | undefined)
    const rows = list(table['w:tr'])
    return rows.flatMap((row, r) => {
      const cells = list(row['w:tc'])
      return cells.map((cell, c) => {
        const borders = (cell['w:tcPr'] as Bun.XML.Element | undefined)?.['w:tcBorders'] as Bun.XML.Element | undefined
        return ['top', 'right', 'bottom', 'left']
          .map(side => {
            const inherited =
              (side === 'top' && r > 0) || (side === 'bottom' && r < rows.length - 1)
                ? 'insideH'
                : (side === 'left' && c > 0) || (side === 'right' && c < cells.length - 1)
                  ? 'insideV'
                  : side
            const attrs =
              (borders?.[`w:${side}`] as Bun.XML.Element | undefined) ??
              (tableBorders?.[`w:${inherited}`] as Bun.XML.Element | undefined) ??
              {}
            const type = attrs['@w:val'] as string | undefined
            if (type === 'nil' || type === 'none') return `border-${side}: none;`
            const styles: Record<string, string> = {
              single: 'solid',
              double: 'double',
              dotted: 'dotted',
              dashed: 'dashed'
            }
            const style = styles[type ?? '']
            if (!style) return ''
            const size = Number(attrs['@w:sz'] ?? 8) / 8
            const color = /^[\da-fA-F]{6}$/.test((attrs['@w:color'] as string) ?? '') ? attrs['@w:color'] : '000000'
            return `border-${side}: ${size}pt ${style} #${color};`
          })
          .join(' ')
      })
    })
  })
}

/**
 * Pandoc's `+styles` keeps custom Word style *names* (data-custom-style) but
 * drops their formatting, so rebuild the missing CSS from the docx itself.
 */
export function wordStylesToCss(docxBytes: Uint8Array) {
  const xml = strFromU8(unzipSync(docxBytes)['word/styles.xml']!)
  const stylesRoot = Bun.XML.parse(xml)['w:styles'] as Bun.XML.Element
  const rules: string[] = []

  for (const style of list(stylesRoot['w:style'])) {
    // Built-in styles become real tags (h1, em, ...), so only custom ones matter.
    if (style['@w:customStyle'] !== '1') continue
    // Pandoc writes w:name, not w:styleId ("alert Char", not "alertChar").
    const name = (style['w:name'] as Bun.XML.Element | undefined)?.['@w:val'] as string | undefined
    const rPr = style['w:rPr'] as Bun.XML.Element | undefined
    if (!name || !rPr) continue

    const decls: string[] = []
    const color = (rPr['w:color'] as Bun.XML.Element | undefined)?.['@w:val'] as string | undefined
    const size = (rPr['w:sz'] as Bun.XML.Element | undefined)?.['@w:val'] as string | undefined
    const font = (rPr['w:rFonts'] as Bun.XML.Element | undefined)?.['@w:ascii'] as string | undefined
    if (color && /^[0-9A-Fa-f]{6}$/.test(color)) decls.push(`color: #${color}`)
    if (size) decls.push(`font-size: ${Number(size) / 2}pt`) // w:sz is half-points
    if (font) decls.push(`font-family: "${font}", serif`)
    if ('w:b' in rPr) decls.push('font-weight: bold')
    if ('w:i' in rPr) decls.push('font-style: italic')
    if (!decls.length) continue

    rules.push(`  [data-custom-style="${name}"] { ${decls.join('; ')}; }`)
  }
  return rules.join('\n')
}
