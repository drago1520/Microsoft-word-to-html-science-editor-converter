import { strFromU8, unzipSync } from "fflate";

/**
 * Pandoc's `+styles` keeps custom Word style *names* (data-custom-style) but
 * drops their formatting, so rebuild the missing CSS from the docx itself.
 */
export function wordStylesToCss(docxBytes: Uint8Array) {
  const xml = strFromU8(unzipSync(docxBytes)["word/styles.xml"]!);
  const rules: string[] = [];

  for (const [style] of xml.matchAll(/<w:style [^>]*>[\s\S]*?<\/w:style>/g)) {
    // Built-in styles become real tags (h1, em, ...), so only custom ones matter.
    if (!style.includes('w:customStyle="1"')) continue;
    // Pandoc writes w:name, not w:styleId ("alert Char", not "alertChar").
    const name = style.match(/<w:name w:val="([^"]*)"/)?.[1];
    const rPr = style.match(/<w:rPr>[\s\S]*?<\/w:rPr>/)?.[0] ?? "";
    if (!name) continue;

    const decls: string[] = [];
    const color = rPr.match(/<w:color w:val="([0-9A-Fa-f]{6})"/)?.[1];
    const size = rPr.match(/<w:sz w:val="(\d+)"/)?.[1];
    const font = rPr.match(/<w:rFonts [^>]*w:ascii="([^"]*)"/)?.[1];
    if (color) decls.push(`color: #${color}`);
    if (size) decls.push(`font-size: ${Number(size) / 2}pt`); // w:sz is half-points
    if (font) decls.push(`font-family: "${font}", serif`);
    if (rPr.includes("<w:b/>")) decls.push("font-weight: bold");
    if (rPr.includes("<w:i/>")) decls.push("font-style: italic");
    if (!decls.length) continue;

    rules.push(`  [data-custom-style="${name}"] { ${decls.join("; ")}; }`);
  }
  return rules.join("\n");
}