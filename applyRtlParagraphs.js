// applyRtlParagraphs.js
import pkg from "jszip";
const JSZip = pkg;

/**
 * מקבל Buffer של DOCX,
 * פותח את word/document.xml,
 * דואג שלכל פסקה יהיה:
 *   - <w:bidi/>
 *   - <w:jc w:val="left"/>  (יישום לפי מה ש-Word יוצר בפועל)
 * ומחזיר Buffer חדש.
 */
export async function applyRtlParagraphs(buffer) {
  console.log("applyRtlParagraphs: Starting…");

  const zip = await JSZip.loadAsync(buffer);
  const path = "word/document.xml";

  const xmlFile = zip.file(path);
  if (!xmlFile) {
    console.warn("applyRtlParagraphs: word/document.xml not found, returning original buffer.");
    return buffer;
  }

  let xml = await xmlFile.async("string");
  const originalLength = xml.length;
  console.log("applyRtlParagraphs: Loaded document.xml, length:", originalLength);

  // עוברים על כל <w:p ...>...</w:p>
  const fixedXml = xml.replace(
    /<w:p([^>]*)>([\s\S]*?)<\/w:p>/g,
    (fullMatch, attrs, inner) => {
       if (
      fullMatch.includes('<w:pStyle w:val="Header"') ||
      fullMatch.includes('<w:pStyle w:val="Footer"')
    ) {
      return fullMatch; // מחזירים כמו שהוא
    }
      const hasPPr = /<w:pPr[\s\S]*?<\/w:pPr>/.test(inner);

      if (hasPPr) {
        // מטפלים ב-pPr הראשון בלבד בתוך הפסקה
       const newInner = inner.replace(
  /<w:pPr([^>]*)>([\s\S]*?)<\/w:pPr>/,
  (pprMatch, pprAttrs, pprInner) => {
    let modified = pprInner;

    // 1. bidi – אם אין, נוסיף בתחילת הבלוק
    if (!pprMatch.includes("<w:bidi")) {
      modified = `<w:bidi/>${modified}`;
    }

    // 2. jc – מתחשבים במרכז
    if (/<w:jc\b[^>]*w:val="center"/.test(pprMatch)) {
      // יש center – לא נוגעים, רק bidi נשאר
    } else if (/<w:jc\b[^>]*>/.test(pprMatch)) {
      modified = modified.replace(
        /<w:jc\b[^>]*>/g,
        '<w:jc w:val="left"/>'
      );
    } else {
      modified = `<w:jc w:val="left"/>${modified}`;
    }

    return `<w:pPr${pprAttrs}>${modified}</w:pPr>`;
  }
);


        return `<w:p${attrs}>${newInner}</w:p>`;
      } else {
        // אין pPr בכלל – ניצור אחד חדש בתחילת הפסקה
        const rtlPr = `<w:pPr><w:bidi/><w:jc w:val="left"/></w:pPr>`;
        return `<w:p${attrs}>${rtlPr}${inner}</w:p>`;
      }
    }
  );

  zip.file(path, fixedXml);

  console.log(
    "applyRtlParagraphs: Done. New length:",
    fixedXml.length,
    "(diff:",
    fixedXml.length - originalLength,
    ")"
  );

  const newBuffer = await zip.generateAsync({ type: "nodebuffer" });
  console.log("applyRtlParagraphs: Buffer regenerated.");

  return newBuffer;
}