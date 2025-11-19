// applyRtlSettings.js
import pkg from "jszip";
const JSZip = pkg;

/**
 * מקבל DOCX כ-Buffer, פותח את ה-ZIP,
 * מטפל ב:
 *  - word/settings.xml   → הגדרות RTL למסמך
 *  - word/styles.xml     → סגנון Normal וברירת מחדל RTL
 *  - word/numbering.xml  → רשימות/מספור RTL
 * ומחזיר Buffer חדש.
 */
export async function applyRtlSettings(buffer) {
  console.log("applyRtlSettings: Starting…");

  const zip = await JSZip.loadAsync(buffer);

  await ensureRtlInSettings(zip);
  await ensureRtlInStyles(zip);
  await ensureRtlInNumbering(zip);

  const newBuffer = await zip.generateAsync({ type: "nodebuffer" });
  console.log("applyRtlSettings: Done, buffer regenerated.");

  return newBuffer;
}

/* ------------------------------------------------------------------
 * 1. SETTINGS.XML – מסמך כ־RTL (bidi, he-IL, rtlGutter וכו')
 * ------------------------------------------------------------------ */

async function ensureRtlInSettings(zip) {
  const path = "word/settings.xml";
  const xmlFile = zip.file(path);

  if (!xmlFile) {
    console.warn("ensureRtlInSettings: settings.xml not found, skipping.");
    return;
  }

  let xml = await xmlFile.async("string");
  const originalLength = xml.length;

  console.log("ensureRtlInSettings: Loaded settings.xml, length:", originalLength);

  // ודא שיש <w:bidi/>
  if (!xml.includes("<w:bidi")) {
    console.log("ensureRtlInSettings: Inserting <w:bidi/> + RTL tags at end of <w:settings>.");
    xml = xml.replace(
      /<\/w:settings>/,
      `  <w:bidi/>\n  <w:rtlGutter/>\n  <w:themeFontLang w:val="he-IL"/>\n  <w:overrideTableDirection w:val="rightToLeft"/>\n</w:settings>`
    );
  } else {
    // לוודא rtlGutter
    if (!xml.includes("<w:rtlGutter")) {
      console.log("ensureRtlInSettings: Inserting <w:rtlGutter/> after <w:bidi/>");
      xml = xml.replace("<w:bidi/>", "<w:bidi/>\n  <w:rtlGutter/>");
    }
    // לוודא themeFontLang
    if (!xml.includes("w:themeFontLang")) {
      console.log("ensureRtlInSettings: Inserting <w:themeFontLang w:val=\"he-IL\"/> after <w:bidi/>");
      xml = xml.replace(
        "<w:bidi/>",
        '<w:bidi/>\n  <w:themeFontLang w:val="he-IL"/>'
      );
    }
    // לוודא overrideTableDirection
    if (!xml.includes("w:overrideTableDirection")) {
      console.log(
        "ensureRtlInSettings: Inserting <w:overrideTableDirection w:val=\"rightToLeft\"/> after <w:bidi/>"
      );
      xml = xml.replace(
        "<w:bidi/>",
        '<w:bidi/>\n  <w:overrideTableDirection w:val="rightToLeft"/>'
      );
    }
  }

  zip.file(path, xml);

  console.log(
    "ensureRtlInSettings: Done. New length:",
    xml.length,
    "(diff:",
    xml.length - originalLength,
    ")"
  );
}

/* ------------------------------------------------------------------
 * 2. STYLES.XML – Normal ו־docDefaults כ־RTL + יישור לימין
 * ------------------------------------------------------------------ */

async function ensureRtlInStyles(zip) {
  const path = "word/styles.xml";
  const xmlFile = zip.file(path);

  if (!xmlFile) {
    console.warn("ensureRtlInStyles: styles.xml not found, skipping.");
    return;
  }

  let xml = await xmlFile.async("string");
  const originalLength = xml.length;

  console.log("ensureRtlInStyles: Loaded styles.xml, length:", originalLength);

  // 2.1 docDefaults – יישור לימין כברירת מחדל
  if (!/\<w:pPrDefault\>[\s\S]*?<w:jc\b[\s\S]*?w:val="right"/.test(xml)) {
    console.log("ensureRtlInStyles: Setting default paragraph alignment to right.");
    xml = xml.replace(
      /<w:pPrDefault>\s*<w:pPr>/,
      `<w:pPrDefault>
    <w:pPr>
      <w:jc w:val="right"/>`
    );
  }

  // 2.2 style Normal – לוודא <w:rtl/> בתוך <w:rPr>
  const normalStyleRPrRegex =
    /(<w:style\b[^>]*w:styleId="Normal"[^>]*>[\s\S]*?<w:rPr\b[^>]*>)([\s\S]*?)(<\/w:rPr>)/;

  if (
    !/w:style\b[^>]*w:styleId="Normal"[\s\S]*?<w:rPr[\s\S]*?<w:rtl\b/.test(xml)
  ) {
    console.log('ensureRtlInStyles: Inserting <w:rtl/> into style "Normal".');
    xml = xml.replace(normalStyleRPrRegex, (match, start, middle, end) => {
      if (middle.includes("<w:rtl")) return match;
      return `${start}\n      <w:rtl/>\n${middle}${end}`;
    });
  }

  // 2.3 style Normal – לוודא יישור לימין בפסקה <w:pPr>
  const normalStylePPrRegex =
    /(<w:style\b[^>]*w:styleId="Normal"[^>]*>[\s\S]*?<w:pPr\b[^>]*>)([\s\S]*?)(<\/w:pPr>)/;

  if (
    !/w:style\b[^>]*w:styleId="Normal"[\s\S]*?<w:pPr[\s\S]*?<w:jc\b[\s\S]*?w:val="right"/.test(
      xml
    )
  ) {
    console.log('ensureRtlInStyles: Setting paragraph alignment "Normal" to right.');
    xml = xml.replace(normalStylePPrRegex, (match, start, middle, end) => {
      if (middle.includes("<w:jc")) return match;
      return `${start}\n      <w:jc w:val="right"/>\n${middle}${end}`;
    });
  }

  zip.file(path, xml);

  console.log(
    "ensureRtlInStyles: Done. New length:",
    xml.length,
    "(diff:",
    xml.length - originalLength,
    ")"
  );
}

/* ------------------------------------------------------------------
 * 3. NUMBERING.XML – רשימות/מספור RTL (יישור לימין והזחה מימין)
 * ------------------------------------------------------------------ */

async function ensureRtlInNumbering(zip) {
  const path = "word/numbering.xml";
  const xmlFile = zip.file(path);

  if (!xmlFile) {
    console.warn("ensureRtlInNumbering: numbering.xml not found, skipping.");
    return;
  }

  let xml = await xmlFile.async("string");
  const originalLength = xml.length;

  console.log("ensureRtlInNumbering: Loaded numbering.xml, length:", originalLength);

  function replaceAndCount(str, regex, replacement, label) {
    const matches = [...str.matchAll(regex)];
    if (!matches.length) {
      console.log(`ensureRtlInNumbering: No matches for ${label}.`);
      return str;
    }
    console.log(
      `ensureRtlInNumbering: Replacing ${matches.length} occurrence(s) of ${label}.`
    );
    return str.replace(regex, replacement);
  }

  // 3.1 יישור הרשימה לימין (lvlJc)
  const lvlJcRegex =
    /<w:lvlJc\b[^>]*w:val="(?:left|center)"[^>]*\/>/g;
  xml = replaceAndCount(
    xml,
    lvlJcRegex,
    '<w:lvlJc w:val="right"/>',
    "w:lvlJc left/center → right"
  );

  // 3.2 הזחה – left → right
  const leftIndentRegex =
    /<w:ind\b([^>]*?)\bw:left="(\d+)"([^>]*?)\/>/g;
  xml = replaceAndCount(
    xml,
    leftIndentRegex,
    '<w:ind$1 w:right="$2"$3/>',
    "w:ind w:left → w:right"
  );

  zip.file(path, xml);

  console.log(
    "ensureRtlInNumbering: Done. New length:",
    xml.length,
    "(diff:",
    xml.length - originalLength,
    ")"
  );
}
