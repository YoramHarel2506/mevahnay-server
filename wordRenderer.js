import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  Header,
  Footer,
  PageNumber,
  Table,
  TableRow,
  TableCell,
  WidthType,
  BorderStyle
} from "docx";

/* ======================================================================
   RLM – לשימור כיוון כתיבה
   ====================================================================== */
const RLM = "\u200F\u200F";

/* ======================================================================
   STYLES – כל העיצובים במקום אחד
   ====================================================================== */
const BASE_FONT = "Arial";

const Styles = {
  title: { fontSize: 26, bold: true, spacingAfter: 400 },
  sectionHeader: { fontSize: 22, bold: true, spacingAfter: 350 },
  sectionTitle: { fontSize: 20, bold: true, spacingAfter: 300 },
  textBlock: { fontSize: 14, spacingAfter: 200, lineSpacing: 1.2 },
  instructions: { fontSize: 14, italics: true, spacingAfter: 300 },
  openQuestion: { fontSize: 16, bold: true, spacingAfter: 150 },
  mcqQuestion: { fontSize: 16, bold: true, spacingAfter: 120 },
  mcqOption: { fontSize: 14, spacingAfter: 100 },
  answerLine: { fontSize: 10, spacingAfter: 150 },
  generic: { fontSize: 14, spacingAfter: 200 },
  teacherHeading: { fontSize: 22, bold: true, spacingAfter: 300 },
};

/* ======================================================================
   RTL TEXT RUN
   ====================================================================== */
function rtlText(text, style) {
  return new TextRun({
    text: `${RLM}${text}`,
    bold: !!style.bold,
    italics: !!style.italics,
    size: (style.fontSize || 14) * 2,
    font: BASE_FONT,
    rightToLeft: true,
    color: "000000",
  });
}

/* ======================================================================
   RTL PARAGRAPH
   ====================================================================== */
function rtlParagraph(children, style, rtl) {
  return new Paragraph({
    children,
    alignment:
      style?.align === "center"
        ? AlignmentType.CENTER
        : rtl
        ? AlignmentType.RIGHT
        : AlignmentType.LEFT,
    rightToLeft: rtl,
    bidi: rtl,
    textDirection: rtl ? "rtl" : "ltr",
    spacing: { after: style.spacingAfter || 200 },
  });
}

/* ======================================================================
   MULTI-LINE TEXT
   ====================================================================== */
function textToParagraphs(text, style, rtl) {
  return text.split("\n").map((line) =>
    rtlParagraph([rtlText(line, style)], style, rtl)
  );
}

/* ======================================================================
   TITLE
   ====================================================================== */
function renderTitle(block, rtl) {
  return [
    rtlParagraph(
      [
        rtlText(block.text, {
          bold: true,
          fontSize: Styles.title.fontSize,
        }),
      ],
      { ...Styles.title, align: "center" },
      rtl
    ),
  ];
}

/* ======================================================================
   SUBTITLE / INSTRUCTIONS
   ====================================================================== */
function renderSubtitle(block, rtl) {
  return textToParagraphs(block.text, Styles.subtitle, rtl);
}

function renderInstructions(block, rtl) {
  return textToParagraphs(block.text, Styles.instructions, rtl);
}

function renderTextBlock(block, rtl) {
  return textToParagraphs(block.text, Styles.textBlock, rtl);
}

/* ======================================================================
   SECTION HEADER — חלק א / חלק ב / חלק ג
   ====================================================================== */
function renderSectionHeader(index, rtl) {
  const parts = ["א", "ב", "ג", "ד", "ה", "ו", "ז"];
  const label = `חלק ${parts[index]} `;

  return [
    rtlParagraph(
      [rtlText(label, Styles.sectionHeader)],
      { ...Styles.sectionHeader, align: "center" },
      rtl
    ),
  ];
}

/* ======================================================================
   SECTION TITLE + TOP LINE
   ====================================================================== */
function renderSectionTitleWithLine(title, rtl) {
  return [
    new Paragraph({
      children: [],
      border: {
        top: { color: "999999", size: 4 },
      },
      spacing: { after: 150 },
    }),

    rtlParagraph([rtlText(title, Styles.sectionTitle)], Styles.sectionTitle, rtl),
  ];
}

/* ======================================================================
   TEXT BOX – קטע מידע
   ====================================================================== */
function renderTextBox(content, rtl) {
  const inside = content.split("\n").map((line) =>
    rtlParagraph([rtlText(line, Styles.textBlock)], Styles.textBlock, rtl)
  );

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            children: inside,
            borders: {
              top: { style: BorderStyle.SINGLE, size: 4, color: "555555" },
              bottom: { style: BorderStyle.SINGLE, size: 4, color: "555555" },
              left: { style: BorderStyle.SINGLE, size: 4, color: "555555" },
              right: { style: BorderStyle.SINGLE, size: 4, color: "555555" },
            },
            margins: {
              top: 200,
              bottom: 200,
              left: 200,
              right: 200,
            },
          }),
        ],
      }),
    ],
  });
}

/* ======================================================================
   ANSWER LINES (5 lines)
   ====================================================================== */
function renderAnswerLinesBlock(rtl) {
const dashed = "_".repeat(80);
  const lines = [];

  for (let i = 0; i < 5; i++) {
    lines.push(
      rtlParagraph(
        [
          new TextRun({
            text: `${RLM}${dashed}`,
            size: Styles.answerLine.fontSize * 2,
            color: "777777",
            font: BASE_FONT,
            rightToLeft: rtl,
          }),
        ],
        Styles.answerLine,
        rtl
      )
    );
  }

  return lines;
}

/* ======================================================================
   OPEN QUESTION
   ====================================================================== */
function renderOpenQuestion(block, rtl) {
  const paras = [];

  paras.push(rtlParagraph([rtlText(" ", Styles.generic)], Styles.generic, rtl));

  const qText = `${block.id}. ${block.text} (${block.points} נק')`;

  paras.push(
    rtlParagraph([rtlText(qText, Styles.openQuestion)], Styles.openQuestion, rtl)
  );

  paras.push(...renderAnswerLinesBlock(rtl));

  paras.push(rtlParagraph([rtlText(" ", Styles.generic)], Styles.generic, rtl));

  return paras;
}

/* ======================================================================
   MULTIPLE CHOICE
   ====================================================================== */
function renderMCQ(block, rtl) {
  const paras = [];

  paras.push(rtlParagraph([rtlText(" ", Styles.generic)], Styles.generic, rtl));

  const qText = `${block.id}. ${block.text} (${block.points} נק')`;

  paras.push(
    rtlParagraph([rtlText(qText, Styles.mcqQuestion)], Styles.mcqQuestion, rtl)
  );

  const rtlLabels = ["א.", "ב.", "ג.", "ד.", "ה."];

  block.options.forEach((opt, i) => {
    paras.push(
      rtlParagraph(
        [rtlText(`${rtlLabels[i]} ${opt}`, Styles.mcqOption)],
        Styles.mcqOption,
        rtl
      )
    );
  });

  paras.push(rtlParagraph([rtlText(" ", Styles.generic)], Styles.generic, rtl));

  return paras;
}

/* ======================================================================
   HEADER
   ====================================================================== */
function buildHeader(title, rtl) {
  let clean = title.trim();
  if (clean.startsWith("מבחן")) clean = clean.replace(/^מבחן[- ]*/, "");

  const finalText = `מבחן – ${clean}`;

  return new Header({
    children: [
      rtlParagraph(
        [rtlText(finalText, { fontSize: 10, bold: false })],
        { fontSize: 14, bold: true, align: "center" },
        rtl
      ),
    ],
  });
}

/* ======================================================================
   FOOTER
   ====================================================================== */



/* ======================================================================
   SOLUTION BOOKLET
   ====================================================================== */
function renderTeacherGuideHeader(rtl) {
  return [
    rtlParagraph(
      [rtlText("מחברת פתרונות – למורה", Styles.teacherHeading)],
      { ...Styles.teacherHeading, align: "center" },
      rtl
    ),
    new Paragraph({
      children: [],
      border: { bottom: { color: "000000", size: 6 } },
      spacing: { after: 300 },
    }),
  ];
}

function renderTeacherGuideOpen(open, rtl) {
  const paras = [];
  if (!open) return paras;

  paras.push(
    rtlParagraph(
      [rtlText("פתרונות – שאלות פתוחות", { fontSize: 18, bold: true })],
      { fontSize: 18, bold: true },
      rtl
    )
  );

  Object.keys(open).forEach((qid) => {
    paras.push(
      rtlParagraph(
        [rtlText(`${qid}. ${open[qid]}`, Styles.textBlock)],
        Styles.textBlock,
        rtl
      )
    );
  });

  return paras;
}

function renderTeacherGuideMCQ(mcq, rtl) {
  const paras = [];
  if (!mcq) return paras;

  paras.push(
    rtlParagraph(
      [rtlText("פתרונות – שאלות אמריקאיות", { fontSize: 18, bold: true })],
      { fontSize: 18, bold: true },
      rtl
    )
  );

  Object.keys(mcq).forEach((qid) => {
    paras.push(
      rtlParagraph(
        [rtlText(`${qid}. ${mcq[qid]}`, Styles.textBlock)],
        Styles.textBlock,
        rtl
      )
    );
  });

  return paras;
}

function buildFooter(rtl) {
  if (rtl) {
    return new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          rightToLeft: true,
          bidi: true,
          children: [
            new TextRun({
              text: `${RLM}עמוד `,
              font: BASE_FONT,
              size: 20,
              rightToLeft: true
            }),
            new TextRun({
              children: [PageNumber.CURRENT],
              font: BASE_FONT,
              size: 20,
              rightToLeft: true
            }),
            new TextRun({
              text: " מתוך ",
              font: BASE_FONT,
              size: 20,
              rightToLeft: true
            }),
            new TextRun({
              children: [PageNumber.TOTAL_PAGES],
              font: BASE_FONT,
              size: 20,
              rightToLeft: true
            }),
          ]
        })
      ]
    });
  }

  // LTR
  return new Footer({
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({
            text: "Page ",
            font: BASE_FONT,
            size: 20
          }),
          new TextRun({
            children: [PageNumber.CURRENT],
            font: BASE_FONT,
            size: 20
          }),
          new TextRun({
            text: " of ",
            font: BASE_FONT,
            size: 20
          }),
          new TextRun({
            children: [PageNumber.TOTAL_PAGES],
            font: BASE_FONT,
            size: 20
          }),
        ]
      })
    ]
  });
}

function renderTeacherGuide(json, rtl) {
  if (!json.teacherGuide) return [];

  return [
    ...renderTeacherGuideHeader(rtl),
    ...renderTeacherGuideOpen(json.teacherGuide.open, rtl),
    rtlParagraph([rtlText(" ", Styles.generic)], Styles.generic, rtl),
    ...renderTeacherGuideMCQ(json.teacherGuide.multipleChoice, rtl),
  ];
}

/* ======================================================================
   PAGE SETUP
   ====================================================================== */
const pageSetup = {
  margin: {
    top: 1440,
    bottom: 1440,
    left: 1134,
    right: 1134,
  },
};

/* ======================================================================
   MAIN EXPORT
   ====================================================================== */
export async function renderExamToDocx(examJson) {

  if (typeof examJson === "string") {
  try {
    console.log("🔵 Parsing examJson string…");
    examJson = JSON.parse(examJson);
  } catch (e) {
    console.error("❌ FAILED TO PARSE examJson:", e, examJson);
    throw e;
  }
}  
  const rtl = examJson.direction !== "ltr";
  const paragraphs = [];
  let index = 0;

  if (examJson.title) {
    paragraphs.push(...renderTitle({ text: examJson.title }, rtl));
  }

  if (examJson.instructions) {
    paragraphs.push(...renderInstructions({ text: examJson.instructions }, rtl));
  }

  examJson.sections?.forEach((section) => {
    paragraphs.push(...renderSectionHeader(index++, rtl));

    if (section.sectionTitle)
      paragraphs.push(...renderSectionTitleWithLine(section.sectionTitle, rtl));

    if (section.content) paragraphs.push(renderTextBox(section.content, rtl));

    section.questions?.open?.forEach((q) =>
      paragraphs.push(...renderOpenQuestion(q, rtl))
    );

    section.questions?.multipleChoice?.forEach((q) =>
      paragraphs.push(...renderMCQ(q, rtl))
    );
  });

  const teacher = renderTeacherGuide(examJson, rtl);

  if (teacher.length > 0) {
    paragraphs.push(new Paragraph({ pageBreakBefore: true }));
    paragraphs.push(...teacher);
  }

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: { font: BASE_FONT, language: rtl ? "he-IL" : "en-US" },
          paragraph: {
            rightToLeft: rtl,
            bidi: rtl,
            alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
          },
        },
      },
    },
    sections: [
      {
        headers: { default: buildHeader(examJson.title, rtl) },
        footers: { default: buildFooter(rtl) },
        properties: { rightToLeft: rtl, ...pageSetup },
        children: paragraphs,
      },
    ],
  });

  return Packer.toBuffer(doc);
}
