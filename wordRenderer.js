import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
} from "docx";

/* ==========================================================================
   RLM FORCE – SECRET SAUCE
   ========================================================================== */

const RLM = "\u200F\u200F"; // double RLM = RTL anchor

/* ==========================================================================
   DEFAULT STYLES
   ========================================================================== */

const Styles = {
  title: { fontSize: 32, bold: true, spacingAfter: 500 },
  subtitle: { fontSize: 22, bold: true, spacingAfter: 350 },
  textBlock: { fontSize: 14, spacingAfter: 200 },
  instructions: { fontSize: 14, italics: true, spacingAfter: 300 },
  openQuestion: { fontSize: 16, bold: true, spacingAfter: 300 },
  mcqQuestion: { fontSize: 16, bold: true, spacingAfter: 150 },
  mcqOption: { fontSize: 14, spacingAfter: 120 },
  generic: { fontSize: 14, spacingAfter: 200 },
};

/* ==========================================================================
   RTL TEXT RUN – with RLM
   ========================================================================== */

function rtlText(text, style) {
  return new TextRun({
    text: `${RLM}${text}`, // ⭐ RLM injected here
    bold: style.bold || false,
    italics: style.italics || false,
    size: style.fontSize ? style.fontSize * 2 : 24, // docx uses half-points
    color: "000000",
    language: "he-IL",
    rightToLeft: true,
  });
}

/* ==========================================================================
   RTL PARAGRAPH – FULL RTL
   ========================================================================== */

function rtlParagraph(children, style, rtl) {
  return new Paragraph({
    children,
    alignment:
      style?.align === "center"
        ? AlignmentType.CENTER
        : rtl
        ? AlignmentType.RIGHT
        : AlignmentType.LEFT,

    bidirectional: rtl,
    rightToLeft: rtl,
    bidi: rtl,
    textDirection: "rtl",

    spacing: { after: style?.spacingAfter || 200 },
  });
}

/* ==========================================================================
   MULTI-LINE TEXT (each line gets RLM)
   ========================================================================== */

function textToParagraphs(text, style, rtl) {
  return text.split("\n").map((line) =>
    rtlParagraph([rtlText(line, style)], style, rtl)
  );
}

/* ==========================================================================
   RENDERERS
   ========================================================================== */

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

function renderSubtitle(block, rtl) {
  return textToParagraphs(block.text, Styles.subtitle, rtl);
}

function renderInstructions(block, rtl) {
  return textToParagraphs(block.text, Styles.instructions, rtl);
}

function renderTextBlock(block, rtl) {
  return textToParagraphs(block.text, Styles.textBlock, rtl);
}

function renderOpenQuestion(block, rtl) {
  return [
    rtlParagraph(
      [rtlText(`${block.id}. ${block.text}`, Styles.openQuestion)],
      Styles.openQuestion,
      rtl
    ),
  ];
}

function renderMCQ(block, rtl) {
  const paras = [];

  // שורת השאלה
  paras.push(
    rtlParagraph(
      [rtlText(`${block.id}. ${block.text}`, Styles.mcqQuestion)],
      Styles.mcqQuestion,
      rtl
    )
  );

  // תשובות בחירה
  block.options.forEach((opt) => {
    paras.push(
      rtlParagraph(
        [rtlText(`• ${opt}`, Styles.mcqOption)],
        Styles.mcqOption,
        rtl
      )
    );
  });

  return paras;
}

function renderGeneric(block, rtl) {
  return textToParagraphs(
    block.text || JSON.stringify(block),
    Styles.generic,
    rtl
  );
}

/* ==========================================================================
   MAIN EXPORT — CLEAN RTL VERSION WITH RLM
   ========================================================================== */

export async function renderExamToDocx(examJson) {
  const paragraphs = [];

  const rtl = examJson.direction === "ltr" ? false : true;

  if (examJson.title) {
    paragraphs.push(...renderTitle({ text: examJson.title }, rtl));
  }

  if (examJson.instructions) {
    paragraphs.push(
      ...renderInstructions({ text: examJson.instructions }, rtl)
    );
  }

  examJson.sections?.forEach((section) => {
    if (section.sectionTitle) {
      paragraphs.push(
        ...renderSubtitle({ text: section.sectionTitle }, rtl)
      );
    }

    if (section.content) {
      paragraphs.push(
        ...renderTextBlock({ text: section.content }, rtl)
      );
    }

    section.questions?.open?.forEach((q) => {
      paragraphs.push(...renderOpenQuestion(q, rtl));
    });

    section.questions?.multipleChoice?.forEach((q) => {
      paragraphs.push(...renderMCQ(q, rtl));
    });
  });

  const doc = new Document({
    styles: {
      default: {
        document: {
          run: {
            language: rtl ? "he-IL" : "en-US",
          },
          paragraph: {
            rightToLeft: rtl,
            bidi: rtl,
            alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
          },
        },
      },
      paragraphStyles: [
        {
          id: "Normal",
          name: "Normal",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: {
            language: rtl ? "he-IL" : "en-US",
            rightToLeft: rtl,
          },
          paragraph: {
            rightToLeft: rtl,
            bidi: rtl,
            alignment: rtl ? AlignmentType.RIGHT : AlignmentType.LEFT,
          },
        },
      ],
    },

    sections: [
      {
        properties: { rightToLeft: rtl },
        children: paragraphs,
      },
    ],
  });

  return await Packer.toBuffer(doc);
}
