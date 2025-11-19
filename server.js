import express from "express";
import cors from "cors";

// מחולל המסמך שלנו
import { renderExamToDocx } from "./wordRenderer.js";

// שלב 1: שינוי settings/styles/numbering למסמך RTL
import { applyRtlSettings } from "./applyRtlSettings.js";

// שלב 2: כפיית RTL על כל הפסקאות ב-document.xml
import { applyRtlParagraphs } from "./applyRtlParagraphs.js";

const app = express();
app.use(cors());
app.use(express.json());

app.post("/generate-docx", async (req, res) => {
  try {
    const examJson = req.body;

    console.log("→ Starting DOCX generation request…");

    // 1) יצירת DOCX רגיל (מ־docx)
    const baseDoc = await renderExamToDocx(examJson);
    console.log("✔ Base DOCX generated");

    // 2) RTL-level במסמך (settings/styles/numbering)
    const rtlDoc1 = await applyRtlSettings(baseDoc);
    console.log("✔ applyRtlSettings done");

    // 3) כפיית RTL על כל הפסקאות ב-document.xml
    const rtlDoc2 = await applyRtlParagraphs(rtlDoc1);
    console.log("✔ applyRtlParagraphs done");

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="exam_${Date.now()}.docx"`
    );

    return res.send(rtlDoc2);
  } catch (err) {
    console.error("❌ DOCX generation failed:", err);
    return res.status(500).send({ error: "DOCX creation failed" });
  }
});

app.listen(3000, () =>
  console.log("🚀 WORD RTL SERVER RUNNING ON PORT 3000")
);
