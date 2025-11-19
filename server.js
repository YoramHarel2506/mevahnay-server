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
    const rtl = examJson.direction !== "ltr"; // ברירת מחדל: RTL אם לא כתוב ltr

    console.log("→ Starting DOCX generation request…");

    // 1) יצירת DOCX רגיל
    let docBuffer = await renderExamToDocx(examJson);
    console.log("✔ Base DOCX generated");

    // 2) אם זה מבחן RTL – מפעילים את תיקוני ה-XML
    if (rtl) {
      docBuffer = await applyRtlSettings(docBuffer);
      console.log("✔ applyRtlSettings done");

      docBuffer = await applyRtlParagraphs(docBuffer);
      console.log("✔ applyRtlParagraphs done");
    } else {
      console.log("ℹ direction=ltr → skipping RTL post-processing");
    }

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="exam_${Date.now()}.docx"`
    );

    return res.send(docBuffer);
  } catch (err) {
    console.error("❌ DOCX generation failed:", err);
    return res.status(500).send({ error: "DOCX creation failed" });
  }
});


app.listen(3000, () =>
  console.log("🚀 WORD RTL SERVER RUNNING ON PORT 3000")
);
