import express from "express";
import cors from "cors";
import { Document, Packer, Paragraph, TextRun } from "docx";

const app = express();
app.use(cors());
app.use(express.json());

// בדיקת שרת
app.get("/", (req, res) => {
  res.send({ status: "server running", name: "Mevahnay API" });
});

// יצירת מסמך DOCX ממבחן
app.post("/generate-docx", async (req, res) => {
  try {
    const exam = req.body;

    const paragraphs = [];

    // כותרת
    paragraphs.push(
      new Paragraph({
        children: [
          new TextRun({ text: exam.title || "Exam", bold: true, size: 40 })
        ],
        spacing: { after: 300 },
      })
    );

    // הוראות
    if (exam.instructions) {
      paragraphs.push(
        new Paragraph({
          children: [new TextRun({ text: exam.instructions, italics: true })],
          spacing: { after: 200 },
        })
      );
    }

    // שאלות
    exam.questions.forEach((q, index) => {
      paragraphs.push(
        new Paragraph({
          text: `${index + 1}. ${q.q}`,
          spacing: { after: 100 },
        })
      );
    });

    const doc = new Document({
      sections: [{ children: paragraphs }],
    });

    // יצירת buffer במקום שמירה לקובץ
    const buffer = await Packer.toBuffer(doc);

    // החזרת המסמך ישירות להורדה
    res.setHeader("Content-Disposition", "attachment; filename=exam.docx");
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );

    return res.send(buffer);

  } catch (error) {
    console.error(error);
    return res.status(500).send({ success: false, message: "Failed to generate DOCX" });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => console.log("Server running on port", port));
