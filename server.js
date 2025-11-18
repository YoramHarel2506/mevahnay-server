import express from "express";
import cors from "cors";
import { Document, Packer, Paragraph, TextRun } from "docx";
import fs from "fs";
import path from "path";

const app = express();
app.use(cors());
app.use(express.json());

app.get("/test", (req, res) => {
  res.send("route works!");
});


// בדיקת שרת
app.get("/", (req, res) => {
  res.send({ status: "server running", name: "Mevahnay API" });
});

// הגשת קבצים סטטיים
app.use("/files", express.static("files"));

// יצירת מסמך WORD ממבחן
app.post("/generate-docx", async (req, res) => {
  try {
    const exam = req.body; // JSON מהאפליקציה

    // יצירת פסקאות
    const paragraphs = [];

    // כותרת
    paragraphs.push(
      new Paragraph({
        children: [new TextRun({ text: exam.title || "Exam", bold: true, size: 40 })],
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

    // יצירת מסמך
    const doc = new Document({
      sections: [{ children: paragraphs }],
    });

    // שם קובץ
    const fileName = `exam_${Date.now()}.docx`;
    const filePath = path.join("files", fileName);

    // יצירת הקובץ
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);

    // החזרת URL להורדה
    return res.send({
      success: true,
      url: `https://mevahnay-server.onrender.com/files/${fileName}`,
    });
  } catch (error) {
    console.error(error);
    return res.status(500).send({ success: false, message: "Failed to generate DOCX" });
  }
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log("Server running on port", port);
});
