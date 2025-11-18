import express from "express";
import cors from "cors";

const app = express();
app.use(cors());
app.use(express.json());

// ברירת מחדל — לבדיקה
app.get("/", (req, res) => {
  res.send({ status: "server running", name: "Mevahnay API" });
});

// יצירת מבחן לדוגמה
app.post("/create-test", (req, res) => {
  const { title, questions } = req.body;

  res.send({
    success: true,
    message: `Test '${title}' created successfully`,
    totalQuestions: questions?.length || 0
  });
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log("Server running on port", port);
});
