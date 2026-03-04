const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.static('public')); // To serve HTML, CSS, logo.png, etc.

const workbook = xlsx.readFile(path.join(__dirname, 'results.xlsx'));
const sheet = workbook.Sheets['result']; // Only one sheet now
const allData = xlsx.utils.sheet_to_json(sheet);

// API to fetch student result
app.get('/result', (req, res) => {
  const studentClass = req.query.class?.trim();
  const roll = req.query.roll?.trim();

  if (!studentClass || !roll) {
    return res.json({ error: "Class and Roll number are required." });
  }

  // Filter student by class and roll
  const student = allData.find(s =>
    String(s.Class).trim() === studentClass && String(s.Roll).trim() === roll
  );

  if (!student) {
    return res.json({ error: "Student not found." });
  }

  // Prepare marks
  const marks = [];
  let total = 0;
  const fullMarksPerSubject = 100;

  for (const key in student) {
    if (!['Class', 'Roll', 'Name', 'FatherName'].includes(key)) {
      const obtained = parseFloat(student[key]) || 0;
      marks.push({
        subject: key,
        fullMarks: fullMarksPerSubject,
        obtainedMarks: obtained
      });
      total += obtained;
    }
  }

  const percentage = (total / (marks.length * fullMarksPerSubject)) * 100;
  const division =
    percentage >= 60 ? 'First' :
    percentage >= 45 ? 'Second' :
    percentage >= 30 ? 'Third' :
    'Fail';

  res.json({
    schoolName: "STAR PUBLIC SCHOOL",
    schoolAddress: "Main road Mathia Bazar, Maghwal",
    studentName: student.Name,
    fatherName: student.FatherName,
    class: studentClass,
    roll: roll,
    marks,
    total,
    percentage: percentage.toFixed(2),
    division,
    description: division === "Fail" ? "Needs Improvement." : "Keep up the good work!"
  });
});

app.listen(PORT, () => {
  console.log(`✅ Server is running at http://localhost:${PORT}`);
});
