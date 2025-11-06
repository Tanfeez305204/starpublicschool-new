const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.static('public')); // frontend folder

// ✅ Read Excel safely
const readSheet = (sheetName) => {
  try {
    const wb = xlsx.readFile(path.join(__dirname, 'results.xlsx'));
    const sheet = wb.Sheets[sheetName];
    return sheet ? xlsx.utils.sheet_to_json(sheet) : null;
  } catch {
    return null;
  }
};

const safeValue = (val, fallback = "") =>
  val === undefined || val === null || String(val).trim() === "" ? fallback : String(val).trim();

const isNumeric = (val) => !isNaN(parseFloat(val)) && isFinite(val);

app.get('/result', (req, res) => {
  const queryClass = req.query.class?.trim().toUpperCase();
  const roll = req.query.roll?.trim();
  const terminal = req.query.terminal?.trim().toLowerCase();

  if (!queryClass || !roll)
    return res.json({ error: "Class and Roll number are required." });
  if (!terminal)
    return res.json({ error: "Please select terminal." });

  const readSheet = (sheetName) => {
    try {
      const wb = xlsx.readFile(path.join(__dirname, 'results.xlsx'));
      const sheet = wb.Sheets[sheetName];
      return sheet ? xlsx.utils.sheet_to_json(sheet) : null;
    } catch {
      return null;
    }
  };

  const safeValue = (val, fallback = "") =>
    val === undefined || val === null || String(val).trim() === ""
      ? fallback
      : String(val).trim();

  const isNumeric = (val) => !isNaN(parseFloat(val)) && isFinite(val);

  const sheets = {
    first: readSheet("result_1st") || [],
    second: readSheet("result_2nd") || [],
    third: readSheet("result_3rd") || [],
    annual: readSheet("result_annual") || [],
  };

  const findStudent = (sheet) =>
    sheet.find(
      (s) =>
        String(s.Class).trim().toUpperCase() === queryClass &&
        String(s.Roll).trim() === roll
    );

  const data = {
    first: findStudent(sheets.first),
    second: findStudent(sheets.second),
    third: findStudent(sheets.third),
    annual: findStudent(sheets.annual),
  };

  if (!data.first && !data.second && !data.third && !data.annual) {
    return res.json({ error: "Student not found." });
  }

  const normalize = (s) =>
    s.trim().toLowerCase().replace(/\./g, "").replace(/\s+/g, "");

  const allKeys = [
    ...Object.keys(data.first || {}),
    ...Object.keys(data.second || {}),
    ...Object.keys(data.third || {}),
    ...Object.keys(data.annual || {}),
  ].filter(
    (k) => !["Class", "Roll", "Name", "FatherName", "Father Name"].includes(k)
  );

  const subjects = allKeys.reduce((unique, sub) => {
    const n = normalize(sub);
    if (!unique.some((u) => normalize(u) === n)) unique.push(sub);
    return unique;
  }, []);

  const marks = subjects.map((sub) => ({
    subject: sub,
    fullMarks: 100,
    passMarks: 30,
    firstTerm: ["1st", "2nd", "3rd", "annual"].includes(terminal)
      ? data.first
        ? safeValue(data.first[sub])
        : ""
      : "",
    secondTerm: ["2nd", "3rd", "annual"].includes(terminal)
      ? data.second
        ? safeValue(data.second[sub])
        : ""
      : "",
    thirdTerm: ["3rd", "annual"].includes(terminal)
      ? data.third
        ? safeValue(data.third[sub])
        : ""
      : "",
    AnnTerm:
      terminal === "annual"
        ? data.annual
          ? safeValue(data.annual[sub])
          : ""
        : "",
  }));

  const calcTotal = (sheetData) =>
    sheetData
      ? Object.keys(sheetData)
          .filter(
            (k) => !["Class", "Roll", "Name", "FatherName", "Father Name"].includes(k)
          )
          .reduce(
            (sum, k) =>
              sum + (isNumeric(sheetData[k]) ? parseFloat(sheetData[k]) : 0),
            0
          )
      : 0;

  const totals = {
    first: calcTotal(data.first),
    second: calcTotal(data.second),
    third: calcTotal(data.third),
    annual: calcTotal(data.annual),
  };

  const termKeys = ["first", "second", "third", "annual"];
  const totalFullMarks = subjects.length * 100;

  // ✅ Calculate percentage for each term
  const percentages = {};
  termKeys.forEach((k) => {
    const termCount = termKeys.indexOf(k) + 1;
    const shownTotals = termKeys.slice(0, termCount).map((t) => totals[t]);
    const shownTotal = shownTotals.reduce((a, b) => a + (b || 0), 0);
    percentages[k] = ((shownTotal / (totalFullMarks * termCount)) * 100).toFixed(2);
  });

  const termMap = { "1st": "first", "2nd": "second", "3rd": "third", annual: "annual" };
  const selectedTerm = termMap[terminal] || "first";
  const percentage = parseFloat(percentages[selectedTerm] || 0);

  const division =
    percentage >= 60
      ? "First"
      : percentage >= 45
      ? "Second"
      : percentage >= 30
      ? "Third"
      : "Fail";

  res.json({
    schoolName: "STAR PUBLIC SCHOOL",
    schoolAddress: "Main road Mathia Bazar, Meghwal",
    studentName: safeValue(
      data.first?.Name ||
        data.second?.Name ||
        data.third?.Name ||
        data.annual?.Name
    ),
    fatherName: safeValue(
      data.first?.FatherName ||
        data.second?.FatherName ||
        data.third?.FatherName ||
        data.annual?.FatherName
    ),
    class: safeValue(queryClass),
    roll: safeValue(roll),
    terminal,
    marks,
    totals,
    totalFullMarks,
    percentageFirst: percentages.first,
    percentageSecond: percentages.second,
    percentageThird: percentages.third,
    percentageAnnual: percentages.annual,
    division,
    description:
      division === "Fail"
        ? "Needs Improvement."
        : "Keep up the good work!",
  });
});
// Read Excel sheet for provisional certificate
app.get('/provisional', (req, res) => {
  const queryClass = req.query.class?.trim().toUpperCase();
  const roll = req.query.roll?.trim();

  if (!queryClass || !roll) return res.json({ error: "Class and Roll required." });

  const sheet = readSheet("result_2nd") || []; // ya jo sheet provisional data ke liye ho
  const student = sheet.find(
    s => String(s.Class).trim().toUpperCase() === queryClass && String(s.Roll).trim() === roll
  );

  if (!student) return res.json({ error: "Student not found." });

  // Map Excel fields to certificate fields
  res.json({
    studentName: safeValue(student.Name),
    fatherName: safeValue(student.FatherName),
    schoolName: safeValue(student.SchoolName || "STAR PUBLIC SCHOOL, MATHIA"),
    class: safeValue(student.RollCode || student.Class),
    rollNo: safeValue(student.Roll),
    year: safeValue(student.Year || "Second Term Exam 2025"),
    division: safeValue(student.Division),
    date: new Date().toLocaleDateString('en-GB'),
    pcNo: safeValue(student.PCNo)
  });
});


app.listen(PORT, () => console.log(`✅ Server running at http://localhost:${PORT}`));
