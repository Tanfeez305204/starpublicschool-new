
const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const cors = require('cors');
const app = express();
const PORT = process.env.PORT || 3000;
 
app.use(cors());
app.use(express.static('public')); // frontend folder

//  Read Excel safely
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

  if (!data.first && !data.second && !data.third && !data.annual)
    return res.json({ error: "Student not found." });

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

  const lowerClasses = [
    "NURSERY-A", "NURSERY-B", "NURSERY-C",
    "L.K.G-A", "L.K.G-B", "U.K.G-A", "U.K.G-B"
  ];
  const isLowerClass = lowerClasses.includes(queryClass);

  // ✅ Filter Science/SST for lower class
  const visibleSubjects = subjects.filter((sub) => {
    const name = sub.trim().toUpperCase();
    if (isLowerClass && (name.includes("SCIENCE") || name.includes("S.S.T")))
      return false; // Hide
    return true;
  });

  const marks = visibleSubjects.map((sub) => {
    const normalizedSub = sub.trim().toUpperCase();
    const isDrawing = normalizedSub.includes("DRAWING");

    let fullMarks = 100;
    let passMarks = 30;

    // ✅ Drawing → grade only
    if (isDrawing) {
      fullMarks = "Grade";
      passMarks = "-";
    }

    const getVal = (termData) => {
      if (!termData) return "AB";
      const val = termData[sub];
      if (val === undefined || val === null || String(val).trim() === "")
        return "AB";
      const v = String(val).trim();
      if (["-", "_"].includes(v)) return "NA";
      return isNumeric(v) ? parseFloat(v) : v;
    };

    return {
      subject: sub,
      fullMarks,
      passMarks,
      firstTerm: ["1st", "2nd", "3rd", "annual"].includes(terminal)
        ? getVal(data.first)
        : 0,
      secondTerm: ["2nd", "3rd", "annual"].includes(terminal)
        ? getVal(data.second)
        : 0,
      thirdTerm: ["3rd", "annual"].includes(terminal)
        ? getVal(data.third)
        : 0,
      AnnTerm: terminal === "annual" ? getVal(data.annual) : 0,
    };
  });

  // ✅ Total Calculation — exclude Drawing
  const calcTotal = (sheetData) =>
    sheetData
      ? Object.keys(sheetData)
          .filter(
            (k) =>
              !["Class", "Roll", "Name", "FatherName", "Father Name"].includes(k) &&
              !k.toLowerCase().includes("drawing")
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

  // ✅ Total full marks calculation
  const totalFullMarks = subjects.filter(sub => {
    const name = sub.trim().toUpperCase();
    if (name.includes("DRAWING")) return false;
    if (isLowerClass && (name.includes("SCIENCE") || name.includes("S.S.T"))) return false;
    return true;
  }).length * 100;

  // ✅ Percentage logic
  const termKeys = ["first", "second", "third", "annual"];
  const percentages = {};
  termKeys.forEach((term) => {
    const totalObtained = totals[term] || 0;
    const totalFull = totalFullMarks || 1;
    percentages[term] = ((totalObtained / totalFull) * 100).toFixed(2);
  });

  // ✅ Division
  const division = {};
  termKeys.forEach((k) => {
    const termData = data[k];
    const perc = parseFloat(percentages[k] || 0);
    let hasIncomplete = false;

    if (termData) {
      Object.keys(termData).forEach((key) => {
        if (["Class", "Roll", "Name", "FatherName", "Father Name"].includes(key))
          return;
        const val = String(termData[key] || "").trim().toUpperCase();
        if (["", "AB", "NA", "-", "_"].includes(val)) hasIncomplete = true;
      });
    }

    division[k] = hasIncomplete
      ? "INCOMPLETE"
      : perc >= 60
      ? "First"
      : perc >= 45
      ? "Second"
      : perc >= 30
      ? "Third"
      : "Fail";
  });

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
      Object.values(division).includes("Fail")
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

  // Normalize keys
  const normalizedStudent = {};
  for (const key in student) {
    const normalizedKey = key.trim().toLowerCase().replace(/\s+/g, '');
    normalizedStudent[normalizedKey] = student[key];
  }

  // Map Excel fields to certificate fields
  res.json({
    studentName: safeValue(normalizedStudent["name"]),
    fatherName: safeValue(normalizedStudent["fathername"]),
    schoolName: safeValue(normalizedStudent["schoolname"] || "STAR PUBLIC SCHOOL, MATHIA"),
    class: safeValue(normalizedStudent["rollcode"] || normalizedStudent["class"]),
    rollNo: safeValue(normalizedStudent["roll"]),
    year: safeValue(normalizedStudent["year"] || "Second Term Exam 2025"),
    division: safeValue(normalizedStudent["division"]),
    date: new Date().toLocaleDateString('en-GB'),
    pcNo: safeValue(normalizedStudent["pcno"])
  });
});

app.listen(PORT, () => console.log(`✅ Server running at http://localhost:${PORT}`));
