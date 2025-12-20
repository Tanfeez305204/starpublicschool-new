
const fs = require("fs");
const express = require('express');
const xlsx = require('xlsx');
const path = require('path');
const cors = require('cors');
const app = express();
const bodyParser = require("body-parser");
const session = require("express-session");
const nodemailer = require("nodemailer");
const bcrypt = require("bcrypt");
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
    (k) => !["Class", "Roll", "Name", "FatherName", "Father Name","P_C NO","pcNO"].includes(k)
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
      if (val === undefined || val === null || val==="AB" || String(val).trim() === "")
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
const isResultAvailable = (termData, visibleSubjects) => {
  if (!termData) return false;

  return visibleSubjects.some((sub) => {
    const val = termData[sub];
    if (val === undefined || val === null) return false;
    const v = String(val).trim().toUpperCase();
    return !["", "AB", "-", "NA"].includes(v);
  });
};

  // ✅ Total Calculation — exclude Drawing
  const calcTotal = (sheetData) =>
    sheetData
      ? Object.keys(sheetData)
          .filter(
            (k) =>
              !["Class", "Roll", "Name", "FatherName", "Father Name","P_C NO","pcNO"].includes(k) &&
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
        if (["", "AB", "-"].includes(val)) hasIncomplete = true;
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
const terminalMap = {
  "1st": data.first,
  "2nd": data.second,
  "3rd": data.third,
  "annual": data.annual,
};

const selectedTermData = terminalMap[terminal];

const resultAvailable = isResultAvailable(
  selectedTermData,
  visibleSubjects
);

if (!resultAvailable) {
  return res.json({ error: "Result not available." });
}

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
  data.first?.FatherName || data.first?.["Father Name"] ||
    data.second?.FatherName || data.second?.["Father Name"] ||
    data.third?.FatherName || data.third?.["Father Name"] ||
    data.annual?.FatherName || data.annual?.["Father Name"]
),

    class: safeValue(queryClass),
    roll: safeValue(roll),
    terminal,
    session:
  terminal === "Annual" ? "2025-26" : "",

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
function generatePCNumber(existingPCs) {
  let pc;
  do {
    pc = "000" + Math.floor(100000 + Math.random() * 900000);
  } while (existingPCs.includes(pc));
  return pc;
}

app.get('/provisional', (req, res) => {
  let queryClass = (req.query.class || "").trim().toUpperCase();
  const roll = (req.query.roll || "").trim();

  const classMap = {
    NURA: "NURSERY-A",
    NURB: "NURSERY-B",
    NURC: "NURSERY-C",
    LKGA: "L.K.G-A",
    LKGB: "L.K.G-B",
    UKGA: "U.K.G-A",
    UKGB: "U.K.G-B"
  };

  if (classMap[queryClass]) queryClass = classMap[queryClass];

  if (!queryClass || !roll) {
    return res.json({ error: "Class and Roll required." });
  }

  // Read Excel
  const workbook = xlsx.readFile("results.xlsx");
  const sheetName = "result_annual";
  const sheet = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

  // Find student row directly (no normalization for finding)
  const student = sheet.find(
    s => String(s.Class).trim().toUpperCase() === queryClass && String(s.Roll).trim() === roll
  );

  if (!student) return res.json({ error: "Student not found." });

  // 🔹 PC number logic
  // Check if Excel already has a PC number
  let existingPC = student["P_C NO"] || student["P C NO"] || student["PC NO"];
  if (!existingPC || existingPC.trim() === "") {
    const existingPCs = sheet
      .map(r => r["P_C NO"] || r["P C NO"] || r["PC NO"] || "")
      .filter(Boolean);

    const newPC = generatePCNumber(existingPCs);

    // Save PC in Excel
    student["P_C NO"] = newPC;
    const updatedSheet = xlsx.utils.json_to_sheet(sheet);
    workbook.Sheets[sheetName] = updatedSheet;
    xlsx.writeFile(workbook, "results.xlsx");

    existingPC = newPC; // use for response
  }
const calculateDivision = (studentRow) => {
  let total = 0;
  let subjectCount = 0;
  let hasIncomplete = false;

  Object.keys(studentRow).forEach((key) => {
    if (
      ["Class","Roll","Name","Father Name","FatherName","P_C NO","P C NO","PC NO","Division","School Name","Year"].includes(key)
    ) return;

    const val = String(studentRow[key] || "").trim().toUpperCase();

   

    if (!isNaN(val)) {
      total += Number(val);
      subjectCount++;
    }
  });

  if (hasIncomplete) return "INCOMPLETE";

  const fullMarks = subjectCount * 100 || 1;
  const percent = (total / fullMarks) * 100;

  if (percent >= 60) return "First";
  if (percent >= 45) return "Second";
  if (percent >= 30) return "Third";
  return "Fail";
};

  // Response
  res.json({
    studentName: safeValue(student["Name"]),
    fatherName: safeValue(student["Father Name"]),
    schoolName: safeValue(student["School Name"] || "STAR PUBLIC SCHOOL, MATHIA"),
    class: safeValue(student["Class"]),
    rollNo: safeValue(student["Roll"]),
    year: safeValue(student["Year"] || "Second Term Exam 2025-26"),
division: safeValue(student["Division"] || calculateDivision(student)),
    date: new Date().toLocaleDateString("en-GB"),
    pcNo: existingPC
  });
});





app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(require("cors")());

// Session setup
app.use(session({
    secret: "admin_secret_key",
    resave: false,
    saveUninitialized: true,
}));

// Prevent caching for all routes (important after logout)
app.use((req, res, next) => {
    res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    res.setHeader("Pragma", "no-cache");
    res.setHeader("Expires", "0");
    res.setHeader("Surrogate-Control", "no-store");
    next();
});

// ----------------- Admin Credentials -----------------
const adminFile = path.join(__dirname, "admin.json");

// If admin.json does not exist, create it with default password
if (!fs.existsSync(adminFile)) {
    const defaultAdmin = {
        email: "t01auheed@gmail.com",
        passwordHash: bcrypt.hashSync("Admin@123", 10)
    };
    fs.writeFileSync(adminFile, JSON.stringify(defaultAdmin, null, 2), "utf-8");
}

let admin = JSON.parse(fs.readFileSync(adminFile, "utf-8"));

// OTP store
let otpStore = {};

// ----------------- Nodemailer Setup -----------------
const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
        user: "t01auheed@gmail.com", // your Gmail
        pass: "kkyt yulu dmas uxuf", // Gmail App Password
    },
});

// ----------------- Helper -----------------
function isAdminLoggedIn(req, res, next) {
    if (req.session && req.session.admin) {
        return next();
    } else {
        return res.redirect("/admin.html");
    }
}

// ----------------- Routes -----------------

// Admin login
app.post("/admin/login", async (req, res) => {
    const { email, password } = req.body;

    if (email.toLowerCase() !== admin.email.toLowerCase()) {
        return res.send("<h3>Invalid email. <a href='/admin.html'>Try again</a></h3>");
    }

    const match = await bcrypt.compare(password, admin.passwordHash);
    if (!match) {
        return res.send("<h3>Invalid password. <a href='/admin.html'>Try again</a></h3>");
    }

    req.session.admin = true;

    res.redirect("/provisional.html");
});

// Admin logout
app.get("/admin/logout", (req, res) => {
    req.session.destroy(err => {
        if (err) return res.send("Error logging out.");
        res.redirect("/admin.html");
    });
});

// Serve provisional page only if logged in
app.get("/provisional.html", isAdminLoggedIn, (req, res) => {
    res.sendFile(path.join(__dirname, "public", "provisional.html"));
});

// Request OTP for password reset
app.post("/admin/request-otp", async (req, res) => {
    const { email } = req.body;
    if (email.toLowerCase() !== admin.email.toLowerCase()) return res.json({ error: "Invalid email." });

    const otp = Math.floor(100000 + Math.random() * 900000).toString();
    const expires = Date.now() + 5 * 60 * 1000; // 5 minutes
    otpStore[email] = { otp, expires };

    try {
        await transporter.sendMail({
            from: '"Admin Reset" <t01auheed@gmail.com>',
            to: email,
            subject: "Your OTP for Admin Password Reset",
            text: `Your OTP is ${otp}. It expires in 5 minutes.`,
        });
        res.json({ success: true, message: "OTP sent to your email." });
    } catch (err) {
        console.error(err);
        res.json({ error: "Failed to send OTP. Check email credentials." });
    }
});

// Reset password using OTP
app.post("/admin/reset-password", async (req, res) => {
    const { email, otp, newPassword } = req.body;
    if (email.toLowerCase() !== admin.email.toLowerCase()) return res.json({ error: "Invalid email." });

    const record = otpStore[email];
    if (!record) return res.json({ error: "OTP not requested." });
    if (Date.now() > record.expires) return res.json({ error: "OTP expired." });
    if (record.otp !== otp) return res.json({ error: "Invalid OTP." });

    // Update password in memory and persist
    admin.passwordHash = await bcrypt.hash(newPassword, 10);
    fs.writeFileSync(adminFile, JSON.stringify(admin, null, 2), "utf-8");

    delete otpStore[email];
    res.json({ success: true, message: "Password updated successfully." });
});

// Protected dashboard
app.get("/admin/dashboard", isAdminLoggedIn, (req, res) => {
    res.send("<h1>Welcome Admin!</h1><a href='/admin/logout'>Logout</a>");
});

app.listen(PORT, () => console.log(`✅ Server running at http://localhost:${PORT}`));
