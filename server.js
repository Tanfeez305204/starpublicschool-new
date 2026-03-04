
require("dotenv").config();   // âœ… SABSE UPAR

const fs = require("fs");
const express = require("express");
const Excel = require("exceljs");
const path = require("path");
const cors = require("cors");
const bodyParser = require("body-parser");
const session = require("express-session");
const nodemailer = require("nodemailer");
const bcrypt = require("bcrypt");
const { createClient } = require("@supabase/supabase-js");

const app = express();
const PORT = process.env.PORT || 3000;

const Mailjet = require("node-mailjet");

const mailjet = Mailjet.apiConnect(
  process.env.SMTP_USER,
  process.env.SMTP_PASS
);

 
app.use(cors());

//  Read Excel safely (using exceljs, returns array of row objects)
const sanitizeCellObject = (obj) => {
  if (obj && typeof obj === "object") {
    if (Array.isArray(obj)) {
      return obj.map(sanitizeCellObject);
    }
    const clean = {};
    for (const key of Object.keys(obj)) {
      if (key === "__proto__" || key === "constructor" || key === "prototype") continue;
      let val = obj[key];
      if (typeof val === "object" && val !== null) val = sanitizeCellObject(val);
      clean[key] = val;
    }
    return clean;
  }
  return obj;
};

async function readSheet(sheetName) {
  try {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(path.join(__dirname, 'results.xlsx'));
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) return null;

    const headerRow = worksheet.getRow(1);
    const headers = headerRow.values.slice(1); // first cell is null
    const rows = [];
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber === 1) return; // skip header
      const rowObj = {};
      row.values.slice(1).forEach((v, i) => {
        rowObj[headers[i]] = v;
      });
      rows.push(rowObj);
    });
    return sanitizeCellObject(rows);
  } catch (err) {
    console.error("Error reading Excel sheet", err);
    return null;
  }
}

const safeValue = (val, fallback = "") =>
  val === undefined || val === null || String(val).trim() === "" ? fallback : String(val).trim();

const isNumeric = (val) => !isNaN(parseFloat(val)) && isFinite(val);

const normalizeClassToken = (cls = "") =>
  String(cls)
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "")
    .replace(/\./g, "")
    .replace(/_/g, "")
    .replace(/-/g, "");

const CLASS_ALIAS_MAP = {
  NURA: "NURSERY-A",
  NURB: "NURSERY-B",
  NURC: "NURSERY-C",
  NURSERYA: "NURSERY-A",
  NURSERYB: "NURSERY-B",
  NURSERYC: "NURSERY-C",
  LKGA: "L.K.G-A",
  LKGB: "L.K.G-B",
  UKGA: "U.K.G-A",
  UKGB: "U.K.G-B",
  CLASS1: "CLASS 1",
  CLASS2: "CLASS 2",
  CLASS3: "CLASS 3",
  CLASS4: "CLASS 4",
  CLASS5: "CLASS 5",
  CLASS6: "CLASS 6",
  CLASS7: "CLASS 7",
  CLASS8: "CLASS 8",
};

const canonicalizeClass = (cls = "") => {
  const token = normalizeClassToken(cls);
  return CLASS_ALIAS_MAP[token] || String(cls).trim().toUpperCase();
};

const CLASS_SORT_ORDER = [
  "NURSERYA",
  "NURSERYB",
  "NURSERYC",
  "LKGA",
  "LKGB",
  "UKGA",
  "UKGB",
  "CLASS1",
  "CLASS2",
  "CLASS3",
  "CLASS4",
  "CLASS5",
  "CLASS6",
  "CLASS7",
  "CLASS8",
];

const normalizeClassSortToken = (className = "") => {
  const canonicalClassName = canonicalizeClass(className);
  const token = normalizeClassToken(canonicalClassName);
  return /^\d+$/.test(token) ? `CLASS${token}` : token;
};

const classSortIndex = (className = "") => {
  const token = normalizeClassSortToken(className);
  const index = CLASS_SORT_ORDER.indexOf(token);
  return index === -1 ? 999 : index;
};

const parseRollForSort = (rollValue = "") => {
  const rawRoll = safeValue(rollValue);
  if (/^\d+$/.test(rawRoll)) {
    return {
      isNumeric: true,
      numeric: Number.parseInt(rawRoll, 10),
      text: rawRoll,
    };
  }

  return {
    isNumeric: false,
    numeric: Number.POSITIVE_INFINITY,
    text: rawRoll.toUpperCase(),
  };
};

const compareRollForSort = (leftRoll = "", rightRoll = "") => {
  const left = parseRollForSort(leftRoll);
  const right = parseRollForSort(rightRoll);

  if (left.isNumeric && right.isNumeric && left.numeric !== right.numeric) {
    return left.numeric - right.numeric;
  }

  if (left.isNumeric !== right.isNumeric) {
    return left.isNumeric ? -1 : 1;
  }

  return left.text.localeCompare(right.text, undefined, { numeric: true });
};

const compareStudentsForDirectory = (leftRow, rightRow) => {
  const leftClass = canonicalizeClass(leftRow?.classes?.class_name || "");
  const rightClass = canonicalizeClass(rightRow?.classes?.class_name || "");

  const classRankDiff = classSortIndex(leftClass) - classSortIndex(rightClass);
  if (classRankDiff !== 0) return classRankDiff;

  const classNameDiff = leftClass.localeCompare(rightClass, undefined, {
    numeric: true,
    sensitivity: "base",
  });
  if (classNameDiff !== 0) return classNameDiff;

  const rollDiff = compareRollForSort(leftRow?.roll_no, rightRow?.roll_no);
  if (rollDiff !== 0) return rollDiff;

  const leftName = safeValue(leftRow?.full_name);
  const rightName = safeValue(rightRow?.full_name);
  return leftName.localeCompare(rightName, undefined, {
    numeric: true,
    sensitivity: "base",
  });
};

const LOWER_CLASS_SET = new Set([
  "NURSERY-A",
  "NURSERY-B",
  "NURSERY-C",
  "L.K.G-A",
  "L.K.G-B",
  "U.K.G-A",
  "U.K.G-B",
]);

const isLowerClassName = (cls = "") => LOWER_CLASS_SET.has(canonicalizeClass(cls));

const normalizeSubjectToken = (sub = "") =>
  String(sub).trim().toUpperCase().replace(/\s+/g, "").replace(/\./g, "");

const isScienceOrSstSubject = (sub = "") => {
  const token = normalizeSubjectToken(sub);
  return token === "SCIENCE" || token === "SST" || token === "SOCIALSCIENCE";
};

const SCHOOL_NAME = "STAR PUBLIC SCHOOL";
const SCHOOL_ADDRESS = "Main road Mathia Bazar, Meghwal";
const DEFAULT_ACADEMIC_SESSION = "2025-26";

const EXAM_NAME_BY_CODE = {
  "1st": "First Terminal",
  "2nd": "Second Terminal",
  "3rd": "Third Terminal",
  annual: "Annual",
};

const DEFAULT_STANDARD_SUBJECTS = [
  "Hindi",
  "English",
  "Math",
  "Science",
  "S.S.T",
  "GK",
  "Sanskrit / Urdu",
  "Drawing",
  "Computer / Table",
];

const DEFAULT_LOWER_SUBJECTS = [
  "Hindi",
  "English",
  "Math",
  "GK",
  "Drawing",
  "Computer / Table",
];

const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const isRetryableNetworkError = (error) => {
  const message = safeValue(error?.message);
  const causeMessage = safeValue(error?.cause?.message);
  const code = safeValue(error?.cause?.code || error?.code);
  const combined = `${message} ${causeMessage} ${code}`;
  return /fetch failed|UND_ERR_SOCKET|ECONNRESET|ETIMEDOUT|EAI_AGAIN|ENETUNREACH|ECONNREFUSED|EPIPE/i.test(
    combined
  );
};

const supabaseFetchWithRetry = async (input, init) => {
  const maxAttempts = 3;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      const response = await fetch(input, init);

      if (
        response.status >= 500 &&
        response.status < 600 &&
        attempt < maxAttempts
      ) {
        await wait(150 * attempt);
        continue;
      }

      return response;
    } catch (error) {
      if (!isRetryableNetworkError(error) || attempt >= maxAttempts) {
        throw error;
      }

      console.warn(
        `Supabase transient network failure (attempt ${attempt}/${maxAttempts}). Retrying...`
      );
      await wait(150 * attempt);
    }
  }

  throw new Error("Supabase request failed after retries.");
};

const SUPABASE_URL = safeValue(process.env.SUPABASE_URL);
const SUPABASE_SERVICE_ROLE_KEY = safeValue(
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_ANON_KEY
);

const supabase =
  SUPABASE_URL && SUPABASE_SERVICE_ROLE_KEY
    ? createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, {
        auth: { persistSession: false, autoRefreshToken: false },
        global: { fetch: supabaseFetchWithRetry },
      })
    : null;

if (!supabase) {
  console.warn(
    "Supabase is not configured. Set SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY to enable database mode."
  );
}

const termToExamCodeMap = {
  first: "1st",
  second: "2nd",
  third: "3rd",
  annual: "annual",
};

const normalizeExamCode = (code = "") => {
  const token = String(code).trim().toLowerCase();
  if (token === "first" || token === "1st") return "1st";
  if (token === "second" || token === "2nd") return "2nd";
  if (token === "third" || token === "3rd") return "3rd";
  if (token === "annual" || token === "final") return "annual";
  return "";
};

const toDbSubjectCode = (subjectName = "") =>
  String(subjectName)
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");

const isSupabaseEnabled = () => Boolean(supabase);

const getDefaultSubjectsForClass = (className = "") => {
  const names = isLowerClassName(className)
    ? DEFAULT_LOWER_SUBJECTS
    : DEFAULT_STANDARD_SUBJECTS;
  return names.map((subjectName, idx) => ({
    subjectName,
    subjectCode: toDbSubjectCode(subjectName),
    fullMarks: 100,
    passMarks: 30,
    displayOrder: idx + 1,
  }));
};

const shouldExcludeSubject = (subjectName = "", className = "") => {
  const token = normalizeSubjectToken(subjectName);
  if (!token) return true;
  if (token.includes("DRAWING")) return true;
  if (isLowerClassName(className) && isScienceOrSstSubject(subjectName))
    return true;
  return false;
};

const SPECIAL_MARK_CODES = new Set(["AB", "NA", "-"]);

const parseMarkValue = (rawValue) => {
  if (rawValue === null || rawValue === undefined) {
    return { marks_obtained: null, grade: null };
  }

  if (typeof rawValue === "number" && Number.isFinite(rawValue)) {
    return {
      marks_obtained: Math.max(0, Math.min(100, rawValue)),
      grade: null,
    };
  }

  const text = String(rawValue).trim();
  if (!text) {
    return { marks_obtained: null, grade: null };
  }

  const normalized = text.toUpperCase();
  if (SPECIAL_MARK_CODES.has(normalized)) {
    return { marks_obtained: null, grade: normalized };
  }

  const asNumber = Number(text);
  if (Number.isFinite(asNumber)) {
    return {
      marks_obtained: Math.max(0, Math.min(100, asNumber)),
      grade: null,
    };
  }

  return { marks_obtained: null, grade: normalized };
};

const markToDisplayValue = (markRow) => {
  if (!markRow) return "";
  if (markRow.marks_obtained !== null && markRow.marks_obtained !== undefined) {
    const value = Number(markRow.marks_obtained);
    return Number.isFinite(value) ? value : "";
  }
  return safeValue(markRow.grade);
};

const isIncompleteMark = (value) => {
  const v = safeValue(value).toUpperCase();
  return v === "" || SPECIAL_MARK_CODES.has(v);
};

const getTermKeys = () => ["first", "second", "third", "annual"];

const calculateTermStats = (termValues, subjectConfigs, className) => {
  const visibleSubjects = subjectConfigs.filter(
    (sub) => !shouldExcludeSubject(sub.subject, className)
  );

  let totalObtained = 0;
  let hasIncomplete = false;

  visibleSubjects.forEach((subject) => {
    const value = termValues[subject.subject];
    const normalized = safeValue(value).toUpperCase();

    if (isIncompleteMark(normalized)) {
      hasIncomplete = true;
      return;
    }

    const asNumber = Number(value);
    if (Number.isFinite(asNumber)) {
      totalObtained += asNumber;
    }
  });

  const computedFullMarks =
    visibleSubjects.reduce(
      (sum, subject) => sum + (Number(subject.fullMarks) || 100),
      0
    ) || 1;
  const totalFullMarks = isLowerClassName(className)
    ? 600
    : computedFullMarks;

  const isClass1To8 =
    /^CLASS[-\s]?[1-8]$/i.test(className) || /^[1-8]$/i.test(className);
  const numValidSubjects = visibleSubjects.length || 1;

  const percentage = isClass1To8
    ? (totalObtained / numValidSubjects).toFixed(2)
    : ((totalObtained / totalFullMarks) * 100).toFixed(2);

  let division;
  if (hasIncomplete) division = "INCOMPLETE";
  else if (Number(percentage) >= 60) division = "First";
  else if (Number(percentage) >= 45) division = "Second";
  else if (Number(percentage) >= 30) division = "Third";
  else division = "Fail";

  return {
    totalObtained,
    totalFullMarks,
    percentage,
    division,
    visibleSubjects,
  };
};

async function ensureExamRows() {
  if (!isSupabaseEnabled()) return [];

  const requestedCodes = ["1st", "2nd", "3rd", "annual"];
  const { data: existing, error: existingError } = await supabase
    .from("exams")
    .select("id, exam_code")
    .in("exam_code", requestedCodes);

  if (existingError) throw existingError;

  const existingCodes = new Set((existing || []).map((exam) => exam.exam_code));
  const missingCodes = requestedCodes.filter((code) => !existingCodes.has(code));

  if (missingCodes.length) {
    const rows = missingCodes.map((examCode) => ({
      exam_code: examCode,
      exam_name: EXAM_NAME_BY_CODE[examCode] || examCode,
      academic_session: DEFAULT_ACADEMIC_SESSION,
      is_active: true,
    }));

    const { error: insertError } = await supabase.from("exams").insert(rows);
    if (insertError) throw insertError;
  }

  const { data: refreshed, error: refreshedError } = await supabase
    .from("exams")
    .select("id, exam_code, exam_name, academic_session")
    .in("exam_code", requestedCodes);

  if (refreshedError) throw refreshedError;
  return refreshed || [];
}

async function getClassRow(className, createIfMissing = false) {
  if (!isSupabaseEnabled()) return null;

  const canonicalClass = canonicalizeClass(className);
  const classCode = normalizeClassToken(canonicalClass);

  const { data: existing, error: existingError } = await supabase
    .from("classes")
    .select("id, class_name, class_code")
    .eq("class_code", classCode)
    .maybeSingle();

  if (existingError) throw existingError;
  if (existing) return existing;
  if (!createIfMissing) return null;

  const { data: created, error: createdError } = await supabase
    .from("classes")
    .insert({
      class_name: canonicalClass,
      class_code: classCode,
    })
    .select("id, class_name, class_code")
    .single();

  if (createdError) throw createdError;
  return created;
}

async function ensureSubjectRow(subjectConfig) {
  if (!isSupabaseEnabled()) return null;

  const { subjectName, subjectCode, fullMarks, passMarks } = subjectConfig;
  const code = toDbSubjectCode(subjectCode || subjectName);

  const { data: existing, error: existingError } = await supabase
    .from("subjects")
    .select("id, subject_name, subject_code, full_marks, pass_marks")
    .eq("subject_code", code)
    .maybeSingle();

  if (existingError) throw existingError;
  if (existing) return existing;

  const { data: created, error: createdError } = await supabase
    .from("subjects")
    .insert({
      subject_name: subjectName,
      subject_code: code,
      full_marks: Number(fullMarks) || 100,
      pass_marks: Number(passMarks) || 30,
    })
    .select("id, subject_name, subject_code, full_marks, pass_marks")
    .single();

  if (createdError) throw createdError;
  return created;
}

async function getSubjectsForClass(className, createIfMissing = false) {
  if (!isSupabaseEnabled()) return [];

  const classRow = await getClassRow(className, createIfMissing);
  if (!classRow) return [];

  const fetchSubjects = async () => {
    const { data, error } = await supabase
      .from("class_subjects")
      .select(
        "display_order, subjects:subject_id(id, subject_name, subject_code, full_marks, pass_marks)"
      )
      .eq("class_id", classRow.id)
      .order("display_order", { ascending: true });

    if (error) throw error;

    return (data || [])
      .filter((row) => row.subjects)
      .map((row) => ({
        subject: row.subjects.subject_name,
        subjectCode: row.subjects.subject_code,
        fullMarks: Number(row.subjects.full_marks) || 100,
        passMarks: Number(row.subjects.pass_marks) || 30,
        displayOrder: Number(row.display_order) || 1,
        subjectId: row.subjects.id,
      }));
  };

  let classSubjects = await fetchSubjects();
  if (classSubjects.length || !createIfMissing) {
    return classSubjects;
  }

  const defaults = getDefaultSubjectsForClass(className);
  for (const subjectConfig of defaults) {
    const subjectRow = await ensureSubjectRow(subjectConfig);
    const { error: linkError } = await supabase.from("class_subjects").upsert(
      {
        class_id: classRow.id,
        subject_id: subjectRow.id,
        display_order: subjectConfig.displayOrder,
      },
      { onConflict: "class_id,subject_id" }
    );
    if (linkError) throw linkError;
  }

  classSubjects = await fetchSubjects();
  return classSubjects;
}

async function getStudentRow(classId, roll) {
  if (!isSupabaseEnabled()) return null;

  const { data, error } = await supabase
    .from("students")
    .select(
      "id, class_id, roll_no, full_name, father_name, address, phone, dob, gender, created_at, updated_at"
    )
    .eq("class_id", classId)
    .eq("roll_no", safeValue(roll))
    .maybeSingle();

  if (error) throw error;
  return data;
}

async function upsertStudentInDb(payload) {
  if (!isSupabaseEnabled()) return null;

  const classRow = await getClassRow(payload.className, true);
  const rollNo = safeValue(payload.roll);
  if (!rollNo) {
    throw new Error("Roll number is required.");
  }

  const existing = await getStudentRow(classRow.id, rollNo);

  const data = {
    class_id: classRow.id,
    roll_no: rollNo,
    full_name: safeValue(payload.name),
    father_name: safeValue(payload.fatherName) || null,
    address: safeValue(payload.address),
    phone: safeValue(payload.phone),
    gender: safeValue(payload.gender),
    dob: safeValue(payload.dob) || null,
  };

  if (!data.full_name) {
    throw new Error("Student name is required.");
  }

  if (existing) {
    const { data: updated, error: updateError } = await supabase
      .from("students")
      .update(data)
      .eq("id", existing.id)
      .select(
        "id, class_id, roll_no, full_name, father_name, address, phone, dob, gender, created_at, updated_at"
      )
      .single();

    if (updateError) throw updateError;
    return { student: updated, classRow };
  }

  const { data: inserted, error: insertError } = await supabase
    .from("students")
    .insert(data)
    .select(
      "id, class_id, roll_no, full_name, father_name, address, phone, dob, gender, created_at, updated_at"
    )
    .single();

  if (insertError) throw insertError;
  return { student: inserted, classRow };
}

async function getOrCreateExam(examCode) {
  if (!isSupabaseEnabled()) return null;

  const normalizedCode = normalizeExamCode(examCode);
  if (!normalizedCode) return null;

  const exams = await ensureExamRows();
  return exams.find((exam) => exam.exam_code === normalizedCode) || null;
}

async function getMarksForStudentExam(studentId, examId) {
  if (!isSupabaseEnabled()) return [];

  const { data, error } = await supabase
    .from("marks")
    .select("id, student_id, exam_id, subject_id, marks_obtained, grade, remark")
    .eq("student_id", studentId)
    .eq("exam_id", examId);

  if (error) throw error;
  return data || [];
}

async function upsertMarksForStudent({
  className,
  roll,
  examCode,
  marks,
  studentName,
  fatherName,
}) {
  if (!isSupabaseEnabled()) {
    throw new Error("Supabase is not configured.");
  }

  const classRow = await getClassRow(className, true);
  let student = await getStudentRow(classRow.id, roll);

  if (!student) {
    const created = await upsertStudentInDb({
      className,
      roll,
      name: studentName,
      fatherName,
    });
    student = created.student;
  } else if (safeValue(studentName) || safeValue(fatherName)) {
    const update = await upsertStudentInDb({
      className,
      roll,
      name: safeValue(studentName) || student.full_name,
      fatherName: safeValue(fatherName) || student.father_name,
      address: student.address,
      phone: student.phone,
      gender: student.gender,
      dob: student.dob,
    });
    student = update.student;
  }

  const exam = await getOrCreateExam(examCode);
  if (!exam) {
    throw new Error("Invalid exam code.");
  }

  const subjects = await getSubjectsForClass(className, true);
  if (!subjects.length) {
    throw new Error("No subjects configured for this class.");
  }

  const subjectByCode = new Map(subjects.map((sub) => [sub.subjectCode, sub]));
  const subjectByNameToken = new Map(
    subjects.map((sub) => [normalizeSubjectToken(sub.subject), sub])
  );

  const rowsToUpsert = [];
  const subjectIdsToDelete = [];

  for (const markInput of marks || []) {
    const rawCode = toDbSubjectCode(markInput.subjectCode || markInput.subject || "");
    const byCode = subjectByCode.get(rawCode);
    const byName = subjectByNameToken.get(
      normalizeSubjectToken(markInput.subjectName || markInput.subject || "")
    );
    const subject = byCode || byName;
    if (!subject) continue;

    const parsedMark = parseMarkValue(markInput.value);
    if (
      parsedMark.marks_obtained === null &&
      safeValue(parsedMark.grade) === ""
    ) {
      subjectIdsToDelete.push(subject.subjectId);
      continue;
    }

    rowsToUpsert.push({
      student_id: student.id,
      exam_id: exam.id,
      subject_id: subject.subjectId,
      marks_obtained: parsedMark.marks_obtained,
      grade: parsedMark.grade,
      remark: safeValue(markInput.remark),
    });
  }

  if (subjectIdsToDelete.length) {
    const { error: deleteError } = await supabase
      .from("marks")
      .delete()
      .eq("student_id", student.id)
      .eq("exam_id", exam.id)
      .in("subject_id", subjectIdsToDelete);

    if (deleteError) throw deleteError;
  }

  if (rowsToUpsert.length) {
    const { error: upsertError } = await supabase.from("marks").upsert(rowsToUpsert, {
      onConflict: "student_id,exam_id,subject_id",
    });

    if (upsertError) throw upsertError;
  }

  return { student, classRow, exam };
}

async function getOrCreatePcNumber(studentId) {
  if (!isSupabaseEnabled()) return "";

  const { data: existing, error: existingError } = await supabase
    .from("provisional_certificates")
    .select("pc_no")
    .eq("student_id", studentId)
    .maybeSingle();

  if (existingError) throw existingError;
  if (existing?.pc_no) return existing.pc_no;

  const { data: allRows, error: allRowsError } = await supabase
    .from("provisional_certificates")
    .select("pc_no");

  if (allRowsError) throw allRowsError;

  const existingPCs = (allRows || [])
    .map((row) => safeValue(row.pc_no))
    .filter(Boolean);

  for (let attempt = 0; attempt < 10; attempt += 1) {
    const pcNo = generatePCNumber(existingPCs);

    const { data: inserted, error: insertError } = await supabase
      .from("provisional_certificates")
      .insert({
        student_id: studentId,
        pc_no: pcNo,
      })
      .select("pc_no")
      .single();

    if (!insertError) {
      return inserted.pc_no;
    }

    if (insertError.code !== "23505") {
      throw insertError;
    }
  }

  throw new Error("Unable to generate provisional certificate number.");
}

async function buildResultFromSupabase({ className, roll, terminal }) {
  if (!isSupabaseEnabled()) {
    return { status: "disabled" };
  }

  const classRow = await getClassRow(className, false);
  if (!classRow) {
    return { status: "not_found" };
  }

  const student = await getStudentRow(classRow.id, roll);
  if (!student) {
    return { status: "not_found" };
  }

  const subjectConfigs = await getSubjectsForClass(className, true);
  const exams = await ensureExamRows();

  const examIdToCode = new Map(exams.map((exam) => [exam.id, exam.exam_code]));
  const examIds = exams.map((exam) => exam.id);
  const subjectById = new Map(subjectConfigs.map((subject) => [subject.subjectId, subject]));

  const { data: marksRows, error: marksError } = await supabase
    .from("marks")
    .select("subject_id, exam_id, marks_obtained, grade")
    .eq("student_id", student.id)
    .in("exam_id", examIds);

  if (marksError) throw marksError;

  const termValues = {
    "1st": {},
    "2nd": {},
    "3rd": {},
    annual: {},
  };

  (marksRows || []).forEach((markRow) => {
    const examCode = examIdToCode.get(markRow.exam_id);
    const subject = subjectById.get(markRow.subject_id);
    if (!examCode || !subject) return;

    termValues[examCode][subject.subject] = markToDisplayValue(markRow);
  });

  const visibleSubjects = subjectConfigs.filter(
    (subject) => !shouldExcludeSubject(subject.subject, className)
  );

  const marks = visibleSubjects.map((subject) => {
    const getValue = (examCode) => {
      const value = termValues[examCode]?.[subject.subject];
      if (value === undefined || value === null) return "";
      return value;
    };

    return {
      subject: subject.subject,
      fullMarks: subject.fullMarks,
      passMarks: subject.passMarks,
      firstTerm: ["1st", "2nd", "3rd", "annual"].includes(terminal)
        ? getValue("1st")
        : 0,
      secondTerm: ["2nd", "3rd", "annual"].includes(terminal)
        ? getValue("2nd")
        : 0,
      thirdTerm: ["3rd", "annual"].includes(terminal)
        ? getValue("3rd")
        : 0,
      AnnTerm: terminal === "annual" ? getValue("annual") : 0,
    };
  });

  const totals = {
    first: calculateTermStats(termValues["1st"], subjectConfigs, className).totalObtained,
    second: calculateTermStats(termValues["2nd"], subjectConfigs, className).totalObtained,
    third: calculateTermStats(termValues["3rd"], subjectConfigs, className).totalObtained,
    annual: calculateTermStats(termValues.annual, subjectConfigs, className).totalObtained,
  };

  const statsByKey = {
    first: calculateTermStats(termValues["1st"], subjectConfigs, className),
    second: calculateTermStats(termValues["2nd"], subjectConfigs, className),
    third: calculateTermStats(termValues["3rd"], subjectConfigs, className),
    annual: calculateTermStats(termValues.annual, subjectConfigs, className),
  };

  const selectedTermValues = termValues[terminal] || {};
  const resultAvailable = visibleSubjects.some((subject) => {
    const value = selectedTermValues[subject.subject];
    return !isIncompleteMark(value);
  });

  if (!resultAvailable) {
    return { status: "found", error: "Result not available." };
  }

  return {
    status: "found",
    payload: {
      schoolName: SCHOOL_NAME,
      schoolAddress: SCHOOL_ADDRESS,
      studentName: safeValue(student.full_name),
      fatherName: safeValue(student.father_name),
      class: safeValue(classRow.class_name),
      roll: safeValue(student.roll_no),
      terminal,
      session: DEFAULT_ACADEMIC_SESSION,
      marks,
      totals,
      totalFullMarks: statsByKey.annual.totalFullMarks,
      percentageFirst: statsByKey.first.percentage,
      percentageSecond: statsByKey.second.percentage,
      percentageThird: statsByKey.third.percentage,
      percentageAnnual: statsByKey.annual.percentage,
      division: {
        first: statsByKey.first.division,
        second: statsByKey.second.division,
        third: statsByKey.third.division,
        annual: statsByKey.annual.division,
      },
      description:
        Object.values({
          first: statsByKey.first.division,
          second: statsByKey.second.division,
          third: statsByKey.third.division,
          annual: statsByKey.annual.division,
        }).includes("Fail")
          ? "Needs Improvement."
          : "Keep up the good work!",
    },
  };
}

async function buildProvisionalFromSupabase({ className, roll }) {
  if (!isSupabaseEnabled()) return { status: "disabled" };

  const classRow = await getClassRow(className, false);
  if (!classRow) return { status: "not_found" };

  const student = await getStudentRow(classRow.id, roll);
  if (!student) return { status: "not_found" };

  const annualExam = await getOrCreateExam("annual");
  if (!annualExam) return { status: "not_found" };

  const subjectConfigs = await getSubjectsForClass(className, true);
  const annualMarksRows = await getMarksForStudentExam(student.id, annualExam.id);
  const subjectById = new Map(subjectConfigs.map((subject) => [subject.subjectId, subject]));

  const annualValues = {};
  annualMarksRows.forEach((markRow) => {
    const subject = subjectById.get(markRow.subject_id);
    if (!subject) return;
    annualValues[subject.subject] = markToDisplayValue(markRow);
  });

  const annualStats = calculateTermStats(annualValues, subjectConfigs, className);
  const pcNo = await getOrCreatePcNumber(student.id);

  return {
    status: "found",
    payload: {
      studentName: safeValue(student.full_name),
      fatherName: safeValue(student.father_name),
      schoolName: `${SCHOOL_NAME}, MATHIA`,
      class: safeValue(classRow.class_name),
      rollNo: safeValue(student.roll_no),
      year: `Annual Exam ${annualExam.academic_session || DEFAULT_ACADEMIC_SESSION}`,
      percentageAnnual: annualStats.percentage,
      division: annualStats.division,
      date: new Date().toLocaleDateString("en-GB"),
      pcNo,
    },
  };
}

const calculatePercentageAndDivision = (
  studentRow,
  subjects,
  isLowerClass
) => {
  if (!studentRow) {
    return { percentage: "0.00", division: "INCOMPLETE" };
  }

  let totalObtained = 0;
  let hasIncomplete = false;

  const validSubjects = subjects.filter(sub => {
    const name = sub.trim().toUpperCase();
    if (name.includes("DRAWING")) return false;
    if (isLowerClass && isScienceOrSstSubject(sub))
      return false;
    return true;
  });

  validSubjects.forEach(sub => {
    const val = String(studentRow[sub] || "").trim().toUpperCase();

    if (["", "AB", "-", "NA"].includes(val)) {
      hasIncomplete = true;
      return;
    }

    if (!isNaN(val)) {
      totalObtained += Number(val);
    }
  });

  const totalFullMarks = validSubjects.length * 100 || 1;
  const percentage = ((totalObtained / totalFullMarks) * 100).toFixed(2);

  let division;
  if (hasIncomplete) division = "INCOMPLETE";
  else if (percentage >= 60) division = "First";
  else if (percentage >= 45) division = "Second";
  else if (percentage >= 30) division = "Third";
  else division = "Fail";

  return { percentage, division };
};

app.get('/result', async (req, res) => {
  const queryClass = canonicalizeClass(req.query.class || "");
  const roll = req.query.roll?.trim();
  const terminal = req.query.terminal?.trim().toLowerCase();

  if (!queryClass || !roll)
    return res.json({ error: "Class and Roll number are required." });
  if (!terminal)
    return res.json({ error: "Please select terminal." });

  try {
    const dbResult = await buildResultFromSupabase({
      className: queryClass,
      roll,
      terminal,
    });

    if (dbResult.status === "found") {
      if (dbResult.error) {
        return res.json({ error: dbResult.error });
      }
      return res.json(dbResult.payload);
    }
  } catch (dbErr) {
    console.error("Supabase result read failed. Falling back to Excel.", dbErr);
  }

  const sheets = {
    first: (await readSheet("result_1st")) || [],
    second: (await readSheet("result_2nd")) || [],
    third: (await readSheet("result_3rd")) || [],
    annual: (await readSheet("result_annual")) || [],
  };

  const findStudent = (sheet) =>
    sheet.find(
      (s) =>
        canonicalizeClass(s.Class) === queryClass &&
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
  ];

  const nonSubjectKeys = [
    "Class",
    "Roll",
    "Name",
    "FatherName",
    "Father Name",
    "P_C NO",
    "pcNO",
  ];

  const subjects = allKeys.reduce((unique, sub) => {
    if (nonSubjectKeys.includes(sub)) return unique;
    const n = normalize(sub);
    if (!unique.some((u) => normalize(u) === n)) unique.push(sub);
    return unique;
  }, []);

  const isLowerClass = isLowerClassName(queryClass);

  const isExcludedSubject = (sub) => {
    const name = normalize(String(sub || ""));
    if (!name) return true;
    if (name.includes("drawing")) return true;
    if (isLowerClass && isScienceOrSstSubject(sub))
      return true;
    return false;
  };

  // Exclude hidden subjects from rendering and calculations using one shared rule.
  const visibleSubjects = subjects.filter((sub) => !isExcludedSubject(sub));

  const marks = visibleSubjects.map((sub) => {
    let fullMarks = 100;
    let passMarks = 30;

    const getVal = (termData) => {
      if (!termData) return "";
      const val = termData[sub];
      if (val === undefined || val === null) return "";
      if (typeof val === "string") return val.trim();
      return val;
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

  // âœ… Total Calculation â€” exclude Drawing (count only numeric values)
  const calcTotal = (sheetData) => {
    if (!sheetData) return 0;
    return Object.keys(sheetData)
      .filter(
        (k) => !nonSubjectKeys.includes(k) && !isExcludedSubject(k)
      )
      .reduce(
        (sum, k) => {
          const val = sheetData[k];
          return sum + (isNumeric(val) ? parseFloat(val) : 0);
        },
        0
      );
  };

  const totals = {
    first: calcTotal(data.first),
    second: calcTotal(data.second),
    third: calcTotal(data.third),
    annual: calcTotal(data.annual),
  };

  // âœ… Total full marks calculation
  const totalFullMarks = isLowerClass ? 600 : visibleSubjects.length * 100;

  // âœ… Percentage logic
  // For classes 1-8: obtained / numSubjects * 100 (where numSubjects = 8 typically)
  const isClass1To8 = /^CLASS[-\s]?[1-8]$/i.test(queryClass) || /^[1-8]$/i.test(queryClass);
  const numValidSubjects = visibleSubjects.length;
  
  const termKeys = ["first", "second", "third", "annual"];
  const percentages = {};
  termKeys.forEach((term) => {
    const totalObtained = totals[term] || 0;
    let percentage;
    
    if (isClass1To8 && numValidSubjects > 0) {
      // For class 1-8: obtained / numSubjects * 100
      percentage = ((totalObtained / numValidSubjects)).toFixed(2);
    } else {
      // For other classes: standard calculation
      const totalFull = totalFullMarks || 1;
      percentage = ((totalObtained / totalFull) * 100).toFixed(2);
    }
    percentages[term] = percentage;
  });

  // âœ… Division
  const division = {};
  termKeys.forEach((k) => {
    const termData = data[k];
    const perc = parseFloat(percentages[k] || 0);
    let hasIncomplete = false;

    if (termData) {
      Object.keys(termData).forEach((key) => {
        if (nonSubjectKeys.includes(key) || isExcludedSubject(key))
          return;
        const val = String(termData[key] || "").trim().toUpperCase();
        if (["", "AB", "-", "NA"].includes(val)) hasIncomplete = true;
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
  terminal === "annual" ? DEFAULT_ACADEMIC_SESSION : "",

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

app.get('/provisional', async (req, res) => {
  let queryClass = canonicalizeClass(req.query.class || "");
  const roll = (req.query.roll || "").trim();

  if (!queryClass || !roll) {
    return res.json({ error: "Class and Roll required." });
  }

  try {
    const dbResult = await buildProvisionalFromSupabase({
      className: queryClass,
      roll,
    });

    if (dbResult.status === "found") {
      return res.json(dbResult.payload);
    }
  } catch (dbErr) {
    console.error(
      "Supabase provisional read failed. Falling back to Excel.",
      dbErr
    );
  }

  // Read Excel using exceljs
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile("results.xlsx");
  const sheetName = "result_annual";
  const worksheet = workbook.getWorksheet(sheetName);
  const headerRow = worksheet.getRow(1);
  const headers = headerRow.values.slice(1);
  const sheet = [];
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const rowObj = {};
    row.values.slice(1).forEach((v, i) => {
      rowObj[headers[i]] = v;
    });
    sheet.push(rowObj);
  });
  sanitizeCellObject(sheet);

  // Find student row with canonical class matching.
  const student = sheet.find(
    s => canonicalizeClass(s.Class) === queryClass && String(s.Roll).trim() === roll
  );

  if (!student) return res.json({ error: "Student not found." });

  // ðŸ”¹ PC number logic
  // Check if Excel already has a PC number
  let existingPC = student["P_C NO"] || student["P C NO"] || student["PC NO"];
  if (!existingPC || existingPC.trim() === "") {
    const existingPCs = sheet
      .map(r => r["P_C NO"] || r["P C NO"] || r["PC NO"] || "")
      .filter(Boolean);

    const newPC = generatePCNumber(existingPCs);

    // Save PC in Excel via exceljs
    // find the actual row index in worksheet
    let targetRow;
    worksheet.eachRow((row, rowNumber) => {
      if (
        canonicalizeClass(row.getCell(headers.indexOf("Class") + 1).value) === queryClass &&
        String(row.getCell(headers.indexOf("Roll") + 1).value).trim() === roll
      ) {
        targetRow = row;
      }
    });
    if (targetRow) {
      const pcColIndex = headers.indexOf("P_C NO") + 1 || headers.indexOf("P C NO") + 1 || headers.indexOf("PC NO") + 1;
      if (pcColIndex) {
        targetRow.getCell(pcColIndex).value = newPC;
      }
      await workbook.xlsx.writeFile("results.xlsx");
    }

    existingPC = newPC; // use for response
  }
  // ---------- SAME LOGIC AS /result ----------

  const isLowerClass = isLowerClassName(queryClass);

  const ignoreKeys = [
    "Class","Roll","Name","Father Name","FatherName",
    "P_C NO","P C NO","PC NO","Division","School Name","Year"
  ];

  const subjects = Object.keys(student).filter(k => !ignoreKeys.includes(k));

  let totalObtained = 0;
  let hasIncomplete = false;

  const validSubjects = subjects.filter(sub => {
    const name = sub.trim().toUpperCase();
    if (name.includes("DRAWING")) return false;
    if (isLowerClass && isScienceOrSstSubject(sub))
      return false;
    return true;
  });

  validSubjects.forEach(sub => {
    const val = String(student[sub] || "").trim().toUpperCase();
    if (["", "AB", "-", "NA"].includes(val)) {
      hasIncomplete = true;
      return;
    }
    if (!isNaN(val)) {
      totalObtained += Number(val);
    }
  });

  const totalFullMarks = isLowerClass ? 600 : (validSubjects.length * 100 || 1);
  
  // âœ… For class 1-8: obtained / numSubjects * 100
  const isClass1To8 = /^CLASS[-\s]?[1-8]$/i.test(queryClass) || /^[1-8]$/i.test(queryClass);
  let percentageAnnual;
  
  if (isClass1To8 && validSubjects.length > 0) {
    percentageAnnual = ((totalObtained / validSubjects.length)).toFixed(2);
  } else {
    percentageAnnual = ((totalObtained / totalFullMarks) * 100).toFixed(2);
  }

  let division;
  if (hasIncomplete) division = "INCOMPLETE";
  else if (percentageAnnual >= 60) division = "First";
  else if (percentageAnnual >= 45) division = "Second";
  else if (percentageAnnual >= 30) division = "Third";
  else division = "Fail";


  // Response
 res.json({
  studentName: safeValue(student["Name"]),
  fatherName: safeValue(student["Father Name"]),
  schoolName: safeValue(student["School Name"] || "STAR PUBLIC SCHOOL, MATHIA"),
  class: safeValue(student["Class"]),
  rollNo: safeValue(student["Roll"]),
  year: safeValue(student["Year"] || "Annual Exam 2025-26"),
  percentageAnnual,
  division,
  date: new Date().toLocaleDateString("en-GB"),
  pcNo: existingPC
});

});


app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());

// Session setup
app.use(session({
    secret: process.env.JWT_SECRET || "admin_secret_key",
    resave: false,
    saveUninitialized: true,
    cookie: {
      httpOnly: true,
      sameSite: "lax",
      secure: false,
    },
}));

// Prevent caching for all routes (important after logout)
app.use((req, res, next) => {
    res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, proxy-revalidate");
    res.setHeader("Pragma", "no-cache");
    res.setHeader("Expires", "0");
    res.setHeader("Surrogate-Control", "no-store");
    next();
});

app.get("/dashboard.html", isAdminLoggedIn, (req, res) => {
  res.sendFile(path.join(__dirname, "public", "dashboard.html"));
});

app.get("/provisional.html", isAdminLoggedIn, (req, res) => {
  res.sendFile(path.join(__dirname, "public", "provisional.html"));
});

app.use(express.static("public"));

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

// ----------------- Nodemailer Setup -----------------

const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: Number(process.env.SMTP_PORT),
  secure: false,
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS
  }
});

transporter.verify((err) => {
  if (err) {
    console.error("SMTP VERIFY FAILED:", err);
  } else {
    console.log("SMTP READY âœ…");
  }
});

// ----------------- Helper -----------------
function isAdminLoggedIn(req, res, next) {
    if (req.session && req.session.admin) {
        return next();
    } else {
        return res.redirect("/admin.html");
    }
}

function isAdminApiAuthenticated(req, res, next) {
  if (req.session && req.session.admin) {
    return next();
  }
  return res.status(401).json({ error: "Unauthorized" });
}

function requireSupabase(req, res, next) {
  if (!isSupabaseEnabled()) {
    return res.status(500).json({
      error:
        "Supabase is not configured. Add SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY in .env.",
    });
  }
  return next();
}

// ----------------- Routes -----------------

app.get(
  "/api/admin/bootstrap",
  isAdminApiAuthenticated,
  requireSupabase,
  async (req, res) => {
    try {
      const exams = await ensureExamRows();
      const { data: classes, error: classError } = await supabase
        .from("classes")
        .select("id, class_name, class_code")
        .order("class_name", { ascending: true });

      if (classError) throw classError;

      const examOrder = ["1st", "2nd", "3rd", "annual"];
      const orderedExams = [...(exams || [])].sort(
        (a, b) =>
          examOrder.indexOf(a.exam_code) - examOrder.indexOf(b.exam_code)
      );

      const classesByToken = new Map();
      (classes || []).forEach((row) => {
        const normalizedClassName = canonicalizeClass(row.class_name);
        const token = normalizeClassToken(normalizedClassName);
        if (!token) return;

        if (!classesByToken.has(token)) {
          classesByToken.set(token, {
            id: row.id,
            className: normalizedClassName,
            classCode: token,
          });
        }
      });

      const cleanedClasses = Array.from(classesByToken.values()).sort((a, b) => {
        const rankDiff = classSortIndex(a.className) - classSortIndex(b.className);
        if (rankDiff !== 0) return rankDiff;
        return a.className.localeCompare(b.className);
      });

      return res.json({
        classes: cleanedClasses,
        exams: orderedExams.map((exam) => ({
          id: exam.id,
          examCode: exam.exam_code,
          examName: exam.exam_name,
          session: exam.academic_session,
        })),
      });
    } catch (error) {
      console.error("Failed to load admin bootstrap data:", error);
      return res.status(500).json({ error: "Failed to load dashboard data." });
    }
  }
);

app.get(
  "/api/admin/students",
  isAdminApiAuthenticated,
  requireSupabase,
  async (req, res) => {
    try {
      const className = safeValue(req.query.className || req.query.class);
      const search = safeValue(req.query.search).replace(/[%_,]/g, "");
      const pageRaw = Number.parseInt(req.query.page, 10);
      const page = Number.isFinite(pageRaw) && pageRaw > 0 ? pageRaw : 1;
      const pageSize = 30;

      let query = supabase
        .from("students")
        .select(
          "id, class_id, roll_no, full_name, father_name, address, phone, dob, gender, created_at, updated_at, classes:class_id(class_name, class_code)",
          { count: "exact" }
        );

      if (className) {
        const classRow = await getClassRow(className, false);
        if (!classRow) {
          return res.json({
            students: [],
            pagination: {
              page,
              pageSize,
              total: 0,
              totalPages: 0,
              hasPrev: false,
              hasNext: false,
            },
          });
        }
        query = query.eq("class_id", classRow.id);
      }

      if (search) {
        query = query.or(
          `roll_no.ilike.%${search}%,full_name.ilike.%${search}%`
        );
      }

      const { data, error, count } = await query;
      if (error) throw error;

      const sortedRows = (data || []).sort(compareStudentsForDirectory);
      const total = Number.isFinite(count) ? count : sortedRows.length;
      const totalPages = total === 0 ? 0 : Math.ceil(total / pageSize);
      const safePage = totalPages === 0 ? 1 : Math.min(page, totalPages);
      const from = (safePage - 1) * pageSize;
      const to = from + pageSize;
      const pagedRows = sortedRows.slice(from, to);

      return res.json({
        students: pagedRows.map((row) => ({
          id: row.id,
          roll: row.roll_no,
          name: row.full_name,
          fatherName: row.father_name || "",
          address: row.address || "",
          phone: row.phone || "",
          dob: row.dob || "",
          gender: row.gender || "",
          className: canonicalizeClass(row.classes?.class_name || ""),
          classCode: normalizeClassToken(row.classes?.class_code || ""),
        })),
        pagination: {
          page: safePage,
          pageSize,
          total,
          totalPages,
          hasPrev: safePage > 1,
          hasNext: safePage < totalPages,
        },
      });
    } catch (error) {
      console.error("Failed to load students:", error);
      return res.status(500).json({ error: "Failed to load students." });
    }
  }
);

app.post(
  "/api/admin/students",
  isAdminApiAuthenticated,
  requireSupabase,
  async (req, res) => {
    try {
      const payload = {
        className: safeValue(req.body.className || req.body.class),
        roll: safeValue(req.body.roll || req.body.rollNo),
        name: safeValue(req.body.name || req.body.studentName),
        fatherName: safeValue(req.body.fatherName),
        address: safeValue(req.body.address),
        phone: safeValue(req.body.phone),
        dob: safeValue(req.body.dob),
        gender: safeValue(req.body.gender),
      };

      if (!payload.className || !payload.roll || !payload.name) {
        return res
          .status(400)
          .json({ error: "Class, roll number and student name are required." });
      }

      const { student, classRow } = await upsertStudentInDb(payload);

      return res.json({
        success: true,
        student: {
          id: student.id,
          roll: student.roll_no,
          name: student.full_name,
          fatherName: student.father_name || "",
          address: student.address || "",
          phone: student.phone || "",
          dob: student.dob || "",
          gender: student.gender || "",
          className: classRow.class_name,
          classCode: classRow.class_code,
        },
      });
    } catch (error) {
      console.error("Failed to save student:", error);
      const message = safeValue(error.message) || "Failed to save student.";
      const status = /required|invalid/i.test(message) ? 400 : 500;
      return res.status(status).json({ error: message });
    }
  }
);

app.get(
  "/api/admin/subjects",
  isAdminApiAuthenticated,
  requireSupabase,
  async (req, res) => {
    try {
      const className = safeValue(req.query.className || req.query.class);
      if (!className) {
        return res.status(400).json({ error: "Class is required." });
      }

      const subjects = await getSubjectsForClass(className, true);
      return res.json({
        className: canonicalizeClass(className),
        subjects: subjects.map((subject) => ({
          subjectId: subject.subjectId,
          subject: subject.subject,
          subjectCode: subject.subjectCode,
          fullMarks: subject.fullMarks,
          passMarks: subject.passMarks,
          displayOrder: subject.displayOrder,
        })),
      });
    } catch (error) {
      console.error("Failed to load subjects:", error);
      return res.status(500).json({ error: "Failed to load subjects." });
    }
  }
);

app.get(
  "/api/admin/marks",
  isAdminApiAuthenticated,
  requireSupabase,
  async (req, res) => {
    try {
      const className = safeValue(req.query.className || req.query.class);
      const roll = safeValue(req.query.roll || req.query.rollNo);
      const examCode = normalizeExamCode(req.query.examCode || req.query.exam);

      if (!className || !roll || !examCode) {
        return res
          .status(400)
          .json({ error: "Class, roll number and exam are required." });
      }

      const classRow = await getClassRow(className, false);
      if (!classRow) {
        return res.status(404).json({ error: "Class not found." });
      }

      const student = await getStudentRow(classRow.id, roll);
      if (!student) {
        return res.status(404).json({ error: "Student not found." });
      }

      const exam = await getOrCreateExam(examCode);
      if (!exam) {
        return res.status(400).json({ error: "Invalid exam code." });
      }

      const subjects = await getSubjectsForClass(className, true);
      const marksRows = await getMarksForStudentExam(student.id, exam.id);
      const markBySubjectId = new Map(
        marksRows.map((markRow) => [markRow.subject_id, markRow])
      );

      return res.json({
        student: {
          id: student.id,
          roll: student.roll_no,
          name: student.full_name,
          fatherName: student.father_name || "",
          className: classRow.class_name,
        },
        exam: {
          id: exam.id,
          examCode: exam.exam_code,
          examName: exam.exam_name,
          session: exam.academic_session,
        },
        subjects: subjects.map((subject) => {
          const markRow = markBySubjectId.get(subject.subjectId);
          return {
            subjectId: subject.subjectId,
            subject: subject.subject,
            subjectCode: subject.subjectCode,
            fullMarks: subject.fullMarks,
            passMarks: subject.passMarks,
            value: markToDisplayValue(markRow),
            remark: safeValue(markRow?.remark),
          };
        }),
      });
    } catch (error) {
      console.error("Failed to load marks:", error);
      return res.status(500).json({ error: "Failed to load marks." });
    }
  }
);

app.post(
  "/api/admin/marks",
  isAdminApiAuthenticated,
  requireSupabase,
  async (req, res) => {
    try {
      const className = safeValue(req.body.className || req.body.class);
      const roll = safeValue(req.body.roll || req.body.rollNo);
      const examCode = normalizeExamCode(
        req.body.examCode || req.body.exam || req.body.terminal
      );
      const studentName = safeValue(req.body.studentName || req.body.name);
      const fatherName = safeValue(req.body.fatherName);
      const marks = Array.isArray(req.body.marks) ? req.body.marks : [];

      if (!className || !roll || !examCode) {
        return res
          .status(400)
          .json({ error: "Class, roll number and exam are required." });
      }

      if (!marks.length) {
        return res.status(400).json({ error: "At least one subject mark is required." });
      }

      const saveResult = await upsertMarksForStudent({
        className,
        roll,
        examCode,
        marks,
        studentName,
        fatherName,
      });

      const savedMarks = await getMarksForStudentExam(
        saveResult.student.id,
        saveResult.exam.id
      );
      const subjects = await getSubjectsForClass(className, true);
      const markBySubjectId = new Map(
        savedMarks.map((markRow) => [markRow.subject_id, markRow])
      );

      return res.json({
        success: true,
        message: "Marks saved successfully.",
        student: {
          id: saveResult.student.id,
          roll: saveResult.student.roll_no,
          name: saveResult.student.full_name,
          fatherName: saveResult.student.father_name || "",
          className: saveResult.classRow.class_name,
        },
        exam: {
          id: saveResult.exam.id,
          examCode: saveResult.exam.exam_code,
          examName: saveResult.exam.exam_name,
          session: saveResult.exam.academic_session,
        },
        subjects: subjects.map((subject) => ({
          subjectId: subject.subjectId,
          subject: subject.subject,
          subjectCode: subject.subjectCode,
          fullMarks: subject.fullMarks,
          passMarks: subject.passMarks,
          value: markToDisplayValue(markBySubjectId.get(subject.subjectId)),
          remark: safeValue(markBySubjectId.get(subject.subjectId)?.remark),
        })),
      });
    } catch (error) {
      console.error("Failed to save marks:", error);
      const message =
        safeValue(error.message) || "Unable to save marks right now.";
      const status = /required|invalid|not configured/i.test(message) ? 400 : 500;
      return res.status(status).json({ error: message });
    }
  }
);

// Admin login
app.post("/admin/login", async (req, res) => {
    const { email, password } = req.body;
    const loginEmail = safeValue(email).toLowerCase();

    if (!loginEmail || !safeValue(password)) {
        return res.send("<h3>Email and password are required. <a href='/admin.html'>Try again</a></h3>");
    }

    if (loginEmail !== admin.email.toLowerCase()) {
        return res.send("<h3>Invalid email. <a href='/admin.html'>Try again</a></h3>");
    }

    const match = await bcrypt.compare(password, admin.passwordHash);
    if (!match) {
        return res.send("<h3>Invalid password. <a href='/admin.html'>Try again</a></h3>");
    }

    req.session.admin = true;

    res.redirect("/dashboard.html");
});

// Admin logout
app.get("/admin/logout", (req, res) => {
    req.session.destroy(err => {
        if (err) return res.send("Error logging out.");
        res.redirect("/admin.html");
    });
});

app.post("/admin/logout", (req, res) => {
  req.session.destroy((err) => {
    if (err) {
      return res.status(500).json({ error: "Error logging out." });
    }
    return res.json({ success: true });
  });
});

// Request OTP for password reset

const otpStore = {}; // server start me ek baar

app.post("/admin/request-otp", async (req, res) => {
  try {
    const { email } = req.body;
    const normalizedEmail = email.toLowerCase();

    if (normalizedEmail !== admin.email.toLowerCase()) {
      return res.json({ error: "Invalid email." });
    }

    // generate OTP
    const otp = Math.floor(100000 + Math.random() * 900000).toString();
    const expires = Date.now() + 5 * 60 * 1000;

    otpStore[normalizedEmail] = { otp, expires };

    // âœ… SEND OTP USING MAILJET API (NO SMTP)
    await mailjet
      .post("send", { version: "v3.1" })
      .request({
        Messages: [
          {
            From: {
              Email: "t01auheed@gmail.com", // VERIFIED MAILJET EMAIL
              Name: "Star Public School"
            },
            To: [
              {
                Email: normalizedEmail
              }
            ],
            Subject: "Star Public School â€“ Admin Verification Code",
            TextPart: `Your OTP is ${otp}. Valid for 5 minutes.`,
            HTMLPart: `
              <h3>Star Public School</h3>
              <p>Your admin verification code is:</p>
              <h1>${otp}</h1>
              <p>This code is valid for 5 minutes.</p>
            `
          }
        ]
      });

    console.log("OTP SENT:", otp); // debug
    res.json({ success: true, message: "OTP sent successfully" });

  } catch (err) {
    console.error("OTP SEND FAIL:", err);
    res.status(500).json({ error: "Failed to send OTP" });
  }
});



// Reset password using OTP
app.post("/admin/reset-password", async (req, res) => {
  try {
    const { email, otp, newPassword } = req.body;
    const normalizedEmail = email.toLowerCase();

    const record = otpStore[normalizedEmail];
    if (!record) return res.json({ error: "OTP not requested" });
    if (Date.now() > record.expires) return res.json({ error: "OTP expired" });
    if (record.otp !== otp) return res.json({ error: "Invalid OTP" });

    admin.passwordHash = await bcrypt.hash(newPassword, 10);
    fs.writeFileSync(adminFile, JSON.stringify(admin, null, 2));

    delete otpStore[normalizedEmail];

    res.json({ success: true, message: "Password updated successfully" });

  } catch (err) {
    console.error("RESET ERROR:", err);
    res.status(500).json({ error: "Server error" });
  }
});

// Protected dashboard
app.get("/admin/dashboard", isAdminLoggedIn, (req, res) => {
    res.redirect("/dashboard.html");
});

app.listen(PORT, () => console.log(`âœ… Server running at http://localhost:${PORT}`));
