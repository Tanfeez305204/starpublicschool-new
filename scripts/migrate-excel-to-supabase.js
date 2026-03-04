require("dotenv").config();

const path = require("path");
const Excel = require("exceljs");
const { createClient } = require("@supabase/supabase-js");

const RESULTS_FILE =
  process.env.RESULTS_FILE || path.join(__dirname, "..", "results.xlsx");
const ACADEMIC_SESSION = process.env.ACADEMIC_SESSION || "2025-26";
const DRY_RUN = process.argv.includes("--dry-run");

const SHEET_EXAM_MAP = {
  result_1st: { examCode: "1st", examName: "First Terminal" },
  result_2nd: { examCode: "2nd", examName: "Second Terminal" },
  result_3rd: { examCode: "3rd", examName: "Third Terminal" },
  result_annual: { examCode: "annual", examName: "Annual" },
};

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
};

const NON_SUBJECT_HEADERS = new Set([
  "class",
  "roll",
  "name",
  "fathername",
  "father name",
  "father_name",
  "p_c no",
  "p c no",
  "pc no",
  "pcno",
  "division",
  "school name",
  "year",
]);

const SPECIAL_MARK_CODES = new Set(["AB", "NA", "-"]);

const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

const safe = (value, fallback = "") => {
  if (value === undefined || value === null) return fallback;
  const text = String(value).trim();
  return text || fallback;
};

const normalizeClassToken = (cls = "") =>
  String(cls)
    .trim()
    .toUpperCase()
    .replace(/\s+/g, "")
    .replace(/\./g, "")
    .replace(/_/g, "")
    .replace(/-/g, "");

const canonicalizeClass = (cls = "") => {
  const token = normalizeClassToken(cls);
  return CLASS_ALIAS_MAP[token] || safe(cls).toUpperCase();
};

const normalizeHeader = (header = "") =>
  String(header).trim().toLowerCase().replace(/\s+/g, " ");

const toSubjectCode = (subjectName = "") =>
  String(subjectName)
    .trim()
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "");

const canonicalizeSubjectName = (header = "") => {
  const raw = safe(header);
  const token = raw.toUpperCase().replace(/[^A-Z0-9]/g, "");

  if (token === "SANSKRITURDU") return "Sanskrit / Urdu";
  if (token === "COMPUTERTABLE") return "Computer / Table";
  if (token === "SST" || token === "SOCIALSCIENCE") return "S.S.T";
  if (token === "MATHS") return "Math";

  return raw
    .replace(/\s*\/\s*/g, " / ")
    .replace(/\s+/g, " ")
    .replace(/^s\.s\.t$/i, "S.S.T");
};

const parseMarkCell = (value) => {
  if (value === undefined || value === null) return null;
  if (typeof value === "number" && Number.isFinite(value)) {
    return {
      marks_obtained: Math.max(0, Math.min(100, value)),
      grade: null,
    };
  }

  const text = safe(value);
  if (!text) return null;

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

const getWorkbookRows = (worksheet) => {
  const headerCells = worksheet.getRow(1).values.slice(1);
  const headers = headerCells.map((h) => safe(h));
  const rows = [];

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return;
    const obj = {};
    row.values.slice(1).forEach((value, idx) => {
      obj[headers[idx]] = value;
    });
    rows.push(obj);
  });

  return { headers, rows };
};

const chunk = (arr, size) => {
  const out = [];
  for (let idx = 0; idx < arr.length; idx += size) {
    out.push(arr.slice(idx, idx + size));
  }
  return out;
};

const isRetryableNetworkError = (error) => {
  const message = safe(error?.message);
  const causeMessage = safe(error?.cause?.message);
  const code = safe(error?.cause?.code || error?.code);
  const combined = `${message} ${causeMessage} ${code}`;
  return /fetch failed|UND_ERR_SOCKET|ECONNRESET|ETIMEDOUT|EAI_AGAIN|ENETUNREACH|ECONNREFUSED|EPIPE/i.test(
    combined
  );
};

const supabaseFetchWithRetry = async (input, init) => {
  const maxAttempts = 4;

  for (let attempt = 1; attempt <= maxAttempts; attempt += 1) {
    try {
      const response = await fetch(input, init);
      if (
        response.status >= 500 &&
        response.status < 600 &&
        attempt < maxAttempts
      ) {
        await wait(200 * attempt);
        continue;
      }
      return response;
    } catch (error) {
      if (!isRetryableNetworkError(error) || attempt >= maxAttempts) {
        throw error;
      }
      console.warn(
        `Transient Supabase network error (${safe(
          error.message
        )}) on attempt ${attempt}/${maxAttempts}. Retrying...`
      );
      await wait(200 * attempt);
    }
  }

  throw new Error("Supabase request failed after retries.");
};

const createSupabaseClient = () => {
  const supabaseUrl = safe(process.env.SUPABASE_URL);
  const serviceKey = safe(process.env.SUPABASE_SERVICE_ROLE_KEY);
  if (!supabaseUrl || !serviceKey) {
    throw new Error(
      "SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are required in .env"
    );
  }

  return createClient(supabaseUrl, serviceKey, {
    auth: { persistSession: false, autoRefreshToken: false },
    global: { fetch: supabaseFetchWithRetry },
  });
};

async function upsertInChunks({
  supabase,
  table,
  rows,
  onConflict,
  chunkSize = 500,
  label,
}) {
  if (!rows.length) return;
  const parts = chunk(rows, chunkSize);
  let done = 0;

  for (const part of parts) {
    const { error } = await supabase
      .from(table)
      .upsert(part, { onConflict, ignoreDuplicates: false });
    if (error) throw error;
    done += part.length;
    console.log(`  -> ${label}: ${done}/${rows.length}`);
  }
}

async function main() {
  console.log("Starting Excel -> Supabase migration");
  console.log(`File: ${RESULTS_FILE}`);
  console.log(`Session: ${ACADEMIC_SESSION}`);
  console.log(`Mode: ${DRY_RUN ? "DRY RUN" : "APPLY"}`);

  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile(RESULTS_FILE);

  const classesByCode = new Map();
  const subjectsByCode = new Map();
  const classSubjectOrder = new Map();
  const studentsByKey = new Map();
  const marksByKey = new Map();
  const provisionalByStudentKey = new Map();

  for (const [sheetName, examInfo] of Object.entries(SHEET_EXAM_MAP)) {
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) {
      console.warn(`Sheet not found: ${sheetName}. Skipping.`);
      continue;
    }

    const { headers, rows } = getWorkbookRows(worksheet);
    const subjectHeaders = headers.filter(
      (header) => !NON_SUBJECT_HEADERS.has(normalizeHeader(header))
    );
    console.log(
      `Reading ${sheetName}: ${rows.length} rows, ${subjectHeaders.length} subject columns`
    );

    rows.forEach((row) => {
      const className = canonicalizeClass(row["Class"]);
      const roll = safe(row["Roll"]);
      if (!className || !roll) return;

      const classCode = normalizeClassToken(className);
      classesByCode.set(classCode, { className, classCode });

      const studentKey = `${classCode}::${roll}`;
      const studentName = safe(row["Name"]);
      const fatherName = safe(row["Father Name"] || row["FatherName"]);

      const existingStudent = studentsByKey.get(studentKey) || {
        classCode,
        className,
        roll,
        fullName: "",
        fatherName: "",
      };

      studentsByKey.set(studentKey, {
        ...existingStudent,
        fullName: existingStudent.fullName || studentName,
        fatherName: existingStudent.fatherName || fatherName,
      });

      const pcNo = safe(row["P_C NO"] || row["P C NO"] || row["PC NO"] || row["pcNO"]);
      if (pcNo) {
        provisionalByStudentKey.set(studentKey, pcNo);
      }

      if (!classSubjectOrder.has(classCode)) {
        classSubjectOrder.set(classCode, new Map());
      }
      const classSubjectMap = classSubjectOrder.get(classCode);

      subjectHeaders.forEach((header) => {
        const subjectName = canonicalizeSubjectName(header);
        const subjectCode = toSubjectCode(subjectName);
        if (!subjectCode) return;

        if (!subjectsByCode.has(subjectCode)) {
          subjectsByCode.set(subjectCode, {
            subjectCode,
            subjectName,
            fullMarks: 100,
            passMarks: 30,
          });
        }

        if (!classSubjectMap.has(subjectCode)) {
          classSubjectMap.set(subjectCode, classSubjectMap.size + 1);
        }

        const parsed = parseMarkCell(row[header]);
        if (!parsed) return;

        const markCompositeKey = `${examInfo.examCode}::${studentKey}::${subjectCode}`;
        marksByKey.set(markCompositeKey, {
          examCode: examInfo.examCode,
          studentKey,
          subjectCode,
          marks_obtained: parsed.marks_obtained,
          grade: parsed.grade,
          remark: "",
        });
      });
    });
  }

  console.log("Extracted summary:");
  console.log(`  Classes: ${classesByCode.size}`);
  console.log(`  Subjects: ${subjectsByCode.size}`);
  console.log(`  Students: ${studentsByKey.size}`);
  console.log(`  Marks: ${marksByKey.size}`);
  console.log(`  Provisional PCs: ${provisionalByStudentKey.size}`);

  if (DRY_RUN) {
    console.log("Dry run complete. No DB changes applied.");
    return;
  }

  const supabase = createSupabaseClient();

  const examRows = Object.values(SHEET_EXAM_MAP).map((exam) => ({
    exam_code: exam.examCode,
    exam_name: exam.examName,
    academic_session: ACADEMIC_SESSION,
    is_active: true,
  }));
  await upsertInChunks({
    supabase,
    table: "exams",
    rows: examRows,
    onConflict: "exam_code",
    chunkSize: 50,
    label: "exams",
  });

  await upsertInChunks({
    supabase,
    table: "classes",
    rows: Array.from(classesByCode.values()).map((row) => ({
      class_name: row.className,
      class_code: row.classCode,
    })),
    onConflict: "class_code",
    chunkSize: 200,
    label: "classes",
  });

  await upsertInChunks({
    supabase,
    table: "subjects",
    rows: Array.from(subjectsByCode.values()).map((row) => ({
      subject_name: row.subjectName,
      subject_code: row.subjectCode,
      full_marks: row.fullMarks,
      pass_marks: row.passMarks,
    })),
    onConflict: "subject_code",
    chunkSize: 200,
    label: "subjects",
  });

  const [{ data: classRows, error: classRowsError }, { data: subjectRows, error: subjectRowsError }, { data: examRowsDb, error: examRowsDbError }] =
    await Promise.all([
      supabase.from("classes").select("id,class_name,class_code"),
      supabase.from("subjects").select("id,subject_code"),
      supabase.from("exams").select("id,exam_code").in(
        "exam_code",
        Object.values(SHEET_EXAM_MAP).map((x) => x.examCode)
      ),
    ]);

  if (classRowsError) throw classRowsError;
  if (subjectRowsError) throw subjectRowsError;
  if (examRowsDbError) throw examRowsDbError;

  const classIdByCode = new Map(
    (classRows || []).map((row) => [row.class_code, row.id])
  );
  const subjectIdByCode = new Map(
    (subjectRows || []).map((row) => [row.subject_code, row.id])
  );
  const examIdByCode = new Map(
    (examRowsDb || []).map((row) => [row.exam_code, row.id])
  );

  const classSubjectRows = [];
  classSubjectOrder.forEach((subjectOrderMap, classCode) => {
    const classId = classIdByCode.get(classCode);
    if (!classId) return;
    subjectOrderMap.forEach((displayOrder, subjectCode) => {
      const subjectId = subjectIdByCode.get(subjectCode);
      if (!subjectId) return;
      classSubjectRows.push({
        class_id: classId,
        subject_id: subjectId,
        display_order: displayOrder,
        is_active: true,
      });
    });
  });

  await upsertInChunks({
    supabase,
    table: "class_subjects",
    rows: classSubjectRows,
    onConflict: "class_id,subject_id",
    chunkSize: 500,
    label: "class_subjects",
  });

  const studentRowsToUpsert = Array.from(studentsByKey.values()).map((student) => {
    const classId = classIdByCode.get(student.classCode);
    if (!classId) return null;

    return {
      class_id: classId,
      roll_no: student.roll,
      full_name: safe(student.fullName, `Student ${student.roll}`),
      father_name: safe(student.fatherName) || null,
      is_active: true,
    };
  }).filter(Boolean);

  await upsertInChunks({
    supabase,
    table: "students",
    rows: studentRowsToUpsert,
    onConflict: "class_id,roll_no",
    chunkSize: 500,
    label: "students",
  });

  const classIds = Array.from(classIdByCode.values());
  const studentIdByStudentKey = new Map();
  if (classIds.length) {
    const { data: studentRowsDb, error: studentRowsDbError } = await supabase
      .from("students")
      .select("id,class_id,roll_no")
      .in("class_id", classIds);
    if (studentRowsDbError) throw studentRowsDbError;

    const classCodeById = new Map((classRows || []).map((row) => [row.id, row.class_code]));
    (studentRowsDb || []).forEach((row) => {
      const classCode = classCodeById.get(row.class_id);
      if (!classCode) return;
      const studentKey = `${classCode}::${safe(row.roll_no)}`;
      studentIdByStudentKey.set(studentKey, row.id);
    });
  }

  const markRowsToUpsert = [];
  marksByKey.forEach((mark) => {
    const studentId = studentIdByStudentKey.get(mark.studentKey);
    const examId = examIdByCode.get(mark.examCode);
    const subjectId = subjectIdByCode.get(mark.subjectCode);
    if (!studentId || !examId || !subjectId) return;

    markRowsToUpsert.push({
      student_id: studentId,
      exam_id: examId,
      subject_id: subjectId,
      marks_obtained: mark.marks_obtained,
      grade: mark.grade,
      remark: mark.remark,
    });
  });

  await upsertInChunks({
    supabase,
    table: "marks",
    rows: markRowsToUpsert,
    onConflict: "student_id,exam_id,subject_id",
    chunkSize: 1000,
    label: "marks",
  });

  const provisionalRows = [];
  provisionalByStudentKey.forEach((pcNo, studentKey) => {
    const studentId = studentIdByStudentKey.get(studentKey);
    if (!studentId) return;
    provisionalRows.push({
      student_id: studentId,
      pc_no: pcNo,
    });
  });

  let provisionalDone = 0;
  for (const row of provisionalRows) {
    const { error } = await supabase
      .from("provisional_certificates")
      .upsert(row, { onConflict: "student_id", ignoreDuplicates: false });
    if (error) {
      console.warn(
        `  -> provisional_certificates skipped for student_id=${row.student_id}: ${error.message}`
      );
      continue;
    }
    provisionalDone += 1;
  }

  console.log("Migration complete:");
  console.log(`  exams: ${examRows.length}`);
  console.log(`  classes: ${classesByCode.size}`);
  console.log(`  subjects: ${subjectsByCode.size}`);
  console.log(`  class_subjects: ${classSubjectRows.length}`);
  console.log(`  students: ${studentRowsToUpsert.length}`);
  console.log(`  marks: ${markRowsToUpsert.length}`);
  console.log(`  provisional_certificates: ${provisionalDone}/${provisionalRows.length}`);
}

main().catch((error) => {
  console.error("Migration failed:", error);
  process.exit(1);
});