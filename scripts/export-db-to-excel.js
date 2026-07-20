require('dotenv').config();
const Excel = require('exceljs');
const { createClient } = require('@supabase/supabase-js');

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const EXPORT_FILE = process.env.EXPORT_FILE || 'db-backup.xlsx';
const TABLES = [
  'classes',
  'subjects',
  'class_subjects',
  'students',
  'exams',
  'marks',
  'provisional_certificates',
  'academic_year_promotions',
  'v_student_marks',
];

if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY) {
  console.error('SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are required in .env.');
  process.exit(1);
}

const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, {
  auth: { persistSession: false, autoRefreshToken: false },
});

async function fetchTable(table) {
  const { data, error } = await supabase.from(table).select('*');
  if (error) {
    throw new Error(`Failed to fetch ${table}: ${error.message || JSON.stringify(error)}`);
  }
  return data || [];
}

function stringifyValue(value) {
  if (value === null || value === undefined) return '';
  if (typeof value === 'object') return JSON.stringify(value);
  return String(value);
}

async function main() {
  const workbook = new Excel.Workbook();
  const sheet = workbook.addWorksheet('db_backup');

  sheet.addRow(['table_name', 'record_index', 'record_json']);

  let totalRows = 0;
  for (const table of TABLES) {
    console.log(`Fetching ${table}...`);
    const rows = await fetchTable(table);
    rows.forEach((row, index) => {
      const json = JSON.stringify(row);
      sheet.addRow([table, index + 1, json]);
    });
    totalRows += rows.length;
    console.log(`  ${table}: ${rows.length} rows`);
  }

  sheet.columns = [
    { header: 'table_name', key: 'table_name', width: 22 },
    { header: 'record_index', key: 'record_index', width: 14 },
    { header: 'record_json', key: 'record_json', width: 120 },
  ];

  console.log(`Writing ${EXPORT_FILE} (${totalRows} rows)...`);
  await workbook.xlsx.writeFile(EXPORT_FILE);
  console.log('Backup complete.');
}

main().catch((err) => {
  console.error('Export failed:', err.message || err);
  process.exit(1);
});
