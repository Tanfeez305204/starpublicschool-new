require('dotenv').config();
const { createClient } = require('@supabase/supabase-js');

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const DRY_RUN = !process.argv.includes('--yes');

const PRESERVE_TABLES = new Set(['classes', 'class_subjects', 'subjects']);
const DELETE_ORDER = [
  'marks',
  'provisional_certificates',
  'academic_year_promotions',
  'exams',
  'students',
];

if (!SUPABASE_URL || !SUPABASE_KEY) {
  console.error('SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are required in .env.');
  process.exit(1);
}

const supabase = createClient(SUPABASE_URL, SUPABASE_KEY, {
  auth: { persistSession: false, autoRefreshToken: false },
});

async function getTableCount(table) {
  const { count, error } = await supabase
    .from(table)
    .select('id', { count: 'exact', head: true });

  if (error) {
    throw new Error(`Failed to count ${table}: ${error.message || JSON.stringify(error)}`);
  }

  return count || 0;
}

async function deleteTable(table) {
  console.log(`Deleting all rows from ${table}...`);
  const { error } = await supabase.from(table).delete().not('id', 'is', null);
  if (error) {
    throw new Error(`Failed to delete ${table}: ${error.message || JSON.stringify(error)}`);
  }
  console.log(`  ${table} deleted.`);
}

async function main() {
  console.log('Database cleanup plan:');
  console.log('  Preserve tables: classes, class_subjects, subjects');
  console.log(`  Mode: ${DRY_RUN ? 'DRY RUN' : 'APPLY'}\n`);

  const tablesToClean = [];
  for (const table of DELETE_ORDER) {
    if (PRESERVE_TABLES.has(table)) continue;
    tablesToClean.push(table);
  }

  for (const table of tablesToClean) {
    const count = await getTableCount(table);
    console.log(`  ${table}: ${count} row(s)`);
  }

  if (DRY_RUN) {
    console.log('\nDry run complete. No deletions were applied.');
    console.log('Run with --yes to delete these rows.');
    return;
  }

  for (const table of tablesToClean) {
    await deleteTable(table);
  }

  console.log('\nCleanup complete. Confirm the preserved tables still contain data.');
}

main().catch((err) => {
  console.error('Cleanup failed:', err.message || err);
  process.exit(1);
});
