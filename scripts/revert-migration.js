require('dotenv').config();
const { createClient } = require('@supabase/supabase-js');

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const ACADEMIC_SESSION = process.env.ACADEMIC_SESSION || '';
const SHEET_FILTER = process.env.SHEET_FILTER || '';
const EXAM_CODES = process.env.EXAM_CODES || SHEET_FILTER || '1st,2nd,3rd,annual';
const DRY_RUN = !process.argv.includes('--yes');

if (!SUPABASE_URL || !SUPABASE_KEY) {
  console.error('SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY are required in environment.');
  process.exit(1);
}

if (!ACADEMIC_SESSION) {
  console.error('ACADEMIC_SESSION must be set to identify which session to revert.');
  process.exit(1);
}

const supabase = createClient(SUPABASE_URL, SUPABASE_KEY, {
  auth: { persistSession: false, autoRefreshToken: false },
});

async function run() {
  const examCodes = EXAM_CODES.split(',').map(s => s.trim()).filter(Boolean);

  console.log('Revert migration plan:');
  console.log(`  Session: ${ACADEMIC_SESSION}`);
  console.log(`  Exam codes: ${examCodes.join(', ')}`);
  console.log(`  Dry run: ${DRY_RUN ? 'YES' : 'NO (will apply deletions)'}\n`);

  // Find exam ids matching session + exam codes
  const { data: exams, error: exErr } = await supabase
    .from('exams')
    .select('id, exam_code')
    .in('exam_code', examCodes)
    .eq('academic_session', ACADEMIC_SESSION);

  if (exErr) {
    console.error('Failed to query exams:', exErr.message || exErr);
    process.exit(1);
  }

  if (!exams || exams.length === 0) {
    console.log('No matching exams found for given session and exam codes. Nothing to do.');
    return;
  }

  const examIds = exams.map(e => e.id);
  console.log(`Found exams: ${exams.map(e => `${e.exam_code}(id=${e.id})`).join(', ')}`);

  // Count marks to delete
  const { count: marksCount, error: countErr } = await supabase
    .from('marks')
    .select('id', { count: 'exact', head: true })
    .in('exam_id', examIds);

  if (countErr) {
    console.error('Failed to count marks:', countErr.message || countErr);
    process.exit(1);
  }

  console.log(`Marks matching exams: ${marksCount}`);

  if (DRY_RUN) {
    console.log('\nDry run complete — no deletions performed.');
    return;
  }

  // Delete marks
  console.log('\nDeleting marks...');
  const { error: delMarksErr } = await supabase
    .from('marks')
    .delete()
    .in('exam_id', examIds);

  if (delMarksErr) {
    console.error('Failed to delete marks:', delMarksErr.message || delMarksErr);
    process.exit(1);
  }
  console.log('Marks deleted.');

  // Delete exams
  console.log('Deleting exams...');
  const { error: delExErr } = await supabase
    .from('exams')
    .delete()
    .in('id', examIds);

  if (delExErr) {
    console.error('Failed to delete exams:', delExErr.message || delExErr);
    process.exit(1);
  }
  console.log('Exams deleted.');

  console.log('\nRevert complete. Verify your Supabase tables to confirm.');
}

run().catch(err => {
  console.error('Revert failed:', err);
  process.exit(1);
});
