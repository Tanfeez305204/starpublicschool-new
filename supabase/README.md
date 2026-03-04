# Supabase Setup

This project now supports database-backed student and marks management.

## 1. Create tables

1. Open your Supabase project.
2. Go to SQL Editor.
3. Run `supabase/schema.sql`.
4. If your existing DB has mandatory father name, run `supabase/father-name-optional.sql`.

## 2. Configure environment variables

Copy `.env.example` to `.env` and set:

- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY`
- existing mail/session settings

## 3. Start the server

```bash
npm start
```

## 4. Migrate Excel data to Supabase

Dry run:

```bash
npm run migrate:excel:dry
```

Apply migration:

```bash
npm run migrate:excel
```

## 5. Admin flow

1. Login at `/admin.html`.
2. You will be redirected to `/dashboard.html`.
3. Add/update student details.
4. Enter marks by:
   - class
   - roll number
   - exam term
   - subject-wise marks

## Notes

- `/result` and `/provisional` now read from Supabase first.
- If Supabase is unavailable, legacy Excel fallback is used.
- Marks accept numeric values or codes like `AB`, `NA`, `-`.
- `father_name` is optional in `students` table.
