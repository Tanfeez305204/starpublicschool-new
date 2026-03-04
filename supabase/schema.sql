begin;

create extension if not exists pgcrypto;

create or replace function public.set_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

create table if not exists public.classes (
  id bigserial primary key,
  class_name text not null unique,
  class_code text not null unique,
  description text default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.subjects (
  id bigserial primary key,
  subject_name text not null,
  subject_code text not null unique,
  full_marks integer not null default 100 check (full_marks > 0),
  pass_marks integer not null default 30 check (pass_marks >= 0 and pass_marks <= full_marks),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.class_subjects (
  class_id bigint not null references public.classes(id) on delete cascade,
  subject_id bigint not null references public.subjects(id) on delete restrict,
  display_order integer not null default 1,
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  primary key (class_id, subject_id)
);

create table if not exists public.students (
  id uuid primary key default gen_random_uuid(),
  class_id bigint not null references public.classes(id) on delete restrict,
  roll_no text not null,
  full_name text not null,
  father_name text,
  address text default '',
  phone text default '',
  dob date,
  gender text default '' check (gender in ('', 'Male', 'Female', 'Other')),
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  unique (class_id, roll_no)
);

alter table if exists public.students
  alter column father_name drop not null;

create table if not exists public.exams (
  id bigserial primary key,
  exam_code text not null unique check (exam_code in ('1st', '2nd', '3rd', 'annual')),
  exam_name text not null,
  academic_session text not null,
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.marks (
  id bigserial primary key,
  student_id uuid not null references public.students(id) on delete cascade,
  exam_id bigint not null references public.exams(id) on delete cascade,
  subject_id bigint not null references public.subjects(id) on delete restrict,
  marks_obtained numeric(5,2),
  grade text,
  remark text default '',
  entered_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  constraint ck_marks_range check (marks_obtained is null or (marks_obtained >= 0 and marks_obtained <= 100)),
  constraint ck_grade_len check (grade is null or char_length(grade) <= 20),
  unique (student_id, exam_id, subject_id)
);

create table if not exists public.provisional_certificates (
  id bigserial primary key,
  student_id uuid not null unique references public.students(id) on delete cascade,
  pc_no text not null unique,
  issued_on date not null default current_date,
  created_at timestamptz not null default now()
);

create index if not exists idx_students_roll on public.students(roll_no);
create index if not exists idx_students_class on public.students(class_id);
create index if not exists idx_marks_student_exam on public.marks(student_id, exam_id);
create index if not exists idx_marks_exam_subject on public.marks(exam_id, subject_id);

-- updated_at triggers

drop trigger if exists trg_classes_updated_at on public.classes;
create trigger trg_classes_updated_at
before update on public.classes
for each row execute function public.set_updated_at();

drop trigger if exists trg_subjects_updated_at on public.subjects;
create trigger trg_subjects_updated_at
before update on public.subjects
for each row execute function public.set_updated_at();

drop trigger if exists trg_students_updated_at on public.students;
create trigger trg_students_updated_at
before update on public.students
for each row execute function public.set_updated_at();

drop trigger if exists trg_exams_updated_at on public.exams;
create trigger trg_exams_updated_at
before update on public.exams
for each row execute function public.set_updated_at();

drop trigger if exists trg_marks_updated_at on public.marks;
create trigger trg_marks_updated_at
before update on public.marks
for each row execute function public.set_updated_at();

-- Seed classes
insert into public.classes (class_name, class_code)
values
  ('NURSERY-A', 'NURSERYA'),
  ('NURSERY-B', 'NURSERYB'),
  ('NURSERY-C', 'NURSERYC'),
  ('L.K.G-A', 'LKGA'),
  ('L.K.G-B', 'LKGB'),
  ('U.K.G-A', 'UKGA'),
  ('U.K.G-B', 'UKGB'),
  ('CLASS 1', 'CLASS1'),
  ('CLASS 2', 'CLASS2'),
  ('CLASS 3', 'CLASS3'),
  ('CLASS 4', 'CLASS4'),
  ('CLASS 5', 'CLASS5'),
  ('CLASS 6', 'CLASS6'),
  ('CLASS 7', 'CLASS7'),
  ('CLASS 8', 'CLASS8')
on conflict (class_code) do update
set class_name = excluded.class_name;

-- Seed exams
insert into public.exams (exam_code, exam_name, academic_session)
values
  ('1st', 'First Terminal', '2025-26'),
  ('2nd', 'Second Terminal', '2025-26'),
  ('3rd', 'Third Terminal', '2025-26'),
  ('annual', 'Annual', '2025-26')
on conflict (exam_code) do update
set exam_name = excluded.exam_name,
    academic_session = excluded.academic_session;

-- Seed subjects
insert into public.subjects (subject_name, subject_code, full_marks, pass_marks)
values
  ('Hindi', 'HINDI', 100, 30),
  ('English', 'ENGLISH', 100, 30),
  ('Math', 'MATH', 100, 30),
  ('Science', 'SCIENCE', 100, 30),
  ('S.S.T', 'SST', 100, 30),
  ('GK', 'GK', 100, 30),
  ('Sanskrit / Urdu', 'SANSKRIT_URDU', 100, 30),
  ('Drawing', 'DRAWING', 100, 30),
  ('Computer / Table', 'COMPUTER_TABLE', 100, 30)
on conflict (subject_code) do update
set subject_name = excluded.subject_name,
    full_marks = excluded.full_marks,
    pass_marks = excluded.pass_marks;

-- Lower classes: Nursery/LKG/UKG
insert into public.class_subjects (class_id, subject_id, display_order, is_active)
select
  c.id,
  s.id,
  case s.subject_code
    when 'HINDI' then 1
    when 'ENGLISH' then 2
    when 'MATH' then 3
    when 'GK' then 4
    when 'DRAWING' then 5
    when 'COMPUTER_TABLE' then 6
    else 99
  end as display_order,
  true
from public.classes c
join public.subjects s
  on s.subject_code in ('HINDI', 'ENGLISH', 'MATH', 'GK', 'DRAWING', 'COMPUTER_TABLE')
where c.class_code in ('NURSERYA', 'NURSERYB', 'NURSERYC', 'LKGA', 'LKGB', 'UKGA', 'UKGB')
on conflict (class_id, subject_id) do update
set display_order = excluded.display_order,
    is_active = excluded.is_active;

-- Class 1-8
insert into public.class_subjects (class_id, subject_id, display_order, is_active)
select
  c.id,
  s.id,
  case s.subject_code
    when 'HINDI' then 1
    when 'ENGLISH' then 2
    when 'MATH' then 3
    when 'SCIENCE' then 4
    when 'SST' then 5
    when 'GK' then 6
    when 'SANSKRIT_URDU' then 7
    when 'DRAWING' then 8
    when 'COMPUTER_TABLE' then 9
    else 99
  end as display_order,
  true
from public.classes c
join public.subjects s
  on s.subject_code in ('HINDI', 'ENGLISH', 'MATH', 'SCIENCE', 'SST', 'GK', 'SANSKRIT_URDU', 'DRAWING', 'COMPUTER_TABLE')
where c.class_code in ('CLASS1', 'CLASS2', 'CLASS3', 'CLASS4', 'CLASS5', 'CLASS6', 'CLASS7', 'CLASS8')
on conflict (class_id, subject_id) do update
set display_order = excluded.display_order,
    is_active = excluded.is_active;

create or replace view public.v_student_marks as
select
  st.id as student_id,
  c.class_name,
  st.roll_no,
  st.full_name,
  st.father_name,
  e.exam_code,
  e.exam_name,
  e.academic_session,
  su.subject_name,
  su.subject_code,
  m.marks_obtained,
  m.grade,
  m.remark,
  m.entered_at
from public.marks m
join public.students st on st.id = m.student_id
join public.classes c on c.id = st.class_id
join public.exams e on e.id = m.exam_id
join public.subjects su on su.id = m.subject_id;

commit;
