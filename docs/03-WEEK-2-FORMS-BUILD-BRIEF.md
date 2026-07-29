# SIS MVP — Week 2 Forms Build Brief (for OpenCode)

## 0 · Recommended sequence & how to use

You forwarded three options — here's the sensible order:

1. **Do Option 2 first (test the tables).** Before layering forms, run `CreateAllTables` in a fresh `SIS_MVP_Allal.accdb`, confirm all **13 tables** appear with relationships intact and seed data loaded, and fix any schema errors. Ten minutes now saves hours of form debugging later.
2. **Then Option 1 (build the forms)** — this brief specs the four Week-2 forms.
3. **Option 3 (§4 Results) is premature:** you can't write Results without pilot data (means, SDs, ANCOVA output). What *is* useful now is a **Results scaffold/template** — say the word and I'll produce it separately.

**How to read this brief:** I reference tables by the confirmed names from your Week-1 build (`tblStudent`, `tblClass`, etc.) and fields *logically* — map each to the exact field names in your `02-SIS-MVP-BUILD-SPEC-UPDATED.md`. Where a form should call a helper, I point to `mod_Utils` (matricule generation, validation, calculations).

> ⚠️ **One schema question to confirm first:** how are **academic subjects** represented versus **vocational modules**? Your 13 tables include `tblModule` ("modules per specialization") but no generic `tblSubject`. If academic subjects (Maths, Arabe, Français…) are also stored in `tblModule`, `frmGradeEntry` works as-is; if they live elsewhere, adjust its record source. Lock this before building grade entry.

## 1 · frmDashboard (unbound)

**Purpose:** startup landing screen — KPIs + quick navigation.

**Record source:** unbound. Populate via `DCount` / `DLookup` or small aggregate queries on load.

**Controls:**
- KPI tiles: *Active students* (`DCount` on `tblStudent` where status = Active) · *Classes this year* (`DCount` on `tblClass` for current year from `tblConfig`) · *Today's attendance %* (from `tblAttendance` where date = `Date()`) · *Pending bulletins*.
- Navigation buttons → `frmStudentEntry`, `frmClassManager`, `frmGradeEntry`.
- Optional: a listbox of recent activity.

**Events / logic:** `Form_Load` fills KPI labels; a Refresh button re-runs them. Read current academic year from `tblConfig`.

**Done when:** set as the startup form; KPIs reflect the seed data on open.

## 2 · frmStudentEntry (bound → tblStudent)

**Purpose:** register/edit a student and view their enrollment.

**Record source:** `tblStudent` (main). Subform: `tblEnrollment` (child) linked on StudentID.

**Controls:**
- Bilingual name fields (FR + AR), **matricule (read-only, auto-generated)**, gender, DOB, place of birth, address, guardian name + phone, photo path, status combo.
- Enrollment subform: shows the student's class history; a combo bound to `tblClass` assigns them to a class for the current year (writes a `tblEnrollment` row).

**Events / logic:**
- On new record (a **New** button or `Form_BeforeInsert`): call the **matricule-generation** helper in `mod_Utils` to populate the matricule.
- On save (`Form_BeforeUpdate`): run `mod_Utils` validation (required fields present, DOB plausible, unique matricule). Cancel + message on failure.
- Set Arabic text controls to RTL.

**Done when:** you can add a student, matricule auto-assigns, assign them to a class, and the record persists to `tblStudent` + `tblEnrollment`.

## 3 · frmClassManager (bound → tblClass)

**Purpose:** create/manage class sections and see the roster.

**Record source:** `tblClass`. Lookup combos bound to `tblLevel` (academic), `tblSpecialization` + `tblDiploma` (vocational), and `tblTeacher` (homeroom).

**Controls:**
- Class name, academic year (default from `tblConfig`), **Track toggle (Academic / Vocational)**, capacity, homeroom teacher.
- The Track toggle shows/hides the right selectors: *Academic* → `tblLevel`; *Vocational* → `tblSpecialization` + `tblDiploma`.
- Read-only list of `tblModule` rows for the chosen specialization (so the user sees what the class studies).
- Roster subform: students via `tblEnrollment` → `tblStudent`, with a live count vs capacity.

**Events / logic:** Track toggle visibility; capacity check on enrollment; refresh module list on specialization change.

**Done when:** you can create one **academic** class (e.g., a lycée level) and one **vocational** class (e.g., BTS Informatique), the roster subform shows enrolled students, and capacity is enforced.

## 4 · frmGradeEntry (continuous / editable)

**Purpose:** fast mark entry per Class × Module × Term, with automatic averages.

**Record source:** a query returning students of the selected class alongside their `tblGrade` rows for the selected module + term (editable continuous subform).

**Controls:**
- Header combos: **Class** (`tblClass`), **Module/Subject** (`tblModule`, filtered by the class's specialization/level — see the schema question in §0), **Term** (from `tblConfig` or a terms table).
- Detail rows: student name (read-only) + **Contrôle continu** and **Composition** inputs.
- Footer: **live class average** in this module.

**Events / logic:**
- `AfterUpdate` on a mark input → validate 0–20 via `mod_Utils` (flag/reject out-of-range), compute subject mark = **(CC + Composition × 2) / 3** (confirm this is your rule), recompute the footer average.
- Save writes to `tblGrade`.

**mod_Utils calls:** `ValidateMark` (0–20), subject-average and class-average calculators.

**Done when:** you can enter marks for a class/module/term, out-of-range values are rejected, averages compute live, data saves to `tblGrade`, and reopening shows the saved marks.

## 5 · Build order, then Week 3

**Build order (each form is independently testable):** frmDashboard → frmStudentEntry → frmClassManager → frmGradeEntry. Keep the **bilingual FR/AR + RTL toggle** as a polish pass *after* core CRUD works — don't let it block functionality.

**Deferred to Week 3 (per the original spec):** `rptBulletin` (the printable weighted-average + rank report card — the demo centrepiece), `frmAttendanceRegister`, and the PDF export helper. Those turn the working data into the artifact you show the school.

**Reminder:** the interactive mockup already visualises the *target* look of the dashboard, grade entry, and bulletin — use it as the reference for what these Access forms should ultimately feel like.


---

Linked from: [[00-MOC-Projects]]