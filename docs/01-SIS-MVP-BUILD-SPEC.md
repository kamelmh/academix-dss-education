# Student Information System — MVP Build Spec (MS Access)

> **Source:** HyperAgent (Opus 4.8) — 2026-07-27
> **Status:** Build-ready spec for OpenCode
> **Tech:** MS Access (.accdb) + VBA
> **Build time:** ~2-3 weeks

---

## 1. Scope & Tech-Stack Decision

### MVP boundary (the wedge)
Build **four capabilities only** — the daily-pain items a school admin does by hand:

**Students · Enrollment · Grades · Attendance**

**Deferred to later phases:** Fees, Staff (beyond a stub), Discipline, Inventory, the Curriculum Designer, and all AI features. Resist building these until the school says *"we need this."*

### Tech-stack decision → MS Access (.accdb) for the MVP

| Why Access now | Trade-off / mitigation |
|---|---|
| Reuses existing Access + VBA skill (Academix DSS) — build in ~2-3 weeks | Concurrency ceiling (~10-20 practical users). Fine for one school |
| Fits the **"concrete / owned / offline"** principle (Decision D1) | Not web-accessible. Acceptable for demo + single-site validation |
| Zero hosting cost, demo-ready fast | — |
| Bilingual + RTL + report generation all doable natively | — |

**Migration path:** keep surrogate integer keys and avoid Access-only field quirks so the schema ports cleanly to SQL Server/PostgreSQL/SQLite when scaling.

---

## 2. Data Model

Surrogate `AutoNumber` primary keys throughout. Bilingual text stored as parallel `_fr` / `_ar` fields.

### Tables

**tblAcademicYear** — YearID (PK), Label (`2026-2027`), StartDate, EndDate, IsCurrent (Yes/No)

**tblLevel** — LevelID (PK), Name_fr, Name_ar, Cycle (Elementary/Middle/Secondary/Vocational), OrderIndex

**tblSpecialization** — SpecializationID (PK), Code, Name_fr, Name_ar, FieldName, DurationMonths

**tblClass** — ClassID (PK), YearID (FK), LevelID (FK), Name (`1AS-A`), Track (Academic/Vocational), SpecializationID (FK, nullable), Capacity

**tblStudent** — StudentID (PK), Matricule (unique), LastName_fr, FirstName_fr, LastName_ar, FirstName_ar, Gender, DOB, PlaceOfBirth, NationalID, Address, Phone, GuardianName, GuardianPhone, PhotoPath, EnrollStatus (Active/Withdrawn/Graduated), CreatedAt

**tblEnrollment** — EnrollmentID (PK), StudentID (FK), ClassID (FK), YearID (FK), EnrollDate, Status, Notes

**tblSubject** — SubjectID (PK), Name_fr, Name_ar, Coefficient (Number), Track, LevelID (FK, nullable)

**tblClassSubject** — ClassSubjectID (PK), ClassID (FK), SubjectID (FK), TeacherID (FK, nullable)

**tblTeacher** — TeacherID (PK), LastName, FirstName, Phone, Email

**tblTerm** — TermID (PK), YearID (FK), Name (Trimestre 1/2/3), StartDate, EndDate

**tblGrade** — GradeID (PK), StudentID (FK), SubjectID (FK), TermID (FK), AssessType (Continuous/Exam/Composition), Mark (0-20), MaxMark (default 20), DateRecorded, TeacherID (FK)

**tblAttendance** — AttendanceID (PK), StudentID (FK), ADate, Session (AM/PM or period), Status (Present/Absent/Late/Excused), Justification, RecordedBy

### Relationships
```
tblAcademicYear 1——∞ tblClass ∞——1 tblLevel
tblClass 1——∞ tblEnrollment ∞——1 tblStudent
tblClass 1——∞ tblClassSubject ∞——1 tblSubject
tblStudent 1——∞ tblGrade ∞——1 tblSubject ,  tblGrade ∞——1 tblTerm
tblStudent 1——∞ tblAttendance
tblClass ∞——1 tblSpecialization  (vocational only)
```

**Grading convention:** marks out of **20**; trimestre average = Σ(mark×coefficient) / Σ(coefficient).

---

## 3. Screens / Forms

| Form | Purpose | Key elements |
|---|---|---|
| **frmDashboard** | Landing / KPIs | Active-student count, class count, today's attendance %, quick-action buttons |
| **frmStudent** | Register + profile | Bilingual fields, photo, national ID; subform = enrollment history |
| **frmEnrollment** | Assign student → class | Cascading Level → Class combos; capacity check |
| **frmGradeEntry** | Enter marks fast | Pick Class + Subject + Term → datasheet of students; 0-20 validation |
| **frmAttendanceRegister** | Daily register | Pick Class + Date → one-click "all present," then flag exceptions |
| **frmReports** | Report launcher | Buttons for bulletin, class list, attendance summary |

---

## 4. Reports (the demo value)

| Report | Contents |
|---|---|
| **rptBulletin** (report card) | Per student / trimestre: subjects, marks, coefficients, weighted average, class rank, appreciation line. Bilingual header. **← the wow factor.** |
| **rptClassList** | Class roster with photos, matricule, guardian phone |
| **rptAttendanceSummary** | Per class / term absence + lateness counts, justified vs. unjustified |

All exportable to **PDF** via VBA (`GenerateBulletinPDF`).

---

## 5. VBA Automation

| Function | Does |
|---|---|
| `CalcTermAverage(StudentID, TermID)` | Coefficient-weighted trimestre average |
| `CalcClassRank(ClassID, TermID)` | Rank all students in a class for a term |
| `GenerateBulletinPDF(StudentID, TermID)` | Render + export a report card to PDF |
| `BulkAttendance(ClassID, ADate, Status)` | Mark whole class present, then edit exceptions |
| `ValidateMark(value)` | Guard 0-20 on entry |
| `SetFormDirection(frm, isArabic)` | Toggle RTL/LTR for bilingual forms |

---

## 6. Build Sequence (~2-3 weeks)

### Week 1 — Foundation
- [ ] Create all tables + relationships + referential integrity
- [ ] Seed a fake school's worth of sample data
- [ ] Build `frmStudent` + `frmEnrollment`

### Week 2 — Grades
- [ ] Build `frmGradeEntry` with 0-20 validation
- [ ] Write `CalcTermAverage` + `CalcClassRank`
- [ ] Build `rptBulletin` + `GenerateBulletinPDF`

### Week 3 — Attendance + polish
- [ ] Build `frmAttendanceRegister` + `BulkAttendance`
- [ ] Build `frmDashboard` + `rptAttendanceSummary` + `rptClassList`
- [ ] Load realistic demo data, RTL polish, package `.accdb`

### Then
Walk former school through it → capture reaction / change list (validation gate).

---

## 7. Bilingual & Algerian Context

- **Language:** parallel `_fr` / `_ar` fields; Arabic rendered **RTL** on forms and reports.
- **Grading:** marks **/20**; three **trimestres**; subject **coefficients**; weighted averages and class ranks.
- **Identity:** **Matricule** as the human-facing student key; national ID captured.
- **Bulletin layout:** mirror the standard Algerian report-card format.
- **Dual-track hook:** `Track` + `SpecializationID` on `tblClass` — the seam where Curriculum Designer plugs in later.

---

## 8. Handoff & Connections

### For OpenCode
This spec is implementation-ready: every table, field, form, report, and VBA function is named. Build in order (§6).

### How it connects to the roadmap
- **→ Curriculum Designer:** `tblSpecialization` + `tblSubject` are the attach points.
- **→ Taallim:** student records can later feed the learning module.
- **→ Business model:** a working SIS is the demo that converts the former school into first paying customer.

---

**Tags:** #sis #mvp #access #tech/vba #build-spec #hyperagent #school-system


---

Linked from: [[00-MOC-Projects]]