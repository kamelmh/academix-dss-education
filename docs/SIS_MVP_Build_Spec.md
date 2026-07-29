# Student Information System — MVP Build Spec (MS Access)

*Build-ready specification for the SIS MVP (the wedge product of the Allal dual-track school system). Hand to OpenCode / implement in MS Access.*

> **Provenance note:** this is the original MVP spec. The Week-1 Access build expanded the schema to **13 tables** (adding `tblField`, `tblDiploma`, `tblModule`, `tblFee`, `tblConfig`); vocational subjects are modelled as `tblModule`. Treat the table list below as the conceptual core, and defer to `02-SIS-MVP-BUILD-SPEC-UPDATED.md` / `mod_CreateTables.bas` for the authoritative field names.

---

## 1 · Scope & Tech-Stack Decision

### MVP boundary (the wedge)
Build **four capabilities only** — the daily-pain items a school admin does by hand:

**Students · Enrollment · Grades · Attendance**

**Deferred to later phases:** Fees, Staff (beyond a stub), Discipline, Inventory, the Curriculum Designer, and all AI features. Resist building these until the school says *"we need this."*

### Tech-stack decision → MS Access (`.accdb`) for the MVP
| Why Access now | Trade-off / mitigation |
|---|---|
| Reuses your existing Access + VBA skill (Academix DSS) — you can build it in ~2–3 weeks | Concurrency ceiling (~10–20 practical users). Fine for one school; not multi-tenant SaaS |
| Fits the **"concrete / owned / offline"** principle (Decision D1) — the school owns the `.accdb` | Not web-accessible. Acceptable for a demo + single-site validation |
| Zero hosting cost, demo-ready fast | — |
| Bilingual + RTL + report generation all doable natively | — |

**Migration path (design for it now):** keep surrogate integer keys and avoid Access-only field quirks in the core tables so the schema ports cleanly. When you scale to multi-school SaaS, move the **backend** to SQL Server Express / PostgreSQL / SQLite and the **front-end** to web — or split *Access front-end + SQL back-end* as an interim step. The schema below is written to survive that move.

**Recommendation:** build the Access MVP → validate with your former school → let real demand decide the web migration.

---

## 2 · Data Model (tables, fields, relationships)

Surrogate `AutoNumber` primary keys throughout (portable to SQL identity columns). Bilingual text stored as parallel `_fr` / `_ar` fields.

**tblAcademicYear** — YearID (PK), Label (`2026-2027`), StartDate, EndDate, IsCurrent (Yes/No)

**tblLevel** — LevelID (PK), Name_fr, Name_ar, Cycle (Elementary/Middle/Secondary/Vocational), OrderIndex

**tblSpecialization** *(vocational; stub in MVP, filled later from your template)* — SpecializationID (PK), Code, Name_fr, Name_ar, FieldName, DurationMonths

**tblClass** — ClassID (PK), YearID (FK), LevelID (FK), Name (`1AS-A`), Track (Academic/Vocational), SpecializationID (FK, nullable), Capacity

**tblStudent** — StudentID (PK), Matricule (unique), LastName_fr, FirstName_fr, LastName_ar, FirstName_ar, Gender, DOB, PlaceOfBirth, NationalID, Address, Phone, GuardianName, GuardianPhone, PhotoPath, EnrollStatus (Active/Withdrawn/Graduated), CreatedAt

**tblEnrollment** — EnrollmentID (PK), StudentID (FK), ClassID (FK), YearID (FK), EnrollDate, Status, Notes

**tblSubject** — SubjectID (PK), Name_fr, Name_ar, Coefficient (Number), Track, LevelID (FK, nullable)

**tblClassSubject** *(junction — which subjects a class takes)* — ClassSubjectID (PK), ClassID (FK), SubjectID (FK), TeacherID (FK, nullable)

**tblTeacher** *(stub for assignment)* — TeacherID (PK), LastName, FirstName, Phone, Email

**tblTerm** — TermID (PK), YearID (FK), Name (Trimestre 1/2/3), StartDate, EndDate

**tblGrade** — GradeID (PK), StudentID (FK), SubjectID (FK), TermID (FK), AssessType (Continuous/Exam/Composition), Mark (0–20), MaxMark (default 20), DateRecorded, TeacherID (FK)

**tblAttendance** — AttendanceID (PK), StudentID (FK), ADate, Session (AM/PM or period), Status (Present/Absent/Late/Excused), Justification, RecordedBy

### Relationships (enforce referential integrity)
```
tblAcademicYear 1──∞ tblClass ∞──1 tblLevel
tblClass 1──∞ tblEnrollment ∞──1 tblStudent
tblClass 1──∞ tblClassSubject ∞──1 tblSubject
tblStudent 1──∞ tblGrade ∞──1 tblSubject ,  tblGrade ∞──1 tblTerm
tblStudent 1──∞ tblAttendance
tblClass ∞──1 tblSpecialization  (vocational only)
```

**Grading convention:** marks out of **20**; trimestre average = Σ(mark×coefficient) / Σ(coefficient).

---

## 3 · Screens / Forms

| Form | Purpose | Key elements |
|---|---|---|
| **frmDashboard** | Landing / KPIs | Active-student count, class count, today's attendance %, quick-action buttons |
| **frmStudent** | Register + profile | Bilingual fields, photo, national ID; subform = enrollment history |
| **frmEnrollment** | Assign student → class (current year) | Cascading Level → Class combos; capacity check |
| **frmGradeEntry** | Enter marks fast | Pick Class + Subject + Term → datasheet of students; 0–20 validation |
| **frmAttendanceRegister** | Daily register | Pick Class + Date → one-click "all present," then flag exceptions |
| **frmReports** | Report launcher | Buttons for bulletin, class list, attendance summary |

Arabic labels rendered RTL; French default LTR. Keyboard-first data entry (teachers/admins move fast).

---

## 4 · Reports (the demo value)

These are what make an administrator say *"we need this"* — they replace hours of manual work.

| Report | Contents |
|---|---|
| **rptBulletin** (report card) | Per student / trimestre: subjects, marks, coefficients, **weighted average**, **class rank**, appreciation line. Bilingual header matching the familiar Algerian bulletin layout. **← the wow factor.** |
| **rptClassList** | Class roster with photos, matricule, guardian phone |
| **rptAttendanceSummary** | Per class / term absence + lateness counts, with justified vs. unjustified split |

All exportable to **PDF** via VBA (`GenerateBulletinPDF`). PDF bulletins are the artifact you hand parents — instant credibility in a demo.

---

## 5 · VBA Automation

| Function | Does |
|---|---|
| `CalcTermAverage(StudentID, TermID)` | Coefficient-weighted trimestre average |
| `CalcClassRank(ClassID, TermID)` | Rank all students in a class for a term |
| `GenerateBulletinPDF(StudentID, TermID)` | Render + export a report card to PDF |
| `BulkAttendance(ClassID, ADate, Status)` | Mark whole class present, then edit exceptions |
| `ValidateMark(value)` | Guard 0–20 on entry |
| `SetFormDirection(frm, isArabic)` | Toggle RTL/LTR for bilingual forms |

This is squarely in your Academix-DSS wheelhouse — the same VBA patterns (recordset loops, report automation, validation) you've already shipped.

---

## 6 · Build Sequence (~2–3 weeks)

- [ ] **Week 1 — Foundation**
  - [ ] Create all tables + relationships + referential integrity
  - [ ] Seed a fake school's worth of sample data (for a convincing demo)
  - [ ] Build `frmStudent` + `frmEnrollment`
- [ ] **Week 2 — Grades**
  - [ ] Build `frmGradeEntry` with 0–20 validation
  - [ ] Write `CalcTermAverage` + `CalcClassRank`
  - [ ] Build `rptBulletin` + `GenerateBulletinPDF`
- [ ] **Week 3 — Attendance + polish**
  - [ ] Build `frmAttendanceRegister` + `BulkAttendance`
  - [ ] Build `frmDashboard` + `rptAttendanceSummary` + `rptClassList`
  - [ ] Load realistic demo data, RTL polish, package `.accdb`
- [ ] **Then:** walk your former school through it → capture their reaction / change list (validation gate)

---

## 7 · Bilingual & Algerian Context

- **Language:** parallel `_fr` / `_ar` fields on names, subjects, levels; Arabic rendered **RTL** on forms and reports.
- **Grading:** marks **/20**; three **trimestres**; subject **coefficients**; weighted averages and class ranks — the norms Algerian schools and parents already expect.
- **Identity:** **Matricule** as the human-facing student key; national ID captured.
- **Bulletin layout:** mirror the standard Algerian report-card format so it feels instantly official.
- **Dual-track hook:** `Track` + `SpecializationID` on `tblClass` already distinguish academic vs. vocational classes — the seam where the Curriculum Designer plugs in later.

---

## 8 · Handoff & Connections

### For OpenCode
This spec is implementation-ready: every table, field, form, report, and VBA function is named. Hand it over section-by-section and build in order (§6).

### How it connects to the rest of the roadmap
- **→ Curriculum Designer:** `tblSpecialization` + `tblSubject` are the attach points. Populate them once you fill `SCHOOL_SPECIALIZATIONS_TEMPLATE.md` with your school's real programs.
- **→ Taallim:** student records here can later feed the Taallim learning module — per-student flashcard / MCQ decks keyed to `StudentID`.
- **→ Business model:** a working SIS is the demo that converts your former school from *design partner* into *first paying customer*.


---

Linked from: [[00-MOC-Projects]]