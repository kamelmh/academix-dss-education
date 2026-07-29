# SIS MVP Build Spec — Access Database

> **Target:** MS Access 2016+ (.accdb)
> **Language:** VBA (reuse Academix DSS patterns)
> **Bilingual:** Arabic + French
> **School:** Allal Badr El-Din (confirmed: 5 levels, 4 fields, 22 specializations)

---

## Database Schema

### tblLevel (Academic Levels)
| Column | Type | Description |
|--------|------|-------------|
| LevelID | AutoNumber (PK) | Unique ID |
| Name_fr | Text (50) | French name |
| Name_ar | Text (50) | Arabic name |
| Cycle | Text (20) | "Middle" or "Secondary" |
| OrderIndex | Integer | Sort order (1-5) |

### tblField (Vocational Fields)
| Column | Type | Description |
|--------|------|-------------|
| FieldID | AutoNumber (PK) | Unique ID |
| Name_fr | Text (50) | French name |
| Name_ar | Text (50) | Arabic name |
| Code | Text (10) | Short code (IT, COM, ELEC, MECH) |

### tblDiploma (Diploma Types)
| Column | Type | Description |
|--------|------|-------------|
| DiplomaID | AutoNumber (PK) | Unique ID |
| Name_fr | Text (100) | French name |
| Name_ar | Text (100) | Arabic name |
| Abbreviation | Text (10) | CAP, BEP, BTS |
| DurationMonths | Integer | 12 or 24 |
| LevelRequired | Integer | FK to tblLevel |

### tblSpecialization (Vocational Specializations)
| Column | Type | Description |
|--------|------|-------------|
| SpecID | Text (10) (PK) | e.g., "IT-01" |
| FieldID | Integer | FK to tblField |
| Name_fr | Text (100) | French name |
| Name_ar | Text (100) | Arabic name |
| DiplomaID | Integer | FK to tblDiploma |
| DurationMonths | Integer | 12 or 24 |

### tblModule (Modules per Specialization)
| Column | Type | Description |
|--------|------|-------------|
| ModuleID | AutoNumber (PK) | Unique ID |
| SpecID | Text (10) | FK to tblSpecialization |
| Name_fr | Text (100) | French module name |
| Name_ar | Text (100) | Arabic module name |
| Hours | Integer | Teaching hours |
| OrderIndex | Integer | Sort order |

### tblStudent (Students)
| Column | Type | Description |
|--------|------|-------------|
| StudentID | AutoNumber (PK) | Unique ID |
| Matricule | Text (20) | Student number (unique) |
| LastName_ar | Text (50) | Arabic last name |
| FirstName_ar | Text (50) | Arabic first name |
| LastName_fr | Text (50) | French last name |
| FirstName_fr | Text (50) | French first name |
| DateOfBirth | Date/Time | Birth date |
| PlaceOfBirth | Text (50) | Birth place |
| Gender | Text (1) | M/F |
| Phone | Text (20) | Contact phone |
| Address | Text (100) | Address |
| EnrollmentDate | Date/Time | Registration date |
| LevelID | Integer | FK to tblLevel |
| SpecID | Text (10) | FK to tblSpecialization |
| Status | Text (20) | Active/Graduated/Withdrawn |

### tblTeacher (Teachers/Instructors)
| Column | Type | Description |
|--------|------|-------------|
| TeacherID | AutoNumber (PK) | Unique ID |
| LastName_ar | Text (50) | Arabic last name |
| FirstName_ar | Text (50) | Arabic first name |
| LastName_fr | Text (50) | French last name |
| FirstName_fr | Text (50) | French first name |
| Phone | Text (20) | Contact phone |
| Email | Text (100) | Email address |
| Specialty | Text (100) | Teaching specialty |
| HireDate | Date/Time | Employment date |
| Status | Text (20) | Active/Inactive |

### tblClass (Class Sections)
| Column | Type | Description |
|--------|------|-------------|
| ClassID | AutoNumber (PK) | Unique ID |
| ClassName | Text (50) | e.g., "IT-CAP-2026-A" |
| LevelID | Integer | FK to tblLevel |
| SpecID | Text (10) | FK to tblSpecialization |
| AcademicYear | Text (10) | e.g., "2026-2027" |
| TeacherID | Integer | FK to tblTeacher (class advisor) |
| MaxCapacity | Integer | Max students |

### tblEnrollment (Student-Class Links)
| Column | Type | Description |
|--------|------|-------------|
| EnrollmentID | AutoNumber (PK) | Unique ID |
| StudentID | Integer | FK to tblStudent |
| ClassID | Integer | FK to tblClass |
| AcademicYear | Text (10) | e.g., "2026-2027" |
| Status | Text (20) | Active/Transferred/Withdrawn |

### tblGrade (Grades/Marks)
| Column | Type | Description |
|--------|------|-------------|
| GradeID | AutoNumber (PK) | Unique ID |
| StudentID | Integer | FK to tblStudent |
| ModuleID | Integer | FK to tblModule |
| ExamType | Text (20) | Continuous/Final/Resit |
| Grade | Double | 0-20 scale |
| Comment | Text (200) | Teacher comment |
| DateRecorded | Date/Time | When entered |

### tblAttendance (Attendance Tracking)
| Column | Type | Description |
|--------|------|-------------|
| AttendanceID | AutoNumber (PK) | Unique ID |
| StudentID | Integer | FK to tblStudent |
| ClassID | Integer | FK to tblClass |
| AttendanceDate | Date/Time | Date |
| Status | Text (10) | Present/Absent/Late/Excused |
| Note | Text (100) | Reason if absent |

### tblFee (Fee Management)
| Column | Type | Description |
|--------|------|-------------|
| FeeID | AutoNumber (PK) | Unique ID |
| StudentID | Integer | FK to tblStudent |
| FeeType | Text (50) | Tuition/Registration/Exam |
| Amount | Currency | Fee amount |
| DueDate | Date/Time | Payment deadline |
| PaidDate | Date/Time | When paid |
| Status | Text (20) | Pending/Paid/Overdue |
| ReceiptNo | Text (20) | Receipt number |

### tblConfig (System Configuration)
| Column | Type | Description |
|--------|------|-------------|
| ConfigKey | Text (50) (PK) | Configuration key |
| ConfigValue | Text (200) | Configuration value |
| Description | Text (200) | What this config does |

---

## Relationships

```
tblLevel ← tblSpecialization (via LevelRequired)
tblField ← tblSpecialization (via FieldID)
tblDiploma ← tblSpecialization (via DiplomaID)
tblSpecialization ← tblModule (via SpecID)
tblLevel ← tblClass (via LevelID)
tblSpecialization ← tblClass (via SpecID)
tblTeacher ← tblClass (via TeacherID)
tblStudent ← tblEnrollment (via StudentID)
tblClass ← tblEnrollment (via ClassID)
tblStudent ← tblGrade (via StudentID)
tblModule ← tblGrade (via ModuleID)
tblStudent ← tblAttendance (via StudentID)
tblClass ← tblAttendance (via ClassID)
tblStudent ← tblFee (via StudentID)
```

---

## Forms (8 UserForms)

### frmDashboard
- Main navigation form
- KPIs: Total students, enrollment by field, attendance rate
- Quick access to all forms

### frmStudentEntry
- Student registration form
- Bilingual labels (FR/AR)
- Auto-generate matricule
- Level/Specialization dropdowns

### frmClassManager
- Create/edit class sections
- Assign teachers to classes
- Set academic year

### frmGradeEntry
- Enter grades by module
- Calculate averages
- Generate bulletin

### frmAttendance
- Daily attendance tracking
- Mark present/absent/late/excused
- Monthly report

### frmFeeManagement
- Track payments
- Generate receipts
- Overdue alerts

### frmReports
- Student list by class
- Grade reports
- Attendance reports
- Fee reports

### frmConfig
- Academic year settings
- Fee amounts
- System preferences

---

## VBA Modules

### mod_Config
- Constants for table/column names
- Connection string
- Academic year variables

### mod_Utils
- Auto-generate matricule
- Validate Arabic/French text
- Date formatting

### mod_Student
- CRUD operations for students
- Search by name/matricule
- Enroll in class

### mod_Grade
- Enter grades
- Calculate averages
- Generate bulletin data

### mod_Attendance
- Mark attendance
- Monthly summary
- Absence alerts

### mod_Fee
- Track payments
- Overdue detection
- Receipt generation

---

## Build Sequence (3 Weeks)

### Week 1: Foundation
- [ ] Create database (.accdb)
- [ ] Create all tables with relationships
- [ ] Create mod_Config
- [ ] Create mod_Utils
- [ ] Seed data (levels, fields, diplomas, specializations)

### Week 2: Core Forms
- [ ] frmDashboard
- [ ] frmStudentEntry
- [ ] frmClassManager
- [ ] frmGradeEntry

### Week 3: Reports & Polish
- [ ] frmAttendance
- [ ] frmFeeManagement
- [ ] frmReports
- [ ] frmConfig
- [ ] Test with sample data

---

**Tags:** #sis #mvp #access #build-spec #database #tables


---

Linked from: [[00-MOC-Projects]]