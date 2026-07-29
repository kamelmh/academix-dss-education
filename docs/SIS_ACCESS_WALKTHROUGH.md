# SIS MVP — Step-by-Step Access Walkthrough

> **Purpose:** Create a new Access database + run CreateAllTables + SeedDemoData
> **Prerequisites:** MS Access installed, VBA modules in `SIS_MVP_VBA\` folder
> **Last Updated:** 2026-07-27

---

## Step 1: Create a New Access Database

1. Open **Microsoft Access**
2. Click **Blank Database** (or File → New → Blank Database)
3. Name it: **`SIS_MVP_v0.1.accdb`**
4. Save it to: `C:\Users\Admin\My Drive\LifeWorkspace\04_Ideas_&_Projects\EdTech_System\SIS_MVP_VBA\`
5. Click **Create** — a blank table opens automatically
6. **Close** the blank table (we don't need it — we'll create tables via VBA)

---

## Step 2: Import the VBA Modules

### 2a. Open the VBA Editor

- Press **Alt + F11** (opens Microsoft Visual Basic for Applications)

### 2b. Import mod_CreateTables

1. In the VBA Editor menu: **File → Import File...**
2. Navigate to: `SIS_MVP_VBA\mod_CreateTables.bas`
3. Click **Open** — the module appears in the left panel under **Modules**

### 2c. Import mod_SeedDemo

1. **File → Import File...**
2. Navigate to: `SIS_MVP_VBA\mod_SeedDemo.bas`
3. Click **Open**

### 2d. Import mod_Utils (optional but recommended)

1. **File → Import File...**
2. Navigate to: `SIS_MVP_VBA\mod_Utils.bas`
3. Click **Open**

### 2e. Verify

You should see **three modules** in the left panel:
- `mod_CreateTables`
- `mod_SeedDemo`
- `mod_Utils`

If you see them, close the VBA Editor (Alt + Q or the X button) and return to Access.

---

## Step 3: Create All Tables

1. Make sure the database is open and no table is open
2. Open the **Immediate Window**: press **Ctrl + G** (or View → Immediate Window)
3. Type the following and press **Enter**:

```vba
Call CreateAllTables
```

### Expected Result

Access will execute `DoCmd.RunSQL` for each CREATE TABLE statement. You may see brief flashes as tables are created. At the end, check the **Navigation Pane** (left side):

You should see **13 tables**:
- `tblConfig`
- `tblLevel`
- `tblField`
- `tblDiploma`
- `tblSpecialization`
- `tblTeacher`
- `tblStudent`
- `tblClass`
- `tblModule`
- `tblEnrollment`
- `tblGrade`
- `tblAttendance`
- `tblFee`

### If You Get an Error

| Error | Cause | Fix |
|-------|-------|-----|
| "Table already exists" | Tables were created before | Run `DoCmd.RunSQL "DELETE FROM tblConfig"` for each table, or delete them manually in the Navigation Pane (right-click → Delete) |
| "Syntax error" | SQL statement issue | Check that `mod_CreateTables.bas` was imported correctly (no truncation) |
| "Permission denied" | Database is read-only | Right-click the .accdb file → Properties → uncheck Read-only |

---

## Step 4: Seed Demo Data

1. In the Immediate Window (**Ctrl + G**), type:

```vba
Call SeedDemoData
```

2. Press **Enter**

### Expected Result

The macro will:
- Insert **8 levels** (CP through BAC)
- Insert **4 fields** (IT, Commerce, Electricity, Mechanics)
- Insert **3 diplomas** (CAP, BEP, BTS)
- Insert **12 specializations** (4 per field)
- Insert **18 modules** (4-5 per specialization)
- Insert **2 teachers** (with Arabic + French names)
- Insert **6 students** (3 male, 3 female, with Arabic + French names)
- Insert **6 classes** (1 per level, first year of each)
- Insert **6 enrollments** (1 per student)
- Insert **18 grades** (3 per student: Arabic, French, English)
- Insert **12 attendance records** (2 per student)
- Insert **3 fee records** (1 per student)

### If You Get an Error

| Error | Cause | Fix |
|-------|-------|-----|
| "No such field" | mod_SeedDemo field names don't match mod_CreateTables | Re-import mod_SeedDemo.bas (use the corrected version) |
| "Duplicate primary key" | Data already seeded | Delete all data: right-click each table → Open → delete rows, or re-run `Call CreateAllTables` after deleting tables |
| "Type mismatch" | Date/number format issue | Check that AcademicYear is TEXT ("2026-2027"), not a number |

---

## Step 5: Verify the Data

### 5a. Open a Table

- In the Navigation Pane, double-click **`tblStudent`**
- You should see 6 rows with names like:
  - LastName_fr: Benali, FirstName_fr: Ahmed
  - LastName_ar: بنالي, FirstName_ar: أحمد

### 5b. Check Relationships

1. Database Tools → **Relationships**
2. Add all tables
3. You should see:
   - tblLevel → tblSpecialization (via LevelID)
   - tblField → tblSpecialization (via FieldID)
   - tblSpecialization → tblModule (via SpecID)
   - tblLevel → tblClass (via LevelID)
   - tblSpecialization → tblClass (via SpecID)
   - tblTeacher → tblClass (via TeacherID)
   - tblStudent → tblEnrollment (via StudentID)
   - tblClass → tblEnrollment (via ClassID)
   - tblStudent → tblGrade (via StudentID)
   - tblModule → tblGrade (via ModuleID)
   - tblStudent → tblAttendance (via StudentID)
   - tblClass → tblAttendance (via ClassID)
   - tblStudent → tblFee (via StudentID)

### 5c. Quick Queries

Run these in the Immediate Window to verify:

```vba
' Count students
Debug.Print DCount("*", "tblStudent")

' Count grades
Debug.Print DCount("*", "tblGrade")

' List all classes
Dim rs As DAO.Recordset
Set rs = CurrentDb.OpenRecordset("SELECT ClassName FROM tblClass")
Do Until rs.EOF
    Debug.Print rs!ClassName
    rs.MoveNext
Loop
rs.Close
```

---

## Step 6: Test the Mockup (Optional)

If you want to see the interactive web mockup:

1. Open the mockup file:
   ```
   C:\Users\Admin\My Drive\LifeWorkspace\04_Ideas_&_Projects\EdTech_System\SIS_MVP_VBA\sis_mockup.html
   ```
2. Double-click → opens in your default browser
3. Click through: **Tableau de bord → Gestion des Élèves → Gestion des Notes → Relevé de Notes**
4. Toggle **AR / FR** in the top-right corner
5. Click **Imprimer le bulletin** to preview the bilingual report card

---

## Step 7: Dashboard (Once Forms Are Built)

After OpenCode builds the UserForms (Week 2+):

1. Open Access → **Create → Form Design**
2. Or run: `Call ShowDashboard` (if mod_Dashboard.bas is imported)

---

## Quick Reference

| Action | Command |
|--------|---------|
| Create tables | `Call CreateAllTables` |
| Seed demo data | `Call SeedDemoData` |
| Clean everything | `Call CleanupWorkbook` |
| Count records | `Debug.Print DCount("*", "tblStudent")` |
| Open VBA Editor | **Alt + F11** |
| Open Immediate Window | **Ctrl + G** |

---

## Troubleshooting Checklist

- [ ] .accdb is not read-only
- [ ] "Enable Content" was clicked
- [ ] All three .bas files imported successfully (check Modules folder in VBA Editor)
- [ ] No old tables exist (delete them manually if needed)
- [ ] mod_SeedDemo.bas uses correct field names (LastName_fr, not DOB, etc.)

---

**Author:** MAHI Kamel Abdelghani
**Last Updated:** 2026-07-27


---

Linked from: [[00-MOC-Projects]]