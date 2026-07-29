# SIS MVP — Week 2 Forms: Build Walkthrough + VBA

## 0 · Before you start (read once)

**Two conventions for this whole guide:**

1. **Field-name mapping.** I use logical field names (e.g., `Matricule`, `LastName_fr`, `CC`, `Compo`). If your `02-SIS-MVP-BUILD-SPEC-UPDATED.md` used different names, change them where flagged — the *logic* is what matters. If you paste me that spec (or `mod_CreateTables.bas`), I'll rewrite every snippet with your exact names so it's copy-paste-clean.
2. **Where VBA goes.** Form-event code lives in the form's own module: open the form in Design view → *Property Sheet* → *Event* tab → click the event (e.g., *Before Update*) → `[Event Procedure]` → the `…` button opens the VBA editor at the right stub. Reusable functions (like `NextMatricule`) go in a **standard module** (e.g., extend `mod_Utils`).

**One reference to confirm:** the recordset code below uses **DAO**. In the VBA editor → *Tools ▸ References* → ensure *Microsoft Office x.x Access Database Engine Object Library* (DAO) is ticked. It usually is by default.

**Build order (each is independently testable):** frmDashboard → frmStudentEntry → frmClassManager → frmGradeEntry.

## 1 · frmDashboard — KPIs (almost no VBA)

**Create:** Design view → new **unbound** form, name it `frmDashboard`.

**KPI tiles = unbound textboxes with domain-aggregate `Control Source` expressions** (no code needed). Set each textbox's *Control Source* to:

| Textbox | Control Source expression |
|---|---|
| `txtActiveStudents` | `=DCount("*","tblStudent","Status='Active'")` |
| `txtClasses` | `=DCount("*","tblClass")` |
| `txtAttendanceToday` | `=Format(DCount("*","tblAttendance","Status='Present' And ADate=Date()")/Abs(Nz(DCount("*","tblAttendance","ADate=Date()"),1)),"0%")` |
| `txtPendingBulletins` | `=DCount("*","tblStudent","Status='Active'")` |

*(Adjust `Status`, `ADate` to your real field names.)*

**Navigation buttons:** add command buttons; in each `Click` event:
```vba
Private Sub cmdStudents_Click()
    DoCmd.OpenForm "frmStudentEntry"
End Sub
```
(Repeat for `frmClassManager`, `frmGradeEntry`.)

**Refresh button (optional):**
```vba
Private Sub cmdRefresh_Click()
    Me.Requery
End Sub
```

**Done when:** opening `frmDashboard` shows live counts from your seed data.

## 2 · frmStudentEntry — auto-matricule + validation

**Create:** Design view → *Record Source* = `tblStudent`. Add bound controls for the name fields (FR + AR), gender, DOB, contact, guardian, status; make **`Matricule` read-only** (`Locked = Yes`, `Enabled = No`). Add an enrollment subform later.

**A) Auto-generate the matricule on a new record.** If `mod_Utils` already has a matricule generator, call it instead of the function below. Otherwise put this in a standard module:
```vba
Public Function NextMatricule() As String
    Dim yr As String, prefix As String, seq As Long
    yr = Format(Date, "yyyy")
    prefix = yr & "-"
    ' assumes Matricule looks like "2026-0001"
    seq = Nz(DMax("Val(Mid([Matricule],6))", "tblStudent", _
            "Left([Matricule],5)='" & prefix & "'"), 0) + 1
    NextMatricule = prefix & Format(seq, "0000")
End Function
```
Then, in the form module:
```vba
Private Sub Form_BeforeInsert(Cancel As Integer)
    If Len(Nz(Me![Matricule], "")) = 0 Then Me![Matricule] = NextMatricule()
End Sub
```

**B) Validate on save** (form module):
```vba
Private Sub Form_BeforeUpdate(Cancel As Integer)
    If Len(Nz(Me![LastName_fr], "")) = 0 Or Len(Nz(Me![FirstName_fr], "")) = 0 Then
        MsgBox "Nom et prénom obligatoires.", vbExclamation
        Cancel = True: Exit Sub
    End If
    If IsNull(Me![DOB]) Then
        MsgBox "Date de naissance obligatoire.", vbExclamation
        Cancel = True: Exit Sub
    End If
End Sub
```

**C) Arabic RTL:** set the AR textboxes' *Right to Left* property to *Yes*.

**Done when:** you add a student, the matricule auto-fills (e.g., `2026-0001`), missing required fields are blocked, and the record saves to `tblStudent`.

## 3 · frmClassManager — dual-track toggle

**Create:** *Record Source* = `tblClass`. Add: `txtClassName`, a **`cboTrack`** combo with values `Academic`/`Vocational`, `cboLevel` (bound to `tblLevel`), `cboSpecialization` (bound to `tblSpecialization`), `cboDiploma` (bound to `tblDiploma`), `txtCapacity`, `cboTeacher` (bound to `tblTeacher`). Add a roster subform (students via `tblEnrollment`) later.

**Show the right selectors for the chosen track** (form module):
```vba
Private Sub SetTrackVisibility()
    Dim isVoc As Boolean
    isVoc = (Nz(Me![cboTrack], "") = "Vocational")
    Me![cboSpecialization].Visible = isVoc
    Me![cboDiploma].Visible = isVoc
    Me![cboLevel].Visible = Not isVoc
End Sub

Private Sub Form_Current()
    SetTrackVisibility
End Sub

Private Sub cboTrack_AfterUpdate()
    SetTrackVisibility
End Sub
```

**Capacity guard (optional)** — on the enrollment action, compare `DCount` of the class's enrollments against `txtCapacity` and block if full.

**Done when:** switching `cboTrack` to *Vocational* reveals Specialization + Diploma and hides Level (and vice-versa); a class saves to `tblClass`.

## 4 · frmGradeEntry — grade sheet + live average

**Concept:** header combos pick **Class + Module + Term**; a *Load* button ensures a `tblGrade` row exists for each enrolled student, then binds an editable subform; a footer shows the live class average.

> ⚠️ **Assumption to confirm:** this treats `tblGrade` as having `StudentID, ModuleID, TermID, CC, Compo`. If your grade table instead stores a single `Mark` + `AssessType`, tell me and I'll swap the query/calc — the structure stays the same.

**Enforce 0–20 with no code:** on the `CC` and `Compo` textboxes set *Validation Rule* = `Between 0 And 20` and *Validation Text* = `La note doit être entre 0 et 20.`

**Subject mark (unbound textbox, no code):** *Control Source* =
`=IIf(IsNull([CC]) Or IsNull([Compo]),Null,([CC]+[Compo]*2)/3)`

**Load button** (header form module):
```vba
Private Sub cmdLoad_Click()
    Dim cls As Long, modID As Long, term As Long, sql As String
    cls = Nz(Me![cboClass], 0): modID = Nz(Me![cboModule], 0): term = Nz(Me![cboTerm], 0)
    If cls = 0 Or modID = 0 Or term = 0 Then
        MsgBox "Choisir classe, module et trimestre.", vbExclamation: Exit Sub
    End If
    ' create any missing grade rows for enrolled students
    sql = "INSERT INTO tblGrade (StudentID, ModuleID, TermID) " & _
          "SELECT e.StudentID, " & modID & ", " & term & " FROM tblEnrollment e " & _
          "WHERE e.ClassID=" & cls & " AND NOT EXISTS (SELECT 1 FROM tblGrade g " & _
          "WHERE g.StudentID=e.StudentID AND g.ModuleID=" & modID & " AND g.TermID=" & term & ")"
    CurrentDb.Execute sql, dbFailOnError
    Me![subGrades].Form.RecordSource = _
        "SELECT g.GradeID, s.LastName_fr & ' ' & s.FirstName_fr AS Eleve, g.CC, g.Compo " & _
        "FROM tblGrade g INNER JOIN tblStudent s ON g.StudentID=s.StudentID " & _
        "WHERE g.ModuleID=" & modID & " AND g.TermID=" & term & _
        " AND g.StudentID IN (SELECT StudentID FROM tblEnrollment WHERE ClassID=" & cls & ")"
    RecalcClassAverage
End Sub
```

**Live class average** (put in the header form module; call from the subform's `AfterUpdate` via `Parent.RecalcClassAverage`):
```vba
Public Sub RecalcClassAverage()
    Dim rs As DAO.Recordset, t As Double, n As Long
    On Error Resume Next
    Set rs = Me![subGrades].Form.RecordsetClone
    If Not (rs Is Nothing) Then
        If Not (rs.BOF And rs.EOF) Then
            rs.MoveFirst
            Do While Not rs.EOF
                If Not IsNull(rs![CC]) And Not IsNull(rs![Compo]) Then
                    t = t + (rs![CC] + rs![Compo] * 2) / 3: n = n + 1
                End If
                rs.MoveNext
            Loop
        End If
    End If
    Me![txtClassAvg] = IIf(n > 0, Format(t / n, "0.00") & " / 20", "—")
End Sub
```

**Done when:** picking a class/module/term + *Load* lists the students, out-of-range marks are refused, the subject mark computes per row, and the footer shows the live class average; marks persist to `tblGrade`.

## 5 · Wire-up & test checklist

- [ ] Set `frmDashboard` as the startup form: *File ▸ Options ▸ Current Database ▸ Display Form*.
- [ ] frmDashboard shows correct KPI counts from seed data.
- [ ] frmStudentEntry: new student → matricule auto-fills; required-field validation fires; saves.
- [ ] frmClassManager: track toggle shows/hides the right selectors; class saves.
- [ ] frmGradeEntry: Load builds the sheet; 0–20 enforced; subject mark + class average compute; saves.
- [ ] Commit a backup copy of `SIS_MVP_Allal.accdb` once all four pass.

**Then Week 3:** `rptBulletin` (weighted-average + rank report card, PDF export) and `frmAttendanceRegister` — the pieces that turn this data into the artifact you show the school.

**Prefer a one-click builder instead?** I can generate a `mod_CreateForms.bas` that builds these forms programmatically (like your `CreateAllTables`) — just note it's harder to verify blind, so the designer + paste-VBA route above is the safer first pass.


---

Linked from: [[00-MOC-Projects]]