# SIS MVP — Week 3: rptBulletin (Report Card + PDF)

*Week-3 build walkthrough for `rptBulletin` — the Algerian report card with weighted average, class rank, mention, and PDF export (single + whole class). Built on the 13-table schema. Field names assumed; map to `02-SIS-MVP-BUILD-SPEC-UPDATED.md`.*

---

## 0 · Overview & assumptions

**Goal:** a printable, bilingual **bulletin** (report card) that, for a chosen student + trimestre, lists each subject with its mark, coefficient, and weighted points, then computes the **coefficient-weighted general average**, the **class rank**, a **mention**, and exports to **PDF** — for one student or a whole class in one click.

**Build order:** (1) helper query `qryGradeCoef` → (2) report source query `qryBulletin` → (3) the `rptBulletin` layout → (4) the VBA functions → (5) the `frmBulletinPicker` form with Preview / Export buttons.

> ✅ **Grade model confirmed:** `tblGrade` stores **`CC` (controle continu) + `Compo` (composition)**, not a single mark. The **subject mark** is computed as **(CC + Compo×2)/3** (matching `frmGradeEntry`) inside the queries in §1–§2 and aliased `Mark`, so §3–§4 are untouched.

**Two notes:**
- The weighting **(CC + Compo×2)/3** lives in one place per query — if your school uses a different rule (e.g. `(CC + Compo)/2`), change it there and nowhere else.
- `qryGradeCoef` counts only subjects where **both** CC and Compo are entered, so partial bulletins don't skew the average. For the report footer's general average to match the rank exactly on partial bulletins, bind that textbox to **`=TermAverage([StudentID],[TermID])`** (§4) instead of `=Sum([WeightedPoints])/Sum([Coef])`. For fully-graded bulletins the two are identical.
- Map field names (`Name_fr`, `Coefficient`, `Matricule`, `ClassName`, `ModuleID`, `TermID`, `CC`, `Compo`) to your `02-SIS-MVP-BUILD-SPEC-UPDATED.md`.

---

## 1 · Helper query — qryGradeCoef

This exposes each grade's **computed subject mark** next to its coefficient, so the average/rank functions can aggregate it. Create a query in SQL view, save as **`qryGradeCoef`**:

```sql
SELECT g.StudentID, g.TermID,
       ((g.CC + g.Compo*2)/3) AS Mark,
       m.Coefficient
FROM tblGrade AS g
INNER JOIN tblModule AS m ON g.ModuleID = m.ModuleID
WHERE g.CC Is Not Null AND g.Compo Is Not Null;
```
The computed subject mark is aliased **`Mark`**, so `TermAverage` and `ClassRank` in §4 work **unchanged**. The `WHERE` clause keeps partially-graded subjects out of the weighted average (numerator and denominator stay aligned).

---

## 2 · Report source query — qryBulletin

Drives the report's detail rows; it filters to the student + term chosen on the picker form (§5). Save as **`qryBulletin`**:

```sql
SELECT s.StudentID, s.Matricule,
       s.LastName_fr, s.FirstName_fr, s.LastName_ar, s.FirstName_ar,
       c.ClassName, c.Track,
       m.Name_fr AS Subject, m.Name_ar AS SubjectAr, m.Coefficient AS Coef,
       g.TermID, g.CC, g.Compo,
       ((g.CC + g.Compo*2)/3) AS Mark,
       (((g.CC + g.Compo*2)/3) * m.Coefficient) AS WeightedPoints
FROM ((((tblGrade AS g
     INNER JOIN tblModule AS m ON g.ModuleID = m.ModuleID)
     INNER JOIN tblStudent AS s ON g.StudentID = s.StudentID)
     INNER JOIN tblEnrollment AS e ON s.StudentID = e.StudentID)
     INNER JOIN tblClass AS c ON e.ClassID = c.ClassID)
WHERE g.StudentID = [Forms]![frmBulletinPicker]![cboStudent]
  AND g.TermID   = [Forms]![frmBulletinPicker]![cboTerm];
```
The report still shows **all** subjects (a subject missing its CC/Compo simply shows a blank `Mark`). `Mark` and `WeightedPoints` are now computed from `CC`/`Compo`, so the report layout (§3) and the VBA (§4) need **no other changes** — they read the same `Mark`, `Coef`, and `WeightedPoints` columns as before.

---

## 3 · The report — rptBulletin layout

Create a report (*Create > Report Design*), set its **Record Source = `qryBulletin`**, save as **`rptBulletin`**.

**Report Header** (official look):
- Two labels, centred: `الجمهورية الجزائرية الديمقراطية الشعبية` and `Republique Algerienne Democratique et Populaire`.
- Ministry line — make it a textbox that adapts to the class track:
  `=IIf([Track]="Vocational","Ministere de la Formation et de l'Enseignement Professionnels","Ministere de l'Education Nationale")`
- School block: `Groupe Scolaire Allal · مجمع الأمل` (your logo image optional).
- Title textbox: `="Bulletin de Notes — Trimestre " & [Forms]![frmBulletinPicker]![cboTerm]`
- Student line: `=[LastName_fr] & " " & [FirstName_fr] & "  |  Matricule: " & [Matricule] & "  |  Classe: " & [ClassName]`

**Detail** (one row per subject) — bind textboxes to: `Subject` (+ small `SubjectAr` in grey), `Mark` (format `0.00`), `Coef`, `WeightedPoints` (`0.00`), and an **Appreciation** textbox:
`=IIf([Mark]>=15,"Tres bien",IIf([Mark]>=12,"Bien",IIf([Mark]>=10,"Assez bien","Insuffisant")))`
Conditional-format the `Mark` box red when `< 10`.

**Report Footer** (the payoff):
- `txtTotalCoef` = `=Sum([Coef])`
- `txtTotalPoints` = `=Sum([WeightedPoints])`
- **`txtGeneralAvg`** = `=Sum([WeightedPoints])/Sum([Coef])` (format `0.00`) — the coefficient-weighted general average.
- **`txtRank`** = `=ClassRank([StudentID],[TermID])` (VBA, §4).
- **`txtMention`** = `=IIf([txtGeneralAvg]>=16,"Felicitations",IIf([txtGeneralAvg]>=14,"Encouragements",IIf([txtGeneralAvg]>=12,"Tableau d'honneur",IIf([txtGeneralAvg]>=10,"Passable","Doit fournir plus d'efforts"))))`
- **`txtDecision`** = `=IIf([txtGeneralAvg]>=10,"Admis(e)","A surveiller")`
- Signature labels: `Le Professeur Principal` · `Le Directeur`.

Set the AR controls' *Right to Left* = Yes.

---

## 4 · VBA — weighted average & class rank

Put these in a **standard module** (e.g., extend `mod_Utils`). `TermAverage` is the coefficient-weighted mean; `ClassRank` ranks the student among classmates for that term.

```vba
Public Function TermAverage(ByVal pStudentID As Long, ByVal pTermID As Long) As Double
    Dim tp As Variant, tc As Variant
    tp = DSum("[Mark]*[Coefficient]", "qryGradeCoef", _
              "StudentID=" & pStudentID & " AND TermID=" & pTermID)
    tc = DSum("[Coefficient]", "qryGradeCoef", _
              "StudentID=" & pStudentID & " AND TermID=" & pTermID)
    If Nz(tc, 0) = 0 Then TermAverage = 0 Else TermAverage = Nz(tp, 0) / tc
End Function

Public Function ClassRank(ByVal pStudentID As Long, ByVal pTermID As Long) As String
    Dim clsID As Long, myAvg As Double, better As Long, total As Long
    Dim rs As DAO.Recordset
    clsID = Nz(DLookup("ClassID", "tblEnrollment", "StudentID=" & pStudentID), 0)
    If clsID = 0 Then ClassRank = "-": Exit Function
    myAvg = TermAverage(pStudentID, pTermID)
    Set rs = CurrentDb.OpenRecordset("SELECT StudentID FROM tblEnrollment WHERE ClassID=" & clsID, dbOpenSnapshot)
    Do While Not rs.EOF
        If TermAverage(rs!StudentID, pTermID) > myAvg Then better = better + 1
        total = total + 1
        rs.MoveNext
    Loop
    rs.Close
    ClassRank = (better + 1) & " / " & total
End Function
```
*(Ties share the higher rank; refine later if your school ranks ties differently.)*

---

## 5 · Picker form + PDF export

Create **`frmBulletinPicker`** (unbound) with combos **`cboClass`** (from `tblClass`), **`cboStudent`** (from `tblStudent`; optionally filter by `cboClass`), **`cboTerm`** (from your terms), and three buttons.

```vba
Private Sub cmdPreview_Click()
    If IsNull(Me.cboStudent) Or IsNull(Me.cboTerm) Then
        MsgBox "Choisir un eleve et un trimestre.", vbExclamation: Exit Sub
    End If
    DoCmd.OpenReport "rptBulletin", acViewPreview
End Sub

Private Sub cmdExportPDF_Click()
    Dim mat As String, fnm As String
    If IsNull(Me.cboStudent) Or IsNull(Me.cboTerm) Then
        MsgBox "Choisir un eleve et un trimestre.", vbExclamation: Exit Sub
    End If
    mat = Nz(DLookup("Matricule", "tblStudent", "StudentID=" & Me.cboStudent), Me.cboStudent)
    fnm = CurrentProject.Path & "\Bulletin_" & mat & "_T" & Me.cboTerm & ".pdf"
    DoCmd.OutputTo acOutputReport, "rptBulletin", acFormatPDF, fnm, True  ' opens the PDF
    MsgBox "Bulletin exporte : " & fnm
End Sub

' Bonus: export every student in the selected class at once
Private Sub cmdExportClass_Click()
    Dim rs As DAO.Recordset, folder As String, mat As String, fnm As String
    If IsNull(Me.cboClass) Or IsNull(Me.cboTerm) Then
        MsgBox "Choisir une classe et un trimestre.", vbExclamation: Exit Sub
    End If
    folder = CurrentProject.Path & "\Bulletins_T" & Me.cboTerm & "\"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    Set rs = CurrentDb.OpenRecordset("SELECT StudentID FROM tblEnrollment WHERE ClassID=" & Me.cboClass, dbOpenSnapshot)
    Do While Not rs.EOF
        Me.cboStudent = rs!StudentID: Me.Refresh: DoEvents
        mat = Nz(DLookup("Matricule", "tblStudent", "StudentID=" & rs!StudentID), rs!StudentID)
        fnm = folder & "Bulletin_" & mat & ".pdf"
        DoCmd.OutputTo acOutputReport, "rptBulletin", acFormatPDF, fnm, False
        rs.MoveNext
    Loop
    rs.Close
    MsgBox "Bulletins generes dans : " & folder
End Sub
```
`acFormatPDF` is built into Access 2010+ (no add-in needed).

---

## 6 · Bilingual, Algerian conventions & test

- **Bilingual:** pair French labels with Arabic (e.g., *Bulletin de Notes / كشف النقاط*, *Moyenne Generale / المعدل العام*); set Arabic controls' *Right to Left = Yes*.
- **Conventions:** marks `/20` (format `0.00`), coefficients, trimestres, matricule, weighted general average, class rank, mention — the layout Algerian parents expect.
- **Match the mockup:** the interactive mockup's bulletin screen is your visual reference for spacing and hierarchy.

**Done when:**
- [ ] Pick a student + trimestre → **Preview** shows subjects, weighted average, correct rank, and mention.
- [ ] **Export PDF** writes `Bulletin_<matricule>_T<n>.pdf` and opens it.
- [ ] **Export class** produces one PDF per enrolled student in a `Bulletins_T<n>` folder.
- [ ] A student with any mark `< 10` shows it in red; the mention/decision read correctly.
- [ ] Back up `SIS_MVP_Allal.accdb` — this is the demo build you show the school.


---

Linked from: [[00-MOC-Projects]]