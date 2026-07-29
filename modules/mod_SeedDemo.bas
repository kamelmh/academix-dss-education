Attribute VB_Name = "mod_SeedDemo"
Option Compare Database
Option Explicit

' ============================================================================
'  mod_SeedDemo  —  load convincing demo data into SIS_MVP_Allal.accdb
'  Run from the Immediate window (Ctrl+G):   Call SeedDemoData
'  Safe to re-run: it clears the dynamic tables first, then reloads.
'
'  SCHEMA MATCH: This module matches mod_CreateTables.bas field names exactly.
'  Tables used: tblLevel, tblField, tblDiploma, tblSpecialization, tblModule,
'               tblTeacher, tblStudent, tblClass, tblEnrollment, tblGrade,
'               tblAttendance, tblConfig.
'
'  KEY SCHEMA NOTES:
'    - No tblAcademicYear / tblTerm — year stored in tblConfig
'    - tblGrade: ExamType + Grade (single field, NOT CC+Compo)
'    - tblClass: ClassName, SpecID, AcademicYear (NOT Name, YearID, Track)
'    - tblStudent: LastName_fr/FirstName_fr/LastName_ar/FirstName_ar, DateOfBirth
'    - tblModule: Hours (NOT Coefficient)
'    - tblSpecialization: SpecID (Text PK, NOT AutoNumber)
'    - tblDiploma: Abbreviation (NOT Code)
' ============================================================================

Private Const CUR_YEAR As String = "2026-2027"

' ---- Algerian name pools (Fr | Ar paired by index) ------------------------
Private mLastFr As Variant, mLastAr As Variant
Private mMFr As Variant, mMAr As Variant
Private mFFr As Variant, mFAr As Variant
Private mPOB As Variant

Private Sub InitPools()
    mLastFr = Array("Benali", "Kaddour", "Bouzid", "Cherif", "Meziane", "Brahimi", _
                    "Zerrouki", "Haddad", "Saidi", "Toumi", "Belkacem", "Mansouri")
    mLastAr = Array(ChrW(1576) & ChrW(1606) & " " & ChrW(1593) & ChrW(1604) & ChrW(1610), _
                    ChrW(1602) & ChrW(1583) & ChrW(1608) & ChrW(1585), _
                    ChrW(1576) & ChrW(1608) & ChrW(1586) & ChrW(1610) & ChrW(1583), _
                    ChrW(1588) & ChrW(1585) & ChrW(1610) & ChrW(1601), _
                    ChrW(1605) & ChrW(1586) & ChrW(1610) & ChrW(1575) & ChrW(1606), _
                    ChrW(1576) & ChrW(1585) & ChrW(1575) & ChrW(1607) & ChrW(1610) & ChrW(1605) & ChrW(1610), _
                    ChrW(1586) & ChrW(1585) & ChrW(1608) & ChrW(1602) & ChrW(1610), _
                    ChrW(1581) & ChrW(1583) & ChrW(1575) & ChrW(1583), _
                    ChrW(1587) & ChrW(1593) & ChrW(1610) & ChrW(1583) & ChrW(1610), _
                    ChrW(1578) & ChrW(1608) & ChrW(1605) & ChrW(1610), _
                    ChrW(1576) & ChrW(1604) & ChrW(1602) & ChrW(1575) & ChrW(1587) & ChrW(1605), _
                    ChrW(1605) & ChrW(1606) & ChrW(1589) & ChrW(1608) & ChrW(1585) & ChrW(1610))
    mMFr = Array("Yacine", "Amine", "Mohamed", "Ahmed", "Khaled", "Riad", "Bilal", "Sofiane", "Islam", "Adel")
    mMAr = Array(ChrW(1610) & ChrW(1575) & ChrW(1587) & ChrW(1610) & ChrW(1606), _
                 ChrW(1571) & ChrW(1605) & ChrW(1610) & ChrW(1606), _
                 ChrW(1605) & ChrW(1581) & ChrW(1605) & ChrW(1583), _
                 ChrW(1571) & ChrW(1581) & ChrW(1605) & ChrW(1583), _
                 ChrW(1582) & ChrW(1575) & ChrW(1604) & ChrW(1583), _
                 ChrW(1585) & ChrW(1610) & ChrW(1575) & ChrW(1590), _
                 ChrW(1576) & ChrW(1604) & ChrW(1575) & ChrW(1604), _
                 ChrW(1587) & ChrW(1601) & ChrW(1610) & ChrW(1575) & ChrW(1606), _
                 ChrW(1573) & ChrW(1587) & ChrW(1604) & ChrW(1575) & ChrW(1605), _
                 ChrW(1593) & ChrW(1575) & ChrW(1583) & ChrW(1604))
    mFFr = Array("Nour", "Sara", "Imane", "Lina", "Feriel", "Amira", "Meriem", "Rania", "Asma", "Hadil")
    mFAr = Array(ChrW(1606) & ChrW(1608) & ChrW(1585), _
                 ChrW(1587) & ChrW(1575) & ChrW(1585) & ChrW(1577), _
                 ChrW(1573) & ChrW(1610) & ChrW(1605) & ChrW(1575) & ChrW(1606), _
                 ChrW(1604) & ChrW(1610) & ChrW(1606) & ChrW(1575), _
                 ChrW(1601) & ChrW(1585) & ChrW(1610) & ChrW(1575) & ChrW(1604), _
                 ChrW(1571) & ChrW(1605) & ChrW(1610) & ChrW(1585) & ChrW(1577), _
                 ChrW(1605) & ChrW(1585) & ChrW(1610) & ChrW(1605), _
                 ChrW(1585) & ChrW(1575) & ChrW(1606) & ChrW(1610) & ChrW(1575), _
                 ChrW(1571) & ChrW(1587) & ChrW(1605) & ChrW(1575) & ChrW(1569), _
                 ChrW(1607) & ChrW(1583) & ChrW(1610) & ChrW(1604))
    mPOB = Array("El Bayadh", "Alger", "Oran", "Blida", ChrW(1587) & ChrW(1600) & ChrW(1578) & ChrW(1610) & ChrW(1601), _
                 "Constantine", "Tlemcen", "Annaba")
End Sub

' ============================================================================
Public Sub SeedDemoData()
    Dim db As DAO.Database
    On Error GoTo EH
    Set db = CurrentDb
    InitPools
    Randomize 7   ' fixed seed => reproducible demo

    If MsgBox("This DELETES existing students, classes, grades and attendance, " & _
              "then loads demo data. Continue?", vbYesNo + vbExclamation, "Seed demo") = vbNo Then Exit Sub

    db.Execute "DELETE FROM tblAttendance", dbFailOnError
    db.Execute "DELETE FROM tblGrade", dbFailOnError
    db.Execute "DELETE FROM tblEnrollment", dbFailOnError
    db.Execute "DELETE FROM tblStudent", dbFailOnError
    db.Execute "DELETE FROM tblClass", dbFailOnError
    db.Execute "DELETE FROM tblTeacher", dbFailOnError

    ' Ensure reference data exists
    Dim lvl1AM As Long, lvl2AM As Long, lvl3AM As Long, lvl4AM As Long, lvlBAC As Long
    Dim fldIT As Long, fldComm As Long, fldElec As Long, fldMech As Long
    Dim dipCAP As Long, dipBEP As Long, dipBTS As Long
    Dim specIT1 As Long, specComm1 As Long, specElec1 As Long
    Dim yID As String

    ' Levels (matching mod_CreateTables.bas seed data: "1ère année moyenne" etc.)
    lvl1AM = EnsureLevel(db, "1ère année moyenne", ChrW(1575) & ChrW(1604) & ChrW(1572) & ChrW(1608) & ChrW(1604) & ChrW(1610) & ChrW(1577) & " " & ChrW(1605) & ChrW(1578) & ChrW(1608) & ChrW(1587) & ChrW(1591), "Middle", 1)
    lvl2AM = EnsureLevel(db, "2ème année moyenne", ChrW(1575) & ChrW(1604) & ChrW(1579) & ChrW(1606) & ChrW(1610) & ChrW(1577) & " " & ChrW(1605) & ChrW(1578) & ChrW(1608) & ChrW(1587) & ChrW(1591), "Middle", 2)
    lvl3AM = EnsureLevel(db, "3ème année moyenne", ChrW(1575) & ChrW(1604) & ChrW(1579) & ChrW(1575) & ChrW(1604) & ChrW(1579) & ChrW(1577) & " " & ChrW(1605) & ChrW(1578) & ChrW(1608) & ChrW(1587) & ChrW(1591), "Middle", 3)
    lvl4AM = EnsureLevel(db, "4ème année moyenne", ChrW(1575) & ChrW(1604) & ChrW(1585) & ChrW(1575) & ChrW(1576) & ChrW(1593) & ChrW(1577) & " " & ChrW(1579) & ChrW(1606) & ChrW(1608) & ChrW(1610), "Secondary", 4)
    lvlBAC = EnsureLevel(db, "Baccalauréat", ChrW(1575) & ChrW(1576) & ChrW(1603) & ChrW(1575) & ChrW(1604) & ChrW(1608) & ChrW(1585) & ChrW(1610) & ChrW(1575), "Secondary", 5)

    ' Fields
    fldIT = EnsureField(db, "IT", ChrW(1575) & ChrW(1604) & ChrW(1605) & ChrW(1593) & ChrW(1604) & ChrW(1608) & ChrW(1605) & ChrW(1575) & ChrW(1578) & " " & ChrW(1575) & ChrW(1604) & ChrW(1587) & ChrW(1593) & ChrW(1610) & ChrW(1575) & ChrW(1578))
    fldComm = EnsureField(db, "COM", ChrW(1575) & ChrW(1604) & ChrW(1575) & ChrW(1587) & ChrW(1578) & ChrW(1602) & ChrW(1583) & ChrW(1575) & ChrW(1605))
    fldElec = EnsureField(db, "ELC", ChrW(1575) & ChrW(1604) & ChrW(1603) & ChrW(1607) & ChrW(1585) & ChrW(1575) & ChrW(1576) & ChrW(1575) & ChrW(1578))
    fldMech = EnsureField(db, "MEC", ChrW(1575) & ChrW(1604) & ChrW(1605) & ChrW(1610) & ChrW(1603) & ChrW(1575) & ChrW(1606) & ChrW(1610) & ChrW(1603) & ChrW(1575))

    ' Diplomas
    dipCAP = EnsureDiploma(db, "CAP", ChrW(1575) & ChrW(1604) & ChrW(1576) & ChrW(1591) & ChrW(1575) & ChrW(1604) & ChrW(1579) & ChrW(1606) & ChrW(1575) & ChrW(1574) & ChrW(1610) & ChrW(1577), "CAP", 2)
    dipBEP = EnsureDiploma(db, "BEP", ChrW(1575) & ChrW(1604) & ChrW(1576) & ChrW(1591) & ChrW(1575) & ChrW(1604) & ChrW(1579) & ChrW(1606) & ChrW(1575) & ChrW(1574) & ChrW(1610) & ChrW(1577), "BEP", 2)
    dipBTS = EnsureDiploma(db, "BTS", ChrW(1587) & ChrW(1600) & ChrW(1578) & ChrW(1610) & ChrW(1601) & ChrW(1610) & ChrW(1577) & " " & ChrW(1578) & ChrW(1606) & ChrW(1589) & ChrW(1610) & ChrW(1585) & ChrW(1610) & ChrW(1577), "BTS", 2)

    ' Specializations
    specIT1 = EnsureSpec(db, "IT-1", "Informatique", "IT", fldIT, dipBTS, 2, _
                         ChrW(1575) & ChrW(1604) & ChrW(1605) & ChrW(1593) & ChrW(1604) & ChrW(1608) & ChrW(1605) & ChrW(1575) & ChrW(1578) & " " & ChrW(1575) & ChrW(1604) & ChrW(1587) & ChrW(1593) & ChrW(1610) & ChrW(1575) & ChrW(1578))
    specComm1 = EnsureSpec(db, "CM-1", "Commerce", "CM", fldComm, dipBTS, 2, _
                           ChrW(1575) & ChrW(1604) & ChrW(1575) & ChrW(1587) & ChrW(1578) & ChrW(1602) & ChrW(1583) & ChrW(1575) & ChrW(1605))
    specElec1 = EnsureSpec(db, "EL-1", ChrW(1571) & ChrW(1603) & ChrW(1607) & ChrW(1585) & ChrW(1575) & ChrW(1576) & ChrW(1575) & ChrW(1578), "EL", fldElec, dipCAP, 2, _
                           ChrW(1575) & ChrW(1604) & ChrW(1603) & ChrW(1607) & ChrW(1585) & ChrW(1575) & ChrW(1578))

    ' Academic year
    yID = EnsureYear(db)

    ' --- teachers ---
    Dim tk As Long, ts As Long, tr As Long
    tk = AddTeacher(db, "Djelloul", "Karim", "Djelloul", ChrW(1583) & ChrW(1580) & ChrW(1604) & ChrW(1608) & ChrW(1604), "0551 23 45 67")
    ts = AddTeacher(db, "Haddad", "Samia", ChrW(1581) & ChrW(1583) & ChrW(1575) & ChrW(1583), ChrW(1587) & ChrW(1575) & ChrW(1605) & ChrW(1610) & ChrW(1575), "0661 78 90 12")
    tr = AddTeacher(db, "Brahimi", "Rachid", ChrW(1576) & ChrW(1585) & ChrW(1575) & ChrW(1607) & ChrW(1610) & ChrW(1605) & ChrW(1610), ChrW(1585) & ChrW(1575) & ChrW(1588) & ChrW(1610) & ChrW(1583), "0770 11 22 33")

    ' --- classes: 1 academic + 2 vocational (shows the dual track) ---
    Dim cA As Long, cIT As Long, cEL As Long
    cA = AddClass(db, "1AM-A", lvl1AM, specIT1, yID, tk, 35)
    cIT = AddClass(db, "BTS Info 1", lvlBAC, specIT1, yID, ts, 25)
    cEL = AddClass(db, "CAP Elec 1", lvl4AM, specElec1, yID, tr, 25)

    ' --- demo subjects/modules per class (create + return IDs & hours) ---
    Dim acadMods() As Long, acadHrs() As Double
    Dim itMods() As Long, itHrs() As Double
    Dim elMods() As Long, elHrs() As Double
    BuildModules db, "ACAD", acadMods, acadHrs
    BuildModules db, "IT", itMods, itHrs
    BuildModules db, "EL", elMods, elHrs

    ' --- students + enrollments + grades + attendance ---
    Dim seq As Long: seq = 0
    SeedClass db, yID, cA, 12, acadMods, acadHrs, seq
    SeedClass db, yID, cIT, 10, itMods, itHrs, seq
    SeedClass db, yID, cEL, 8, elMods, elHrs, seq

    MsgBox "Demo loaded: 3 classes, " & seq & " students, Exam grades, 5 days of attendance.", vbInformation
    Exit Sub
EH:
    MsgBox "Seed error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "(Field-name mismatch — check the schema notes at the top.)", vbCritical
End Sub

' ============================================================================
'  Reference-row helpers (look up, or create a minimal row if missing)
' ============================================================================
Private Function EnsureYear(db As DAO.Database) As String
    Dim v: v = DLookup("ConfigValue", "tblConfig", "ConfigKey='AcademicYear'")
    If Not IsNull(v) Then EnsureYear = CStr(v): Exit Function
    EnsureYear = CUR_YEAR
End Function

Private Function EnsureLevel(db As DAO.Database, sNameFr As String, sNameAr As String, cycle As String, idx As Long) As Long
    Dim v: v = DLookup("LevelID", "tblLevel", "Name_fr='" & sNameFr & "'")
    If Not IsNull(v) Then EnsureLevel = v: Exit Function
    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset("tblLevel", dbOpenDynaset)
    rs.AddNew: rs!Name_fr = sNameFr: rs!Name_ar = sNameAr: rs!Cycle = cycle: rs!OrderIndex = idx: rs.Update
    rs.Bookmark = rs.LastModified: EnsureLevel = rs!LevelID: rs.Close
End Function

Private Function EnsureField(db As DAO.Database, sCode As String, sNameAr As String) As Long
    Dim v: v = DLookup("FieldID", "tblField", "Code='" & sCode & "'")
    If Not IsNull(v) Then EnsureField = v: Exit Function
    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset("tblField", dbOpenDynaset)
    rs.AddNew: rs!Code = sCode: rs!Name_fr = sCode: rs!Name_ar = sNameAr: rs.Update
    rs.Bookmark = rs.LastModified: EnsureField = rs!FieldID: rs.Close
End Function

Private Function EnsureDiploma(db As DAO.Database, sAbbr As String, sNameFr As String, sNameAr As String, months As Long) As Long
    Dim v: v = DLookup("DiplomaID", "tblDiploma", "Abbreviation='" & sAbbr & "'")
    If Not IsNull(v) Then EnsureDiploma = v: Exit Function
    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset("tblDiploma", dbOpenDynaset)
    rs.AddNew: rs!Abbreviation = sAbbr: rs!Name_fr = sNameFr: rs!Name_ar = sNameAr
    rs!DurationMonths = months: rs!LevelRequired = 4: rs.Update
    rs.Bookmark = rs.LastModified: EnsureDiploma = rs!DiplomaID: rs.Close
End Function

Private Function EnsureSpec(db As DAO.Database, sID As String, sNameFr As String, sNameAr As String, _
                             fldID As Long, dipID As Long, months As Long, sNameArFull As String) As Long
    Dim v: v = DLookup("SpecID", "tblSpecialization", "SpecID='" & sID & "'")
    If Not IsNull(v) Then EnsureSpec = v: Exit Function
    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset("tblSpecialization", dbOpenDynaset)
    rs.AddNew: rs!SpecID = sID: rs!Name_fr = sNameFr: rs!Name_ar = sNameArFull
    rs!FieldID = fldID: rs!DiplomaID = dipID: rs!DurationMonths = months: rs.Update
    rs.Bookmark = rs.LastModified: EnsureSpec = rs!SpecID: rs.Close
End Function

' ============================================================================
'  Insert helpers (return new AutoNumber IDs)
' ============================================================================
Private Function AddTeacher(db As DAO.Database, sLastFr As String, sFirstFr As String, _
                             sLastAr As String, sFirstAr As String, sPhone As String) As Long
    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset("tblTeacher", dbOpenDynaset)
    rs.AddNew
    rs!LastName_fr = sLastFr: rs!FirstName_fr = sFirstFr
    rs!LastName_ar = sLastAr: rs!FirstName_ar = sFirstAr
    rs!Phone = sPhone: rs!Status = "Active"
    rs.Update: rs.Bookmark = rs.LastModified: AddTeacher = rs!TeacherID: rs.Close
End Function

Private Function AddClass(db As DAO.Database, sClassName As String, lvlID As Long, _
                           specID As Long, yrID As String, teacherID As Long, cap As Long) As Long
    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset("tblClass", dbOpenDynaset)
    rs.AddNew
    rs!ClassName = sClassName: rs!LevelID = lvlID: rs!SpecID = specID
    rs!AcademicYear = yrID: rs!TeacherID = teacherID: rs!MaxCapacity = cap
    rs.Update: rs.Bookmark = rs.LastModified: AddClass = rs!ClassID: rs.Close
End Function

Private Function AddModule(db As DAO.Database, specID As String, nameFr As String, _
                            nameAr As String, hrs As Double, Optional orderIdx As Long = 0) As Long
    Dim rs As DAO.Recordset: Set rs = db.OpenRecordset("tblModule", dbOpenDynaset)
    rs.AddNew: rs!SpecID = specID: rs!Name_fr = nameFr: rs!Name_ar = nameAr
    rs!Hours = hrs: rs!OrderIndex = orderIdx: rs.Update
    rs.Bookmark = rs.LastModified: AddModule = rs!ModuleID: rs.Close
End Function

' Build a small subject/module set for a track; return parallel arrays of IDs + hours
Private Sub BuildModules(db As DAO.Database, track As String, ByRef ids() As Long, ByRef hrs() As Double)
    Dim fr As Variant, ar As Variant, h As Variant, sp As Variant, i As Long
    Select Case track
      Case "ACAD"
        sp = Array("ACAD", "ACAD", "ACAD", "ACAD", "ACAD", "ACAD", "ACAD", "ACAD")
        fr = Array("Mathématiques", "Sciences Physiques", "Sciences Naturelles", "Langue Arabe", _
                   "Langue Française", "Langue Anglaise", "Histoire-Géographie", "Éducation Physique")
        ar = Array(ChrW(1575) & ChrW(1604) & ChrW(1585) & ChrW(1610) & ChrW(1575) & ChrW(1590) & ChrW(1610) & ChrW(1575) & ChrW(1578), _
                   ChrW(1575) & ChrW(1604) & ChrW(1593) & ChrW(1604) & ChrW(1608) & ChrW(1605) & ChrW(1575) & ChrW(1578) & " " & ChrW(1575) & ChrW(1604) & ChrW(1601) & ChrW(1610) & ChrW(1586) & ChrW(1610) & ChrW(1575) & ChrW(1569) & ChrW(1577), _
                   ChrW(1593) & ChrW(1604) & ChrW(1608) & ChrW(1605) & " " & ChrW(1575) & ChrW(1604) & ChrW(1591) & ChrW(1576) & ChrW(1610) & ChrW(1593) & ChrW(1577) & " " & ChrW(1608) & ChrW(1575) & ChrW(1604) & ChrW(1581) & ChrW(1610) & ChrW(1575) & ChrW(1572), _
                   ChrW(1575) & ChrW(1604) & ChrW(1604) & ChrW(1610) & ChrW(1594) & ChrW(1577) & " " & ChrW(1575) & ChrW(1604) & ChrW(1593) & ChrW(1585) & ChrW(1576) & ChrW(1610) & ChrW(1577), _
                   ChrW(1575) & ChrW(1604) & ChrW(1604) & ChrW(1610) & ChrW(1594) & ChrW(1577) & " " & ChrW(1575) & ChrW(1604) & ChrW(1601) & ChrW(1585) & ChrW(1606) & ChrW(1587) & ChrW(1610) & ChrW(1577), _
                   ChrW(1575) & ChrW(1604) & ChrW(1604) & ChrW(1610) & ChrW(1594) & ChrW(1577) & " " & ChrW(1575) & ChrW(1604) & ChrW(1573) & ChrW(1606) & ChrW(1580) & ChrW(1604) & ChrW(1610) & ChrW(1586) & ChrW(1610) & ChrW(1577), _
                   ChrW(1575) & ChrW(1604) & ChrW(1578) & ChrW(1575) & ChrW(1585) & ChrW(1610) & ChrW(1582) & " " & ChrW(1608) & ChrW(1575) & ChrW(1604) & ChrW(1580) & ChrW(1580) & ChrW(1585) & ChrW(1601) & ChrW(1610) & ChrW(1575), _
                   ChrW(1575) & ChrW(1604) & ChrW(1578) & ChrW(1585) & ChrW(1576) & ChrW(1610) & ChrW(1577) & " " & ChrW(1575) & ChrW(1604) & ChrW(1576) & ChrW(1583) & ChrW(1606) & ChrW(1610) & ChrW(1577))
        h = Array(5, 4, 4, 3, 3, 2, 2, 1)
      Case "IT"
        sp = Array("IT-1", "IT-1", "IT-1", "IT-1", "IT-1", "IT-1")
        fr = Array("Programmation", "Bases de données", "Réseaux informatiques", "Bureautique", "Anglais technique", "Stage pratique")
        ar = Array(ChrW(1575) & ChrW(1604) & ChrW(1576) & ChrW(1585) & ChrW(1605) & ChrW(1580) & ChrW(1577), _
                   ChrW(1602) & ChrW(1608) & ChrW(1575) & ChrW(1593) & ChrW(1583) & " " & ChrW(1575) & ChrW(1604) & ChrW(1576) & ChrW(1610) & ChrW(1575) & ChrW(1606) & ChrW(1575) & ChrW(1578), _
                   ChrW(1575) & ChrW(1604) & ChrW(1588) & ChrW(1576) & ChrW(1603) & ChrW(1575) & ChrW(1578), _
                   ChrW(1575) & ChrW(1604) & ChrW(1605) & ChrW(1603) & ChrW(1578) & ChrW(1576) & ChrW(1610) & ChrW(1577), _
                   ChrW(1575) & ChrW(1604) & ChrW(1573) & ChrW(1606) & ChrW(1580) & ChrW(1604) & ChrW(1610) & ChrW(1586) & ChrW(1610) & ChrW(1577) & " " & ChrW(1575) & ChrW(1604) & ChrW(1578) & ChrW(1602) & ChrW(1606) & ChrW(1610) & ChrW(1577), _
                   ChrW(1575) & ChrW(1604) & ChrW(1578) & ChrW(1585) & ChrW(1576) & ChrW(1589) & " " & ChrW(1575) & ChrW(1604) & ChrW(1578) & ChrW(1591) & ChrW(1576) & ChrW(1610) & ChrW(1602) & ChrW(1610))
        h = Array(5, 4, 4, 2, 2, 4)
      Case Else ' EL
        sp = Array("EL-1", "EL-1", "EL-1", "EL-1", "EL-1", "EL-1")
        fr = Array("Électrotechnique", "Installations électriques", "Sécurité", "Maths appliquées", "Anglais technique", "Stage pratique")
        ar = Array(ChrW(1575) & ChrW(1604) & ChrW(1603) & ChrW(1607) & ChrW(1585) & ChrW(1575) & ChrW(1608) & ChrW(1610) & ChrW(1577) & " " & ChrW(1575) & ChrW(1604) & ChrW(1578) & ChrW(1602) & ChrW(1606) & ChrW(1610) & ChrW(1577), _
                   ChrW(1575) & ChrW(1604) & ChrW(1578) & ChrW(1581) & ChrW(1585) & ChrW(1610) & ChrW(1603) & ChrW(1575) & ChrW(1578) & " " & ChrW(1575) & ChrW(1604) & ChrW(1603) & ChrW(1607) & ChrW(1585) & ChrW(1575) & ChrW(1576) & ChrW(1610) & ChrW(1575) & ChrW(1569), _
                   ChrW(1575) & ChrW(1604) & ChrW(1571) & ChrW(1605) & ChrW(1606) & " " & ChrW(1608) & ChrW(1575) & ChrW(1604) & ChrW(1608) & ChrW(1602) & ChrW(1575) & ChrW(1610) & ChrW(1577), _
                   ChrW(1575) & ChrW(1604) & ChrW(1585) & ChrW(1610) & ChrW(1575) & ChrW(1590) & ChrW(1610) & ChrW(1575) & ChrW(1578) & " " & ChrW(1575) & ChrW(1604) & ChrW(1578) & ChrW(1589) & ChrW(1604) & ChrW(1610) & ChrW(1577), _
                   ChrW(1575) & ChrW(1604) & ChrW(1573) & ChrW(1606) & ChrW(1580) & ChrW(1604) & ChrW(1610) & ChrW(1586) & ChrW(1610) & ChrW(1577) & " " & ChrW(1575) & ChrW(1604) & ChrW(1578) & ChrW(1602) & ChrW(1606) & ChrW(1610) & ChrW(1577), _
                   ChrW(1575) & ChrW(1604) & ChrW(1578) & ChrW(1585) & ChrW(1576) & ChrW(1589) & " " & ChrW(1575) & ChrW(1604) & ChrW(1578) & ChrW(1591) & ChrW(1576) & ChrW(1610) & ChrW(1602) & ChrW(1610))
        h = Array(4, 4, 2, 2, 2, 4)
    End Select
    ReDim ids(LBound(fr) To UBound(fr))
    ReDim hrs(LBound(fr) To UBound(fr))
    For i = LBound(fr) To UBound(fr)
        ids(i) = AddModule(db, CStr(sp(i)), CStr(fr(i)), CStr(ar(i)), CDbl(h(i)), i + 1)
        hrs(i) = CDbl(h(i))
    Next i
End Sub

' ============================================================================
'  Seed one class: students + enrollment + exam grades + attendance
' ============================================================================
Private Sub SeedClass(db As DAO.Database, yID As String, classID As Long, _
                      nStud As Long, mods() As Long, hrs() As Double, ByRef seq As Long)
    Dim rsS As DAO.Recordset, rsE As DAO.Recordset, rsG As DAO.Recordset, rsA As DAO.Recordset
    Set rsS = db.OpenRecordset("tblStudent", dbOpenDynaset)
    Set rsE = db.OpenRecordset("tblEnrollment", dbOpenDynaset)
    Set rsG = db.OpenRecordset("tblGrade", dbOpenDynaset)
    Set rsA = db.OpenRecordset("tblAttendance", dbOpenDynaset)

    Dim i As Long, j As Long, d As Long, sid As Long
    For i = 1 To nStud
        seq = seq + 1
        Dim male As Boolean: male = (Rnd < 0.5)
        Dim ni As Long, li As Long
        ni = Int(Rnd * 10): li = Int(Rnd * 12)
        Dim fFr As String, fAr As String, lFr As String, lAr As String
        If male Then fFr = mMFr(ni): fAr = mMAr(ni) Else fFr = mFFr(ni): fAr = mFAr(ni)
        lFr = mLastFr(li): lAr = mLastAr(li)
        Dim ability As Double: ability = 8 + Rnd * 9   ' 8..17 latent ability

        rsS.AddNew
        rsS!Matricule = Format(Year(Date), "0000") & "-" & Format(seq, "0000")
        rsS!LastName_fr = lFr: rsS!FirstName_fr = fFr
        rsS!LastName_ar = lAr: rsS!FirstName_ar = fAr
        rsS!Gender = IIf(male, "M", "F")
        rsS!DateOfBirth = DateSerial(2009 - Int(Rnd * 4), 1 + Int(Rnd * 12), 1 + Int(Rnd * 27))
        rsS!PlaceOfBirth = mPOB(Int(Rnd * 8))
        rsS!EnrollmentDate = DateSerial(2026, 9, 5)
        rsS!Status = "Active"
        rsS.Update: rsS.Bookmark = rsS.LastModified: sid = rsS!StudentID

        rsE.AddNew
        rsE!StudentID = sid: rsE!ClassID = classID: rsE!AcademicYear = CUR_YEAR
        rsE!Status = "Active"
        rsE.Update

        ' grades per module (ExamType + single Grade field, 0..20)
        For j = LBound(mods) To UBound(mods)
            rsG.AddNew
            rsG!StudentID = sid: rsG!ModuleID = mods(j)
            rsG!ExamType = "CC"
            rsG!Grade = Clamp20(ability + Noise(2.5))
            rsG!DateRecorded = DateSerial(2026, 11, 10 + Int(Rnd * 20))
            rsG.Update
        Next j

        ' attendance: 5 school days (Sun 9 Nov -> Thu 13 Nov 2026), mostly Present
        For d = 0 To 4
            rsA.AddNew
            rsA!StudentID = sid: rsA!ClassID = classID
            rsA!AttendanceDate = DateSerial(2026, 11, 9) + d
            rsA!Status = AttStatus()
            rsA.Update
        Next d
    Next i

    rsS.Close: rsE.Close: rsG.Close: rsA.Close
End Sub

' ---- small numeric helpers ----
Private Function Noise(sd As Double) As Double
    Noise = (Rnd + Rnd + Rnd - 1.5) * sd   ' ~N(0, sd), quick approximation
End Function
Private Function Clamp20(v As Double) As Double
    If v < 0 Then v = 0
    If v > 20 Then v = 20
    Clamp20 = Int(v * 4 + 0.5) / 4          ' round to nearest 0.25
End Function
Private Function AttStatus() As String
    Dim r As Double: r = Rnd
    If r < 0.9 Then AttStatus = "Present" _
    Else If r < 0.95 Then AttStatus = "Absent" _
    Else If r < 0.98 Then AttStatus = "Late" _
    Else AttStatus = "Excused"
End Function
