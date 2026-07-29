Option Compare Database
Option Explicit

' =============================================================================
' mod_Utils — SIS MVP Utility Functions
' School: Allal Badr El-Din (El Bayadh)
' =============================================================================

' -----------------------------------------------------------------------------
' GenerateMatricule — Create unique student number
' Format: YYYY-FIELD-NNNN (e.g., 2026-IT-0001)
' -----------------------------------------------------------------------------
Public Function GenerateMatricule(ByVal SpecID As String) As String
    Dim sql As String
    Dim rs As DAO.Recordset
    Dim fieldCode As String
    Dim yearStr As String
    Dim seqNum As Long
    
    ' Get field code from specialization
    sql = "SELECT f.Code FROM tblField f INNER JOIN tblSpecialization s ON f.FieldID = s.FieldID " & _
          "WHERE s.SpecID = '" & SpecID & "'"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    
    If rs.EOF Then
        GenerateMatricule = "ERROR"
        rs.Close
        Exit Function
    End If
    
    fieldCode = rs!Code
    rs.Close
    
    ' Get current year
    yearStr = Format(Year(Date), "YYYY")
    
    ' Get next sequence number
    sql = "SELECT COUNT(*) AS Cnt FROM tblStudent WHERE SpecID = '" & SpecID & "'"
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    seqNum = rs!Cnt + 1
    rs.Close
    
    ' Build matricule
    GenerateMatricule = yearStr & "-" & fieldCode & "-" & Format(seqNum, "0000")
End Function

' -----------------------------------------------------------------------------
' ValidateArabic — Check if text contains Arabic characters
' -----------------------------------------------------------------------------
Public Function ValidateArabic(ByVal txt As String) As Boolean
    Dim i As Long
    Dim charCode As Long
    
    If Len(txt) = 0 Then
        ValidateArabic = False
        Exit Function
    End If
    
    For i = 1 To Len(txt)
        charCode = AscW(Mid(txt, i, 1))
        ' Arabic Unicode range: 0x0600 to 0x06FF
        If charCode >= &H600 And charCode <= &H6FF Then
            ValidateArabic = True
            Exit Function
        End If
    Next i
    
    ValidateArabic = False
End Function

' -----------------------------------------------------------------------------
' ValidateFrench — Check if text contains French/Latin characters
' -----------------------------------------------------------------------------
Public Function ValidateFrench(ByVal txt As String) As Boolean
    Dim i As Long
    Dim charCode As Long
    
    If Len(txt) = 0 Then
        ValidateFrench = False
        Exit Function
    End If
    
    For i = 1 To Len(txt)
        charCode = AscW(Mid(txt, i, 1))
        ' Latin Unicode range: 0x0020 to 0x007E
        If charCode >= &H20 And charCode <= &H7E Then
            ValidateFrench = True
            Exit Function
        End If
    Next i
    
    ValidateFrench = False
End Function

' -----------------------------------------------------------------------------
' FormatDateFr — Format date in French style (DD/MM/YYYY)
' -----------------------------------------------------------------------------
Public Function FormatDateFr(ByVal dt As Date) As String
    FormatDateFr = Format(dt, "DD/MM/YYYY")
End Function

' -----------------------------------------------------------------------------
' FormatDateAr — Format date in Arabic style (YYYY/MM/DD)
' -----------------------------------------------------------------------------
Public Function FormatDateAr(ByVal dt As Date) As String
    FormatDateAr = Format(dt, "YYYY/MM/DD")
End Function

' -----------------------------------------------------------------------------
' CalculateAverage — Calculate average grade for a student
' -----------------------------------------------------------------------------
Public Function CalculateAverage(ByVal StudentID As Long, Optional ByVal ModuleID As Long = 0) As Double
    Dim sql As String
    Dim rs As DAO.Recordset
    Dim total As Double
    Dim count As Long
    
    sql = "SELECT Grade FROM tblGrade WHERE StudentID = " & StudentID
    If ModuleID > 0 Then
        sql = sql & " AND ModuleID = " & ModuleID
    End If
    
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    
    total = 0
    count = 0
    
    Do While Not rs.EOF
        If Not IsNull(rs!Grade) Then
            total = total + rs!Grade
            count = count + 1
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    
    If count > 0 Then
        CalculateAverage = total / count
    Else
        CalculateAverage = 0
    End If
End Function

' -----------------------------------------------------------------------------
' GetStudentCount — Count students in a class
' -----------------------------------------------------------------------------
Public Function GetStudentCount(ByVal ClassID As Long) As Long
    Dim sql As String
    
    sql = "SELECT COUNT(*) AS Cnt FROM tblEnrollment WHERE ClassID = " & ClassID & _
          " AND Status = 'Active'"
    
    GetStudentCount = DLookup("Cnt", "tblEnrollment", "ClassID = " & ClassID & _
                              " AND Status = 'Active'")
End Function

' -----------------------------------------------------------------------------
' GetClassName — Get full class name
' -----------------------------------------------------------------------------
Public Function GetClassName(ByVal ClassID As Long) As String
    Dim sql As String
    Dim rs As DAO.Recordset
    
    sql = "SELECT c.ClassName, s.Name_fr, s.Name_ar, f.Code " & _
          "FROM (tblClass c INNER JOIN tblSpecialization s ON c.SpecID = s.SpecID) " & _
          "INNER JOIN tblField f ON s.FieldID = f.FieldID " & _
          "WHERE c.ClassID = " & ClassID
    
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    
    If rs.EOF Then
        GetClassName = "Unknown"
    Else
        GetClassName = rs!Code & " - " & rs!Name_fr & " (" & rs!ClassName & ")"
    End If
    
    rs.Close
End Function

' -----------------------------------------------------------------------------
' GetFieldName — Get field name in French or Arabic
' -----------------------------------------------------------------------------
Public Function GetFieldName(ByVal FieldID As Long, Optional ByVal lang As String = "fr") As String
    Dim sql As String
    
    If lang = "ar" Then
        GetFieldName = DLookup("Name_ar", "tblField", "FieldID = " & FieldID)
    Else
        GetFieldName = DLookup("Name_fr", "tblField", "FieldID = " & FieldID)
    End If
End Function

' -----------------------------------------------------------------------------
' GetLevelName — Get level name in French or Arabic
' -----------------------------------------------------------------------------
Public Function GetLevelName(ByVal LevelID As Long, Optional ByVal lang As String = "fr") As String
    If lang = "ar" Then
        GetLevelName = DLookup("Name_ar", "tblLevel", "LevelID = " & LevelID)
    Else
        GetLevelName = DLookup("Name_fr", "tblLevel", "LevelID = " & LevelID)
    End If
End Function

' -----------------------------------------------------------------------------
' GetDiplomaName — Get diploma name in French or Arabic
' -----------------------------------------------------------------------------
Public Function GetDiplomaName(ByVal DiplomaID As Long, Optional ByVal lang As String = "fr") As String
    If lang = "ar" Then
        GetDiplomaName = DLookup("Name_ar", "tblDiploma", "DiplomaID = " & DiplomaID)
    Else
        GetDiplomaName = DLookup("Name_fr", "tblDiploma", "DiplomaID = " & DiplomaID)
    End If
End Function

' -----------------------------------------------------------------------------
' GetAttendanceRate — Calculate attendance rate for a student
' -----------------------------------------------------------------------------
Public Function GetAttendanceRate(ByVal StudentID As Long, Optional ByVal ClassID As Long = 0) As Double
    Dim sql As String
    Dim rs As DAO.Recordset
    Dim total As Long
    Dim present As Long
    
    sql = "SELECT Status FROM tblAttendance WHERE StudentID = " & StudentID
    If ClassID > 0 Then
        sql = sql & " AND ClassID = " & ClassID
    End If
    
    Set rs = CurrentDb.OpenRecordset(sql, dbOpenSnapshot)
    
    total = 0
    present = 0
    
    Do While Not rs.EOF
        total = total + 1
        If rs!Status = "Present" Then
            present = present + 1
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    
    If total > 0 Then
        GetAttendanceRate = (present / total) * 100
    Else
        GetAttendanceRate = 0
    End If
End Function
