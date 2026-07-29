Option Compare Database
Option Explicit

' =============================================================================
' mod_CreateTables — SIS MVP Table Creation Module
' School: Allal Badr El-Din (El Bayadh)
' =============================================================================

Public Sub CreateAllTables()
    On Error GoTo ErrHandler
    
    Debug.Print "Creating SIS MVP tables..."
    
    ' Create tables in order (respecting foreign key dependencies)
    Call CreatetblLevel
    Call CreatetblField
    Call CreatetblDiploma
    Call CreatetblSpecialization
    Call CreatetblModule
    Call CreatetblTeacher
    Call CreatetblStudent
    Call CreatetblClass
    Call CreatetblEnrollment
    Call CreatetblGrade
    Call CreatetblAttendance
    Call CreatetblFee
    Call CreatetblConfig
    
    Debug.Print "All tables created successfully!"
    MsgBox "Tables created successfully!", vbInformation, "SIS MVP"
    
    Exit Sub
    
ErrHandler:
    Debug.Print "Error in CreateAllTables: " & Err.Description
    MsgBox "Error creating tables: " & Err.Description, vbCritical, "SIS MVP"
End Sub

' -----------------------------------------------------------------------------
' tblLevel — Academic Levels (1AM, 2AM, 3AM, 4AM, BAC)
' -----------------------------------------------------------------------------
Private Sub CreatetblLevel()
    Dim sql As String
    
    If DCount("*", "MSysObjects", "Name='tblLevel'") > 0 Then
        Debug.Print "tblLevel already exists, skipping."
        Exit Sub
    End If
    
    sql = "CREATE TABLE tblLevel (" & _
          "LevelID AUTOINCREMENT PRIMARY KEY, " & _
          "Name_fr TEXT(50) NOT NULL, " & _
          "Name_ar TEXT(50) NOT NULL, " & _
          "Cycle TEXT(20) NOT NULL, " & _
          "OrderIndex INTEGER NOT NULL)"
    
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "tblLevel created."
    
    ' Seed data
    CurrentDb.Execute "INSERT INTO tblLevel (Name_fr, Name_ar, Cycle, OrderIndex) VALUES ('1ère année moyenne', 'الأولى متوسط', 'Middle', 1)"
    CurrentDb.Execute "INSERT INTO tblLevel (Name_fr, Name_ar, Cycle, OrderIndex) VALUES ('2ème année moyenne', 'الثانية متوسط', 'Middle', 2)"
    CurrentDb.Execute "INSERT INTO tblLevel (Name_fr, Name_ar, Cycle, OrderIndex) VALUES ('3ème année moyenne', 'الثالثة متوسط', 'Middle', 3)"
    CurrentDb.Execute "INSERT INTO tblLevel (Name_fr, Name_ar, Cycle, OrderIndex) VALUES ('4ème année moyenne', 'الرابعة ثانوي', 'Secondary', 4)"
    CurrentDb.Execute "INSERT INTO tblLevel (Name_fr, Name_ar, Cycle, OrderIndex) VALUES ('Baccalauréat', 'البكالوريا', 'Secondary', 5)"
    
    Debug.Print "tblLevel seeded."
End Sub

' -----------------------------------------------------------------------------
' tblField — Vocational Fields (IT, Commerce, Electricity, Mechanics)
' -----------------------------------------------------------------------------
Private Sub CreatetblField()
    Dim sql As String
    
    If DCount("*", "MSysObjects", "Name='tblField'") > 0 Then
        Debug.Print "tblField already exists, skipping."
        Exit Sub
    End If
    
    sql = "CREATE TABLE tblField (" & _
          "FieldID AUTOINCREMENT PRIMARY KEY, " & _
          "Name_fr TEXT(50) NOT NULL, " & _
          "Name_ar TEXT(50) NOT NULL, " & _
          "Code TEXT(10) NOT NULL)"
    
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "tblField created."
    
    ' Seed data
    CurrentDb.Execute "INSERT INTO tblField (Name_fr, Name_ar, Code) VALUES ('Informatique et Numérique', 'المعلوماتية والرقمية', 'IT')"
    CurrentDb.Execute "INSERT INTO tblField (Name_fr, Name_ar, Code) VALUES ('Commerce', 'التجارة', 'COM')"
    CurrentDb.Execute "INSERT INTO tblField (Name_fr, Name_ar, Code) VALUES ('Électricité', 'الكهرباء', 'ELEC')"
    CurrentDb.Execute "INSERT INTO tblField (Name_fr, Name_ar, Code) VALUES ('Mécanique', 'الميكانيكا', 'MECH')"
    
    Debug.Print "tblField seeded."
End Sub

' -----------------------------------------------------------------------------
' tblDiploma — Diploma Types (CAP, BEP, BTS)
' -----------------------------------------------------------------------------
Private Sub CreatetblDiploma()
    Dim sql As String
    
    If DCount("*", "MSysObjects", "Name='tblDiploma'") > 0 Then
        Debug.Print "tblDiploma already exists, skipping."
        Exit Sub
    End If
    
    sql = "CREATE TABLE tblDiploma (" & _
          "DiplomaID AUTOINCREMENT PRIMARY KEY, " & _
          "Name_fr TEXT(100) NOT NULL, " & _
          "Name_ar TEXT(100) NOT NULL, " & _
          "Abbreviation TEXT(10) NOT NULL, " & _
          "DurationMonths INTEGER NOT NULL, " & _
          "LevelRequired INTEGER NOT NULL)"
    
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "tblDiploma created."
    
    ' Seed data
    CurrentDb.Execute "INSERT INTO tblDiploma (Name_fr, Name_ar, Abbreviation, DurationMonths, LevelRequired) VALUES ('Certificat d''Aptitude Professionnelle', 'شهادة الكفاءة المهنية', 'CAP', 12, 3)"
    CurrentDb.Execute "INSERT INTO tblDiploma (Name_fr, Name_ar, Abbreviation, DurationMonths, LevelRequired) VALUES ('Brevet d''Études Professionnelles', 'شهادة دراسات مهنية', 'BEP', 24, 4)"
    CurrentDb.Execute "INSERT INTO tblDiploma (Name_fr, Name_ar, Abbreviation, DurationMonths, LevelRequired) VALUES ('Brevet de Technicien Supérieur', 'شهادة تقني سام', 'BTS', 24, 5)"
    
    Debug.Print "tblDiploma seeded."
End Sub

' -----------------------------------------------------------------------------
' tblSpecialization — 22 Specializations across 4 fields
' -----------------------------------------------------------------------------
Private Sub CreatetblSpecialization()
    Dim sql As String
    
    If DCount("*", "MSysObjects", "Name='tblSpecialization'") > 0 Then
        Debug.Print "tblSpecialization already exists, skipping."
        Exit Sub
    End If
    
    sql = "CREATE TABLE tblSpecialization (" & _
          "SpecID TEXT(10) PRIMARY KEY, " & _
          "FieldID INTEGER NOT NULL, " & _
          "Name_fr TEXT(100) NOT NULL, " & _
          "Name_ar TEXT(100) NOT NULL, " & _
          "DiplomaID INTEGER NOT NULL, " & _
          "DurationMonths INTEGER NOT NULL)"
    
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "tblSpecialization created."
    
    ' IT specializations
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('IT-01', 1, 'Assistant administrateur', 'مساعد إداري', 1, 12)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('IT-02', 1, 'Développeur web', 'مطور ويب', 1, 12)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('IT-03', 1, 'Maintenance informatique', 'صيانة الحاسوب', 1, 12)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('IT-04', 1, 'Intégrateur web', 'متكامل ويب', 2, 24)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('IT-05', 1, 'Administrateur systèmes et réseaux', 'مسؤول الأنظمة والشبكات', 2, 24)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('IT-06', 1, 'Technicien en cybersécurité', 'تقني في الأمن السيبراني', 3, 24)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('IT-07', 1, 'Développeur d''applications', 'مطور تطبيقات', 3, 24)"
    
    ' Commerce specializations
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('COM-01', 2, 'Comptable', 'محاسب', 1, 12)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('COM-02', 2, 'Assistant commercial', 'مساعد تجاري', 1, 12)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('COM-03', 2, 'Secrétaire assistant(e)', 'سكرتير تنفيذي', 1, 12)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('COM-04', 2, 'Commerce international', 'التجارة الدولية', 2, 24)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('COM-05', 2, 'Gestion des entreprises', 'إدارة المؤسسات', 2, 24)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('COM-06', 2, 'Marketing et commerce', 'التسويق والتجارة', 3, 24)"
    
    ' Electricity specializations
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('ELEC-01', 3, 'Installateur électricien', 'كهربائي تركيب', 1, 12)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('ELEC-02', 3, 'Électricien de maintenance', 'كهربائي صيانة', 1, 12)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('ELEC-03', 3, 'Électrotechnique', 'الكهرباء الصناعية', 2, 24)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('ELEC-04', 3, 'Électronique industrielle', 'الإلكترونيك الصناعي', 2, 24)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('ELEC-05', 3, 'Froid et climatisation', 'التكييف والتبريد', 3, 24)"
    
    ' Mechanics specializations
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('MECH-01', 4, 'Mécanique auto', 'ميكانيكا السيارات', 1, 12)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('MECH-02', 4, 'Soudure', 'اللحام', 1, 12)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('MECH-03', 4, 'Mécanique industrielle', 'الميكانيك الصناعي', 2, 24)"
    CurrentDb.Execute "INSERT INTO tblSpecialization (SpecID, FieldID, Name_fr, Name_ar, DiplomaID, DurationMonths) VALUES ('MECH-04', 4, 'Maintenance industrielle', 'الصيانة الصناعية', 3, 24)"
    
    Debug.Print "tblSpecialization seeded (22 specializations)."
End Sub

' -----------------------------------------------------------------------------
' tblModule — Modules per Specialization
' -----------------------------------------------------------------------------
Private Sub CreatetblModule()
    Dim sql As String
    
    If DCount("*", "MSysObjects", "Name='tblModule'") > 0 Then
        Debug.Print "tblModule already exists, skipping."
        Exit Sub
    End If
    
    sql = "CREATE TABLE tblModule (" & _
          "ModuleID AUTOINCREMENT PRIMARY KEY, " & _
          "SpecID TEXT(10) NOT NULL, " & _
          "Name_fr TEXT(100) NOT NULL, " & _
          "Name_ar TEXT(100) NOT NULL, " & _
          "Hours INTEGER NOT NULL, " & _
          "OrderIndex INTEGER NOT NULL)"
    
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "tblModule created."
    
    ' IT-01 modules
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('IT-01', 'Informatique de base', 'أساسيات الحاسوب', 60, 1)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('IT-01', 'Réseaux', 'الشبكات', 60, 2)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('IT-01', 'Systèmes d''exploitation', 'أنظمة التشغيل', 60, 3)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('IT-01', 'Administration', 'الإدارة', 60, 4)"
    
    ' IT-02 modules
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('IT-02', 'HTML/CSS', 'HTML/CSS', 60, 1)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('IT-02', 'JavaScript', 'جافا سكربت', 60, 2)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('IT-02', 'PHP', 'بي إتش بي', 60, 3)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('IT-02', 'Bases de données', 'قواعد البيانات', 60, 4)"
    
    ' IT-03 modules
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('IT-03', 'Matériel informatique', 'المعدات الحاسوبية', 60, 1)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('IT-03', 'Diagnostic', 'التشخيص', 60, 2)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('IT-03', 'Réparation', 'الإصلاح', 60, 3)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('IT-03', 'Prévention', 'الوقاية', 60, 4)"
    
    ' COM-01 modules
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('COM-01', 'Comptabilité générale', 'المحاسبة العامة', 60, 1)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('COM-01', 'Facturation', 'الفوترة', 60, 2)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('COM-01', 'Paie', 'الأجور', 60, 3)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('COM-01', 'Excel', 'إكسل', 60, 4)"
    
    ' ELEC-01 modules
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('ELEC-01', 'Électricité générale', 'الكهرباء العامة', 60, 1)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('ELEC-01', 'Câblage', 'الأسلاك', 60, 2)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('ELEC-01', 'Normes', 'المعايير', 60, 3)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('ELEC-01', 'Sécurité', 'الأمان', 60, 4)"
    
    ' MECH-01 modules
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('MECH-01', 'Moteur', 'المحرك', 60, 1)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('MECH-01', 'Transmission', 'النقل', 60, 2)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('MECH-01', 'Freinage', 'الكبح', 60, 3)"
    CurrentDb.Execute "INSERT INTO tblModule (SpecID, Name_fr, Name_ar, Hours, OrderIndex) VALUES ('MECH-01', 'Diagnostic', 'التشخيص', 60, 4)"
    
    Debug.Print "tblModule seeded (24 modules for 6 specializations)."
End Sub

' -----------------------------------------------------------------------------
' tblTeacher — Teachers/Instructors
' -----------------------------------------------------------------------------
Private Sub CreatetblTeacher()
    Dim sql As String
    
    If DCount("*", "MSysObjects", "Name='tblTeacher'") > 0 Then
        Debug.Print "tblTeacher already exists, skipping."
        Exit Sub
    End If
    
    sql = "CREATE TABLE tblTeacher (" & _
          "TeacherID AUTOINCREMENT PRIMARY KEY, " & _
          "LastName_ar TEXT(50) NOT NULL, " & _
          "FirstName_ar TEXT(50) NOT NULL, " & _
          "LastName_fr TEXT(50) NOT NULL, " & _
          "FirstName_fr TEXT(50) NOT NULL, " & _
          "Phone TEXT(20), " & _
          "Email TEXT(100), " & _
          "Specialty TEXT(100), " & _
          "HireDate DATE/TIME, " & _
          "Status TEXT(20) DEFAULT 'Active')"
    
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "tblTeacher created."
End Sub

' -----------------------------------------------------------------------------
' tblStudent — Students
' -----------------------------------------------------------------------------
Private Sub CreatetblStudent()
    Dim sql As String
    
    If DCount("*", "MSysObjects", "Name='tblStudent'") > 0 Then
        Debug.Print "tblStudent already exists, skipping."
        Exit Sub
    End If
    
    sql = "CREATE TABLE tblStudent (" & _
          "StudentID AUTOINCREMENT PRIMARY KEY, " & _
          "Matricule TEXT(20) NOT NULL, " & _
          "LastName_ar TEXT(50) NOT NULL, " & _
          "FirstName_ar TEXT(50) NOT NULL, " & _
          "LastName_fr TEXT(50) NOT NULL, " & _
          "FirstName_fr TEXT(50) NOT NULL, " & _
          "DateOfBirth DATE/TIME, " & _
          "PlaceOfBirth TEXT(50), " & _
          "Gender TEXT(1), " & _
          "Phone TEXT(20), " & _
          "Address TEXT(100), " & _
          "EnrollmentDate DATE/TIME, " & _
          "LevelID INTEGER, " & _
          "SpecID TEXT(10), " & _
          "Status TEXT(20) DEFAULT 'Active')"
    
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "tblStudent created."
End Sub

' -----------------------------------------------------------------------------
' tblClass — Class Sections
' -----------------------------------------------------------------------------
Private Sub CreatetblClass()
    Dim sql As String
    
    If DCount("*", "MSysObjects", "Name='tblClass'") > 0 Then
        Debug.Print "tblClass already exists, skipping."
        Exit Sub
    End If
    
    sql = "CREATE TABLE tblClass (" & _
          "ClassID AUTOINCREMENT PRIMARY KEY, " & _
          "ClassName TEXT(50) NOT NULL, " & _
          "LevelID INTEGER NOT NULL, " & _
          "SpecID TEXT(10) NOT NULL, " & _
          "AcademicYear TEXT(10) NOT NULL, " & _
          "TeacherID INTEGER, " & _
          "MaxCapacity INTEGER DEFAULT 30)"
    
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "tblClass created."
End Sub

' -----------------------------------------------------------------------------
' tblEnrollment — Student-Class Links
' -----------------------------------------------------------------------------
Private Sub CreatetblEnrollment()
    Dim sql As String
    
    If DCount("*", "MSysObjects", "Name='tblEnrollment'") > 0 Then
        Debug.Print "tblEnrollment already exists, skipping."
        Exit Sub
    End If
    
    sql = "CREATE TABLE tblEnrollment (" & _
          "EnrollmentID AUTOINCREMENT PRIMARY KEY, " & _
          "StudentID INTEGER NOT NULL, " & _
          "ClassID INTEGER NOT NULL, " & _
          "AcademicYear TEXT(10) NOT NULL, " & _
          "Status TEXT(20) DEFAULT 'Active')"
    
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "tblEnrollment created."
End Sub

' -----------------------------------------------------------------------------
' tblGrade — Grades/Marks
' -----------------------------------------------------------------------------
Private Sub CreatetblGrade()
    Dim sql As String
    
    If DCount("*", "MSysObjects", "Name='tblGrade'") > 0 Then
        Debug.Print "tblGrade already exists, skipping."
        Exit Sub
    End If
    
    sql = "CREATE TABLE tblGrade (" & _
          "GradeID AUTOINCREMENT PRIMARY KEY, " & _
          "StudentID INTEGER NOT NULL, " & _
          "ModuleID INTEGER NOT NULL, " & _
          "ExamType TEXT(20) NOT NULL, " & _
          "Grade DOUBLE, " & _
          "Comment TEXT(200), " & _
          "DateRecorded DATE/TIME)"
    
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "tblGrade created."
End Sub

' -----------------------------------------------------------------------------
' tblAttendance — Attendance Tracking
' -----------------------------------------------------------------------------
Private Sub CreatetblAttendance()
    Dim sql As String
    
    If DCount("*", "MSysObjects", "Name='tblAttendance'") > 0 Then
        Debug.Print "tblAttendance already exists, skipping."
        Exit Sub
    End If
    
    sql = "CREATE TABLE tblAttendance (" & _
          "AttendanceID AUTOINCREMENT PRIMARY KEY, " & _
          "StudentID INTEGER NOT NULL, " & _
          "ClassID INTEGER NOT NULL, " & _
          "AttendanceDate DATE/TIME NOT NULL, " & _
          "Status TEXT(10) NOT NULL, " & _
          "Note TEXT(100))"
    
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "tblAttendance created."
End Sub

' -----------------------------------------------------------------------------
' tblFee — Fee Management
' -----------------------------------------------------------------------------
Private Sub CreatetblFee()
    Dim sql As String
    
    If DCount("*", "MSysObjects", "Name='tblFee'") > 0 Then
        Debug.Print "tblFee already exists, skipping."
        Exit Sub
    End If
    
    sql = "CREATE TABLE tblFee (" & _
          "FeeID AUTOINCREMENT PRIMARY KEY, " & _
          "StudentID INTEGER NOT NULL, " & _
          "FeeType TEXT(50) NOT NULL, " & _
          "Amount CURRENCY, " & _
          "DueDate DATE/TIME, " & _
          "PaidDate DATE/TIME, " & _
          "Status TEXT(20) DEFAULT 'Pending', " & _
          "ReceiptNo TEXT(20))"
    
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "tblFee created."
End Sub

' -----------------------------------------------------------------------------
' tblConfig — System Configuration
' -----------------------------------------------------------------------------
Private Sub CreatetblConfig()
    Dim sql As String
    
    If DCount("*", "MSysObjects", "Name='tblConfig'") > 0 Then
        Debug.Print "tblConfig already exists, skipping."
        Exit Sub
    End If
    
    sql = "CREATE TABLE tblConfig (" & _
          "ConfigKey TEXT(50) PRIMARY KEY, " & _
          "ConfigValue TEXT(200), " & _
          "Description TEXT(200))"
    
    CurrentDb.Execute sql, dbFailOnError
    Debug.Print "tblConfig created."
    
    ' Seed configuration
    CurrentDb.Execute "INSERT INTO tblConfig (ConfigKey, ConfigValue, Description) VALUES ('SchoolName_ar', 'المؤسسة الخاصة علال بدر الدين', 'School name in Arabic')"
    CurrentDb.Execute "INSERT INTO tblConfig (ConfigKey, ConfigValue, Description) VALUES ('SchoolName_fr', 'Institut Privé Allal Badr El-Din', 'School name in French')"
    CurrentDb.Execute "INSERT INTO tblConfig (ConfigKey, ConfigValue, Description) VALUES ('AcademicYear', '2026-2027', 'Current academic year')"
    CurrentDb.Execute "INSERT INTO tblConfig (ConfigKey, ConfigValue, Description) VALUES ('Location', 'El Bayadh, Algeria', 'School location')"
    
    Debug.Print "tblConfig seeded."
End Sub
