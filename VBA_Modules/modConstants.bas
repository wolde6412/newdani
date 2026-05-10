Option Explicit

' ========================================================================================
' Module: modConstants
' Purpose: Enterprise-grade centralized configuration and constants management.
'          Eliminates magic strings/numbers scattered throughout codebase.
'          Single source of truth for system configuration.
' Author: Senior Enterprise Architect
' Version: 1.0
' ========================================================================================

' --- SYSTEM METADATA ---
Public Const CONST_SYS_NAME As String = "Sunday School Management System"
Public Const CONST_SYS_VERSION As String = "1.0.0-ENTERPRISE"
Public Const CONST_SYS_AUTHOR As String = "Enterprise Development Team"
Public Const CONST_BUILD_DATE As String = "2026-05-10"

' --- WORKSHEET NAMES ---
Public Const WS_DASHBOARD As String = "Dashboard"
Public Const WS_STUDENTS As String = "Students"
Public Const WS_RESULTS As String = "Results"
Public Const WS_ATTENDANCE As String = "Attendance"
Public Const WS_TEACHERS As String = "Teachers"
Public Const WS_SCHEDULE As String = "Schedule"
Public Const WS_USERS As String = "Users"
Public Const WS_SETTINGS As String = "Settings"
Public Const WS_SUBJECTS As String = "Subjects"
Public Const WS_LOGS As String = "Logs"
Public Const WS_TEMP As String = "Temp"
Public Const WS_REPORTCARD As String = "ReportCard"

' --- TABLE NAMES ---
Public Const TBL_STUDENTS As String = "tblStudents"
Public Const TBL_USERS As String = "tblUsers"
Public Const TBL_LOGS As String = "tblLogs"
Public Const TBL_SETTINGS As String = "tblSettings"
Public Const TBL_ATTENDANCE As String = "tblAttendance"
Public Const TBL_RESULTS As String = "tblResults"
Public Const TBL_TEACHERS As String = "tblTeachers"
Public Const TBL_SCHEDULE As String = "tblSchedule"

' --- CLASS LEVELS (ACADEMIC) ---
Public Const CLASS_GRADE_1 As String = "Grade 1"
Public Const CLASS_GRADE_2 As String = "Grade 2"
Public Const CLASS_GRADE_3 As String = "Grade 3"
Public Const CLASS_GRADE_4 As String = "Grade 4"
Public Const CLASS_GRADE_5 As String = "Grade 5"
Public Const CLASS_GRADE_6 As String = "Grade 6"
Public Const CLASS_GRADE_7 As String = "Grade 7"
Public Const CLASS_GRADE_8 As String = "Grade 8"

Public Function GetAllClasses() As Variant
    ' Returns array of all class levels for use in ComboBoxes
    GetAllClasses = Array(CLASS_GRADE_1, CLASS_GRADE_2, CLASS_GRADE_3, CLASS_GRADE_4, _
                         CLASS_GRADE_5, CLASS_GRADE_6, CLASS_GRADE_7, CLASS_GRADE_8)
End Function

' --- SECTIONS ---
Public Const SECTION_A As String = "A"
Public Const SECTION_B As String = "B"
Public Const SECTION_C As String = "C"
Public Const SECTION_D As String = "D"

Public Function GetAllSections() As Variant
    GetAllSections = Array(SECTION_A, SECTION_B, SECTION_C, SECTION_D)
End Function

' --- SUBJECTS (CORE CURRICULUM) ---
Public Const SUBJECT_FAITH As String = "Faith"
Public Const SUBJECT_GEEZ As String = "Geez"
Public Const SUBJECT_HISTORY As String = "History"
Public Const SUBJECT_MORALITY As String = "Morality"
Public Const SUBJECT_LITURGY As String = "Liturgy"
Public Const SUBJECT_SAINTS As String = "Saints"

Public Function GetAllSubjects() As Variant
    GetAllSubjects = Array(SUBJECT_FAITH, SUBJECT_GEEZ, SUBJECT_HISTORY, _
                          SUBJECT_MORALITY, SUBJECT_LITURGY, SUBJECT_SAINTS)
End Function

' --- EXAM TERMS ---
Public Const EXAM_TERM_1 As String = "Term 1"
Public Const EXAM_TERM_2 As String = "Term 2"
Public Const EXAM_TERM_3 As String = "Term 3"
Public Const EXAM_FINAL As String = "Final Exam"
Public Const EXAM_MIDTERM As String = "Midterm"

Public Function GetAllExamTerms() As Variant
    GetAllExamTerms = Array(EXAM_TERM_1, EXAM_TERM_2, EXAM_TERM_3, EXAM_FINAL, EXAM_MIDTERM)
End Function

' --- ATTENDANCE STATUS ---
Public Const ATTENDANCE_PRESENT As String = "Present"
Public Const ATTENDANCE_ABSENT As String = "Absent"
Public Const ATTENDANCE_LATE As String = "Late"
Public Const ATTENDANCE_EXCUSED As String = "Excused"

Public Function GetAllAttendanceStatuses() As Variant
    GetAllAttendanceStatuses = Array(ATTENDANCE_PRESENT, ATTENDANCE_ABSENT, _
                                    ATTENDANCE_LATE, ATTENDANCE_EXCUSED)
End Function

' --- GENDER ---
Public Const GENDER_MALE As String = "Male"
Public Const GENDER_FEMALE As String = "Female"
Public Const GENDER_OTHER As String = "Other"

Public Function GetAllGenders() As Variant
    GetAllGenders = Array(GENDER_MALE, GENDER_FEMALE, GENDER_OTHER)
End Function

' --- DAYS OF WEEK ---
Public Const DAY_MONDAY As String = "Monday"
Public Const DAY_TUESDAY As String = "Tuesday"
Public Const DAY_WEDNESDAY As String = "Wednesday"
Public Const DAY_THURSDAY As String = "Thursday"
Public Const DAY_FRIDAY As String = "Friday"
Public Const DAY_SATURDAY As String = "Saturday"
Public Const DAY_SUNDAY As String = "Sunday"

Public Function GetAllDaysOfWeek() As Variant
    GetAllDaysOfWeek = Array(DAY_MONDAY, DAY_TUESDAY, DAY_WEDNESDAY, DAY_THURSDAY, _
                            DAY_FRIDAY, DAY_SATURDAY, DAY_SUNDAY)
End Function

' --- GRADING SCALE ---
Public Const GRADE_A As String = "A"
Public Const GRADE_B As String = "B"
Public Const GRADE_C As String = "C"
Public Const GRADE_D As String = "D"
Public Const GRADE_F As String = "F"

Public Const GRADE_A_MIN As Double = 90
Public Const GRADE_B_MIN As Double = 80
Public Const GRADE_C_MIN As Double = 70
Public Const GRADE_D_MIN As Double = 60
Public Const GRADE_F_MIN As Double = 0

' --- SCORE RANGES ---
Public Const SCORE_MIN As Double = 0
Public Const SCORE_MAX As Double = 100
Public Const SUBJECT_COUNT As Integer = 6 ' Faith, Geez, History, Morality, Liturgy, Saints

' --- SECURITY & SESSIONS ---
Public Const SESSION_TIMEOUT_MINUTES As Integer = 30
Public Const MAX_LOGIN_ATTEMPTS As Integer = 5
Public Const LOCKOUT_DURATION_MINUTES As Integer = 15
Public Const PASSWORD_MIN_LENGTH As Integer = 8
Public Const PASSWORD_REQUIRE_UPPERCASE As Boolean = True
Public Const PASSWORD_REQUIRE_LOWERCASE As Boolean = True
Public Const PASSWORD_REQUIRE_NUMBERS As Boolean = True
Public Const PASSWORD_REQUIRE_SPECIAL As Boolean = False

' --- PAGINATION ---
Public Const PAGE_SIZE_SMALL As Integer = 10
Public Const PAGE_SIZE_MEDIUM As Integer = 25
Public Const PAGE_SIZE_LARGE As Integer = 50
Public Const DEFAULT_PAGE_SIZE As Integer = PAGE_SIZE_MEDIUM

' --- PERFORMANCE THRESHOLDS ---
Public Const QUERY_TIMEOUT_SECONDS As Integer = 30
Public Const SLOW_QUERY_THRESHOLD_MS As Long = 500
Public Const BULK_OPERATION_BATCH_SIZE As Integer = 100

' --- UI TIMING (MILLISECONDS) ---
Public Const UI_TRANSITION_DURATION As Integer = 300
Public Const UI_TOAST_DURATION As Integer = 3000
Public Const UI_FORM_LOAD_TIMEOUT As Integer = 5000

' --- DEFAULT ADMIN CREDENTIALS ---
Public Const DEFAULT_ADMIN_USERNAME As String = "admin"
Public Const DEFAULT_ADMIN_PASSWORD_HASH As String = "" ' Will be generated; never store plaintext

' --- ERROR CODES (Enterprise Standard) ---
Public Const ERR_VALIDATION_FAILED As Integer = 1001
Public Const ERR_DATABASE_NOT_FOUND As Integer = 1002
Public Const ERR_RECORD_NOT_FOUND As Integer = 1003
Public Const ERR_DUPLICATE_ENTRY As Integer = 1004
Public Const ERR_PERMISSION_DENIED As Integer = 1005
Public Const ERR_SESSION_EXPIRED As Integer = 1006
Public Const ERR_INVALID_CREDENTIALS As Integer = 1007
Public Const ERR_SYSTEM_LOCKED As Integer = 1008
Public Const ERR_DATA_INTEGRITY As Integer = 1009
Public Const ERR_OPERATION_TIMEOUT As Integer = 1010
Public Const ERR_RESOURCE_UNAVAILABLE As Integer = 1011
Public Const ERR_UNKNOWN As Integer = 9999

' --- LOG LEVELS ---
Public Enum LogLevel
    DEBUG = 1
    INFO = 2
    WARNING = 3
    ERROR = 4
    CRITICAL = 5
End Enum

' --- AUDIT ACTION TYPES ---
Public Const ACTION_LOGIN As String = "LOGIN"
Public Const ACTION_LOGOUT As String = "LOGOUT"
Public Const ACTION_CREATE As String = "CREATE"
Public Const ACTION_UPDATE As String = "UPDATE"
Public Const ACTION_DELETE As String = "DELETE"
Public Const ACTION_EXPORT As String = "EXPORT"
Public Const ACTION_IMPORT As String = "IMPORT"
Public Const ACTION_PRINT As String = "PRINT"
Public Const ACTION_BACKUP As String = "BACKUP"
Public Const ACTION_RESTORE As String = "RESTORE"
Public Const ACTION_CONFIG_CHANGE As String = "CONFIG_CHANGE"
Public Const ACTION_PERMISSION_CHANGE As String = "PERMISSION_CHANGE"

' --- FEATURE FLAGS (For A/B Testing & Gradual Rollout) ---
Public Const FEATURE_DARK_MODE_ENABLED As Boolean = False
Public Const FEATURE_BULK_IMPORT_ENABLED As Boolean = True
Public Const FEATURE_ADVANCED_SEARCH_ENABLED As Boolean = True
Public Const FEATURE_ANALYTICS_DASHBOARD_ENABLED As Boolean = False
Public Const FEATURE_EMAIL_NOTIFICATIONS_ENABLED As Boolean = False
Public Const FEATURE_SMS_NOTIFICATIONS_ENABLED As Boolean = False
Public Const FEATURE_OFFLINE_MODE_ENABLED As Boolean = False

' --- CACHING SETTINGS ---
Public Const CACHE_ENABLED As Boolean = True
Public Const CACHE_TTL_SECONDS As Integer = 300 ' 5 minutes
Public Const CACHE_MAX_ITEMS As Integer = 1000
Public Const CACHE_EVICTION_POLICY As String = "LRU" ' Least Recently Used

Public Function GetConstantValue(ByVal ConstantName As String) As Variant
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Runtime lookup of constants by name (useful for UI-driven config)
    ' USAGE: value = GetConstantValue("CLASS_GRADE_1")
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    
    Select Case ConstantName
        Case "SYS_NAME": GetConstantValue = CONST_SYS_NAME
        Case "SYS_VERSION": GetConstantValue = CONST_SYS_VERSION
        Case "SCORE_MIN": GetConstantValue = SCORE_MIN
        Case "SCORE_MAX": GetConstantValue = SCORE_MAX
        Case "SESSION_TIMEOUT_MINUTES": GetConstantValue = SESSION_TIMEOUT_MINUTES
        Case Else: GetConstantValue = ""
    End Select
End Function
