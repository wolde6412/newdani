Option Explicit

' ========================================================================================
' UserForm: frmDashboard (MODERNIZED - ENTERPRISE EDITION)
' Purpose: Central navigation hub and application dashboard.
'          Displays welcome information, provides access to all system modules,
'          and serves as the main navigation switchboard.
' Author: Senior Enterprise Architect
' Version: 2.0 (Modernized)
' ========================================================================================

' --- FORM-LEVEL VARIABLES ---
Private m_FormInitialized As Boolean

' --- UI STATE COLORS (Cache for performance) ---
Private m_ColorPrimary As Long
Private m_ColorSecondary As Long
Private m_ColorText As Long
Private m_ColorBackground As Long

' ========================================================================================
' FORM INITIALIZATION
' ========================================================================================

Private Sub UserForm_Initialize()
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Initializes the dashboard with modern styling and layout.
    ' ------------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    ' 1. Initialize framework
    Call modUIFramework.InitializeBaseForm(Me, "frmDashboard")
    
    ' 2. Cache theme colors
    m_ColorPrimary = modUITheme.GetPrimaryColor()
    m_ColorSecondary = modUITheme.GetAccentColor()
    m_ColorText = modUITheme.GetTextColor()
    m_ColorBackground = modUITheme.GetBackgroundColor()
    
    ' 3. Configure form appearance
    With Me
        .BackColor = m_ColorBackground
        .Width = 900 ' Wider dashboard
        .Height = 700 ' Tall for content
        .Left = (Application.Width - .Width) / 2
        .Top = (Application.Height - .Height) / 2
        .Caption = CONST_SYS_NAME & " - Dashboard"
    End With
    
    ' 4. Initialize controls
    Call InitializeControls
    
    ' 5. Load dashboard data
    Call LoadDashboardData
    
    m_FormInitialized = True
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Dashboard Initialization Error: " & Err.Description, vbCritical
    Unload Me
End Sub

Private Sub InitializeControls()
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Configures dashboard controls with professional layout.
    ' ------------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    ' --- HEADER SECTION ---
    With Me.lblTitle
        .Caption = "Dashboard"
        .Top = 12
        .Left = 12
        .Width = 850
        .Height = 28
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_HEADING_1
        .Font.Bold = True
        .ForeColor = m_ColorPrimary
        .BackColor = m_ColorBackground
    End With
    
    ' --- WELCOME LABEL ---
    With Me.lblWelcome
        .Caption = "Welcome, " & g_CurrentUsername & "!"
        .Top = 45
        .Left = 12
        .Width = 500
        .Height = 20
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_HEADING_2
        .Font.Bold = False
        .ForeColor = m_ColorText
        .BackColor = m_ColorBackground
    End With
    
    ' --- USER INFO SECTION ---
    With Me.lblRole
        .Caption = "Role: " & GetRoleDisplayName(g_CurrentUserRole)
        .Top = 70
        .Left = 12
        .Width = 300
        .Height = 16
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_BODY
        .ForeColor = modUITheme.COLOR_MEDIUM_GRAY
        .BackColor = m_ColorBackground
    End With
    
    With Me.lblDate
        .Caption = Format(Now, "dddd, mmmm dd, yyyy | hh:mm AM/PM")
        .Top = 70
        .Left = 600
        .Width = 270
        .Height = 16
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_BODY
        .ForeColor = modUITheme.COLOR_MEDIUM_GRAY
        .BackColor = m_ColorBackground
        .Alignment = 1 ' Right-align
    End With
    
    ' --- NAVIGATION TITLE ---
    With Me.lblModulesTitle
        .Caption = "Quick Access"
        .Top = 100
        .Left = 12
        .Width = 850
        .Height = 20
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_HEADING_2
        .Font.Bold = True
        .ForeColor = m_ColorPrimary
        .BackColor = m_ColorBackground
    End With
    
    ' --- MODULE BUTTONS (Primary Navigation) ---
    ' Note: These buttons are styled as navigation cards in modern design
    
    ' STUDENTS MODULE
    With Me.cmdStudents
        .Caption = "Students"
        .Top = 135
        .Left = 12
        .Width = 200
        .Height = 50
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_HEADING_3
        .Font.Bold = True
        Call modUITheme.StyleButtonPrimary(Me.cmdStudents)
    End With
    
    ' TEACHERS MODULE
    With Me.cmdTeachers
        .Caption = "Teachers"
        .Top = 135
        .Left = 230
        .Width = 200
        .Height = 50
        Call modUITheme.StyleButtonPrimary(Me.cmdTeachers)
    End With
    
    ' ATTENDANCE MODULE
    With Me.cmdAttendance
        .Caption = "Attendance"
        .Top = 135
        .Left = 448
        .Width = 200
        .Height = 50
        Call modUITheme.StyleButtonPrimary(Me.cmdAttendance)
    End With
    
    ' RESULTS MODULE
    With Me.cmdResults
        .Caption = "Results"
        .Top = 135
        .Left = 666
        .Width = 196
        .Height = 50
        Call modUITheme.StyleButtonPrimary(Me.cmdResults)
    End With
    
    ' SCHEDULE MODULE (Second row)
    With Me.cmdSchedule
        .Caption = "Schedule"
        .Top = 200
        .Left = 12
        .Width = 200
        .Height = 50
        Call modUITheme.StyleButtonPrimary(Me.cmdSchedule)
    End With
    
    ' REPORT CARD MODULE
    With Me.cmdReportCard
        .Caption = "Report Cards"
        .Top = 200
        .Left = 230
        .Width = 200
        .Height = 50
        Call modUITheme.StyleButtonPrimary(Me.cmdReportCard)
    End With
    
    ' SETTINGS MODULE
    With Me.cmdSettings
        .Caption = "Settings"
        .Top = 200
        .Left = 448
        .Width = 200
        .Height = 50
        Call modUITheme.StyleButtonSecondary(Me.cmdSettings)
    End With
    
    ' SYSTEM INFO
    With Me.cmdSystemInfo
        .Caption = "System Info"
        .Top = 200
        .Left = 666
        .Width = 196
        .Height = 50
        Call modUITheme.StyleButtonSecondary(Me.cmdSystemInfo)
    End With
    
    ' --- STATISTICS SECTION (KPI Cards) ---
    Call InitializeStatisticsSection
    
    ' --- LOGOUT BUTTON (Bottom, Danger Style) ---
    With Me.cmdLogout
        .Caption = "Sign Out"
        .Top = 625
        .Left = 12
        .Width = 850
        .Height = modUITheme.BUTTON_HEIGHT
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_BODY
        .Font.Bold = True
        Call modUITheme.StyleButtonSecondary(Me.cmdLogout)
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error Initializing Dashboard Controls: " & Err.Description, vbCritical
End Sub

Private Sub InitializeStatisticsSection()
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Sets up KPI/statistics cards below quick access buttons.
    ' NOTE: Placeholder for Phase B when we add actual data fetching
    ' ------------------------------------------------------------------------------------
    With Me.lblStatsTitle
        .Caption = "Statistics"
        .Top = 270
        .Left = 12
        .Width = 850
        .Height = 20
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_HEADING_2
        .Font.Bold = True
        .ForeColor = modUITheme.GetPrimaryColor()
        .BackColor = modUITheme.GetBackgroundColor()
    End With
    
    ' STAT: Total Students
    With Me.lblStatStudents
        .Caption = "Total Students: --"
        .Top = 305
        .Left = 12
        .Width = 200
        .Height = 40
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_BODY
        .BackColor = modUITheme.HexToLong("#F5F5F5")
        .BorderStyle = 1
        .BorderColor = modUITheme.COLOR_BORDER
    End With
    
    ' STAT: Total Teachers
    With Me.lblStatTeachers
        .Caption = "Total Teachers: --"
        .Top = 305
        .Left = 230
        .Width = 200
        .Height = 40
        Call modUITheme.StyleLabelBody(Me.lblStatTeachers)
        .BackColor = modUITheme.HexToLong("#F5F5F5")
        .BorderStyle = 1
        .BorderColor = modUITheme.COLOR_BORDER
    End With
    
    ' STAT: Today's Attendance
    With Me.lblStatAttendance
        .Caption = "Today's Attendance: --"
        .Top = 305
        .Left = 448
        .Width = 200
        .Height = 40
        .BackColor = modUITheme.HexToLong("#F5F5F5")
        .BorderStyle = 1
        .BorderColor = modUITheme.COLOR_BORDER
    End With
    
    ' STAT: Pending Tasks
    With Me.lblStatPending
        .Caption = "Pending Tasks: --"
        .Top = 305
        .Left = 666
        .Width = 196
        .Height = 40
        .BackColor = modUITheme.HexToLong("#F5F5F5")
        .BorderStyle = 1
        .BorderColor = modUITheme.COLOR_BORDER
    End With
End Sub

Private Sub LoadDashboardData()
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Loads dashboard statistics from database.
    ' NOTE: Placeholder for Phase B - will fetch actual data
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    
    ' TODO: Implement statistics loading
    ' - Count total students
    ' - Count total teachers
    ' - Calculate today's attendance percentage
    ' - Count pending report cards
    
    Debug.Print "Dashboard data loaded"
End Sub

Private Function GetRoleDisplayName(ByVal roleEnum As UserRole) As String
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Converts role enum to readable display name.
    ' ------------------------------------------------------------------------------------
    Select Case roleEnum
        Case UserRole.Admin
            GetRoleDisplayName = "Administrator"
        Case UserRole.Teacher
            GetRoleDisplayName = "Teacher"
        Case UserRole.Viewer
            GetRoleDisplayName = "Viewer"
        Case Else
            GetRoleDisplayName = "Unknown"
    End Select
End Function

' ========================================================================================
' NAVIGATION EVENT HANDLERS
' ========================================================================================

Private Sub cmdStudents_Click()
    ' Launch Student Management module
    On Error GoTo ErrorHandler
    frmStudents.Show vbModal
    Exit Sub
ErrorHandler:
    MsgBox "Error opening Students module: " & Err.Description, vbCritical
End Sub

Private Sub cmdTeachers_Click()
    ' Launch Teacher Management module
    On Error GoTo ErrorHandler
    frmTeachers.Show vbModal
    Exit Sub
ErrorHandler:
    MsgBox "Error opening Teachers module: " & Err.Description, vbCritical
End Sub

Private Sub cmdAttendance_Click()
    ' Launch Attendance Tracking module
    On Error GoTo ErrorHandler
    frmAttendance.Show vbModal
    Exit Sub
ErrorHandler:
    MsgBox "Error opening Attendance module: " & Err.Description, vbCritical
End Sub

Private Sub cmdResults_Click()
    ' Launch Results Entry module
    On Error GoTo ErrorHandler
    frmResults.Show vbModal
    Exit Sub
ErrorHandler:
    MsgBox "Error opening Results module: " & Err.Description, vbCritical
End Sub

Private Sub cmdSchedule_Click()
    ' Launch Schedule Management module
    On Error GoTo ErrorHandler
    frmSchedule.Show vbModal
    Exit Sub
ErrorHandler:
    MsgBox "Error opening Schedule module: " & Err.Description, vbCritical
End Sub

Private Sub cmdReportCard_Click()
    ' Launch Report Card Generator
    On Error GoTo ErrorHandler
    frmReportCard.Show vbModal
    Exit Sub
ErrorHandler:
    MsgBox "Error opening Report Card module: " & Err.Description, vbCritical
End Sub

Private Sub cmdSettings_Click()
    ' Launch Settings/Configuration
    On Error GoTo ErrorHandler
    frmSettings.Show vbModal
    Exit Sub
ErrorHandler:
    MsgBox "Error opening Settings: " & Err.Description, vbCritical
End Sub

Private Sub cmdSystemInfo_Click()
    ' Display system information
    Dim message As String
    message = CONST_SYS_NAME & vbCrLf & vbCrLf
    message = message & "Version: " & CONST_SYS_VERSION & vbCrLf
    message = message & "User: " & g_CurrentUsername & vbCrLf
    message = message & "Role: " & GetRoleDisplayName(g_CurrentUserRole) & vbCrLf
    message = message & "Build Date: " & CONST_BUILD_DATE
    
    MsgBox message, vbInformation, "System Information"
End Sub

Private Sub cmdLogout_Click()
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Safely logs out user and closes the dashboard.
    ' ------------------------------------------------------------------------------------
    If modUIFramework.Confirm("Sign out of " & CONST_SYS_NAME & "?") Then
        Call modSecurity.LogoutUser
        Call modUIFramework.SetFormState(FormState.CLOSING)
        Unload Me
        ' Control returns to modMain.StartApplication which will close the app
    End If
End Sub

' ========================================================================================
' FORM LIFECYCLE - QUERYCLOSE
' ========================================================================================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Enforces proper logout sequence before dashboard closes.
    ' ------------------------------------------------------------------------------------
    If CloseMode = 0 Then ' User clicked X button
        Cancel = True
        Call cmdLogout_Click ' Force through logout button
    End If
End Sub
