Option Explicit

' ========================================================================================
' UserForm: frmLogin (MODERNIZED - ENTERPRISE EDITION)
' Purpose: Professional authentication gateway with modern UI/UX patterns.
'          Entry point to the Sunday School Management System.
'          Implements secure login with visual feedback, validation, and error handling.
' Author: Senior Enterprise Architect
' Version: 2.0 (Modernized)
' ========================================================================================

' --- FORM-LEVEL VARIABLES ---
Private m_LoginAttempts As Integer
Private m_FormInitialized As Boolean
Private m_IsAuthenticating As Boolean

' --- UI STATE COLORS (Cache for performance) ---
Private m_ColorPrimary As Long
Private m_ColorError As Long
Private m_ColorSuccess As Long
Private m_ColorText As Long
Private m_ColorBackground As Long

' ========================================================================================
' FORM INITIALIZATION & LIFECYCLE
' ========================================================================================

Private Sub UserForm_Initialize()
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Initializes the login form with modern styling and defaults.
    '          Called automatically when form loads into memory.
    ' ARCHITECTURE: Deferred initialization (controls set up after Show for performance)
    ' ------------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    ' 1. Initialize framework
    Call modUIFramework.InitializeBaseForm(Me, "frmLogin")
    
    ' 2. Cache theme colors for performance (avoid repeated function calls)
    m_ColorPrimary = modUITheme.GetPrimaryColor()
    m_ColorError = modUITheme.GetErrorColor()
    m_ColorSuccess = modUITheme.GetSuccessColor()
    m_ColorText = modUITheme.GetTextColor()
    m_ColorBackground = modUITheme.GetBackgroundColor()
    
    ' 3. Configure form appearance
    With Me
        .BackColor = m_ColorBackground
        .Width = 450 ' Narrower form for focus (not 800px desktop)
        .Height = 550 ' Taller for breathing room
        .Left = (Application.Width - .Width) / 2 ' Center horizontally
        .Top = (Application.Height - .Height) / 2 ' Center vertically
        .Caption = CONST_SYS_NAME & " - Login"
        .StartUpPosition = 1 ' CenterOwner (once Shown)
    End With
    
    ' 4. Initialize controls
    Call InitializeControls
    
    ' 5. Set defaults and focus
    m_LoginAttempts = 0
    m_IsAuthenticating = False
    m_FormInitialized = True
    
    ' 6. Set initial focus
    Call modUIFramework.SetFocusOnControl(Me, "txtUsername")
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Login Form Initialization Error: " & Err.Description, vbCritical, "System Error"
    Unload Me
End Sub

Private Sub InitializeControls()
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Configures all form controls with professional styling.
    '          Separated from Initialize for clarity and testability.
    ' ------------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    ' --- HEADER SECTION ---
    With Me.lblTitle
        .Caption = "Welcome"
        .Top = 12
        .Left = 12
        .Width = 400
        .Height = 28
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_HEADING_1
        .Font.Bold = True
        .ForeColor = m_ColorPrimary
        .BackColor = m_ColorBackground
    End With
    
    With Me.lblSubtitle
        .Caption = "Sign in to " & CONST_SYS_NAME
        .Top = 42
        .Left = 12
        .Width = 400
        .Height = 16
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_BODY - 1
        .Font.Bold = False
        .ForeColor = modUITheme.COLOR_MEDIUM_GRAY
        .BackColor = m_ColorBackground
    End With
    
    ' --- DIVIDER (VISUAL SEPARATOR) ---
    With Me.lblDivider1
        .BackColor = modUITheme.COLOR_BORDER
        .Top = 62
        .Left = 12
        .Width = 400
        .Height = 1
    End With
    
    ' --- USERNAME LABEL & INPUT ---
    With Me.lblUsername
        .Caption = "Username"
        .Top = 80
        .Left = 12
        .Width = 100
        .Height = 16
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_LABEL
        .Font.Bold = True
        .ForeColor = m_ColorText
        .BackColor = m_ColorBackground
    End With
    
    With Me.txtUsername
        .Top = 98
        .Left = 12
        .Width = 400
        .Height = modUITheme.TEXTBOX_HEIGHT
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_BODY
        .BackColor = modUITheme.COLOR_WHITE
        .ForeColor = m_ColorText
        .BorderColor = modUITheme.COLOR_BORDER
        .BorderStyle = 1 ' Single line
        .Text = ""
        .PasswordChar = ""
    End With
    
    ' --- PASSWORD LABEL & INPUT ---
    With Me.lblPassword
        .Caption = "Password"
        .Top = 135
        .Left = 12
        .Width = 100
        .Height = 16
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_LABEL
        .Font.Bold = True
        .ForeColor = m_ColorText
        .BackColor = m_ColorBackground
    End With
    
    With Me.txtPassword
        .Top = 153
        .Left = 12
        .Width = 400
        .Height = modUITheme.TEXTBOX_HEIGHT
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_BODY
        .BackColor = modUITheme.COLOR_WHITE
        .ForeColor = m_ColorText
        .BorderColor = modUITheme.COLOR_BORDER
        .BorderStyle = 1 ' Single line
        .Text = ""
        .PasswordChar = "*" ' Mask password input
    End With
    
    ' --- "REMEMBER ME" CHECKBOX (Optional Feature) ---
    With Me.chkRememberMe
        .Caption = "Remember me on this computer"
        .Top = 190
        .Left = 12
        .Width = 200
        .Height = 16
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_SMALL
        .Value = False
        .BackColor = m_ColorBackground
        .ForeColor = m_ColorText
    End With
    
    ' --- ERROR MESSAGE LABEL (Hidden by default) ---
    With Me.lblError
        .Caption = ""
        .Top = 215
        .Left = 12
        .Width = 400
        .Height = 32
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_SMALL
        .Font.Bold = False
        .ForeColor = m_ColorError
        .BackColor = modUITheme.HexToLong("#FFEBEE") ' Light red background
        .TextAlign = fmTextAlignLeft
        .WordWrap = True
        .BorderColor = m_ColorError
        .BorderStyle = 1
        .Visible = False ' Hidden until needed
    End With
    
    ' --- DIVIDER 2 ---
    With Me.lblDivider2
        .BackColor = modUITheme.COLOR_BORDER
        .Top = 260
        .Left = 12
        .Width = 400
        .Height = 1
    End With
    
    ' --- LOGIN BUTTON (Primary CTA) ---
    With Me.cmdLogin
        .Caption = "Sign In"
        .Top = 280
        .Left = 12
        .Width = 400
        .Height = modUITheme.BUTTON_HEIGHT
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_BODY
        .Font.Bold = True
        Call modUITheme.StyleButtonPrimary(Me.cmdLogin)
    End With
    
    ' --- CANCEL BUTTON (Secondary) ---
    With Me.cmdCancel
        .Caption = "Cancel"
        .Top = 320
        .Left = 12
        .Width = 400
        .Height = modUITheme.BUTTON_HEIGHT
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_BODY
        .Font.Bold = False
        Call modUITheme.StyleButtonSecondary(Me.cmdCancel)
    End With
    
    ' --- FOOTER SECTION (OPTIONAL FUTURE LINKS) ---
    With Me.lblFooter
        .Caption = "Forgot password? Contact your administrator."
        .Top = 365
        .Left = 12
        .Width = 400
        .Height = 16
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = modUITheme.FONT_SIZE_SMALL
        .Font.Bold = False
        .Font.Italic = True
        .ForeColor = modUITheme.COLOR_MEDIUM_GRAY
        .BackColor = m_ColorBackground
    End With
    
    ' --- VERSION INFO (Bottom right) ---
    With Me.lblVersion
        .Caption = "v" & CONST_SYS_VERSION
        .Top = 515
        .Left = 12
        .Width = 100
        .Height = 12
        .Font.Name = modUITheme.FONT_PRIMARY
        .Font.Size = 8
        .Font.Bold = False
        .ForeColor = modUITheme.COLOR_MEDIUM_GRAY
        .BackColor = m_ColorBackground
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error Initializing Controls: " & Err.Description, vbCritical
End Sub

' ========================================================================================
' AUTHENTICATION LOGIC
' ========================================================================================

Private Sub cmdLogin_Click()
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Triggered when user clicks Login button or presses Enter in password field.
    '          Validates credentials and initiates authentication sequence.
    ' SECURITY: Implements rate limiting and lockout after failed attempts.
    ' ------------------------------------------------------------------------------------
    Dim strUsername As String
    Dim strPassword As String
    Dim isAuthenticated As Boolean
    
    On Error GoTo ErrorHandler
    
    ' 1. RATE LIMITING CHECK (Security: Prevent brute force)
    If m_LoginAttempts >= modConstants.MAX_LOGIN_ATTEMPTS Then
        Call DisplayError "Too many login attempts. System locked for " & _
                         modConstants.LOCKOUT_DURATION_MINUTES & " minutes."
        Call modUIFramework.DisableAllControls(Me)
        Call SetupLockoutTimer
        Exit Sub
    End If
    
    ' 2. FORM VALIDATION (Client-side before DB query)
    strUsername = Trim(Me.txtUsername.Text)
    strPassword = Trim(Me.txtPassword.Text)
    
    If Not ValidateLoginForm(strUsername, strPassword) Then
        Exit Sub
    End If
    
    ' 3. SET PROCESSING STATE
    m_IsAuthenticating = True
    Call modUIFramework.SetProcessingState(Me, True, "Authenticating...")
    Call HideErrorMessage ' Clear previous errors
    DoEvents ' Allow UI update
    
    ' 4. AUTHENTICATE USER (Pass to security engine)
    isAuthenticated = modSecurity.AuthenticateUser(strUsername, strPassword)
    
    ' 5. HANDLE AUTHENTICATION RESULT
    If isAuthenticated Then
        ' SUCCESS PATH
        Call DisplaySuccess "Welcome, " & g_CurrentUsername & "!"
        Call modUIFramework.SetFormState(FormState.CLOSING)
        DoEvents
        Application.Wait (Now + TimeValue("0:00:00.5")) ' Brief success display
        ' Form will be unloaded by modMain.StartApplication after this exits
        Unload Me
    Else
        ' FAILURE PATH
        m_LoginAttempts = m_LoginAttempts + 1
        Call DisplayError "Invalid username or password. " & _
                         (modConstants.MAX_LOGIN_ATTEMPTS - m_LoginAttempts) & " attempts remaining."
        Call ResetPasswordField
    End If
    
    m_IsAuthenticating = False
    Call modUIFramework.SetProcessingState(Me, False)
    
    Exit Sub
    
ErrorHandler:
    m_IsAuthenticating = False
    Call modUIFramework.SetProcessingState(Me, False)
    Call DisplayError "Authentication Error: " & Err.Description
    Debug.Print "frmLogin.cmdLogin_Click Error: " & Err.Description
End Sub

Private Function ValidateLoginForm(ByVal username As String, ByVal password As String) As Boolean
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Client-side validation of login inputs.
    ' RETURNS: True if valid, False if validation failed
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    
    ' Check for empty fields
    If username = "" Then
        Call DisplayError "Please enter your username."
        Call modUIFramework.SetFocusOnControl(Me, "txtUsername")
        ValidateLoginForm = False
        Exit Function
    End If
    
    If password = "" Then
        Call DisplayError "Please enter your password."
        Call modUIFramework.SetFocusOnControl(Me, "txtPassword")
        ValidateLoginForm = False
        Exit Function
    End If
    
    ' Check minimum length
    If Len(username) < 3 Then
        Call DisplayError "Username must be at least 3 characters."
        ValidateLoginForm = False
        Exit Function
    End If
    
    If Len(password) < 1 Then ' Already checked empty, but be explicit
        Call DisplayError "Password cannot be empty."
        ValidateLoginForm = False
        Exit Function
    End If
    
    ValidateLoginForm = True
End Function

Private Sub SetupLockoutTimer()
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Sets up a timer for account lockout after too many failed attempts.
    ' NOTE: In Phase B, this will use proper Excel-based timer mechanism
    ' ------------------------------------------------------------------------------------
    ' TODO: Implement actual timeout
    ' For now, form is just disabled
    Debug.Print "Login attempt limit reached. User locked out temporarily."
End Sub

' ========================================================================================
' UI FEEDBACK MECHANISMS
' ========================================================================================

Private Sub DisplayError(ByVal message As String)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Shows error message in the designated label with styling.
    ' PARAMETERS: message - The error text to display
    ' ------------------------------------------------------------------------------------
    With Me.lblError
        .Caption = message
        .ForeColor = m_ColorError
        .BackColor = modUITheme.HexToLong("#FFEBEE") ' Light red
        .Visible = True
    End With
End Sub

Private Sub DisplaySuccess(ByVal message As String)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Shows success message in the designated label with styling.
    ' PARAMETERS: message - The success text to display
    ' ------------------------------------------------------------------------------------
    With Me.lblError ' Reuse the label, just change styling
        .Caption = message
        .ForeColor = m_ColorSuccess
        .BackColor = modUITheme.HexToLong("#E8F5E9") ' Light green
        .Visible = True
    End With
End Sub

Private Sub HideErrorMessage()
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Hides the error/message label.
    ' ------------------------------------------------------------------------------------
    Me.lblError.Visible = False
    Me.lblError.Caption = ""
End Sub

Private Sub ResetPasswordField()
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Clears password field and returns focus (security best practice).
    ' ------------------------------------------------------------------------------------
    Me.txtPassword.Text = ""
    Call modUIFramework.SetFocusOnControl(Me, "txtPassword")
End Sub

' ========================================================================================
' BUTTON EVENT HANDLERS
' ========================================================================================

Private Sub cmdCancel_Click()
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Safely closes the login form and exits the application.
    ' ------------------------------------------------------------------------------------
    If modUIFramework.Confirm("Exit " & CONST_SYS_NAME & "?") Then
        Call modUIFramework.SetFormState(FormState.CLOSING)
        ' Reset global login state
        g_IsLoggedIn = False
        Unload Me
    End If
End Sub

' ========================================================================================
' KEYBOARD EVENT HANDLERS (Usability)
' ========================================================================================

Private Sub txtUsername_KeyDown(ByVal KeyCode As MSForms.ReturnKeyCode, ByVal Shift As Integer)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Allows Tab key to move to password field.
    ' ------------------------------------------------------------------------------------
    If KeyCode = vbKeyTab Then
        KeyCode = 0 ' Suppress default behavior
        Call modUIFramework.SetFocusOnControl(Me, "txtPassword")
    End If
End Sub

Private Sub txtPassword_KeyDown(ByVal KeyCode As MSForms.ReturnKeyCode, ByVal Shift As Integer)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Allows Enter key to submit login.
    '          Allows Shift+Tab to go back to username field.
    ' ------------------------------------------------------------------------------------
    If KeyCode = vbKeyReturn Then
        KeyCode = 0 ' Suppress default
        Call cmdLogin_Click ' Trigger login
    ElseIf KeyCode = vbKeyTab And Shift = 1 Then ' Shift+Tab
        KeyCode = 0
        Call modUIFramework.SetFocusOnControl(Me, "txtUsername")
    End If
End Sub

' ========================================================================================
' FORM LIFECYCLE - QUERYCLOSE
' ========================================================================================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Prevents bypassing login by clicking red X button.
    '          Enforces using Cancel button for proper cleanup.
    ' PARAMETERS:
    '   Cancel - Set to True to prevent closing
    '   CloseMode - 0 = X button, 1 = Unload, 2 = CloseButton
    ' ------------------------------------------------------------------------------------
    If CloseMode = 0 Then ' User clicked X button
        Cancel = True
        ' Silently close instead of showing message (better UX)
        If modUIFramework.Confirm("Exit " & CONST_SYS_NAME & "?") Then
            Cancel = False
        End If
    End If
End Sub

' ========================================================================================
' INITIALIZATION BOOTSTRAP (Called by modMain)
' ========================================================================================

Public Sub BootstrapLogin()
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Ensures default admin account exists before login form is shown.
    '          Called from modMain.StartApplication.
    ' ------------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    ' Create admin account if it doesn't exist
    Call modSecurity.CreateDefaultAdmin
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Bootstrap Error: " & Err.Description, vbCritical
End Sub
