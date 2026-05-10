Option Explicit

' ========================================================================================
' Module: modUITheme
' Purpose: Enterprise-grade centralized UI theming engine.
'          Provides consistent colors, fonts, spacing, and styling across all forms.
'          Supports light/dark themes with runtime switching.
'          Single source of truth for visual design.
' Author: Senior Enterprise Architect
' Version: 1.0
' ========================================================================================

' --- COLOR PALETTE (RGB format for Excel compatibility) ---

' PRIMARY BRAND COLORS
Public Const COLOR_PRIMARY_BLUE As Long = 2070783 ' #1F7F9F - Professional Blue
Public Const COLOR_PRIMARY_ACCENT As Long = 13421619 ' #CC9933 - Gold Accent
Public Const COLOR_SECONDARY_TEAL As Long = 8367959 ' #7FA5A7 - Soft Teal

' NEUTRAL COLORS
Public Const COLOR_WHITE As Long = 16777215 ' #FFFFFF
Public Const COLOR_LIGHT_GRAY As Long = 15329769 ' #F0F0F9 - Very Light
Public Const COLOR_MEDIUM_GRAY As Long = 11842740 ' #B4B4B4 - Mid Gray
Public Const COLOR_DARK_GRAY As Long = 4210752 ' #404040 - Dark Gray
Public Const COLOR_BLACK As Long = 0 ' #000000

' STATUS COLORS
Public Const COLOR_SUCCESS_GREEN As Long = 3329330 ' #32CD32 - Bright Green
Public Const COLOR_WARNING_ORANGE As Long = 4634598 ' #FF9900 - Warning Orange
Public Const COLOR_ERROR_RED As Long = 255 ' #FF0000 - Error Red
Public Const COLOR_INFO_CYAN As Long = 16776960 ' #00FFFF - Info Cyan

' SEMANTIC COLORS
Public Const COLOR_HOVER_OVERLAY As Long = 14935011 ' #E3E3E3 - Subtle Hover
Public Const COLOR_FOCUS_RING As Long = 2070783 ' #1F7F9F - Focus Indicator
Public Const COLOR_DISABLED As Long = 12632256 ' #C0C0C0 - Disabled State
Public Const COLOR_BORDER As Long = 10921638 ' #A6A6A6 - Border Color

' DARK MODE VARIANTS (Future Implementation)
Public Const COLOR_DARK_BG As Long = 2105376 ' #202020
Public Const COLOR_DARK_SURFACE As Long = 3355443 ' #333333
Public Const COLOR_DARK_TEXT As Long = 16777215 ' #FFFFFF

' --- FONT CONFIGURATION ---

Public Const FONT_PRIMARY As String = "Segoe UI" ' Modern, professional
Public Const FONT_MONOSPACE As String = "Consolas" ' Code/technical data
Public Const FONT_SIZE_HEADING_1 As Integer = 18
Public Const FONT_SIZE_HEADING_2 As Integer = 14
Public Const FONT_SIZE_HEADING_3 As Integer = 12
Public Const FONT_SIZE_BODY As Integer = 11
Public Const FONT_SIZE_SMALL As Integer = 9
Public Const FONT_SIZE_LABEL As Integer = 10

' --- SPACING & LAYOUT (PIXELS) ---

Public Const SPACING_XS As Integer = 2
Public Const SPACING_SM As Integer = 4
Public Const SPACING_MD As Integer = 8
Public Const SPACING_LG As Integer = 12
Public Const SPACING_XL As Integer = 16
Public Const SPACING_XXL As Integer = 24

Public Const BORDER_RADIUS_SM As Integer = 2
Public Const BORDER_RADIUS_MD As Integer = 4
Public Const BORDER_RADIUS_LG As Integer = 8

Public Const BORDER_WIDTH_THIN As Integer = 1
Public Const BORDER_WIDTH_MEDIUM As Integer = 2
Public Const BORDER_WIDTH_THICK As Integer = 3

' --- CONTROL DIMENSIONS ---

Public Const BUTTON_HEIGHT As Integer = 28
Public Const BUTTON_WIDTH_SMALL As Integer = 75
Public Const BUTTON_WIDTH_MEDIUM As Integer = 120
Public Const BUTTON_WIDTH_LARGE As Integer = 150

Public Const TEXTBOX_HEIGHT As Integer = 24
Public Const COMBOBOX_HEIGHT As Integer = 24
Public Const LISTBOX_HEIGHT As Integer = 150

Public Const FORM_PADDING As Integer = 12
Public Const FORM_MIN_WIDTH As Integer = 400
Public Const FORM_MIN_HEIGHT As Integer = 300

' --- SHADOW & DEPTH ---

Public Const SHADOW_OFFSET_X As Integer = 1
Public Const SHADOW_OFFSET_Y As Integer = 2
Public Const SHADOW_BLUR_RADIUS As Integer = 4
Public Const SHADOW_OPACITY As Double = 0.15 ' 0-1 scale

' --- ANIMATION & TRANSITION ---

Public Const ANIMATION_DURATION_SHORT As Integer = 150 ' milliseconds
Public Const ANIMATION_DURATION_MEDIUM As Integer = 300
Public Const ANIMATION_DURATION_LONG As Integer = 500

' --- THEME ENUMERATION ---

Public Enum ThemeMode
    LIGHT = 1
    DARK = 2
    AUTO = 3 ' System default
End Enum

' --- GLOBAL THEME STATE ---

Public g_CurrentTheme As ThemeMode

' ========================================================================================
' COLOR UTILITY FUNCTIONS
' ========================================================================================

Public Function GetPrimaryColor() As Long
    ' Returns the primary brand color (theme-aware)
    GetPrimaryColor = COLOR_PRIMARY_BLUE
End Function

Public Function GetAccentColor() As Long
    ' Returns the accent color for highlights
    GetAccentColor = COLOR_PRIMARY_ACCENT
End Function

Public Function GetSuccessColor() As Long
    ' Returns green for success states
    GetSuccessColor = COLOR_SUCCESS_GREEN
End Function

Public Function GetErrorColor() As Long
    ' Returns red for error states
    GetErrorColor = COLOR_ERROR_RED
End Function

Public Function GetWarningColor() As Long
    ' Returns orange for warnings
    GetWarningColor = COLOR_WARNING_ORANGE
End Function

Public Function GetTextColor() As Long
    ' Returns foreground text color (theme-aware)
    Select Case g_CurrentTheme
        Case ThemeMode.DARK
            GetTextColor = COLOR_DARK_TEXT
        Case Else
            GetTextColor = COLOR_BLACK
    End Select
End Function

Public Function GetBackgroundColor() As Long
    ' Returns background color (theme-aware)
    Select Case g_CurrentTheme
        Case ThemeMode.DARK
            GetBackgroundColor = COLOR_DARK_BG
        Case Else
            GetBackgroundColor = COLOR_WHITE
    End Select
End Function

Public Function GetBorderColor() As Long
    ' Returns border color (theme-aware)
    GetBorderColor = COLOR_BORDER
End Function

Public Function GetDisabledColor() As Long
    ' Returns color for disabled controls
    GetDisabledColor = COLOR_DISABLED
End Function

' ========================================================================================
' FONT UTILITY FUNCTIONS
' ========================================================================================

Public Function GetHeadingFont() As String
    GetHeadingFont = FONT_PRIMARY
End Function

Public Function GetBodyFont() As String
    GetBodyFont = FONT_PRIMARY
End Function

Public Function GetMonospaceFont() As String
    GetMonospaceFont = FONT_MONOSPACE
End Function

' ========================================================================================
' SPACING UTILITY FUNCTIONS
' ========================================================================================

Public Function GetFormPadding() As Integer
    GetFormPadding = FORM_PADDING
End Function

Public Function GetSpacing(ByVal Size As String) As Integer
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Semantic spacing lookup by size name
    ' USAGE: left = 12 + GetSpacing("md")
    ' ------------------------------------------------------------------------------------
    Select Case LCase(Size)
        Case "xs": GetSpacing = SPACING_XS
        Case "sm": GetSpacing = SPACING_SM
        Case "md": GetSpacing = SPACING_MD
        Case "lg": GetSpacing = SPACING_LG
        Case "xl": GetSpacing = SPACING_XL
        Case "xxl": GetSpacing = SPACING_XXL
        Case Else: GetSpacing = SPACING_MD
    End Select
End Function

' ========================================================================================
' THEME MANAGEMENT
' ========================================================================================

Public Sub InitializeTheme(Optional ByVal Theme As ThemeMode = ThemeMode.LIGHT)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Sets up the application theme on startup
    ' PARAMETERS: Theme - LIGHT, DARK, or AUTO
    ' ------------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    g_CurrentTheme = Theme
    
    Debug.Print "UI Theme Initialized: " & Theme
    Exit Sub
    
ErrorHandler:
    Debug.Print "Theme Initialization Error: " & Err.Description
End Sub

Public Sub SetTheme(ByVal Theme As ThemeMode)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Changes the application theme at runtime
    ' PARAMETERS: Theme - LIGHT or DARK
    ' ------------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    g_CurrentTheme = Theme
    
    ' TODO: In Phase A-B, add code to refresh all open forms
    ' This will require iterating through VBA.UserForms collection
    
    Debug.Print "Theme Changed To: " & Theme
    Exit Sub
    
ErrorHandler:
    Debug.Print "Theme Change Error: " & Err.Description
End Sub

Public Function GetCurrentTheme() As ThemeMode
    GetCurrentTheme = g_CurrentTheme
End Function

' ========================================================================================
' COLOR CONVERSION UTILITIES
' ========================================================================================

Public Function RGBToLong(ByVal Red As Integer, ByVal Green As Integer, ByVal Blue As Integer) As Long
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Convert RGB values to Excel color format (BGR order)
    ' ------------------------------------------------------------------------------------
    RGBToLong = (Blue * 65536) + (Green * 256) + Red
End Function

Public Function HexToLong(ByVal HexColor As String) As Long
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Convert hex color string to Excel Long format
    ' USAGE: color = HexToLong("#1F7F9F")
    ' ------------------------------------------------------------------------------------
    Dim cleanHex As String
    cleanHex = Replace(HexColor, "#", "")
    HexToLong = CLng("&H" & cleanHex)
End Function

Public Function LongToHex(ByVal ColorLong As Long) As String
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Convert Excel Long color to hex string
    ' ------------------------------------------------------------------------------------
    LongToHex = "#" & Format(Hex(ColorLong), "000000")
End Function

' ========================================================================================
' BUTTON STYLING HELPERS
' ========================================================================================

Public Sub StyleButtonPrimary(ByRef btn As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Applies primary button styling (blue, prominent)
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    With btn
        .BackColor = GetPrimaryColor()
        .ForeColor = COLOR_WHITE
        .Font.Name = GetBodyFont()
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = False
        .Height = BUTTON_HEIGHT
    End With
End Sub

Public Sub StyleButtonSecondary(ByRef btn As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Applies secondary button styling (gray, less prominent)
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    With btn
        .BackColor = COLOR_LIGHT_GRAY
        .ForeColor = GetTextColor()
        .Font.Name = GetBodyFont()
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = False
        .Height = BUTTON_HEIGHT
        .BorderColor = GetBorderColor()
    End With
End Sub

Public Sub StyleButtonDanger(ByRef btn As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Applies danger button styling (red, for destructive actions)
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    With btn
        .BackColor = GetErrorColor()
        .ForeColor = COLOR_WHITE
        .Font.Name = GetBodyFont()
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = True
        .Height = BUTTON_HEIGHT
    End With
End Sub

Public Sub StyleButtonSuccess(ByRef btn As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Applies success button styling (green, for positive actions)
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    With btn
        .BackColor = GetSuccessColor()
        .ForeColor = COLOR_WHITE
        .Font.Name = GetBodyFont()
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = False
        .Height = BUTTON_HEIGHT
    End With
End Sub

' ========================================================================================
' LABEL STYLING HELPERS
' ========================================================================================

Public Sub StyleLabelHeading(ByRef lbl As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Styles label as main heading
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    With lbl
        .Font.Name = GetHeadingFont()
        .Font.Size = FONT_SIZE_HEADING_1
        .Font.Bold = True
        .ForeColor = GetTextColor()
        .BackColor = GetBackgroundColor()
    End With
End Sub

Public Sub StyleLabelSubheading(ByRef lbl As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Styles label as subheading
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    With lbl
        .Font.Name = GetHeadingFont()
        .Font.Size = FONT_SIZE_HEADING_2
        .Font.Bold = True
        .ForeColor = GetPrimaryColor()
        .BackColor = GetBackgroundColor()
    End With
End Sub

Public Sub StyleLabelBody(ByRef lbl As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Styles label as regular body text
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    With lbl
        .Font.Name = GetBodyFont()
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = False
        .ForeColor = GetTextColor()
        .BackColor = GetBackgroundColor()
    End With
End Sub

Public Sub StyleLabelSmall(ByRef lbl As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Styles label as small text (helper text, captions)
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    With lbl
        .Font.Name = GetBodyFont()
        .Font.Size = FONT_SIZE_SMALL
        .Font.Bold = False
        .ForeColor = COLOR_MEDIUM_GRAY
        .BackColor = GetBackgroundColor()
    End With
End Sub

Public Sub StyleLabelError(ByRef lbl As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Styles label for error messages (red text)
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    With lbl
        .Font.Name = GetBodyFont()
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = False
        .ForeColor = GetErrorColor()
        .BackColor = GetBackgroundColor()
    End With
End Sub

Public Sub StyleLabelSuccess(ByRef lbl As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Styles label for success messages (green text)
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    With lbl
        .Font.Name = GetBodyFont()
        .Font.Size = FONT_SIZE_BODY
        .Font.Bold = False
        .ForeColor = GetSuccessColor()
        .BackColor = GetBackgroundColor()
    End With
End Sub

' ========================================================================================
' TEXTBOX STYLING HELPERS
' ========================================================================================

Public Sub StyleTextboxStandard(ByRef txt As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Applies standard textbox styling
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    With txt
        .Font.Name = GetBodyFont()
        .Font.Size = FONT_SIZE_BODY
        .BackColor = COLOR_WHITE
        .ForeColor = GetTextColor()
        .BorderColor = GetBorderColor()
        .BorderStyle = 1 ' Single line border
        .Height = TEXTBOX_HEIGHT
    End With
End Sub

Public Sub StyleTextboxDisabled(ByRef txt As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Styles textbox in disabled state
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    With txt
        .Font.Name = GetBodyFont()
        .Font.Size = FONT_SIZE_BODY
        .BackColor = COLOR_LIGHT_GRAY
        .ForeColor = COLOR_MEDIUM_GRAY
        .BorderColor = GetDisabledColor()
        .Enabled = False
    End With
End Sub

' ========================================================================================
' FORM STYLING HELPERS
' ========================================================================================

Public Sub StyleFormDefault(ByRef frm As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Applies standard form styling
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    With frm
        .BackColor = GetBackgroundColor()
        .Font.Name = GetBodyFont()
        .Font.Size = FONT_SIZE_BODY
    End With
End Sub
