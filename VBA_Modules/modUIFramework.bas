Option Explicit

' ========================================================================================
' Module: modUIFramework
' Purpose: Core UI framework for enterprise-grade UserForm management.
'          Handles form initialization, lifecycle, validation, and event management.
'          Provides reusable UI patterns and standards across all forms.
' Author: Senior Enterprise Architect
' Version: 1.0
' ========================================================================================

' --- FORM LIFECYCLE STATES ---
Public Enum FormState
    INITIALIZING = 1
    READY = 2
    LOADING = 3
    PROCESSING = 4
    ERROR = 5
    CLOSING = 6
End Enum

' --- FORM RESULT CODES ---
Public Enum FormResult
    OK = 1
    CANCEL = 2
    RETRY = 3
    ERROR_VALIDATION = 4
    ERROR_DATABASE = 5
End Enum

' --- GLOBAL FORM STATE TRACKER ---
Public g_FormState As FormState
Public g_LastFormResult As FormResult

' ========================================================================================
' FORM INITIALIZATION & LIFECYCLE
' ========================================================================================

Public Sub InitializeBaseForm(ByRef frm As Object, ByVal FormName As String)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Standard initialization for all enterprise UserForms.
    '          Call this from every form's UserForm_Initialize event.
    ' PARAMETERS:
    '   frm - The UserForm object (use 'Me' from within form code)
    '   FormName - Display name of the form for logging
    ' EXAMPLE: Call modUIFramework.InitializeBaseForm(Me, "frmStudents")
    ' ------------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    g_FormState = FormState.INITIALIZING
    
    ' 1. Set default properties
    With frm
        Call modUITheme.StyleFormDefault(frm)
        .AutoRedraw = False ' Reduce flicker
        .ShowModal = False ' Allow background interaction (we'll control with vbModal)
    End With
    
    ' 2. Theme initialization
    Call modUITheme.InitializeTheme(ThemeMode.LIGHT)
    
    ' 3. Disable visual updates during setup
    Application.ScreenUpdating = False
    
    ' 4. Log form initialization
    Debug.Print "[" & FormName & "] Initializing..."
    
    g_FormState = FormState.READY
    
    ' 5. Re-enable updates
    Application.ScreenUpdating = True
    
    Exit Sub
    
ErrorHandler:
    g_FormState = FormState.ERROR
    MsgBox "Form Initialization Error: " & Err.Description, vbCritical, "System Error"
End Sub

Public Sub SetFormState(ByVal NewState As FormState)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Changes the current form state and logs transition.
    ' ------------------------------------------------------------------------------------
    Dim previousState As FormState
    previousState = g_FormState
    g_FormState = NewState
    
    Debug.Print "Form State: " & previousState & " -> " & NewState
End Sub

Public Function GetFormState() As FormState
    GetFormState = g_FormState
End Function

Public Sub SetLoadingState(ByRef frm As Object, ByVal IsLoading As Boolean, Optional ByVal LoadingMessage As String = "")
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Shows/hides loading indicator on the form.
    ' PARAMETERS:
    '   frm - The form object
    '   IsLoading - True to show, False to hide
    '   LoadingMessage - Optional message (e.g., "Loading students...")
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    
    If IsLoading Then
        SetFormState FormState.LOADING
        ' Look for a lblLoading label on the form
        If Not frm.lblLoading Is Nothing Then
            frm.lblLoading.Visible = True
            If LoadingMessage <> "" Then
                frm.lblLoading.Caption = LoadingMessage
            End If
        End If
        ' Disable main controls to prevent user interaction
        Application.StatusBar = "Loading data. Please wait..."
    Else
        SetFormState FormState.READY
        ' Hide loading indicator
        If Not frm.lblLoading Is Nothing Then
            frm.lblLoading.Visible = False
        End If
        ' Restore status bar
        Application.StatusBar = False
    End If
End Sub

Public Sub SetProcessingState(ByRef frm As Object, ByVal IsProcessing As Boolean, Optional ByVal ProcessingMessage As String = "")
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Shows processing state (for operations, saving, etc.)
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    
    If IsProcessing Then
        SetFormState FormState.PROCESSING
        ' Disable form buttons to prevent double-submit
        Dim ctrl As Object
        For Each ctrl In frm.Controls
            If InStr(ctrl.Name, "cmd") > 0 Then ' Disable all buttons
                ctrl.Enabled = False
            End If
        Next ctrl
        If ProcessingMessage <> "" Then
            Application.StatusBar = ProcessingMessage
        Else
            Application.StatusBar = "Processing. Please wait..."
        End If
    Else
        SetFormState FormState.READY
        ' Re-enable buttons
        For Each ctrl In frm.Controls
            If InStr(ctrl.Name, "cmd") > 0 Then
                ctrl.Enabled = True
            End If
        Next ctrl
        Application.StatusBar = False
    End If
End Sub

' ========================================================================================
' INPUT VALIDATION FRAMEWORK
' ========================================================================================

Public Function ValidateRequired(ByVal value As String, ByVal FieldName As String) As Boolean
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Validates that a field is not empty.
    ' RETURNS: True if valid, False if empty
    ' ------------------------------------------------------------------------------------
    If Trim(value) = "" Then
        ShowFieldError FieldName & " is required."
        ValidateRequired = False
    Else
        ValidateRequired = True
    End If
End Function

Public Function ValidateEmail(ByVal email As String) As Boolean
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Validates email format using basic regex pattern.
    ' RETURNS: True if valid format, False if invalid
    ' ------------------------------------------------------------------------------------
    Dim pattern As String
    ' Simple email pattern: anything@anything.anything
    If InStr(email, "@") > 0 And InStr(email, ".") > InStr(email, "@") Then
        ValidateEmail = True
    Else
        ShowFieldError "Please enter a valid email address."
        ValidateEmail = False
    End If
End Function

Public Function ValidatePhone(ByVal phone As String) As Boolean
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Validates phone format (basic numeric check).
    ' RETURNS: True if valid, False if invalid
    ' ------------------------------------------------------------------------------------
    Dim i As Integer, digitCount As Integer
    For i = 1 To Len(phone)
        If IsNumeric(Mid(phone, i, 1)) Then digitCount = digitCount + 1
    Next i
    
    If digitCount >= 7 Then ' At least 7 digits
        ValidatePhone = True
    Else
        ShowFieldError "Please enter a valid phone number (at least 7 digits)."
        ValidatePhone = False
    End If
End Function

Public Function ValidateDate(ByVal dateStr As String, ByVal FieldName As String) As Boolean
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Validates that a string is a valid date.
    ' RETURNS: True if valid date, False if invalid
    ' ------------------------------------------------------------------------------------
    If IsDate(dateStr) Then
        ValidateDate = True
    Else
        ShowFieldError FieldName & " must be a valid date."
        ValidateDate = False
    End If
End Function

Public Function ValidateNumeric(ByVal value As String, ByVal FieldName As String) As Boolean
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Validates that a string is numeric.
    ' RETURNS: True if numeric, False if not
    ' ------------------------------------------------------------------------------------
    If IsNumeric(value) Then
        ValidateNumeric = True
    Else
        ShowFieldError FieldName & " must be a number."
        ValidateNumeric = False
    End If
End Function

Public Function ValidateRange(ByVal value As Double, ByVal MinVal As Double, ByVal MaxVal As Double, _
                              ByVal FieldName As String) As Boolean
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Validates that a numeric value is within a range.
    ' RETURNS: True if in range, False if outside
    ' ------------------------------------------------------------------------------------
    If value >= MinVal And value <= MaxVal Then
        ValidateRange = True
    Else
        ShowFieldError FieldName & " must be between " & MinVal & " and " & MaxVal & "."
        ValidateRange = False
    End If
End Function

Public Function ValidateSelection(ByVal selectedIndex As Integer, ByVal FieldName As String) As Boolean
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Validates that a ComboBox has a selection.
    ' RETURNS: True if selected, False if no selection
    ' USAGE: if ValidateSelection(cmbClass.ListIndex, "Class") then...
    ' ------------------------------------------------------------------------------------
    If selectedIndex >= 0 Then ' -1 means no selection
        ValidateSelection = True
    Else
        ShowFieldError "Please select a " & FieldName & "."
        ValidateSelection = False
    End If
End Function

Public Function ValidateLength(ByVal value As String, ByVal MinLen As Integer, ByVal MaxLen As Integer, _
                               ByVal FieldName As String) As Boolean
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Validates string length constraints.
    ' RETURNS: True if within bounds, False if outside
    ' ------------------------------------------------------------------------------------
    Dim actualLen As Integer
    actualLen = Len(Trim(value))
    
    If actualLen >= MinLen And actualLen <= MaxLen Then
        ValidateLength = True
    Else
        ShowFieldError FieldName & " must be between " & MinLen & " and " & MaxLen & " characters."
        ValidateLength = False
    End If
End Function

' ========================================================================================
' ERROR & FEEDBACK UI
' ========================================================================================

Public Sub ShowFieldError(ByVal ErrorMessage As String)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Displays a field-level error message.
    ' PARAMETERS: ErrorMessage - The error text to display
    ' ------------------------------------------------------------------------------------
    ' In Phase A-B, we'll enhance this to support in-form error labels
    ' For now, using MsgBox as fallback
    MsgBox ErrorMessage, vbExclamation, "Validation Error"
End Sub

Public Sub ShowNotification(ByVal Message As String, ByVal Type As String)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Displays a toast-style notification (success, info, warning, error).
    ' PARAMETERS:
    '   Message - The notification text
    '   Type - "success", "info", "warning", "error"
    ' NOTE: In Phase A-B, this will be enhanced for true toast notifications
    ' ------------------------------------------------------------------------------------
    Dim Icon As VbMsgBoxStyle
    Dim Title As String
    
    Select Case LCase(Type)
        Case "success"
            Icon = vbInformation
            Title = "Success"
        Case "warning"
            Icon = vbExclamation
            Title = "Warning"
        Case "error"
            Icon = vbCritical
            Title = "Error"
        Case Else ' info
            Icon = vbInformation
            Title = "Information"
    End Select
    
    MsgBox Message, Icon, Title
End Sub

Public Sub ShowSuccess(ByVal Message As String)
    ShowNotification Message, "success"
End Sub

Public Sub ShowError(ByVal Message As String)
    ShowNotification Message, "error"
End Sub

Public Sub ShowWarning(ByVal Message As String)
    ShowNotification Message, "warning"
End Sub

Public Sub ShowInfo(ByVal Message As String)
    ShowNotification Message, "info"
End Sub

Public Function Confirm(ByVal Message As String, Optional ByVal Title As String = "Confirm") As Boolean
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Shows a confirmation dialog.
    ' RETURNS: True if user clicked Yes, False if No
    ' ------------------------------------------------------------------------------------
    Confirm = (MsgBox(Message, vbYesNo + vbQuestion, Title) = vbYes)
End Function

' ========================================================================================
' CONTROL MANAGEMENT
' ========================================================================================

Public Sub ClearAllTextboxes(ByRef frm As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Clears all textbox controls on a form.
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    Dim ctrl As Object
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is MSForms.TextBox Then
            ctrl.Text = ""
        End If
    Next ctrl
End Sub

Public Sub ClearAllComboboxes(ByRef frm As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Resets all combobox controls on a form.
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    Dim ctrl As Object
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is MSForms.ComboBox Then
            ctrl.ListIndex = -1
        End If
    Next ctrl
End Sub

Public Sub DisableAllControls(ByRef frm As Object, Optional ByVal ExceptControl As String = "")
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Disables all controls on a form except specified ones.
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    Dim ctrl As Object
    For Each ctrl In frm.Controls
        If ctrl.Name <> ExceptControl Then
            ctrl.Enabled = False
        End If
    Next ctrl
End Sub

Public Sub EnableAllControls(ByRef frm As Object)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Enables all controls on a form.
    ' ------------------------------------------------------------------------------------
    On Error Resume Next
    Dim ctrl As Object
    For Each ctrl In frm.Controls
        ctrl.Enabled = True
    Next ctrl
End Sub

Public Sub SetFocusOnControl(ByRef frm As Object, ByVal ControlName As String)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Sets focus to a specific control by name.
    ' ------------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    frm.Controls(ControlName).SetFocus
    Exit Sub
ErrorHandler:
    Debug.Print "SetFocusOnControl: Control " & ControlName & " not found."
End Sub

' ========================================================================================
' DATA BINDING HELPERS
' ========================================================================================

Public Sub PopulateComboFromArray(ByRef comb As Object, ByVal dataArray As Variant)
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Populates a ComboBox from an array.
    ' PARAMETERS:
    '   comb - ComboBox control
    '   dataArray - Array of values
    ' EXAMPLE: Call PopulateComboFromArray(cmbClass, GetAllClasses())
    ' ------------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    comb.Clear
    
    For i = LBound(dataArray) To UBound(dataArray)
        comb.AddItem dataArray(i)
    Next i
    
    Exit Sub
ErrorHandler:
    Debug.Print "PopulateComboFromArray Error: " & Err.Description
End Sub

Public Sub PopulateListFromTable(ByRef lst As Object, ByVal TableName As String, _
                                  ByVal DisplayColumn As String, Optional ByVal FilterColumn As String = "", _
                                  Optional ByVal FilterValue As Variant = "")
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Populates a ListBox from a database table.
    ' PARAMETERS:
    '   lst - ListBox control
    '   TableName - Name of the database table (e.g., "tblStudents")
    '   DisplayColumn - Column to display (e.g., "FirstName")
    '   FilterColumn - Optional column to filter by
    '   FilterValue - Value to filter for
    ' EXAMPLE: Call PopulateListFromTable(lstStudents, "tblStudents", "FirstName", "Class", "Grade 1")
    ' ------------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    
    Dim lo As ListObject
    Dim lr As ListRow
    Dim displayValue As String
    
    Set lo = modDatabase.GetTable(TableName)
    If lo Is Nothing Then Err.Raise vbObjectError + 1, , "Table not found: " & TableName
    
    lst.Clear
    
    For Each lr In lo.ListRows
        ' Apply filter if specified
        If FilterColumn <> "" Then
            If Intersect(lr.Range, lo.ListColumns(FilterColumn).Range).Value <> FilterValue Then
                GoTo NextRow
            End If
        End If
        
        displayValue = Intersect(lr.Range, lo.ListColumns(DisplayColumn).Range).Value
        lst.AddItem displayValue
        
NextRow:
    Next lr
    
    Exit Sub
ErrorHandler:
    Debug.Print "PopulateListFromTable Error: " & Err.Description
End Sub

' ========================================================================================
' HELPER FUNCTIONS
' ========================================================================================

Public Function IsNumeric(ByVal value As String) As Boolean
    ' ------------------------------------------------------------------------------------
    ' PURPOSE: Checks if a string represents a numeric value.
    ' ------------------------------------------------------------------------------------
    On Error GoTo ErrorHandler
    Dim dummy As Double
    dummy = CDbl(Trim(value))
    IsNumeric = True
    Exit Function
ErrorHandler:
    IsNumeric = False
End Function
