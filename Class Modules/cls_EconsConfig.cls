VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_EconsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Single source of truth for the econs PowerPoint export configuration.
'
' Holds the number of outputs, each output's display name + named range, and
' the list of cases to print. Also owns load/save of the last-used preset via
' the per-user registry (SaveSetting/GetSetting), so the dialog can re-open
' pre-populated with the previous successful run.

Private Const MAX_OUTPUTS As Long = 5
Private Const REG_APP As String = "Vantage"
Private Const REG_SECTION As String = "Econs"

Private mNumOutputs As Long
Private mNames(1 To MAX_OUTPUTS) As String
Private mRanges(1 To MAX_OUTPUTS) As String
Private mCaseRange As String
Private mCases As String
Private mReplaceExisting As Boolean
Private mTitleCaseFirst As Boolean

' Advanced: up to two extra per-case outputs (the "3rd / 4th" outputs). Each has
' a name, a named range, and the set of cases it applies to.
Private Const MAX_ADVANCED As Long = 2
Private mAdvName(1 To MAX_ADVANCED) As String
Private mAdvRange(1 To MAX_ADVANCED) As String
Private mAdvCases(1 To MAX_ADVANCED) As String

Private Sub Class_Initialize()
    Call LoadDefaults
End Sub

Public Property Get MaxOutputs() As Long
    MaxOutputs = MAX_OUTPUTS
End Property

Public Property Get NumOutputs() As Long
    NumOutputs = mNumOutputs
End Property

Public Property Let NumOutputs(ByVal value As Long)
    If value < 1 Then value = 1
    If value > MAX_OUTPUTS Then value = MAX_OUTPUTS
    mNumOutputs = value
End Property

Public Property Get OutputName(ByVal idx As Long) As String
    OutputName = mNames(idx)
End Property

Public Property Let OutputName(ByVal idx As Long, ByVal value As String)
    mNames(idx) = value
End Property

Public Property Get OutputRange(ByVal idx As Long) As String
    OutputRange = mRanges(idx)
End Property

Public Property Let OutputRange(ByVal idx As Long, ByVal value As String)
    mRanges(idx) = value
End Property

Public Property Get CaseRangeName() As String
    CaseRangeName = mCaseRange
End Property

Public Property Let CaseRangeName(ByVal value As String)
    mCaseRange = value
End Property

Public Property Get ReplaceExisting() As Boolean
    ReplaceExisting = mReplaceExisting
End Property

Public Property Let ReplaceExisting(ByVal value As Boolean)
    mReplaceExisting = value
End Property

' False -> append titles as "Output | Case"; True -> "Case | Output".
Public Property Get TitleCaseFirst() As Boolean
    TitleCaseFirst = mTitleCaseFirst
End Property

Public Property Let TitleCaseFirst(ByVal value As Boolean)
    mTitleCaseFirst = value
End Property

Public Property Get MaxAdvanced() As Long
    MaxAdvanced = MAX_ADVANCED
End Property

Public Property Get AdvName(ByVal idx As Long) As String
    AdvName = mAdvName(idx)
End Property

Public Property Let AdvName(ByVal idx As Long, ByVal value As String)
    mAdvName(idx) = value
End Property

Public Property Get AdvRange(ByVal idx As Long) As String
    AdvRange = mAdvRange(idx)
End Property

Public Property Let AdvRange(ByVal idx As Long, ByVal value As String)
    mAdvRange(idx) = value
End Property

Public Property Get AdvCasesText(ByVal idx As Long) As String
    AdvCasesText = mAdvCases(idx)
End Property

Public Property Let AdvCasesText(ByVal idx As Long, ByVal value As String)
    mAdvCases(idx) = value
End Property

' Cases an advanced output applies to, as a 0-based string array.
Public Function AdvCasesArray(ByVal idx As Long) As Variant
    AdvCasesArray = SplitTrim(mAdvCases(idx))
End Function

' True if advanced slot idx has any of its three fields filled in.
Public Function AdvHasContent(ByVal idx As Long) As Boolean
    AdvHasContent = (Trim$(mAdvName(idx)) <> "" _
                Or Trim$(mAdvRange(idx)) <> "" _
                Or Trim$(mAdvCases(idx)) <> "")
End Function

' True if advanced slot idx is fully specified: a name, a named range, and at
' least one case. A slot with only some of these is neither complete nor empty.
Public Function AdvIsComplete(ByVal idx As Long) As Boolean
    Dim arr As Variant

    AdvIsComplete = False
    If Trim$(mAdvName(idx)) = "" Then Exit Function
    If Trim$(mAdvRange(idx)) = "" Then Exit Function

    arr = SplitTrim(mAdvCases(idx))
    If Not IsArray(arr) Then Exit Function
    If UBound(arr) < 0 Then Exit Function

    AdvIsComplete = True
End Function

' True if advanced output idx is fully defined and applies to caseName.
Public Function AdvAppliesTo(ByVal idx As Long, ByVal caseName As String) As Boolean
    Dim arr As Variant, j As Long

    AdvAppliesTo = False
    If Not AdvIsComplete(idx) Then Exit Function

    arr = SplitTrim(mAdvCases(idx))
    For j = LBound(arr) To UBound(arr)
        If StrComp(CStr(arr(j)), caseName, vbTextCompare) = 0 Then
            AdvAppliesTo = True
            Exit Function
        End If
    Next j
End Function

Public Property Get CasesText() As String
    CasesText = mCases
End Property

Public Property Let CasesText(ByVal value As String)
    mCases = value
End Property

' Return the trimmed, non-empty case names as a 0-based string array.
' Returns an empty array if none are present.
Public Function CasesArray() As Variant
    CasesArray = SplitTrim(mCases)
End Function

' Split a comma-separated string into a 0-based array of trimmed, non-empty
' tokens (empty Array() when there are none).
Private Function SplitTrim(ByVal s As String) As Variant
    Dim parts As Variant
    Dim result() As String
    Dim count As Long
    Dim i As Long
    Dim token As String

    If Trim$(s) = "" Then
        SplitTrim = Array()
        Exit Function
    End If

    parts = Split(s, ",")
    ReDim result(0 To UBound(parts))
    count = 0
    For i = LBound(parts) To UBound(parts)
        token = Trim$(parts(i))
        If token <> "" Then
            result(count) = token
            count = count + 1
        End If
    Next i

    If count = 0 Then
        SplitTrim = Array()
    Else
        ReDim Preserve result(0 To count - 1)
        SplitTrim = result
    End If
End Function

' Blank the output names and ranges and clear every case selection (main and
' advanced). NumOutputs, the case-range name, and the title/replace options are
' left as-is. Used by the main dialog's Reset All.
Public Sub ClearAll()
    Dim i As Long

    For i = 1 To MAX_OUTPUTS
        mNames(i) = ""
        mRanges(i) = ""
    Next i
    mCases = ""
    For i = 1 To MAX_ADVANCED
        mAdvName(i) = ""
        mAdvRange(i) = ""
        mAdvCases(i) = ""
    Next i
End Sub

' Reset to the built-in starter configuration.
Public Sub LoadDefaults()
    Dim defaultNames As Variant
    Dim defaultRanges As Variant
    Dim i As Long

    defaultNames = Array("Cash Flows", "Valuations & Returns", "Operating Build", "Output 4", "Output 5")
    defaultRanges = Array("cf", "ret", "op", "out4", "out5")

    mNumOutputs = 2
    For i = 1 To MAX_OUTPUTS
        mNames(i) = defaultNames(i - 1)
        mRanges(i) = defaultRanges(i - 1)
    Next i
    mCaseRange = "case"
    mCases = "Base, Upside, Downside"
    mReplaceExisting = False
    mTitleCaseFirst = False

    For i = 1 To MAX_ADVANCED
        mAdvName(i) = ""
        mAdvRange(i) = ""
        mAdvCases(i) = ""
    Next i
End Sub

' Load the last-used preset from the registry, falling back to defaults for any
' value that has not been saved yet.
Public Sub LoadFromRegistry()
    Dim raw As String
    Dim i As Long

    Call LoadDefaults

    On Error GoTo Fallback

    raw = GetSetting(REG_APP, REG_SECTION, "NumOutputs", "")
    If raw <> "" And IsNumeric(raw) Then NumOutputs = CLng(raw)

    For i = 1 To MAX_OUTPUTS
        raw = GetSetting(REG_APP, REG_SECTION, "Name" & i, "")
        If raw <> "" Then mNames(i) = raw
        raw = GetSetting(REG_APP, REG_SECTION, "Range" & i, "")
        If raw <> "" Then mRanges(i) = raw
    Next i

    raw = GetSetting(REG_APP, REG_SECTION, "CaseRange", "")
    If raw <> "" Then mCaseRange = raw

    raw = GetSetting(REG_APP, REG_SECTION, "Cases", "")
    If raw <> "" Then mCases = raw

    raw = GetSetting(REG_APP, REG_SECTION, "Replace", "")
    If raw <> "" Then mReplaceExisting = (raw = "1")

    raw = GetSetting(REG_APP, REG_SECTION, "TitleCaseFirst", "")
    If raw <> "" Then mTitleCaseFirst = (raw = "1")

    For i = 1 To MAX_ADVANCED
        raw = GetSetting(REG_APP, REG_SECTION, "AdvName" & i, "")
        If raw <> "" Then mAdvName(i) = raw
        raw = GetSetting(REG_APP, REG_SECTION, "AdvRange" & i, "")
        If raw <> "" Then mAdvRange(i) = raw
        raw = GetSetting(REG_APP, REG_SECTION, "AdvCases" & i, "")
        If raw <> "" Then mAdvCases(i) = raw
    Next i

Fallback:
    On Error GoTo 0
End Sub

' Persist the current configuration as the remembered preset. Call this only
' after a successful paste.
Public Sub SaveToRegistry()
    Dim i As Long

    On Error Resume Next
    SaveSetting REG_APP, REG_SECTION, "NumOutputs", CStr(mNumOutputs)
    For i = 1 To MAX_OUTPUTS
        SaveSetting REG_APP, REG_SECTION, "Name" & i, mNames(i)
        SaveSetting REG_APP, REG_SECTION, "Range" & i, mRanges(i)
    Next i
    SaveSetting REG_APP, REG_SECTION, "CaseRange", mCaseRange
    SaveSetting REG_APP, REG_SECTION, "Cases", mCases
    SaveSetting REG_APP, REG_SECTION, "Replace", IIf(mReplaceExisting, "1", "0")
    SaveSetting REG_APP, REG_SECTION, "TitleCaseFirst", IIf(mTitleCaseFirst, "1", "0")
    For i = 1 To MAX_ADVANCED
        SaveSetting REG_APP, REG_SECTION, "AdvName" & i, mAdvName(i)
        SaveSetting REG_APP, REG_SECTION, "AdvRange" & i, mAdvRange(i)
        SaveSetting REG_APP, REG_SECTION, "AdvCases" & i, mAdvCases(i)
    Next i
    On Error GoTo 0
End Sub

' Validate the active configuration. Returns False and a user-facing message in
' errMsg when something is missing or out of range.
Public Function Validate(ByRef errMsg As String) As Boolean
    Dim i As Long

    Validate = False

    If mNumOutputs < 1 Or mNumOutputs > MAX_OUTPUTS Then
        errMsg = "Number of outputs must be between 1 and " & MAX_OUTPUTS & "."
        Exit Function
    End If

    For i = 1 To mNumOutputs
        If Trim$(mNames(i)) = "" Then
            errMsg = "Output #" & i & " is missing a display name."
            Exit Function
        End If
        If Trim$(mRanges(i)) = "" Then
            errMsg = "Output #" & i & " is missing a named range."
            Exit Function
        End If
    Next i

    If Trim$(mCaseRange) = "" Then
        errMsg = "Enter the case input cell's named range."
        Exit Function
    End If

    If UBound(CasesArray()) < 0 Then
        errMsg = "Enter at least one case name (comma-separated)."
        Exit Function
    End If

    If Not ValidateAdvanced(errMsg) Then Exit Function

    Validate = True
End Function

' Validate only the advanced outputs. Each slot must be either fully empty or
' fully specified (name + named range + at least one case). Shared by the main
' Validate and the advanced dialog's OK handler so both enforce the same rule.
Public Function ValidateAdvanced(ByRef errMsg As String) As Boolean
    Dim i As Long

    ValidateAdvanced = False
    For i = 1 To MAX_ADVANCED
        If AdvHasContent(i) And Not AdvIsComplete(i) Then
            errMsg = "Extra output #" & (i + 2) & " is incomplete." & vbCrLf & _
                     "Provide a name, a named range, and at least one case " & _
                     "- or clear all three fields."
            Exit Function
        End If
    Next i

    ValidateAdvanced = True
End Function

' --- History serialization ------------------------------------------------
'
' Serialize the whole configuration (basic + advanced) to a single string for
' history storage, and restore it. The field separator is Chr(1), which cannot
' appear in typed names/ranges, so no escaping is needed. The field order and
' count are fixed; Deserialize rejects a string whose field count does not match.

Private Function FieldSep() As String
    FieldSep = Chr$(1)
End Function

' Number of separated fields a serialized record must contain.
Private Function SerialFieldCount() As Long
    SerialFieldCount = 1 + 2 * MAX_OUTPUTS + 4 + 3 * MAX_ADVANCED
End Function

Public Function Serialize() As String
    Dim sep As String, i As Long, s As String

    sep = FieldSep()
    s = CStr(mNumOutputs)
    For i = 1 To MAX_OUTPUTS
        s = s & sep & mNames(i) & sep & mRanges(i)
    Next i
    s = s & sep & mCaseRange
    s = s & sep & mCases
    s = s & sep & IIf(mReplaceExisting, "1", "0")
    s = s & sep & IIf(mTitleCaseFirst, "1", "0")
    For i = 1 To MAX_ADVANCED
        s = s & sep & mAdvName(i) & sep & mAdvRange(i) & sep & mAdvCases(i)
    Next i

    Serialize = s
End Function

' Restore configuration from a Serialize() string. Returns False (leaving the
' config untouched) when the string is malformed.
Public Function Deserialize(ByVal data As String) As Boolean
    Dim parts As Variant, i As Long, idx As Long

    Deserialize = False
    parts = Split(data, FieldSep())
    If UBound(parts) <> SerialFieldCount() - 1 Then Exit Function

    idx = 0
    NumOutputs = CLng(parts(idx)): idx = idx + 1     ' property clamps to range
    For i = 1 To MAX_OUTPUTS
        mNames(i) = CStr(parts(idx)): idx = idx + 1
        mRanges(i) = CStr(parts(idx)): idx = idx + 1
    Next i
    mCaseRange = CStr(parts(idx)): idx = idx + 1
    mCases = CStr(parts(idx)): idx = idx + 1
    mReplaceExisting = (CStr(parts(idx)) = "1"): idx = idx + 1
    mTitleCaseFirst = (CStr(parts(idx)) = "1"): idx = idx + 1
    For i = 1 To MAX_ADVANCED
        mAdvName(i) = CStr(parts(idx)): idx = idx + 1
        mAdvRange(i) = CStr(parts(idx)): idx = idx + 1
        mAdvCases(i) = CStr(parts(idx)): idx = idx + 1
    Next i

    Deserialize = True
End Function

' One-line summary of the selected output names (with "+N adv" when advanced
' outputs are configured), for the history list's first column.
Public Function TitleSummary() As String
    Dim i As Long, s As String

    For i = 1 To mNumOutputs
        If Trim$(mNames(i)) <> "" Then
            If s <> "" Then s = s & ", "
            s = s & mNames(i)
        End If
    Next i

    ' Name the advanced (extra) outputs explicitly rather than just counting them.
    For i = 1 To MAX_ADVANCED
        If AdvIsComplete(i) Then
            If s <> "" Then s = s & ", "
            s = s & mAdvName(i)
        End If
    Next i

    TitleSummary = s
End Function

' The selected cases, for the history list's second column.
Public Function CasesSummary() As String
    CasesSummary = mCases
End Function

' A full, multi-line description of the run for the history details pane, so
' every output and case stays visible regardless of the dialog size.
Public Function DetailSummary() As String
    Dim i As Long, s As String, names As String

    For i = 1 To mNumOutputs
        If Trim$(mNames(i)) <> "" Then
            If names <> "" Then names = names & ", "
            names = names & mNames(i)
        End If
    Next i

    s = "Outputs: " & names
    s = s & vbNewLine & "Cases: " & mCases

    For i = 1 To MAX_ADVANCED
        If AdvIsComplete(i) Then
            s = s & vbNewLine & "Extra output " & (i + 2) & " (" & mAdvName(i) & "): " & mAdvCases(i)
        End If
    Next i

    DetailSummary = s
End Function
