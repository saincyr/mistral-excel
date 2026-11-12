Attribute VB_Name = "MistralChat"
Option Explicit

' ============================================================
'  Mistral Chat for Excel 2016 for Mac (v15.x)  -  VBA version
'  No add-in required. Calls the Mistral API via curl.
'
'  HOW TO USE
'   1. Paste your API key in MISTRAL_KEY below.
'   2. Type a prompt in a cell, select it, and run MistralChat
'      (Tools > Macro > Macros, or a shape assigned to it).
'      The reply + actions are logged on the "Mistral Log" sheet.
' ============================================================

Private Const MISTRAL_KEY As String = "PASTE_YOUR_MISTRAL_API_KEY_HERE"
Private Const MODEL As String = "mistral-small-latest"

Private Const SYSTEM_PROMPT As String = _
    "You are an assistant embedded in the user is Excel workbook." & vbLf & _
    "You receive the workbook context (sheet names and sample data) and a user prompt." & vbLf & _
    "Answer ONLY with plain-text lines in these formats:" & vbLf & _
    "REPLY: one or two short sentences, in the same language as the user" & vbLf & _
    "FORMULA|SheetName|CellReference|ExcelFormula" & vbLf & _
    "VALUE|SheetName|CellReference|single value" & vbLf & _
    "Rules:" & vbLf & _
    "- Use several FORMULA or VALUE lines for several cells." & vbLf & _
    "- Formulas: EN syntax, comma separators, starting with =." & vbLf & _
    "- Only use sheet names that appear in the context. Never invent sheet names." & vbLf & _
    "- Do not use double quotes, backslashes or pipe characters in any line." & vbLf & _
    "- If no cell should be written, output only the REPLY line."

' ------------------------------------------------------------
' MAIN ENTRY POINT
' ------------------------------------------------------------
Public Sub MistralChat()
    Dim prompt As String, context As String, raw As String, content As String

    If Left$(MISTRAL_KEY, 6) = "PASTE_" Then
        MsgBox "First open the VBA editor (Tools > Macro > Visual Basic Editor)," & vbLf & _
               "open the MistralChat module and replace" & vbLf & _
               "PASTE_YOUR_MISTRAL_API_KEY_HERE with your key from console.mistral.ai", _
               vbExclamation, "Mistral Chat"
        Exit Sub
    End If

    prompt = GetPrompt()
    If Len(Trim$(prompt)) = 0 Then
        MsgBox "Type your prompt in a cell, select it, then run MistralChat again.", _
               vbInformation, "Mistral Chat"
        Exit Sub
    End If

    Application.StatusBar = "Mistral: reading workbook and thinking..."
    On Error GoTo Fail
    context = BuildContext()
    raw = CallMistral("WORKBOOK CONTEXT:" & vbLf & context & vbLf & vbLf & _
                      "USER PROMPT:" & vbLf & prompt)

    If InStr(raw, Chr$(34) & "content" & Chr$(34)) = 0 Then
        Application.StatusBar = False
        MsgBox "Mistral API error:" & vbLf & Left$(raw, 400), vbCritical, "Mistral Chat"
        Exit Sub
    End If

    content = ExtractContent(raw)
    ApplyOutput content, prompt
    Application.StatusBar = False
    Exit Sub

Fail:
    Application.StatusBar = False
    MsgBox "Error: " & Err.Description, vbCritical, "Mistral Chat"
End Sub

' ------------------------------------------------------------
' PROMPT INPUT: named cell AI_PROMPT, else selected cell, else InputBox
' ------------------------------------------------------------
Private Function GetPrompt() As String
    Dim c As Range
    On Error Resume Next
    Set c = ThisWorkbook.Names("AI_PROMPT").RefersToRange
    On Error GoTo 0
    If Not c Is Nothing Then
        GetPrompt = CStr(c.Value)
    ElseIf Len(CStr(ActiveCell.Value)) > 0 Then
        GetPrompt = CStr(ActiveCell.Value)
    Else
        GetPrompt = InputBox("Your prompt for Mistral:", "Mistral Chat")
    End If
End Function

' ------------------------------------------------------------
' CONTEXT: every sheet, used-range address, first 25 rows x 12 cols
' ------------------------------------------------------------
Private Function BuildContext() As String
    Dim ws As Worksheet, ur As Range
    Dim s As String, v As String
    Dim r As Long, c As Long, rowsN As Long, colsN As Long

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Mistral Log" Then GoTo NextSheet
        s = s & "SHEET: " & ws.Name & " - "
        Set ur = Nothing
        On Error Resume Next
        Set ur = ws.UsedRange
        On Error GoTo 0
        If ur Is Nothing Then
            s = s & "(empty)" & vbLf
        Else
            s = s & "used range " & ur.Address(False, False) & vbLf
            rowsN = ur.Rows.Count: If rowsN > 25 Then rowsN = 25
            colsN = ur.Columns.Count: If colsN > 12 Then colsN = 12
            For r = 1 To rowsN
                For c = 1 To colsN
                    v = CStr(ur.Cells(r, c).Value)
                    If Len(v) > 40 Then v = Left$(v, 40)
                    s = s & v & ";"
                Next c
                s = s & vbLf
            Next r
        End If
NextSheet:
    Next ws
    BuildContext = s
End Function

' ------------------------------------------------------------
' API CALL via AppleScript > curl  (works on Excel 15.x for Mac)
' ------------------------------------------------------------
Private Function CallMistral(ByVal userMsg As String) As String
    Dim dq As String, esc As String
    Dim body As String, cmd As String, appleScript As String

    dq = Chr$(34)
    esc = Chr$(92) & dq                  ' backslash + quote (AppleScript escape)

    userMsg = Clean(userMsg)

    body = "{""model"":""" & MODEL & """,""temperature"":0,""messages"":[{""role"":""system"",""content"":""" & Clean(SYSTEM_PROMPT) & """},{""role"":""user"",""content"":""" & userMsg & """}]}"

    cmd = "curl -s --max-time 60 https://api.mistral.ai/v1/chat/completions" & _
          " -H 'Content-Type: application/json'" & _
          " -H 'Authorization: Bearer " & MISTRAL_KEY & "'" & _
          " -d '" & body & "'"

    appleScript = "do shell script " & dq & Replace(cmd, dq, esc) & dq
    CallMistral = MacScript(appleScript)
End Function

' Remove characters that would break shell quoting / JSON
Private Function Clean(ByVal s As String) As String
    s = Replace(s, Chr$(92), "/")        ' backslash
    s = Replace(s, """", " ")            ' double quote
    s = Replace(s, "'", " ")             ' apostrophe
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Clean = Trim$(s)
End Function

' ------------------------------------------------------------
' PARSE the JSON response: extract choices[0].message.content
' ------------------------------------------------------------
Private Function ExtractContent(ByVal raw As String) As String
    Dim m1 As String, m2 As String, p As Long, q As Long, s As String

    m1 = Chr$(34) & "content" & Chr$(34) & ":" & Chr$(34)
    p = InStr(raw, m1)
    If p = 0 Then ExtractContent = raw: Exit Function
    p = p + Len(m1)

    m2 = Chr$(34) & "}," & Chr$(34) & "finish"
    q = InStr(p, raw, m2)
    If q = 0 Then q = Len(raw)

    s = Mid$(raw, p, q - p)
    s = Replace(s, "\" & "n", vbLf)      ' \n -> newline
    s = Replace(s, "\" & Chr$(34), Chr$(34))  ' \" -> "
    s = Replace(s, "\" & "\", "\")            ' \\ -> \
    ExtractContent = s
End Function

' ------------------------------------------------------------
' APPLY the model output: REPLY / FORMULA / VALUE lines
' ------------------------------------------------------------
Private Sub ApplyOutput(ByVal content As String, ByVal prompt As String)
    Dim lines() As String, ln As String, parts() As String
    Dim replyTxt As String, actLog As String, result As String
    Dim i As Long
    Dim logSheet As Worksheet

    lines = Split(content, vbLf)
    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) > 0 Then
            If Len(ln) >= 6 And UCase$(Left$(ln, 6)) = "REPLY:" Then
                replyTxt = Trim$(Mid$(ln, 7))
            ElseIf Len(ln) >= 8 And UCase$(Left$(ln, 8)) = "FORMULA|" Then
                parts = Split(Mid$(ln, 9), "|")
                If UBound(parts) >= 3 Then
                    actLog = actLog & WriteFormula(parts(0), parts(1), JoinFrom(parts, 3)) & vbLf
                End If
            ElseIf Len(ln) >= 6 And UCase$(Left$(ln, 6)) = "VALUE|" Then
                parts = Split(Mid$(ln, 7), "|")
                If UBound(parts) >= 3 Then
                    actLog = actLog & WriteValue(parts(0), parts(1), JoinFrom(parts, 3)) & vbLf
                End If
            End If
        End If
    Next i

    ' --- chat log sheet ---
    Set logSheet = Nothing
    On Error Resume Next
    Set logSheet = ThisWorkbook.Worksheets("Mistral Log")
    On Error GoTo 0
    If logSheet Is Nothing Then
        Set logSheet = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        logSheet.Name = "Mistral Log"
        logSheet.Range("A1").Value = "Time"
        logSheet.Range("B1").Value = "You"
        logSheet.Range("C1").Value = "Mistral"
        logSheet.Range("D1").Value = "Actions"
    End If

    Dim nextRow As Long
    nextRow = logSheet.Cells(logSheet.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2
    logSheet.Cells(nextRow, 1).Value = Format(Now, "hh:nn")
    logSheet.Cells(nextRow, 2).Value = prompt
    logSheet.Cells(nextRow, 3).Value = replyTxt
    logSheet.Cells(nextRow, 4).Value = actLog
    logSheet.Rows.AutoFit

    ' --- result popup ---
    result = replyTxt
    If Len(actLog) > 0 Then result = result & vbLf & vbLf & actLog
    MsgBox result, vbInformation, "Mistral Chat"
End Sub

Private Function JoinFrom(ByVal p() As String, ByVal startIdx As Long) As String
    Dim i As Long, s As String
    For i = startIdx To UBound(p)
        s = s & p(i)
        If i < UBound(p) Then s = s & "|"
    Next i
    JoinFrom = s
End Function

Private Function GetSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Set ws = ActiveSheet
    Set GetSheet = ws
End Function

Private Function WriteFormula(ByVal sheetName As String, ByVal cellRef As String, _
                              ByVal formula As String) As String
    Dim ws As Worksheet
    On Error GoTo Fail
    Set ws = GetSheet(sheetName)
    If Left$(formula, 1) <> "=" Then formula = "=" & formula
    ws.Range(cellRef).Formula = formula
    WriteFormula = "OK - formula written to " & ws.Name & "!" & cellRef
    Exit Function
Fail:
    WriteFormula = "FAILED - " & sheetName & "!" & cellRef & ": " & Err.Description
End Function

Private Function WriteValue(ByVal sheetName As String, ByVal cellRef As String, _
                            ByVal v As String) As String
    Dim ws As Worksheet
    On Error GoTo Fail
    Set ws = GetSheet(sheetName)
    ws.Range(cellRef).Value = v
    WriteValue = "OK - value written to " & ws.Name & "!" & cellRef
    Exit Function
Fail:
    WriteValue = "FAILED - " & sheetName & "!" & cellRef & ": " & Err.Description
End Function
