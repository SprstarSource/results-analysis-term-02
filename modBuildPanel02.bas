Attribute VB_Name = "modBuildPanel02"
Option Explicit

' Import only the matching term panel into each independent tool workbook.
Private Const PANEL_SHEET As String = "ControlPanel"
Private Const RESULTS_FILE As String = "Term02Results.xlsx"
Private Const SOURCE_FOLDER As String = "Source02"
Private Const STATUS_SHEET As String = "Term02Status"
Private Const BUTTON_PAD As Double = 8
Private Const BUTTON_HEIGHT As Double = 54
Private Const BUTTON_GAP As Double = 12

Private Function PanelActions() As Variant
    ' Second Term Analysis, Cumulative Term Average and Split Term Results work on the raw files of Source02: no previous step is required.
    PanelActions = Array( _
        Array("btnCollect", "Second Term Results", "CollectTerm02Results", False, &H60AE27), _
        Array("btnDistribute", "Second Term Analysis", "modBuildPanel02.Panel_RunAnalysis", False, &HB98029), _
        Array("btnAverage", "Cumulative Term Average", "CumulativeTermAverage", False, &HAD448E), _
        Array("btnSplitTerms", "Split Term Results", "modSplitTermResults.SplitTermResults", False, &H85A016), _
        Array("btnExamen", "Examen Analysis", "modExamenAnalysis02.Run_ExamenAnalysis02", False, &HA05568), _
        Array("btnCompare", "Compare", "modCompare.Run_Compare", False, &H00B0E8F0), _
        Array("btnOrientationAll", "Term 02 Orientation", "modOrientationAll02.Run_OrientationAll02", False, &HB98029))
End Function

Public Sub BuildControlPanel()
    Dim ws As Worksheet, btn As Shape, actions As Variant, item As Variant
    Dim i As Long, count As Long, columns As Long, gridRow As Long, gridCol As Long
    Dim rowItems As Long, canvasWidth As Double, rowWidth As Double, btnWidth As Double
    Dim originLeft As Double, originTop As Double, x As Double, y As Double
    Dim prevScreen As Boolean, errorText As String

    prevScreen = Application.ScreenUpdating
    On Error GoTo BuildError
    Application.ScreenUpdating = False
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 801, , "The tool workbook is read-only."
    Set ws = PanelSheet()
    If ws Is Nothing Then
        If ThisWorkbook.ProtectStructure Then Err.Raise vbObjectError + 802, , "The workbook structure is protected."
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = PANEL_SHEET
    End If
    If ws.ProtectContents Or ws.ProtectDrawingObjects Then
        Err.Raise vbObjectError + 803, , "Unprotect the control panel sheet before rebuilding it."
    End If
    ws.Activate
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    With ws
        .Cells.UnMerge
        .Cells.Clear
        .DisplayRightToLeft = True
        .Columns("A").ColumnWidth = 3
        .Columns("B").ColumnWidth = 5
        .Columns("C").ColumnWidth = 36
        .Columns("D").ColumnWidth = 14
        .Columns("E").ColumnWidth = 14
        .Columns("F:H").ColumnWidth = 20
        .Rows(1).RowHeight = 48
        .Rows(2).RowHeight = 14
        ' The title bar ends exactly where the buttons end (column H).
        With .Range("B1:H1")
            .Merge
            .Value = Panel_Ar("62A 62D 644 64A 644 20 646 62A 627 626 62C 20 627 644 641 635 644 20 30 32") & "  |  Second Term Management Dashboard"
            .Font.Name = "Segoe UI"
            .Font.Size = 16
            .ShrinkToFit = True
            .WrapText = False
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(28, 55, 103)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End With

    actions = PanelActions()
    count = UBound(actions) - LBound(actions) + 1
    columns = 2
    ' Buttons sit on the right side of the sheet, inside columns F:H (next to the tables).
    btnWidth = (ws.Range("F3:H3").Width - 2 * BUTTON_PAD - (columns - 1) * BUTTON_GAP) / columns
    canvasWidth = columns * btnWidth + (columns - 1) * BUTTON_GAP
    originLeft = ws.Range("F3").Left + BUTTON_PAD
    originTop = ws.Range("F3").Top
    For i = 0 To count - 1
        item = actions(i)
        gridRow = i \ columns
        gridCol = i Mod columns
        rowItems = count - gridRow * columns
        If rowItems > columns Then rowItems = columns
        rowWidth = rowItems * btnWidth + (rowItems - 1) * BUTTON_GAP
        x = originLeft + (canvasWidth - rowWidth) / 2 + gridCol * (btnWidth + BUTTON_GAP)
        y = originTop + gridRow * (BUTTON_HEIGHT + BUTTON_GAP)
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, btnWidth, BUTTON_HEIGHT)
        With btn
            .Name = CStr(item(0))
            .Adjustments(1) = 0.12
            .Placement = xlFreeFloating
            .Line.Weight = 1.25
            With .TextFrame2
                .TextRange.Text = CStr(item(1))
                .TextRange.Font.Name = "Segoe UI"
                .TextRange.Font.Size = 12
                .TextRange.Font.Bold = msoTrue
                .TextRange.ParagraphFormat.Alignment = msoAlignCenter
                .VerticalAnchor = msoAnchorMiddle
                .WordWrap = msoTrue
                .MarginLeft = 7
                .MarginRight = 7
                .MarginTop = 4
                .MarginBottom = 4
            End With
        End With
    Next i
    y = originTop + ((count + columns - 1) \ columns) * (BUTTON_HEIGHT + BUTTON_GAP)
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, originLeft, y, canvasWidth, 40)
    With btn
        .Name = "btnOpenOutput"
        .Fill.Solid: .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Line.ForeColor.RGB = RGB(89, 89, 89)
        .OnAction = PanelMacroAction("modBuildPanel02.Panel_OpenOutput")
        With .TextFrame2
            .TextRange.Text = Panel_Ar("641 62A 62D 20 645 62C 644 62F 20 627 644 625 635 62F 627 631")
            .TextRange.Font.Size = 14: .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            .VerticalAnchor = msoAnchorMiddle
        End With
    End With
    Panel_BuildOutputTables ws
    If Not ActiveWindow Is Nothing Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
    End If
    RefreshPanelState
    CheckPanelMacros actions
    RemoveOldWorkflowNames
    Panel_SyncDurations
    Panel_DeleteStatusSheet
    ws.Activate
    Application.ScreenUpdating = prevScreen
    MsgBox "Control panel updated: " & count & " buttons in two columns.", vbInformation, "Done"
    Exit Sub
BuildError:
    errorText = CStr(Err.Number) & " : " & Err.Description
    Application.ScreenUpdating = prevScreen
    MsgBox "Panel build failed: " & errorText, vbExclamation, "Error"
End Sub

Public Sub RefreshPanelState()
    Dim ws As Worksheet, btn As Shape, actions As Variant, item As Variant
    Dim i As Long, enabled As Boolean, resultsExist As Boolean
    Set ws = PanelSheet()
    If ws Is Nothing Then Exit Sub
    resultsExist = PanelResultsExist()
    actions = PanelActions()
    For i = LBound(actions) To UBound(actions)
        item = actions(i)
        Set btn = Nothing
        On Error Resume Next
        Set btn = ws.Shapes(CStr(item(0)))
        Err.Clear
        On Error GoTo 0
        If Not btn Is Nothing Then
            enabled = True
            If CBool(item(3)) Then enabled = resultsExist
            With btn
                .Fill.Solid
                If enabled Then
                    .Fill.ForeColor.RGB = CLng(item(4))
                    .Line.ForeColor.RGB = CLng(item(4))
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                    .OnAction = PanelMacroAction(CStr(item(2)))
                Else
                    .Fill.ForeColor.RGB = RGB(220, 220, 220)
                    .Line.ForeColor.RGB = RGB(180, 180, 180)
                    .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(120, 120, 120)
                    .OnAction = PanelMacroAction("ResultsGuard")
                End If
            End With
        End If
    Next i
    Panel_BuildOutputTables ws
End Sub

Private Sub Panel_BuildOutputTables(ByVal ws As Worksheet)
    Dim map As Object, key As Variant, fso As Object, folder As Object, f As Object
    Dim r As Long, startRow As Long, outDir As String, idx As Long, isDone As Boolean, d As Double
    On Error Resume Next
    Panel_SyncDurations
    startRow = 3
    ws.Range("B" & startRow & ":E500").Clear
    ws.Rows(startRow & ":500").RowHeight = ws.StandardHeight
    Panel_TableTitle ws, startRow, Panel_Ar("639 645 644 64A 629 20 627 644 62A 648 62C 64A 647")
    Panel_TableHeader ws, startRow + 1
    Panel_StepRow ws, startRow + 2, 1, Panel_Ar("627 633 62A 64A 631 627 62F 20 627 644 645 639 62F 644 627 62A")
    Panel_StepRow ws, startRow + 3, 2, Panel_Ar("627 633 62A 64A 631 627 62F 20 627 644 631 63A 628 627 62A")
    Panel_StepRow ws, startRow + 4, 3, Panel_Ar("625 646 62C 627 632 20 627 644 62A 648 62C 64A 647")
    Panel_StepRow ws, startRow + 5, 4, Panel_Ar("62A 635 62F 64A 631 20 646 62A 627 626 62C 20 627 644 62A 648 62C 64A 647")

    ' Orientation output files belong to the orientation table, right after its four steps.
    Set fso = CreateObject("Scripting.FileSystemObject")
    outDir = ThisWorkbook.Path & Application.PathSeparator & PanelOutputFolderName()
    r = startRow + 6
    If fso.FolderExists(outDir) Then
        For Each f In fso.GetFolder(outDir).Files
            If Panel_IsListableOutput(fso, f.Name) Then
                If Panel_IsOrientationFile(f.Name) Then
                    Panel_TableRow ws, r, r - startRow - 1, f.Name, Panel_TxtDone(), -1
                    r = r + 1
                End If
            End If
        Next f
    End If

    startRow = r + 1
    Panel_TableTitle ws, startRow, Panel_Ar("645 644 641 627 62A 20 627 644 62A 62D 644 64A 644")
    Panel_TableHeader ws, startRow + 1
    Set map = baf_GetFileMap()
    idx = 0
    r = startRow + 2
    For Each key In map.Keys
        idx = idx + 1
        isDone = Panel_HasOutputStem(outDir, CStr(key))
        d = -1
        ' Real duration measured by the analysis run; shown only once the file is done.
        If isDone Then
            d = Panel_ReadNumber("PanelDur02_" & idx)
        Else
            Panel_StoreNumber "PanelDur02_" & idx, 0      ' file missing: forget its old duration
        End If
        Panel_TableRow ws, r, idx, CStr(key), IIf(isDone, Panel_TxtDone(), Panel_TxtPending()), d
        r = r + 1
    Next key

    startRow = r + 1
    Panel_TableTitle ws, startRow, Panel_Ar("645 62E 631 62C 627 62A 20 627 644 623 632 631 627 631 20 627 644 623 62E 631 649")
    Panel_TableHeader ws, startRow + 1
    If fso.FolderExists(outDir) Then
        Set folder = fso.GetFolder(outDir)
        r = startRow + 2
        For Each f In folder.Files
            If Panel_IsListableOutput(fso, f.Name) Then
                If Not Panel_IsAnalysisStem(map, f.Name) And Not Panel_IsOrientationFile(f.Name) Then
                    Panel_TableRow ws, r, r - startRow - 1, f.Name, Panel_TxtDone(), -1
                    r = r + 1
                End If
            End If
        Next f
    End If
    Err.Clear
End Sub

' Excel files only; the "~$" lock files created while a workbook is open are ignored.
Private Function Panel_IsListableOutput(ByVal fso As Object, ByVal fileName As String) As Boolean
    If Left$(fileName, 2) = "~$" Then Exit Function
    Panel_IsListableOutput = (LCase$(fso.GetExtensionName(fileName)) Like "xls*")
End Function

' Orientation outputs contain the Arabic word for "orientation" in their name.
Private Function Panel_IsOrientationFile(ByVal fileName As String) As Boolean
    Panel_IsOrientationFile = (InStr(1, fileName, Panel_Ar("627 644 62A 648 62C 64A 647"), vbBinaryCompare) > 0)
End Function

Private Function Panel_HasOutputStem(ByVal folder As String, ByVal stem As String) As Boolean
    Dim f As String
    f = Dir$(folder & Application.PathSeparator & stem & "*.xlsx")
    Panel_HasOutputStem = (Len(f) > 0)
End Function

Private Function Panel_IsAnalysisStem(ByVal map As Object, ByVal fileName As String) As Boolean
    Dim key As Variant
    For Each key In map.Keys
        If StrComp(Left$(fileName, Len(CStr(key))), CStr(key), vbTextCompare) = 0 Then
            Panel_IsAnalysisStem = True
            Exit Function
        End If
    Next key
End Function

Private Sub Panel_TableTitle(ByVal ws As Worksheet, ByVal r As Long, ByVal title As String)
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
        .Merge: .Value2 = title: .Font.Bold = True: .Font.Size = 14
        .Font.Color = RGB(28, 55, 103): .HorizontalAlignment = xlRight
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    ws.Rows(r).RowHeight = 25
End Sub

Private Sub Panel_TableHeader(ByVal ws As Worksheet, ByVal r As Long)
    ws.Cells(r, 2).Value2 = "#"
    ws.Cells(r, 3).Value2 = Panel_Ar("627 644 645 644 641 20 2F 20 627 644 62E 637 648 629")
    ws.Cells(r, 4).Value2 = Panel_Ar("627 644 62D 627 644 629")
    ws.Cells(r, 5).Value2 = Panel_Ar("627 644 645 62F 629 20 28 62B 29")
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
        .Font.Bold = True: .Interior.Color = RGB(28, 55, 103): .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
    End With
End Sub

' duration < 0 means "no duration to show".
Private Sub Panel_TableRow(ByVal ws As Worksheet, ByVal r As Long, ByVal idx As Long, ByVal label As String, ByVal statusText As String, ByVal duration As Double)
    ws.Cells(r, 2).Value2 = idx: ws.Cells(r, 3).Value2 = label: ws.Cells(r, 4).Value2 = statusText
    If duration >= 0 Then
        ws.Cells(r, 5).Value2 = duration
        ws.Cells(r, 5).NumberFormat = "0.0"
    End If
    With ws.Range(ws.Cells(r, 2), ws.Cells(r, 5))
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(191, 191, 191)
        .VerticalAlignment = xlCenter
    End With
    ws.Range(ws.Cells(r, 2), ws.Cells(r, 2)).HorizontalAlignment = xlCenter
    ws.Cells(r, 4).HorizontalAlignment = xlCenter
    ws.Cells(r, 5).HorizontalAlignment = xlCenter
    If StrComp(statusText, Panel_TxtDone(), vbBinaryCompare) = 0 Then
        ws.Cells(r, 4).Interior.Color = RGB(39, 174, 96)
        ws.Cells(r, 4).Font.Color = RGB(255, 255, 255)
        ws.Cells(r, 4).Font.Bold = True
    End If
End Sub

Private Function Panel_TxtDone() As String
    Panel_TxtDone = Panel_Ar("623 646 62C 632")
End Function

Private Function Panel_TxtPending() As String
    Panel_TxtPending = Panel_Ar("641 64A 20 627 644 627 646 62A 638 627 631")
End Function

' ---- Durations kept in hidden workbook names (no status sheet needed) ----
Private Sub Panel_StoreNumber(ByVal nm As String, ByVal v As Double)
    On Error Resume Next
    ThisWorkbook.Names(nm).Delete
    If v > 0 Then ThisWorkbook.Names.Add Name:=nm, RefersTo:="=" & Trim$(Str$(Round(v, 2))), Visible:=False
    Err.Clear
End Sub

Private Function Panel_ReadNumber(ByVal nm As String) As Double
    Dim s As String
    Panel_ReadNumber = -1
    On Error Resume Next
    s = ThisWorkbook.Names(nm).RefersTo
    If Err.Number <> 0 Or Len(s) < 2 Then Err.Clear: Exit Function
    Panel_ReadNumber = Val(Mid$(s, 2))
    If Panel_ReadNumber <= 0 Then Panel_ReadNumber = -1
End Function

'' Called by modOrientationAll02 when a step finishes (seconds <= 0 forgets the value).
Public Sub Panel_RecordStepDuration(ByVal stepNo As Long, ByVal seconds As Double)
    Panel_StoreNumber "PanelStep02_" & stepNo, seconds
End Sub

' Called by modAnalysisBuilderBuildAll02 when an analysis file is finished (seconds <= 0 forgets the value).
Public Sub Panel_RecordFileDuration(ByVal fileNo As Long, ByVal seconds As Double)
    Panel_StoreNumber "PanelDur02_" & fileNo, seconds
End Sub

' Copies the per-file durations written by the analysis run into hidden names.
Private Sub Panel_SyncDurations()
    Dim sh As Worksheet, map As Object, i As Long, v As Variant
    Set sh = Panel_StatusSheet()
    If sh Is Nothing Then Exit Sub
    On Error Resume Next
    Set map = baf_GetFileMap()
    If map Is Nothing Then Exit Sub
    For i = 1 To map.Count
        v = sh.Cells(11 + i, 7).Value2
        If IsNumeric(v) And Len(CStr(v)) > 0 Then
            Panel_StoreNumber "PanelDur02_" & i, CDbl(v)
        Else
            Panel_StoreNumber "PanelDur02_" & i, 0
        End If
    Next i
    Err.Clear
End Sub

Private Function Panel_StatusSheet() As Worksheet
    On Error Resume Next
    Set Panel_StatusSheet = ThisWorkbook.Worksheets(STATUS_SHEET)
    On Error GoTo 0
End Function

Private Sub Panel_DeleteStatusSheet()
    Dim sh As Worksheet, prev As Boolean
    Set sh = Panel_StatusSheet()
    If sh Is Nothing Then Exit Sub
    prev = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    sh.Delete
    Err.Clear
    On Error GoTo 0
    Application.DisplayAlerts = prev
End Sub

Private Function Panel_OrientationStepStatus02(ByVal stepNo As Long) As String
    Dim ok As Boolean
    On Error Resume Next
    Select Case stepNo
        Case 1: ok = modTransfer02.Transfer02Succeeded()
        Case 2: ok = modStreamPreferences02.Preferences02Succeeded()
        Case 3: ok = modOrientation02.Orientation02Succeeded() And Panel_OutputFolderExists()
        Case 4: ok = modOrientationResults02.OrientationResults02Succeeded() And Panel_OutputFolderExists()
    End Select
    Err.Clear
    Panel_OrientationStepStatus02 = IIf(ok, Panel_TxtDone(), Panel_TxtPending())
End Function

' One row of the orientation table; a duration is shown only for a step that is done.
Private Sub Panel_StepRow(ByVal ws As Worksheet, ByVal r As Long, ByVal n As Long, ByVal label As String)
    Dim st As String, d As Double
    st = Panel_OrientationStepStatus02(n)
    d = -1
    If StrComp(st, Panel_TxtDone(), vbBinaryCompare) = 0 Then
        d = Panel_ReadNumber("PanelStep02_" & n)
    Else
        Panel_StoreNumber "PanelStep02_" & n, 0       ' not done (or output folder deleted): forget the old duration
    End If
    Panel_TableRow ws, r, n, label, st, d
End Sub

Private Function Panel_OutputFolderExists() As Boolean
    Dim attrs As Long
    On Error Resume Next
    attrs = GetAttr(ThisWorkbook.Path & Application.PathSeparator & PanelOutputFolderName())
    If Err.Number = 0 Then Panel_OutputFolderExists = ((attrs And vbDirectory) <> 0)
    Err.Clear
End Function

Private Function Panel_Ar(ByVal codes As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(codes, " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    Panel_Ar = s
End Function

Public Sub Panel_RunAnalysis()
    Dim errNumber As Long, errDescription As String
    On Error GoTo RunFailed
    Application.Run PanelMacroAction("modAnalysisBuilderBuildAll02.Build_All")
    Panel_SyncDurations
    Panel_DeleteStatusSheet
    RefreshPanelState
    Exit Sub
RunFailed:
    errNumber = Err.Number
    errDescription = Err.Description
    Panel_SyncDurations
    Panel_DeleteStatusSheet
    RefreshPanelState
    Err.Raise errNumber, "Panel_RunAnalysis", errDescription
End Sub

Public Sub Panel_OpenOutput()
    Dim p As String, fso As Object
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "Save the tool workbook before opening its output folder.", vbExclamation, "Output folder"
        Exit Sub
    End If
    p = ThisWorkbook.Path & Application.PathSeparator & PanelOutputFolderName()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(p) Then fso.CreateFolder p
    ThisWorkbook.FollowHyperlink Address:=p, NewWindow:=True
End Sub

Public Sub ResultsGuard()
    MsgBox "Run Second Term Results first to generate " & RESULTS_FILE & ".", _
           vbExclamation, "Step 1 Required"
End Sub

' Compatibility with buttons created by earlier panel versions.
Public Sub DistributeGuard()
    ResultsGuard
End Sub
Public Sub AverageGuard()
    ResultsGuard
End Sub

' Deletes the hidden names (_Term02_*) written by the old step-by-step workflow, which no longer exists.
Private Sub RemoveOldWorkflowNames()
    Dim i As Long
    On Error Resume Next
    For i = ThisWorkbook.Names.Count To 1 Step -1
        If InStr(1, ThisWorkbook.Names(i).Name, "_Term02_", vbTextCompare) > 0 Then ThisWorkbook.Names(i).Delete
    Next i
    Err.Clear
End Sub

Private Function PanelSheet() As Worksheet
    On Error Resume Next
    Set PanelSheet = ThisWorkbook.Worksheets(PANEL_SHEET)
    On Error GoTo 0
End Function

Private Function PanelResultsExist() As Boolean
    Dim attrs As Long, p As String
    If Len(ThisWorkbook.Path) = 0 Then Exit Function
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then Exit Function
    p = ThisWorkbook.Path & Application.PathSeparator & PanelOutputFolderName() & Application.PathSeparator & RESULTS_FILE
    On Error GoTo Missing
    attrs = GetAttr(p)
    PanelResultsExist = ((attrs And vbDirectory) = 0)
    Exit Function
Missing:
    PanelResultsExist = False
End Function

' Output folder that holds Term02Results.xlsx (hex code points, same name as in modCollector)
Private Function PanelOutputFolderName() As String
    Dim parts() As String, i As Long, s As String
    parts = Split("0645 0639 0627 0644 062C 0629 0020 0628 064A 0627 0646 0627 062A 0020 0627 0644 0641 0635 0644 0020 0030 0032", " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    PanelOutputFolderName = s
End Function

Private Function PanelMacroAction(ByVal macroName As String) As String
    PanelMacroAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & macroName
End Function

Private Sub CheckPanelMacros(ByVal actions As Variant)
    Dim i As Long, item As Variant, missing As String
    For i = LBound(actions) To UBound(actions)
        item = actions(i)
        If Not PanelMacroExists(CStr(item(2))) Then missing = missing & "  - " & CStr(item(2)) & vbCrLf
    Next i
    If Len(missing) > 0 Then
        MsgBox "Import the missing modules before using these actions:" & vbCrLf & missing, _
               vbExclamation, "Missing Macros"
    End If
End Sub

Private Function PanelMacroExists(ByVal macroName As String) As Boolean
    Dim comp As Object, codeMod As Object, txt As String
    Dim moduleName As String, procedureName As String, dotPos As Long
    On Error GoTo AccessDenied
    dotPos = InStrRev(macroName, ".")
    procedureName = macroName
    If dotPos > 0 Then
        moduleName = Left$(macroName, dotPos - 1)
        procedureName = Mid$(macroName, dotPos + 1)
    End If
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 1 Then
            If Len(moduleName) = 0 Or StrComp(comp.Name, moduleName, vbTextCompare) = 0 Then
                Set codeMod = comp.CodeModule
                If codeMod.CountOfLines > 0 Then
                    txt = codeMod.Lines(1, codeMod.CountOfLines)
                    If InStr(1, txt, "Sub " & procedureName & "(", vbTextCompare) > 0 Then
                        PanelMacroExists = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next comp
    Exit Function
AccessDenied:
    ' Macro inspection is optional when VBA project access is disabled.
    PanelMacroExists = True
    Err.Clear
End Function
