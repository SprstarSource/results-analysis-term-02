Attribute VB_Name = "modCollector02"
Option Explicit

'==============================================================================
' Module  : modCollector02
' Purpose : SINGLE collection engine of the tool. Reads every source file found in
'           \Source02\ exactly as it is (.xls, .xlsb, .xlsm, .xlsx): nothing is
'           converted, moved, renamed or saved inside \Source02\.
'           Only real MARKS files are collected, with TWO checks that work together:
'             CHECK 1 (structure): row 6 must hold the identity headers - B name, C birth
'                       date, D sex. The stream-wishes file (headers in row 2) fails here.
'             CHECK 2 (marks): from column F, row 7 onward the cells must be averages
'                       (numbers, or text that stands for a number - converted to a real
'                       number). A file that holds mostly text there fails.
'           A file that fails is skipped, never stops the run, and the report names the
'           check that rejected it. Files are never recognised by their name.
'           BuildSourceWorkbook()  -> collected workbook, in memory (open, unsaved).
'                                     Used by ResultFlow02 and modTransfer02.
'           CollectTerm02Results() -> same workbook saved as the Arabic-named file returned by
'                                     ResultsFileName() (second term results) inside the
'                                     output folder next to the tool (created when missing).
' Entry   : CollectTerm02Results()
'==============================================================================

Private Const DATA_START_ROW       As Long = 6      ' header row in the source sheet
Private Const TRAILING_ROWS_SKIP   As Long = 1      ' footer rows after last used row (original logic: -1). Set 0 if none
Private Const IDENTITY_LAST_COL    As Long = 5      ' columns A:E = identity columns (AutoFit)
Private Const COL_WIDTH            As Double = 5.7  ' width for columns F:LastCol
Private Const FONT_SIZE            As Long = 14
Private Const FONT_ARABIC          As String = "Arabic UI Display Light"
Private Const FONT_NUMERIC         As String = "Bahnschrift SemiBold"
Private Const FONT_FALLBACK        As String = "Calibri"
Private Const SOURCE_FOLDER_NAME   As String = "Source02"
Private Const FIRST_MARK_COL       As Long = 6      ' column F : first average column
Private Const FIRST_MARK_ROW       As Long = 7      ' row 7   : first student row

Private Const PREFERRED_FONT As String = "Cairo"
Private Const FALLBACK_FONT  As String = "Arial"
Private Const HEADER_FONT_SIZE As Double = 11
Private Const HEADER_ROW_HEIGHT As Double = 30


Private g_ActiveArabicFont  As String
Private g_ActiveNumericFont As String

' ==============================================================================
' HELPERS
' ==============================================================================
Private Function FileExistsPath(ByVal p As String) As Boolean
    On Error Resume Next
    FileExistsPath = ((GetAttr(p) And vbDirectory) = 0)
    If Err.Number <> 0 Then FileExistsPath = False: Err.Clear
End Function

Private Function SortedNames(ByVal col As Collection) As Variant
    Dim arr() As String, i As Long, j As Long, tmp As String
    If col.Count = 0 Then SortedNames = Array(): Exit Function
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count: arr(i) = CStr(col(i)): Next i
    For i = 2 To UBound(arr)
        tmp = arr(i): j = i - 1
        Do While j >= 1
            If StrComp(arr(j), tmp, vbTextCompare) <= 0 Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
    SortedNames = arr
End Function

' ==============================================================================
' COLLECTION OF THE SOURCE FILES INTO ONE WORKBOOK (fonts + dynamic tables)
' ==============================================================================
' Checks the Windows font registry (HKLM + HKCU). If the check itself cannot run,
' returns True (keeps the requested font instead of forcing the fallback).
Private Function FontExists(ByVal FontName As String) As Boolean
    Const HKCU As Long = &H80000001
    Const HKLM As Long = &H80000002
    Dim reg As Object, hv As Variant
    Dim names As Variant, types As Variant
    Dim i As Long, p As Long, nm As String, enumOK As Boolean

    On Error GoTo Fallback
    Set reg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

    For Each hv In Array(HKLM, HKCU)
        names = Empty
        If reg.EnumValues(CLng(hv), "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", names, types) = 0 Then
            enumOK = True
            If IsArray(names) Then
                For i = LBound(names) To UBound(names)
                    nm = CStr(names(i))
                    p = InStr(nm, " (")
                    If p > 0 Then nm = Left(nm, p - 1)          ' strip "(TrueType)" / "(OpenType)"
                    ' exact name, or family name of a variable font (e.g. "Bahnschrift" for "Bahnschrift SemiBold")
                    If StrComp(nm, FontName, vbTextCompare) = 0 Then FontExists = True: Exit Function
                    If StrComp(Left(FontName, Len(nm) + 1), nm & " ", vbTextCompare) = 0 Then FontExists = True: Exit Function
                Next i
            End If
        End If
    Next hv

    FontExists = Not enumOK     ' enumeration worked and nothing matched -> False; enumeration failed -> True
    Exit Function
Fallback:
    FontExists = True
End Function

Private Sub SetupFonts()
    If FontExists(FONT_ARABIC) Then g_ActiveArabicFont = FONT_ARABIC Else g_ActiveArabicFont = FONT_FALLBACK
    If FontExists(FONT_NUMERIC) Then g_ActiveNumericFont = FONT_NUMERIC Else g_ActiveNumericFont = FONT_FALLBACK
End Sub

' ==============================================================================
' STRUCTURE CHECK (engine). A results file has its identity headers in row
' DATA_START_ROW: B = name, C = birth date, D = sex. A file with any other layout
' (for example the wishes file, whose headers are in row 2) is skipped and named
' in the report; it never stops the run and it is never recognised by its file name.
' ==============================================================================
Private Function SrcHasResultsLayout(ByVal ws As Worksheet) As Boolean
    SrcHasResultsLayout = SrcHeaderHas(ws, 2, "627 633 645") _
                      And SrcHeaderHas(ws, 3, "645 64A 644 627 62F") _
                      And SrcHeaderHas(ws, 4, "62C 646 633")
End Function

Private Function SrcHeaderHas(ByVal ws As Worksheet, ByVal col As Long, ByVal wordCodes As String) As Boolean
    Dim v As Variant, h As String, w As String
    On Error Resume Next
    v = ws.Cells(DATA_START_ROW, col).MergeArea.Cells(1, 1).Value2
    On Error GoTo 0
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    h = SrcNormHeader(CStr(v))
    w = SrcNormHeader(SrcWord(wordCodes))
    SrcHeaderHas = (Len(w) > 0 And InStr(h, w) > 0)
End Function

Private Function SrcWord(ByVal codes As String) As String
    Dim p As Variant, s As String
    For Each p In Split(codes, " ")
        If Len(p) > 0 Then s = s & ChrW(CLng("&H" & p))
    Next p
    SrcWord = s
End Function

Private Function SrcNormHeader(ByVal t As String) As String
    t = Replace(t, ChrW(&H623), ChrW(&H627))    ' alef variants -> alef
    t = Replace(t, ChrW(&H625), ChrW(&H627))
    t = Replace(t, ChrW(&H622), ChrW(&H627))
    t = Replace(t, ChrW(&H649), ChrW(&H64A))    ' alef maksura -> yaa
    t = Replace(t, ChrW(&H629), ChrW(&H647))    ' taa marbuta -> haa
    t = Replace(t, ChrW(&H640), "")             ' tatweel
    t = Replace(t, ChrW(&HA0), " ")
    t = Replace(t, ChrW(&H200E), "")
    t = Replace(t, ChrW(&H200F), "")
    t = Replace(t, vbTab, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbCr, " ")
    SrcNormHeader = LCase$(Trim$(t))
End Function

' True when the workbook has the sheets of the wishes file (only used to word the report)
Private Function SrcHasWishSheets(ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, "Tecno", vbTextCompare) = 0 Or StrComp(ws.Name, "Litr", vbTextCompare) = 0 Then
            SrcHasWishSheets = True
            Exit Function
        End If
    Next ws
End Function

' ==============================================================================
' CHECK 2 - MARKS : averages must start at F7 (a wishes file holds text there)
' ==============================================================================
' Text that stands for a number ("12.5", "12,5", Arabic-Indic digits, stray spaces or
' hidden direction marks) -> True and the value in d. Any real text -> False.
Private Function TextToMark(ByVal s As String, ByRef d As Double) As Boolean
    Dim i As Long, code As Long, res As String
    Dim digits As Long, dots As Long

    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1))
        If code < 0 Then code = code + 65536
        Select Case code
            Case 48 To 57
                res = res & ChrW(code): digits = digits + 1
            Case &H660 To &H669
                res = res & ChrW(code - &H660 + 48): digits = digits + 1
            Case &H6F0 To &H6F9
                res = res & ChrW(code - &H6F0 + 48): digits = digits + 1
            Case 46, 44, &H66B
                res = res & ".": dots = dots + 1
            Case 45, &H2212
                If Len(res) > 0 Then Exit Function
                res = "-"
            Case 43
                If Len(res) > 0 Then Exit Function
            Case 32, 160, &H200E, &H200F, &H202A To &H202E, &HFEFF&
                ' blank / hidden direction mark: ignored
            Case Else
                Exit Function
        End Select
    Next i
    If digits = 0 Or dots > 1 Then Exit Function
    d = Val(res)                        ' Val always reads "." (independent of the regional settings)
    TextToMark = True
End Function

' Scans the score block (column F onward, from row 7). Text that stands for a number is
' converted IN the array to a real number ("convert to number"). Returns how many cells
' are numbers and how many are real texts.
Private Sub PrepareMarksBlock(ByRef arr As Variant, ByRef numCount As Long, ByRef textCount As Long)
    Dim r As Long, c As Long, v As Variant, d As Double, t As String

    numCount = 0: textCount = 0
    If UBound(arr, 2) < FIRST_MARK_COL Then Exit Sub
    For r = FIRST_MARK_ROW - DATA_START_ROW + 1 To UBound(arr, 1)      ' array row 1 = header row
        For c = FIRST_MARK_COL To UBound(arr, 2)
            v = arr(r, c)
            Select Case VarType(v)
                Case vbEmpty, vbNull
                    ' empty cell
                Case vbDouble, vbSingle, vbInteger, vbLong, vbCurrency, vbDecimal
                    numCount = numCount + 1
                Case vbString
                    t = Trim$(Replace(CStr(v), ChrW(160), " "))
                    If Len(t) > 0 Then
                        If TextToMark(CStr(v), d) Then
                            arr(r, c) = d
                            numCount = numCount + 1
                        Else
                            textCount = textCount + 1
                        End If
                    End If
                Case Else
                    textCount = textCount + 1        ' dates, booleans, errors
            End Select
        Next c
    Next r
End Sub

Public Function ProcessSingleSourceFile(ByRef wbResult As Workbook, _
                                        ByVal SourcePath As String, _
                                        ByVal fileName As String, _
                                        ByVal SheetNumber As Long, _
                                        ByRef ReasonOut As String) As Boolean
    Dim wbSource As Workbook, wsSource As Worksheet, wsResult As Worksheet
    Dim LastRow As Long, LastCol As Long, DataArr As Variant
    Dim errNum As Long, errDesc As String, lastDataRow As Long, wrapEndRow As Long
    Dim idLastCol As Long, lastFind As Range
    Dim tbl As ListObject, tblRange As Range, seenHeaders As Object
    Dim h As Long, headerText As String, genCounter As Long
    Dim numCount As Long, textCount As Long

    ProcessSingleSourceFile = False
    ReasonOut = ""
    On Error GoTo EH

    On Error Resume Next
    Set wbSource = Workbooks.Open(fileName:=SourcePath & fileName, UpdateLinks:=False, _
                                  ReadOnly:=True, AddToMRU:=False, Password:="")
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If wbSource Is Nothing Or errNum <> 0 Then
        ReasonOut = "OPEN_FAILED | " & errDesc: Exit Function
    End If

    If wbSource.Sheets.Count = 0 Then
        ReasonOut = "EMPTY_WORKBOOK (no sheets)": wbSource.Close SaveChanges:=False: Exit Function
    End If

    Set wsSource = wbSource.Sheets(1)

    On Error Resume Next
    LastCol = wsSource.Cells(DATA_START_ROW, wsSource.Columns.Count).End(xlToLeft).Column
    Set lastFind = wsSource.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, _
                                       SearchDirection:=xlPrevious)
    If lastFind Is Nothing Then LastRow = 0 Else LastRow = lastFind.Row - TRAILING_ROWS_SKIP
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If errNum <> 0 Then
        ReasonOut = "DETECTION_FAILED | " & errDesc: wbSource.Close SaveChanges:=False: Exit Function
    End If

    If LastRow < DATA_START_ROW Or LastCol < 1 Then
        ReasonOut = "EMPTY_WORKBOOK (no data at row " & DATA_START_ROW & ")"
        wbSource.Close SaveChanges:=False: Exit Function
    End If

    ' ---- CHECK 1 (structure): not a results layout -> skipped and named in the report
    If Not SrcHasResultsLayout(wsSource) Then
        If SrcHasWishSheets(wbSource) Then
            ReasonOut = "CHECK 1 (headers) - WISHES FILE layout (sheets Tecno / Litr), not a results file, ignored"
        Else
            ReasonOut = "CHECK 1 (headers) - DIFFERENT_STRUCTURE (identity headers not found in row " & DATA_START_ROW & "), ignored"
        End If
        wbSource.Close SaveChanges:=False: Exit Function
    End If

    ' ---- Read source data (before any sheet is created in the result workbook)
    On Error Resume Next
    DataArr = wsSource.Range(wsSource.Cells(DATA_START_ROW, 1), wsSource.Cells(LastRow, LastCol)).Value
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If errNum <> 0 Or Not IsArray(DataArr) Then
        ReasonOut = "DATA_READ_FAILED | " & errDesc
        wbSource.Close SaveChanges:=False
        Exit Function
    End If

    ' ---- CHECK 2 (marks): averages must start at F7. Text that stands for a number is
    '      converted to a real number; a file whose F7 onward is text is skipped.
    PrepareMarksBlock DataArr, numCount, textCount
    If numCount = 0 Or numCount <= textCount Then
        ReasonOut = "CHECK 2 (marks) - from F" & FIRST_MARK_ROW & " onward the cells are text, not averages" & _
                    " (numbers: " & numCount & ", texts: " & textCount & "), ignored"
        wbSource.Close SaveChanges:=False
        Exit Function
    End If

    On Error Resume Next
    Set wsResult = wbResult.Sheets.Add(After:=wbResult.Sheets(wbResult.Sheets.Count))
    errNum = Err.Number: errDesc = Err.Description: Err.Clear
    On Error GoTo EH

    If wsResult Is Nothing Then
        ReasonOut = "ADD_SHEET_FAILED | " & errDesc: wbSource.Close SaveChanges:=False: Exit Function
    End If

    On Error Resume Next
    wsResult.Name = Format(SheetNumber, "00")
    Err.Clear
    On Error GoTo EH

    wsResult.DisplayRightToLeft = True

    lastDataRow = 1 + UBound(DataArr, 1)     ' header at row 2 + data rows

    ' ---- Number formats BEFORE writing data (so "@" columns store text)
    With wsResult
        .Columns("B").NumberFormat = "@"
        .Columns("C").NumberFormat = "yyyy-mm-dd"
        If LastCol >= 4 Then .Columns("D").NumberFormat = "@"
        If LastCol >= 5 Then .Columns("E").NumberFormat = "@"
        If LastCol >= 6 Then .Range(.Columns("F"), .Columns(LastCol)).NumberFormat = "0.00"
    End With

    wsResult.Range("A2").Resize(UBound(DataArr, 1), UBound(DataArr, 2)).Value = DataArr

    ' ---- Unique, non-empty headers (case-insensitive, like Excel tables)
    Set seenHeaders = CreateObject("Scripting.Dictionary")
    seenHeaders.CompareMode = vbTextCompare
    genCounter = 0

    For h = 1 To LastCol
        headerText = Trim(CStr(wsResult.Cells(2, h).Value))
        If headerText = "" Or seenHeaders.Exists(headerText) Then
            genCounter = genCounter + 1
            Do While seenHeaders.Exists("Column" & genCounter): genCounter = genCounter + 1: Loop
            wsResult.Cells(2, h).Value = "Column" & genCounter
            seenHeaders.Add "Column" & genCounter, True
        Else
            seenHeaders.Add headerText, True
        End If
    Next h

    ' ---- Dynamic table (range ends exactly at the last data row)
    Set tblRange = wsResult.Range(wsResult.Cells(2, 1), wsResult.Cells(lastDataRow, LastCol))
    Set tbl = Nothing

    On Error Resume Next
    Set tbl = wsResult.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    If Err.Number = 0 Then
        tbl.Name = "Data" & Format(SheetNumber, "00")
        tbl.TableStyle = ""
        tbl.ShowTableStyleColumnStripes = False
        tbl.ShowTableStyleRowStripes = False
        tbl.ShowTableStyleFirstColumn = False
        tbl.ShowTableStyleLastColumn = False
        tbl.ShowAutoFilter = True
    End If
    Err.Clear
    On Error GoTo EH

    ' ---- Formatting
    With wsResult
        .Columns("B").Font.Name = g_ActiveArabicFont: .Columns("B").Font.Size = FONT_SIZE
        .Columns("C").Font.Name = g_ActiveNumericFont: .Columns("C").Font.Size = FONT_SIZE
        If LastCol >= 4 Then .Columns("D").Font.Name = g_ActiveArabicFont: .Columns("D").Font.Size = FONT_SIZE
        If LastCol >= 5 Then .Columns("E").Font.Name = g_ActiveArabicFont: .Columns("E").Font.Size = FONT_SIZE
        If LastCol >= 6 Then
            With .Range(.Columns("F"), .Columns(LastCol))
                .Font.Name = g_ActiveNumericFont
                .Font.Size = FONT_SIZE
            End With
        End If

        wrapEndRow = lastDataRow
        If wrapEndRow < 2 Then wrapEndRow = 2
        .Range(.Cells(2, 1), .Cells(wrapEndRow, LastCol)).WrapText = False

        With .Range(.Cells(1, 1), .Cells(lastDataRow, LastCol))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        ' Identity columns: AutoFit. Score columns: fixed width, rotated bold headers.
        idLastCol = IDENTITY_LAST_COL
        If LastCol < idLastCol Then idLastCol = LastCol
        .Range(.Columns(1), .Columns(idLastCol)).EntireColumn.AutoFit

        If LastCol >= 6 Then .Range(.Columns(IDENTITY_LAST_COL + 1), .Columns(LastCol)).ColumnWidth = COL_WIDTH

        If LastCol >= 4 Then
            .Range(.Cells(2, 4), .Cells(2, LastCol)).Orientation = xlUpward
            .Range(.Cells(2, 4), .Cells(2, LastCol)).Font.Bold = True
        End If
    End With

    ' Apply title formatting last so table column fonts do not override it.
    PlaceHeaderTitle wsSource, wsResult

    On Error Resume Next
    wbSource.Close SaveChanges:=False
    Set wbSource = Nothing
    Err.Clear
    On Error GoTo 0

    ProcessSingleSourceFile = True
    Exit Function

EH:
    ReasonOut = "UNEXPECTED | " & Err.Number & " : " & Err.Description
    Err.Clear
    On Error Resume Next
    If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
    If Not wsResult Is Nothing Then wsResult.Delete
    Err.Clear
    ProcessSingleSourceFile = False
End Function

Public Sub CollectAllSources(ByRef wbResult As Workbook, _
                             ByVal SourcePath As String, _
                             ByRef CopiedCount As Long, _
                             ByRef SkippedCount As Long, _
                             ByRef SkippedList As String)
    Dim fileName As String, SheetNumber As Long, Reason As String
    Dim names As Collection, sorted As Variant, i As Long
    Dim f As String, ext As String, baseName As String, dotPos As Long

    CopiedCount = 0: SkippedCount = 0: SheetNumber = 1: SkippedList = ""

    ' Collect names first (Dir cannot be nested with other Dir calls)
    Set names = New Collection
    fileName = Dir(SourcePath & "*.xls*")
    Do While fileName <> ""
        If Left(fileName, 2) <> "~$" Then
            dotPos = InStrRev(fileName, ".")
            ext = LCase(Mid(fileName, dotPos + 1))
            If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Or ext = "xlsb" Then names.Add fileName
        End If
        fileName = Dir
    Loop

    If names.Count = 0 Then
        SkippedList = "  - " & SOURCE_FOLDER_NAME & " folder is empty or contains no Excel files:" & vbCrLf & _
                      "    " & SourcePath & vbCrLf
        Exit Sub
    End If

    sorted = SortedNames(names)     ' deterministic sheet order (01, 02, ...)
    SetupFonts

    For i = LBound(sorted) To UBound(sorted)
        f = CStr(sorted(i))
        dotPos = InStrRev(f, ".")
        ext = LCase(Mid(f, dotPos + 1))
        baseName = Left(f, dotPos - 1)

        If StrComp(SourcePath & f, ThisWorkbook.FullName, vbTextCompare) = 0 Then
            ' skip the macro workbook itself
        ElseIf IsOpenInExcel(SourcePath & f) Then
            ' never touch (or close) a file the user already has open
            SkippedCount = SkippedCount + 1
            SkippedList = SkippedList & "  - " & f & " : ALREADY OPEN in Excel (close it and run again)" & vbCrLf
        ElseIf ext <> "xlsx" And FileExistsPath(SourcePath & baseName & ".xlsx") Then
            SkippedCount = SkippedCount + 1
            SkippedList = SkippedList & "  - " & f & " : DUPLICATE (.xlsx version exists)" & vbCrLf
        ElseIf ProcessSingleSourceFile(wbResult, SourcePath, f, SheetNumber, Reason) Then
            CopiedCount = CopiedCount + 1: SheetNumber = SheetNumber + 1
        Else
            SkippedCount = SkippedCount + 1
            SkippedList = SkippedList & "  - " & f & " : " & Reason & vbCrLf
        End If
    Next i
End Sub

Private Function IsOpenInExcel(ByVal fullPath As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            IsOpenInExcel = True
            Exit Function
        End If
    Next wb
End Function

' ==============================================================================
' SHARED ENGINE: builds the collected workbook from \Source02\ IN MEMORY.
' Returns True with wbOut OPEN and UNSAVED (the caller must close it with
' SaveChanges:=False, or save it). Returns False when nothing usable was found;
' the reasons are in skippedList. The source files are only read, never changed.
' Calculation mode is not touched. AutomationSecurity / DisplayAlerts are restored.
' ==============================================================================
Public Function BuildSourceWorkbook(ByVal srcFolder As String, _
                                    ByRef wbOut As Workbook, _
                                    ByRef copied As Long, _
                                    ByRef skipped As Long, _
                                    ByRef skippedList As String) As Boolean
    Dim wb As Workbook, oldSec As Long, oldAlerts As Boolean

    BuildSourceWorkbook = False
    Set wbOut = Nothing
    copied = 0: skipped = 0: skippedList = ""

    If LCase(Left(srcFolder, 4)) = "http" Then
        skippedList = "  - The workbook is on OneDrive/SharePoint (URL path). Move it to a local folder first." & vbCrLf
        Exit Function
    End If
    If Right(srcFolder, 1) <> "\" Then srcFolder = srcFolder & "\"
    If Dir(srcFolder, vbDirectory) = "" Then
        skippedList = "  - " & SOURCE_FOLDER_NAME & " folder not found: " & srcFolder & vbCrLf
        Exit Function
    End If

    oldSec = Application.AutomationSecurity
    oldAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable   ' block Auto_Open in .xlsm/.xls
    On Error GoTo EH

    Set wb = Workbooks.Add
    Do While wb.Sheets.Count > 1: wb.Sheets(wb.Sheets.Count).Delete: Loop

    CollectAllSources wb, srcFolder, copied, skipped, skippedList

    If copied > 0 Then
        If wb.Sheets.Count > 1 Then
            If Left(wb.Sheets(1).Name, 5) = "Sheet" And wb.Sheets(1).UsedRange.Cells.Count = 1 Then wb.Sheets(1).Delete
        End If
        Set wbOut = wb
        BuildSourceWorkbook = True
    Else
        wb.Close SaveChanges:=False
    End If

Done:
    Application.AutomationSecurity = oldSec
    Application.DisplayAlerts = oldAlerts
    Exit Function

EH:
    skippedList = skippedList & "  - Build failed: " & Err.Number & " : " & Err.Description & vbCrLf
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set wbOut = Nothing
    BuildSourceWorkbook = False
    Resume Done
End Function

' ==============================================================================
' OUTPUT FOLDER (created next to the tool workbook when missing)
' ==============================================================================
' Name of the output folder (hex code points)
Private Function OutputFolderName() As String
    Dim parts() As String, i As Long, s As String
    parts = Split("0645 0639 0627 0644 062C 0629 0020 0628 064A 0627 0646 0627 062A 0020 0627 0644 0641 0635 0644 0020 0030 0032", " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    OutputFolderName = s
End Function

' File name of the saved workbook (hex code points + ".xlsx")
Private Function ResultsFileName() As String
    Dim parts() As String, i As Long, s As String
    parts = Split("0646 062A 0627 0626 062C 0020 0627 0644 0641 0635 0644 0020 0627 0644 062B 0627 0646 064A", " ")
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then s = s & ChrW(CLng("&H" & parts(i)))
    Next i
    ResultsFileName = s & ".xlsx"
End Function

Private Function FolderExistsPath(ByVal p As String) As Boolean
    On Error Resume Next
    FolderExistsPath = ((GetAttr(p) And vbDirectory) <> 0)
    If Err.Number <> 0 Then FolderExistsPath = False: Err.Clear
End Function

Private Function EnsureFolder(ByVal p As String) As Boolean
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    If Not FolderExistsPath(p) Then
        On Error Resume Next
        MkDir p
        Err.Clear
        On Error GoTo 0
    End If
    EnsureFolder = FolderExistsPath(p)
End Function

' ==============================================================================
' MAIN ENTRY POINT (Collect button): saves the collected workbook as the second term results file
' ==============================================================================
Public Sub CollectTerm02Results()
    Dim wbResult As Workbook
    Dim copied As Long, skipped As Long, skippedList As String, report As String
    Dim srcPath As String, outDir As String, outPath As String, failMsg As String
    Dim oldCalc As Long, saved As Boolean

    If ThisWorkbook.Path = "" Then
        MsgBox "Please save this workbook before running the macro.", vbExclamation, "Error": Exit Sub
    End If

    If LCase(Left(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "This workbook is on OneDrive/SharePoint (URL path)." & vbCrLf & _
               "Move it to a local folder, or sync the folder, then run again.", vbExclamation, "Error": Exit Sub
    End If

    srcPath = ThisWorkbook.Path & "\" & SOURCE_FOLDER_NAME & "\"
    outDir = ThisWorkbook.Path & "\" & OutputFolderName()
    outPath = outDir & "\" & ResultsFileName()

    If Dir(srcPath, vbDirectory) = "" Then
        MsgBox SOURCE_FOLDER_NAME & " folder not found:" & vbCrLf & srcPath, vbExclamation, "Error": Exit Sub
    End If

    oldCalc = Application.Calculation

    Application.ScreenUpdating = False: Application.Calculation = xlCalculationManual
    Application.EnableEvents = False: Application.DisplayAlerts = False

    On Error GoTo Fail

    If BuildSourceWorkbook(srcPath, wbResult, copied, skipped, skippedList) Then
        If Not EnsureFolder(outDir) Then Err.Raise vbObjectError + 910, , "The output folder could not be created: " & outDir
        wbResult.SaveAs fileName:=outPath, FileFormat:=xlOpenXMLWorkbook
        saved = True
    End If
    GoTo CleanExit

Fail:
    failMsg = Err.Number & " : " & Err.Description
    Err.Clear

CleanExit:
    On Error Resume Next
    If Not wbResult Is Nothing Then wbResult.Close SaveChanges:=False
    Set wbResult = Nothing
    Application.ScreenUpdating = True: Application.Calculation = oldCalc
    Application.EnableEvents = True: Application.DisplayAlerts = True
    Err.Clear
    On Error GoTo 0

    report = "Process completed!" & vbCrLf & vbCrLf & _
             "--- Data Collection (directly from " & SOURCE_FOLDER_NAME & ") ---" & vbCrLf & _
             "Copied files : " & copied & vbCrLf & "Skipped files: " & skipped & vbCrLf & vbCrLf
    If saved Then
        report = report & "Output: " & outPath
    Else
        report = report & "No usable source file was found - nothing was saved."
    End If

    If failMsg <> "" Then report = report & vbCrLf & vbCrLf & "ERROR: " & failMsg
    If skipped > 0 Or (Not saved And skippedList <> "") Then report = report & vbCrLf & vbCrLf & "Skipped / empty workbooks:" & vbCrLf & skippedList

    ' Sync the Control Panel state after a successful collection
    ' (optional, late-bound: works even when the control panel module is not imported)
    If failMsg = "" And saved Then
        On Error Resume Next
        Application.Run "'" & ThisWorkbook.Name & "'!RefreshPanelState"
        Err.Clear
        On Error GoTo 0
    End If

    MsgBox report, IIf(failMsg = "" And saved, vbInformation, vbExclamation), "Second Term Collection Report"
End Sub

Private Sub PlaceHeaderTitle(ByVal wsSource As Worksheet, ByVal wsTarget As Worksheet)

    Dim rawText As String
    Dim partTitle As String
    Dim partGroup As String

    ' 1. Capture and clean the text from A5
    rawText = CleanText(CStr(wsSource.Range("A5").Value))

    ' 2. Split at the end of the school year
    SplitAtSchoolYear rawText, partTitle, partGroup


    ' 3. Write and format both merged ranges
    WriteMergedCell wsTarget.Range("B1:F1"), partTitle
    WriteMergedCell wsTarget.Range("H1:Q1"), partGroup

    ' Merged cells do not auto-fit, so set the row height explicitly
    wsTarget.Rows(1).RowHeight = HEADER_ROW_HEIGHT

End Sub

'----------------------------------------------------------
' Cleans text: line breaks, non-breaking spaces, control
' characters and repeated spaces.
'----------------------------------------------------------
Private Function CleanText(ByVal s As String) As String

    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, Chr(160), " ")                       ' non-breaking space
    s = Application.WorksheetFunction.Clean(s)          ' control characters
    CleanText = Application.WorksheetFunction.Trim(s)   ' extra spaces

End Function

'----------------------------------------------------------
' Splits text into:
'   part1 = everything up to the end of the school year
'   part2 = the remainder (class / group name)
' If no year is found, part1 = full text and part2 = "".
'----------------------------------------------------------
Private Sub SplitAtSchoolYear(ByVal txt As String, _
                              ByRef part1 As String, _
                              ByRef part2 As String)

    Dim rx As Object
    Dim m As Object
    Dim splitPos As Long
    Dim digit As String
    Dim seps As String

    part1 = txt
    part2 = ""
    If Len(txt) = 0 Then Exit Sub

    ' Western (0-9) and Arabic-Indic (U+0660-U+0669) digits
    digit = "[0-9\u0660-\u0669]"
    ' Separators: slash, backslash, hyphen, en dash, em dash
    seps = "[/\\\-\u2013\u2014]"

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True

    ' First choice: full school-year pattern (e.g. 2025-2026)
    rx.Pattern = digit & "{4}\s*" & seps & "\s*" & digit & "{4}"

    If Not rx.Test(txt) Then
        ' Fallback: a single 4-digit year
        rx.Pattern = digit & "{4}"
        If Not rx.Test(txt) Then Exit Sub
    End If

    Set m = rx.Execute(txt)
    splitPos = m(0).FirstIndex + m(0).Length

    part1 = Trim(Left(txt, splitPos))

    If splitPos < Len(txt) Then
        part2 = Trim(Mid(txt, splitPos + 1))

        ' Remove any separator left at the start of the group name
        Do While Len(part2) > 0
            If InStr(1, "-:/" & ChrW(&H2013) & ChrW(&H2014), Left(part2, 1)) = 0 Then Exit Do
            part2 = Trim(Mid(part2, 2))
        Loop
    End If

End Sub

'----------------------------------------------------------
' Merges the range, writes the text and applies formatting.
'----------------------------------------------------------
Private Sub WriteMergedCell(ByVal rng As Range, ByVal txt As String)

    With rng
        .UnMerge                                ' remove any previous merge
        .ClearContents                          ' clear leftovers
        .Merge                                  ' merge the range
        .Value = txt
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .ReadingOrder = xlRTL                   ' right-to-left reading order
        .Font.Name = GetAvailableFont(PREFERRED_FONT, FALLBACK_FONT)
        .Font.Size = HEADER_FONT_SIZE
        .Font.Bold = True
    End With

End Sub

'----------------------------------------------------------
' Returns the preferred font if installed, otherwise the fallback.
'----------------------------------------------------------
Private Function GetAvailableFont(ByVal wanted As String, _
                                  ByVal fallback As String) As String

    If IsFontInstalled(wanted) Then
        GetAvailableFont = wanted
    Else
        GetAvailableFont = fallback
    End If

End Function

'----------------------------------------------------------
' Checks the system and per-user font folders for a file whose
' name starts with the font name (e.g. Cairo-Regular.ttf).
'----------------------------------------------------------
Private Function IsFontInstalled(ByVal fontName As String) As Boolean

    Dim folders(1 To 2) As String
    Dim i As Long

    folders(1) = Environ("WINDIR") & "\Fonts\"
    folders(2) = Environ("LOCALAPPDATA") & "\Microsoft\Windows\Fonts\"

    On Error Resume Next
    For i = 1 To 2
        If Len(Dir(folders(i) & fontName & "*.*")) > 0 Then
            IsFontInstalled = True
            Exit Function
        End If
    Next i
    On Error GoTo 0

End Function
