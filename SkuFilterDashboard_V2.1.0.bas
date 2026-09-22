Attribute VB_Name = "CC_Dashboard_V2_1"

Option Explicit

Private gMultiSelect As MultiSelectHandler

Private Const DASH      As String = "Dashboard"
Private Const LISTS     As String = "_FilterLists"
Private Const ALL_TAG   As String = "(All)"
Private Const HDR_ROW   As Long = 5          ' header row of the parameter table
Private Const P_ROW1    As Long = 6          ' first parameter row
Private Const R_COL     As Long = 8          ' results start in column F
Private Const MAX_LIST  As Long = 5000       ' max distinct values per dropdown

    ' developed by DSU11425
    ' Software Version 2.1.0

Public Sub Build_CC_Dashboard()
    Dim wsP As Worksheet, wsL As Worksheet, wsD As Worksheet, wsL2 As Worksheet
    Dim dat As Variant, loc As Variant
    Dim nCols As Long, i As Long, r As Long, listCol As Long
    Dim vals() As String, n As Long
    Dim isNum As Boolean

    On Error GoTo EH
    Application.ScreenUpdating = False

    If Not ResolveDataSheets(wsP, wsL) Then
        MsgBox "Could not find the product data." & vbCrLf & vbCrLf & DescribeSheets(), vbExclamation
        GoTo CLEANUP
    End If

    dat = ReadUsed(wsP)
    If IsEmpty(dat) Then
        MsgBox "The product sheet is empty.", vbExclamation
        GoTo CLEANUP
    End If

    If Not wsL Is Nothing Then
        loc = ReadUsed(wsL)                       ' legacy: separate sku/location lookup sheet
    Else
        Dim locColB As Long, rr As Long
        locColB = FindLocationColumn(dat)          ' single sheet: location_name is a product column
        If locColB > 0 Then
            ReDim loc(1 To UBound(dat, 1), 1 To 2)
            For rr = 2 To UBound(dat, 1)
                loc(rr, 2) = dat(rr, locColB)
            Next rr
        End If
    End If

    KillSheet DASH
    KillSheet LISTS

    Set wsL2 = TB.Worksheets.Add(After:=TB.Sheets(TB.Sheets.Count))
    wsL2.Name = LISTS
    Set wsD = TB.Worksheets.Add(After:=TB.Sheets(TB.Sheets.Count))
    wsD.Name = DASH

    nCols = UBound(dat, 2)

    '--- titles -------------------------------------------------
    With wsD
        .Range("A1").Value = "CC BUILD DASHBOARD"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Cells(HDR_ROW, 1).Value = "PARAMETER"
        .Cells(HDR_ROW, 2).Value = "VALUE  /  MIN"
        .Cells(HDR_ROW, 3).Value = "MAX"
        .Cells(HDR_ROW, 4).Value = "TYPE"
        .Cells(HDR_ROW, 5).Value = "KEY"
        .Range(.Cells(HDR_ROW, 1), .Cells(HDR_ROW, 5)).Font.Bold = True
        .Range(.Cells(HDR_ROW, 1), .Cells(HDR_ROW, 5)).Interior.Color = RGB(221, 235, 247)
    End With

    r = P_ROW1
    listCol = 0

    '--- one row per product column -----------------------------
    For i = 1 To nCols
        isNum = ColumnIsNumeric(dat, i)
        wsD.Cells(r, 1).Value = CStr(dat(1, i))
        wsD.Cells(r, 4).Value = IIf(isNum, "Number", "Text")
        wsD.Cells(r, 5).Value = "COL:" & i

        If isNum Then
            wsD.Cells(r, 2).Value = ""
            wsD.Cells(r, 3).Value = ""
        Else
            DistinctFromData dat, i, vals, n
            listCol = listCol + 1
            WriteList wsL2, listCol, vals, n
            AttachValidation wsD.Cells(r, 2), wsL2, listCol, n + 1
            wsD.Cells(r, 2).Value = ALL_TAG
        End If
        r = r + 1
    Next i

    '--- location parameters ------------------------------------
    AddLocParam wsD, wsL2, loc, r, listCol, 0, "LOCATION (full)", "LOC:F"
    AddLocParam wsD, wsL2, loc, r, listCol, 1, "LOCATION part 1 (XXX)", "LOC:1"
    AddLocParam wsD, wsL2, loc, r, listCol, 2, "LOCATION part 2 (YY)", "LOC:2"
    AddLocParam wsD, wsL2, loc, r, listCol, 3, "LOCATION part 3 (ZZZ)", "LOC:3"

    ' has-location yes/no
    wsD.Cells(r, 1).Value = "HAS LOCATION?"
    wsD.Cells(r, 4).Value = "Text"
    wsD.Cells(r, 5).Value = "LOC:H"
    ReDim vals(1 To 2)
    vals(1) = "Yes": vals(2) = "No"
    listCol = listCol + 1
    WriteList wsL2, listCol, vals, 2
    AttachValidation wsD.Cells(r, 2), wsL2, listCol, 3
    wsD.Cells(r, 2).Value = ALL_TAG
    r = r + 1

    '--- cosmetics ----------------------------------------------
    With wsD
        .Range(.Cells(P_ROW1, 1), .Cells(r - 1, 4)).Borders.LineStyle = xlContinuous
        .Range(.Cells(P_ROW1, 1), .Cells(r - 1, 4)).Borders.Color = RGB(190, 190, 190)
        .Range(.Cells(P_ROW1, 2), .Cells(r - 1, 3)).Interior.Color = RGB(255, 251, 224)
        .Columns("A").ColumnWidth = 22
        .Columns("B").ColumnWidth = 45
        .Columns("C").ColumnWidth = 8
        .Columns("D").ColumnWidth = 8
        .Columns("E").Hidden = True
        .Columns("F").ColumnWidth = 5
        .Columns("G").ColumnWidth = 5
        .Columns("H").ColumnWidth = 14
        .Columns("I").ColumnWidth = 50
        .Columns("J").ColumnWidth = 25
        .Columns("K").ColumnWidth = 7
        .Cells(HDR_ROW, R_COL).Value = "SKU"
        .Cells(HDR_ROW, R_COL + 1).Value = "PRODUCT NAME"
        .Cells(HDR_ROW, R_COL + 2).Value = "LOCATION"
        .Cells(HDR_ROW, R_COL + 3).Value = "STOCK"
        .Range(.Cells(HDR_ROW, R_COL), .Cells(HDR_ROW, R_COL + 3)).Font.Bold = True
        .Range(.Cells(HDR_ROW, R_COL), .Cells(HDR_ROW, R_COL + 3)).Interior.Color = RGB(221, 235, 247)
        .Cells(3, R_COL).Value = "Result: not run yet"
        .Cells(3, R_COL).Font.Bold = True
        .Rows(HDR_ROW + 1).Select
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(P_ROW1, 1).Select
    End With

    AddButton wsD, 10, 20, 130, 30, "APPLY FILTER", "ApplyFilters"
    AddButton wsD, 150, 20, 70, 30, "RESET", "ResetFilters"
    AddButton wsD, 230, 20, 100, 30, "REBUILD", "BuildDashboard"
    AddButton wsD, 340, 20, 100, 30, "Q-VEG", "QuickFilter_QVEG"

    wsL2.Visible = xlSheetVeryHidden
    wsD.Cells(1, 1).Select
    EnableMultiSelect

CLEANUP:
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "BuildDashboard failed: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------
' 2. APPLY THE FILTERS
'------------------------------------------------------------------
Public Sub ApplyFilters()
    Dim wsP As Worksheet, wsL As Worksheet, wsD As Worksheet
    Dim dat As Variant, loc As Variant
    Dim dLoc As Object
    Dim lastP As Long, i As Long, rw As Long, c As Long
    Dim key As String, typ As String, crit As String, critMax As String
    Dim ok As Boolean, hit As Long
    Dim locTxt As String

    On Error GoTo EH
    Set wsD = SheetOrNothing(DASH)
    If wsD Is Nothing Then MsgBox "Run BuildDashboard first.", vbExclamation: Exit Sub
    If Not ResolveDataSheets(wsP, wsL) Then
        MsgBox "Product data not found." & vbCrLf & vbCrLf & DescribeSheets(), vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    dat = ReadUsed(wsP)

    Dim locCol As Long, stockCol As Long
    locCol = 0
    If Not wsL Is Nothing Then
        loc = ReadUsed(wsL)
        Set dLoc = BuildLocMap(loc)
    Else
        locCol = FindLocationColumn(dat)
    End If
    stockCol = FindStockCol(dat)

    lastP = wsD.Cells(wsD.Rows.Count, 5).End(xlUp).Row

    ' Group by SKU: a SKU can appear as several batch rows (same SKU,
    ' same or different location, different expiry) — matching batches
    ' get combined into one output row: stock summed, distinct locations
    ' joined with ", ".
    Dim skuDict As Object, idx As Long, skuKey As String
    Dim outSku() As Variant, outName() As Variant, stockSum() As Double
    Dim locSets() As Object
    Set skuDict = CreateObject("Scripting.Dictionary")
    ReDim outSku(1 To UBound(dat, 1))
    ReDim outName(1 To UBound(dat, 1))
    ReDim stockSum(1 To UBound(dat, 1))
    ReDim locSets(1 To UBound(dat, 1))
    hit = 0

    For rw = 2 To UBound(dat, 1)
        ok = True
        locTxt = ""
        If Not wsL Is Nothing Then
            If dLoc.Exists(KeyOf(dat(rw, 1))) Then locTxt = dLoc(KeyOf(dat(rw, 1)))
        ElseIf locCol > 0 Then
            locTxt = NormalizeLoc(AsText(dat(rw, locCol)))
        End If

        For i = P_ROW1 To lastP
            key = CStr(wsD.Cells(i, 5).Value)
            If key <> "" Then
                typ = CStr(wsD.Cells(i, 4).Value)
                crit = CStr(wsD.Cells(i, 2).Value)
                critMax = CStr(wsD.Cells(i, 3).Value)

                If Left$(key, 4) = "COL:" Then
                    c = CLng(Mid$(key, 5))
                    If typ = "Number" Then
                        If Not NumMatch(dat(rw, c), crit, critMax) Then ok = False
                    Else
                        If Not TextMatch(AsText(dat(rw, c)), crit) Then ok = False
                    End If
                Else
                    Select Case key
                        Case "LOC:F"
                            If Not TextMatch(locTxt, NormalizeCrit(crit, 0)) Then ok = False
                        Case "LOC:1", "LOC:2", "LOC:3"
                            If Not TextMatch(LocPart(locTxt, CLng(Mid$(key, 5))), _
                                             NormalizeCrit(crit, CLng(Mid$(key, 5)))) Then ok = False
                        Case "LOC:H"
                            If Trim$(crit) <> "" And crit <> ALL_TAG Then
                                If StrComp(crit, "Yes", vbTextCompare) = 0 And locTxt = "" Then ok = False
                                If StrComp(crit, "No", vbTextCompare) = 0 And locTxt <> "" Then ok = False
                            End If
                    End Select
                End If
            End If
            If Not ok Then Exit For
        Next i

        If ok Then
            skuKey = KeyOf(dat(rw, 1))
            If Not skuDict.Exists(skuKey) Then
                hit = hit + 1
                skuDict.Add skuKey, hit
                outSku(hit) = dat(rw, 1)
                outName(hit) = dat(rw, 2)
                stockSum(hit) = 0
                Set locSets(hit) = CreateObject("Scripting.Dictionary")
            End If
            idx = skuDict(skuKey)

            If locTxt <> "" Then
                If Not locSets(idx).Exists(locTxt) Then locSets(idx).Add locTxt, 1
            End If
            If stockCol > 0 Then
                If IsNumeric(dat(rw, stockCol)) Then stockSum(idx) = stockSum(idx) + CDbl(dat(rw, stockCol))
            End If
        End If
    Next rw

    ' clear old results
    wsD.Range(wsD.Cells(HDR_ROW + 1, R_COL), wsD.Cells(wsD.Rows.Count, R_COL + 3)).Clear

    If hit > 0 Then
        Dim out() As Variant, k As Long, locStr As String, kk As Variant
        ReDim out(1 To hit, 1 To 4)
        For k = 1 To hit
            out(k, 1) = outSku(k)
            out(k, 2) = outName(k)
            locStr = ""
            For Each kk In locSets(k).Keys
                locStr = locStr & IIf(locStr = "", "", ", ") & kk
            Next kk
            out(k, 3) = locStr
            If stockCol > 0 Then out(k, 4) = stockSum(k) Else out(k, 4) = ""
        Next k

        wsD.Cells(HDR_ROW + 1, R_COL).Resize(hit, 4).Value = out

        ' sort the results A-Z by LOCATION (blanks sort first)
        If hit > 1 Then
            wsD.Range(wsD.Cells(HDR_ROW + 1, R_COL), wsD.Cells(HDR_ROW + hit, R_COL + 3)).Sort _
                Key1:=wsD.Cells(HDR_ROW + 1, R_COL + 2), Order1:=xlAscending, _
                Header:=xlNo, MatchCase:=False, Orientation:=xlTopToBottom
        End If

        If stockCol = 0 Then wsD.Cells(HDR_ROW, R_COL + 3).Value = "STOCK (col n/a)"

        With wsD.Range(wsD.Cells(HDR_ROW + 1, R_COL), wsD.Cells(HDR_ROW + hit, R_COL + 3))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(210, 210, 210)
        End With
    End If

    wsD.Cells(3, R_COL).Value = "Result: " & hit & " SKU(s)  (" & (UBound(dat, 1) - 1) & " data rows scanned)"
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "ApplyFilters failed: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------
' 3. RESET
'------------------------------------------------------------------
Public Sub ResetFilters()
    Dim wsD As Worksheet, i As Long, lastP As Long
    Set wsD = SheetOrNothing(DASH)
    If wsD Is Nothing Then Exit Sub
    lastP = wsD.Cells(wsD.Rows.Count, 5).End(xlUp).Row
    For i = P_ROW1 To lastP
        If CStr(wsD.Cells(i, 4).Value) = "Number" Then
            wsD.Cells(i, 2).ClearContents
            wsD.Cells(i, 3).ClearContents
        Else
            wsD.Cells(i, 2).Value = ALL_TAG
        End If
    Next i
    wsD.Range(wsD.Cells(HDR_ROW + 1, R_COL), wsD.Cells(wsD.Rows.Count, R_COL + 3)).Clear
    wsD.Cells(3, R_COL).Value = "Result: not run yet"
End Sub

'==================================================================
' HELPERS
'==================================================================
' Shows exactly which workbook and sheets were checked when data can't be
' found — so a mismatch (wrong workbook, stale module, unexpected sheet
' names) is visible immediately instead of guessed at.
Private Function DescribeSheets() As String
    Dim wb As Workbook, ws As Worksheet, s As String
    On Error Resume Next
    Set wb = TB()
    If wb Is Nothing Then DescribeSheets = "(no active workbook detected)": Exit Function
    s = "Looking in workbook: " & wb.Name & vbCrLf & "Sheets seen: "
    For Each ws In wb.Worksheets
        s = s & ws.Name & " | "
    Next ws
    DescribeSheets = s
End Function

Private Function ResolveDataSheets(ByRef wsP As Worksheet, ByRef wsL As Worksheet) As Boolean
    Dim ws As Worksheet, col As New Collection, i As Long
    Dim h1 As String, h2 As String

    Set wsL = Nothing
    For Each ws In TB.Worksheets
        If ws.Name <> DASH And ws.Name <> LISTS Then col.Add ws
    Next ws
    If col.Count = 0 Then Exit Function

    ' single sheet (e.g. a CSV opened in Excel): location lives in the
    ' product list itself, there is no separate lookup sheet to find
    If col.Count = 1 Then
        Set wsP = col(1)
        ResolveDataSheets = True
        Exit Function
    End If

    ' the location sheet is the one whose headers mention "location"
    For i = 1 To col.Count
        Set ws = col(i)
        h1 = LCase$(CStr(ws.Cells(1, 1).Value) & "|" & CStr(ws.Cells(1, 2).Value))
        If InStr(h1, "location") > 0 Then
            Set wsL = ws
            Exit For
        End If
    Next i

    If wsL Is Nothing Then
        Set wsP = col(1): Set wsL = col(2)
    Else
        For i = 1 To col.Count
            If Not col(i) Is wsL Then Set wsP = col(i): Exit For
        Next i
    End If
    ResolveDataSheets = Not wsP Is Nothing
End Function

' Locates the location column when it's embedded directly in the product
' list (single-sheet workbooks) rather than a separate lookup sheet.
Private Function FindLocationColumn(dat As Variant) As Long
    Dim c As Long, h As String
    For c = 1 To UBound(dat, 2)
        h = LCase$(Trim$(Replace(CStr(dat(1, c)), " ", "_")))
        If h = "location_name" Or h = "location" Then FindLocationColumn = c: Exit Function
    Next c
    For c = 1 To UBound(dat, 2)
        h = LCase$(Trim$(Replace(CStr(dat(1, c)), " ", "_")))
        If InStr(h, "location") > 0 Then FindLocationColumn = c: Exit Function
    Next c
End Function

Private Function ReadUsed(ws As Worksheet) As Variant
    Dim lr As Long, lc As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lr < 1 Or lc < 1 Then Exit Function
    If lr = 1 And lc = 1 And IsEmpty(ws.Cells(1, 1)) Then Exit Function
    ReadUsed = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value
End Function

' Locates the stock-on-hand column by header text so it still works if the
' product list's column order shifts between exports. Prefers an exact
' header match, then falls back to any header that says "stock" without
' also being a reserved/transit/blocked/buffer bucket.
Private Function FindStockCol(dat As Variant) As Long
    Dim c As Long, h As String
    For c = 1 To UBound(dat, 2)
        h = LCase$(Trim$(CStr(dat(1, c))))
        If h = "stock_on_hand" Or h = "stock on hand" Or h = "stock" Then
            FindStockCol = c
            Exit Function
        End If
    Next c
    For c = 1 To UBound(dat, 2)
        h = LCase$(Trim$(CStr(dat(1, c))))
        If InStr(h, "stock") > 0 And InStr(h, "reserved") = 0 And InStr(h, "transit") = 0 _
           And InStr(h, "blocked") = 0 And InStr(h, "buffer") = 0 And InStr(h, "putaway") = 0 Then
            FindStockCol = c
            Exit Function
        End If
    Next c
End Function

Private Function BuildLocMap(loc As Variant) As Object
    Dim d As Object, i As Long, k As String, v As String
    Set d = CreateObject("Scripting.Dictionary")
    If IsEmpty(loc) Then Set BuildLocMap = d: Exit Function
    For i = 2 To UBound(loc, 1)
        k = KeyOf(loc(i, 1))
        If UBound(loc, 2) >= 2 Then v = NormalizeLoc(AsText(loc(i, 2))) Else v = ""
        If k <> "" Then
            If Not d.Exists(k) Then
                d.Add k, v
            ElseIf d(k) = "" And v <> "" Then
                d(k) = v
            ElseIf v <> "" And InStr(1, d(k) & ";", v & ";", vbTextCompare) = 0 Then
                d(k) = d(k) & "; " & v          ' SKU stored in several bins
            End If
        End If
    Next i
    Set BuildLocMap = d
End Function

Private Function KeyOf(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then
        KeyOf = Format$(CDbl(v), "0.##########")
    Else
        KeyOf = Trim$(CStr(v))
    End If
End Function

Private Function AsText(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        AsText = IIf(v, "TRUE", "FALSE")
    ElseIf IsNumeric(v) Then
        AsText = Format$(CDbl(v), "0.##########")
    ElseIf VarType(v) = vbDate Then
        AsText = Format$(v, "yyyy-mm-dd")
    Else
        AsText = Trim$(CStr(v))
    End If
End Function

Private Function LocPart(ByVal locTxt As String, ByVal idx As Long) As String
    Dim p() As String
    If locTxt = "" Then Exit Function
    p = Split(Split(locTxt, ";")(0), "-")
    If idx - 1 <= UBound(p) Then LocPart = Trim$(p(idx - 1))
End Function

'--- location is always XXX-YY-ZZZ : 8 characters of data --------
' F01-8-5  ->  F01-08-005        f1-2-7  ->  F01-02-007
Private Function NormalizeLoc(ByVal s As String) As String
    Dim p() As String
    s = UCase$(Replace(Replace(Trim$(s), " ", ""), "_", "-"))
    If s = "" Then Exit Function
    p = Split(s, "-")
    If UBound(p) <> 2 Then NormalizeLoc = s: Exit Function
    NormalizeLoc = PadPart(p(0), 3) & "-" & PadPart(p(1), 2) & "-" & PadPart(p(2), 3)
End Function

' pads a section with leading zeros up to n characters
Private Function PadPart(ByVal s As String, ByVal n As Long) As String
    Dim i As Long, digitsOnly As Boolean
    s = Trim$(s)
    If s = "" Then PadPart = String$(n, "0"): Exit Function
    digitsOnly = True
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then digitsOnly = False: Exit For
    Next i
    If digitsOnly Then
        If Len(s) < n Then
            PadPart = String$(n - Len(s), "0") & s
        Else
            PadPart = s
        End If
    Else
        ' e.g. "F1" in a 3-char section -> "F01"  (zeros inserted after the letters)
        i = 1
        Do While i <= Len(s) And Not (Mid$(s, i, 1) >= "0" And Mid$(s, i, 1) <= "9")
            i = i + 1
        Loop
        If i <= Len(s) And Len(s) < n Then
            PadPart = Left$(s, i - 1) & String$(n - Len(s), "0") & Mid$(s, i)
        Else
            PadPart = s
        End If
    End If
End Function

' normalizes what the user typed, leaving wildcard entries untouched
Private Function NormalizeCrit(ByVal crit As String, ByVal part As Long) As String
    Dim p() As String, i As Long, s As String, res As String
    crit = Trim$(crit)
    If crit = "" Or crit = ALL_TAG Then NormalizeCrit = crit: Exit Function
    p = Split(crit, ";")
    For i = LBound(p) To UBound(p)
        s = Trim$(p(i))
        If s <> "" Then
            If InStr(s, "*") = 0 And InStr(s, "?") = 0 Then
                Select Case part
                    Case 0: s = NormalizeLoc(s)
                    Case 1: s = PadPart(UCase$(s), 3)
                    Case 2: s = PadPart(UCase$(s), 2)
                    Case 3: s = PadPart(UCase$(s), 3)
                End Select
            Else
                s = UCase$(s)
            End If
            res = res & IIf(res = "", "", ";") & s
        End If
    Next i
    NormalizeCrit = res
End Function

Private Function TextMatch(ByVal val As String, ByVal crit As String) As Boolean
    Dim parts() As String, i As Long, p As String
    crit = Trim$(crit)
    If crit = "" Or crit = ALL_TAG Then TextMatch = True: Exit Function
    ' whole criterion as a literal first (values may themselves contain ";")
    If StrComp(val, crit, vbTextCompare) = 0 Then TextMatch = True: Exit Function
    parts = Split(crit, ";")
    For i = LBound(parts) To UBound(parts)
        p = Trim$(parts(i))
        If p <> "" Then
            If InStr(p, "*") > 0 Or InStr(p, "?") > 0 Then
                If LCase$(val) Like LCase$(p) Then TextMatch = True: Exit Function
            ElseIf StrComp(val, p, vbTextCompare) = 0 Then
                TextMatch = True: Exit Function
            End If
        End If
    Next i
End Function

Private Function NumMatch(ByVal v As Variant, ByVal lo As String, ByVal hi As String) As Boolean
    Dim d As Double
    lo = Trim$(lo): hi = Trim$(hi)
    If lo = "" And hi = "" Then NumMatch = True: Exit Function
    If Not IsNumeric(v) Then Exit Function
    d = CDbl(v)
    If lo <> "" Then
        If Not IsNumeric(lo) Then Exit Function
        If d < CDbl(lo) Then Exit Function
    End If
    If hi <> "" Then
        If Not IsNumeric(hi) Then Exit Function
        If d > CDbl(hi) Then Exit Function
    End If
    NumMatch = True
End Function

Private Function ColumnIsNumeric(dat As Variant, ByVal c As Long) As Boolean
    Dim i As Long, seen As Long, num As Long, d As Object
    Set d = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(dat, 1)
        If Not IsEmpty(dat(i, c)) And Not IsError(dat(i, c)) Then
            seen = seen + 1
            If IsNumeric(dat(i, c)) And VarType(dat(i, c)) <> vbBoolean Then num = num + 1
            If Not d.Exists(AsText(dat(i, c))) Then d.Add AsText(dat(i, c)), 1
        End If
    Next i
    ' numeric only when every value is a number AND it is not a short code list
    ColumnIsNumeric = (seen > 0 And num = seen And d.Count > 12 And c <> 1)
End Function

Private Sub DistinctFromData(dat As Variant, ByVal c As Long, ByRef vals() As String, ByRef n As Long)
    Dim d As Object, i As Long, s As String
    Set d = CreateObject("Scripting.Dictionary")
    For i = 2 To UBound(dat, 1)
        s = AsText(dat(i, c))
        If s = "" Then s = "(blank)"
        If Not d.Exists(s) Then
            If d.Count >= MAX_LIST Then Exit For
            d.Add s, 1
        End If
    Next i
    n = d.Count
    If n = 0 Then ReDim vals(1 To 1): vals(1) = "(blank)": n = 1: Exit Sub
    ReDim vals(1 To n)
    For i = 0 To n - 1
        vals(i + 1) = d.Keys()(i)
    Next i
    SortStrings vals, 1, n
End Sub

Private Sub AddLocParam(wsD As Worksheet, wsL2 As Worksheet, loc As Variant, _
                        ByRef r As Long, ByRef listCol As Long, _
                        ByVal part As Long, ByVal caption As String, ByVal key As String)
    Dim d As Object, i As Long, s As String, vals() As String, n As Long
    Set d = CreateObject("Scripting.Dictionary")

    If Not IsEmpty(loc) Then
        For i = 2 To UBound(loc, 1)
            If UBound(loc, 2) >= 2 Then
                s = NormalizeLoc(AsText(loc(i, 2)))
                If s <> "" Then
                    If part > 0 Then s = LocPart(s, part)
                    If s <> "" Then If Not d.Exists(s) Then If d.Count < MAX_LIST Then d.Add s, 1
                End If
            End If
        Next i
    End If

    n = d.Count
    If n = 0 Then
        ReDim vals(1 To 1): vals(1) = "(none)": n = 1
    Else
        ReDim vals(1 To n)
        For i = 0 To n - 1
            vals(i + 1) = d.Keys()(i)
        Next i
        SortStrings vals, 1, n
    End If

    wsD.Cells(r, 1).Value = caption
    wsD.Cells(r, 4).Value = "Text"
    wsD.Cells(r, 5).Value = key
    listCol = listCol + 1
    WriteList wsL2, listCol, vals, n
    AttachValidation wsD.Cells(r, 2), wsL2, listCol, n + 1
    wsD.Cells(r, 2).Value = ALL_TAG
    r = r + 1
End Sub

Private Sub WriteList(wsL2 As Worksheet, ByVal listCol As Long, vals() As String, ByVal n As Long)
    Dim arr() As Variant, i As Long
    ReDim arr(1 To n + 1, 1 To 1)
    arr(1, 1) = ALL_TAG
    For i = 1 To n
        arr(i + 1, 1) = vals(i)
    Next i
    wsL2.Cells(1, listCol).Resize(n + 1, 1).Value = arr
End Sub

Private Sub AttachValidation(rng As Range, wsL2 As Worksheet, ByVal listCol As Long, ByVal n As Long)
    Dim f As String
    f = "=" & LISTS & "!" & wsL2.Cells(1, listCol).Resize(n, 1).Address(True, True)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertInformation, Formula1:=f
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    rng.Validation.ShowError = False        ' allow typed values / wildcards
    On Error GoTo 0
End Sub

Private Sub SortStrings(a() As String, ByVal lo As Long, ByVal hi As Long)
    Dim i As Long, j As Long, p As String, t As String
    i = lo: j = hi: p = a((lo + hi) \ 2)
    Do While i <= j
        Do While StrComp(a(i), p, vbTextCompare) < 0: i = i + 1: Loop
        Do While StrComp(a(j), p, vbTextCompare) > 0: j = j - 1: Loop
        If i <= j Then
            t = a(i): a(i) = a(j): a(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then SortStrings a, lo, j
    If i < hi Then SortStrings a, i, hi
End Sub

Private Sub AddButton(ws As Worksheet, ByVal l As Single, ByVal t As Single, _
                      ByVal w As Single, ByVal h As Single, _
                      ByVal caption As String, ByVal macro As String)
    Dim shp As Shape
    Set shp = ws.Shapes.AddFormControl(xlButtonControl, l, t, w, h)
    shp.OnAction = ThisWorkbook.Name & "!" & macro
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.Characters.Font.Bold = True
End Sub

' The workbook the dashboard is built in.
' Works whether this module sits in the data workbook or in PERSONAL.XLSB:
' ThisWorkbook = the file holding the code, TB = the file holding the data.
Private Function TB() As Workbook
    If ActiveWorkbook Is Nothing Then
        Set TB = ThisWorkbook
    ElseIf ActiveWorkbook Is ThisWorkbook Then
        Set TB = ThisWorkbook
    Else
        Set TB = ActiveWorkbook
    End If
End Function

Private Function SheetOrNothing(ByVal nm As String) As Worksheet
    On Error Resume Next
    Set SheetOrNothing = TB.Worksheets(nm)
    On Error GoTo 0
End Function

' Hooks the multi-select class to the running Excel instance so every
' dropdown pick on the Dashboard sheet is intercepted and merged instead
' of overwriting the cell. Safe to call repeatedly.
Public Sub EnableMultiSelect()
    If gMultiSelect Is Nothing Then Set gMultiSelect = New MultiSelectHandler
    Set gMultiSelect.App = Application
End Sub

'==================================================================
' CASCADING CATEGORY FILTERS  (parent_category -> category -> subcategory)
'------------------------------------------------------------------
' Selecting a value in a higher level narrows what the lower levels'
' dropdowns can show — the same way Excel's own column filters only
' offer values still present in the currently-filtered rows.
' Called automatically by MultiSelectHandler whenever a dropdown
' cell changes; never needs to be run manually.
'==================================================================

' Finds the dashboard rows for parent_category / category / subcategory
' by header text (0 = that field isn't in this product list at all).
Private Sub FindHierarchyRows(wsD As Worksheet, ByRef rParent As Long, ByRef rCategory As Long, ByRef rSub As Long)
    Dim lastP As Long, i As Long, h As String
    rParent = 0: rCategory = 0: rSub = 0
    lastP = wsD.Cells(wsD.Rows.Count, 5).End(xlUp).Row
    For i = P_ROW1 To lastP
        h = LCase$(Trim$(Replace(CStr(wsD.Cells(i, 1).Value), " ", "_")))
        Select Case h
            Case "parent_category": rParent = i
            Case "category": rCategory = i
            Case "subcategory": rSub = i
        End Select
    Next i
End Sub

' Finds a product-list column by header text (0 if not present).
Private Function ColOfHeader(dat As Variant, ByVal name As String) As Long
    Dim c As Long, h As String
    name = LCase$(Trim$(Replace(name, " ", "_")))
    For c = 1 To UBound(dat, 2)
        h = LCase$(Trim$(Replace(CStr(dat(1, c)), " ", "_")))
        If h = name Then ColOfHeader = c: Exit Function
    Next c
End Function

' Distinct values of targetCol among rows that match every constraint
' column/selection pair (constraints use the same "(All)" / ";" / "*"
' rules as the main filter, via TextMatch).
Private Sub ComputeConstrainedList(dat As Variant, ByVal targetCol As Long, _
                                   constraintCols As Variant, constraintSels As Variant, _
                                   ByRef vals() As String, ByRef n As Long)
    Dim d As Object, r As Long, i As Long, ok As Boolean, s As String
    Set d = CreateObject("Scripting.Dictionary")
    If targetCol = 0 Then n = 0: Exit Sub

    For r = 2 To UBound(dat, 1)
        ok = True
        For i = LBound(constraintCols) To UBound(constraintCols)
            If constraintCols(i) > 0 Then
                If Not TextMatch(AsText(dat(r, constraintCols(i))), CStr(constraintSels(i))) Then
                    ok = False
                    Exit For
                End If
            End If
        Next i
        If ok Then
            s = AsText(dat(r, targetCol))
            If s = "" Then s = "(blank)"
            If Not d.Exists(s) Then If d.Count < MAX_LIST Then d.Add s, 1
        End If
    Next r

    n = d.Count
    If n > 0 Then
        ReDim vals(1 To n)
        For i = 0 To n - 1
            vals(i + 1) = d.Keys()(i)
        Next i
        SortStrings vals, 1, n
    End If
End Sub

' Rewrites a dropdown's own source list in place — reading the range the
' cell's validation already points at (so no need to resize the
' validation itself), writing "(All)" + the narrowed values at the top,
' and blanking the rest. Excel's list validation stops at the first
' blank cell in its source range, so this is enough to hide values that
' no longer apply, exactly like Excel's own filter narrowing.
Private Sub RewriteListInPlace(cell As Range, vals() As String, ByVal n As Long)
    Dim f As String, parts() As String, sh As String, addr As String
    Dim rng As Range, cap As Long, outArr() As Variant, i As Long

    On Error GoTo EH
    f = cell.Validation.Formula1
    If Left$(f, 1) = "=" Then f = Mid$(f, 2)
    parts = Split(f, "!")
    If UBound(parts) <> 1 Then Exit Sub
    sh = Replace(parts(0), "'", "")
    addr = parts(1)
    Set rng = TB().Sheets(sh).Range(addr)

    cap = rng.Rows.Count
    If n + 1 > cap Then n = cap - 1      ' safety clamp; shouldn't normally trigger

    ReDim outArr(1 To cap, 1 To 1)
    outArr(1, 1) = ALL_TAG
    For i = 1 To n
        outArr(i + 1, 1) = vals(i)
    Next i
    rng.Value = outArr
    Exit Sub
EH:
    ' cell has no recognizable list validation to rewrite — nothing to do
End Sub

' Recomputes the Category and/or Sub Category lists to match whatever is
' currently selected in the levels above them.
Private Sub RefreshCascade(ByVal rParent As Long, ByVal rCategory As Long, ByVal rSub As Long, _
                            ByVal doCategory As Boolean, ByVal doSub As Boolean)
    Dim wsD As Worksheet, wsP As Worksheet, wsL As Worksheet, dat As Variant
    Dim colParent As Long, colCategory As Long, colSub As Long
    Dim selParent As String, selCategory As String
    Dim catVals() As String, catN As Long, subVals() As String, subN As Long

    Set wsD = SheetOrNothing(DASH)
    If wsD Is Nothing Then Exit Sub
    If Not ResolveDataSheets(wsP, wsL) Then Exit Sub
    dat = ReadUsed(wsP)
    If IsEmpty(dat) Then Exit Sub

    colParent = ColOfHeader(dat, "parent_category")
    colCategory = ColOfHeader(dat, "category")
    colSub = ColOfHeader(dat, "subcategory")

    If rParent > 0 Then selParent = CStr(wsD.Cells(rParent, 2).Value)
    If rCategory > 0 Then selCategory = CStr(wsD.Cells(rCategory, 2).Value)

    If doCategory And rCategory > 0 And colCategory > 0 Then
        ComputeConstrainedList dat, colCategory, Array(colParent), Array(selParent), catVals, catN
        RewriteListInPlace wsD.Cells(rCategory, 2), catVals, catN
    End If

    If doSub And rSub > 0 And colSub > 0 Then
        ComputeConstrainedList dat, colSub, Array(colParent, colCategory), Array(selParent, selCategory), subVals, subN
        RewriteListInPlace wsD.Cells(rSub, 2), subVals, subN
    End If
End Sub

' Called by MultiSelectHandler right after any dropdown cell settles on
' its new value. Cheap no-op for every row except Parent Category and
' Category — those two cascade into the level(s) below them.
Public Sub OnParamChanged(ByVal changedRow As Long)
    Dim wsD As Worksheet, rParent As Long, rCategory As Long, rSub As Long
    Set wsD = SheetOrNothing(DASH)
    If wsD Is Nothing Then Exit Sub
    FindHierarchyRows wsD, rParent, rCategory, rSub

    If rParent > 0 And changedRow = rParent Then
        RefreshCascade rParent, rCategory, rSub, True, True
    ElseIf rCategory > 0 And changedRow = rCategory Then
        RefreshCascade rParent, rCategory, rSub, False, True
    End If
End Sub

'==================================================================
' QUICK FILTER BUTTONS
'==================================================================

' Q-VEG: Parent Category = Fruit & Veg + Ready To Eat,
'        Category = every category under those two EXCEPT Dates & Dried Fruit.
Public Sub QuickFilter_QVEG()
    Dim wsD As Worksheet, wsP As Worksheet, wsL As Worksheet, dat As Variant
    Dim rParent As Long, rCategory As Long, rSub As Long
    Dim colParent As Long, colCategory As Long
    Dim catVals() As String, catN As Long, i As Long, res As String
    Const PARENT_SEL As String = "Fruit & Veg; Ready To Eat"
    Const EXCLUDE_CAT As String = "Dates & Dried Fruit"

    Set wsD = SheetOrNothing(DASH)
    If wsD Is Nothing Then MsgBox "Run BuildDashboard first.", vbExclamation: Exit Sub
    If Not ResolveDataSheets(wsP, wsL) Then
        MsgBox "Product data not found." & vbCrLf & vbCrLf & DescribeSheets(), vbExclamation
        Exit Sub
    End If
    dat = ReadUsed(wsP)
    If IsEmpty(dat) Then Exit Sub

    FindHierarchyRows wsD, rParent, rCategory, rSub
    If rParent = 0 Or rCategory = 0 Then
        MsgBox "This product list has no parent_category / category columns.", vbExclamation
        Exit Sub
    End If
    colParent = ColOfHeader(dat, "parent_category")
    colCategory = ColOfHeader(dat, "category")

    Application.EnableEvents = False
    wsD.Cells(rParent, 2).Value = PARENT_SEL

    ComputeConstrainedList dat, colCategory, Array(colParent), Array(PARENT_SEL), catVals, catN
    res = ""
    For i = 1 To catN
        If StrComp(catVals(i), EXCLUDE_CAT, vbTextCompare) <> 0 Then
            res = res & IIf(res = "", "", "; ") & catVals(i)
        End If
    Next i
    If res = "" Then res = ALL_TAG
    wsD.Cells(rCategory, 2).Value = res
    Application.EnableEvents = True

    RefreshCascade rParent, rCategory, rSub, True, True
    ApplyFilters
End Sub

Private Sub KillSheet(ByVal nm As String)
    Dim ws As Worksheet
    Set ws = SheetOrNothing(nm)
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Visible = xlSheetVisible
        ws.Delete
        Application.DisplayAlerts = True
    End If
End Sub
