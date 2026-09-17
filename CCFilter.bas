Attribute VB_Name = "CCDashboard"
'==================================================================
' SKU FILTER DASHBOARD
' developed by DSU11425
' Software Version 1.0.2
'------------------------------------------------------------------
Option Explicit

Private Const DASH      As String = "Dashboard"
Private Const LISTS     As String = "_FilterLists"
Private Const ALL_TAG   As String = "(All)"
Private Const HDR_ROW   As Long = 5          ' header row of the parameter table
Private Const P_ROW1    As Long = 6          ' first parameter row
Private Const R_COL     As Long = 6          ' results start in column F
Private Const MAX_LIST  As Long = 5000       ' max distinct values per dropdown

'------------------------------------------------------------------
Public Sub Build_CC_Dashboard()
    Dim wsP As Worksheet, wsL As Worksheet, wsD As Worksheet, wsL2 As Worksheet
    Dim dat As Variant, loc As Variant
    Dim nCols As Long, i As Long, r As Long, listCol As Long
    Dim vals() As String, n As Long
    Dim isNum As Boolean

    On Error GoTo EH
    Application.ScreenUpdating = False

    If Not ResolveDataSheets(wsP, wsL) Then
        MsgBox "Could not find the two data sheets." & vbCrLf & _
               "The workbook needs a product sheet and a location sheet.", vbExclamation
        GoTo CLEANUP
    End If

    dat = ReadUsed(wsP)
    If IsEmpty(dat) Then
        MsgBox "The product sheet is empty.", vbExclamation
        GoTo CLEANUP
    End If
    loc = ReadUsed(wsL)

    KillSheet DASH
    KillSheet LISTS

    Set wsL2 = TB.Worksheets.Add(After:=TB.Sheets(TB.Sheets.Count))
    wsL2.Name = LISTS
    Set wsD = TB.Worksheets.Add(After:=TB.Sheets(TB.Sheets.Count))
    wsD.Name = DASH

    nCols = UBound(dat, 2)

    '--- titles -------------------------------------------------
    With wsD
        .Range("A1").Value = "SKU FILTER DASHBOARD"
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A2").Value = "Set any parameters below, leave the rest as " & ALL_TAG & ", then press APPLY FILTER."
        .Range("A3").Value = "Text filters accept: one value, several separated by "";"" , or wildcards such as *choc*"
        .Range("A4").Value = "Locations are padded to XXX-YY-ZZZ (8 characters of data): typing F01-8-5 finds F01-08-005"

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
            AttachValidation wsD.Cells(r, 2), wsL2, listCol, n
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
        .Columns("A").ColumnWidth = 28
        .Columns("B").ColumnWidth = 30
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 9
        .Columns("E").Hidden = True
        .Columns("F").ColumnWidth = 14
        .Columns("G").ColumnWidth = 60
        .Columns("H").ColumnWidth = 16
        .Cells(HDR_ROW, R_COL).Value = "SKU"
        .Cells(HDR_ROW, R_COL + 1).Value = "PRODUCT NAME"
        .Cells(HDR_ROW, R_COL + 2).Value = "LOCATION"
        .Range(.Cells(HDR_ROW, R_COL), .Cells(HDR_ROW, R_COL + 2)).Font.Bold = True
        .Range(.Cells(HDR_ROW, R_COL), .Cells(HDR_ROW, R_COL + 2)).Interior.Color = RGB(221, 235, 247)
        .Cells(3, R_COL).Value = "Result: not run yet"
        .Cells(3, R_COL).Font.Bold = True
        .Rows(HDR_ROW + 1).Select
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(P_ROW1, 1).Select
    End With

    AddButton wsD, 10, 20, 130, 30, "APPLY FILTER", "ApplyFilters"
    AddButton wsD, 150, 20, 110, 30, "RESET", "ResetFilters"
    AddButton wsD, 268, 20, 140, 30, "REBUILD LISTS", "BuildDashboard"

    wsL2.Visible = xlSheetVeryHidden
    wsD.Cells(1, 1).Select

CLEANUP:
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "BuildDashboard failed: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------
Public Sub ApplyFilters()
    Dim wsP As Worksheet, wsL As Worksheet, wsD As Worksheet
    Dim dat As Variant, loc As Variant
    Dim dLoc As Object
    Dim lastP As Long, i As Long, rw As Long, c As Long
    Dim key As String, typ As String, crit As String, critMax As String
    Dim ok As Boolean, hit As Long
    Dim out() As Variant, locTxt As String, parts() As String

    On Error GoTo EH
    Set wsD = SheetOrNothing(DASH)
    If wsD Is Nothing Then MsgBox "Run BuildDashboard first.", vbExclamation: Exit Sub
    If Not ResolveDataSheets(wsP, wsL) Then MsgBox "Data sheets not found.", vbExclamation: Exit Sub

    Application.ScreenUpdating = False
    dat = ReadUsed(wsP)
    loc = ReadUsed(wsL)
    Set dLoc = BuildLocMap(loc)

    lastP = wsD.Cells(wsD.Rows.Count, 5).End(xlUp).Row

    ReDim out(1 To UBound(dat, 1), 1 To 3)
    hit = 0

    For rw = 2 To UBound(dat, 1)
        ok = True
        locTxt = ""
        If dLoc.Exists(KeyOf(dat(rw, 1))) Then locTxt = dLoc(KeyOf(dat(rw, 1)))

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
            hit = hit + 1
            out(hit, 1) = dat(rw, 1)
            out(hit, 2) = dat(rw, 2)
            out(hit, 3) = locTxt
        End If
    Next rw

    ' clear old results
    wsD.Range(wsD.Cells(HDR_ROW + 1, R_COL), wsD.Cells(wsD.Rows.Count, R_COL + 2)).Clear

    If hit > 0 Then
        wsD.Cells(HDR_ROW + 1, R_COL).Resize(hit, 3).Value = out
        With wsD.Range(wsD.Cells(HDR_ROW + 1, R_COL), wsD.Cells(HDR_ROW + hit, R_COL + 2))
            .Borders.LineStyle = xlContinuous
            .Borders.Color = RGB(210, 210, 210)
        End With
    End If

    wsD.Cells(3, R_COL).Value = "Result: " & hit & " SKU(s) of " & (UBound(dat, 1) - 1)
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "ApplyFilters failed: " & Err.Description, vbCritical
End Sub

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
    wsD.Range(wsD.Cells(HDR_ROW + 1, R_COL), wsD.Cells(wsD.Rows.Count, R_COL + 2)).Clear
    wsD.Cells(3, R_COL).Value = "Result: not run yet"
End Sub

'==================================================================
Private Function ResolveDataSheets(ByRef wsP As Worksheet, ByRef wsL As Worksheet) As Boolean
    Dim ws As Worksheet, col As New Collection, i As Long
    Dim h1 As String, h2 As String

    For Each ws In TB.Worksheets
        If ws.Name <> DASH And ws.Name <> LISTS Then col.Add ws
    Next ws
    If col.Count < 2 Then Exit Function

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
    ResolveDataSheets = Not (wsP Is Nothing Or wsL Is Nothing)
End Function

Private Function ReadUsed(ws As Worksheet) As Variant
    Dim lr As Long, lc As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lr < 1 Or lc < 1 Then Exit Function
    If lr = 1 And lc = 1 And IsEmpty(ws.Cells(1, 1)) Then Exit Function
    ReadUsed = ws.Range(ws.Cells(1, 1), ws.Cells(lr, lc)).Value
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
