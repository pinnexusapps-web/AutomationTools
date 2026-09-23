Attribute VB_Name = "MultiSelectHandler_v1_2"
Option Explicit

Public WithEvents App As Application

Private Const DASH_NAME As String = "Dashboard"
Private Const ALL_TAG   As String = "(All)"
Private Const FIRST_ROW As Long = 6      ' first parameter row (must match P_ROW1)
Private Const TYPE_COL  As Long = 4      ' "Text" / "Number" column
Private Const VAL_COL   As Long = 2      ' the dropdown cell column

    ' developed by DSU11425
    ' Software Version 1.2.1

Private Sub App_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Dim typ As String
    Dim newVal As String, oldVal As String, res As String
    Dim toks() As String, i As Long, found As Boolean

    On Error GoTo EH

    If Sh.Name <> DASH_NAME Then Exit Sub
    If Target.Cells.Count <> 1 Then Exit Sub
    If Target.Column <> VAL_COL Then Exit Sub
    If Target.Row < FIRST_ROW Then Exit Sub

    typ = ""
    typ = CStr(Sh.Cells(Target.Row, TYPE_COL).Value)
    If typ <> "Text" Then Exit Sub          ' Number rows use Min/Max, not a list

    Application.EnableEvents = False

    newVal = Trim$(CStr(Target.Value))

    ' Recover the pre-click value via Undo (this is what makes an in-cell
    ' dropdown "remember" prior picks instead of just overwriting them).
    oldVal = ""
    On Error Resume Next
    Application.Undo
    oldVal = Trim$(CStr(Target.Value))
    Target.Value = newVal
    On Error GoTo EH

    If newVal = "" Then
        Target.Value = ALL_TAG
    ElseIf newVal = ALL_TAG Then
        Target.Value = ALL_TAG
    ElseIf oldVal = "" Or oldVal = ALL_TAG Then
        Target.Value = newVal
    Else
        toks = Split(oldVal, ";")
        found = False
        For i = LBound(toks) To UBound(toks)
            If StrComp(Trim$(toks(i)), newVal, vbTextCompare) = 0 Then
                found = True
                Exit For
            End If
        Next i

        If found Then
            res = ""
            For i = LBound(toks) To UBound(toks)
                If StrComp(Trim$(toks(i)), newVal, vbTextCompare) <> 0 And Trim$(toks(i)) <> "" Then
                    res = res & IIf(res = "", "", "; ") & Trim$(toks(i))
                End If
            Next i
            If res = "" Then res = ALL_TAG
        Else
            res = oldVal & "; " & newVal
        End If
        Target.Value = res
    End If

    ' Parent Category / Category rows cascade into the level(s) below —
    ' still inside EnableEvents=False so none of that triggers this handler again.
    On Error Resume Next
    OnParamChanged Target.Row
    On Error GoTo EH

    Application.EnableEvents = True
    Exit Sub

EH:
    Application.EnableEvents = True
End Sub
