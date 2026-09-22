Attribute VB_Name = "GenerateRoaster_V1_0"
Sub GenerateRoaster()
    Dim wsSchedule As Worksheet
    Dim wsRoaster As Worksheet
    Dim lastRowSched As Long, targetRow As Long
    Dim r As Long, col As Long
    Dim empId As String, empPos As String, shiftText As String, dateVal As String
    Dim startTime As String, endTime As String
    Dim startDateObj As Date, endDateObj As Date
    Dim startingPoint As String
    
    ' Software Version 1.1
    ' developed by DSU11425

    startingPoint = "Falaj_Hazza" ' <<<<< EDIT HERE
    
    Set wsSchedule = ActiveSheet
    
    On Error Resume Next
    Set wsRoaster = ActiveWorkbook.Sheets("Roaster")
    On Error GoTo 0
    
    If Not wsRoaster Is Nothing Then
        Application.DisplayAlerts = False
        wsRoaster.Delete
        Application.DisplayAlerts = True
    End If
    
    Set wsRoaster = ActiveWorkbook.Sheets.Add(After:=wsSchedule)
    wsRoaster.Name = "Roaster"
    wsRoaster.Cells(3, 1).Value = "shift_id"
    wsRoaster.Cells(3, 2).Value = "adjustment_type"
    wsRoaster.Cells(3, 3).Value = "employee_id"
    wsRoaster.Cells(3, 4).Value = "starting_point_id"
    wsRoaster.Cells(3, 5).Value = "start_date"
    wsRoaster.Cells(3, 6).Value = "end_date"
    wsRoaster.Cells(3, 7).Value = "start_time (local)"
    wsRoaster.Cells(3, 8).Value = "end_time (local)"
    wsRoaster.Cells(3, 9).Value = "shift_slots"
    
    wsRoaster.Rows(3).Font.Bold = True
    
    targetRow = 4
    lastRowSched = wsSchedule.Cells(wsSchedule.Rows.Count, "D").End(xlUp).Row
    
    Dim dRow As Long
    Dim dateHeaders(1 To 7) As String
    
    For r = 1 To lastRowSched
        If UCase(Trim(wsSchedule.Cells(r, 1).Value)) = "EMPLPYEE ID" Or UCase(Trim(wsSchedule.Cells(r, 1).Value)) = "EMPLOYEE ID" Then
            dRow = r - 1
            If dRow > 0 Then
                For col = 4 To 10
                    dateHeaders(col - 3) = wsSchedule.Cells(dRow, col).Value
                Next col
            End If
            r = r + 1
        End If
        
        If r <= lastRowSched And wsSchedule.Cells(r, 1).Value <> "" Then
            empId = Trim(wsSchedule.Cells(r, 1).Value)
            empPos = Trim(wsSchedule.Cells(r, 3).Value)
            
            If IsTargetPosition(empPos) Then
                For col = 4 To 10
                    shiftText = Trim(wsSchedule.Cells(r, col).Value)
                    dateVal = dateHeaders(col - 3)
                    
                    If shiftText <> "" And UCase(shiftText) <> "OFF" Then
                        Call ParseShiftTimes(shiftText, startTime, endTime)
                        If startTime <> "" And endTime <> "" Then
                            startDateObj = CDate(dateVal)
                            endDateObj = startDateObj
                            
                            If CInt(Left(endTime, 2)) < CInt(Left(startTime, 2)) Then
                                endDateObj = startDateObj + 1
                            End If
                            
                            wsRoaster.Cells(targetRow, 2).Value = "CREATE"
                            wsRoaster.Cells(targetRow, 3).Value = empId
                            wsRoaster.Cells(targetRow, 4).Value = startingPoint
                            wsRoaster.Cells(targetRow, 5).Value = Format(startDateObj, "M/d/yyyy")
                            wsRoaster.Cells(targetRow, 6).Value = Format(endDateObj, "M/d/yyyy")
                            wsRoaster.Cells(targetRow, 7).Value = startTime
                            wsRoaster.Cells(targetRow, 8).Value = endTime
                            
                            targetRow = targetRow + 1
                        End If
                    End If
                Next col
            End If
        End If
    Next r
    
    MsgBox "Roaster sheet successfully created with Night Shift date handling up to row " & (targetRow - 1) & "!", vbInformation, "Done"
End Sub

Function IsTargetPosition(pos As String) As Boolean
    Dim p As String
    p = UCase(Trim(pos))
    If InStr(p, "SUPERVISOR") > 0 Or InStr(p, "PICKER") > 0 Then
        IsTargetPosition = True
    Else
        IsTargetPosition = False
    End If
End Function

Sub ParseShiftTimes(ByVal sText As String, ByRef sTime As String, ByRef eTime As String)
    sText = UCase(Trim(sText))
    sTime = ""
    eTime = ""
    
    Dim cleanText As String
    cleanText = Replace(sText, " - ", "-")
    cleanText = Replace(cleanText, "  ", " ")
    
    Dim parts() As String
    If InStr(cleanText, "-") > 0 Then
        parts = Split(cleanText, "-")
        sTime = ConvertTo24Hour(Trim(parts(0)))
        eTime = ConvertTo24Hour(Trim(parts(1)))
    ElseIf InStr(cleanText, "TO") > 0 Then
        parts = Split(cleanText, " TO ")
        sTime = ConvertTo24Hour(Trim(parts(0)))
        eTime = ConvertTo24Hour(Trim(parts(1)))
    End If
End Sub

Function ConvertTo24Hour(tStr As String) As String
    tStr = UCase(Trim(tStr))
    Dim timePart As String
    Dim ampm As String
    
    If InStr(tStr, "AM") > 0 Then
        ampm = "AM"
        timePart = Trim(Replace(tStr, "AM", ""))
    ElseIf InStr(tStr, "PM") > 0 Then
        ampm = "PM"
        timePart = Trim(Replace(tStr, "PM", ""))
    Else
        ConvertTo24Hour = tStr
        Exit Function
    End If
    
    Dim hourVal As Integer, minVal As Integer
    If InStr(timePart, ":") > 0 Then
        Dim subParts() As String
        subParts = Split(timePart, ":")
        hourVal = CInt(subParts(0))
        minVal = CInt(subParts(1))
    Else
        hourVal = CInt(timePart)
        minVal = 0
    End If
    
    If ampm = "PM" And hourVal < 12 Then
        hourVal = hourVal + 12
    ElseIf ampm = "AM" And hourVal = 12 Then
        hourVal = 0
    End If
    
    ConvertTo24Hour = Format(hourVal, "00") & ":" & Format(minVal, "00")
End Function
