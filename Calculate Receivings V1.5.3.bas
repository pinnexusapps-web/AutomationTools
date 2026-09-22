Attribute VB_Name = "CalculateReceivings_V1_5"
Sub CalculateReceivings()
    Dim wsMain As Worksheet
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim finalSum As Double
    Dim skuCell As Range
    Dim skuRow As Long, skuCol As Long
    Dim skuColLetter As String
    Dim i As Long, r As Long
    Dim retryCount As Integer
    Dim tblRange As Range

    ' developed by DSU11425
    ' Software Version 1.5.3

    On Error Resume Next
    Set wsMain = ActiveWorkbook.Sheets("Sheet1")
    Set wsData = ActiveWorkbook.Sheets(1)
    On Error GoTo 0

    If wsMain Is Nothing Then
        MsgBox "Error: Please make sure there is a sheet named 'Sheet1' in this workbook!", vbCritical, "Sheet Missing"
        Exit Sub
    End If

FindSKULabel:
    Set skuCell = wsMain.Cells.Find(What:="SKU", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If skuCell Is Nothing Then
        MsgBox "Error: Could not find the 'SKU' column header in Sheet1!", vbCritical, "Header Missing"
        Exit Sub
    End If

    skuRow = skuCell.Row
    skuCol = skuCell.Column

    If skuRow > 1 Then
        For r = skuRow - 1 To 1 Step -1
            If Application.WorksheetFunction.CountA(wsMain.Rows(r)) = 0 Then
                wsMain.Rows(r).Delete shift:=xlUp
            End If
        Next r
        Set skuCell = wsMain.Cells.Find(What:="SKU", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        skuRow = skuCell.Row
        skuCol = skuCell.Column
    End If

    If skuCol > 1 Then
        For i = skuCol - 1 To 1 Step -1
            If Application.WorksheetFunction.CountA(wsMain.Columns(i)) = 0 Then
                wsMain.Columns(i).Delete shift:=xlToLeft
            End If
        Next i
        Set skuCell = wsMain.Cells.Find(What:="SKU", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        skuRow = skuCell.Row
        skuCol = skuCell.Column
    End If

    skuColLetter = Split(wsMain.Cells(1, skuCol).Address, "$")(1)
    lastRow = wsMain.Cells(wsMain.Rows.Count, skuColLetter).End(xlUp).Row

    If lastRow <= skuRow Then
        If retryCount = 0 Then
            wsMain.Cells(skuRow, skuCol).Insert shift:=xlToRight
            retryCount = 1
            GoTo FindSKULabel
        Else
            MsgBox "Warning: No data found under the SKU column even after adjustment!", vbExclamation, "No Data"
            Exit Sub
        End If
    End If

    lastCol = wsMain.Cells(skuRow, wsMain.Columns.Count).End(xlToLeft).Column

    Dim colSupplierSKU As Long, colUnitsPerCase As Long, colTotalOrderedCases As Long
    Dim colConfirmedQty As Long, colTotalReceivedCases As Long, colReceivedQty As Long
    Dim colInternalTax As Long, colVAT As Long, colGrossUnitCost As Long
    Dim colCost As Long, colNetDiscountedCost As Long, colUnitCost As Long

    colSupplierSKU = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Supplier SKU")
    colUnitsPerCase = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Units Per Case")
    colTotalOrderedCases = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Total Ordered Cases")
    colConfirmedQty = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Confirmed Quantity")
    colTotalReceivedCases = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Total Received Cases")
    colReceivedQty = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Received Quantity")
    colInternalTax = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Internal Tax")
    colVAT = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "VAT")
    colGrossUnitCost = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Gross Unit Cost")
    colCost = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Cost")
    colNetDiscountedCost = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Net Discounted Cost")

    If colReceivedQty > 0 Then
        For r = skuRow + 1 To lastRow
            wsMain.Cells(r, colReceivedQty).Formula = _
                "=VLOOKUP(" & skuColLetter & r & ", '" & wsData.Name & "'!$L:$AB, 17, 0)"
        Next r
    End If

    Dim colsToDelete() As Long
    colsToDelete = SortDescendingNonZero(Array(colUnitsPerCase, colTotalOrderedCases, _
        colConfirmedQty, colTotalReceivedCases, colInternalTax, colVAT, _
        colGrossUnitCost, colCost, colNetDiscountedCost))

    For i = LBound(colsToDelete) To UBound(colsToDelete)
        If colsToDelete(i) > 0 Then
            wsMain.Columns(colsToDelete(i)).Delete shift:=xlToLeft
        End If
    Next i

    Set skuCell = wsMain.Cells.Find(What:="SKU", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    skuRow = skuCell.Row
    skuCol = skuCell.Column
    skuColLetter = Split(wsMain.Cells(1, skuCol).Address, "$")(1)
    lastRow = wsMain.Cells(wsMain.Rows.Count, skuColLetter).End(xlUp).Row
    lastCol = wsMain.Cells(skuRow, wsMain.Columns.Count).End(xlToLeft).Column

    colSupplierSKU = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Supplier SKU")
    colReceivedQty = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Received Quantity")
    colUnitCost = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Unit Cost")

    If colSupplierSKU > 0 And lastRow > skuRow Then
        Dim sortRange As Range
        Set sortRange = wsMain.Range(wsMain.Cells(skuRow, skuCol), wsMain.Cells(lastRow, lastCol))
        sortRange.Sort Key1:=wsMain.Cells(skuRow, colSupplierSKU), Order1:=xlAscending, _
                        Header:=xlYes, MatchCase:=False, Orientation:=xlTopToBottom
    End If

    Dim colOrderedQty As Long
    Dim colInvoiceQty As Long, colTCN As Long, colGrossCost As Long
    Dim colBarcode As Long, colTCNCost As Long

    colOrderedQty = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Ordered Quantity")
    colInvoiceQty = FindHeaderCol(wsMain, skuRow, skuCol, lastCol + 5, "Invoice Received Qty")
    
    If colInvoiceQty = 0 Then
        colInvoiceQty = colUnitCost + 1
        wsMain.Cells(skuRow, colInvoiceQty).Value = "Invoice Received Qty"
    End If

    colTCN = colInvoiceQty + 1
    wsMain.Cells(skuRow, colTCN).Value = "TCN"
    colGrossCost = colInvoiceQty + 2
    wsMain.Cells(skuRow, colGrossCost).Value = "Gross Cost"
    colBarcode = colInvoiceQty + 3
    wsMain.Cells(skuRow, colBarcode).Value = "Barcode"
    colTCNCost = colInvoiceQty + 4
    wsMain.Cells(skuRow, colTCNCost).Value = "TCN Cost"

    lastCol = colTCNCost

    Dim rqColLetter As String, ucColLetter As String, invColLetter As String, oqColLetter As String
    rqColLetter = Split(wsMain.Cells(1, colReceivedQty).Address, "$")(1)
    ucColLetter = Split(wsMain.Cells(1, colUnitCost).Address, "$")(1)
    invColLetter = Split(wsMain.Cells(1, colInvoiceQty).Address, "$")(1)
    oqColLetter = Split(wsMain.Cells(1, colOrderedQty).Address, "$")(1)

    For r = skuRow + 1 To lastRow
        If Trim(CStr(wsMain.Cells(r, colInvoiceQty).Value)) = "" Then
            wsMain.Cells(r, colInvoiceQty).Formula = "=" & oqColLetter & r
        End If

        wsMain.Cells(r, colTCN).Formula = "=" & invColLetter & r & "-" & rqColLetter & r
        wsMain.Cells(r, colGrossCost).Formula = _
            "=" & rqColLetter & r & "*1.05*" & ucColLetter & r
        wsMain.Cells(r, colBarcode).Formula = _
            "=VLOOKUP(" & skuColLetter & r & ", '" & wsData.Name & "'!$L:$P, 5, 0)"

        Dim tcnColLetterRow As String
        tcnColLetterRow = Split(wsMain.Cells(1, colTCN).Address, "$")(1)
        wsMain.Cells(r, colTCNCost).Formula = _
            "=1.05*" & tcnColLetterRow & r & "*" & ucColLetter & r
    Next r

    Dim tcnColLetter As String
    tcnColLetter = Split(wsMain.Cells(1, colTCN).Address, "$")(1)

    Dim highlightRange As Range
    Set highlightRange = wsMain.Range(wsMain.Cells(skuRow + 1, skuCol), wsMain.Cells(lastRow, lastCol))
    highlightRange.FormatConditions.Delete
    
    ' Condition 1: Received Qty = 0 -> Yellow
    highlightRange.FormatConditions.Add Type:=xlExpression, Formula1:="=$" & rqColLetter & (skuRow + 1) & "=0"
    highlightRange.FormatConditions(1).Interior.Color = RGB(255, 255, 0)
    
    ' Condition 2: Received Qty > 0 AND TCN > 0 -> Light Red
    highlightRange.FormatConditions.Add Type:=xlExpression, Formula1:="=AND($" & rqColLetter & (skuRow + 1) & ">0, $" & tcnColLetter & (skuRow + 1) & ">0)"
    highlightRange.FormatConditions(2).Interior.Color = RGB(255, 204, 204)

    Set tblRange = wsMain.Range(wsMain.Cells(skuRow, skuCol), wsMain.Cells(lastRow, lastCol))
    With tblRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(211, 211, 211)
    End With
    wsMain.Range(wsMain.Columns(skuCol), wsMain.Columns(lastCol)).AutoFit

    Dim gcColLetter As String, tccColLetter As String
    Dim tcnCostSum As Double
    gcColLetter = Split(wsMain.Cells(1, colGrossCost).Address, "$")(1)
    tccColLetter = Split(wsMain.Cells(1, colTCNCost).Address, "$")(1)

    wsMain.Cells(lastRow + 2, colGrossCost - 1).Value = "Total SUM:"
    wsMain.Cells(lastRow + 2, colGrossCost).Formula = _
        "=SUM(" & gcColLetter & (skuRow + 1) & ":" & gcColLetter & lastRow & ")"
    finalSum = Application.WorksheetFunction.Sum( _
        wsMain.Range(gcColLetter & (skuRow + 1) & ":" & gcColLetter & lastRow))

    wsMain.Cells(lastRow + 2, colTCNCost - 1).Value = "TCN Total:"
    wsMain.Cells(lastRow + 2, colTCNCost).Formula = _
         "=SUM(" & tccColLetter & (skuRow + 1) & ":" & tccColLetter & lastRow & ")"
    tcnCostSum = Application.WorksheetFunction.Sum( _
        wsMain.Range(tccColLetter & (skuRow + 1) & ":" & tccColLetter & lastRow))

    Dim fullTotalSum As Double
    wsMain.Cells(lastRow + 3, colTCNCost - 1).Value = "Full Total:"
    wsMain.Cells(lastRow + 3, colTCNCost).Formula = _
        "=" & gcColLetter & (lastRow + 2) & "+" & tccColLetter & (lastRow + 2)
    fullTotalSum = finalSum + tcnCostSum

    MsgBox "Process completed successfully!" & vbCrLf & vbCrLf & _
           "==============================" & vbCrLf & _
           " TOTAL GROSS COST: " & Format(finalSum, "#,##0.00") & vbCrLf & _
           " TCN TOTAL: " & Format(tcnCostSum, "#,##0.00") & vbCrLf & _
           " FULL TOTAL: " & Format(fullTotalSum, "#,##0.00") & vbCrLf & _
           "==============================", vbInformation, "Success"
End Sub

Private Function FindHeaderCol(ws As Worksheet, headerRow As Long, startCol As Long, endCol As Long, headerText As String) As Long
    Dim c As Long
    FindHeaderCol = 0
    For c = startCol To endCol
        If LCase(Trim(CStr(ws.Cells(headerRow, c).Value))) = LCase(Trim(headerText)) Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c
End Function

Private Function SortDescendingNonZero(arr As Variant) As Long()
    Dim n As Long, i As Long, j As Long, temp As Long
    Dim result() As Long
    n = UBound(arr) - LBound(arr) + 1
    ReDim result(0 To n - 1)
    For i = 0 To n - 1
        result(i) = CLng(arr(LBound(arr) + i))
    Next i
    For i = 0 To n - 2
        For j = 0 To n - 2 - i
             If result(j) < result(j + 1) Then
                temp = result(j)
                result(j) = result(j + 1)
                result(j + 1) = temp
            End If
        Next j
    Next i
    SortDescendingNonZero = result
End Function


Sub ExtractTCNSummary()
    Dim wsMain As Worksheet
    Dim skuCell As Range
    Dim skuRow As Long, skuCol As Long, lastCol As Long
    Dim lastRow As Long
    Dim skuColLetter As String
    Dim colSKU As Long, colTCN As Long, colUnitCost As Long
    Dim r As Long, outRow As Long
    Dim colM As Long, colN As Long, colO As Long

    ' developed by DSU11425
    ' Software Version 1.0.0

    On Error Resume Next
    Set wsMain = ActiveWorkbook.Sheets("Sheet1")
    On Error GoTo 0

    If wsMain Is Nothing Then
        MsgBox "Error: Please make sure there is a sheet named 'Sheet1' in this workbook!", vbCritical, "Sheet Missing"
        Exit Sub
    End If

    Set skuCell = wsMain.Cells.Find(What:="SKU", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If skuCell Is Nothing Then
        MsgBox "Error: Could not find the 'SKU' column header in Sheet1!", vbCritical, "Header Missing"
        Exit Sub
    End If

    skuRow = skuCell.Row
    skuCol = skuCell.Column
    skuColLetter = Split(wsMain.Cells(1, skuCol).Address, "$")(1)
    lastRow = wsMain.Cells(wsMain.Rows.Count, skuColLetter).End(xlUp).Row
    lastCol = wsMain.Cells(skuRow, wsMain.Columns.Count).End(xlToLeft).Column

    colSKU = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "SKU")
    colTCN = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "TCN")
    colUnitCost = FindHeaderCol(wsMain, skuRow, skuCol, lastCol, "Unit Cost")

    If colTCN = 0 Or colUnitCost = 0 Then
        MsgBox "Error: Could not find the 'TCN' or 'Unit Cost' column." & vbCrLf & _
               "Please run CalculateReceivings first.", vbCritical, "Columns Missing"
        Exit Sub
    End If

    ' Fixed output columns: M, N, O
    colM = wsMain.Range("M1").Column
    colN = wsMain.Range("N1").Column
    colO = wsMain.Range("O1").Column

    ' Clear any previous summary before rebuilding it
    wsMain.Range(wsMain.Cells(skuRow, colM), wsMain.Cells(wsMain.Rows.Count, colO)).ClearContents

    wsMain.Cells(skuRow, colM).Value = "SKU"
    wsMain.Cells(skuRow, colN).Value = "QTY"
    wsMain.Cells(skuRow, colO).Value = "price(no vat)"

    outRow = skuRow + 1
    For r = skuRow + 1 To lastRow
        Dim tcnVal As Double
        tcnVal = 0
        If IsNumeric(wsMain.Cells(r, colTCN).Value) Then tcnVal = wsMain.Cells(r, colTCN).Value

        If tcnVal <> 0 Then
            wsMain.Cells(outRow, colM).Value = wsMain.Cells(r, colSKU).Value
            wsMain.Cells(outRow, colN).Value = tcnVal
            wsMain.Cells(outRow, colO).Value = wsMain.Cells(r, colUnitCost).Value
            outRow = outRow + 1
        End If
    Next r

    If outRow > skuRow + 1 Then
        Dim tblRange As Range
        Set tblRange = wsMain.Range(wsMain.Cells(skuRow, colM), wsMain.Cells(outRow - 1, colO))
        With tblRange.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(211, 211, 211)
        End With
    End If

    wsMain.Range(wsMain.Columns(colM), wsMain.Columns(colO)).AutoFit

    MsgBox "TCN summary table built successfully!" & vbCrLf & vbCrLf & _
           (outRow - skuRow - 1) & " row(s) with non-zero TCN found.", _
           vbInformation, "Success"
End Sub
