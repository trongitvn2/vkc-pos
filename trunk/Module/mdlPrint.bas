Attribute VB_Name = "mdlPrint"


Public Sub Print_Buffer(BillNO As Double)
On Error GoTo handle
    Dim rs As New Recordset
    Dim DescArr() As String
    Dim SQL As String
    Dim ReceiptReport As CRAXDDRT.Report
    Dim iReport As CRAXDDRT.Report
    DescArr = LoadLanguage(LngFile, "#02:005:")
    Dim cmd As New ADODB.Command
    If ArrayFlag(SF(0), 5) = 0 Then
            SQL = "SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.KarDiscount, Invoice_Totals.CustNum, Invoice_Totals.Discount, Invoice_Totals.Total_Price, Invoice_Totals.Service_Charge, Invoice_Totals.VATFee, Invoice_Totals.Adjustment1, Invoice_Totals.Adj2Rate, Invoice_Totals.Adj1Rate, Invoice_Totals.Adjustment2, Invoice_Totals.Adjustment3, Invoice_Totals.Adjustment4, Invoice_Totals.AddMoney, Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered, Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID, Invoice_Totals.InvType, Invoice_Itemized.ItemNum, Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer, Sum(Invoice_Itemized.Quantity*Invoice_Itemized.PricePer) AS Amt, Invoice_Itemized.DiffItemName, Invoice_Totals.Orig_OnHoldID, Invoice_Totals.OrderMan, Right([OpenTime],8) AS TimeIn, Right([ClosingTime],8) AS TimeOut, Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount" & _
            " FROM ((((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number) INNER JOIN Inventory ON Invoice_Itemized.ItemNum = Inventory.ItemNum) INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID) INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo" & _
            " WHERE (((Invoice_Itemized.ItemNum)<>'KAR') AND ((Invoice_Totals.Invoice_Number)=" & BillNO & "))" & _
            " GROUP BY Invoice_Totals.Invoice_Number, Invoice_Totals.KarDiscount, Invoice_Totals.CustNum, Invoice_Totals.Discount, Invoice_Totals.Total_Price, Invoice_Totals.Service_Charge, Invoice_Totals.VATFee, Invoice_Totals.Adjustment1, Invoice_Totals.Adjustment2, Invoice_Totals.Adj2Rate, Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment3, Invoice_Totals.Adjustment4, Invoice_Totals.AddMoney, Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered, Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID, Invoice_Totals.InvType, Invoice_Itemized.ItemNum, Invoice_Itemized.PricePer, Invoice_Itemized.DiffItemName, Invoice_Totals.Orig_OnHoldID, Invoice_Totals.OrderMan, Right([OpenTime],8), Right([ClosingTime],8), Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount" & _
            " ORDER BY Invoice_Itemized.ItemNum"
    Else
        SQL = "SELECT Invoice_Totals.Invoice_Number,Invoice_Totals.KarDiscount, Invoice_Totals.CustNum," & _
            " Invoice_Totals.Discount, Invoice_Totals.Total_Price," & _
            " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adjustment1," & _
            " Invoice_Totals.Adjustment2,Invoice_Totals.Adjustment3," & _
            " Invoice_Totals.Adj2Rate, Invoice_Totals.Adj1Rate," & _
            " Invoice_Totals.Adjustment4,Invoice_Totals.AddMoney," & _
            " Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered," & _
            " Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID," & _
            " Invoice_Totals.Station_ID,Invoice_Totals.InvType,Invoice_Itemized.ItemNum, " & _
            " Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer," & _
            " sum(Invoice_Itemized.Quantity*Invoice_Itemized.PricePer) as Amt, " & _
            " Invoice_Itemized.DiffItemName ,Invoice_Totals.Orig_OnHoldID,Invoice_Totals.OrderMan, " & _
            " Right([OpenTime],8) AS TimeIn, Right([ClosingTime],8) AS TimeOut, Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount, MainGroup.GroupNo, MainGroup.GroupName " & _
            " FROM ((((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number) INNER JOIN Inventory ON Invoice_Itemized.ItemNum = Inventory.ItemNum) INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID) INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo " & _
            " Where Invoice_Itemized.ItemNum<>'KAR' and Invoice_Totals.Invoice_Number=" & BillNO & _
            " group by Invoice_Itemized.ItemNum,Invoice_Totals.Invoice_Number," & _
            " Invoice_Totals.CustNum,Invoice_Totals.Discount,Invoice_Totals.KarDiscount," & _
            " Invoice_Totals.Total_Price,Invoice_Totals.Grand_Total," & _
            " Invoice_Totals.Amt_Tendered,Invoice_Totals.Amt_Change," & _
            " Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID," & _
            " Invoice_Itemized.PricePer, Invoice_Itemized.DiffItemName ," & _
            " Invoice_Totals.Orig_OnHoldID,Invoice_Totals.OrderMan, Invoice_Totals.InvType, " & _
            " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adj2Rate, Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment1,Invoice_Totals.Adjustment2,Invoice_Totals.Adjustment3," & _
            " Invoice_Totals.Adjustment4,Invoice_Totals.AddMoney, Right([OpenTime],8), Right([ClosingTime],8), Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount, MainGroup.GroupNo, MainGroup.GroupName" & _
            " order by Invoice_Itemized.ItemNum"
            
   End If
    
    Set crBalanceA6 = Nothing
    Set crBalance = Nothing
    If ArrayFlag(SF(3), 8) = 1 Then
        Set ReceiptReport = crBalance
    Else
        Set ReceiptReport = crBalanceA6
    End If
  
        cmd.ActiveConnection = cnData
        cmd.CommandType = adCmdText
        cmd.CommandText = SQL
        cmd.Execute
    With ReceiptReport
        .Database.AddADOCommand cnData, cmd
        .txtPluCode.SetUnboundFieldSource "{ado.DiffItemName}"
        .txtPluName.SetUnboundFieldSource "{ado.ItemNum}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtCost.SetUnboundFieldSource "{ado.PricePer}"
        .txtAmt.SetUnboundFieldSource "{ado.Amt}"
        .txtBillNo.SetUnboundFieldSource "{ado.Invoice_Number}"
        .txtCashier.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtDiscount.SetUnboundFieldSource "{ado.Discount}"
        .txtTable.SetUnboundFieldSource "{ado.Orig_OnHoldID}"
        .txtserver.SetUnboundFieldSource "{ado.Station_ID}"
        .txtOrder.SetUnboundFieldSource "{ado.OrderMan}"
        .txtAdj1.SetUnboundFieldSource "{ado.Adjustment1}"
        .txtAdj1Rate.SetUnboundFieldSource "{ado.Adj1Rate}"
        .txtAdj2.SetUnboundFieldSource "{ado.Adjustment2}"
        .txtAdj2Rate.SetUnboundFieldSource "{ado.Adj2Rate}"
        .txtAdj3.SetUnboundFieldSource "{ado.Adjustment3}"
        .txtAdj4.SetUnboundFieldSource "{ado.Adjustment4}"
        .txtSev.SetUnboundFieldSource "{ado.Service_Charge}"
        .txtVAT.SetUnboundFieldSource "{ado.VATFee}"
        .txtMoney.SetUnboundFieldSource "{ado.AddMoney}"
        .PrintCount.SetUnboundFieldSource "{ado.InvType}"
            .lblPhuthu.SetText DescArr(14)
            .lblPhuphi.SetText DescArr(27)
            .lblAdj1.SetText DescArr(25)
            .lblAdj2.SetText DescArr(26)
        If Style = 1 Then
            .lblTitle.SetText DescArr(1)
        Else
            .lblTitle.SetText DescArr(24)
            If ArrayFlag(SF(0), 5) = 1 Then
                .txtMainGroup.SetUnboundFieldSource "{ado.GroupNo}"
            End If
        End If
        'canh le
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        
        With .txtCost
            
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtMoney
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAdj4
            
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAdj3
            
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAdj2
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAdj1
            
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAmt
            
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtChange
            
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        
        With .txtMainTotal
            
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtServAmt
            
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
'        With .txtQty
'
'            .DecimalPlaces = DecimalQtyNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
        
        With .txtTotal
            
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtTotalAmt
            
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
       
    End With
        Set iReport = ReceiptReport
        
    With frmShowBillBalance
        .Report = iReport
        .Show vbModal
        
    End With
Exit Sub
handle:
Exit Sub
    MsgBox Err.Number & Err.Description & " Print_Buffer"

End Sub
