VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} cGeneralReport 
   ClientHeight    =   8145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11175
   OleObjectBlob   =   "crGeneralReport.dsx":0000
End
Attribute VB_Name = "cGeneralReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim rscompany As New ADODB.Recordset
Dim rsAdjustment As New ADODB.Recordset

Private Sub Report_Initialize()
If cnData.State = 0 Then
    Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881")
End If
 Set rsAdjustment = Open_Table(cnData, "Adjustment")

End Sub

Private Sub Report_Terminate()
    On Error GoTo Handle
        Set rscompany = Nothing
        Set rsAdjustment = Nothing
    Exit Sub
Handle:
    MsgBox Me.Name & " Report_Terminate " & Err.Number & " " & Err.Description
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Section10_Format"
End Sub

Private Sub Section12_Format(ByVal pFormattingInfo As Object)
If CDbl("0" & txtAmountOpen.Value) = 0 Then Section12.Suppress = True
End Sub

Private Sub Section14_Format(ByVal pFormattingInfo As Object)
If CDbl("0" & txtAmountCA.Value) = 0 Then Section14.Suppress = True
End Sub

Private Sub Section15_Format(ByVal pFormattingInfo As Object)
If CDbl("0" & txtAmountOA.Value) = 0 Then Section15.Suppress = True
End Sub

Private Sub Section16_Format(ByVal pFormattingInfo As Object)
If CDbl("0" & txtAmountCheck.Value) = 0 Then Section16.Suppress = True
End Sub

Private Sub Section17_Format(ByVal pFormattingInfo As Object)
If CDbl("0" & txtAmountCredit.Value) = 0 Then Section17.Suppress = True
End Sub

Private Sub Section18_Format(ByVal pFormattingInfo As Object)
If CDbl("0" & txtAmountGiftCard.Value) = 0 Then Section18.Suppress = True
End Sub

Private Sub Section11_Format(ByVal pFormattingInfo As Object)
If CDbl("0" & txtDelNotAmt.Value) = 0 Then Section11.Suppress = True
End Sub

Private Sub Section13_Format(ByVal pFormattingInfo As Object)
If CDbl("0" & txtDeleteAmt.Value) = 0 Then Section13.Suppress = True
End Sub

Private Sub Section19_Format(ByVal pFormattingInfo As Object)
txtReadTotal.SetText readnumber(CDbl("0" & Abs(txtCashInDrawer.Value))) & " ÆÂng./."
End Sub

Private Sub Section2_Format(ByVal pFormattingInfo As Object)
If CDbl("0" & txtDisAmt.Value) = 0 Then Section2.Suppress = True
End Sub

Private Sub Section20_Format(ByVal pFormattingInfo As Object)
If adjAmt1.Value = 0 Then Section20.Suppress = True
    With rsAdjustment
        .Find "AdjNo='01'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj1.SetText .Fields("AdjName")
        End If
    End With
End Sub

Private Sub Section21_Format(ByVal pFormattingInfo As Object)
    If txtKaAmt.Value = 0 Then Section21.Suppress = True
End Sub

Private Sub Section22_Format(ByVal pFormattingInfo As Object)
    If txtServAmt.Value = 0 Then Section22.Suppress = True
End Sub

Private Sub Section23_Format(ByVal pFormattingInfo As Object)
If AdjAmt2.Value = 0 Then Section23.Suppress = True
With rsAdjustment
        .Find "AdjNo='02'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj2.SetText .Fields("AdjName")
        End If
    End With
End Sub

Private Sub Section24_Format(ByVal pFormattingInfo As Object)
    If AmtEx.Value = 0 Then Section24.Suppress = True
End Sub

Private Sub Section25_Format(ByVal pFormattingInfo As Object)
    If txtAmountReceiveMoney.Value = 0 Then Section25.Suppress = True
End Sub

Private Sub Section26_Format(ByVal pFormattingInfo As Object)
    If AmountDiskar.Value = 0 Then Section26.Suppress = True
End Sub

Private Sub Section3_Format(ByVal pFormattingInfo As Object)
'txtDrawer.SetText Format(txtCashInDrawer.Value + CDbl("0" & txtReceiptAmt.Value) & CDbl("0" & txtExpenseAmt.Value), formatNum)
If CDbl("0" & txtTotal.Value) = 0 Then Section3.Suppress = True
End Sub

Private Sub Section4_Format(ByVal pFormattingInfo As Object)
If CDbl("0" & txtExpenseAmt.Value) = 0 Then Section4.Suppress = True
End Sub

Private Sub Section5_Format(ByVal pFormattingInfo As Object)
If CDbl("0" & txtReceiptAmt.Value) = 0 Then Section5.Suppress = True
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    Set rscompany = Open_Table(cnData, "Setup")
    If rscompany.RecordCount > 0 Then
        lblInfor1.SetText rscompany.Fields("Company_info_1")
        lblInfor2.SetText rscompany.Fields("Company_info_2")
        lblInfor3.SetText "ßﬁa chÿ :" & rscompany.Fields("Company_info_3")
        lblInfor4.SetText "ßi÷n thoπi :" & rscompany.Fields("Company_info_4") & Space(10) & "Fax:" & rscompany.Fields("Company_info_5")
        Picture1.SetOleLocation (rscompany!Image)
    End If
Exit Sub
Handle:

MsgBox "Bao loi" & Err.Number & " " & Err.Description
End Sub

Private Sub Section9_Format(ByVal pFormattingInfo As Object)
   '
End Sub
