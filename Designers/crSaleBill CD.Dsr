VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crSaleBill 
   ClientHeight    =   9570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10185
   OleObjectBlob   =   "crSaleBill.dsx":0000
End
Attribute VB_Name = "crSaleBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim DescArr() As String
Dim rsserver As New ADODB.Recordset
Dim rscompany As New ADODB.Recordset
Dim rsFCRate As New ADODB.Recordset
Dim rsAdjustment As New ADODB.Recordset
Dim rsMainGroup As New ADODB.Recordset

Dim rsInventory As New ADODB.Recordset

Private Sub Report_Terminate()
    Set rsserver = Nothing
    Set rsFCRate = Nothing
    Set rsAdjustment = Nothing
    Set rsMainGroup = Nothing
    Set rsInventory = Nothing
End Sub

Private Sub Report_Initialize()
    DescArr = LoadLanguage(LngFile, "#02:005:")
    If cnData.State = 0 Then
        Set cnData = Get_Connection(WorkingFolder & "\database.mdb", "200587")
    End If
    Set rsserver = Open_Table(cnData, "Table_Diagram_Sections")
    Set rscompany = Open_Table(cnData, "Setup")
    Set rsFCRate = Open_Table(cnData, "Media")
    Set rsAdjustment = Open_Table(cnData, "Adjustment")
    Set rsMainGroup = Open_Table(cnData, "MainGroup")
    Set rsInventory = Open_Table(cnData, "Inventory")
'
'        If ArrayFlag(SF(5), 1) = 1 Then
'            Section38.Suppress = True
'            Section39.Suppress = True
'            Section8.Suppress = False
'        End If
'        If ArrayFlag(SF(5), 2) = 1 Then
'            Section38.Suppress = False
'            Section39.Suppress = True
'            Section8.Suppress = True
'        End If
'        If ArrayFlag(SF(5), 3) = 1 Then
'            Section38.Suppress = True
'            Section39.Suppress = False
'            Section8.Suppress = True
'        End If
'        If ArrayFlag(SF(5), 4) = 1 Then
'            Section38.Suppress = True
'            Section39.Suppress = True
'            Section8.Suppress = True
'        End If
'

End Sub

Private Sub Section1_Format(ByVal pFormattingInfo As Object)
    On Error GoTo handle
        txtFooter.SetText rscompany!Invoice_Notes_1 & "-" & rscompany!Invoice_Notes_2
        txtNote.SetText rscompany!Invoice_Notes_3
        txtLine4.SetText rscompany!Invoice_Notes_4
    Exit Sub

handle:
    MsgBox "B¸o lçi _CrSaleBill_Section1_format"
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
On Error GoTo handle

With rsInventory
    .Find "ItemNum='" & txtPluName.Value & "'", , adSearchForward, adBookmarkFirst
    If Not .EOF Then
        If ArrayFlag(.Fields("F1"), 3) = 1 Then
            With txtQty
                .DecimalPlaces = DecimalQtyNumber
                .DecimalSymbol = DecimalMark
                .ThousandsSeparators = True
                .ThousandSymbol = DigitGroupMark
            End With
        Else
            With txtQty
                .DecimalPlaces = 0
                .DecimalSymbol = DecimalMark
                .ThousandsSeparators = True
                .ThousandSymbol = DigitGroupMark
            End With
        End If
    End If
End With

Exit Sub
handle:
    MsgBox Err.Number & Err.Description & Me.Name & " Section10_Format"
End Sub

'Private Sub Section10_Format(ByVal pFormattingInfo As Object)
'    lblCost.SetText Format(txtCost.Value, "#,##0")
'End Sub

Private Sub Section12_Format(ByVal pFormattingInfo As Object)
    lblTotal.SetText DescArr(8)
End Sub
'
'Private Sub Section13_Format(ByVal pFormattingInfo As Object)
'Dim S As String
'Dim FCName As String
'Dim i As Integer
'
'If ArrayFlag(SF(3), 1) = 1 Then
'    With rsFCRate
'        For i = 1 To 25
'            .Find "MediaID='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
'            If Not .EOF Then
'                S = .Fields("F")
'                If Mid(Right("00000000" & HexToBin(S), 8), 2, 1) = 1 Then
'                    FCName = .Fields("MediaName")
'                    If CDbl("0" & .Fields("FCRate")) <> 0 Then
'                        FCRate.SetText .Fields("SYMBOL") & Format(txtTotalAmt.Value / CDbl("0" & .Fields("FCRate")), "#,##0.00")
'                        txtFC.SetText FCName
'                        Exit Sub
'                    End If
'
'                End If
'            End If
'        Next
'    Section13.Suppress = True
'    End With
'Else
'    Section13.Suppress = True
'End If
'
'End Sub

'Private Sub Section14_Format(ByVal pFormattingInfo As Object)
'Dim S As String
'Dim FCName As String
'Dim i As Integer
'If ArrayFlag(SF(3), 1) = 1 Then
'    With rsFCRate
'        For i = 1 To 25
'            .Find "MediaID='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
'            If Not .EOF Then
'                S = .Fields("F")
'                If Mid(Right("00000000" & HexToBin(S), 8), 2, 1) = 1 Then
'                    FCName = .Fields("MediaName")
'                    If CDbl("0" & .Fields("FCRate")) <> 0 Then
'                        txtFCAmt.SetText .Fields("SYMBOL") & Format((txtTotalAmt.Value + txtAmtDist.Value) / CDbl("0" & .Fields("FCRate")), "#,##0.00")
'                        txtFCName.SetText FCName
'                        GoTo 1
'                    End If
'
'                End If
'            End If
'        Next
'    Section14.Suppress = True
'    End With
'Else
'    Section14.Suppress = True
'End If
'1: If Section5.Suppress = True Then Section14.Suppress = True
'End Sub

Private Sub Section16_Format(ByVal pFormattingInfo As Object)
    If txtTotalAmt.Value = TxtTotal.Value Then
        Section16.Suppress = True
    End If
End Sub

Private Sub Section17_Format(ByVal pFormattingInfo As Object)
    If CDbl("0" & txtDiscount.Value) = 0 And CDbl(txtChange.Value) = 0 Then
        Section17.Suppress = True
    End If
    If CDbl(txtChange.Value) = 0 Then
        Section17.Suppress = True
    End If
    lblChange.SetText DescArr(11)
End Sub



Private Sub Section19_Format(ByVal pFormattingInfo As Object)
On Error GoTo handle
    With rsInventory
        .Find "ItemNum='" & txtPluName.Value & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If Len(.Fields("Modify_Number")) > 1 Then
                txtName2.SetText .Fields("Modify_Number")
            Else
                Section19.Suppress = True
            End If
        End If
    End With
Exit Sub
handle:
MsgBox Err.Number & Err.Description & " Section25_Format"
End Sub

Private Sub Section21_Format(ByVal pFormattingInfo As Object)
    If txtServAmt.Value = 0 Then Section21.Suppress = True
End Sub


'Private Sub Section22_Format(ByVal pFormattingInfo As Object)
'If txtAdj4.Value = 0 Then Section22.Suppress = True
'    With rsAdjustment
'        .Find "AdjNo='04'", , adSearchForward, adBookmarkFirst
'        If Not .EOF Then
'            lblAdj4.SetText .Fields("AdjName")
'            RateAdj4.SetText .Fields("AdjRate")
'        End If
'    End With
'End Sub

'Private Sub Section23_Format(ByVal pFormattingInfo As Object)
'If txtAdj3.Value = 0 Then Section23.Suppress = True
'    With rsAdjustment
'        .Find "AdjNo='03'", , adSearchForward, adBookmarkFirst
'        If Not .EOF Then
'            lblAdj3.SetText .Fields("AdjName")
'            RateAdj3.SetText .Fields("AdjRate")
'        End If
'    End With
'End Sub

Private Sub Section24_Format(ByVal pFormattingInfo As Object)
If txtAdj2.Value = 0 Then Section24.Suppress = True
    
End Sub

Private Sub Section25_Format(ByVal pFormattingInfo As Object)
If txtAdj1.Value = 0 Then Section25.Suppress = True
End Sub

Private Sub Section27_Format(ByVal pFormattingInfo As Object)
On Error GoTo handle
    With rsMainGroup
        .Find "GroupNo='" & Trim$(txtMaingroup.Value) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtGroupName.SetText .Fields("GroupName") & ":"
        End If
    End With
        If ArrayFlag(SF(0), 5) = 0 Then Section27.Suppress = True
Exit Sub
handle:
MsgBox Err.Number & Err.Description & Me.Name & "  Section27_Format"

End Sub

'
'Private Sub Section28_Format(ByVal pFormattingInfo As Object)
'If txtdatcoc.Value = 0 Then Section28.Suppress = True
'End Sub

Private Sub Section3_Format(ByVal pFormattingInfo As Object)
    'If txtMethod.Value = "CA" Then
        lblTender.SetText DescArr(10)
    'Else
    '
    'End If
    If CDbl("0" & txtDiscount.Value) = 0 And CDbl(txtChange.Value) = 0 Then
        Section3.Suppress = True
    End If
    If CDbl(txtChange.Value) = 0 Then
        Section3.Suppress = True
    End If
End Sub

Private Sub Section30_Format(ByVal pFormattingInfo As Object)
    If txtMoney.Value = 0 Then Section30.Suppress = True
End Sub

Private Sub Section32_Format(ByVal pFormattingInfo As Object)
    If txtVAT.Value = 0 Then Section32.Suppress = True
End Sub


'Private Sub Section34_Format(ByVal pFormattingInfo As Object)
'    If txtNVPV.Value = 0 Then Section34.Suppress = True
'End Sub

Private Sub Section34_Format(ByVal pFormattingInfo As Object)
On Error GoTo handle
    Dim rsEmployee As New ADODB.Recordset
    Set rsEmployee = Open_Table(cnData, "Employee")
        If txtNVPV.Value <> 0 Then
            With rsEmployee
                .Find "Cashier_ID='" & Trim(txtNVPV.Value) & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    lblNVPV.SetText .Fields("EmpName")
                End If
            End With
        Else
            Section34.Suppress = True
        End If
Exit Sub
handle:
MsgBox Err.Number & Err.Description & Me.Name & " Section34_Format-Hien thi nhan vien phuc vu "
End Sub

Private Sub Section35_Format(ByVal pFormattingInfo As Object)
    lblCashier.SetText DescArr(13)
    txtCashName.SetText userName
End Sub

Private Sub Section36_Format(ByVal pFormattingInfo As Object)
On Error GoTo handle
    If txtCustomer.Value = "101" Or txtCustomer.Value = "" Then
        Section36.Suppress = True
        Exit Sub
    End If
    Dim rscust As New ADODB.Recordset
    Set rscust = Open_Table(cnData, "Customer")
    With rscust
        .Find "CustNum='" & txtCustomer.Value & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblCustomer.SetText .Fields("CustName")
        End If
    End With
Exit Sub
handle:
MsgBox Err.Number & Err.Description & Me.Name & "  Section36_Format"
End Sub

Private Sub Section37_Format(ByVal pFormattingInfo As Object)
    
    lblTitle.SetText "Hãa ®¬n t¹m tÝnh" 'DescArr(1)
'    lblBill.SetText DescArr(2)
'    'lblTable.SetText DescArr(3)
'    lblServer.SetText DescArr(16)
'    lblDate.SetText DescArr(17)
'    lblTime.SetText DescArr(18)
    '
    lblDatetime.SetText gfCONVERT_STRING_TO_DATE(txtDate.Value)
    With rsserver
    .Find "Location_ID='" & txtserver.Value & "'", , adSearchForward, adBookmarkFirst
    If Not .EOF Then
        lblTextSever.SetText .Fields("Section_ID")
    End If
End With
 
End Sub

'Private Sub Section38_Format(ByVal pFormattingInfo As Object)
'On Error GoTo handle
'    Picture2.SetOleLocation (rscompany!Image)
'
'Exit Sub
'handle:
'    MsgBox Err.Number & Err.Description & Me.Name
'End Sub



'Private Sub Section39_Format(ByVal pFormattingInfo As Object)
'    If rscompany.RecordCount > 0 Then
'        lblText1.SetText rscompany.Fields("Company_info_1")
'        lblText2.SetText rscompany.Fields("Company_info_2")
'        lblText3.SetText rscompany.Fields("Company_info_3")
'        lblText4.SetText rscompany.Fields("Company_info_4") & "-" & rscompany.Fields("Company_info_5")
'
'    End If
'End Sub



Private Sub Section5_Format(ByVal pFormattingInfo As Object)
    lblDiscount.SetText DescArr(9)
    'lblTotal1.SetText DescArr(15)
    If CDbl("0" & txtDiscount.Value) = 0 Then
        Section5.Suppress = True
    End If
End Sub

Private Sub Section6_Format(ByVal pFormattingInfo As Object)
'    lblItems.SetText DescArr(4)
'    lblQty.SetText DescArr(5)
'    lblPrice.SetText DescArr(6)
'    lblAmt.SetText DescArr(7)
End Sub

'
'Private Sub Section8_Format(ByVal pFormattingInfo As Object)
'On Error GoTo handle
'    If rscompany.RecordCount > 0 Then
'        lblInfor1.SetText rscompany.Fields("Company_info_1")
'        lblInfor2.SetText rscompany.Fields("Company_info_2")
'        lblInfor3.SetText rscompany.Fields("Company_info_3")
'        lblInfor4.SetText rscompany.Fields("Company_info_4") & "-" & rscompany.Fields("Company_info_5")
'        Picture1.SetOleLocation (rscompany!Image)
'
'    End If
'Exit Sub
'handle:
'
' MsgBox Err.Number & " File logo kh«ng t­¬ng thÝch", vbMsgBoxHelpButton, "Gióp ®ì"
'
'End Sub

'Private Sub Section9_Format(ByVal pFormattingInfo As Object)
'     lblRead.SetText DescArr(12)
'    If TxtTotal.Value < 0 Then
'        txtRead.SetText readnumber(CDbl("0" & Abs(TxtTotal.Value))) & " ®ång"
'    Else
'        txtRead.SetText readnumber(CDbl("0" & TxtTotal.Value)) & " ®ång"
'    End If
'
'End Sub
