VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crSaleBill75 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16425
   OleObjectBlob   =   "crSaleBill75.dsx":0000
End
Attribute VB_Name = "crSaleBill75"
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
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\database.mdb", "100881administrator")
'    End If
    Set rsserver = Open_Table(cnData, "Table_Diagram_Sections")
    Set rscompany = Open_Table(cnData, "Setup")
    Set rsFCRate = Open_Table(cnData, "Media")
    Set rsAdjustment = Open_Table(cnData, "Adjustment")
    Set rsMainGroup = Open_Table(cnData, "MainGroup")
    Set rsInventory = Open_Table(cnData, "Inventory")
'
        If ArrayFlag(SF(5), 1) = 1 Then
            Section8.Suppress = False
        End If
        If ArrayFlag(SF(5), 2) = 1 Then
            Section38.Suppress = False
        End If
        If ArrayFlag(SF(5), 3) = 1 Then
            Section39.Suppress = False
        End If
        If ArrayFlag(SF(5), 4) = 1 Then
            Section38.Suppress = True
            Section39.Suppress = True
            Section8.Suppress = True
        End If
'

End Sub

Private Sub Section1_Format(ByVal pFormattingInfo As Object)
    On Error GoTo Handle
    txtFoot1.SetText rscompany!Invoice_Notes_1
    If txtFoot1.Text = "" Then Section1.Suppress = True
    Exit Sub
Handle:
    MsgBox "B¸o lçi _CrSaleBill_Section1_format"
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle

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
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Section10_Format"
End Sub

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
End Sub



Private Sub Section19_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
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
Handle:
MsgBox Err.Number & Err.Description & " Section25_Format"
End Sub

Private Sub Section21_Format(ByVal pFormattingInfo As Object)
    If txtServAmt.Value = 0 Then Section21.Suppress = True
End Sub

Private Sub Section22_Format(ByVal pFormattingInfo As Object)
If txtAdj4.Value = 0 Then Section22.Suppress = True
With rsAdjustment
        .Find "AdjNo='04'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj4.SetText .Fields("AdjName")
        End If
    End With
End Sub

Private Sub Section23_Format(ByVal pFormattingInfo As Object)
If txtAdj3.Value = 0 Then Section23.Suppress = True
With rsAdjustment
        .Find "AdjNo='03'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj3.SetText .Fields("AdjName")
        End If
    End With
End Sub

Private Sub Section24_Format(ByVal pFormattingInfo As Object)
    If txtAdj2.Value = 0 Then Section24.Suppress = True
    With rsAdjustment
        .Find "AdjNo='02'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj2.SetText .Fields("AdjName")
        End If
    End With
End Sub

Private Sub Section25_Format(ByVal pFormattingInfo As Object)
If txtAdj1.Value = 0 Then Section25.Suppress = True
With rsAdjustment
        .Find "AdjNo='01'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj1.SetText .Fields("AdjName")
        End If
    End With
End Sub

Private Sub Section27_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    With rsMainGroup
        .Find "GroupNo='" & Trim$(txtMaingroup.Value) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtGroupName.SetText .Fields("GroupName") & ":"
        End If
    End With
        If ArrayFlag(SF(0), 5) = 0 Then Section27.Suppress = True
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Section27_Format"
End Sub

Private Sub Section28_Format(ByVal pFormattingInfo As Object)
    If txtCash.Value = TxtTotal.Value _
    Or txtCash.Value = txtPayment.Value _
    Or txtCash.Value = txtOAPAYMENT.Value _
    Or txtCash.Value = txtCCPAYMENT.Value _
    Or txtCash.Value = txtCTPAYMENT.Value _
    Or txtCash.Value = txtROAPAYMENT.Value _
    Or txtCash.Value = txtGCPAYMENT.Value Then Section28.Suppress = True
End Sub

Private Sub Section29_Format(ByVal pFormattingInfo As Object)
    txtFoot5.SetText rscompany!Invoice_Notes_5
    If txtFoot5.Text = "" Then Section29.Suppress = True
End Sub

Private Sub Section3_Format(ByVal pFormattingInfo As Object)
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

Private Sub Section31_Format(ByVal pFormattingInfo As Object)
    txtFoot4.SetText rscompany!Invoice_Notes_4
    If txtFoot4.Text = "" Then Section31.Suppress = True
End Sub

Private Sub Section32_Format(ByVal pFormattingInfo As Object)
    If txtVAT.Value = 0 Then Section32.Suppress = True
End Sub


Private Sub Section33_Format(ByVal pFormattingInfo As Object)
    txtFoot3.SetText rscompany!Invoice_Notes_3
    If txtFoot3.Text = "" Then Section33.Suppress = True
End Sub

Private Sub Section34_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    Dim rsEmployee As New ADODB.Recordset
    Set rsEmployee = Open_Table(cnData, "Employee")
        If txtNVPV.Value <> "" Then
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
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section34_Format-Hien thi nhan vien phuc vu "
End Sub

Private Sub Section35_Format(ByVal pFormattingInfo As Object)
    txtCashName.SetText userName
End Sub

Private Sub Section36_Format(ByVal pFormattingInfo As Object)
    If LineDisc.Value = 0 Then Section36.Suppress = True
End Sub

Private Sub Section38_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    Picture2.SetOleLocation (rscompany!Image)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Kh«ng t×m thÊy file Logo"
End Sub

Private Sub Section39_Format(ByVal pFormattingInfo As Object)
    If rscompany.RecordCount > 0 Then
        lblText1.SetText rscompany.Fields("Company_info_1")
        lblText2.SetText rscompany.Fields("Company_info_2")
        lblText3.SetText rscompany.Fields("Company_info_3")
        lblText4.SetText rscompany.Fields("Company_info_4")
        lblText5.SetText rscompany.Fields("Company_info_5")

    End If
End Sub

Private Sub Section40_Format(ByVal pFormattingInfo As Object)
    txtFoot2.SetText rscompany!Invoice_Notes_2
    If txtFoot2.Text = "" Then Section40.Suppress = True
End Sub

Private Sub Section41_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If txtCustomer.Value = "101" Or txtCustomer.Value = "" Then
        Section41.Suppress = True
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
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Section41_Format"
End Sub

Private Sub Section42_Format(ByVal pFormattingInfo As Object)
    If SumLineDisc.Value = 0 Then Section42.Suppress = True
End Sub

Private Sub Section43_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If txtsokhach.Value = 0 Then Section43.Suppress = True
Exit Sub
Handle: MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub Section44_Format(ByVal pFormattingInfo As Object)
    If txtCAPAYMENT.Value = 0 Then Section44.Suppress = True
    If ArrayFlag(SF(5), 8) = 1 Then Section44.Suppress = True
End Sub

Private Sub Section45_Format(ByVal pFormattingInfo As Object)
    If txtOAPAYMENT.Value = 0 Then Section45.Suppress = True
    If ArrayFlag(SF(5), 8) = 1 Then Section45.Suppress = True
End Sub

Private Sub Section46_Format(ByVal pFormattingInfo As Object)
    If txtCTPAYMENT.Value = 0 Then Section46.Suppress = True
    If ArrayFlag(SF(5), 8) = 1 Then Section46.Suppress = True
End Sub

Private Sub Section47_Format(ByVal pFormattingInfo As Object)
    If txtCCPAYMENT.Value = 0 Then Section47.Suppress = True
    If ArrayFlag(SF(5), 8) = 1 Then Section47.Suppress = True
End Sub

Private Sub Section48_Format(ByVal pFormattingInfo As Object)
    If txtROAPAYMENT.Value = 0 Then Section48.Suppress = True
    If ArrayFlag(SF(5), 8) = 1 Then Section48.Suppress = True
End Sub

Private Sub Section49_Format(ByVal pFormattingInfo As Object)
    If txtGCPAYMENT.Value = 0 Then Section49.Suppress = True
    If ArrayFlag(SF(5), 8) = 1 Then Section49.Suppress = True
End Sub

Private Sub Section5_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If CDbl("0" & txtDiscount.Value) = 0 Then
        Section5.Suppress = True
    End If
     Dim rsMixmatch As New ADODB.Recordset
    Set rsMixmatch = Open_Table(cnData, "Promotion")
    With rsMixmatch
        .Find "Pro_ID='" & Format(txtMixmatch.Value, "00") & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblDiscount.SetText .Fields("Pro_Name")
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - Section5_Format"
End Sub

Private Sub Section50_Format(ByVal pFormattingInfo As Object)
 If ArrayFlag(SF(5), 6) = 1 Then Section50.Suppress = True
End Sub

Private Sub Section51_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle

If ArrayFlag(SF(5), 7) = 1 Then
        Section51.Suppress = True
        Exit Sub
    End If
    With rsserver
    .Find "Location_ID='" & txtserver.Value & "'", , adSearchForward, adBookmarkFirst
    If Not .EOF Then
        lblTextSever.SetText .Fields("Section_ID")
    End If
End With
Exit Sub
Handle:
MsgBox Err.Number & " Section51_Format"

End Sub

Private Sub Section52_Format(ByVal pFormattingInfo As Object)
If ArrayFlag(SF(5), 5) = 1 Then Section52.Suppress = True
End Sub

Private Sub Section53_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
lblDateIn.SetText gfCONVERT_STRING_TO_DATE(txtDateIn.Value)
lblDateOut.SetText gfCONVERT_STRING_TO_DATE(txtDateOut.Value)
Exit Sub
Handle:
MsgBox Err.Number & " Section53_Format"
End Sub

Private Sub Section54_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If txtDatcoc.Value = 0 Then Section54.Suppress = True
Exit Sub
Handle:
MsgBox Err.Number & " Section54_Format"


End Sub

Private Sub Section55_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle

If txtFinal.Value = txtCash.Value Then Section55.Suppress = True
Exit Sub
Handle:
MsgBox Err.Number & " Section55_Format"

End Sub

Private Sub Section56_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
If txtAdj5.Value = 0 Then Section56.Suppress = True
With rsAdjustment
        .Find "AdjNo='05'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj5.SetText .Fields("AdjName")
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & " Section56_Format"
End Sub

Private Sub Section57_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
If txtAdj6.Value = 0 Then Section57.Suppress = True
With rsAdjustment
        .Find "AdjNo='06'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj6.SetText .Fields("AdjName")
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & " Section9_Format"
End Sub

'
Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If rscompany.RecordCount > 0 Then
        lblInfor1.SetText rscompany.Fields("Company_info_1")
        lblInfor2.SetText rscompany.Fields("Company_info_2")
        lblInfor3.SetText rscompany.Fields("Company_info_3")
        lblInfor4.SetText rscompany.Fields("Company_info_4") & "-" & rscompany.Fields("Company_info_5")
        Picture1.SetOleLocation (rscompany!Image)

    End If
Exit Sub
Handle:
MsgBox Err.Number & " Kh«ng t×m thÊy file logo! Vµo CÊu h×nh hÖ thèng-->Th«ng tin ®Çu cuèi H§-->T.Tin ®ång bé d÷ liÖu--> KÝch vµo khung Logo h×nh--> chän file logo (*.bmp) -->OK"
End Sub

Private Sub Section9_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle

    If TxtTotal.Value < 0 Then
        txtRead.SetText readnumber(CDbl("0" & Abs(txtCash.Value))) & " ®ång"
    Else
        txtRead.SetText readnumber(CDbl("0" & txtCash.Value)) & " ®ång"
    End If
Exit Sub
Handle:
MsgBox Err.Number & " Section9_Format"
End Sub

