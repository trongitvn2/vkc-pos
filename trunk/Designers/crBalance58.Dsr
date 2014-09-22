VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crBalance58 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16425
   OleObjectBlob   =   "crBalance58.dsx":0000
End
Attribute VB_Name = "crBalance58"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim DescArr() As String
Dim rsserver As New ADODB.Recordset
Dim rscompany As New ADODB.Recordset
Dim rsSystem As New ADODB.Recordset
Dim rsAdjustment As New ADODB.Recordset
Dim rsMainGroup As New ADODB.Recordset
Dim rsFCRate As New ADODB.Recordset
Dim rsInventory As New ADODB.Recordset
Private Sub Report_Initialize()
On Error GoTo Handle
    DescArr = LoadLanguage(LngFile, "#02:005:")
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\database.mdb", "100881administrator")
'    End If
    Set rsserver = Open_Table(cnData, "Table_Diagram_Sections")
    Set rscompany = Open_Table(cnData, "Setup")
    Set rsSystem = Open_Table(cnData, "SystemFlag")
    Set rsAdjustment = Open_Table(cnData, "Adjustment")
    Set rsMainGroup = Open_Table(cnData, "MainGroup")
    Set rsFCRate = Open_Table(cnData, "Media")
    Set rsInventory = Open_Table(cnData, "Inventory")

    With rsSystem
    If rsSystem.State <> 0 And .RecordCount > 0 Then .MoveFirst
        .Find "SF=06", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If ArrayFlag(.Fields("Data"), 1) = 1 Then
                Section8.Suppress = False
            End If
            If ArrayFlag(.Fields("Data"), 2) = 1 Then
                Section31.Suppress = False
            End If
            If ArrayFlag(.Fields("Data"), 3) = 1 Then
                Section32.Suppress = False
            End If
            If ArrayFlag(.Fields("Data"), 4) = 1 Then
                Section31.Suppress = True
                Section32.Suppress = True
                Section8.Suppress = True
            End If
        End If
    End With
    Exit Sub
Handle:
    MsgBox "B¸o lçi Report_Initialize"
End Sub

Private Sub Report_Terminate()
    Set rsserver = Nothing
    Set rsSystem = Nothing
    Set rsAdjustment = Nothing
    Set rsMainGroup = Nothing
    Set rsInventory = Nothing
End Sub

Private Sub Section1_Format(ByVal pFormattingInfo As Object)
    On Error GoTo Handle
        If TxtTotal.Value < 0 Then
            txtRead.SetText "©m" & readnumber(CDbl("0" & Abs(txtCash.Value))) & " ®ång"
        Else
            txtRead.SetText readnumber(CDbl("0" & txtCash.Value)) & " ®ång"
        End If
        
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
        If ArrayFlag(.Fields("F3"), 5) = 1 Then Section10.Suppress = True
    End If
End With

Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Section10_Format"
End Sub
Private Sub Section13_Format(ByVal pFormattingInfo As Object)
    txtFoot1.SetText rscompany!Invoice_Notes_1
    If txtFoot1.Text = "" Then Section13.Suppress = True
End Sub

Private Sub Section14_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle

Dim rsEmployee As New ADODB.Recordset
Set rsEmployee = Open_Table(cnData, "Employee")
    If txtOrder.Value <> "" Then
        With rsEmployee
            .Find "Cashier_ID='" & Trim(txtOrder.Value) & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtOrderName.SetText .Fields("EmpName")
            End If
        End With
    Else
        Section14.Suppress = True
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section14_Format"

End Sub


Private Sub Section15_Format(ByVal pFormattingInfo As Object)
    If TxtTotal.Value = txtTotalAmt.Value Then Section15.Suppress = True
End Sub

Private Sub Section16_Format(ByVal pFormattingInfo As Object)
    If txtSev.Value = 0 Then Section16.Suppress = True
End Sub

Private Sub Section17_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle

If txtAdj4.Value = 0 Then Section17.Suppress = True
    
    With rsAdjustment
    .MoveFirst
        .Find "AdjNo='04'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj4.SetText .Fields("AdjName")
        End If
    End With

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section17_Format"

End Sub

Private Sub Section18_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle

If txtAdj3.Value = 0 Then Section18.Suppress = True
    With rsAdjustment
        .Find "AdjNo='03'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj3.SetText .Fields("AdjName")
        End If
    End With

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section18_Format"

End Sub

Private Sub Section19_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle

If txtAdj2.Value = 0 Then Section19.Suppress = True
    With rsAdjustment
        .Find "AdjNo='02'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj2.SetText .Fields("AdjName")
        End If
    End With

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section19_Format"
End Sub

Private Sub Section2_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle

Dim S As String
Dim FCName As String
Dim i As Integer
With rsSystem
    .Find "SF='04'", , adSearchForward, adBookmarkFirst
    If Not .EOF Then
        S = .Fields("Data")
    End If
End With
If Left(Right("00000000" & HexToBin(S), 8), 1) = 1 Then
    With rsFCRate
        For i = 1 To 25
            .Find "MediaID='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                S = .Fields("F")
                If Mid(Right("00000000" & HexToBin(S), 8), 2, 1) = 1 Then
                    FCName = .Fields("MediaName")
                    If CDbl("0" & .Fields("FCRate")) <> 0 Then
                        txtFCAmount.SetText .Fields("SYMBOL") & Format(TxtTotal.Value / CDbl("0" & .Fields("FCRate")), "#,##0.00")
                        txtFCName.SetText FCName
                        GoTo 1
                    End If

                End If
            End If
        Next
    Section2.Suppress = True
    End With
Else
    Section2.Suppress = True
End If
1: If Section5.Suppress = True Then Section2.Suppress = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section20_Format"

End Sub

Private Sub Section20_Format(ByVal pFormattingInfo As Object)

On Error GoTo Handle

If txtAdj1.Value = 0 Then Section20.Suppress = True
    With rsAdjustment
        .Find "AdjNo='01'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj1.SetText .Fields("AdjName")
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section2_Format"

End Sub

Private Sub Section22_Format(ByVal pFormattingInfo As Object)
Dim S As String
On Error GoTo Handle
    With rsMainGroup
        .Find "GroupNo='" & Trim$(txtMaingroup.Value) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtGroupName.SetText .Fields("GroupName") & ":"
        End If
    End With
    With rsSystem
        .Find "SF='01'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            S = .Fields("Data")
        End If
        If ArrayFlag(S, 5) = 0 Then Section22.Suppress = True
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Section22_Format"

End Sub

Private Sub Section23_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    txtFoot5.SetText rscompany!Invoice_Notes_5
    If txtFoot5.Text = "" Then Section23.Suppress = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section23_Format"
End Sub

Private Sub Section25_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    With rsInventory
        .Find "ItemNum='" & txtPluName.Value & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If Len(.Fields("Modify_Number")) > 1 Then
                txtName2.SetText .Fields("Modify_Number")
            Else
                Section25.Suppress = True
            End If
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & " Section25_Format"
End Sub

Private Sub Section26_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If txtMoney.Value = 0 Then Section26.Suppress = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section26_Format"

End Sub

Private Sub Section27_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    txtCashName.SetText userName
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section28_Format"

End Sub

Private Sub Section28_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If txtVAT.Value = 0 Then Section28.Suppress = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section28_Format"

End Sub

Private Sub Section29_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    txtFoot4.SetText rscompany!Invoice_Notes_4
    If txtFoot4.Text = "" Then Section29.Suppress = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section29_Format"

End Sub



Private Sub Section3_Format(ByVal pFormattingInfo As Object)
Dim S As String
Dim FCName As String
Dim i As Integer
With rsSystem
    .Find "SF='04'", , adSearchForward, adBookmarkFirst
    If Not .EOF Then
        S = .Fields("Data")
    End If
End With
If Left(Right("00000000" & HexToBin(S), 8), 1) = 1 Then
    With rsFCRate
        For i = 1 To 25
            .Find "MediaID='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                S = .Fields("F")
                If Mid(Right("00000000" & HexToBin(S), 8), 2, 1) = 1 Then
                    FCName = .Fields("MediaName")
                    If CDbl("0" & .Fields("FCRate")) <> 0 Then
                        FCRate.SetText .Fields("SYMBOL") & Format(txtTotalAmt.Value / CDbl("0" & .Fields("FCRate")), "#,##0.00")
                        txtFC.SetText FCName
                        Exit Sub
                    End If

                End If
            End If
        Next
    Section3.Suppress = True
    End With
Else
    Section3.Suppress = True
End If
End Sub



Private Sub Section31_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle

    If rscompany.RecordCount > 0 Then
        Picture2.SetOleLocation (rscompany!Image)
    End If
    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & " kh«ng t×m thÊy file logo lín"
End Sub

Private Sub Section32_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    lblText1.SetText rscompany!Company_Info_1
    lblText2.SetText rscompany!Company_Info_2
    lblText3.SetText rscompany!Company_Info_3
    lblText4.SetText rscompany!Company_Info_4
    lblText5.SetText rscompany!Company_Info_5
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & " Load logo chu"
End Sub

Private Sub Section33_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    txtFoot3.SetText rscompany!Invoice_Notes_3
    If txtFoot3.Text = "" Then Section33.Suppress = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section33_Format"

End Sub

Private Sub Section34_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    txtFoot2.SetText rscompany!Invoice_Notes_2
    If txtFoot2.Text = "" Then Section34.Suppress = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section34_Format"

End Sub

Private Sub Section35_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If txtCustomerID.Value = "101" Then
        Section35.Suppress = True
    Else
        Dim rscust As New ADODB.Recordset
        Set rscust = Open_Table(cnData, "Customer")
        With rscust
            .Find "CustNum='" & txtCustomerID.Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtCustName.SetText .Fields("CustName")
            End If
            
        End With
        CloseRecordset rscust
    End If

Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Section35_Format"
End Sub

Private Sub Section36_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If LineDisc.Value = 0 Then Section36.Suppress = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section37_Format"
End Sub

Private Sub Section37_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If SumLineDiscAmt.Value = 0 Then Section37.Suppress = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section37_Format"
End Sub

Private Sub Section38_Format(ByVal pFormattingInfo As Object)
    On Error GoTo Handle
        If CDbl("0" & txtsokhach.Value) = 0 Then Section38.Suppress = True
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub Section39_Format(ByVal pFormattingInfo As Object)
    lblDateIn.SetText gfCONVERT_STRING_TO_DATE(txtDateIn.Value)
    lblDateOut.SetText gfCONVERT_STRING_TO_DATE(txtDateOut.Value)
End Sub

Private Sub Section40_Format(ByVal pFormattingInfo As Object)
       If ArrayFlag(SF(5), 6) = 1 Then Section40.Suppress = True
End Sub

Private Sub Section41_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If ArrayFlag(SF(5), 7) = 1 Then
        Section41.Suppress = True
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
MsgBox Err.Number & Err.Description & Me.name & " Section41_Format"

End Sub

Private Sub Section42_Format(ByVal pFormattingInfo As Object)
If ArrayFlag(SF(5), 5) = 1 Then Section42.Suppress = True
End Sub

Private Sub Section43_Format(ByVal pFormattingInfo As Object)
If txtReserved.Value = 0 Then Section43.Suppress = True
End Sub

Private Sub Section44_Format(ByVal pFormattingInfo As Object)
If txtFinal.Value = txtCash.Value Then Section44.Suppress = True
End Sub

Private Sub Section45_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle

If txtAdj5.Value = 0 Then Section45.Suppress = True
    With rsAdjustment
        .Find "AdjNo='05'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj5.SetText .Fields("AdjName")
        End If
    End With

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section45_Format"


End Sub

Private Sub Section46_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle

If txtAdj6.Value = 0 Then Section46.Suppress = True
    With rsAdjustment
        .Find "AdjNo='06'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj6.SetText .Fields("AdjName")
        End If
    End With

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section46_Format"

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
MsgBox Err.Number & Err.Description & Me.name & " Section5_Format"
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If rscompany.RecordCount > 0 Then
        lblInfor1.SetText rscompany!Company_Info_1
        lblInfor2.SetText rscompany!Company_Info_2
        lblInfor3.SetText rscompany!Company_Info_3
        lblInfor4.SetText rscompany!Company_Info_4 & "-" & rscompany!Company_Info_5
        Picture1.SetOleLocation (rscompany!Image)
    End If
Exit Sub
Handle:
    MsgBox Err.Number & " Kh«ng t×m thÊy file logo! Vµo CÊu h×nh hÖ thèng-->Th«ng tin ®Çu cuèi H§-->T.Tin ®ång bé d÷ liÖu--> KÝch vµo khung Logo h×nh--> chän file logo (*.bmp) -->OK"
End Sub

Private Sub Section9_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If txtCash.Value = TxtTotal.Value Then Section9.Suppress = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section37_Format"
End Sub
