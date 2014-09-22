VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crBillList 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15315
   OleObjectBlob   =   "BillList.dsx":0000
End
Attribute VB_Name = "crBillList"
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

Private Sub Report_Terminate()
    Set rsserver = Nothing
    Set rsFCRate = Nothing
    Set rsAdjustment = Nothing
    Set rsMainGroup = Nothing
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

End Sub

Private Sub Section1_Format(ByVal pFormattingInfo As Object)
   
    txtFooter.SetText rscompany!Invoice_Notes_1 & "-" & rscompany!Invoice_Notes_2
    txtNote.SetText rscompany!Invoice_Notes_3
    On Error GoTo Handle
        
    Exit Sub

Handle:
    MsgBox "B¸o lçi _CrSaleBill_Section1_format"
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle

Dim rsInventory As New ADODB.Recordset
Set rsInventory = Open_Table(cnData, "Inventory")
With rsInventory
    .Find "ItemNum='" & PluName.Value & "'", , adSearchForward, adBookmarkFirst
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


Private Sub Section11_Format(ByVal pFormattingInfo As Object)
If txtReceiveMoney.Value = 0 Then Section11.Suppress = True
End Sub

Private Sub Section12_Format(ByVal pFormattingInfo As Object)
If txtSerCharge.Value = 0 Then Section12.Suppress = True
End Sub

Private Sub Section13_Format(ByVal pFormattingInfo As Object)
If txtDiscount.Value = 0 Then Section13.Suppress = True
End Sub


Private Sub Section16_Format(ByVal pFormattingInfo As Object)
    If txtPayment.Value = txtsumAmt.Value Then Section16.Suppress = True
End Sub

Private Sub Section18_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If txtCustomer.Value = "101" Or txtCustomer.Value = "" Then
        Section18.Suppress = True
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
MsgBox Err.Number & Err.Description & Me.name & "  Section36_Format"

End Sub

Private Sub Section19_Format(ByVal pFormattingInfo As Object)
    Dim rscash As New ADODB.Recordset
    Set rscash = LoadPasswordData
    With rscash
        .Find "ID='" & txtCashier.Value & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtCashName.SetText .Fields("UserName")
        End If
    End With
End Sub

Private Sub Section20_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
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
            Section20.Suppress = True
        End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section34_Format-Hien thi nhan vien phuc vu "

End Sub

Private Sub Section21_Format(ByVal pFormattingInfo As Object)
    txtRead.SetText readnumber(txtPayment.Value) & " ®ång./."
End Sub

Private Sub Section3_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    With rsMainGroup
        .Find "GroupNo='" & Trim$(txtMaingroup.Value) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblMainGroup.SetText .Fields("GroupName") & ":"
        End If
    End With

        If ArrayFlag(SF(0), 5) = 0 Then Section3.Suppress = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Section27_Format"

End Sub



Private Sub Section31_Format(ByVal pFormattingInfo As Object)
    With rsserver
        .Find "Location_ID='" & txtserver.Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                lblTextSever.SetText .Fields("Section_ID")
                lblTable.SetText DescArr(3)
            End If
    End With
    lblDatetime.SetText gfCONVERT_STRING_TO_DATE(txtDate.Value)
End Sub

Private Sub Section32_Format(ByVal pFormattingInfo As Object)
    lblItems.SetText DescArr(4)
    lblQty.SetText DescArr(5)
    lblPrice.SetText DescArr(6)
    lblAmt.SetText DescArr(7)
End Sub


Private Sub Section4_Format(ByVal pFormattingInfo As Object)
    If txtAdjustment2.Value = 0 Then Section4.Suppress = True
End Sub


Private Sub Section5_Format(ByVal pFormattingInfo As Object)
    If txtAdjustment1.Value = 0 Then Section5.Suppress = True
End Sub



Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
Dim i As Integer
    
    With rsserver
        .Find "Location_ID='" & txtserver.Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                lblTextSever.SetText .Fields("Section_ID")
                lblTable.SetText DescArr(3)
            End If
    End With
    
Exit Sub
Handle:

MsgBox Err.Number & " Kh«ng t×m thÊy file logo! Vµo CÊu h×nh hÖ thèng-->Th«ng tin ®Çu cuèi H§-->T.Tin ®ång bé d÷ liÖu--> KÝch vµo khung Logo h×nh--> chän file logo (*.bmp) -->OK"
End Sub


