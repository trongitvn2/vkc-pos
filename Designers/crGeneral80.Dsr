VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crGeneral80 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10590
   OleObjectBlob   =   "crGeneral80.dsx":0000
End
Attribute VB_Name = "crGeneral80"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim rscompany As New ADODB.Recordset
Dim rsAdjustment As New ADODB.Recordset

Private Sub Report_Initialize()
'If cnData.State = 0 Then
'    Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'End If
 Set rsAdjustment = Open_Table(cnData, "Adjustment")

End Sub

Private Sub Report_Terminate()
    On Error GoTo Handle
        Set rscompany = Nothing
        Set rsAdjustment = Nothing
    Exit Sub
Handle:
    MsgBox Me.name & " Report_Terminate " & Err.Number & " " & Err.Description
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Section10_Format"
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
If CDbl("0" & Abs(txtAmountGiftCard.Value)) = 0 Then Section18.Suppress = True
End Sub

Private Sub Section11_Format(ByVal pFormattingInfo As Object)
If Val("0" & Abs(txtDelNotAmt.Value)) = 0 Then Section11.Suppress = True
End Sub

Private Sub Section13_Format(ByVal pFormattingInfo As Object)
If CDbl(txtDeleteAmt.Value) = 0 Then Section13.Suppress = True
End Sub



Private Sub Section2_Format(ByVal pFormattingInfo As Object)
If CDbl("0" & Abs(txtDisAmt.Value)) = 0 Then Section2.Suppress = True
End Sub

Private Sub Section20_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
If adjAmt1.Value = 0 Then Section20.Suppress = True
    With rsAdjustment
        .Find "AdjNo='01'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj1.SetText .Fields("AdjName")
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub Section21_Format(ByVal pFormattingInfo As Object)
    If txtCountLineDisc.Value = 0 Then Section21.Suppress = True
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
    If txtSokhach.Value = 0 Then Section24.Suppress = True
End Sub

Private Sub Section25_Format(ByVal pFormattingInfo As Object)
    If txtAmountReceiveMoney.Value = 0 Then Section25.Suppress = True
End Sub


Private Sub Section26_Format(ByVal pFormattingInfo As Object)
If CDbl("0" & txtAmountROA.Value) = 0 Then Section26.Suppress = True
End Sub

Private Sub Section27_Format(ByVal pFormattingInfo As Object)
    If txtVATAmount.Value = 0 Then Section27.Suppress = True
End Sub

Private Sub Section28_Format(ByVal pFormattingInfo As Object)
If txtAmountReserve.Value = 0 Then Section28.Suppress = True
End Sub

Private Sub Section29_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
If AdjAmt6.Value = 0 Then Section29.Suppress = True
    With rsAdjustment
        .Find "AdjNo='06'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj6.SetText .Fields("AdjName")
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub Section30_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
If AdjAmt5.Value = 0 Then Section30.Suppress = True
    With rsAdjustment
        .Find "AdjNo='05'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj5.SetText .Fields("AdjName")
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub Section31_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
If AdjAmt4.Value = 0 Then Section31.Suppress = True
    With rsAdjustment
        .Find "AdjNo='04'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj4.SetText .Fields("AdjName")
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub Section32_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
If AdjAmt3.Value = 0 Then Section32.Suppress = True
    With rsAdjustment
        .Find "AdjNo='03'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblAdj3.SetText .Fields("AdjName")
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
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
        lblInfor3.SetText "Add:" & rscompany.Fields("Company_info_3")
        lblInfor4.SetText "Tel:" & rscompany.Fields("Company_info_4") & Space(10) & "Fax:" & rscompany.Fields("Company_info_5")
'        Picture1.SetOleLocation (rscompany!Image)
    End If
Exit Sub
Handle:

MsgBox "Bao loi" & Err.Number & " " & Err.Description
End Sub

Private Sub Section9_Format(ByVal pFormattingInfo As Object)
   '
End Sub
