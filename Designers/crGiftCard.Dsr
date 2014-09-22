VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crGiftCard 
   ClientHeight    =   9360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   OleObjectBlob   =   "crGiftCard.dsx":0000
End
Attribute VB_Name = "crGiftCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim rsCashier As New ADODB.Recordset
Dim rscompany As New ADODB.Recordset

Private Sub Report_Initialize()
    Set rsCashier = LoadPasswordData
End Sub

Private Sub Report_Terminate()
    On Error GoTo Handle
        Set rscompany = Nothing
        Set rsCashier = Nothing
    Exit Sub
Handle:
    MsgBox Me.Name & " Report_Terminate " & Err.Number & " " & Err.Description
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    Set rscompany = Open_Table(cnData, "Setup")
    If rscompany.RecordCount > 0 Then
        lblInfor1.SetText rscompany.Fields("Company_info_1")
        lblInfor2.SetText rscompany.Fields("Company_info_2")
    End If
    lblDateExpired.SetText gfCONVERT_STRING_TO_DATE(txtDateExpired.Value)
    lblDateOpen.SetText "Ngµy " & Day(gfCONVERT_STRING_TO_DATE(txtDateOpen.Value)) & " Th¸ng " & Month(gfCONVERT_STRING_TO_DATE(txtDateOpen.Value)) & " N¨m " & Year(gfCONVERT_STRING_TO_DATE(txtDateOpen.Value))
     txtRead.SetText readnumber(CDbl("0" & txtAmount.Value)) & " ®ång"
     txtRead1.SetText readnumber(CDbl("0" & txtAmount.Value)) & " ®ång"
     amount.SetText Format(txtAmount2.Value, "#,###") & "§"
     barcode.SetText txtCardID.Value
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Section10_Format"
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    
Exit Sub
Handle:

MsgBox "Bao loi" & Err.Number & " " & Err.Description
End Sub

