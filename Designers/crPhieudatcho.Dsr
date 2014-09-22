VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crPhieudatcho 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14970
   OleObjectBlob   =   "crPhieudatcho.dsx":0000
End
Attribute VB_Name = "crPhieudatcho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
    Option Explicit

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
    Dim rsInventory As ADODB.Recordset
    Set rsInventory = Open_Table(cnData, "Inventory")
    With rsInventory
        .Find "ItemNum='" & txtmahang.Value & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            dvt.SetText .Fields("Unit")
        End If
    End With
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
Dim rscompany As New ADODB.Recordset
    Set rscompany = Open_Table(cnData, "Setup")
    If rscompany.RecordCount > 0 Then
        lblInfor1.SetText rscompany.Fields("Company_info_1")
        lblInfor2.SetText rscompany.Fields("Company_info_2")
        lblInfor3.SetText "Add:" & rscompany.Fields("Company_info_3")
        lblInfor4.SetText "Tel :" & rscompany.Fields("Company_info_4") & "Fax:" & rscompany.Fields("Company_info_5")
'        Picture1.SetOleLocation (rscompany!Image)
    End If
Exit Sub
Handle:
    MsgBox "Bao loi" & Err.Number & " " & Err.Description
End Sub
'
Private Sub Section6_Format(ByVal pFormattingInfo As Object)
    txtbangchu.SetText readnumber(txtSotien.Value) & " ®ång"
End Sub
