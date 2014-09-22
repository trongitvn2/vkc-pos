VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crMaterial 
   ClientHeight    =   9960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9690
   OleObjectBlob   =   "crMaterial.dsx":0000
End
Attribute VB_Name = "crMaterial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rscompany As New ADODB.Recordset

Private Sub Section11_Format(ByVal pFormattingInfo As Object)
    CostPLU.SetText Format(CDbl(Price.Value), "#,##0")
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)

On Error GoTo Handle
    Set rscompany = Open_Table(cnData, "Setup")
    If rscompany.RecordCount > 0 Then
        lblInfor1.SetText rscompany.Fields("Company_info_1")
        lblInfor2.SetText rscompany.Fields("Company_info_2")
'        lblInfor3.SetText "§/C:" & rscompany.Fields("Company_info_3")
'        lblInfor4.SetText "§T:" & rscompany.Fields("Company_info_4") & Space(10) & "Fax:" & rscompany.Fields("Company_info_5")
'        Picture1.SetOleLocation (rscompany!Image)
    End If
Exit Sub
Handle:

MsgBox "Bao loi" & Err.Number & " " & Err.Description
End Sub
