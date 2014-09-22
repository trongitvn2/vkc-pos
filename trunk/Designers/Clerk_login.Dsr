VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crClerkLogin 
   ClientHeight    =   9480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9585
   OleObjectBlob   =   "Clerk_login.dsx":0000
End
Attribute VB_Name = "crClerkLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Section10_Format(ByVal pFormattingInfo As Object)
    Select Case InOutType.Value
        Case "I"
            lblInOutType.SetText "Vµo"
        Case "O"
            lblInOutType.SetText "Ra"
    End Select

End Sub

Private Sub Section6_Format(ByVal pFormattingInfo As Object)
    Select Case InOutType.Value
        Case "I"
            InOutTitle.SetText "Vµo"
        Case "O"
            InOutTitle.SetText "Ra"
    End Select
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo handle
Dim rscompany As New ADODB.Recordset
    Set rscompany = Open_Table(cnData, "Setup")
    If rscompany.RecordCount > 0 Then
        lblInfor1.SetText rscompany.Fields("Company_info_1")
        lblInfor2.SetText rscompany.Fields("Company_info_2")
        lblInfor3.SetText rscompany.Fields("Company_info_3")
        lblInfor4.SetText rscompany.Fields("Company_info_4") & Space(2) & rscompany.Fields("Company_info_5")
'        Picture1.SetOleLocation (rscompany!Image)
    End If
Exit Sub

handle:

MsgBox "Bao loi - kh«ng t×m thÊy file Logo h×nh" & Err.Number & " " & Err.Description

End Sub
