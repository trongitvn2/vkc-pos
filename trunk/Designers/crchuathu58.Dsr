VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crBanchuathu58 
   ClientHeight    =   9300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10740
   OleObjectBlob   =   "crchuathu58.dsx":0000
End
Attribute VB_Name = "crBanchuathu58"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim rsuser As New ADODB.Recordset

Private Sub Report_Initialize()
    Set rsuser = LoadPasswordData
End Sub

Private Sub Report_Terminate()
    On Error GoTo Handle
       
    Exit Sub
Handle:
    MsgBox Me.Name & " Report_Terminate " & Err.Number & " " & Err.Description
End Sub

Private Sub Section1_Format(ByVal pFormattingInfo As Object)

End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
Dim rsserver As New ADODB.Recordset

Set rsserver = Open_Table(cnData, "Table_Diagram_Sections")
With rsserver
    If Not .EOF Then
        .Find "Location_ID='" & txtserver.Value & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            ServerName.SetText .Fields("Section_ID")
        End If
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
        lblInfor3.SetText rscompany.Fields("Company_info_3")
        lblInfor4.SetText rscompany.Fields("Company_info_4") & Space(2) & rscompany.Fields("Company_info_5")
'        Picture1.SetOleLocation (rscompany!Image)
    End If
Exit Sub

Handle:

MsgBox "Bao loi" & Err.Number & " " & Err.Description
End Sub

