VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crLocation 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10260
   OleObjectBlob   =   "crLocation.dsx":0000
End
Attribute VB_Name = "crLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim rsLocation As New ADODB.Recordset

Dim rscash As New ADODB.Recordset

Private Sub Report_Initialize()
On Error GoTo Handle
'If cnData.State = 0 Then
'    cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'End If
    Set rsLocation = Open_Table(cnData, "Table_Diagram_Sections")
    Set rscash = LoadPasswordData
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Report_Initialize"
End Sub

Private Sub Report_Terminate()
    On Error GoTo Handle
       Set rsLocation = Nothing
    Exit Sub
Handle:
    MsgBox Me.name & " Report_Terminate " & Err.Number & " " & Err.Description
End Sub

Private Sub Section5_Format(ByVal pFormattingInfo As Object)


End Sub

'Private Sub Section1_Format(ByVal pFormattingInfo As Object)
'With rscash
'        .Find "ID='" & Cashier.Value & "'", , adSearchForward, adBookmarkFirst
'        If Not .EOF Then
'            txtName.SetText .Fields("UserName")
'        End If
'    End With
'End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
Dim rscompany As New ADODB.Recordset
    Set rscompany = Open_Table(cnData, "Setup")
    If rscompany.RecordCount > 0 Then
        lblInfor1.SetText rscompany.Fields("Company_info_1")
        lblInfor2.SetText rscompany.Fields("Company_info_2")
        lblInfor3.SetText "Add" & rscompany.Fields("Company_info_3")
        lblInfor4.SetText "Tel:" & rscompany.Fields("Company_info_4") & Space(10) & "Fax:" & rscompany.Fields("Company_info_5")
'        Picture1.SetOleLocation (rscompany!Image)
    End If
Exit Sub

Handle:

MsgBox "Bao loi" & Err.Number & " " & Err.Description
End Sub

