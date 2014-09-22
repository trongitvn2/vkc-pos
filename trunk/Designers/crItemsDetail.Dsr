VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crItemsDetail 
   ClientHeight    =   9300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9810
   OleObjectBlob   =   "crItemsDetail.dsx":0000
End
Attribute VB_Name = "crItemsDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim rscompany As New ADODB.Recordset
Dim rsStore As New ADODB.Recordset
Dim rsLocation As New ADODB.Recordset


Private Sub Report_Initialize()
On Error GoTo Handle
'If cnData.State = 0 Then
'    cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'End If
    Set rsStore = Open_Table(cnData, "Stations_Location")
    Set rsLocation = Open_Table(cnData, "Table_Diagram_Sections")
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Report_Initialize"

End Sub

Private Sub Report_Terminate()
    On Error GoTo Handle
        Set rscompany = Nothing
        Set rsStore = Nothing
        Set rsLocation = Nothing
    Exit Sub
Handle:
    MsgBox Me.name & " Report_Terminate " & Err.Number & " " & Err.Description
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    Set rscompany = Open_Table(cnData, "Setup")
    If rscompany.RecordCount > 0 Then
        lblInfor1.SetText rscompany.Fields("Company_info_1")
        lblInfor2.SetText rscompany.Fields("Company_info_2")
        lblInfor3.SetText "§/C:" & rscompany.Fields("Company_info_3")
        lblInfor4.SetText "§T:" & rscompany.Fields("Company_info_4") & Space(10) & "Fax:" & rscompany.Fields("Company_info_5")
'        Picture1.SetOleLocation (rscompany!Image)
    End If
Exit Sub
Handle:

MsgBox "Bao loi" & Err.Number & " " & Err.Description
End Sub

