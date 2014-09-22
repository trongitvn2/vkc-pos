VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crTableTotal 
   ClientHeight    =   8205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13965
   OleObjectBlob   =   "crTableTotal.dsx":0000
End
Attribute VB_Name = "crTableTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim rsserver As New ADODB.Recordset
Dim rsStore As New ADODB.Recordset

Private Sub Report_Initialize()
On Error GoTo Handle
'If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    Set rsserver = Open_Table(cnData, "Table_Diagram_Sections")
    Set rsStore = Open_Table(cnData, "Stations_Location")
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub Report_Terminate()
    On Error GoTo Handle
       Set rsserver = Nothing
       Set rsStore = Nothing
    Exit Sub
Handle:
    MsgBox Me.name & " Report_Terminate " & Err.Number & " " & Err.Description
End Sub

Private Sub Section12_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    With rsserver
        If Not .EOF Then
            .Find "Location_ID='" & txtserver.Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                ServerName.SetText " - " & .Fields("Section_ID")
            End If
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section12_Format"
End Sub

Private Sub Section5_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    With rsStore
        If Not .EOF Then
            .Find "Station_Number='" & txtStoreID.Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtStoreName.SetText .Fields("Station_Name")
            End If
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "Section5_Format"
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
Dim rscompany As New ADODB.Recordset
    Set rscompany = Open_Table(cnData, "Setup")
    If rscompany.RecordCount > 0 Then
        lblInfor1.SetText rscompany.Fields("Company_info_1")
        lblInfor2.SetText rscompany.Fields("Company_info_2")
        lblInfor3.SetText rscompany.Fields("Company_info_3")
        lblInfor4.SetText rscompany.Fields("Company_info_4") & Space(10) & rscompany.Fields("Company_info_5")
'        Picture1.SetOleLocation (rscompany!Image)
    End If
Exit Sub
Handle:
    MsgBox "Bao loi" & Err.Number & " " & Err.Description
End Sub

Private Sub Section9_Format(ByVal pFormattingInfo As Object)
    txtReadValue.SetText readnumber(CDbl("0" & txtsumAmt.Value)) & " ®ång ./."
End Sub
