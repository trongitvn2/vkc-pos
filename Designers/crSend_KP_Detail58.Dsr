VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crSend_KP_Detail58 
   ClientHeight    =   9480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9585
   OleObjectBlob   =   "crSend_KP_Detail58.dsx":0000
End
Attribute VB_Name = "crSend_KP_Detail58"
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

Private Sub Section1_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    With rsStore
        If .RecordCount > 0 Then
            .Find "Station_Number='" & txtStoreID.Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtStoreName.SetText .Fields("Station_Name")
            End If
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section1_Format"

End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Section10_Format"
End Sub

Private Sub Section11_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    Dim rsCashier As New ADODB.Recordset
    Set rsCashier = LoadPasswordData
    With rsCashier
        .Find "ID='" & txtCashierID.Value & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtCashierName.SetText rsCashier.Fields("UserName")
        Else
            txtCashierName.SetText "Admin"
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section11_Format"
End Sub

Private Sub Section3_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    With rsLocation
        If .RecordCount > 0 Then
            .Find "Location_ID='" & txtStationID.Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtStationName.SetText .Fields("Section_ID")
            End If
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section3_Format"

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
    txtReadTotal.SetText readnumber(CDbl("0" & txtTotalAmt.Value)) & " ®ång./."
End Sub
