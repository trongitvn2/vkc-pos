VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crSale_Location 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16425
   OleObjectBlob   =   "crSale_Location.dsx":0000
End
Attribute VB_Name = "crSale_Location"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim rsLocation As New ADODB.Recordset


Private Sub Section1_Format(ByVal pFormattingInfo As Object)
Dim rsStore As New ADODB.Recordset
If cnData.State = adStateClosed Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
Set rsStore = Open_Table(cnData, "Stations_Location")
With rsStore
    If .RecordCount > 0 Then .MoveFirst
    .Find "Station_Number='" & storeID.Value & "'", , adSearchForward, adBookmarkFirst
    If Not .EOF Then
        storeName.SetText .Fields("Station_Name")
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
        lblInfor4.SetText rscompany.Fields("Company_info_4") & Space(10) & "-" & rscompany.Fields("Company_info_5")
'        Picture1.SetOleLocation (rscompany!Image)
    End If
Exit Sub

Handle:

MsgBox "Bao loi" & Err.Number & " " & Err.Description
End Sub

