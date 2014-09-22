VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crCashierReport80 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16620
   OleObjectBlob   =   "crCashierReport80.dsx":0000
End
Attribute VB_Name = "crCashierReport80"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim rsCashier As New ADODB.Recordset
Dim rscompany As New ADODB.Recordset
Dim DescArrDis() As String


Private Sub Report_Initialize()
    Set rsCashier = LoadPasswordData
    DescArrDis = LoadLanguage(LngFile, "#01:012:")

End Sub

Private Sub Report_Terminate()
    On Error GoTo Handle
        Set rscompany = Nothing
        Set rsCashier = Nothing
    Exit Sub
Handle:
    MsgBox Me.name & " Report_Terminate " & Err.Number & " " & Err.Description
End Sub

Private Sub Section1_Format(ByVal pFormattingInfo As Object)
    On Error GoTo Handle
        With rsCashier
            .Find "ID='" & Field1.Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                CashierName.SetText " - " & .Fields("userName")
            End If
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Section1_Format"
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    lblAdj1.SetText DescArrDis(12)
    lblAdj2.SetText DescArrDis(13)
    lblAdj3.SetText DescArrDis(14)
    lblAdj4.SetText DescArrDis(15)
    lblAdj5.SetText DescArrDis(16)
    lblAdj6.SetText DescArrDis(17)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    Set rscompany = Open_Table(cnData, "Setup")
    If rscompany.RecordCount > 0 Then
        lblInfor1.SetText rscompany.Fields("Company_info_1")
        lblInfor2.SetText rscompany.Fields("Company_info_2")
        lblInfor3.SetText rscompany.Fields("Company_info_3")
        lblInfor4.SetText rscompany.Fields("Company_info_4") & "-" & rscompany.Fields("Company_info_5")
'        Picture1.SetOleLocation (rscompany!Image)
    End If
Exit Sub
Handle:

MsgBox "Bao loi" & Err.Number & " " & Err.Description
End Sub

Private Sub Section9_Format(ByVal pFormattingInfo As Object)
    txtReadTotal.SetText readnumber(CDbl("0" & sumTotal.Value)) & " ®ång./."
End Sub
