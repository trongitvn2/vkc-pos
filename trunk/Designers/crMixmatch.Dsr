VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crMixmatch 
   ClientHeight    =   9960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10680
   OleObjectBlob   =   "crMixmatch.dsx":0000
End
Attribute VB_Name = "crMixmatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rscompany As New ADODB.Recordset
Dim rscust As New ADODB.Recordset

Private Sub Report_Initialize()
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    Set rscompany = Open_Table(cnData, "setup")
    Set rscust = Open_Table(cnData, "Customer")
End Sub

Private Sub Report_Terminate()
    CloseRecordset rscompany
    CloseRecordset rscust
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
Dim rsMixmatch As New ADODB.Recordset
Dim rsuser As New ADODB.Recordset
    Set rsuser = LoadPasswordData
    With rsuser
        .Find "ID='" & Format(txtCashierID.Value, "00") & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblCashier.SetText .Fields("UserName")
        End If
    End With
    With rscust
        .Find "CustNum='" & txtCust.Value & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtCustName.SetText .Fields("CustName")
        End If
    End With
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

