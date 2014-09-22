VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crPhieuchi58 
   ClientHeight    =   9135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15420
   OleObjectBlob   =   "crPhieuChi58.dsx":0000
End
Attribute VB_Name = "crPhieuchi58"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim rsChi As New ADODB.Recordset

Private Sub Report_Initialize()
    On Error GoTo Handle
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsChi = OpenCriticalTable("select * from Expense", cnData)
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Report_Initialize "
End Sub

Private Sub Report_Terminate()
    Set rsChi = Nothing
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
    On Error GoTo Handle
                With rsChi
            .Find "MaChi='" & txtMathu.Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtDiengiai.SetText .Fields("Diengiai")
            End If
        End With
        Bangchu.SetText readnumber(txtSotien.Value) & " ®ång ./."
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Section10_Format"
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
Dim rscompany As New ADODB.Recordset
    Set rscompany = Open_Table(cnData, "Setup")
    If rscompany.RecordCount > 0 Then
        lblInfor1.SetText rscompany.Fields("Company_info_1")
        lblInfor2.SetText rscompany.Fields("Company_info_2")
        lblInfor3.SetText "Ñc:" & rscompany.Fields("Company_info_3")
        lblInfor4.SetText "ÑT :" & rscompany.Fields("Company_info_4") & Space(10) & "Fax:" & rscompany.Fields("Company_info_5")
'        Picture1.SetOleLocation (rscompany!Image)
    End If
    txtNgay.SetText gfCONVERT_STRING_TO_DATE(txtNgaythu.Value)
Exit Sub

Handle:

MsgBox "Bao loi" & Err.Number & " " & Err.Description
    
End Sub

Private Sub Section9_Format(ByVal pFormattingInfo As Object)
    txtDate.SetText "Ngµy " & Right(DateDefault, 2) & " Th¸ng " & Mid(DateDefault, 5, 2) & " N¨m " & Left(DateDefault, 4)
End Sub
