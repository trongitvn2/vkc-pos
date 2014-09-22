VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crStockCard 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14970
   OleObjectBlob   =   "crStockCard.dsx":0000
End
Attribute VB_Name = "crStockCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsInventory As New ADODB.Recordset
Dim rscompany As New ADODB.Recordset

Private Sub Report_Initialize()
    Set rsInventory = Open_Table(cnData, "Inventory")
End Sub
Private Sub Report_Terminate()
    Set rscompany = Nothing
    Set rsInventory = Nothing
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
    On Error GoTo Handle
        With rsInventory
            .Find "ItemNum='" & txtCode.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                Unit.SetText .Fields("Unit")
            End If
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " Section8_Format "
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    Set rscompany = Open_Table(cnData, "Setup")
    If rscompany.RecordCount > 0 Then
        txtinfor1.SetText rscompany.Fields("Company_info_1")
        txtinfor2.SetText rscompany.Fields("Company_info_2")
        txtinfor3.SetText "Add:" & rscompany.Fields("Company_info_3")
        txtinfor4.SetText "Tel:" & rscompany.Fields("Company_info_4") & Space(10) & "Fax:" & rscompany.Fields("Company_info_5")
'        Picture1.SetOleLocation (rscompany!Image)
    End If
Exit Sub
Handle:

MsgBox "Bao loi" & Err.Number & " Section8_Format " & Err.Description

    
End Sub
