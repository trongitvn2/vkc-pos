VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crStockMoveInOut 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10770
   OleObjectBlob   =   "crStockMoveInOut.dsx":0000
End
Attribute VB_Name = "crStockMoveInOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim myColor As Long
Dim rscompany As New ADODB.Recordset
Dim rsInventory As New ADODB.Recordset
Dim rsMaterial As New ADODB.Recordset

Private Sub Report_Initialize()
    Set rsMaterial = Open_Table(cnData, "SetMPLu")
    Set rsInventory = Open_Table(cnData, "Inventory")
End Sub

Private Sub Report_Terminate()
Set rscompany = Nothing
Set rsInventory = Nothing
Set rsMaterial = Nothing
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
  On Error GoTo errHdl
    If LastQty.Value < 0 Then
        myColor = mRed
    Else
        myColor = mWhite
    End If
    
    Section10.BackColor = myColor
With rsInventory
    .Find "ItemNum='" & PluCode.Value & "'", , adSearchForward, adBookmarkFirst
    If Not .EOF Then
        lblName.SetText .Fields("ItemName")
    Else
        With rsMaterial
            .Find "PluCode='" & PluCode.Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                lblName.SetText .Fields("PluName")
            End If
        End With
    End If
End With
    Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - Section10_Format"
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
