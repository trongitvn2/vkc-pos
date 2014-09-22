VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crTranfer58 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11325
   OleObjectBlob   =   "crTranfer58.dsx":0000
End
Attribute VB_Name = "crTranfer58"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim rsSection As New ADODB.Recordset
Dim rsCashier As New ADODB.Recordset

Private Sub Report_Initialize()
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsSection = Open_Table(cnData, "Table_Diagram_Sections")
    Set rsCashier = LoadPasswordData
End Sub


Private Sub Report_Terminate()
    Set crNewBalance = Nothing
    Set rsSection = Nothing
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If rsSection.State <> 0 Then
        With rsSection
            .Find "Location_ID='" & Location.Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                lblLocationID.SetText .Fields("Section_ID")
            End If
        End With
        
        With rsSection
            If .RecordCount > 0 Then .MoveFirst
            .Find "Location_ID='" & LocationDes.Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                lblLocationDes.SetText .Fields("Section_ID")
            End If
        End With
        
    End If
    If rsCashier.State <> 0 Then
            With rsCashier
                .Find "ID='" & CashierID.Value & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    txtCashierName.SetText .Fields("userName")
                End If
        End With
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub Section9_Format(ByVal pFormattingInfo As Object)
    On Error GoTo Handle
    
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Section9_Format"
End Sub
