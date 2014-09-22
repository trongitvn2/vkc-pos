VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crDeleteItems 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   OleObjectBlob   =   "crDeleteItems.dsx":0000
End
Attribute VB_Name = "crDeleteItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim rscompany As New ADODB.Recordset

Private Sub Report_Terminate()
    On Error GoTo Handle
        Set rscompany = Nothing
    Exit Sub
Handle:
    MsgBox Me.name & " Report_Terminate " & Err.Number & " " & Err.Description
End Sub

Private Sub Section1_Format(ByVal pFormattingInfo As Object)
    txtDateGroup.SetText gfCONVERT_STRING_TO_DATE(DateGroup.Value)
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
Dim rsInventory As New ADODB.Recordset

'If cnData.State = 0 Then
'    Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'End If
Set rsInventory = Open_Table(cnData, "Inventory")
If rsInventory.State <> 0 Then
    With rsInventory
        .Find "ItemNum='" & txtPluCode.Value & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            
            If ArrayFlag(.Fields("F1"), 3) = 1 Then
                With txtQty
                    .DecimalPlaces = DecimalQtyNumber
                    .DecimalSymbol = DecimalMark
                    .ThousandsSeparators = True
                    .ThousandSymbol = DigitGroupMark
                End With
            Else
                With txtQty
                    .DecimalPlaces = 0
                    .DecimalSymbol = DecimalMark
                    .ThousandsSeparators = True
                    .ThousandSymbol = DigitGroupMark
                End With
            End If
        End If
    End With
End If

If txtCashier.Value = "131112" Then
    txtCashName.SetText "PTV Administrator"
Else
    With rsuser
        .Find "ID='" & txtCashier.Value & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtCashName.SetText .Fields("UserName")
        End If
    End With
End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Section10_Format"
End Sub


Private Sub Section12_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
Dim rsserver As New ADODB.Recordset
    
Set rsserver = Open_Table(cnData, "Table_Diagram_Sections")
With rsserver
    If Not .EOF Then
        .Find "Location_ID='" & txtserver.Value & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            ServerName.SetText .Fields("Section_ID")
        End If
    End If
End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - Section12_Format"
End Sub

Private Sub Section3_Format(ByVal pFormattingInfo As Object)
If printcount.Value = "" Or printcount.Value = 0 Then
    Section3.Suppress = True
End If
End Sub

Private Sub Section5_Format(ByVal pFormattingInfo As Object)
    If GroupOrder.Value = True Then
        txtOrdered.Suppress = True
        txtOrdered1.Suppress = False
    Else
        txtOrdered.Suppress = False
        txtOrdered1.Suppress = True
    End If
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    Set rscompany = Open_Table(cnData, "Setup")
    If rscompany.RecordCount > 0 Then
        lblInfor1.SetText rscompany.Fields("Company_info_1")
        lblInfor2.SetText rscompany.Fields("Company_info_2")
        lblInfor3.SetText "Add:" & rscompany.Fields("Company_info_3")
        lblInfor4.SetText "Tel :" & rscompany.Fields("Company_info_4") & Space(10) & "Fax:" & rscompany.Fields("Company_info_5")
'        Picture1.SetOleLocation (rscompany!Image)
    End If
Exit Sub
Handle:

MsgBox "Bao loi" & Err.Number & " " & Err.Description
End Sub

Private Sub Section9_Format(ByVal pFormattingInfo As Object)
    txtReadTotal.SetText readnumber(CDbl("0" & txtTotalAmt.Value)) & " ®ång./."
End Sub
