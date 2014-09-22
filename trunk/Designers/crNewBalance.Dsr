VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crNewBalance 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10770
   OleObjectBlob   =   "crNewBalance.dsx":0000
End
Attribute VB_Name = "crNewBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"CrystalReport"
Option Explicit
Dim rsSection As New ADODB.Recordset

Private Sub Report_Initialize()
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsSection = Open_Table(cnData, "Table_Diagram_Sections")
End Sub


Private Sub Report_Terminate()
    Set crNewBalance = Nothing
    Set rsSection = Nothing
End Sub

Private Sub Section1_Format(ByVal pFormattingInfo As Object)
    If txtKitDesc.Value = "" Or txtKitDesc.Value = "-" Then Section1.Suppress = True
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle

Dim rsInventory As New ADODB.Recordset
Set rsInventory = Open_Table(cnData, "Inventory")
With rsInventory
    .Find "ItemNum='" & ItemNum.Value & "'", , adSearchForward, adBookmarkFirst
    If Not .EOF Then
        If ArrayFlag(.Fields("F1"), 3) = 1 Then
            With Qty
                .DecimalPlaces = DecimalQtyNumber
                .DecimalSymbol = DecimalMark
                .ThousandsSeparators = True
                .ThousandSymbol = DigitGroupMark
            End With
        Else
            With Qty
                .DecimalPlaces = 0
                .DecimalSymbol = DecimalMark
                .ThousandsSeparators = True
                .ThousandSymbol = DigitGroupMark
            End With
        End If
    End If
End With

    If Qty.Value < 0 Then
        Qty.Font.Size = 10
        Qty.Font.Bold = True
        Items.Font.Size = 11
        Items.Font.Bold = True
        Price.Font.Size = 10
        Price.Font.Bold = True
        Items.Font.Strikethrough = True
        
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section10_Format "
End Sub

Private Sub Section2_Format(ByVal pFormattingInfo As Object)
    If txtsokhach.Value = 0 Then Section2.Suppress = True
End Sub


Private Sub Section6_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    If rsSection.State <> 0 Then
        With rsSection
            .Find "Location_ID='" & Location.Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                lblLocationID.SetText .Fields("Section_ID")
            End If
        End With
    End If
    If rsuser.State <> 0 Then
        With rsuser
            .Find "ID='" & Cashier.Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtCashierName.SetText .Fields("userName")
            End If
        End With
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub
