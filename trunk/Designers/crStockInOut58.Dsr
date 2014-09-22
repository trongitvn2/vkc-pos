VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crStockInOut58 
   ClientHeight    =   9960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9690
   OleObjectBlob   =   "crStockInOut58.dsx":0000
End
Attribute VB_Name = "crStockInOut58"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CntLine As Integer
Dim myColor As Long
Dim rscompany As New ADODB.Recordset
Dim rsMPlu As New ADODB.Recordset
Dim rsPLU As New ADODB.Recordset


Private Sub Report_BeforeFormatPage(ByVal PageNumber As Long)
  On Error GoTo errHdl
    CntLine = 1
  
  Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - Report_BeforeFormatPage"
  
End Sub

Private Sub Report_Initialize()
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rscompany = Open_Table(cnData, "Setup")
    Set rsMPlu = Open_Table(cnData, "SetMPLU")
    Set rsPLU = Open_Table(cnData, "Inventory")
End Sub

Private Sub Report_Terminate()
    Set rscompany = Nothing
End Sub

Private Sub Section1_Format(ByVal pFormattingInfo As Object)
 Dim rsStock As New ADODB.Recordset
    Set rsStock = Open_Table(cnData, "Stock_List")
    With rsStock
        .Find "ID='" & StockID.Value & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            StockName.SetText .Fields("Stock_Name")
        End If
    End With
    
    Exit Sub
Handle:
        MsgBox Err.Number & Err.Description & _
        Me.name & " Section1_Format"
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
  On Error GoTo errHdl
    
'    If CntLine Mod 2 = 0 Then
'        myColor = mLightGrey
'    Else
        myColor = mWhite
'    End If
    If Qty.Value < 0 Then
        myColor = mRed
    End If
    Section10.BackColor = myColor
    CntLine = CntLine + 1
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - Section10_Format"

End Sub

Private Sub Section11_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    Dim rsDept As New ADODB.Recordset
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    Set rsDept = Open_Table(cnData, "Departments")
    With rsDept
        If Not rsDept.EOF And .RecordCount > 0 Then .MoveFirst
       .Find "Dept_ID='" & GroupA.Value & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            groupName.SetText " - " & .Fields("Description")
        End If
    End With
        
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Section11_Format"
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo Handle
    With rscompany
        txtCompany.SetText !Company_Info_1
        txtAdd.SetText !Company_Info_2
        txtPhone.SetText !Company_Info_3 & " - " & !Company_Info_4
'        Picture1.SetOleLocation (rscompany!Image)
    End With
Exit Sub

Handle:

MsgBox "Kh«ng t×m thÊy file Logo" & Err.Number & " " & Err.Description
End Sub

