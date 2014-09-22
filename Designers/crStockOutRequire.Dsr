VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crStockOutRequire 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16425
   OleObjectBlob   =   "crStockOutRequire.dsx":0000
End
Attribute VB_Name = "crStockOutRequire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iPluDB As ADODB.Recordset
Dim rsReason As New ADODB.Recordset
Dim rsNhapxuat As New ADODB.Recordset
Dim rsCashier As New ADODB.Recordset
Dim rsPLU As New ADODB.Recordset
Dim rscompany As New ADODB.Recordset

Private Sub Report_Initialize()
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsReason = Open_Table(cnData, "Stock_List")
    Set rsNhapxuat = Open_Table(cnData, "InOutType")
    Set rsCashier = LoadPasswordData
    Set rsPLU = Open_Table(cnData, "Inventory")
     Set rscompany = Open_Table(cnData, "Setup")
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
    With rsPLU
        .Find "ItemNum='" & Trim(PluCode.Value) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtUnit.SetText .Fields("Unit")
        End If
    End With
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
    With rscompany
        If .RecordCount > 0 Then
            lblInfor1.SetText .Fields("Company_info_1")
            lblInfor2.SetText .Fields("Company_info_2")
            lblInfor3.SetText .Fields("Company_info_3")
            lblInfor4.SetText .Fields("Company_info_4")
            lblInfor5.SetText .Fields("Company_info_5")
        End If
    End With
    With rsReason
        .Find "ID='" & Trim(stockID.Value) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblKho.SetText .Fields("Stock_name")
        End If
    End With
    
    With rsNhapxuat
        .Find "MaNX='" & Trim(lydonhap.Value) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblLydo.SetText .Fields("Diengiai")
        End If
    End With
    With rsCashier
        .Find "ID='" & Trim(Nguoigiao.Value) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblCashier.SetText .Fields("userName")
        End If
    End With
End Sub


Private Sub Section9_Format(ByVal pFormattingInfo As Object)
Dim ngayxuat As String
ngayxuat = Trim(NgayCT.Value)
    lblDateIn.SetText "Ngµy " & Right(ngayxuat, 2) & " th¸ng " & Mid(ngayxuat, 5, 2) & " n¨m " & Left(ngayxuat, 4)
End Sub
