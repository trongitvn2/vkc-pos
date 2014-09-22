VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crStockIn 
   ClientHeight    =   9510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14970
   OleObjectBlob   =   "crStockIn.dsx":0000
End
Attribute VB_Name = "crStockIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iPluDB As ADODB.Recordset
Dim rsVendor As New ADODB.Recordset
Dim rsNhapxuat As New ADODB.Recordset
Dim rsCashier As New ADODB.Recordset
Dim rsPLU As New ADODB.Recordset

Private Sub Report_Initialize()
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsVendor = Open_Table(cnData, "Vendors")
    Set rsNhapxuat = Open_Table(cnData, "InOutType")
    Set rsCashier = LoadPasswordData
    Set rsPLU = Open_Table(cnData, "SetMPLU")
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
    With rsPLU
        .Find "PluCode='" & Trim(PluCode.Value) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtUnit.SetText .Fields("Unit")
        End If
    End With
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
    lblDateIn.SetText gfCONVERT_STRING_TO_DATE(NgayCT.Value)
    lblDate.SetText gfCONVERT_STRING_TO_DATE(NgayDH.Value)
    With rsVendor
        .Find "Vendor_Number='" & Trim(Donvixuat.Value) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            DVXuat.SetText .Fields("Vendor_Name")
            txtDiachi.SetText .Fields("Address_1") & "-" & .Fields("Address_2")
        End If
    End With
    With rsNhapxuat
        .Find "MaNX='" & Trim(lydonhap.Value) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblLydo.SetText .Fields("Diengiai")
        End If
    End With
    With rsCashier
        .Find "ID='" & Trim(Nguoinhan.Value) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblCashier.SetText .Fields("userName")
        End If
    End With
End Sub

Private Sub Section9_Format(ByVal pFormattingInfo As Object)
    txtReadNumber.SetText readnumber(SumAmt.Value)
End Sub
