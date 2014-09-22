VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} crPLUPurseError 
   ClientHeight    =   7080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9345
   OleObjectBlob   =   "crPLUPurseError.dsx":0000
End
Attribute VB_Name = "crPLUPurseError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arrCap()        As String
Dim mStrSiteID      As String

Private Sub Report_Initialize()
On Error GoTo errHdl

    arrCap = LoadLanguage(LngFile, "#03:049:")
    DDate.Suppress = True
    TTime.Suppress = True
    txtPage.Suppress = True
    
    With SumQty
        .DecimalPlaces = DecimalQtyNumber
        .DecimalSymbol = DecimalMark
        If DigitsGroup > 0 Then .ThousandsSeparators = True
        .ThousandSymbol = DigitGroupMark
        .Suppress = Not gblnShowQty
    End With
    
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description
End Sub

Private Sub Section10_Format(ByVal pFormattingInfo As Object)
On Error GoTo errHdl
    If Site.Value & "" <> "" Then
        txtSitename.SetText ""
    End If
    mStrSiteID = Site.Value
    Site.Suppress = True
    
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description
End Sub


Private Sub Section12_Format(ByVal pFormattingInfo As Object)
On Error GoTo errHdl
    If Net.Value & "" <> "" Then
        txtNetname.SetText arrCap(13)
    End If
    
    Net.Suppress = True
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description
End Sub

Private Sub Section5_Format(ByVal pFormattingInfo As Object)
On Error GoTo errHdl
    If DateOpen.Value & "" <> "" Then
        txtDateOpen.SetText Right(DateOpen.Value, 2) & "/" & _
        Mid(DateOpen.Value, 5, 2) & "/" & Left(DateOpen.Value, 4)
    End If
    
    DateOpen.Suppress = True
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description
End Sub

Private Sub Section6_Format(ByVal pFormattingInfo As Object)
On Error GoTo errHdl

    lblOrder.SetText arrCap(2)
    lblDocNo.SetText arrCap(3)
    lblDateOpen.SetText arrCap(4)
    lblPluCode.SetText arrCap(5)
    lblPluName.SetText arrCap(6)
    lblQty.SetText arrCap(7)
    lblUnit.SetText arrCap(8)

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description
End Sub

Private Sub Section7_Format(ByVal pFormattingInfo As Object)
On Error GoTo errHdl

    lblPage.SetText arrCap(10) & txtPage.Value
    lblDate.SetText arrCap(11) & DDate.Value
    lblTime.SetText arrCap(12) & " : " & TTime.Value

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description
End Sub

Private Sub Section8_Format(ByVal pFormattingInfo As Object)
On Error GoTo errHdl

    lblTitle1.SetText UCase(arrCap(1))

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description
End Sub

'Report Footer
Private Sub Section9_Format(ByVal pFormattingInfo As Object)
On Error GoTo errHdl

    lblReporter.SetText arrCap(9)
    
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description
End Sub
