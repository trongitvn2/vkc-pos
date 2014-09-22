Attribute VB_Name = "mdbOpen_Cash"
Private Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Private Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Private Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Private Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
Const ModuleName As String = "ModuleCashDraw"
Private Type DOCINFO
pDocName As String
pOutputFile As String
pDatatype As String
End Type

Public Property Get PrinterPort(PrinterName As String) As String
On Error Resume Next
Dim X As Printer
For Each X In Printers
If UCase(X.DeviceName) = UCase(PrinterName) Then
PrinterPort = X.Port
Exit For
End If
Next
On Error GoTo 0
End Property

Public Sub OpenPrinterCashDraw(SelectedPrinter As String)
On Error Resume Next
Dim myStr As String
myStr = Chr$(27) & "p" & Chr$(0) & Chr$(97) & Chr$(98)
Call SendToPrn(myStr, SelectedPrinter)
On Error GoTo 0
End Sub

Public Sub SendToPrn(outString As String, PrinterName)
On Error Resume Next
Dim lhPrinter As Long
Dim lReturn As Long
Dim lpcWritten As Long
Dim lDoc As Long
Dim sWrittenData As String
Dim printdata As String
Dim MyDocInfo As DOCINFO
Dim Printerok As Boolean
Dim X As Printer
'kiÓm tra sù tån t¹i cña m¸y in ®­îc g¸n víi m¸y in receipt trong hÖ thèng
For Each X In Printers
If Left(UCase(X.DeviceName), 8) = Left(UCase(PrinterName), 8) Then
Set Printer = X
Printerok = True
Exit For
End If
Next
If Not Printerok Then
    MsgBox "M¸y in kÕt nèi víi Cash Drawer kh«ng tån t¹i!", vbInformation
    Exit Sub
End If
printdata = outString
lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
If lReturn = 0 Then
MsgBox "The Printer Name not recognized."
Exit Sub
End If
MyDocInfo.pDocName = "Open Cash Drawer"
MyDocInfo.pOutputFile = vbNullString
MyDocInfo.pDatatype = vbNullString
lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
Call StartPagePrinter(lhPrinter)
lReturn = WritePrinter(lhPrinter, ByVal printdata, Len(printdata), lpcWritten)
lReturn = EndPagePrinter(lhPrinter)
lReturn = EndDocPrinter(lhPrinter)
lReturn = ClosePrinter(lhPrinter)

On Error GoTo 0

End Sub
