Attribute VB_Name = "mdlPrinter"
Option Explicit
Public Declare Function lstrcpy Lib "kernel32" _
   Alias "lstrcpyA" _
   (ByVal lpString1 As String, _
   ByVal lpString2 As String) _
   As Long

Public Declare Function OpenPrinter Lib "winspool.drv" _
   Alias "OpenPrinterA" _
   (ByVal pPrinterName As String, _
   phPrinter As Long, _
   pDefault As PRINTER_DEFAULTS) _
   As Long

Public Declare Function GetPrinter Lib "winspool.drv" Alias "GetPrinterA" _
   (ByVal hPrinter As Long, _
   ByVal Level As Long, _
   pPrinter As Byte, _
   ByVal cbBuf As Long, _
   pcbNeeded As Long) _
   As Long

Public Declare Function ClosePrinter Lib "winspool.drv" _
   (ByVal hPrinter As Long) _
   As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
   (Destination As Any, _
   Source As Any, _
   ByVal Length As Long)

Public Declare Function EnumJobs Lib "winspool.drv" Alias "EnumJobsA" _
   (ByVal hPrinter As Long, _
   ByVal FirstJob As Long, _
   ByVal NoJobs As Long, _
   ByVal Level As Long, _
   pJob As Byte, _
   ByVal cdBuf As Long, _
   pcbNeeded As Long, _
   pcReturned As Long) _
   As Long
   
' constants for PRINTER_DEFAULTS structure
Public Const PRINTER_ACCESS_USE = &H8
Public Const PRINTER_ACCESS_ADMINISTER = &H4

' constants for DEVMODE structure
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32

Public Type PRINTER_DEFAULTS
   pDatatype As String
   pDevMode As Long
   DesiredAccess As Long
End Type

Public Type DEVMODE
   dmDeviceName As String * CCHDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer
   dmDriverExtra As Integer
   dmFields As Long
   dmOrientation As Integer
   dmPaperSize As Integer
   dmPaperLength As Integer
   dmPaperWidth As Integer
   dmScale As Integer
   dmCopies As Integer
   dmDefaultSource As Integer
   dmPrintQuality As Integer
   dmColor As Integer
   dmDuplex As Integer
   dmYResolution As Integer
   dmTTOption As Integer
   dmCollate As Integer
   dmFormName As String * CCHFORMNAME
   dmLogPixels As Integer
   dmBitsPerPel As Long
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type

Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Type JOB_INFO_2
   JobId As Long
   pPrinterName As Long
   pMachineName As Long
   pUserName As Long
   pDocument As Long
   pNotifyName As Long
   pDatatype As Long
   pPrintProcessor As Long
   pParameters As Long
   pDriverName As Long
   pDevMode As Long
   pStatus As Long
   pSecurityDescriptor As Long
   Status As Long
   Priority As Long
   Position As Long
   StartTime As Long
   UntilTime As Long
   TotalPages As Long
   Size As Long
   Submitted As SYSTEMTIME
   time As Long
   PagesPrinted As Long
End Type

Type PRINTER_INFO_2
   pServerName As Long
   pPrinterName As Long
   pShareName As Long
   pPortName As Long
   pDriverName As Long
   pComment As Long
   pLocation As Long
   pDevMode As Long
   pSepFile As Long
   pPrintProcessor As Long
   pDatatype As Long
   pParameters As Long
   pSecurityDescriptor As Long
   Attributes As Long
   Priority As Long
   DefaultPriority As Long
   StartTime As Long
   UntilTime As Long
   Status As Long
   cJobs As Long
   AveragePPM As Long
End Type

Public Const ERROR_INSUFFICIENT_BUFFER = 122
Public Const PRINTER_STATUS_BUSY = &H200
Public Const PRINTER_STATUS_DOOR_OPEN = &H400000
Public Const PRINTER_STATUS_ERROR = &H2
Public Const PRINTER_STATUS_INITIALIZING = &H8000
Public Const PRINTER_STATUS_IO_ACTIVE = &H100
Public Const PRINTER_STATUS_MANUAL_FEED = &H20
Public Const PRINTER_STATUS_NO_TONER = &H40000
Public Const PRINTER_STATUS_NOT_AVAILABLE = &H1000
Public Const PRINTER_STATUS_OFFLINE = &H80
Public Const PRINTER_STATUS_OUT_OF_MEMORY = &H200000
Public Const PRINTER_STATUS_OUTPUT_BIN_FULL = &H800
Public Const PRINTER_STATUS_PAGE_PUNT = &H80000
Public Const PRINTER_STATUS_PAPER_JAM = &H8
Public Const PRINTER_STATUS_PAPER_OUT = &H10
Public Const PRINTER_STATUS_PAPER_PROBLEM = &H40
Public Const PRINTER_STATUS_PAUSED = &H1
Public Const PRINTER_STATUS_PENDING_DELETION = &H4
Public Const PRINTER_STATUS_PRINTING = &H400
Public Const PRINTER_STATUS_PROCESSING = &H4000
Public Const PRINTER_STATUS_TONER_LOW = &H20000
Public Const PRINTER_STATUS_USER_INTERVENTION = &H100000
Public Const PRINTER_STATUS_WAITING = &H2000
Public Const PRINTER_STATUS_WARMING_UP = &H10000
Public Const JOB_STATUS_PAUSED = &H1
Public Const JOB_STATUS_ERROR = &H2
Public Const JOB_STATUS_DELETING = &H4
Public Const JOB_STATUS_SPOOLING = &H8
Public Const JOB_STATUS_PRINTING = &H10
Public Const JOB_STATUS_OFFLINE = &H20
Public Const JOB_STATUS_PAPEROUT = &H40
Public Const JOB_STATUS_PRINTED = &H80
Public Const JOB_STATUS_DELETED = &H100
Public Const JOB_STATUS_BLOCKED_DEVQ = &H200
Public Const JOB_STATUS_USER_INTERVENTION = &H400
Public Const JOB_STATUS_RESTART = &H800

Public Function GetString(ByVal PtrStr As Long) As String
   Dim StrBuff As String * 256
   
   'Check for zero address
   If PtrStr = 0 Then
      GetString = " "
      Exit Function
   End If
   
   'Copy data from PtrStr to buffer.
   CopyMemory ByVal StrBuff, ByVal PtrStr, 256
   
   'Strip any trailing nulls from string.
   GetString = StripNulls(StrBuff)
End Function

Public Function StripNulls(OriginalStr As String) As String
   'Strip any trailing nulls from input string.
   If (InStr(OriginalStr, Chr(0)) > 0) Then
      OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
   End If

   'Return modified string.
   StripNulls = OriginalStr
End Function

Public Function PtrCtoVbString(Add As Long) As String
    Dim sTemp As String * 512
    Dim x As Long

    x = lstrcpy(sTemp, Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
         PtrCtoVbString = ""
    Else
         PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function

Public Function CheckPrinterStatus(PI2Status As Long) As String
   Dim tempStr As String
   
   If PI2Status = 0 Then   ' Return "Ready"
      CheckPrinterStatus = "Printer Status = Ready" & vbCrLf
   Else
      tempStr = ""   ' Clear
      If (PI2Status And PRINTER_STATUS_BUSY) Then
         tempStr = tempStr & "Busy  "
      End If
      
      If (PI2Status And PRINTER_STATUS_DOOR_OPEN) Then
         tempStr = tempStr & "Printer Door Open  "
      End If
      
      If (PI2Status And PRINTER_STATUS_ERROR) Then
         tempStr = tempStr & "Printer Error  "
      End If
      
      If (PI2Status And PRINTER_STATUS_INITIALIZING) Then
         tempStr = tempStr & "Initializing  "
      End If
      
      If (PI2Status And PRINTER_STATUS_IO_ACTIVE) Then
         tempStr = tempStr & "I/O Active  "
      End If

      If (PI2Status And PRINTER_STATUS_MANUAL_FEED) Then
         tempStr = tempStr & "Manual Feed  "
      End If
      
      If (PI2Status And PRINTER_STATUS_NO_TONER) Then
         tempStr = tempStr & "No Toner  "
      End If
      
      If (PI2Status And PRINTER_STATUS_NOT_AVAILABLE) Then
         tempStr = tempStr & "Not Available  "
      End If
      
      If (PI2Status And PRINTER_STATUS_OFFLINE) Then
         tempStr = tempStr & "Off Line  "
      End If
      
      If (PI2Status And PRINTER_STATUS_OUT_OF_MEMORY) Then
         tempStr = tempStr & "Out of Memory  "
      End If
      
      If (PI2Status And PRINTER_STATUS_OUTPUT_BIN_FULL) Then
         tempStr = tempStr & "Output Bin Full  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PAGE_PUNT) Then
         tempStr = tempStr & "Page Punt  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PAPER_JAM) Then
         tempStr = tempStr & "Paper Jam  "
      End If

      If (PI2Status And PRINTER_STATUS_PAPER_OUT) Then
         tempStr = tempStr & "Paper Out  "
      End If
      
      If (PI2Status And PRINTER_STATUS_OUTPUT_BIN_FULL) Then
         tempStr = tempStr & "Output Bin Full  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PAPER_PROBLEM) Then
         tempStr = tempStr & "Page Problem  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PAUSED) Then
         tempStr = tempStr & "Paused  "
      End If

      If (PI2Status And PRINTER_STATUS_PENDING_DELETION) Then
         tempStr = tempStr & "Pending Deletion  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PRINTING) Then
         tempStr = tempStr & "Printing  "
      End If
      
      If (PI2Status And PRINTER_STATUS_PROCESSING) Then
         tempStr = tempStr & "Processing  "
      End If
      
      If (PI2Status And PRINTER_STATUS_TONER_LOW) Then
         tempStr = tempStr & "Toner Low  "
      End If

      If (PI2Status And PRINTER_STATUS_USER_INTERVENTION) Then
         tempStr = tempStr & "User Intervention  "
      End If
      
      If (PI2Status And PRINTER_STATUS_WAITING) Then
         tempStr = tempStr & "Waiting  "
      End If
      
      If (PI2Status And PRINTER_STATUS_WARMING_UP) Then
         tempStr = tempStr & "Warming Up  "
      End If
      
      'Did you find a known status?
      If Len(tempStr) = 0 Then
         tempStr = "Unknown Status of " & PI2Status
      End If
      
      'Return the Status
      CheckPrinterStatus = "Printer Status = " & tempStr & vbCrLf
   End If
End Function


Public Sub myPrint(ByVal Document As CRAXDDRT.Report, ByVal CurrentPage As Integer, ByVal MaxPage As Long)
On Error GoTo errHdl

    With frmPrint
        .DocumentPrint = Document
        .MaxPageNumber = MaxPage
        .CurrentPageNumber = CurrentPage
        .Show vbModal
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlPrinter- myPrint"
End Sub


Public Function CheckPrinter(PrinterStr As String, JobStr As String, ByVal PrinterName As String) As String
   Dim hPrinter As Long
   Dim ByteBuf As Long
   Dim BytesNeeded As Long
   Dim PI2 As PRINTER_INFO_2
   Dim JI2 As JOB_INFO_2
   Dim PrinterInfo() As Byte
   Dim JobInfo() As Byte
   Dim result As Long
   Dim LastError As Long
   Dim tempStr As String
   Dim NumJI2 As Long
   Dim pDefaults As PRINTER_DEFAULTS
   Dim I As Integer
   
   'Set a default return value if no errors occur.
   CheckPrinter = "Printer info retrieved"
   
   'NOTE: You can pick a printer from the Printers Collection
   'or use the EnumPrinters() API to select a printer name.
   
   'Use the default printer of Printers collection.
   'This is typically, but not always, the system default printer.
   PrinterName = Printer.DeviceName
   
   'Set desired access security setting.
   pDefaults.DesiredAccess = PRINTER_ACCESS_USE
   
   'Call API to get a handle to the printer.
   result = OpenPrinter(PrinterName, hPrinter, pDefaults)
   If result = 0 Then
      'If an error occurred, display an error and exit sub.
      CheckPrinter = "Cannot open printer " & PrinterName & _
         ", Error: " & Err.LastDllError
      Exit Function
   End If

   'Init BytesNeeded
   BytesNeeded = 0

   'Clear the error object of any errors.
   Err.Clear

   'Determine the buffer size that is needed to get printer info.
   result = GetPrinter(hPrinter, 2, 0&, 0&, BytesNeeded)
   
   'Check for error calling GetPrinter.
   If Err.LastDllError <> ERROR_INSUFFICIENT_BUFFER Then
      'Display an error message, close printer, and exit sub.
      CheckPrinter = " > GetPrinter Failed on initial call! <"
      ClosePrinter hPrinter
      Exit Function
   End If
   
   'Note that in Charles Petzold's book "Programming Windows 95," he
   'states that because of a problem with GetPrinter on Windows 95 only, you
   'must allocate a buffer as much as three times larger than the value
   'returned by the initial call to GetPrinter. This is not done here.
   ReDim PrinterInfo(1 To BytesNeeded)
   
   ByteBuf = BytesNeeded
   
   'Call GetPrinter to get the status.
   result = GetPrinter(hPrinter, 2, PrinterInfo(1), ByteBuf, _
     BytesNeeded)
   
   'Check for errors.
   If result = 0 Then
      'Determine the error that occurred.
      LastError = Err.LastDllError()
      
      'Display error message, close printer, and exit sub.
      CheckPrinter = "Couldn't get Printer Status!  Error = " _
         & LastError
      ClosePrinter hPrinter
      Exit Function
   End If

   'Copy contents of printer status byte array into a
   'PRINTER_INFO_2 structure to separate the individual elements.
   CopyMemory PI2, PrinterInfo(1), Len(PI2)
   
   'Check if printer is in ready state.
   tempStr = PI2.Status
   
   'Add printer name, driver, and port to list.
'   PrinterStr = PrinterStr & "Printer Name = " & _
'     GetString(PI2.pPrinterName) & vbCrLf
'   PrinterStr = PrinterStr & "Printer Driver Name = " & _
'     GetString(PI2.pDriverName) & vbCrLf
'   PrinterStr = PrinterStr & "Printer Port Name = " & _
'     GetString(PI2.pPortName) & vbCrLf
'
'   'Call API to get size of buffer that is needed.
'   result = EnumJobs(hPrinter, 0&, &HFFFFFFFF, 2, ByVal 0&, 0&, _
'      BytesNeeded, NumJI2)
'
'   'Check if there are no current jobs, and then display appropriate message.
'   If BytesNeeded = 0 Then
'      JobStr = "No Print Jobs!"
'   Else
'      'Redim byte array to hold info about print job.
'      ReDim JobInfo(0 To BytesNeeded)
'
'      'Call API to get print job info.
'      result = EnumJobs(hPrinter, 0&, &HFFFFFFFF, 2, JobInfo(0), _
'        BytesNeeded, ByteBuf, NumJI2)
'
'      'Check for errors.
'      If result = 0 Then
'         'Get and display error, close printer, and exit sub.
'         LastError = Err.LastDllError
'         CheckPrinter = " > EnumJobs Failed on second call! <  Error = " _
'            & LastError
'         ClosePrinter hPrinter
'         Exit Function
'      End If
'
'   End If
   CheckPrinter = tempStr
   'Close the printer handle.
   ClosePrinter hPrinter
End Function

