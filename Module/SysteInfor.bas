Attribute VB_Name = "SystemInfor"
Public Type SYSTEM_INFO
dwOemID As Long
dwPageSize As Long
lpMinimumApplicationAddress As Long
lpMaximumApplicationAddress As Long
dwActiveProcessorMask As Long
dwNumberOrfProcessors As Long
dwProcessorType As Long
dwAllocationGranularity As Long
dwReserved As Long
End Type
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Public Declare Function GetNetworkParams Lib "IPHlpApi" (FixedInfo As Any, pOutBufLen As Long) As Long
Public Declare Function GetAdaptersInfo Lib "IPHlpApi" (IpAdapterInfo As Any, pOutBufLen As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function EnumPorts Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal Level As Long, ByVal lpbPorts As Long, ByVal cbBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pTo As Any, uFrom As Any, ByVal lSize As Long)
Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GetProcessHeap Lib "kernel32" () As Long
Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Declare Function EnumPrinters Lib "winspool.drv" Alias "EnumPrintersA" (ByVal Flags As Long, ByVal Name As String, ByVal Level As Long, pPrinterEnum As Long, ByVal cdBuf As Long, pcbNeeded As Long, pcReturned As Long) As Long
Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Public Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Long) As Long
Public Declare Function GetKeyboardType Lib "user32" (ByVal nTypeFlag As Long) As Long
Dim CPUCode As String

Public Function TrimSpace(strName As String) As String
    On Error GoTo errHdl

    Dim sResult As String
    Dim k As Integer
    
    For k = 1 To Len(strName)
    DoEvents
        If Mid(strName, k, 1) <> Space(1) Then
            sResult = sResult & Mid(strName, k, 1)
        End If
    Next k
    TrimSpace = sResult
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal - TrimSpace"
End Function


Public Function Lay_CPU()
Dim cpu As String
    Dim SInfo As SYSTEM_INFO
    GetSystemInfo SInfo
    cpu = TrimSpace(str$(SInfo.dwNumberOrfProcessors) & str$(SInfo.dwProcessorType) & str$(SInfo.dwReserved))
    Lay_CPU = cpu
End Function


