Attribute VB_Name = "GetSystemInfor"
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" _
(ByVal lpRootPathName As String, _
ByVal lpVolumeNameBuffer As String, _
ByVal nVolumeNameSize As Long, _
ByRef lpVolumeSerialNumber As Any, _
ByRef lpMaximumComponentLength As Any, _
ByRef lpFileSystemFlags As Any, _
ByVal lpFileSystemNameBuffer As String, _
ByVal nFileSystemNameSize As Long) As Boolean
Public CPUCode, MainSerial, HDD_Code As String

Function readseriemainboard() As String
On Error GoTo Handle
Dim serialMain As String
    Dim objs As Object
    
    Dim obj As Object
    
    Dim WMI As Object
    
    Dim sAns As String
    
    Set WMI = GetObject("WinMgmts:")
    
    Set objs = WMI.InstancesOf("Win32_BaseBoard")
    
    For Each obj In objs
    
    sAns = sAns & obj.SerialNumber
    
    If sAns < objs.count Then sAns = sAns & ","
    
    Next
    readseriemainboard = sAns
Exit Function
Handle:
MsgBox Err.Number & Err.Description & " - readseriemainboard "
End Function


Function readserienumber() As String
On Error GoTo Handle
Dim fso As Object, Drv As Object

        'Create a FileSystemObject object

          Set fso = CreateObject("Scripting.FileSystemObject")

          'Assign the current drive letter if not specified

          Set Drv = fso.GetDrive("C:\")
          With Drv

              If .IsReady Then

                  DriveSerial = Abs(.SerialNumber)

              Else    '"Drive Not Ready!"

                  DriveSerial = -1

              End If

          End With

          'Clean up

          Set Drv = Nothing

          Set fso = Nothing

          readserienumber = DriveSerial
Exit Function
Handle: MsgBox Err.Number & Err.Description & " - readserienumber"
 End Function



