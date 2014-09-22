Attribute VB_Name = "mdlShutdown_Hibernate_Logoff"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''
'Use for WINNT Restart, Logoff, Shutdown
Private Const EWX_LOGOFF = 0
Private Const EWX_SHUTDOWN = 1
Private Const EWX_REBOOT = 2
Private Const EWX_POWEROFF = 8
Private Const EWX_FORCE = 4
Private Const TOKEN_ADJUST_PRIVILEGES = &H20
Private Const TOKEN_QUERY = &H8
Private Const SE_PRIVILEGE_ENABLED = &H2
Private Const ANYSIZE_ARRAY = 1
Private Const VER_PLATFORM_WIN32_NT = 2
'System Power State
Public Enum eSystemPowerState
waSUSPEND
waHIBERNATE
End Enum
Private Type OSVERSIONINFO
dwOSVersionInfoSize As Long
dwMajorVersion As Long
dwMinorVersion As Long
dwBuildNumber As Long
dwPlatformId As Long
szCSDVersion As String * 128
End Type
Private Type LUID
LowPart As Long
HighPart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
pLuid As LUID
Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type
'Suspend|Hibernate
Private Declare Function SetSuspendState Lib "Powrprof" (ByVal Hibernate As Long, ByVal ForceCritical As Long, ByVal DisableWakeEvent As Long) As Long
'Lock Computer
Private Declare Function LockWorkStation Lib "user32.dll" () As Long
'Shutdown, Restart, LogOff
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivileglevelue Lib "advapi32" Alias "LookupPrivileglevelueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As OSVERSIONINFO) As Long

Private Const Process_Force = True
'--------------------------------------------------

'Detect if the program is running under Windows NT
Private Function IsWinNT() As Boolean
Dim myOS As OSVERSIONINFO
myOS.dwOSVersionInfoSize = Len(myOS)
GetVersionEx myOS
IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function
'set the shut down privilege for the current application
Private Sub EnableShutDown()
Dim hProc As Long
Dim hToken As Long
Dim mLUID As LUID
Dim mPriv As TOKEN_PRIVILEGES
Dim mNewPriv As TOKEN_PRIVILEGES
hProc = GetCurrentProcess()
OpenProcessToken hProc, TOKEN_ADJUST_PRIVILEGES + TOKEN_QUERY, hToken
LookupPrivileglevelue "advapi32", "SeShutdownPrivilege", mLUID
mPriv.PrivilegeCount = 1
mPriv.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
mPriv.Privileges(0).pLuid = mLUID
' enable shutdown privilege for the current application
AdjustTokenPrivileges hToken, False, mPriv, 4 + (12 * mPriv.PrivilegeCount), mNewPriv, 4 + (12 * mNewPriv.PrivilegeCount)
End Sub
' Shut Down NT
'Public Sub ShutDownNT(Force As Boolean)
'Dim ret As Long
'Dim Flags As Long
'Flags = EWX_POWEROFF + EWX_SHUTDOWN
'If Force Then Flags = Flags + EWX_FORCE
'If IsWinNT Then EnableShutDown
'ExitWindowsEx Flags, 0
'End Sub
''Restart NT
Public Sub RebootNT(Force As Boolean)
Dim ret As Long
Dim Flags As Long
Flags = EWX_REBOOT
If Force Then Flags = Flags + EWX_FORCE
If IsWinNT Then EnableShutDown
ExitWindowsEx Flags, 0
End Sub
'Log off the current user
Public Sub LogOffNT(Force As Boolean)
Dim ret As Long
Dim Flags As Long
Flags = EWX_LOGOFF
If Force Then Flags = Flags + EWX_FORCE
ExitWindowsEx Flags, 0
End Sub

Public Function SetSystemPowerState(mAction As eSystemPowerState, mForceSuspension As Boolean, mDisableWakeEvents As Boolean) As Boolean
'Supports: Suspend(Stand By), Hibernate
'Platforms: Only Windows 98 or later, Windows 2000 or later
'If Hibernation is not enabled on the target system, it will Suspend instead
If SetSuspendState(mAction, mForceSuspension, mDisableWakeEvents) Then SetSystemPowerState = True
End Function

Public Function LockComputer() As Boolean
'Platforms: Only Windows 2000 or later
If LockWorkStation Then LockComputer = True
End Function

'----------------S? d?ng----------------------------------

'Clock Computer
'LockComputer

'Hibernate
'If SetSystemPowerState(waHIBERNATE, Process_Force, False) Then MsgBox "He thong khong ho tro Hibernate"
'
''StandBy
'If SetSystemPowerState(waSUSPEND, Process_Force, False) Then MsgBox "He thong khong ho tro StandBy"
'
''LogOff
'LogOffNT Process_Force
'
''Restart
'RebootNT Process_Force
''
''Shutdown
Public Sub Tat_may()
    ShutDownNT Process_Force
End Sub



