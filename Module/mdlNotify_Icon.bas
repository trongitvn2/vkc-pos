Attribute VB_Name = "mdlNotify_Icon"
Option Explicit
Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias " Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Const NIM_ADD = &H0

Public Const NIM_DELETE = &H2

Public Const NIM_MODIFY = &H1

Public Const NMPWAIT_NOWAIT = &H1

Public Const NMPWAIT_USE_DEFAULT_WAIT = &H0

Public Const NIF_ICON = &H2

Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200

Public Const WM_LBUTTONDBLCLK = &H203

Public Const WM_LBUTTONDOWN = &H201

Public Const WM_LBUTTONUP = &H202

Public Const WM_MBUTTONDBLCLK = &H209

Public Const WM_MBUTTONDOWN = &H207

Public Const WM_MBUTTONUP = &H208

Public Type NOTIFYICONDATA
        cbSize As Long
        hwnd As Long
        uID As Long
        uFlags As Long
        uCallbackMessage As Long
        hIcon As Long
        szTip As String * 64
End Type

Public theData As NOTIFYICONDATA
