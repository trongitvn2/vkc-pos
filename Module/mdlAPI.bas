Attribute VB_Name = "mdlAPI"
Option Explicit

Public keys(0 To 255) As Byte
Public Const VK_CAPITAL = &H14
Public CapsLockState As Boolean
Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
'Dongle
Declare Function KFUNC Lib "KL2DLL32.DLL" Alias "_KFUNC@16" (ByVal Arg1 As Long, ByVal Arg2 As Long, ByVal Arg3 As Long, ByVal Arg4 As Long) As Long
Declare Function KEYBD Lib "KL2DLL32.DLL" Alias "_KEYBD@4" (ByVal Arg1 As Long) As Integer
Declare Function GETLASTKEYERROR Lib "KL2DLL32.DLL" Alias "_GETLASTKEYERROR@0" () As Long

