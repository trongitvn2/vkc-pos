Attribute VB_Name = "mdlShowHelp"
Option Explicit


Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" _
    (ByVal hWndCaller As Long, ByVal pszFile As String, _
    ByVal uCommand As Long, ByVal dwData As Long) As Long

Private Declare Function HTMLHelp2 Lib "hhctrl.ocx" Alias "HtmlHelpA" _
    (ByVal hWndCaller As Long, ByVal pszFile As String, _
    ByVal uCommand As Long, ByVal dwData As String) As Long

Private Const HH_DISPLAY_TOPIC = &H0
Private Const HH_SET_WIN_TYPE = &H4
Private Const HH_GET_WIN_TYPE = &H5
Private Const HH_GET_WIN_HANDLE = &H6
Private Const HH_DISPLAY_TEXT_POPUP = &HE   ' Display string resource ID or text in a pop-up window.
Private Const HH_HELP_CONTEXT = &HF         ' Display mapped numeric value in  dwData.
Private Const HH_TP_HELP_CONTEXTMENU = &H10 ' Text pop-up help, similar to WinHelp's HELP_CONTEXTMENU.
Private Const HH_TP_HELP_WM_HELP = &H11     ' text pop-up help, similar to WinHelp's HELP_WM_HELP.
Private Const HH_CLOSE_ALL = &H12

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As _
    String, ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub Showhelp(ByVal HelpTopic As String)
    On Error Resume Next
    Dim HelpFile As String
    Dim FormHandle As Long
    Dim HhWnd As Long
    
    FormHandle = frmLogin.hWnd
    
    HelpFile = App.Path & "\Help\VKC-TOUCH_USER'S_GUIDE.CHM"
    If Dir(HelpFile) = "" Then Exit Sub
'    HhWnd = HTMLHelp2(FormHandle, HelpFile, HH_DISPLAY_TOPIC, HelpTopic & ".htm")
    HhWnd = HTMLHelp2(0, HelpFile, HH_DISPLAY_TOPIC, HelpTopic & ".htm")
    If HhWnd = 0 Then GoTo ErrHandle
    Exit Sub
ErrHandle:
    ShellExecute 0, "", HelpFile, 0, 0, 1
End Sub

Public Sub CloseHelp()
'    If Dir(HelpFile) = "" Then Exit Sub
'    HtmlHelp frmMain.hWnd, "", HH_CLOSE_ALL, 0
    HtmlHelp 0, "", HH_CLOSE_ALL, 0
End Sub
