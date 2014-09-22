VERSION 5.00
Object = "{CC34CEB4-5C10-11D1-A40F-00A024229C83}#1.0#0"; "CapturePRO3.dll"
Begin VB.Form PreviewForm 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Khung h×nh"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   ControlBox      =   0   'False
   FillColor       =   &H00FF0000&
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5610
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin CAPTUREPRO3LibCtl.CapturePRO CapturePRO1 
      Height          =   5595
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   9869
      ErrStr          =   "B1279074AA2B3CC8737E7BE41C1ABAB8"
      ErrCode         =   824530348
      ErrInfo         =   -194444910
      _cx             =   10821
      _cy             =   9869
      BorderVisible   =   -1  'True
      BorderWidth     =   1
      Caption         =   ""
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   65535
      AudioBits       =   8
      AudioChannels   =   1
      AudioSampleRate =   11025
      AutoIncrement   =   0   'False
      AutoSave        =   0   'False
      AutoStretch     =   0
      CaptionHeight   =   0
      CaptionLeft     =   0
      CaptionTop      =   0
      CaptionWidth    =   0
      CaptureAudio    =   -1  'True
      ClipCaption     =   0   'False
      FTPPassword     =   ""
      FTPRename       =   0   'False
      FTPUserName     =   ""
      FrameFile       =   ""
      FrameRate       =   30
      HAlign          =   0
      Interval        =   60000
      PICPassword     =   ""
      Preview         =   0   'False
      PreviewRate     =   15
      ProxyServer     =   ""
      ResX            =   0
      ResY            =   0
      SaveJPGChromFactor=   36
      SaveJPGLumFactor=   32
      SaveJPGProgressive=   0   'False
      SaveJPGSubSampling=   2
      ScaleHeight     =   0
      ScalePercent    =   100
      ScaleWidth      =   0
      ShadowText      =   -1  'True
      StreamFile      =   ""
      TimeLimit       =   0
      VAlign          =   0
   End
End
Attribute VB_Name = "PreviewForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Dim cxDiff As Long, cyDiff As Long

Private Sub CapturePRO1_DeviceError(ByVal ErrorCode As Long, ByVal ErrorText As String)
MainForm.LblStatus.Caption = ErrorText
End Sub

Private Sub CapturePRO1_DeviceStatus(ByVal StatusCode As Long, ByVal StatusText As String)
If (StatusCode = IDS_CAP_END) Then
    MainForm.CmdCap.Caption = "Capture Stream"
   Exit Sub
End If

MainForm.LblStatus.Caption = StatusText
DoEvents
End Sub

Private Sub CapturePRO1_DeviceWarning(ByVal WarningCode As Long, ByVal WarningText As String)
  MsgBox "Warning: " + WarningText
End Sub

' FilterListChanged event handling procedure
Private Sub CapturePRO1_FilterListChanged(ByVal FilterCategory As String, ByVal ChangeType As CAPTUREPRO3LibCtl.FILTERLISTCHANGES, ByVal lFilterIndex As Long, ByVal FilterName As String)
   
    If (FilterCategory = "Video Capture Sources") Then
      If (ChangeType = FILTER_ADDED) Then
          If MainForm.Combo1.List(0) = "No Capture Devices" And lFilterIndex = 0 Then
          ' single capture device is plugged into computer
            MainForm.Combo1.RemoveItem (0) ' remove "No Capture Devices" item
            MainForm.Combo1.AddItem (FilterName)
            MainForm.Combo1.ListIndex = 0
            CapturePRO1.VideoDeviceIndex = lFilterIndex
            Exit Sub
          End If
          MainForm.Combo1.AddItem (FilterName)
          If Not CapturePRO1.IsConnected And lFilterIndex <> CapturePRO1.VideoDeviceIndex Then
            If MsgBox(FilterName + " is plugged into your computer. Do you want to use it?", vbYesNo) = vbYes Then
                CapturePRO1.VideoDeviceIndex = lFilterIndex
            End If
          End If
      ElseIf (ChangeType = FILTER_REMOVED) Then
          MainForm.Combo1.RemoveItem (lFilterIndex)
      End If
      If (CapturePRO1.NumDevices > 0) Then
         MainForm.Combo1.ListIndex = CapturePRO1.VideoDeviceIndex ' + 1
      Else
      ' no capture devives in the computer
          MainForm.Combo1.AddItem ("Kh«ng t×m thÊt thiÕt bÞ")
          MainForm.Combo1.ListIndex = 0
      End If
   
      
    End If
    If ChangeType = FILTER_REMOVED And MainForm.CmdConnect.Caption = "Disconnect" And Not CapturePRO1.IsConnected Then
        ' disable controls
        MainForm.OnDisconnect
    End If
  
End Sub

Private Sub CapturePRO1_Resize()

PreviewForm.Width = CapturePRO1.Width + cxDiff
PreviewForm.Height = CapturePRO1.Height + cyDiff

End Sub
Private Sub Form_Load()
    
cxDiff = PreviewForm.Width - CapturePRO1.Width
cyDiff = PreviewForm.Height - CapturePRO1.Height

End Sub
