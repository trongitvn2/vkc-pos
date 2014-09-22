VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MainForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chôp h×nh"
   ClientHeight    =   4185
   ClientLeft      =   210
   ClientTop       =   825
   ClientWidth     =   5775
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleMode       =   0  'User
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   Begin prjTouchScreen.MyButton cmdThoat 
      Height          =   615
      Left            =   4530
      TabIndex        =   26
      Top             =   3540
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "&Tho¸t"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "MainFrm.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton CmdCap 
      Height          =   615
      Left            =   3420
      TabIndex        =   25
      Top             =   3540
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "Chôp h×nh"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "MainFrm.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton CmdPreview 
      Height          =   615
      Left            =   2310
      TabIndex        =   24
      Top             =   3540
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "Xem h×nh ¶nh"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "MainFrm.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton CmdConnect 
      Height          =   615
      Left            =   1200
      TabIndex        =   23
      Top             =   3540
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "KÕt nèi WC"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "MainFrm.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton VideoConfig 
      Height          =   615
      Left            =   90
      TabIndex        =   22
      Top             =   3540
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "CÊu h×nh"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "MainFrm.frx":007C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   150
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame fraDevice 
      Caption         =   "Tïy chän th«ng sè"
      ForeColor       =   &H00FF0000&
      Height          =   3015
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin prjTouchScreen.MyButton CmdCapFN 
         Height          =   585
         Left            =   4590
         TabIndex        =   21
         Top             =   2400
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1032
         BTYPE           =   5
         TX              =   "Browse"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16578804
         BCOLO           =   16777152
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "MainFrm.frx":0098
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton CommandFont 
         Height          =   585
         Left            =   4590
         TabIndex        =   20
         Top             =   1740
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   1032
         BTYPE           =   5
         TX              =   "Font..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16578804
         BCOLO           =   16777152
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "MainFrm.frx":00B4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.TextBox EdtInterval 
         Height          =   375
         Left            =   3930
         TabIndex        =   15
         Top             =   1200
         Width           =   1515
      End
      Begin VB.TextBox EdtCaption 
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   1320
         TabIndex        =   14
         Top             =   2100
         Width           =   3225
      End
      Begin VB.TextBox EdtCapLeft 
         Height          =   375
         Left            =   3870
         TabIndex        =   13
         Top             =   1680
         Width           =   675
      End
      Begin VB.TextBox EdtCapTop 
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   1680
         Width           =   675
      End
      Begin VB.TextBox EdtFile 
         Height          =   390
         Left            =   1320
         TabIndex        =   10
         Top             =   2520
         Width           =   3225
      End
      Begin VB.ComboBox Effect 
         Height          =   345
         ItemData        =   "MainFrm.frx":00D0
         Left            =   1320
         List            =   "MainFrm.frx":00D2
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1230
         Width           =   1575
      End
      Begin VB.ComboBox ResolutionCombo 
         Enabled         =   0   'False
         Height          =   345
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1185
      End
      Begin VB.ComboBox ColorFormatCombo 
         Enabled         =   0   'False
         Height          =   345
         Left            =   3930
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   690
         Width           =   1545
      End
      Begin VB.ComboBox Combo1 
         Height          =   345
         ItemData        =   "MainFrm.frx":00D4
         Left            =   1320
         List            =   "MainFrm.frx":00D6
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   4155
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Kho¶ngTG:"
         Height          =   345
         Left            =   2820
         TabIndex        =   19
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "M· sè h×nh:"
         Height          =   285
         Left            =   90
         TabIndex        =   18
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "C¸ch lÒ tr¸i:"
         Height          =   345
         Left            =   2610
         TabIndex        =   17
         Top             =   1710
         Width           =   1215
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "C¸ch lÒ trªn:"
         Height          =   345
         Left            =   30
         TabIndex        =   16
         Top             =   1680
         Width           =   1245
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "Tªn file:"
         Height          =   255
         Left            =   540
         TabIndex        =   11
         Top             =   2550
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "HiÖu øng:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1260
         Width           =   1035
      End
      Begin VB.Label ResolLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "§é ph©n gi¶i:"
         Enabled         =   0   'False
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   765
         Width           =   1155
      End
      Begin VB.Label ColorLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "Tïy chän mµu s¾c:"
         Enabled         =   0   'False
         Height          =   405
         Left            =   2520
         TabIndex        =   6
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "ThiÕt bÞ:"
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Label LblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   0
      TabIndex        =   3
      Top             =   3030
      Width           =   5565
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit

Const AUDIO_PROP_PAGES = PROP_AUDIO_DEVICE + PROP_AUDIO_RENDERER
Const VIDEO_CONFIG_PROP_PAGES = PROP_VIDEO_DEVICE + PROP_VIDEO_CAPTURE_STREAM + PROP_TV_TUNER

Sub SetCaptureValues()

Dim Index As Integer

    PreviewForm.CapturePRO1.Interval = CLng(EdtInterval.Text)
    PreviewForm.CapturePRO1.CaptionLeft = EdtCapLeft.Text
    PreviewForm.CapturePRO1.CaptionTop = EdtCapTop.Text
    PreviewForm.CapturePRO1.Caption = EdtCaption.Text
    PreviewForm.CapturePRO1.FrameFile = EdtFile.Text

End Sub

Function ItoB(Value As Integer) As Boolean
    If (Value <> 0) Then
       ItoB = True
    Else
       ItoB = False
    End If
End Function

Function BtoI(Value As Boolean) As Integer
    If (Value = True) Then
       BtoI = 1
    Else
       BtoI = 0
    End If
End Function

Private Sub AudioConfig_Click()
    Call PreviewForm.CapturePRO1.ShowFilterPropertyPage(AUDIO_PROP_PAGES, "")
End Sub

Private Sub Capture_Frames_Click()
    CmdCap.Caption = "Chôp h×nh"
    EdtFile.Text = PreviewForm.CapturePRO1.FrameFile
    'frmOptions.Visible = True
End Sub

Private Sub Capture_Streams_Click()
    CmdCap.Caption = "Capture Stream"
    EdtFile.Text = PreviewForm.CapturePRO1.StreamFile
    'frmOptions.Visible = False
    
End Sub

Private Sub CaptureType_Click()
    
        Call Capture_Frames_Click
    
End Sub

Private Sub CmdCap_Click()
    On Error GoTo ErrorHandler
    
    Screen.MousePointer = 11
    SetCaptureValues
    PreviewForm.CapturePRO1.CaptureFrame
    Screen.MousePointer = 0
    PImage = EdtFile.Text
    Me.Hide
    PreviewForm.Hide
    Exit Sub
ErrorHandler:
    MsgBox "Eror code: " + Hex(Err.Number) + " " + Err.Description, , "Error"
    Err.Clear
    Screen.MousePointer = 0
End Sub

Private Sub CmdCapFN_Click()
    EdtFile.Text = GetSaveName(EdtFile.Text, 0)
End Sub

Private Sub CmdConnect_Click()
On Error GoTo ErrorHandler

Dim Index As Integer

If (PreviewForm.CapturePRO1.NumDevices < 1) Then
   MsgBox "Kh«ng t×m thÊy thiÕt bÞ camera n¸o ®­îc l¾p ®Æt trong hÖ thèng", vbInformation
   Exit Sub
End If

If Not PreviewForm.CapturePRO1.IsConnected Then
   Me.MousePointer = vbHourglass
   PreviewForm.CapturePRO1.Connect Combo1.ListIndex
   Me.MousePointer = vbDefault

   CmdPreview.Enabled = True
   'Display Device information
   ResolutionCombo.Enabled = True
   ResolLabel.Enabled = True
    If (PreviewForm.CapturePRO1.NumVideoResolutions > 0) Then
        For Index = 0 To PreviewForm.CapturePRO1.NumVideoResolutions - 1
            ResolutionCombo.List(Index) = PreviewForm.CapturePRO1.ObtainVideoResolutionName(Index)
        Next
    End If
    ResolutionCombo.ListIndex = PreviewForm.CapturePRO1.VideoResolutionIndex

    ColorLabel.Enabled = True
    ColorFormatCombo.Enabled = True
    If (PreviewForm.CapturePRO1.NumVideoColorFormats > 0) Then
        For Index = 0 To PreviewForm.CapturePRO1.NumVideoColorFormats - 1
            ColorFormatCombo.List(Index) = PreviewForm.CapturePRO1.ObtainVideoColorFormatName(Index)
        Next
    End If
    ColorFormatCombo.ListIndex = PreviewForm.CapturePRO1.VideoColorFormatIndex
    
    CmdCap.Enabled = True
    VideoConfig.Enabled = PreviewForm.CapturePRO1.HasFilterPropertyPage(VIDEO_CONFIG_PROP_PAGES)
    CmdConnect.Caption = "Ng¾t kÕt nèi"
    If CmdPreview.Caption = "Èn hiÓn thÞ" Then
      PreviewForm.CapturePRO1.Preview = True
      PreviewForm.Show
    End If
Else
   PreviewForm.CapturePRO1.Disconnect
   OnDisconnect
End If

'Kich hoat CapturePro
Call CmdPreview_Click
''''''''


Exit Sub      ' Exit to avoid error handler.
ErrorHandler:   ' Error-handling routine.
    MsgBox "Eror code: " + Hex(Err.Number) + " " + Err.Description, , "Error"
    Err.Clear   ' Clear Err object fields
    Me.MousePointer = vbDefault
End Sub

' disables connection time controls and hides PreviewForm
Sub OnDisconnect()
   CmdConnect.Caption = "Connect"
   CmdCap.Enabled = False
   CmdPreview.Enabled = False
   ResolLabel.Enabled = False
   ResolutionCombo.Clear
   ResolutionCombo.Enabled = False
   ColorFormatCombo.Clear
   ColorFormatCombo.Enabled = False
   ColorLabel.Enabled = False
   VideoConfig.Enabled = False
   PreviewForm.Hide

End Sub

Private Sub CmdPreview_Click()
On Error GoTo ErrorHandler

If PreviewForm.CapturePRO1.Preview Then
   PreviewForm.CapturePRO1.Preview = False
   PreviewForm.Hide
   CmdPreview.Caption = "Xem h×nh"
Else
    PreviewForm.Show 0, frmItems
    PreviewForm.Left = MainForm.Left + MainForm.Width - 220
    PreviewForm.top = 0
    PreviewForm.CapturePRO1.Preview = True
    CmdPreview.Caption = "Èn hiÓn thÞ"
End If

Exit Sub      ' Exit to avoid error handler.
ErrorHandler:   ' Error-handling routine.
    MsgBox "Eror code: " + CStr(Err.Number) + " " + Err.Description, , "Error"
    Err.Clear   ' Clear Err object fields
End Sub

Private Sub cmdThoat_Click()
    Me.Hide
    PreviewForm.Hide
End Sub

Private Sub CommandFont_Click()
    On Error GoTo ErrorHandler
    With CommonDialog
        .Color = PreviewForm.CapturePRO1.ForeColor
        .FontBold = PreviewForm.CapturePRO1.Font.Bold
        .FontItalic = PreviewForm.CapturePRO1.Font.Italic
        .FontName = PreviewForm.CapturePRO1.Font.Name
        .FontSize = PreviewForm.CapturePRO1.Font.Size
        .FontStrikethru = PreviewForm.CapturePRO1.Font.Strikethrough
        .FontUnderline = PreviewForm.CapturePRO1.Font.Underline
        .Flags = cdlCFApply + cdlCFBoth + cdlCFEffects
        .ShowFont
        PreviewForm.CapturePRO1.ForeColor = .Color
        PreviewForm.CapturePRO1.Font.Bold = .FontBold
        PreviewForm.CapturePRO1.Font.Italic = .FontItalic
        PreviewForm.CapturePRO1.Font.Name = .FontName
        PreviewForm.CapturePRO1.Font.Size = .FontSize
        PreviewForm.CapturePRO1.Font.Strikethrough = .FontStrikethru
        PreviewForm.CapturePRO1.Font.Underline = .FontUnderline
        EdtCaption.ForeColor = CommonDialog.Color
        EdtCaption.Font.Bold = .FontBold
        EdtCaption.Font.Italic = .FontItalic
        EdtCaption.Font.Name = .FontName
        EdtCaption.Font.Size = .FontSize
        EdtCaption.Font.Strikethrough = .FontStrikethru
        EdtCaption.Font.Underline = .FontUnderline
    End With
ErrorHandler:
End Sub

Private Sub ComprConfig_Click()
    Call PreviewForm.CapturePRO1.ShowFilterPropertyPage(PROP_VIDEO_COMPRESSOR, "")
End Sub



Private Sub EdtCaption_Change()
    PreviewForm.CapturePRO1.Caption = EdtCaption.Text
End Sub

Private Sub Effect_Click()
Dim filterIndex As Long
On Error GoTo ErrorHandler
  filterIndex = -1
  If Effect.ListIndex > 0 Then
    filterIndex = PreviewForm.CapturePRO1.FindFilterIndex("Video Effect 1", Effect.List(Effect.ListIndex))
    If filterIndex < 0 Then
        MsgBox "Cannot find " + Effect.List(Effect.ListIndex) + " effect"
        Effect.ListIndex = 0
        Exit Sub
    End If
  End If
  Call PreviewForm.CapturePRO1.SetVideoProcFilter(0, "Video Effect 1", filterIndex)
Exit Sub
ErrorHandler:   ' Error-handling routine.
    MsgBox "Eror code: " + CStr(Err.Number) + " " + Err.Description, , "Error"
    Err.Clear   ' Clear Err object fieldsEnd Sub

End Sub

Private Sub File_Exit_Click()
    Unload Me
End Sub

Private Sub File_Load_Click()
Dim Index As Integer
Dim FileName As String
Dim tmpStr, tmpStr2, tmpStr3 As String
FileName = GetLoadName("", 2)
If Len(FileName) > 0 Then
    Call PreviewForm.CapturePRO1.LoadProfile(FileName)

    If PreviewForm.CapturePRO1.Preview Then
        CmdPreview.Caption = "Èn h×nh"
    Else
        CmdPreview.Caption = "Xem h×nh"
    End If

    If (PreviewForm.CapturePRO1.NumDevices > 0) Then
       Combo1.ListIndex = PreviewForm.CapturePRO1.VideoDeviceIndex ' + 1
    Else
        Combo1.ListIndex = 0
    End If

    EdtFile.Text = PreviewForm.CapturePRO1.FrameFile
   

    Effect.ListIndex = PreviewForm.CapturePRO1.ObtainVideoProcFilterIndex(0) + 1

    ' Frame options
    EdtInterval.Text = PreviewForm.CapturePRO1.Interval
    EdtCapLeft.Text = PreviewForm.CapturePRO1.CaptionLeft
    EdtCapTop.Text = PreviewForm.CapturePRO1.CaptionTop
    EdtCaption.Text = PreviewForm.CapturePRO1.Caption

    If PreviewForm.CapturePRO1.IsConnected Then

        CmdPreview.Enabled = True
       'Display Device information
        ResolutionCombo.Enabled = True
        ResolLabel.Enabled = True
        If (PreviewForm.CapturePRO1.NumVideoResolutions > 0) Then
            For Index = 0 To PreviewForm.CapturePRO1.NumVideoResolutions - 1
                ResolutionCombo.List(Index) = PreviewForm.CapturePRO1.ObtainVideoResolutionName(Index)
            Next
        End If
        ResolutionCombo.ListIndex = PreviewForm.CapturePRO1.VideoResolutionIndex

        ColorLabel.Enabled = True
        ColorFormatCombo.Enabled = True
        If (PreviewForm.CapturePRO1.NumVideoColorFormats > 0) Then
            For Index = 0 To PreviewForm.CapturePRO1.NumVideoColorFormats - 1
                ColorFormatCombo.List(Index) = PreviewForm.CapturePRO1.ObtainVideoColorFormatName(Index)
            Next
        End If
        ColorFormatCombo.ListIndex = PreviewForm.CapturePRO1.VideoColorFormatIndex
        CmdCap.Enabled = True
        VideoConfig.Enabled = PreviewForm.CapturePRO1.HasFilterPropertyPage(VIDEO_CONFIG_PROP_PAGES)
        CmdConnect.Caption = "Ng¾t kÕt nèi"
        If CmdPreview.Caption = "Èn mµn h×nh" Then
            PreviewForm.CapturePRO1.Preview = True
            PreviewForm.Show
        End If
    End If

End If
End Sub

Private Sub File_Save_Click()
Dim FileName As String
FileName = GetSaveName("", 2)
If Len(FileName) > 0 Then
    Call PreviewForm.CapturePRO1.SaveProfile(FileName)
End If
End Sub

Private Sub Form_Load()

Dim Index As Integer

If PreviewForm.CapturePRO1.Preview Then
    CmdPreview.Caption = "Xem h×nh"
Else
    CmdPreview.Caption = "Èn h×nh"
End If

If (PreviewForm.CapturePRO1.NumDevices > 0) Then
   For Index = 0 To PreviewForm.CapturePRO1.NumDevices - 1
      Combo1.List(Index) = PreviewForm.CapturePRO1.ObtainDeviceName(Index)
   Next
Else
   Combo1.List(0) = "Kh«ng cã thiÕt bÞ"
End If
Combo1.ListIndex = 0

Dim i, k, categ As Integer
Dim g As String

Effect.List(0) = "None"
i = PreviewForm.CapturePRO1.ObtainNumFilters("Video Effect 1")
If i > 0 Then
  Effect.Enabled = True
  For k = 1 To i
    Effect.List(k) = PreviewForm.CapturePRO1.ObtainFilterName("Video Effect 1", k - 1)
  Next
  Effect.ListIndex = PreviewForm.CapturePRO1.ObtainVideoProcFilterIndex(0) + 1
Else
  Effect.ListIndex = 0
  Effect.Enabled = False
End If


' Frame options
EdtFile.Text = PreviewForm.CapturePRO1.FrameFile
EdtInterval.Text = PreviewForm.CapturePRO1.Interval
EdtCapLeft.Text = PreviewForm.CapturePRO1.CaptionLeft
EdtCapTop.Text = PreviewForm.CapturePRO1.CaptionTop
EdtCaption.Text = PreviewForm.CapturePRO1.Caption


' Position PreviewForm to the right of MainForm and make sure it's
' within the screen boundaries
If (MainForm.Left + MainForm.Width + PreviewForm.Width > Screen.Width) Then
   MainForm.Left = 0
End If
PreviewForm.Move MainForm.Left + MainForm.Width, MainForm.top
If (PreviewForm.Left > Screen.Width) Then
   PreviewForm.Left = 0
End If
Call CmdConnect_Click

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub

Private Sub ResolutionCombo_Click()
On Error GoTo ErrorHandler
    If Not ResolutionCombo.ListIndex = PreviewForm.CapturePRO1.VideoResolutionIndex Then
        PreviewForm.CapturePRO1.VideoResolutionIndex = ResolutionCombo.ListIndex
    End If
Exit Sub      ' Exit to avoid error handler.
ErrorHandler:   ' Error-handling routine.
    MsgBox "Eror code: " + CStr(Err.Number) + " " + Err.Description, , "Error"
    Err.Clear   ' Clear Err object fields
End Sub

Private Sub ColorFormatCombo_Click()
On Error GoTo ErrorHandler
    If Not ColorFormatCombo.ListIndex = PreviewForm.CapturePRO1.VideoColorFormatIndex Then
        PreviewForm.CapturePRO1.VideoColorFormatIndex = ColorFormatCombo.ListIndex
    End If
Exit Sub      ' Exit to avoid error handler.
ErrorHandler:   ' Error-handling routine.
    MsgBox "Eror code: " + CStr(Err.Number) + " " + Err.Description, , "Error"
    Err.Clear   ' Clear Err object fields
End Sub


Private Sub VideoConfig_Click()
If PreviewForm.CapturePRO1.HasFilterPropertyPage(VIDEO_CONFIG_PROP_PAGES) Then
        Dim props As Long
        props = VIDEO_CONFIG_PROP_PAGES
        If Effect.ListIndex > 0 Then props = props + PROP_VIDEO_PROC_1
        Call PreviewForm.CapturePRO1.ShowFilterPropertyPage(props, "")
        ColorFormatCombo.ListIndex = PreviewForm.CapturePRO1.VideoColorFormatIndex
        ResolutionCombo.ListIndex = PreviewForm.CapturePRO1.VideoResolutionIndex
        
    End If
End Sub
