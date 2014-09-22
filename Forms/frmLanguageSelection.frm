VERSION 5.00
Begin VB.Form frmLanguageSelection 
   Caption         =   "Language Selection"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4605
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   2565
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdExit 
      Height          =   675
      Left            =   3180
      TabIndex        =   9
      Top             =   1830
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1191
      BTYPE           =   3
      TX              =   "&Hñy bá"
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
      BCOL            =   16578804
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLanguageSelection.frx":0000
      PICN            =   "frmLanguageSelection.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdOK 
      Height          =   675
      Left            =   540
      TabIndex        =   8
      Top             =   1830
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   1191
      BTYPE           =   3
      TX              =   "&§ång ý"
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
      BCOL            =   16578804
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLanguageSelection.frx":62B6
      PICN            =   "frmLanguageSelection.frx":62D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdHelp 
      Height          =   675
      Left            =   1890
      TabIndex        =   7
      Top             =   1830
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1191
      BTYPE           =   3
      TX              =   "&Gióp ®ì"
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
      BCOL            =   16578804
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLanguageSelection.frx":690C
      PICN            =   "frmLanguageSelection.frx":6928
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.ComboBox cboFont 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1740
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1140
      Width           =   2535
   End
   Begin VB.TextBox txtLanguage 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1740
      TabIndex        =   4
      Top             =   630
      Width           =   2535
   End
   Begin VB.FileListBox fLanguage 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      TabIndex        =   2
      Top             =   1140
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox cboLanguage 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   1740
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   2535
   End
   Begin VB.Label lblFont 
      Alignment       =   1  'Right Justify
      Caption         =   "Font:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1230
      Width           =   1275
   End
   Begin VB.Label lblLanguageName 
      Alignment       =   1  'Right Justify
      Caption         =   "Language Name:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label lblLanguage 
      Alignment       =   1  'Right Justify
      Caption         =   "Language:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   210
      Width           =   1035
   End
End
Attribute VB_Name = "frmLanguageSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type lstLanguage
    Language As String
    FileName As String
    FontName As String
End Type

Dim AlstLang() As lstLanguage
Dim LastLang As String

Private Sub cboLanguage_Change()
On Error GoTo errHdl

    txtLanguage.Font.name = AlstLang(cboLanguage.ListIndex).FontName
    txtLanguage.Text = AlstLang(cboLanguage.ListIndex).Language
    cbofont.Text = AlstLang(cboLanguage.ListIndex).FontName
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - cboLanguage_Change"
End Sub

Private Sub cboLanguage_Click()
On Error GoTo errHdl

    Call cboLanguage_Change
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - cboLanguage_Click"
End Sub

Private Sub cmdExit_Click()
On Error GoTo errHdl
    Unload Me
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - cmdExit_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHdl

    LngFile = App.Path & "\Language\" & cboLanguage.Text & ".nls"
    LngFolder = App.Path & "\Language\" & cboLanguage.Text
    SaveSettingStr "SYSTEM", "Language", Replace(LngFile, App.Path, ""), myIniFile
    LoadFont CurFont, LngFile
    'arrMessage = LoadLanguage(LngFile, "#06:001:") 'Kun: 09-03-2006
    Unload Me
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - cmdOK_Click"
End Sub

Private Sub Form_Activate()
    Dim DescArr() As String
    DescArr = LoadLanguage(LngFile, "#03:018:")
    If cmdOK.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    cmdHelp.Caption = DescArr(3)
    cmdOK.Caption = DescArr(2)
    cmdExit.Caption = DescArr(4)
End Sub

Private Sub Form_Load()
On Error GoTo errHdl
    Dim i As Integer
    Dim fOpenLang As Integer, strLang As String, NLang As String, fLang As String
    
    For i = 0 To Screen.FontCount - 1 Step 1
        cbofont.AddItem Screen.Fonts(i)
    Next
    LastLang = LngFile
    fLanguage.Path = App.Path & "\Language"
    fLanguage.Pattern = "*.nls"
    If fLanguage.ListCount <= 0 Then Exit Sub
    ReDim AlstLang(fLanguage.ListCount - 1)
    For i = 0 To fLanguage.ListCount - 1
        cboLanguage.AddItem Left(fLanguage.List(i), InStr(fLanguage.List(i), ".") - 1)
        NLang = ""
        fLang = ""
        AlstLang(i).FileName = fLanguage.List(i)
        fOpenLang = FreeFile
        Open App.Path & "\Language\" & AlstLang(i).FileName For Input As #fOpenLang
        Do While Not EOF(fOpenLang)
            DoEvents
            Line Input #fOpenLang, strLang
            If Left(strLang, 8) = "#99:001:" Then
                strLang = Right(strLang, Len(strLang) - 8)
                If Left(strLang, 2) = "01" Then
                    strLang = Right(strLang, Len(strLang) - InStr(strLang, ":"))
                    NLang = Trim(strLang)
                    AlstLang(i).Language = NLang
                ElseIf Left(strLang, 2) = "02" Then
                    strLang = Right(strLang, Len(strLang) - InStr(strLang, ":"))
                    fLang = Trim(strLang)
                    AlstLang(i).FontName = fLang
                End If
            End If
            If fLang <> "" And NLang <> "" Then Exit Do
        Loop
        Close #fOpenLang
    Next
    If cboLanguage.ListCount > 0 Then
        cboLanguage.ListIndex = 0
        For i = 0 To cboLanguage.ListCount - 1
            If InStr(LastLang, cboLanguage.List(i) & ".nls") Then
                cboLanguage.ListIndex = i
                Exit For
            End If
        Next
    End If
    Call cboLanguage_Change
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Form_Load"
End Sub
