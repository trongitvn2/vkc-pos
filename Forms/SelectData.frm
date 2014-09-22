VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSelectData 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chän d÷ liÖu"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4080
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdNext 
      Height          =   825
      Left            =   2880
      TabIndex        =   6
      Tag             =   "L6"
      Top             =   2940
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1455
      BTYPE           =   3
      TX              =   "&TiÕp tôc >>"
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
      MICON           =   "SelectData.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   825
      Left            =   750
      TabIndex        =   5
      Tag             =   "L5"
      Top             =   2940
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1455
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
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "SelectData.frx":001C
      PICN            =   "SelectData.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4590
      Top             =   2835
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraDulieu 
      Caption         =   "Chon d÷ liÖu cho ch­¬ng tr×nh"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Tag             =   "L1"
      Top             =   930
      Width           =   6045
      Begin VB.OptionButton optCreateNew 
         Caption         =   "T¹o míi c¬ së d÷ liÖu"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   495
         TabIndex        =   2
         Tag             =   "L4"
         Top             =   1035
         Width           =   3495
      End
      Begin VB.OptionButton optOldDatabase 
         Caption         =   "Chän d÷ liÖu cã s½n"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   480
         TabIndex        =   1
         Tag             =   "L3"
         Top             =   360
         Value           =   -1  'True
         Width           =   3615
      End
   End
   Begin VB.Label lblCurrentFolder 
      Caption         =   "Th­ môc hiÖn t¹i:"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   225
      TabIndex        =   4
      Tag             =   "L2"
      Top             =   225
      Width           =   1545
   End
   Begin VB.Label lblCurrent 
      Caption         =   "th­ môc"
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   1845
      TabIndex        =   3
      Top             =   210
      Width           =   4545
   End
End
Attribute VB_Name = "frmSelectData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 Option Explicit
 Dim DescArr() As String
Private Sub cmdCancel_Click()
    SaveSettingStr "SYSTEM", "Default Site", DefaultSite, myIniFile
    Unload Me
End Sub

Private Sub cmdNext_Click()
Dim fso As New FileSystemObject
Dim tmpSiteFile As String
If optOldDatabase.Value = True Then
    
   frmOpen.FileType = "Program File|*.Data"
    With frmOpen
        .Show vbModal
    End With
   
ElseIf optCreateNew.Value = True Then
    frmSave.FileType = "Program File|*.Data"
    With frmSave
        .Show , vbModal
    End With
    If frmSave.FileName <> "" Then
        SiteFile = frmSave.FileName
        If Dir(RemoveExtFile(SiteFile), vbDirectory) <> "" Then _
            fso.DeleteFolder RemoveExtFile(SiteFile), True
    End If
    
End If
cmdNext.Enabled = False
cmdCancel.Caption = DescArr(5)

End Sub

Private Sub Form_Activate()
On Error GoTo errHdl
    DescArr = LoadLanguage(LngFile, "#01:001:")
    If cmdCancel.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    fraDulieu.Caption = DescArr(1)
    lblCurrentFolder.Caption = DescArr(2)
    optOldDatabase.Caption = DescArr(3)
    optCreateNew.Caption = DescArr(4)
    cmdCancel.Caption = DescArr(5)
    cmdNext.Caption = DescArr(6)
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Form_Activate"
End Sub

Private Sub Form_Load()
    DescArr = LoadLanguage(LngFile, "#01:001:")
    lblCurrent.Caption = WorkingFolder
    cmdCancel.Caption = DescArr(5)
    cmdNext.Caption = DescArr(6)
End Sub
