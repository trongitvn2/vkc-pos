VERSION 5.00
Begin VB.Form frmOpenBook 
   Caption         =   "Më sæ "
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7575
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpenBook.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdexit 
      Height          =   855
      Left            =   3840
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "Tho¸t"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOpenBook.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdOpen 
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "&Më sæ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOpenBook.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Sæ b¸n hµng th¸ng nµy ch­a ®­îc më. B¹n cã muèn?"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   7455
   End
End
Attribute VB_Name = "frmOpenBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strDateTime As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOpen_Click()
    Dim fso As New FileSystemObject
    fso.CreateFolder WorkingFolder
    fso.CopyFile ReportFolder & "\database.mdb", WorkingFolder & "\Database.mdb", True
    Unload Me
    With frmLogin
        .Me_State = 1
        .Show vbModal
    End With
End Sub

Private Sub Form_Load()
    strDateTime = Format(Month(Date), "00") & Format(Year(Date), "00")
    
End Sub
