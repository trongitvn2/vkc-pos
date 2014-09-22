VERSION 5.00
Begin VB.Form frmCal_select 
   Caption         =   "Lùa chän kho cÇn tÝnh"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8595
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
   ScaleHeight     =   2400
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   -120
      Width           =   8655
      Begin VB.OptionButton opttatca 
         Caption         =   "C¶ 2 kho"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   5880
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
      Begin VB.OptionButton optphu 
         Caption         =   "Kho phô"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   3240
         TabIndex        =   3
         Top             =   480
         Width           =   2415
      End
      Begin VB.OptionButton optchinh 
         Caption         =   "Kho chÝnh"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   2415
      End
   End
   Begin prjTouchScreen.MyButton cmdOK 
      Height          =   735
      Left            =   2640
      TabIndex        =   0
      Top             =   1560
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      BTYPE           =   2
      TX              =   "&TÝnh"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCal_select.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
End
Attribute VB_Name = "frmCal_select"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isOK As Integer

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    opttatca = True
End Sub

Private Sub optchinh_Click()
   isOK = 1
   optchinh = True
End Sub

Private Sub optphu_Click()
    isOK = 2
    optphu = True
    
End Sub

Private Sub opttatca_Click()
    isOK = 3
    opttatca = True
End Sub

Public Property Get Let_state() As Variant
    Let_state = isOK
End Property
