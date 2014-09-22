VERSION 5.00
Begin VB.Form frmLicen_Help 
   Caption         =   " "
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   6645
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   735
         Left            =   4800
         TabIndex        =   1
         Top             =   4440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   "&Close"
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
         MICON           =   "frmLicen_Help.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.Label Label5 
         Caption         =   "H©n h¹nh ®­îc phôc vô quý kh¸ch hµng!"
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   3720
         Width           =   5655
      End
      Begin VB.Label Label4 
         Caption         =   $"frmLicen_Help.frx":001C
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   480
         TabIndex        =   5
         Top             =   2160
         Width           =   5655
      End
      Begin VB.Label Label3 
         Caption         =   "2. Email M· ng­êi dïng vµ M· m¸y vÒ ®Þa chØ: can_vk@phucthanhvinh.com"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   5775
      End
      Begin VB.Label Label2 
         Caption         =   "1. Hç trî kü thuËt - CÊp License Key: 0918.655.887"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   5655
      End
      Begin VB.Label Label1 
         Caption         =   "Vui lßng liªn hÖ theo h­íng dÉn sau ®Ó ®­îc trî gióp"
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   6015
      End
   End
End
Attribute VB_Name = "frmLicen_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub
