VERSION 5.00
Begin VB.Form frmLicense_Actived 
   Caption         =   "Th«ng tin b¶n quyÒn"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   10065
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
   ScaleHeight     =   4815
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   855
      Left            =   3000
      TabIndex        =   0
      Top             =   3720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BTYPE           =   2
      TX              =   "MyButton1"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLicense_Actived.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblActived 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Th«ng tin ng­êi sö dông"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "Ngµy cÊp:"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
End
Attribute VB_Name = "frmLicense_Actived"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
