VERSION 5.00
Begin VB.Form frmPhimso 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPhimso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   7500
      Left            =   0
      TabIndex        =   1
      Top             =   1275
      Width           =   4335
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1165
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   210
         Width           =   1360
         _ExtentX        =   2408
         _ExtentY        =   2064
         BTYPE           =   6
         TX              =   "1"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1165
         Index           =   1
         Left            =   1490
         TabIndex        =   3
         Top             =   210
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   2064
         BTYPE           =   6
         TX              =   "2"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1170
         Index           =   2
         Left            =   2880
         TabIndex        =   4
         Top             =   210
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   2064
         BTYPE           =   6
         TX              =   "3"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1165
         Index           =   3
         Left            =   90
         TabIndex        =   5
         Top             =   1425
         Width           =   1360
         _ExtentX        =   2408
         _ExtentY        =   2064
         BTYPE           =   6
         TX              =   "4"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":0060
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1165
         Index           =   4
         Left            =   1490
         TabIndex        =   6
         Top             =   1425
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   2064
         BTYPE           =   6
         TX              =   "5"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":007C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1165
         Index           =   5
         Left            =   2880
         TabIndex        =   7
         Top             =   1425
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   2064
         BTYPE           =   6
         TX              =   "6"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":0098
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1165
         Index           =   6
         Left            =   90
         TabIndex        =   8
         Top             =   2625
         Width           =   1360
         _ExtentX        =   2408
         _ExtentY        =   2064
         BTYPE           =   6
         TX              =   "7"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":00B4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1165
         Index           =   7
         Left            =   1490
         TabIndex        =   9
         Top             =   2625
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   2064
         BTYPE           =   6
         TX              =   "8"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":00D0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1165
         Index           =   8
         Left            =   2880
         TabIndex        =   10
         Top             =   2625
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   2064
         BTYPE           =   6
         TX              =   "9"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":00EC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1165
         Index           =   9
         Left            =   90
         TabIndex        =   11
         Top             =   3840
         Width           =   1360
         _ExtentX        =   2408
         _ExtentY        =   2064
         BTYPE           =   6
         TX              =   "0"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":0108
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1165
         Index           =   10
         Left            =   1490
         TabIndex        =   12
         Top             =   3840
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   2064
         BTYPE           =   6
         TX              =   "00"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":0124
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   915
         Index           =   11
         Left            =   88
         TabIndex        =   13
         Top             =   5100
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1614
         BTYPE           =   6
         TX              =   "."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":0140
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1170
         Index           =   15
         Left            =   2880
         TabIndex        =   14
         Top             =   3840
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   2064
         BTYPE           =   6
         TX              =   "000"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":015C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdfree 
         Height          =   915
         Left            =   1485
         TabIndex        =   16
         Top             =   5100
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   1614
         BTYPE           =   6
         TX              =   "Free"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":0178
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAdd 
         Height          =   1365
         Left            =   2280
         TabIndex        =   18
         Top             =   6120
         Visible         =   0   'False
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   2408
         BTYPE           =   6
         TX              =   "+"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":0194
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdMinus 
         Height          =   1365
         Left            =   120
         TabIndex        =   19
         Top             =   6120
         Visible         =   0   'False
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2408
         BTYPE           =   6
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":01B0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1365
         Index           =   14
         Left            =   105
         TabIndex        =   20
         Tag             =   "L4"
         Top             =   6120
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   2408
         BTYPE           =   6
         TX              =   "&§ång ý"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmPhimso.frx":01CC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   60
      MaxLength       =   15
      TabIndex        =   0
      Top             =   525
      Width           =   2895
   End
   Begin prjTouchScreen.MyButton cmdAlpha 
      Height          =   750
      Index           =   12
      Left            =   2970
      TabIndex        =   15
      Tag             =   "L2"
      Top             =   510
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1323
      BTYPE           =   6
      TX              =   "&Xãa"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   255
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmPhimso.frx":01E8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblTitle 
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmPhimso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim formCallme As Integer
Dim Invoice_Number As Double
Dim Return_Valued As Double
Dim receiveMoney As Double
Dim isOK As Boolean

Private Sub cmdAdd_Click()
    On Error GoTo Handle
        isOK = True
            Return_Valued = CDbl("0" & txtQty.Text)
            Unload Me
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdAdd_Click"
End Sub

Private Sub cmdAlpha_Click(Index As Integer)
    Select Case Index
        Case 0 To 10, 15:
            If InStr(txtQty.Text, ".") > 0 Then
                txtQty.Text = txtQty.Text & cmdAlpha(Index).Caption
            Else
                txtQty.Text = Format(txtQty.Text & cmdAlpha(Index).Caption, "#,##0")
                txtQty.SelStart = Len(txtQty.Text)
            End If
        Case 11
            If InStr(txtQty.Text, ".") > 0 Then
                txtQty.Text = txtQty.Text
            Else
                txtQty.Text = txtQty.Text & cmdAlpha(Index).Caption
                txtQty.SelStart = Len(txtQty.Text)
            End If
        Case 12:
            txtQty.Text = ""
        
        Case 14:
            isOK = True
            If txtQty.Text = "" Then txtQty.Text = 0
            Select Case formCallme
                Case 2: 'So khach
                    Call Update_Person_Invoice(txtQty.Text)
                Case 3
                    If IsNumeric(txtQty.Text) Then
                        Return_Valued = txtQty.Text
                    Else
                        Return_Valued = 0
                    End If
               
                    
        End Select
        Unload Me
    End Select
End Sub

Private Sub cmdfree_Click()
    txtQty.Text = 100
    Call cmdAlpha_Click(14)
End Sub

Private Sub cmdMinus_Click()
On Error GoTo Handle
        isOK = True
        Return_Valued = -CDbl("0" & txtQty.Text)
        Unload Me
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdMinus_Click"
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    Dim ctrl As Control
    Dim DescArr() As String

    If cmdAlpha(14).Font.name <> CurFont Then Call Set_Language(Me, CurFont)
        DescArr = LoadLanguage(LngFile, "#03:023:")
        Me.Caption = DescArr(1)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
        Next ctrl
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"

End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    Return_Valued = 0
        With txtQty
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub

Public Property Let FormCall(ByVal vNewValue As Variant)
    formCallme = vNewValue
End Property

Public Sub Update_Person_Invoice(qty_Person As String)
On Error GoTo Handle
    Dim Person As Integer
    Dim rsInvoice_Person As New ADODB.Recordset
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsInvoice_Person = Open_Table(cnData, "Invoice_Totals_Person_Mapping")
    Person = CInt("0" & qty_Person)
    With rsInvoice_Person
        .Find "Invoice_Number=" & Invoice_Number, , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields("Invoice_Number") = Invoice_Number
            .Fields("Store_ID") = Store_ID
            .Fields("SeatNum") = Person
            .Update
        Else
            .Fields("Invoice_Number") = Invoice_Number
            .Fields("Store_ID") = Store_ID
            .Fields("SeatNum") = Person
            .Update
        End If
        
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Update_Person_Invoice"
End Sub


Public Property Let Get_Invoice_Num(ByVal vNewValue As Variant)
    Invoice_Number = vNewValue
End Property

Public Property Get Return_Value() As Variant
    Return_Value = Return_Valued
End Property

Public Property Let Return_Value(ByVal vNewValue As Variant)
    Return_Valued = vNewValue
End Property
Public Property Get Get_Receive_Money() As Variant
    Get_Receive_Money = receiveMoney
End Property

Public Property Let Get_Receive_Money(ByVal vNewValue As Variant)
    receiveMoney = vNewValue
End Property

Private Sub Form_Unload(Cancel As Integer)
    If Not isOK Then Return_Valued = 0
    Clipboard.Clear
End Sub

Private Sub txtQty_Change()
On Error GoTo Handle
    txtQty.Text = txtQty.Text
    txtQty.SelStart = Len(txtQty.Text)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtQty_Change"
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    Select Case KeyAscii
        Case 13
            Call cmdAlpha_Click(14)
        Case 8
        Case 45, 48 To 57
        Case Else:   KeyAscii = 0
    End Select
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtQty_KeyPress "
End Sub

