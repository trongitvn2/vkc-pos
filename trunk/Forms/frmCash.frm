VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCash1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "H×nh thøc thanh to¸n"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
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
   Icon            =   "frmCash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000B&
      Height          =   1365
      Left            =   30
      TabIndex        =   19
      Top             =   6840
      Width           =   7335
      Begin prjTouchScreen.MyButton cmdCheck 
         Height          =   1105
         Left            =   1920
         TabIndex        =   20
         Tag             =   "L4"
         Top             =   200
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   1958
         BTYPE           =   6
         TX              =   "&ChuyÓn kho¶n"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   16711680
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCash.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdGifCard 
         Height          =   1105
         Left            =   3720
         TabIndex        =   21
         Tag             =   "L5"
         Top             =   200
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   1958
         BTYPE           =   6
         TX              =   "&PhiÕu quµ tÆng"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   16711680
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCash.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdBalance 
         Height          =   1105
         Left            =   5530
         TabIndex        =   22
         Tag             =   "L6"
         Top             =   200
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   1958
         BTYPE           =   6
         TX              =   "C«ng &nî"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   16711680
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCash.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdCredit 
         Height          =   1105
         Left            =   120
         TabIndex        =   24
         Tag             =   "L7"
         Top             =   200
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   1958
         BTYPE           =   6
         TX              =   "ThÎ tÝn dông"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   16711680
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCash.frx":0060
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
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      Height          =   6825
      Left            =   150
      TabIndex        =   15
      Top             =   30
      Width           =   5145
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5145
         Left            =   60
         TabIndex        =   18
         Top             =   1470
         Width           =   4965
         _ExtentX        =   8758
         _ExtentY        =   9075
         _Version        =   393216
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   705
         Left            =   90
         TabIndex        =   17
         Top             =   630
         Width           =   4875
      End
      Begin VB.Label lblAmountList 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Caption         =   "H×nh thøc thanh to¸n"
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
         Height          =   375
         Left            =   30
         TabIndex        =   16
         Tag             =   "L2"
         Top             =   270
         Width           =   4395
      End
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   5370
      TabIndex        =   13
      Top             =   600
      Width           =   2805
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000B&
      Height          =   6885
      Left            =   5310
      TabIndex        =   0
      Top             =   1320
      Width           =   4215
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   995
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1746
         BTYPE           =   6
         TX              =   "1"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   27.75
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
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmCash.frx":007C
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
         Height          =   995
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1746
         BTYPE           =   6
         TX              =   "2"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   27.75
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
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmCash.frx":0098
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
         Height          =   995
         Index           =   2
         Left            =   2790
         TabIndex        =   3
         Top             =   210
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1746
         BTYPE           =   6
         TX              =   "3"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   27.75
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
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmCash.frx":00B4
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
         Height          =   995
         Index           =   3
         Left            =   90
         TabIndex        =   4
         Top             =   1220
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1746
         BTYPE           =   6
         TX              =   "4"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   27.75
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
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmCash.frx":00D0
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
         Height          =   995
         Index           =   4
         Left            =   1440
         TabIndex        =   5
         Top             =   1220
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1746
         BTYPE           =   6
         TX              =   "5"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   27.75
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
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmCash.frx":00EC
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
         Height          =   995
         Index           =   5
         Left            =   2790
         TabIndex        =   6
         Top             =   1220
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1746
         BTYPE           =   6
         TX              =   "6"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   27.75
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
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmCash.frx":0108
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
         Height          =   995
         Index           =   6
         Left            =   90
         TabIndex        =   7
         Top             =   2240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1746
         BTYPE           =   6
         TX              =   "7"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   27.75
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
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmCash.frx":0124
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
         Height          =   995
         Index           =   7
         Left            =   1440
         TabIndex        =   8
         Top             =   2240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1746
         BTYPE           =   6
         TX              =   "8"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   27.75
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
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmCash.frx":0140
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
         Height          =   995
         Index           =   8
         Left            =   2790
         TabIndex        =   9
         Top             =   2240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1746
         BTYPE           =   6
         TX              =   "9"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   27.75
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
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmCash.frx":015C
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
         Height          =   995
         Index           =   9
         Left            =   90
         TabIndex        =   10
         Top             =   3260
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1746
         BTYPE           =   6
         TX              =   "0"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   27.75
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
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmCash.frx":0178
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
         Height          =   995
         Index           =   10
         Left            =   1440
         TabIndex        =   11
         Top             =   3260
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1746
         BTYPE           =   6
         TX              =   "00"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   27.75
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
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmCash.frx":0194
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
         Height          =   995
         Index           =   11
         Left            =   2790
         TabIndex        =   12
         Top             =   3260
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1746
         BTYPE           =   6
         TX              =   "000"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   27.75
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
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmCash.frx":01B0
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
         Index           =   13
         Left            =   90
         TabIndex        =   23
         Tag             =   "L8"
         Top             =   4320
         Width           =   1960
         _ExtentX        =   3466
         _ExtentY        =   2064
         BTYPE           =   6
         TX              =   "&Tho¸t"
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
         BCOL            =   12632256
         BCOLO           =   16711680
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmCash.frx":01CC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdCash 
         Height          =   2505
         Left            =   2100
         TabIndex        =   26
         Tag             =   "L3"
         Top             =   4320
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   4419
         BTYPE           =   6
         TX              =   "&TiÒn mÆt"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
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
         BCOLO           =   16711680
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCash.frx":01E8
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
   Begin prjTouchScreen.MyButton cmdAlpha 
      Height          =   660
      Index           =   12
      Left            =   8205
      TabIndex        =   25
      Top             =   600
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   1164
      BTYPE           =   6
      TX              =   "Xãa"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12640511
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   -2147483638
      MPTR            =   1
      MICON           =   "frmCash.frx":0204
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblTenderAmount 
      BackColor       =   &H00C0C0C0&
      Caption         =   "NhËp sè tiÒn thanh to¸n"
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
      Height          =   465
      Left            =   5370
      TabIndex        =   14
      Tag             =   "L1"
      Top             =   210
      Width           =   3525
   End
End
Attribute VB_Name = "frmCash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iCash As Boolean
Dim Total, Totals As Double
Dim Payment_Method As String
Dim Customer As String
Dim DescArr() As String
Dim BillNO As Double
Dim rsInvoice_Items As New ADODB.Recordset
Dim rsInvoice_Onhold As New ADODB.Recordset
Dim rsInvoice_Total As New ADODB.Recordset
Dim rsInvoice_Notes As New ADODB.Recordset
Dim isActived As Boolean

Private Sub cmdAlpha_Click(Index As Integer)
    Select Case Index
        Case 0 To 11:
            txtQty.Text = Format(txtQty.Text & cmdAlpha(Index).Caption, "#,##0")
            txtQty.SelStart = Len(txtQty.Text)
        Case 12:
            txtQty.Text = ""
        Case 13:
            Unload Me
            iCash = False
        Unload Me
    End Select
End Sub


Public Property Let GetTotals(ByVal vNewValue As Variant)
    Totals = vNewValue
End Property
Public Property Let GetTotal(ByVal vNewValue As Variant)
    Total = vNewValue
End Property

Private Sub cmdBalance_Click()
    Payment_Method = "OA"
    If Customer = "101" Then
        MsgBox "Kh¸ch v·ng lai kh«ng ®­îc l­u vµo c«ng nî ", vbInformation
        Exit Sub
    Else
        If update_Balance(Customer) = False Then Exit Sub
        Call cmdCash_Click
    End If
End Sub

Private Sub cmdCash_Click()
On Error GoTo Handle
    Dim i As Double
    If CDbl(txtQty.Text) < CDbl(txtAmount.Text) Then
        MsgBox "Sè tiÒn b¹n nhËp nhá h¬n sè tiÒn trªn H§"
    Else
        iCash = True
        'Payment_Method = "CA"
            If gfUpdate_Invoice_Totals = True Then
                'If gfUpdate_Invoice_Itemized = False Then Exit Sub
                Call Update_Invoice_Notes
                If gfDelete_Invoice_Onhold = False Then Exit Sub
            Else
                Exit Sub
            End If
            i = txtQty.Text
         Unload Me
'         Thoat khoi giao dien ban hang
         With frmShowBillSale
            .GetBill = GetBillNo
            .Show vbModal
        End With

'        With frmChange
'            .GetTotal = Totals
'            .GetTender_Amt = i
'            .Show vbModal
'        End With
'        Unload Me
'        Unload frmCashMedia
'        Unload frmOrder
    End If
Exit Sub
Handle:
    Exit Sub
MsgBox Err.Number & Err.Description & Me.name & "  cmdCash_Click"
End Sub

Private Sub cmdCheck_Click()
    Payment_Method = "CH"
    Call cmdCash_Click
End Sub

Private Sub cmdCredit_Click()
    Payment_Method = "CC"
    Call cmdCash_Click
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
Dim ctrl As Control
If isActived = True Then Exit Sub
isActived = True
If cmdCash.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    DescArr = LoadLanguage(LngFile, "#02:003:")
    For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
        Next ctrl
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Activate"
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    iCash = False
    isActived = False
    Payment_Method = "CA"
'        txtQty.Text = Format(Total, formatNum)
'        txtAmount = Format(Total, formatNum)
    If cnData.State = 0 Then
        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    End If
    txtAmount.Locked = True
    Set rsInvoice_Onhold = OpenCriticalTable("select * from Invoice_OnHold", cnData)
    Set rsInvoice_Total = OpenCriticalTable("Select * from Invoice_Totals", cnData)
    Set rsInvoice_Notes = OpenCriticalTable("select * from Invoice_Totals_Notes", cnData)
    'Update gio Karaoke vao bang
        
    With rsInvoice_Notes
        .Find "Invoice_Number=" & BillNO, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If CDbl("0" & .Fields("Karaoke_Amount")) = 0 Then
                Call Update_Invoice_Notes
                Totals = Format(Totals + CDbl("0" & .Fields("Karaoke_Amount")), formatNum)
            End If
            txtQty.Text = Format(Totals, formatNum)
            txtAmount.Text = Format(Totals, formatNum)
        End If
    End With
    
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   Form_Load"
End Sub

Public Property Get GetCustomer() As Variant
    GetCustomer = Customer
End Property

Public Property Let GetCustomer(ByVal vNewValue As Variant)
    Customer = vNewValue
End Property

Public Property Get GetBillNo() As Variant
    GetBillNo = BillNO
End Property

Public Property Let GetBillNo(ByVal vNewValue As Variant)
   BillNO = vNewValue
End Property

Public Property Let Get_Payment_Method(ByVal vNewValue As Variant)
    Payment_Method = vNewValue
End Property


Public Sub Update_Invoice_Notes()
 On Error GoTo Handle
Dim rsLocation As New ADODB.Recordset
  
    With rsInvoice_Notes
    If .State = 0 Then Set rsInvoice_Notes = Open_Table(cnData, "Invoice_Totals_Notes")
      .Find "Invoice_Number=" & BillNO, , adSearchForward, adBookmarkFirst
          If Not .EOF Then
            If UCase(.Fields("ClosingTime")) = "C" Then
                    .Fields("ClosingTime") = DateDefault & Format(Now, "HH:mm:ss")
                    .Update
                End If
      End If
    End With
 
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Update_Invoice_Notes"
End Sub

Public Function gfUpdate_Invoice_Totals() As Boolean
On Error GoTo Handle
    gfUpdate_Invoice_Totals = False
    Set rsInvoice_Total = OpenCriticalTable("select * from Invoice_Totals", cnData)
        With rsInvoice_Total
            .Find "Invoice_Number=" & BillNO, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                !CustNum = Customer
                !Total_Price = Total
                !Grand_Total = Totals
                !Status = "C"
                !cashier_ID = UserID
                !Amt_Tendered = CDbl("0" & txtQty.Text)
                !Amt_Change = CDbl("0" & txtQty.Text) - Totals
                !Payment_Method = Payment_Method
                If Get_Cash_by_Time = True Then
                    !DateTime = DateDefault & Format(Now, "HH:mm:ss")
                End If
                rsInvoice_Total.Update
                .Requery
            End If
        End With
gfUpdate_Invoice_Totals = True
Exit Function

Handle:
    MsgBox Err.Number & Err.Description & Me.name & "gfupdate_Invoice_Totals"
    gfUpdate_Invoice_Totals = False
End Function

Public Function gfDelete_Invoice_Onhold() As Boolean
On Error GoTo Handle
    gfDelete_Invoice_Onhold = False
     With rsInvoice_Onhold
     If .State = 1 And .RecordCount > 0 Then
        .MoveFirst
     Else
        Exit Function
     End If
      .Find "Invoice_Number=" & BillNO, , adSearchForward, adBookmarkFirst
          If Not .EOF Then
              .Delete adAffectCurrent
          End If
    End With
    gfDelete_Invoice_Onhold = True
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " gfDelete_Invoice_Onhold"
    gfDelete_Invoice_Onhold = False
End Function


Private Sub txtQty_Change()
On Error GoTo Handle
    txtQty.Text = Format(txtQty.Text, "#,##0")
    txtQty.SelStart = Len(txtQty.Text)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtQty_Change"
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    Select Case KeyAscii
        Case 13
            Call cmdCash_Click
        Case 8
        Case 48 To 57
        Case Else:   KeyAscii = 0
    End Select
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtQty_KeyPress "
End Sub
'Luu no vao cong no khach hang
'Tham so truyen vao la ma khach hang

Public Function update_Balance(S As String) As Boolean
On Error GoTo Handle
Dim isUpdate As Boolean
    Dim rsCustomer As New ADODB.Recordset
    Dim strCus As String
    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    
    Set rsCustomer = Open_Table(cnData, "Customer")
    With rsCustomer
        If Not .EOF And .RecordCount > 0 Then .MoveFirst
        .Find "CustNum='" & S & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If CDbl("0" & .Fields("Acct_Balance")) + Totals >= CDbl("0" & .Fields("Acct_Max_Balance")) Then
                MsgBox " C«ng nî cña b¹n ®¹t ®Õn møc tèi ®a, vui lßng thanh tãan bít tr­íc khi ghi nî"
                isUpdate = False
            Else
                .Fields("Acct_Balance") = CDbl("0" & .Fields("Acct_Balance")) + Totals
                .Update
                isUpdate = True
            End If
        End If
    End With
    update_Balance = isUpdate
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name
End Function

Public Function Get_Cash_by_Time() As Boolean
On Error GoTo Handle
Dim iOpen As Boolean

    If ArrayFlag(SF(0), 7) = 1 Then iOpen = True
    Get_Cash_by_Time = iOpen
    
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Get_Cash_by_Time"

End Function

Public Property Get Return_Amt() As Variant
    Return_Amt = CDbl(txtQty.Text)
End Property

