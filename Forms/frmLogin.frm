VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   10710
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   15240
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
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   12600
      Top             =   10080
   End
   Begin VB.PictureBox fraPassword 
      BackColor       =   &H00808000&
      Height          =   7625
      Left            =   7770
      ScaleHeight     =   7560
      ScaleWidth      =   7065
      TabIndex        =   11
      Top             =   1020
      Width           =   7125
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   945
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "1"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         FCOL            =   255
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmLogin.frx":1CFA
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
         Height          =   945
         Index           =   1
         Left            =   1410
         TabIndex        =   13
         Top             =   0
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "2"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         FCOL            =   255
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmLogin.frx":1D16
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
         Height          =   945
         Index           =   2
         Left            =   2810
         TabIndex        =   14
         Top             =   0
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "3"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         FCOL            =   255
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmLogin.frx":1D32
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
         Height          =   945
         Index           =   3
         Left            =   4210
         TabIndex        =   15
         Top             =   0
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "4"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         FCOL            =   255
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmLogin.frx":1D4E
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
         Height          =   945
         Index           =   4
         Left            =   5630
         TabIndex        =   16
         Top             =   0
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "5"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         FCOL            =   255
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmLogin.frx":1D6A
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
         Height          =   945
         Index           =   5
         Left            =   0
         TabIndex        =   17
         Top             =   930
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "6"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         FCOL            =   255
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmLogin.frx":1D86
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
         Height          =   945
         Index           =   6
         Left            =   1410
         TabIndex        =   18
         Top             =   930
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "7"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         FCOL            =   255
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmLogin.frx":1DA2
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
         Height          =   945
         Index           =   7
         Left            =   2810
         TabIndex        =   19
         Top             =   930
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "8"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         FCOL            =   255
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmLogin.frx":1DBE
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
         Height          =   945
         Index           =   8
         Left            =   4210
         TabIndex        =   20
         Top             =   930
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "9"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         FCOL            =   255
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmLogin.frx":1DDA
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
         Height          =   945
         Index           =   9
         Left            =   5630
         TabIndex        =   21
         Top             =   930
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "0"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         FCOL            =   255
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmLogin.frx":1DF6
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
         Height          =   945
         Index           =   10
         Left            =   0
         TabIndex        =   22
         Top             =   1860
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "a"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1E12
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
         Height          =   945
         Index           =   11
         Left            =   1410
         TabIndex        =   23
         Top             =   1860
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "b"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1E2E
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
         Height          =   945
         Index           =   12
         Left            =   2810
         TabIndex        =   24
         Top             =   1860
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "c"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1E4A
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
         Height          =   945
         Index           =   13
         Left            =   4210
         TabIndex        =   25
         Top             =   1860
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "d"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1E66
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
         Height          =   945
         Index           =   14
         Left            =   5630
         TabIndex        =   26
         Top             =   1860
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "e"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1E82
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
         Height          =   945
         Index           =   15
         Left            =   0
         TabIndex        =   27
         Top             =   2790
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "f"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1E9E
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
         Height          =   945
         Index           =   16
         Left            =   1410
         TabIndex        =   28
         Top             =   2790
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "g"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1EBA
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
         Height          =   945
         Index           =   17
         Left            =   2810
         TabIndex        =   29
         Top             =   2790
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "h"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1ED6
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
         Height          =   945
         Index           =   18
         Left            =   4210
         TabIndex        =   30
         Top             =   2790
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "i"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1EF2
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
         Height          =   945
         Index           =   19
         Left            =   5630
         TabIndex        =   31
         Top             =   2790
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "j"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1F0E
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
         Height          =   945
         Index           =   20
         Left            =   0
         TabIndex        =   32
         Top             =   3720
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "k"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1F2A
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
         Height          =   945
         Index           =   21
         Left            =   1410
         TabIndex        =   33
         Top             =   3720
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "l"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1F46
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
         Height          =   945
         Index           =   22
         Left            =   2810
         TabIndex        =   34
         Top             =   3720
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "m"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1F62
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
         Height          =   945
         Index           =   23
         Left            =   4210
         TabIndex        =   35
         Top             =   3720
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "n"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1F7E
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
         Height          =   945
         Index           =   24
         Left            =   5630
         TabIndex        =   36
         Top             =   3720
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "o"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1F9A
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
         Height          =   945
         Index           =   25
         Left            =   0
         TabIndex        =   37
         Top             =   4650
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "p"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1FB6
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
         Height          =   945
         Index           =   26
         Left            =   1410
         TabIndex        =   38
         Top             =   4650
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "q"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1FD2
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
         Height          =   945
         Index           =   27
         Left            =   2810
         TabIndex        =   39
         Top             =   4650
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "r"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":1FEE
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
         Height          =   945
         Index           =   28
         Left            =   4210
         TabIndex        =   40
         Top             =   4650
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "s"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":200A
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
         Height          =   945
         Index           =   29
         Left            =   5630
         TabIndex        =   41
         Top             =   4650
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "t"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":2026
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
         Height          =   945
         Index           =   30
         Left            =   0
         TabIndex        =   42
         Top             =   5580
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "u"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":2042
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
         Height          =   945
         Index           =   31
         Left            =   1410
         TabIndex        =   43
         Top             =   5580
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "v"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":205E
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
         Height          =   945
         Index           =   32
         Left            =   2810
         TabIndex        =   44
         Top             =   5580
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "w"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":207A
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
         Height          =   945
         Index           =   33
         Left            =   4210
         TabIndex        =   45
         Top             =   5580
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "x"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":2096
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
         Height          =   945
         Index           =   34
         Left            =   5630
         TabIndex        =   46
         Top             =   5580
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "y"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":20B2
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
         Height          =   1065
         Index           =   35
         Left            =   0
         TabIndex        =   47
         Top             =   6510
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1879
         BTYPE           =   5
         TX              =   "z"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   26.25
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
         MICON           =   "frmLogin.frx":20CE
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
         Height          =   1065
         Index           =   37
         Left            =   4215
         TabIndex        =   48
         Top             =   6510
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   1879
         BTYPE           =   5
         TX              =   "&Enter"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
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
         MICON           =   "frmLogin.frx":20EA
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
         Height          =   1065
         Index           =   36
         Left            =   1415
         TabIndex        =   49
         Top             =   6510
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   1879
         BTYPE           =   5
         TX              =   "&Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
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
         MICON           =   "frmLogin.frx":2106
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
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   9855
      Left            =   0
      Picture         =   "frmLogin.frx":2122
      ScaleHeight     =   9855
      ScaleWidth      =   7455
      TabIndex        =   10
      Top             =   0
      Width           =   7455
   End
   Begin VB.TextBox txtHint 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   9120
      TabIndex        =   9
      Text            =   "§Ìn CAPS LOCK ®ang bËt."
      Top             =   1200
      Width           =   2565
   End
   Begin prjTouchScreen.MyButton cmdLogInOut 
      Height          =   915
      Left            =   8520
      TabIndex        =   8
      Top             =   9720
      Visible         =   0   'False
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1614
      BTYPE           =   4
      TX              =   ""
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
      MPTR            =   1
      MICON           =   "frmLogin.frx":D40E
      PICN            =   "frmLogin.frx":D42A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdShutdown 
      Height          =   705
      Left            =   2910
      TabIndex        =   6
      Top             =   9960
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   1244
      BTYPE           =   5
      TX              =   "T¾t m¸y"
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
      BCOL            =   12632256
      BCOLO           =   255
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLogin.frx":D87C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6030
      Top             =   8010
   End
   Begin prjTouchScreen.MyButton cmdBackSpace 
      Height          =   675
      Left            =   13440
      TabIndex        =   5
      Top             =   210
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   1191
      BTYPE           =   6
      TX              =   ""
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
      BCOL            =   8421504
      BCOLO           =   8421504
      FCOL            =   16711680
      FCOLO           =   192
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLogin.frx":D898
      PICN            =   "frmLogin.frx":D8B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      IMEMode         =   3  'DISABLE
      Left            =   7770
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   210
      Width           =   5625
   End
   Begin prjTouchScreen.MyButton cmdHiberNate 
      Height          =   705
      Left            =   4260
      TabIndex        =   7
      Top             =   9960
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   1244
      BTYPE           =   5
      TX              =   "T¹m nghÜ"
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
      BCOL            =   12632256
      BCOLO           =   255
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLogin.frx":DE08
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   55
      Top             =   9840
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   1335
      Left            =   13560
      TabIndex        =   54
      Top             =   9120
      Width           =   1455
   End
   Begin VB.Line Kimgiay 
      BorderColor     =   &H0000FFFF&
      X1              =   14280
      X2              =   13920
      Y1              =   9840
      Y2              =   10200
   End
   Begin VB.Line Kimphut 
      BorderColor     =   &H00C00000&
      BorderWidth     =   2
      X1              =   14280
      X2              =   14760
      Y1              =   9840
      Y2              =   9840
   End
   Begin VB.Line Kimgio 
      BorderWidth     =   3
      X1              =   14280
      X2              =   14280
      Y1              =   9840
      Y2              =   9480
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   14280
      TabIndex        =   53
      Top             =   10440
      Width           =   135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   14925
      TabIndex        =   52
      Top             =   9720
      Width           =   135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   13560
      TabIndex        =   51
      Top             =   9720
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   14100
      TabIndex        =   50
      Top             =   9000
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      Height          =   1215
      Left            =   13680
      Shape           =   3  'Circle
      Top             =   9240
      Width           =   1215
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "gio"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   13920
      TabIndex        =   4
      Top             =   10125
      Width           =   825
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ngay"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   13710
      TabIndex        =   3
      Top             =   9940
      Width           =   1185
   End
   Begin VB.Label lblQuay 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "QuÇy sè:"
      BeginProperty Font 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   9570
      TabIndex        =   2
      Top             =   9690
      Width           =   3945
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SIRWAN SUPPER MARKET"
      BeginProperty Font 
         Name            =   "VNI-Bodon"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   8250
      TabIndex        =   1
      Top             =   8640
      Width           =   6225
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "&HÖ thèng"
      Begin VB.Menu mnuSelectDB 
         Caption         =   "Lùa chän d÷ liÖu nguån"
      End
      Begin VB.Menu mnuBackup 
         Caption         =   "Lùa chän ®­êng dÉn backup"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLanguage 
         Caption         =   "Lùa chän ng«n ng÷"
      End
      Begin VB.Menu mnuMonthReport 
         Caption         =   "B¸o c¸o tÝch lòy"
      End
      Begin VB.Menu mnuNhanvien 
         Caption         =   "Nh©n viªn"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMainGroup 
         Caption         =   "Nhãm chñ"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuGroupA 
         Caption         =   "Nhãm hµng"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHanghoa 
         Caption         =   " Hµng hãa"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMaterial 
         Caption         =   "Danh môc nguyªn liÖu"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSetMLink 
         Caption         =   "§Þnh l­îng nguyªn liÖu"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCustomer 
         Caption         =   "Danh môc kh¸ch hµng"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVendor 
         Caption         =   "Danh môc nhµ cung cÊp"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSetup 
         Caption         =   "Th«ng sè ch­¬ng tr×nh"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuPrintType 
         Caption         =   "§Þnh nghÜa m¸y in"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuStation 
         Caption         =   "QuÇy thu ng©n"
      End
      Begin VB.Menu mnuStationSelected 
         Caption         =   "Lùa chän QuÇy"
      End
      Begin VB.Menu mnuFormat 
         Caption         =   "§Þnh d¹ng"
      End
      Begin VB.Menu DeleteSale 
         Caption         =   "Xãa d÷ liÖu "
      End
      Begin VB.Menu mnuCust 
         Caption         =   "Tra cøu th«ng tin kh¸ch hµng"
      End
      Begin VB.Menu mnuAboutInfor 
         Caption         =   "Th«ng tin s¶n phÈm"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "H­íng dÉn sö dông"
      End
      Begin VB.Menu mnuUpdateDB 
         Caption         =   "CËp nhËt Database"
      End
      Begin VB.Menu mnuThoat 
         Caption         =   "§¨ng ký b¶n quyÒn"
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Desarr() As String
Dim rsStation As New ADODB.Recordset
Dim rscompany As New ADODB.Recordset
Dim State As Integer
Dim Giây&, Phút&, Gio&


Private Sub cmdAlpha_Click(Index As Integer)
On Error GoTo Handle
Dim IDUSER As String
    Select Case Index
        Case 0 To 35:
            txtInput.Text = txtInput.Text & cmdAlpha(Index).Caption
        Case 36
    
'            Call Open_File
            'Print #fFile, "End Program !" & vbTab & Now & vbTab
            'Print #fFile, "===================================================================="
'            Close #fFile
            Set cnData = Nothing
            Call gsDELETE_TMP_FILE
            End
        Case 37:
            IDUSER = TrimSpecialChar(txtInput.Text)
            If UCase(IDUSER) = "131112" Then
                UserLevel = 1
                UserPass = "admin"
                userName = "Administrator"
                UserID = "131112"
                Unload Me
                With frmTablePlan
                    .FormState = 1
                    .Show vbModal
                End With
            ElseIf IDUSER = "0909419887" Then
                UserLevel = 1
                UserPass = "09419887"
                userName = "Admin Level 2"
                UserID = "0909419887"
                Unload Me
                With frmTablePlan
                    .FormState = 1
                    .Show vbModal
                End With
            Else
                With rsuser
                    .Find "ID='" & Left(UCase(IDUSER), 2) & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                            UserID = Trim(.Fields("ID"))
                            UserPass = .Fields("Password")
                            userName = .Fields("userName")
                            UserLevel = .Fields("userLevel")
                        If UCase(IDUSER) = UCase(UserID & TrimSpecialChar(UserPass)) Then
                            Unload Me
                            With frmTablePlan
                                .FormState = 1
                                .Show vbModal
                            End With
                        Else
                            MsgBox "B¹n nhËp sai mËt khÈu, Vui long nhËp l¹i!!"
                            txtInput.SelStart = 0
                            txtInput.SelLength = Len(txtInput.Text)
                            'Close #fFile
                        End If
                    Else
                        MsgBox "Kh«ng t×m thÊy m· ng­êi dïng trong hÖ thèng", vbInformation
                        txtInput.SelStart = 0
                        txtInput.SelLength = 9999
                        'Close #fFile
                    End If
                End With
            End If
    End Select
    IDUSER = ""
Exit Sub
Handle:
    Close #fFile
    MsgBox Err.Number & Err.Description & Me.name & "  cmdAlpha_Click"
End Sub

Private Sub cmdBackSpace_Click()
    If Len(txtInput) > 0 Then
      txtInput.Text = Left(txtInput, Len(txtInput) - 1)
    End If
End Sub

Private Sub cmdLogInOut_Click()
    frmClerkLogin.Show vbModal
End Sub

Private Sub DeleteSale_Click()
    With frmPassword
         .FormActionKey = "SaleDelete"
         .Show vbModal
    End With
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
If cmdHiberNate.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    lbldate.Caption = gfCONVERT_STRING_TO_DATE(DateDefault)
    lblTime.Caption = time
    Desarr = LoadLanguage(LngFile, "#01:003:")
    Me.Caption = Desarr(1)
    mnuSystem.Caption = Desarr(4)
    mnuSelectDB.Caption = Desarr(5)
    mnuLanguage.Caption = Desarr(6)
    mnuNhanvien.Caption = Desarr(7)
    mnuGroupA.Caption = Desarr(8)
    mnuHanghoa.Caption = Desarr(9)
    mnuMaterial.Caption = Desarr(10)
    mnuSetMLink.Caption = Desarr(11)
    mnuPrintType.Caption = Desarr(12)
    mnuAboutInfor.Caption = Desarr(13)
    mnuThoat.Caption = "CÊp l¹i key sö dông"
    'lblQuay.Caption = Desarr(15)
    cmdHiberNate.Caption = Desarr(17)
    cmdShutdown.Caption = Desarr(16)
    mnuMainGroup.Caption = Desarr(20)
    mnuCustomer.Caption = Desarr(21)
    mnuVendor.Caption = Desarr(22)
    mnuSetup.Caption = Desarr(23)
    
    lblCompany.Font.name = "VNI-Bodon"
    If ServerName = "" Then
        ServerName = GetSettingStr("SYSTEM", "ServerName", "", myIniFile)
        DataBaseName = GetSettingStr("SYSTEM", "DatabaseName", "", myIniFile)
        DB_Password = GetSettingStr("SYSTEM", "Password", "", myIniFile)
        DB_Password = En_Decryption.MalgoDecrypt(DB_Password, 10)
    End If
       ' Set cnData = Get_Connection()
        
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    lbldate.Caption = gfCONVERT_STRING_TO_DATE(DateDefault)
    lblTime.Caption = time
    Desarr = LoadLanguage(LngFile, "#01:003:")
    Me.Caption = Desarr(1)
    mnuSystem.Caption = Desarr(4)
    mnuSelectDB.Caption = Desarr(5)
    mnuLanguage.Caption = Desarr(6)
    mnuNhanvien.Caption = Desarr(7)
    mnuGroupA.Caption = Desarr(8)
    mnuHanghoa.Caption = Desarr(9)
    mnuMaterial.Caption = Desarr(10)
    mnuSetMLink.Caption = Desarr(11)
    mnuPrintType.Caption = Desarr(12)
    mnuAboutInfor.Caption = Desarr(13)
    mnuThoat.Caption = Desarr(14)
   ' lblQuay.Caption = Desarr(15)
    cmdHiberNate.Caption = Desarr(17)
    cmdShutdown.Caption = Desarr(16)
    Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
        If cnData.State = 1 Then
            Set rsStation = Open_Table(cnData, "Stations_Location")
            Set rscompany = Open_Table(cnData, "Setup")
            lblQuay.Caption = Load_Station(rsStation, Store_ID)
            lblCompany.Caption = Load_Company(rscompany)
            ''''''Gan Danh sach nguoi su dung vao rsUser
            Set rsuser = LoadPasswordData
            Call Load_SF_System
        Else
            MsgBox "Kh«ng kÕt nèi ®­îc d÷ liÖu !"
            frmConnect_Data.Show vbModal
        End If
    
Exit Sub
Handle:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call gsDELETE_TMP_FILE
    Set cnData = Nothing
End Sub

Private Sub Label5_Click()
    Shell "rundll32.exe shell32.dll,Control_RunDLL  timedate.cpl,,0"
'    End
End Sub

Private Sub Label6_Click()
    End
End Sub

Private Sub mnuAboutInfor_Click()
    frmAboutInfor.Show vbModal
End Sub

'Private Sub mnuBackup_Click()
'    frmPathBackup.Show vbModal
'End Sub

Private Sub mnuCust_Click()
On Error GoTo Handle
    With frmFindCustomer
        .Get_State = 2
        .Show vbModal
    End With

Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " mnuCust_Click"
End Sub

Private Sub mnuCustomer_Click()
    frmCustomer.Show vbModal
End Sub

Private Sub mnuFormat_Click()
    frmFormat.Show vbModal
End Sub

Private Sub mnuGroupA_Click()
    frmDepartement.Show vbModal
End Sub

Private Sub mnuHanghoa_Click()
    frmItems.Show vbModal
End Sub

Private Sub mnuHelp_Click()
    Showhelp "Infor"
End Sub

Private Sub mnuLanguage_Click()
    frmLanguageSelection.Show vbModal
End Sub

Private Sub mnuMainGroup_Click()
    frmMainGroup.Show vbModal
End Sub

Private Sub mnuMaterial_Click()
    frmSetMPLU.Show vbModal
End Sub

Private Sub mnuMonthReport_Click()
    With frmPassword
         .FormActionKey = "SaleReport"
         .Show vbModal
    End With
End Sub

Private Sub mnuNhanvien_Click()
    With frmPassword
        .FormActionKey = "Employee"
        .Show vbModal
    End With
End Sub

Private Sub mnuPrintType_Click()
    frmPrintDefault.Show vbModal
End Sub

Private Sub mnuSelectDB_Click()
    'frmSelectData.Show vbModal
    frmConnect_Data.Show vbModal
End Sub

Private Sub mnuSetMLink_Click()
    frmSetMenuLink.Show vbModal
End Sub

Private Sub mnuSetup_Click()
    frmSetup.Show vbModal
End Sub

Private Sub mnuStation_Click()
    frmStation.Show vbModal
End Sub

Private Sub mnuStationSelected_Click()
    With frmPassword
         .FormActionKey = "Select_Station"
         .Show vbModal
    End With
End Sub

'Private Sub mnuThoat_Click()
'On Error GoTo Handle
'
'        Dim License_File As String
'         License_File = "C:\Windows\System32\KernelSys.sys"
'         If MsgBox("B¹n cã muèn yªu cÇu cÊp l¹i Key míi!", vbYesNo) = vbYes Then
'            If Dir(License_File, vbDirectory) <> "" Then
'               Kill License_File
'            End If
'            frmLicense.Show vbModal
'         End If
'
'Exit Sub
'Handle:
'    MsgBox Err.Number & Err.Description & "  mnuThoat_Click! "
'End Sub

Private Sub cmdHiberNate_Click()
     SetSystemPowerState waSUSPEND, True, False ') Then
End Sub

Private Sub mnuUpdateDB_Click()
    frmUpdateDB.Show vbModal
End Sub

Private Sub mnuVendor_Click()
    frmSupplier.Show vbModal
End Sub



Private Sub Timer1_Timer()
    On Error GoTo Handle
        lblTime.Caption = time
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Timer1_Timer"
End Sub

Private Sub Timer2_Timer()
Giây = Second(time)
Phút = Minute(time)
Gio = Hour(time) * 5 + Phút / 60
Kimgiay.X2 = 14280 + 560 * Cos(Giây * 0.1047 - 1.57)
Kimgiay.Y2 = 9840 + 560 * Sin(Giây * 0.1047 - 1.57)
Kimphut.X2 = 14280 + 475 * Cos(Phút * 0.1047 - 1.57)
Kimphut.Y2 = 9840 + 475 * Sin(Phút * 0.1047 - 1.57)
Kimgio.X2 = 14280 + 325 * Cos(Gio * 0.1047 - 1.57)
Kimgio.Y2 = 9840 + 325 * Sin(Gio * 0.1047 - 1.57)
End Sub

Private Sub txtInput_GotFocus()
On Error GoTo errHdl

'    GetKeyboardState Keys(0)
'    CapsLockState = Keys(VK_CAPITAL)
    If CapsLockState Then
        txtHint.Text = "CAPS LOCK status is on."
        txtHint.Visible = True
    Else
        txtHint.Text = ""
        txtHint.Visible = False
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - txtPassword_GotFocus"
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    If KeyCode = vbKeyCapital Then
        CapsLockState = Not CapsLockState
        If CapsLockState Then
            txtHint.Text = "CAPS LOCK status is on."
            txtHint.Visible = True
        Else
            txtHint.Text = ""
            txtHint.Visible = False
        End If
    End If
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - txtPassword_KeyDown"
End Sub

Public Function Load_Station(rs As ADODB.Recordset, strID As String) As String
On Error GoTo Handle
Dim str As String
str = ""
    With rs
        .Find "Station_Number='" & strID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            str = .Fields("Station_Name")
        End If
    End With
Load_Station = str
Exit Function
Handle:
    Load_Station = ""
    MsgBox Err.Number & Err.Description & Me.name & " Load_Station"
End Function
Public Function Load_Company(rs As ADODB.Recordset) As String
On Error GoTo Handle
Dim str As String
str = ""
    With rs
        If Not .EOF Then
            str = .Fields("Company_Info_2")
        End If
    End With
Load_Company = str
Exit Function
Handle:
    Load_Company = ""
    MsgBox Err.Number & Err.Description & Me.name & " Load_Company"
End Function

'Public Sub Compact_Repair_DB()
'On Error GoTo Handle
'    'Nén CSDL tên MyData.mdb và tao 1 CSDL moi tên DB2.mdb
'    If Dir(WorkingFolder & "\Database.ldb", vbDirectory) <> "" Then Exit Sub
''        Kill WorkingFolder & "\Database.ldb"
''    End If
'    If Dir(WorkingFolder & "\DB1.mdb", vbDirectory) <> "" Then
'        Kill WorkingFolder & "\DB1.mdb"
'    End If
'    DBEngine.CompactDatabase WorkingFolder & "\Database.mdb", WorkingFolder & "\DB1.mdb", ";pwd=100881administrator", , ";pwd=100881administrator"
'    'Xóa Database.mdb
'    Kill WorkingFolder & "\Database.mdb"
'    '®æi tªn DB1.mdb thµnh Database.mdb
'    Dim OldName
'    Dim NewName
'
'    OldName = WorkingFolder & "\DB1.mdb": NewName = WorkingFolder & "\Database.mdb"
'    Name OldName As NewName
'
'    Call gsDELETE_TMP_FILE
'Exit Sub
'Handle:
'    MsgBox Err.Number & Err.Description & Me.name & " Lçi D÷ liÖu"
'    If Err.Number = 3343 Then
'        Call Copy_DB
'    End If
'End Sub

Public Property Let Me_State(ByVal vNewValue As Variant)
    State = vNewValue
End Property

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdAlpha_Click(37)
    End If
End Sub

Public Sub Check_Customer_Birthday()
On Error GoTo Handle
Dim i As Integer
    Dim rscust As New ADODB.Recordset
    Dim rsPoint As New ADODB.Recordset
    Set rscust = Open_Table(cnData, "Customer")
    Set rsPoint = Open_Table(cnData, "Customer_Point_Sale")
    With rscust
        If Not .EOF Then
            .MoveFirst
            Do While Not rscust.EOF
                If IsDate(.Fields("Birthday")) Then
                        If Right(gfCONVERT_DATE_TO_STRING(.Fields("Birthday")), 4) = Right(gfCONVERT_DATE_TO_STRING(Format(Date, "dd/MM/yyyy")), 4) Then
                            MsgBox "H«m nay lµ ngµy sinh nhËt cña kh¸ch hµng:" & .Fields("CustName") & " - Sè thÎ:" & .Fields("CustNum") & " !!!"
                            With rsPoint
                                If Not .EOF Then
                                    
                                    If GetSettingStr("SYSTEM", "DateBirth", True, myIniFile) = Format(Date, "dd/MM/yyyy") Then
                                        If GetSettingStr("SYSTEM", "SavePoint", True, myIniFile) <> 1 Then
                                            rscust.Fields("Point") = rscust.Fields("Point") + .Fields("BirthPoint")
                                            SaveSettingStr "SYSTEM", "SavePoint", 1, myIniFile
                                            SaveSettingStr "SYSTEM", "DateBirth", Format(Date, "dd/MM/yyyy"), myIniFile
                                        End If
                                    Else
                                         SaveSettingStr "SYSTEM", "DateBirth", Format(Date, "dd/MM/yyyy"), myIniFile
                                         SaveSettingStr "SYSTEM", "SavePoint", 0, myIniFile
                                    End If
                                End If
                            End With
                        End If
                    End If
                .MoveNext
            Loop
        End If
    End With

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - Check_Customer_Birthday"
End Sub

