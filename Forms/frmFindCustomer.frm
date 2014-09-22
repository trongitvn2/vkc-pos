VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmFindCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T×m kiÕm th«ng tin kh¸ch hµng"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
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
   Icon            =   "frmFindCustomer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraKey 
      Height          =   5745
      Left            =   90
      TabIndex        =   2
      Top             =   5400
      Width           =   15135
      Begin VB.TextBox txtFind 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7440
         TabIndex        =   0
         Top             =   270
         Width           =   5925
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   855
         Index           =   0
         Left            =   30
         TabIndex        =   3
         Top             =   1260
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "1"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   855
         Index           =   1
         Left            =   1200
         TabIndex        =   4
         Top             =   1260
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "2"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   855
         Index           =   2
         Left            =   2370
         TabIndex        =   5
         Top             =   1260
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "3"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   855
         Index           =   3
         Left            =   3540
         TabIndex        =   6
         Top             =   1260
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "4"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0060
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   855
         Index           =   4
         Left            =   4710
         TabIndex        =   7
         Top             =   1260
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "5"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":007C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   855
         Index           =   5
         Left            =   5880
         TabIndex        =   8
         Top             =   1260
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "6"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0098
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   855
         Index           =   6
         Left            =   7050
         TabIndex        =   9
         Top             =   1260
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "7"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":00B4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   855
         Index           =   7
         Left            =   8220
         TabIndex        =   10
         Top             =   1260
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "8"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":00D0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   855
         Index           =   8
         Left            =   9390
         TabIndex        =   11
         Top             =   1260
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "9"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":00EC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   855
         Index           =   9
         Left            =   10560
         TabIndex        =   12
         Top             =   1260
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "0"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0108
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   10
         Left            =   1470
         TabIndex        =   13
         Top             =   2190
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "q"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0124
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   11
         Left            =   2670
         TabIndex        =   14
         Top             =   2190
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "w"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0140
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   12
         Left            =   3870
         TabIndex        =   15
         Top             =   2190
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "e"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":015C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   13
         Left            =   5070
         TabIndex        =   16
         Top             =   2190
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "r"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0178
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   14
         Left            =   6270
         TabIndex        =   17
         Top             =   2190
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "t"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0194
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   15
         Left            =   7470
         TabIndex        =   18
         Top             =   2190
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "y"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":01B0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   16
         Left            =   8670
         TabIndex        =   19
         Top             =   2190
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "u"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":01CC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   17
         Left            =   9870
         TabIndex        =   20
         Top             =   2190
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "i"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":01E8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   18
         Left            =   11070
         TabIndex        =   21
         Top             =   2190
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "o"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0204
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   19
         Left            =   12270
         TabIndex        =   22
         Top             =   2190
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "p"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0220
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   20
         Left            =   1860
         TabIndex        =   23
         Top             =   3030
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "a"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":023C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   21
         Left            =   3150
         TabIndex        =   24
         Top             =   3030
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "s"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0258
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   22
         Left            =   4440
         TabIndex        =   25
         Top             =   3030
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "d"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0274
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   23
         Left            =   5730
         TabIndex        =   26
         Top             =   3030
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "f"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0290
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   24
         Left            =   7020
         TabIndex        =   27
         Top             =   3030
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "g"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":02AC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   25
         Left            =   8310
         TabIndex        =   28
         Top             =   3030
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "h"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":02C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   26
         Left            =   9600
         TabIndex        =   29
         Top             =   3030
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "j"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":02E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   27
         Left            =   10890
         TabIndex        =   30
         Top             =   3030
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "k"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0300
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   28
         Left            =   12180
         TabIndex        =   31
         Top             =   3030
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "l"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":031C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   29
         Left            =   2100
         TabIndex        =   32
         Top             =   3870
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "z"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0338
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   30
         Left            =   3270
         TabIndex        =   33
         Top             =   3870
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "x"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0354
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   31
         Left            =   4440
         TabIndex        =   34
         Top             =   3870
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "c"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0370
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   32
         Left            =   5610
         TabIndex        =   35
         Top             =   3870
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "v"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":038C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   33
         Left            =   6780
         TabIndex        =   36
         Top             =   3870
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "b"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":03A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   34
         Left            =   7950
         TabIndex        =   37
         Top             =   3870
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "n"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":03C4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   35
         Left            =   9120
         TabIndex        =   38
         Top             =   3870
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "m"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":03E0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   36
         Left            =   10290
         TabIndex        =   39
         Top             =   3870
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   ","
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":03FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   37
         Left            =   11340
         TabIndex        =   40
         Top             =   3870
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0418
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAddnew 
         Height          =   915
         Left            =   30
         TabIndex        =   41
         Tag             =   "L12"
         Top             =   4740
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1614
         BTYPE           =   1
         TX              =   "&Thªm"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0434
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdEdit 
         Height          =   915
         Left            =   1470
         TabIndex        =   42
         Tag             =   "L13"
         Top             =   4740
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1614
         BTYPE           =   1
         TX              =   "&Söa"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":0450
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   915
         Index           =   41
         Left            =   2970
         TabIndex        =   43
         Top             =   4740
         Width           =   6435
         _ExtentX        =   11351
         _ExtentY        =   1614
         BTYPE           =   1
         TX              =   "Space"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":046C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdEnter 
         Height          =   1785
         Index           =   48
         Left            =   13470
         TabIndex        =   44
         Top             =   3870
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   3149
         BTYPE           =   1
         TX              =   "Enter"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MICON           =   "frmFindCustomer.frx":0488
         PICN            =   "frmFindCustomer.frx":04A4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   38
         Left            =   12390
         TabIndex        =   45
         Top             =   3870
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "/"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":17B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   49
         Left            =   0
         TabIndex        =   46
         Top             =   2190
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "In th«ng tin KH"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":17D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   50
         Left            =   0
         TabIndex        =   47
         Top             =   3030
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "Caplock"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   18
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":17EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   51
         Left            =   0
         TabIndex        =   48
         Top             =   3870
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "Shift"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   18
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":180A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdBackSpace 
         Height          =   795
         Index           =   43
         Left            =   13410
         TabIndex        =   49
         Top             =   270
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1402
         BTYPE           =   3
         TX              =   "Back Space"
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
         BCOL            =   16578804
         BCOLO           =   16578804
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":1826
         PICN            =   "frmFindCustomer.frx":1842
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   855
         Index           =   42
         Left            =   11790
         TabIndex        =   51
         Top             =   1260
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "("
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":19D1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   855
         Index           =   43
         Left            =   12900
         TabIndex        =   52
         Top             =   1260
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   ")"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":19ED
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   44
         Left            =   13560
         TabIndex        =   53
         Top             =   2190
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "-"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":1A09
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   855
         Index           =   45
         Left            =   14010
         TabIndex        =   54
         Top             =   1260
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "&&"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8438015
         BCOLO           =   8438015
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":1A25
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   915
         Index           =   46
         Left            =   11340
         TabIndex        =   55
         Tag             =   "L15"
         Top             =   4740
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   1614
         BTYPE           =   1
         TX              =   "&Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":1A41
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   825
         Index           =   47
         Left            =   13470
         TabIndex        =   56
         Top             =   3030
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1455
         BTYPE           =   1
         TX              =   "_"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":1A5D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdText 
         Height          =   915
         Index           =   48
         Left            =   9450
         TabIndex        =   57
         Tag             =   "L14"
         Top             =   4740
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1614
         BTYPE           =   1
         TX              =   "&Clear"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFindCustomer.frx":1A79
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.Label lblTextFind 
         Caption         =   "NhËp th«ng tin cÇn t×m kiÕm( M· KH, Tªn KH, §Þa chØ, Sè §T...)"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   150
         TabIndex        =   50
         Tag             =   "L16"
         Top             =   330
         Width           =   7335
      End
   End
   Begin MSDataGridLib.DataGrid dtgCustomer 
      Height          =   5325
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   15165
      _ExtentX        =   26749
      _ExtentY        =   9393
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   28
      AllowAddNew     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFindCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FlagShift, FlagCaplock As Boolean
Dim rsCustomer As New ADODB.Recordset
Dim DescArr() As String
Dim strCustFind As String
Dim strFormcall As String
Dim Discount_Cust As Integer
Dim et As Integer
Dim Table_ID As String
Dim Amount  As Double
Dim CustID As String
Dim state_Call, Amount_Get_Point, Pnt As Integer


Private Sub cmdAddNew_Click()
    On Error GoTo Handle
        frmAddNewCust.Show vbModal
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdAddnew_Click"
End Sub

Private Sub cmdBackSpace_Click(Index As Integer)
    If Len(txtFind) > 0 Then
      txtFind.Text = Left(txtFind, Len(txtFind) - 1)
    End If
End Sub

Private Sub cmdEnter_Click(Index As Integer)
'If state_Call = 1 Then
    Select Case strFormcall
        Case "Delivery"
            Call Update_Delivery(CustNo(0))
        Case "CustomerSelect"
            dtgCustomer_DblClick
    End Select
    Unload Me
'End If
Unload Me
End Sub

Private Sub cmdText_Click(Index As Integer)
On Error GoTo Handle
Dim i As Integer
    Select Case Index
        Case 50:
            FlagCaplock = Not FlagCaplock
            If FlagCaplock = True Then
                For i = 10 To 35
                    cmdText(i).Caption = UCase(cmdText(i).Caption)
                Next
            Else
                For i = 10 To 35
                    cmdText(i).Caption = LCase(cmdText(i).Caption)
                Next
            End If
            'FlagCaplock = True
        Case 51:
            FlagShift = True
            For i = 10 To 35
                cmdText(i).Caption = UCase(cmdText(i).Caption)
            Next
        Case 41:
            txtFind.Text = txtFind.Text & Space(1)
        Case 45:
            txtFind.Text = txtFind.Text & Left(cmdText(Index).Caption, 1)
        
        Case 46
            Unload Me
        Case 48
            txtFind.Text = ""
        Case 49
            Call Print_Cust_Infor(CustID)
        Case Else:
        
            txtFind.Text = txtFind.Text & cmdText(Index).Caption
            If FlagShift = True Then
                If FlagCaplock = False Then
                    For i = 10 To 35
                        cmdText(i).Caption = LCase(cmdText(i).Caption)
                    Next
                End If
            End If
    End Select
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  cmdText_Click"
End Sub


Private Sub dtgCustomer_DblClick()
On Error GoTo Handle
Dim rsPromotion As New ADODB.Recordset
Dim Disc, adj1, adj2 As Integer
    Set rsPromotion = OpenCriticalTable("select * from Customer_Type where CustType_ID='" & dtgCustomer.Columns(6).Value & "'", cnData)
        With rsPromotion
            If Not .EOF Then
                Select Case .Fields("Promotion")
                    Case 0
                        
                    Case 1
                        Disc = .Fields("Pro_Value")
                    Case 2
                        adj1 = .Fields("Pro_Value")
                    Case 3
                        adj2 = .Fields("Pro_Value")
                End Select
            End If
        End With
        With frmOrder
            .Get_Adj1 = adj1
            .Get_Adj2 = adj2
            .Get_Discount = Disc
        End With
        CustNo(0) = dtgCustomer.Columns(0).Value
        CustNo(1) = dtgCustomer.Columns(1).Value
        Unload Me
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  dtgCustomer_DblClick"
End Sub

Public Function Get_Point() As Integer
On Error GoTo Handle
    Dim rsPoint As New ADODB.Recordset
    Dim kq, i As Integer
    Set rsPoint = Open_Table(cnData, "Customer_Point_Sale")
    With rsPoint
        If Not .EOF Then
            Amount_Get_Point = .Fields("Amount_Get_Point")
            Pnt = .Fields("Point")
        End If
    End With
    
    i = Int(Amount / Amount_Get_Point)
    Do Until i = 0
        kq = kq + Pnt
    i = i - 1
    Loop
    Get_Point = kq
Exit Function
Handle:
    Get_Point = 0
    MsgBox Err.Number & Err.Description & Me.name & " Get_Point"

End Function


Private Sub Form_Activate()
On Error GoTo Handle
Dim str As String
        If cmdEnter(48).Font.name <> CurFont Then Call Set_Language(Me, CurFont)
        Dim ctrl As Control
        str = "SELECT Customer.CustNum, Customer.CustName, Customer.Company, Customer.Address," & _
        " Customer.Phone, Customer.Fax, Customer.Cust_Type, Customer.TaxCode, Customer.AccountNo, " & _
        " Customer.Birthday,Acct_Balance,Customer.Totals,Customer.Point FROM Customer;"
        DescArr = LoadLanguage(LngFile, "#02:002:")
        Me.Caption = DescArr(1)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
        Next ctrl
        Set rsCustomer = OpenCriticalTable(str, cnData)
        If state_Call = 1 Then
            Call InitDatagrid(rsCustomer)
        Else
            Call Initgrid
        End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Activate"
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    Dim str As String
    str = "SELECT Customer.CustNum, Customer.CustName, Customer.Company, Customer.Address," & _
    " Customer.Phone, Customer.Fax, Customer.Cust_Type, Customer.TaxCode, Customer.AccountNo, " & _
    " Customer.Birthday,Acct_Balance,Customer.Totals,Customer.Point FROM Customer;"
'        If cnData.State = 0 Then
'            Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'        End If
        Set rsCustomer = OpenCriticalTable(str, cnData)
        
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    et = 0
    Set cnData = Nothing
End Sub

Private Sub txtFind_Change()
On Error GoTo errHdl
Dim rsCustGrid As New ADODB.Recordset
'On Error GoTo HandlEErr
With rsCustGrid
    .ActiveConnection = cnData
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .CursorType = adOpenDynamic
    '.Open
End With
'        Call RecreateVStockDB
        With rsCustGrid
            If .State = adStateOpen Then .Close
            
            If InStr(1, Trim(txtFind.Text), "*", vbTextCompare) > 0 Then GoTo 1
            
                .Open "SELECT  CustNum, CustName, Company, Address,Phone,Fax,Cust_Type,TaxCode,AccountNo,Birthday,Acct_Balance FROM Customer WHERE " & _
                 "CustNum LIKE '%" & txtFind.Text & "' OR  CustName LIKE '%" & txtFind.Text & "' OR Company LIKE '%" & txtFind.Text & "' OR  Address LIKE '%" & txtFind.Text & "'" & _
                "OR Phone LIKE '%" & txtFind.Text & "' OR  Fax LIKE '%" & txtFind.Text & "' OR TaxCode LIKE '" & txtFind.Text & "'" & _
                " OR AccountNo LIKE '%" & txtFind.Text & "' OR  Birthday LIKE '%" & txtFind.Text & "'  ORDER BY CustNum ASC"
            
            GoTo 2
1:
                .Open "SELECT  CustNum, CustName, Company, Address,Phone,Fax,Cust_Type,TaxCode,AccountNo,Birthday,Acct_Balance FROM Customer WHERE " & _
                "CustNum LIKE '" & Trim(txtFind.Text) & "%' OR CustName LIKE '" & _
                Left(Trim(txtFind.Text), Len(Trim(txtFind.Text)) - 1) & "%') OR Company LIKE '" & _
                Left(Trim(txtFind.Text), Len(Trim(txtFind.Text)) - 1) & "%')OR Address LIKE '" & _
                Left(Trim(txtFind.Text), Len(Trim(txtFind.Text)) - 1) & "%')OR Phone LIKE '" & _
                Left(Trim(txtFind.Text), Len(Trim(txtFind.Text)) - 1) & "%')OR Fax LIKE '" & _
                Left(Trim(txtFind.Text), Len(Trim(txtFind.Text)) - 1) & "%')OR TaxCode LIKE '" & _
                Left(Trim(txtFind.Text), Len(Trim(txtFind.Text)) - 1) & "%')OR AccountNo LIKE '" & _
                Left(Trim(txtFind.Text), Len(Trim(txtFind.Text)) - 1) & "%')OR Birthday LIKE '" & _
                Left(Trim(txtFind.Text), Len(Trim(txtFind.Text)) - 1) & "%')ORDER BY CustNum ASC"
        
2:
        End With
       
        Call InitDatagrid(rsCustGrid)
       
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "txtFind_Change"
End Sub

Public Sub Initgrid()
    On Error GoTo Handle
    With dtgCustomer
            .Columns(0).Caption = DescArr(2) '"Ma KH"
            .Columns(0).Width = 1500
            .Columns(1).Caption = DescArr(3) '"Ten KH"
            .Columns(1).Width = 3000
            .Columns(1).Alignment = dbgLeft
            
       End With
       Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Initgrid"
End Sub

Public Sub InitDatagrid(rs As ADODB.Recordset)
    On Error GoTo Handle
    With dtgCustomer
          Set .DataSource = rs
            .Columns(0).Caption = DescArr(2) '"Ma KH"
            .Columns(0).Width = 1500
            .Columns(1).Caption = DescArr(3) '"Ten KH"
            .Columns(1).Width = 3000
            .Columns(1).Alignment = dbgLeft
            .Columns(2).Caption = DescArr(4) '"Cty"
            .Columns(2).Width = 2500
            .Columns(2).Alignment = dbgLeft
            .Columns(3).Caption = DescArr(5) '"Dia chi"
            .Columns(3).Alignment = dbgLeft
            .Columns(3).Width = 3000
            .Columns(4).Caption = DescArr(6) ' "Dien thoai"
            .Columns(4).Alignment = dbgCenter
            .Columns(4).Width = 1600
            .Columns(5).Caption = DescArr(7) ' "Fax"
            .Columns(5).Alignment = dbgCenter
            .Columns(5).Width = 1600
            .Columns(6).Caption = DescArr(8) 'Chiet khau
            .Columns(6).Alignment = dbgLeft
            .Columns(6).Width = 1000
            .Columns(7).Caption = DescArr(9) ' "Ma So thue"
            .Columns(7).Alignment = dbgCenter
            .Columns(7).Width = 1600
            .Columns(8).Caption = DescArr(10) ' "So tai khoan
            .Columns(8).Alignment = dbgCenter
            .Columns(8).Width = 1600
            .Columns(9).Caption = DescArr(11) ' Ngay sinh nhat
            .Columns(9).Alignment = dbgCenter
            .Columns(9).Width = 1500
            .Columns(10).Caption = DescArr(17) ' Ngay sinh nhat
            .Columns(10).Alignment = dbgCenter
            .Columns(10).Width = 1500
       End With
       Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " InitDatagrid"
End Sub

Public Property Let FormCall(ByVal vNewValue As Variant)
    strFormcall = vNewValue
End Property
Public Sub Update_Delivery(S As String)
On Error GoTo Handle
    Dim MaxInvoice As Integer

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    If cnData.State <> 0 Then
    End If
    
    Table_ID = S
    MaxInvoice = GetMaxInvoice_Number
    SaveSettingStr "SYSTEM", "MaxInvoice", MaxInvoice, myIniFile
    
    If gfUpdate_Invoice_Totals(MaxInvoice) = True Then
        If gfUpdate_Invoice_OnHold(MaxInvoice) = False Then Exit Sub
        If gfUpdate_Invoice_Notes(MaxInvoice) = False Then Exit Sub
    Else
        Exit Sub
    End If
    
    With frmOrder
        .GetBill_Number = MaxInvoice
        .Get_Secion = "DE"
        .Get_Table_ID = Table_ID
        .Show vbModal
    End With
    'currentBill = ""
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Update_TakeOut"
End Sub


Public Function gfUpdate_Invoice_OnHold(Invoice As Integer) As Boolean
On Error GoTo Handle
    Dim rsinvoice_hold As New ADODB.Recordset
    gfUpdate_Invoice_OnHold = False
    
    Set rsinvoice_hold = OpenCriticalTable("select * from Invoice_OnHold ", cnData)
    With rsinvoice_hold
        .Find "OnHoldID='" & Table_ID & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                'Khong ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                currentBill = .Fields("Invoice_Number")
            Else
                ' ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                currentBill = Invoice
                .addNew
                .Fields("Invoice_Number") = Invoice
                .Fields("OnHoldID") = Table_ID
                .Fields("Cashier_ID") = UserID
                .Fields("Store_ID") = Store_ID
                .Fields("Occupied") = -1
                .Fields("Section_ID") = "DE"
                .Fields("Status") = 0
                .Update
            End If
    End With
    
gfUpdate_Invoice_OnHold = True
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " gfUpdate_Invoice_OnHold"
    gfUpdate_Invoice_OnHold = False
End Function

Public Function gfUpdate_Invoice_Totals(Invoice As Integer) As Boolean
On Error GoTo Handle
    Dim rsInvoice_Total As New ADODB.Recordset
    gfUpdate_Invoice_Totals = False
    
    Set rsInvoice_Total = OpenCriticalTable("select * from Invoice_Totals ", cnData)
    
    With rsInvoice_Total
        .Find "Invoice_Number=" & Invoice, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                Dim rscust As New ADODB.Recordset
                Set rscust = Open_Table(cnData, "Customer")
                    rscust.Find "CustNum='" & .Fields("CustNum") & "'", , adSearchForward, adBookmarkFirst
                    If Not rscust.EOF Then
                        CustNo(0) = .Fields("CustNum")
                        CustNo(1) = rscust!CustName
                        CustNo(2) = rscust!Acct_Balance
                        Discount_Cust = CDbl("0" & rscust.Fields("Discount"))
                    End If
                'Discount_Cust = .Fields("Discount")
            Else
                ' ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                .addNew
                .Fields("Invoice_Number") = Invoice
                .Fields("Store_ID") = Store_ID
                .Fields("CustNum") = CustNo(0)
                .Fields("DateTime") = DateDefault & Format(Now, "HH:mm:ss")
                .Fields("InvoiceNotesUsed") = -1
                .Fields("Status") = "O"
                .Fields("Station_ID") = "DE"
                .Fields("Cashier_ID") = UserID
                .Fields("Payment_MeThod") = "CA"
                .Fields("InvType") = 0
                .Fields("Orig_OnHoldID") = Trim(Table_ID)
'                .Fields("Tax_Rate_ID") = 0
                .Update
            End If
    End With
    
gfUpdate_Invoice_Totals = True
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " gfUpdate_Invoice_Totals"
    gfUpdate_Invoice_Totals = False
End Function


Public Function gfUpdate_Invoice_Notes(Invoice As Integer) As Boolean
On Error GoTo Handle
    Dim rsInvoice_Notes As New ADODB.Recordset
    gfUpdate_Invoice_Notes = False
    
    Set rsInvoice_Notes = Open_Table(cnData, "Invoice_Totals_Notes")
    With rsInvoice_Notes
        .Find "Invoice_Number=" & Invoice, , adSearchForward, adBookmarkFirst
            If .EOF Then
                ' ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                .addNew
                .Fields("Invoice_Number") = Invoice
                .Fields("Store_ID") = Store_ID
                .Fields("OpenTime") = DateDefault & Format(Now, "HH:mm:ss")
                .Fields("ClosingTime") = ""
                .Update
            End If
    End With
gfUpdate_Invoice_Notes = True
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " gfUpdate_Invoice_Notes"
    gfUpdate_Invoice_Notes = False
End Function



Public Property Get Get_Discount_Cust() As Variant
    Get_Discount_Cust = Discount_Cust
End Property


Public Sub Print_Cust_Infor(ByVal CustomerID As String)
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
    Dim iReport As CRAXDDRT.Report

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = "SELECT Customer.CustNum, Customer.CustName, Customer.Point, Customer.Address, Customer.Phone, Customer.Birthday, Customer.TaxCode" & _
          " FROM Customer Where CustNum = '" & CustomerID & "' order by Customer.CustName"
    Set crCustInfor = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crCustInfor
        .Database.AddADOCommand cnData, cmd
        .txtID.SetUnboundFieldSource "{ado.CustNum}"
        .txtName.SetUnboundFieldSource "{ado.CustName}"
        .txtAdd.SetUnboundFieldSource "{ado.Address}"
        .txtPhone.SetUnboundFieldSource "{ado.Phone}"
        .txtPoint.SetUnboundFieldSource "{ado.Point}"
        'canh le
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign
    End With
    Set iReport = crCustInfor
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = iReport
        .Show vbModal
    End With
Exit Sub
Handle:
'MsgBox Err.Number & Err.Description & Me.Name & " Print_Cust_Infor"
End Sub


Public Property Let Get_State(ByVal vNewValue As Variant)

End Property


Private Sub txtFind_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Len(txtFind.Text) > 4 Then
            et = et + 1
            If et = 2 Then
            txtFind.Text = Mid(txtFind.Text, 14, 4)
            
                    With rsCustomer
                    .Find "CustNum='" & txtFind.Text & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        CustNo(0) = .Fields("CustNum")
                        CustNo(1) = .Fields("CustName")
                        CustNo(2) = .Fields("Acct_Balance")
                        Discount_Cust = CDbl("0" & .Fields("Discount"))
                    End If
                    End With
                     Call cmdEnter_Click(48)
                 End If
'            End If
'Unload Me
        Else
            With rsCustomer
            .Find "CustNum='" & txtFind.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                CustNo(0) = .Fields("CustNum")
                CustNo(1) = .Fields("CustName")
                CustNo(2) = .Fields("Acct_Balance")
                Discount_Cust = CDbl("0" & .Fields("Discount"))
            End If
            End With
'             Call cmdEnter_Click(48)
               
        End If
    End If
End Sub


Public Property Let get_Amount(ByVal vNewValue As Variant)
    Amount = vNewValue
End Property
