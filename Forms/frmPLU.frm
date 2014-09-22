VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmItems 
   Caption         =   "Items"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
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
   Icon            =   "frmPLU.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraForm 
      Height          =   11055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15165
      Begin prjTouchScreen.MyButton MyButton1 
         Height          =   375
         Left            =   15000
         TabIndex        =   134
         Top             =   9600
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   661
         BTYPE           =   2
         TX              =   "MyButton1"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPLU.frx":111EA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdUpdate 
         Height          =   1065
         Left            =   13440
         TabIndex        =   132
         Top             =   8700
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1879
         BTYPE           =   5
         TX              =   "CËp nhËt gi¸"
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
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPLU.frx":11206
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid flexPLU 
         Height          =   10815
         Left            =   0
         TabIndex        =   75
         Top             =   120
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   19076
         _Version        =   393216
         BackColorFixed  =   -2147483643
         BackColorBkg    =   -2147483643
         GridColor       =   8421504
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjTouchScreen.MyButton cmdSend 
         Height          =   1065
         Left            =   10260
         TabIndex        =   23
         Tag             =   "L3"
         Top             =   8700
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1879
         BTYPE           =   5
         TX              =   "&L­u"
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
         BCOLO           =   12640511
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPLU.frx":11222
         PICN            =   "frmPLU.frx":1123E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.PictureBox picLabel 
         BackColor       =   &H80000008&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8460
         ScaleHeight     =   675
         ScaleWidth      =   6495
         TabIndex        =   11
         Top             =   150
         Width           =   6555
         Begin VB.Label lblCode 
            Alignment       =   2  'Center
            BackColor       =   &H80000008&
            Caption         =   "Plu Code"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   1560
            TabIndex        =   13
            Top             =   0
            Width           =   2415
         End
         Begin VB.Label lblName 
            Alignment       =   2  'Center
            BackColor       =   &H80000008&
            Caption         =   "Plu Name"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   1560
            TabIndex        =   12
            Top             =   360
            Width           =   2415
         End
      End
      Begin TabDlg.SSTab tabPLU 
         Height          =   7485
         Left            =   8520
         TabIndex        =   1
         Top             =   1110
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   13203
         _Version        =   393216
         TabOrientation  =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Th«ng tin"
         TabPicture(0)   =   "frmPLU.frx":11782
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Flag"
         TabPicture(1)   =   "frmPLU.frx":1179E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "frmList(4)"
         Tab(1).Control(1)=   "frmList(3)"
         Tab(1).Control(2)=   "frmList(2)"
         Tab(1).Control(3)=   "frmList(1)"
         Tab(1).Control(4)=   "frmList(0)"
         Tab(1).Control(5)=   "picFlag"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "Other"
         TabPicture(2)   =   "frmPLU.frx":117BA
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame3"
         Tab(2).ControlCount=   1
         Begin VB.Frame frmList 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4605
            Index           =   4
            Left            =   -74430
            TabIndex        =   70
            Top             =   1140
            Width           =   5655
            Begin VB.PictureBox picList 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4335
               Index           =   4
               Left            =   80
               ScaleHeight     =   4275
               ScaleWidth      =   5460
               TabIndex        =   71
               Top             =   170
               Width           =   5520
               Begin VB.ListBox lstFlag 
                  Height          =   3435
                  Index           =   4
                  Left            =   90
                  Style           =   1  'Checkbox
                  TabIndex        =   73
                  Top             =   600
                  Width           =   5295
               End
               Begin VB.TextBox txtFlag 
                  Alignment       =   2  'Center
                  Height          =   375
                  Index           =   4
                  Left            =   2400
                  TabIndex        =   72
                  Tag             =   "13"
                  Text            =   "Text1"
                  Top             =   120
                  Width           =   735
               End
            End
         End
         Begin VB.Frame frmList 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4485
            Index           =   3
            Left            =   -74430
            TabIndex        =   66
            Top             =   1140
            Width           =   5655
            Begin VB.PictureBox picList 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3135
               Index           =   3
               Left            =   80
               ScaleHeight     =   3075
               ScaleWidth      =   5460
               TabIndex        =   67
               Top             =   170
               Width           =   5520
               Begin VB.ListBox lstFlag 
                  Height          =   2085
                  Index           =   3
                  Left            =   90
                  Style           =   1  'Checkbox
                  TabIndex        =   69
                  Top             =   600
                  Width           =   5295
               End
               Begin VB.TextBox txtFlag 
                  Alignment       =   2  'Center
                  Height          =   375
                  Index           =   3
                  Left            =   2400
                  TabIndex        =   68
                  Tag             =   "12"
                  Text            =   "Text1"
                  Top             =   120
                  Width           =   735
               End
            End
         End
         Begin VB.Frame frmList 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Index           =   2
            Left            =   -74430
            TabIndex        =   62
            Top             =   1050
            Width           =   5655
            Begin VB.PictureBox picList 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4425
               Index           =   2
               Left            =   80
               ScaleHeight     =   4365
               ScaleWidth      =   5460
               TabIndex        =   63
               Top             =   170
               Width           =   5520
               Begin VB.ListBox lstFlag 
                  Height          =   3660
                  Index           =   2
                  Left            =   90
                  Style           =   1  'Checkbox
                  TabIndex        =   65
                  Top             =   600
                  Width           =   5295
               End
               Begin VB.TextBox txtFlag 
                  Alignment       =   2  'Center
                  Height          =   375
                  Index           =   2
                  Left            =   2400
                  TabIndex        =   64
                  Tag             =   "11"
                  Text            =   "Text1"
                  Top             =   120
                  Width           =   735
               End
            End
         End
         Begin VB.Frame frmList 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Index           =   1
            Left            =   -74430
            TabIndex        =   58
            Top             =   1050
            Width           =   5655
            Begin VB.PictureBox picList 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4455
               Index           =   1
               Left            =   80
               ScaleHeight     =   4395
               ScaleWidth      =   5460
               TabIndex        =   59
               Top             =   170
               Width           =   5520
               Begin VB.ListBox lstFlag 
                  Height          =   3660
                  Index           =   1
                  Left            =   90
                  Style           =   1  'Checkbox
                  TabIndex        =   61
                  Top             =   600
                  Width           =   5295
               End
               Begin VB.TextBox txtFlag 
                  Alignment       =   2  'Center
                  Height          =   375
                  Index           =   1
                  Left            =   2400
                  TabIndex        =   60
                  Tag             =   "10"
                  Text            =   "Text1"
                  Top             =   120
                  Width           =   735
               End
            End
         End
         Begin VB.Frame frmList 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4695
            Index           =   0
            Left            =   -74430
            TabIndex        =   54
            Top             =   1050
            Width           =   5655
            Begin VB.PictureBox picList 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4455
               Index           =   0
               Left            =   90
               ScaleHeight     =   4395
               ScaleWidth      =   5460
               TabIndex        =   55
               Top             =   180
               Width           =   5520
               Begin VB.TextBox txtFlag 
                  Alignment       =   2  'Center
                  Height          =   375
                  Index           =   0
                  Left            =   2400
                  TabIndex        =   57
                  Tag             =   "9"
                  Text            =   "Text1"
                  Top             =   120
                  Width           =   735
               End
               Begin VB.ListBox lstFlag 
                  Height          =   3660
                  Index           =   0
                  Left            =   90
                  Style           =   1  'Checkbox
                  TabIndex        =   56
                  Top             =   570
                  Width           =   5295
               End
            End
         End
         Begin VB.Frame Frame3 
            Height          =   5895
            Left            =   -74490
            TabIndex        =   22
            Top             =   210
            Width           =   5805
            Begin prjTouchScreen.MyButton cmdRemovePic 
               Height          =   465
               Left            =   2640
               TabIndex        =   74
               Top             =   5040
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   820
               BTYPE           =   14
               TX              =   "Xãa h×nh ¶nh"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   ".VnArial"
                  Size            =   8.25
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
               MICON           =   "frmPLU.frx":117D6
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               Value           =   0   'False
            End
            Begin MSComDlg.CommonDialog comdlg 
               Left            =   4920
               Top             =   210
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.Frame fraImage 
               BackColor       =   &H00C0E0FF&
               Height          =   3675
               Left            =   720
               TabIndex        =   30
               Top             =   1320
               Width           =   4005
               Begin VB.Image Image1 
                  BorderStyle     =   1  'Fixed Single
                  Height          =   3465
                  Left            =   60
                  Picture         =   "frmPLU.frx":117F2
                  Stretch         =   -1  'True
                  Top             =   150
                  Width           =   3885
               End
            End
            Begin prjTouchScreen.MyButton cmd 
               Height          =   465
               Left            =   870
               TabIndex        =   29
               Tag             =   "L29"
               Top             =   5040
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   820
               BTYPE           =   14
               TX              =   "Chän h×nh ¶nh"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   ".VnArial"
                  Size            =   8.25
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
               MICON           =   "frmPLU.frx":1C1E9
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               Value           =   0   'False
            End
            Begin prjTouchScreen.MyButton cmdModifierPickup 
               Height          =   465
               Left            =   4830
               TabIndex        =   28
               Top             =   810
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   820
               BTYPE           =   14
               TX              =   ""
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   ".VnArial"
                  Size            =   8.25
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
               MICON           =   "frmPLU.frx":1C205
               PICN            =   "frmPLU.frx":1C221
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               Value           =   0   'False
            End
            Begin prjTouchScreen.MyButton cmdCapture 
               Height          =   4185
               Left            =   4830
               TabIndex        =   31
               Tag             =   "L30"
               Top             =   1350
               Width           =   555
               _ExtentX        =   979
               _ExtentY        =   7382
               BTYPE           =   14
               TX              =   "Chôp h×nh"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   ".VnArial"
                  Size            =   8.25
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
               MICON           =   "frmPLU.frx":1DF2B
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
         Begin VB.PictureBox picFlag 
            BackColor       =   &H8000000E&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   -74340
            ScaleHeight     =   435
            ScaleWidth      =   4830
            TabIndex        =   17
            Top             =   510
            Width           =   4890
            Begin VB.OptionButton optFlag 
               BackColor       =   &H80000016&
               Caption         =   "PF-5"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   4
               Left            =   3795
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   30
               Width           =   885
            End
            Begin VB.OptionButton optFlag 
               BackColor       =   &H80000016&
               Caption         =   "PF-4"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   3
               Left            =   2865
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   30
               Width           =   885
            End
            Begin VB.OptionButton optFlag 
               BackColor       =   &H80000016&
               Caption         =   "PF-3"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   2
               Left            =   1935
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   30
               Width           =   885
            End
            Begin VB.OptionButton optFlag 
               BackColor       =   &H80000016&
               Caption         =   "PF-2"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   1
               Left            =   1020
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   30
               Width           =   885
            End
            Begin VB.OptionButton optFlag 
               BackColor       =   &H80000016&
               Caption         =   "PF-1"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   0
               Left            =   105
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   30
               Width           =   885
            End
         End
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6975
            Left            =   375
            TabIndex        =   2
            Top             =   165
            Width           =   6045
            Begin prjTouchScreen.MyButton cmdDept 
               Height          =   435
               Left            =   4800
               TabIndex        =   131
               Top             =   1725
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   767
               BTYPE           =   14
               TX              =   "..."
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   ".VnArial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   2
               FOCUSR          =   -1  'True
               BCOL            =   14215660
               BCOLO           =   14215660
               FCOL            =   16711680
               FCOLO           =   16711680
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmPLU.frx":1DF47
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               Value           =   0   'False
            End
            Begin VB.CommandButton cmdSetcolor 
               Caption         =   "Set mµu"
               Height          =   615
               Left            =   3480
               TabIndex        =   130
               Top             =   6000
               Visible         =   0   'False
               Width           =   1575
            End
            Begin prjTouchScreen.MyButton cmdColor 
               Height          =   375
               Left            =   2400
               TabIndex        =   129
               Top             =   6120
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   661
               BTYPE           =   5
               TX              =   "Color"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   ".VnArial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   2
               FOCUSR          =   -1  'True
               BCOL            =   33023
               BCOLO           =   33023
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmPLU.frx":1DF63
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               Value           =   0   'False
            End
            Begin VB.Frame Frame5 
               Caption         =   "Basic Colors"
               Height          =   2415
               Left            =   1920
               TabIndex        =   79
               Top             =   2280
               Visible         =   0   'False
               Width           =   3975
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H008080FF&
                  Height          =   285
                  Index           =   0
                  Left            =   120
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   127
                  Tag             =   "ED"
                  Top             =   210
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H0080FFFF&
                  Height          =   285
                  Index           =   1
                  Left            =   600
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   126
                  Tag             =   "FD"
                  Top             =   210
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H0080FF80&
                  Height          =   285
                  Index           =   2
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   125
                  Tag             =   "7D"
                  Top             =   210
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H0080FF00&
                  Height          =   285
                  Index           =   3
                  Left            =   1560
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   124
                  Tag             =   "1D"
                  Top             =   210
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00FFFF80&
                  Height          =   285
                  Index           =   4
                  Left            =   2040
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   123
                  Tag             =   "7F"
                  Top             =   210
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00FF8000&
                  Height          =   285
                  Index           =   5
                  Left            =   2520
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   122
                  Tag             =   "0F"
                  Top             =   210
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00C080FF&
                  Height          =   285
                  Index           =   6
                  Left            =   3000
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   121
                  Tag             =   "EE"
                  Top             =   210
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00FF80FF&
                  Height          =   285
                  Index           =   7
                  Left            =   3480
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   120
                  Tag             =   "EF"
                  Top             =   210
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H000000FF&
                  Height          =   285
                  Index           =   8
                  Left            =   120
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   119
                  Tag             =   "E0"
                  Top             =   600
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H0000FFFF&
                  Height          =   285
                  Index           =   9
                  Left            =   600
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   118
                  Tag             =   "FC"
                  Top             =   600
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H0000FF80&
                  Height          =   285
                  Index           =   10
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   117
                  Tag             =   "7C"
                  Top             =   600
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H0040FF00&
                  Height          =   285
                  Index           =   11
                  Left            =   1560
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   116
                  Tag             =   "1C"
                  Top             =   600
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00FFFF00&
                  Height          =   285
                  Index           =   12
                  Left            =   2040
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   115
                  Tag             =   "1F"
                  Top             =   600
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00C08000&
                  Height          =   285
                  Index           =   13
                  Left            =   2520
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   114
                  Tag             =   "0E"
                  Top             =   600
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00C08080&
                  Height          =   285
                  Index           =   14
                  Left            =   3000
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   113
                  Tag             =   "6E"
                  Top             =   600
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00FF00FF&
                  Height          =   285
                  Index           =   15
                  Left            =   3480
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   112
                  Tag             =   "E3"
                  Top             =   600
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00404080&
                  Height          =   285
                  Index           =   16
                  Left            =   120
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   111
                  Tag             =   "64"
                  Top             =   960
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H004080FF&
                  Height          =   285
                  Index           =   17
                  Left            =   600
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   110
                  Tag             =   "EC"
                  Top             =   960
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H0000FF00&
                  Height          =   285
                  Index           =   18
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   109
                  Tag             =   "1C"
                  Top             =   960
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00808000&
                  Height          =   285
                  Index           =   19
                  Left            =   1560
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   108
                  Tag             =   "0D"
                  Top             =   960
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00804000&
                  Height          =   285
                  Index           =   20
                  Left            =   2040
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   107
                  Tag             =   "05"
                  Top             =   960
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00FF8080&
                  Height          =   285
                  Index           =   21
                  Left            =   2520
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   106
                  Tag             =   "6F"
                  Top             =   960
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00400080&
                  Height          =   285
                  Index           =   22
                  Left            =   3000
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   105
                  Tag             =   "60"
                  Top             =   960
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H008000FF&
                  Height          =   285
                  Index           =   23
                  Left            =   3480
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   104
                  Tag             =   "E1"
                  Top             =   960
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00000080&
                  Height          =   285
                  Index           =   24
                  Left            =   120
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   103
                  Tag             =   "60"
                  Top             =   1320
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H000080FF&
                  Height          =   285
                  Index           =   25
                  Left            =   600
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   102
                  Tag             =   "EC"
                  Top             =   1320
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00008000&
                  Height          =   285
                  Index           =   26
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   101
                  Tag             =   "0C"
                  Top             =   1320
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00408000&
                  Height          =   285
                  Index           =   27
                  Left            =   1560
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   100
                  Tag             =   "0C"
                  Top             =   1320
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00FF0000&
                  Height          =   285
                  Index           =   28
                  Left            =   2040
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   99
                  Tag             =   "03"
                  Top             =   1320
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00A00000&
                  Height          =   285
                  Index           =   29
                  Left            =   2520
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   98
                  Tag             =   "01"
                  Top             =   1320
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00800080&
                  Height          =   285
                  Index           =   30
                  Left            =   3000
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   97
                  Tag             =   "61"
                  Top             =   1320
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00FF0080&
                  Height          =   285
                  Index           =   31
                  Left            =   3480
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   96
                  Tag             =   "63"
                  Top             =   1320
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00000040&
                  Height          =   285
                  Index           =   32
                  Left            =   120
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   95
                  Tag             =   "20"
                  Top             =   1680
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00004080&
                  Height          =   285
                  Index           =   33
                  Left            =   600
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   94
                  Tag             =   "64"
                  Top             =   1680
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00004000&
                  Height          =   285
                  Index           =   34
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   93
                  Tag             =   "04"
                  Top             =   1680
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00404000&
                  Height          =   285
                  Index           =   35
                  Left            =   1560
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   92
                  Tag             =   "04"
                  Top             =   1680
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00800000&
                  Height          =   285
                  Index           =   36
                  Left            =   2040
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   91
                  Tag             =   "01"
                  Top             =   1680
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00400000&
                  Height          =   285
                  Index           =   37
                  Left            =   2520
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   90
                  Tag             =   "00"
                  Top             =   1680
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00400040&
                  Height          =   285
                  Index           =   38
                  Left            =   3000
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   89
                  Tag             =   "20"
                  Top             =   1680
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00800040&
                  Height          =   285
                  Index           =   39
                  Left            =   3480
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   88
                  Tag             =   "21"
                  Top             =   1680
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00000000&
                  Height          =   285
                  Index           =   40
                  Left            =   120
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   87
                  Tag             =   "00"
                  Top             =   2040
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00008080&
                  Height          =   285
                  Index           =   41
                  Left            =   600
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   86
                  Tag             =   "6C"
                  Top             =   2040
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00408080&
                  Height          =   285
                  Index           =   42
                  Left            =   1080
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   85
                  Tag             =   "6C"
                  Top             =   2040
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00808080&
                  Height          =   285
                  Index           =   43
                  Left            =   1560
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   84
                  Tag             =   "6D"
                  Top             =   2040
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00808040&
                  Height          =   285
                  Index           =   44
                  Left            =   2040
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   83
                  Tag             =   "2D"
                  Top             =   2040
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00C0C0C0&
                  Height          =   285
                  Index           =   45
                  Left            =   2520
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   82
                  Tag             =   "B6"
                  Top             =   2040
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00400040&
                  Height          =   285
                  Index           =   46
                  Left            =   3000
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   81
                  Tag             =   "20"
                  Top             =   2040
                  Width           =   375
               End
               Begin VB.PictureBox picBasicColor 
                  BackColor       =   &H00FFFFFF&
                  Height          =   285
                  Index           =   47
                  Left            =   3480
                  ScaleHeight     =   225
                  ScaleWidth      =   315
                  TabIndex        =   80
                  Tag             =   "FF"
                  Top             =   2040
                  Width           =   375
               End
               Begin VB.Label lblMove 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Label1"
                  ForeColor       =   &H80000008&
                  Height          =   320
                  Left            =   105
                  TabIndex        =   128
                  Top             =   195
                  Width           =   400
               End
               Begin VB.Shape shpMove 
                  BorderColor     =   &H00404040&
                  BorderStyle     =   3  'Dot
                  DrawMode        =   14  'Copy Pen
                  Height          =   370
                  Left            =   80
                  Shape           =   4  'Rounded Rectangle
                  Top             =   160
                  Width           =   460
               End
            End
            Begin VB.TextBox txtPLU 
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   13
               Left            =   1680
               MaxLength       =   40
               TabIndex        =   76
               Tag             =   "15"
               Top             =   840
               Width           =   3555
            End
            Begin VB.Frame Frame4 
               Caption         =   "Everning Price"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   2955
               Left            =   3960
               TabIndex        =   46
               Tag             =   "L28"
               Top             =   2400
               Width           =   1905
               Begin VB.TextBox txtPLU 
                  BeginProperty Font 
                     Name            =   ".VnArial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   8
                  Left            =   30
                  MaxLength       =   10
                  TabIndex        =   49
                  Tag             =   "10"
                  Top             =   1380
                  Width           =   1620
               End
               Begin VB.TextBox txtPLU 
                  BeginProperty Font 
                     Name            =   ".VnArial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   7
                  Left            =   30
                  MaxLength       =   10
                  TabIndex        =   48
                  Tag             =   "9"
                  Top             =   630
                  Width           =   1620
               End
               Begin VB.TextBox txtPLU 
                  BeginProperty Font 
                     Name            =   ".VnArial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   9
                  Left            =   30
                  MaxLength       =   10
                  TabIndex        =   47
                  Tag             =   "11"
                  Top             =   2160
                  Width           =   1620
               End
               Begin VB.Label lblPrice 
                  Caption         =   "Evering Price &2 :"
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
                  Index           =   9
                  Left            =   120
                  TabIndex        =   52
                  Tag             =   "L20"
                  Top             =   1140
                  Width           =   1545
               End
               Begin VB.Label lblPrice 
                  Caption         =   "Everning Price &1:"
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
                  Index           =   8
                  Left            =   90
                  TabIndex        =   51
                  Tag             =   "L19"
                  Top             =   360
                  Width           =   1515
               End
               Begin VB.Label lblPrice 
                  Caption         =   "Everning Price &3:"
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
                  Index           =   7
                  Left            =   120
                  TabIndex        =   50
                  Tag             =   "L21"
                  Top             =   1920
                  Width           =   1515
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Happy Hour Price"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   2955
               Left            =   1980
               TabIndex        =   39
               Tag             =   "L27"
               Top             =   2400
               Width           =   1905
               Begin VB.TextBox txtPLU 
                  BeginProperty Font 
                     Name            =   ".VnArial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   5
                  Left            =   30
                  MaxLength       =   10
                  TabIndex        =   42
                  Tag             =   "7"
                  Top             =   1380
                  Width           =   1620
               End
               Begin VB.TextBox txtPLU 
                  BeginProperty Font 
                     Name            =   ".VnArial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   4
                  Left            =   30
                  MaxLength       =   10
                  TabIndex        =   41
                  Tag             =   "6"
                  Top             =   600
                  Width           =   1620
               End
               Begin VB.TextBox txtPLU 
                  BeginProperty Font 
                     Name            =   ".VnArial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   6
                  Left            =   30
                  MaxLength       =   10
                  TabIndex        =   40
                  Tag             =   "8"
                  Top             =   2160
                  Width           =   1620
               End
               Begin VB.Label lblPrice 
                  Caption         =   "HH Price &2 :"
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
                  Index           =   6
                  Left            =   120
                  TabIndex        =   45
                  Tag             =   "L17"
                  Top             =   1140
                  Width           =   1515
               End
               Begin VB.Label lblPrice 
                  Caption         =   "HH Price &1:"
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
                  Index           =   5
                  Left            =   90
                  TabIndex        =   44
                  Tag             =   "L16"
                  Top             =   330
                  Width           =   1515
               End
               Begin VB.Label lblPrice 
                  Caption         =   "HH Price &3 :"
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
                  Index           =   4
                  Left            =   120
                  TabIndex        =   43
                  Tag             =   "L18"
                  Top             =   1920
                  Width           =   1515
               End
            End
            Begin VB.Frame fraStandar 
               Caption         =   "Standar Price"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   2955
               Left            =   60
               TabIndex        =   32
               Tag             =   "L26"
               Top             =   2400
               Width           =   1875
               Begin VB.TextBox txtPLU 
                  BeginProperty Font 
                     Name            =   ".VnArial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   3
                  Left            =   30
                  MaxLength       =   10
                  TabIndex        =   37
                  Tag             =   "5"
                  Top             =   2160
                  Width           =   1620
               End
               Begin VB.TextBox txtPLU 
                  BeginProperty Font 
                     Name            =   ".VnArial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   1
                  Left            =   30
                  MaxLength       =   10
                  TabIndex        =   34
                  Tag             =   "3"
                  Top             =   600
                  Width           =   1620
               End
               Begin VB.TextBox txtPLU 
                  BeginProperty Font 
                     Name            =   ".VnArial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Index           =   2
                  Left            =   30
                  MaxLength       =   10
                  TabIndex        =   33
                  Tag             =   "4"
                  Top             =   1380
                  Width           =   1620
               End
               Begin VB.Label lblPrice 
                  Caption         =   "Std Price &3 :"
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
                  Index           =   3
                  Left            =   120
                  TabIndex        =   38
                  Tag             =   "L15"
                  Top             =   1920
                  Width           =   1515
               End
               Begin VB.Label lblPrice 
                  Caption         =   "Std Price &1:"
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
                  Index           =   0
                  Left            =   90
                  TabIndex        =   36
                  Tag             =   "L13"
                  Top             =   330
                  Width           =   1515
               End
               Begin VB.Label lblPrice 
                  Caption         =   "Std Price &2 :"
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
                  Index           =   1
                  Left            =   120
                  TabIndex        =   35
                  Tag             =   "L14"
                  Top             =   1140
                  Width           =   1545
               End
            End
            Begin prjTouchScreen.MyButton cmdPickup 
               Height          =   435
               Index           =   0
               Left            =   4080
               TabIndex        =   16
               Top             =   1725
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   767
               BTYPE           =   14
               TX              =   ""
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   ".VnArial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   2
               FOCUSR          =   -1  'True
               BCOL            =   16777215
               BCOLO           =   16777152
               FCOL            =   16711680
               FCOLO           =   255
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmPLU.frx":1DF7F
               PICN            =   "frmPLU.frx":1DF9B
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               Value           =   0   'False
            End
            Begin VB.TextBox txtPLU 
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   10
               Left            =   930
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   14
               Tag             =   "12"
               Top             =   6105
               Width           =   1485
            End
            Begin VB.ComboBox cboPLU 
               BeginProperty Font 
                  Name            =   ".VnArial NarrowH"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   0
               Left            =   1680
               Style           =   2  'Dropdown List
               TabIndex        =   6
               Tag             =   "2"
               Top             =   1725
               Width           =   2415
            End
            Begin VB.TextBox txtPLU 
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   390
               Index           =   0
               Left            =   1680
               TabIndex        =   5
               Tag             =   "1"
               Top             =   225
               Width           =   3555
            End
            Begin VB.TextBox txtPLU 
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   11
               Left            =   3030
               TabIndex        =   4
               Tag             =   "13"
               Top             =   5490
               Width           =   945
            End
            Begin VB.TextBox txtPLU 
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   12
               Left            =   840
               TabIndex        =   3
               Tag             =   "14"
               Top             =   5460
               Width           =   1605
            End
            Begin VB.Label lblLinkCode 
               Alignment       =   1  'Right Justify
               Caption         =   "&Link PLU Modifier"
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
               TabIndex        =   77
               Tag             =   "L25"
               Top             =   960
               Width           =   1335
            End
            Begin VB.Label lblPrice 
               Caption         =   "Mµu s¾c:"
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
               Index           =   2
               Left            =   90
               TabIndex        =   15
               Tag             =   "L22"
               Top             =   6150
               Width           =   1170
            End
            Begin VB.Label lblSetup 
               Alignment       =   1  'Right Justify
               Caption         =   "Item &Name:"
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
               Index           =   0
               Left            =   90
               TabIndex        =   10
               Tag             =   "L11"
               Top             =   390
               Width           =   1560
            End
            Begin VB.Label lblSetup 
               Alignment       =   1  'Right Justify
               Caption         =   "Link to Group-&A:"
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
               Index           =   1
               Left            =   90
               TabIndex        =   9
               Tag             =   "L12"
               Top             =   1815
               Width           =   1560
            End
            Begin VB.Label lblSetup 
               Alignment       =   1  'Right Justify
               Caption         =   "&Unit:"
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
               Index           =   5
               Left            =   2490
               TabIndex        =   8
               Tag             =   "L23"
               Top             =   5550
               Width           =   525
            End
            Begin VB.Label lblSetup 
               Alignment       =   1  'Right Justify
               Caption         =   "Gi¸ vèn:"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Index           =   4
               Left            =   90
               TabIndex        =   7
               Tag             =   "L24"
               Top             =   5550
               Width           =   810
            End
         End
      End
      Begin prjTouchScreen.MyButton cmdAdd 
         Height          =   1065
         Left            =   8670
         TabIndex        =   24
         Tag             =   "L2"
         Top             =   8700
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1879
         BTYPE           =   5
         TX              =   "&Thªm míi"
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
         BCOLO           =   12640511
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPLU.frx":1FCA5
         PICN            =   "frmPLU.frx":1FCC1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdDelete 
         Height          =   1065
         Left            =   11850
         TabIndex        =   25
         Tag             =   "L4"
         Top             =   8700
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1879
         BTYPE           =   5
         TX              =   "&Xãa"
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
         BCOLO           =   12640511
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPLU.frx":20113
         PICN            =   "frmPLU.frx":2012F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdSearch 
         Height          =   1065
         Left            =   11850
         TabIndex        =   26
         Tag             =   "L5"
         Top             =   9780
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1879
         BTYPE           =   5
         TX              =   "T&×m kiÕm"
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
         BCOLO           =   12640511
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPLU.frx":20769
         PICN            =   "frmPLU.frx":20785
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   1065
         Left            =   13440
         TabIndex        =   27
         Tag             =   "L6"
         Top             =   9780
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1879
         BTYPE           =   5
         TX              =   "&§ãng"
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
         BCOLO           =   12640511
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPLU.frx":20DBF
         PICN            =   "frmPLU.frx":20DDB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdPrint 
         Height          =   1065
         Left            =   10260
         TabIndex        =   78
         Top             =   9780
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1879
         BTYPE           =   5
         TX              =   "In Danh s¸ch m· hµng"
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
         BCOLO           =   12640511
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPLU.frx":27075
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdImports 
         Cancel          =   -1  'True
         Height          =   1065
         Left            =   8670
         TabIndex        =   133
         Top             =   9780
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1879
         BTYPE           =   5
         TX              =   "NhËp danh môc tõ file Excel"
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
         BCOLO           =   12640511
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPLU.frx":27091
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
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim rsVirtual_PLU As New ADODB.Recordset
    Dim i, j, k As Integer
    Dim arrFieldNames() As String
    Dim arrRAM() As String
    Dim array_NewPLUs As String
    Dim fLoad As Boolean
    Dim fUpdate As Boolean
    Dim flagSF4 As Boolean
    Dim fActivate As Boolean
    Dim fUpdateRam As Boolean
    Dim fAddNew As Boolean
    Dim fFlexClick As Boolean
    Dim rsPLU As New ADODB.Recordset
    Dim iLimit As Double
    Dim iListIndex As Byte
    

Public Property Let Get_NewPLUs(ByVal vNewValue As String)
On Error GoTo errHdl

    array_NewPLUs = vNewValue

Exit Property
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - Get_NewPLUs "
End Property

Private Sub cmd_Click()
Dim fso As New FileSystemObject
Dim P As String
    With comdlg
         .FileName = ""
        .Filter = "Image(*.jpg)|*.jpg|*.bmp|*.*"
        .DefaultExt = "*.jpg"
        .InitDir = App.Path
        .ShowOpen
        If .FileName <> "" Then
            PImage = .FileName
            Image1.Picture = LoadPicture(PImage)
            If rsPLU.State <> 0 Then
                With rsPLU
                    .Find "ItemNum='" & flexPLU.TextMatrix(flexPLU.Row, 0) & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        .Fields("Picture") = PImage
                        .Update
                    End If
                End With
            End If
        End If
    End With
End Sub

Private Sub cmdCapture_Click()
    'MainForm.Show vbModal
End Sub

Private Sub cmdColor_Click()
    With Frame5
        .Visible = True
        .top = 4560
        .Left = 2040
    End With
    
End Sub

Private Sub cmdDept_Click()
    frmDepartement.Show vbModal
    SetCombo "Departments", cboPLU(0), "Description", False
End Sub

Private Sub cmdImports_Click()
On Error GoTo errHdl
Unload Me
    With frmUpload
        .FormCall = "Items"
        .Show vbModal
    End With
    frmItems.Show vbModal
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - cmdImports_Click "
End Sub

Private Sub cmdModifierPickup_Click()
On Error GoTo errHdl

    Dim fPickup As Byte
    fPickup = 10
    With frmPickup
        .GetCurrentValue = txtPLU(13).Text
        .GetfPickup = fPickup
        Set .FormCall = Me
        .Show vbModal
    End With
    UpdateData
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - cmdPickup_Click "
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Handle
    Dim SQL As String
    Dim cmd As New ADODB.Command
    Dim iReport As CRAXDDRT.Report
    If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    SQL = "SELECT Inventory.ItemNum, Inventory.ItemName, Departments.Dept_ID, Departments.Description, Inventory.Std_Price1, Inventory.Std_Price2, Inventory.Minstock, Inventory.Unit" & _
         " FROM Inventory INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID"
    
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set crPLUList = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crPLUList
        .Database.AddADOCommand cnData, cmd
        .txtGroupA.SetUnboundFieldSource "{ado.Description}"
        .txtPluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .txtPluName.SetUnboundFieldSource "{ado.ItemName}"
        .txtUnit.SetUnboundFieldSource "{ado.Unit}"
        .txtPrice1.SetUnboundFieldSource "{ado.Std_Price1}"
        .txtPrice2.SetUnboundFieldSource "{ado.Std_Price2}"
        .txtCostPrice.SetUnboundFieldSource "{ado.Minstock}"
        With .txtPrice1
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtPrice2
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtCostPrice
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    Set iReport = crPLUList
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With

Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & " - cmdPrint_Click"
End Sub

Private Sub cmdRemovePic_Click()
    If rsPLU.State <> 0 Then
        With rsPLU
            .Find "ItemNum='" & flexPLU.TextMatrix(flexPLU.Row, 0) & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Fields("Picture") = "00"
                .Update
            End If
        End With
    End If
    Image1.Picture = Nothing
End Sub

Private Sub cmdSend_Click()
On Error GoTo errHdl
    If fUpdate Then
        fUpdate = False
        Add_DataUpdate_To_DB
    End If
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - cmdSend_Click "
End Sub

Private Sub cmdSetcolor_Click()
    cnData.Execute "update Inventory set LimitPrice='" & txtPLU(10).Text & "'"
End Sub

Private Sub flexPLU_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl
    If KeyAscii = 13 Then
        With txtPLU(0)
            .SetFocus
            .SelStart = 0
            .SelLength = 9999
        End With
    End If
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - flexPLU_KeyPress "
End Sub

Private Sub flexPLU_Scroll()
'    flexPLU.Row = flexPLU.Row + 1
End Sub

'            -----------FORM------------
Private Sub Form_Activate()
On Error GoTo errHdl
    Dim DescArr() As String
    Dim ctrl As Control
    Dim iCount As Byte, iTab As Byte
    Dim i As Integer
    If rsVirtual_PLU.State = 0 Then
        cmdClose_Click
        fraForm.Enabled = True
        Exit Sub
    End If
    If fActivate Then
        fraForm.Enabled = True
        Exit Sub
    End If
    fActivate = True
    DescArr = LoadLanguage(LngFile, "#01:017:")
    If cmdSend.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
        Me.Caption = DescArr(1)
    iCount = 0: iTab = 0
    For i = 1 To UBound(DescArr)
    DoEvents
        Select Case i
            Case 7, 8, 9
                tabPLU.TabCaption(iTab) = DescArr(i)
                iTab = iTab + 1
            Case 10 To 25, 71 To 75
                flexPLU.TextMatrix(0, iCount) = DescArr(i)
                iCount = iCount + 1
            Case 23
                With txtPLU(12)
                    If .Tag = "" Then
                       .Text = "0"
                       .Locked = True
                    Else
                        flexPLU.TextMatrix(0, iCount) = DescArr(i)
                        iCount = iCount + 1
                    End If
                End With
        End Select
    Next i
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    fraForm.Enabled = True
    If UCase(UserID) = "131112" Or UserID = "881507" Then cmdSetcolor.Visible = True
Exit Sub
errHdl:
    fraForm.Enabled = True
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - Form_Activate "
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl
    With flexPLU
        If Shift = 2 Then 'xac dinh cac fim duoc click: shift,ctrl,alt
            If KeyCode = vbKeyDown Then ' chon keypreview trong from =true
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    If .Row > 19 Then .TopRow = .Row - 19
                End If
                KeyCode = 0
                flexPLU_Click
            ElseIf KeyCode = vbKeyUp Then 'ctrl + keyup
                If .Row > 1 Then
                    .Row = .Row - 1
                    If .Row < .TopRow Then .TopRow = .Row
                End If
                KeyCode = 0
                flexPLU_Click
            End If
        End If
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - Form_KeyDown "
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        If KeyCode = vbKeyS Then cmdSend_Click: fUpdate = False
        If KeyCode = vbKeyN Then cmdAdd_Click
        If KeyCode = vbKeyF Then cmdSearch_Click
        If KeyCode = vbKeyF4 Then cmdClose_Click
    End If
End Sub

'
Private Sub Form_Load()
On Error GoTo errHdl
    Dim i, j As Integer
    Dim sFieldName As String
    Dim strSql As String
    Dim intCount As Integer
    fraForm.BorderStyle = 0
    fraForm.Enabled = False
    
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsPLU = Open_Table(cnData, "Inventory")
    With rsPLU
    .Sort = "ItemNum ASC"
        For i = 0 To .Fields.count - 1
        DoEvents
            rsVirtual_PLU.Fields.Append .Fields(i).name, .Fields(i).Type, .Fields(i).DefinedSize
            rsVirtual_PLU.Fields(.Fields(i).name).Attributes = adColNullable
        Next i
        rsVirtual_PLU.Fields.Append "fStatus", adVarWChar, 20
        rsVirtual_PLU.Open
        If rsPLU.RecordCount > 0 Then
            .MoveFirst
        Else
            Exit Sub
        End If
        Do While Not .EOF
        DoEvents
            rsVirtual_PLU.addNew
            For j = 0 To .Fields.count - 1
            DoEvents
                If .Fields(j).Value & "" = "" Then
                    rsVirtual_PLU.Fields(j).Value = ""
                Else
                    rsVirtual_PLU.Fields(j).Value = .Fields(j).Value
                End If
            Next j
            rsVirtual_PLU.Fields("fStatus") = "default"
            .Update
            .MoveNext
        Loop
    End With
    flagSF4 = False: fActivate = False
    With rsPLU
        For i = 0 To txtPLU.count - 1 Step 1
        DoEvents
            Select Case i
                Case 0: sFieldName = "ItemName": txtPLU(i).Alignment = 0
                Case 1: sFieldName = "Std_Price1": txtPLU(i).Alignment = 1
                Case 2: sFieldName = "Std_Price2": txtPLU(i).Alignment = 1
                Case 3: sFieldName = "Std_Price3": txtPLU(i).Alignment = 1
                Case 4: sFieldName = "HH_Price1": txtPLU(i).Alignment = 1
                Case 5: sFieldName = "HH_Price2": txtPLU(i).Alignment = 1
                Case 6: sFieldName = "HH_Price3": txtPLU(i).Alignment = 1
                Case 7: sFieldName = "EV_Price1": txtPLU(i).Alignment = 1
                Case 8: sFieldName = "EV_Price2": txtPLU(i).Alignment = 1
                Case 9: sFieldName = "EV_Price3": txtPLU(i).Alignment = 1
                Case 10: sFieldName = "LimitPrice": txtPLU(i).Alignment = 0
                Case 11: sFieldName = "Unit": txtPLU(i).Alignment = 0
                Case 12: sFieldName = "Minstock": txtPLU(i).Alignment = 1
                Case 13: sFieldName = "LinkPLU": txtPLU(i).Alignment = 0
                
            End Select
        Next i
    End With
    'Set rsPLU = Nothing
    InitTagForCtrl
    Initialize
    
Exit Sub
errHdl:
    fraForm.Enabled = True
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - Form_Load "
End Sub

Private Sub Initialize()
On Error GoTo errHdl

    fLoad = False: fUpdate = False: fUpdateRam = False
    Call SetDataInFlex(Sort_By)
    
    Init_Tab_Information '19-01

'    'khoi tao cac combo
    SetCombo "Departments", cboPLU(0), "Description", False
    With flexPLU
'        SetColorFlexGrid flexPLU, 1, 1, .Cols
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
    End With
    
    Call Add_Flag_Items
    SetTextNull
    optFlag(1).Value = True
    LockTxtFlag
    flexPLU_Click
    fLoad = True

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - Initialize "
End Sub

''            ----------COMBOBOX-------------
Private Sub cboPlu_Click(Index As Integer)
On Error GoTo errHdl

    If fLoad Then Call UpdateData

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - cboPlu_Click "
End Sub

Private Sub cboPLU_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 13 Then
        If Index < cboPLU.count - 1 Then
            cboPLU(Index + 1).SetFocus
        Else
            With txtPLU(1)
                .SetFocus
                .SelStart = 0
                .SelLength = 9999
            End With
        End If
    End If

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - cboPLU_KeyPress "
End Sub

'            ----------COMMANDBUTTON-------------
Private Sub cmdAdd_Click()
On Error GoTo errHdl

    Dim iMaxPLU As Double
    Dim arrAddNew() As String
    If fUpdate Then Call Add_DataUpdate_To_DB
    fAddNew = True
    iMaxPLU = flexPLU.Rows - 1
    Dim arrPLUCodes As String
    arrPLUCodes = ";"
    For i = 1 To flexPLU.Rows - 1
    DoEvents
        arrPLUCodes = arrPLUCodes & flexPLU.TextMatrix(i, 0) & ";"
    Next i
    
    Dim CurRow As Integer
    CurRow = flexPLU.Rows - 1
    
1:  With frmAddNewPLU
        Set .FormCall = Me
        .Get_CurPLUs = arrPLUCodes
        .Show vbModal
    End With
    
    Dim tempPrice As Long
    Dim strNewPLUs As String
    
    Hide_Button True
    strNewPLUs = ""
    strNewPLUs = frmAddNewPLU.Get_AddNewRecords
    If strNewPLUs <> "" Then
        fUpdate = True
        With rsVirtual_PLU
            For j = 1 To flexPLU.Rows - 1
            DoEvents
                If j Mod 500 = 0 Then Delay 200
                If InStr(1, strNewPLUs, flexPLU.TextMatrix(j, 0) & ";", vbBinaryCompare) <> 0 Then
                    .MoveFirst
                    .Find "ItemNum='" & flexPLU.TextMatrix(j, 0) & "'"
                    If .EOF Then
                        .addNew
                        For k = 0 To flexPLU.Cols - 2
                        DoEvents
                            Select Case arrFieldNames(k)
                                Case "Std_Price1", "Std_Price2", "Std_Price3", "HH_Price1", "HH_Price2", "HH_Price3", "EV_Price1", "EV_Price2", "EV_Price3"
                                    tempPrice = .Fields(arrFieldNames(k)).Value = FillZeroForString(CStr(tempPrice), .Fields("Std_Price1").DefinedSize)
                                Case Else
                                    If flexPLU.TextMatrix(j, k) = "" Then
                                        .Fields(arrFieldNames(k)).Value = 0
                                    Else
                                        .Fields(arrFieldNames(k)).Value = flexPLU.TextMatrix(j, k)
                                    End If
                            End Select
                        Next k
                    End If
                    !fStatus = "AddNew"
                    .Update
                End If
            Next j
        End With
'        SetColorFlexGrid flexPLU, CurRow + 1, 1, flexPLU.Cols
        Init_Tab_Information
    End If
    fAddNew = False
    flexPLU_Click
    Hide_Button False

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - cmdAdd_Click "
End Sub

Private Sub cmdClose_Click()
On Error GoTo errHdl

    Dim res As Byte
        
    If Not fUpdate Then GoTo 1
    res = MsgBox("B¹n cã muèn l­u th«ng tin thay ®æi ?", vbYesNo)
    Select Case res
        Case vbYes
                Hide_Button True
                Add_DataUpdate_To_DB
                Hide_Button False
        Case vbNo:  GoTo 1
        Case vbCancel: Exit Sub
    End Select
1:
    CloseRecordset rsVirtual_PLU
    Unload Me

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - cmdClose_Click "
End Sub

Private Sub cmddelete_Click()
On Error GoTo errHdl

    Dim arrDelete() As String

    fAddNew = True
    With frmDeletePLU
        Set .FormCall = Me
        .Show vbModal, Me
    End With
    ReDim Preserve arrDelete(0)
    arrDelete = frmDeletePLU.Get_DeleteRecords
    If UBound(arrDelete) > 0 Then
        Hide_Button True
        fUpdate = True
        With rsVirtual_PLU
            .MoveFirst
            For i = 1 To UBound(arrDelete)
            DoEvents
                .Filter = "ItemNum='" & arrDelete(i) & "'"
                If Not .EOF Then
                    !fStatus = "Delete"
                Else
                    .addNew
                    !ItemNum = arrDelete(i)
                    !fStatus = "Delete"
                End If
                .Update
                .Filter = adFilterNone
            Next i
        End With
        Init_Tab_Information
        Hide_Button False
    End If
    fAddNew = False

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - cmdDelete_Click "
End Sub
'
Private Sub cmdPickup_Click(Index As Integer)
On Error GoTo errHdl

    Dim fPickup As Byte
    
    Select Case Index
        Case 0: fPickup = 1
        Case 1: fPickup = 2
        Case 2: fPickup = 3
    End Select
    With frmPickup
        .GetCurrentValue = cboPLU(Index).Text
        .GetfPickup = fPickup
        Set .FormCall = Me
        .Show vbModal
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - cmdPickup_Click "
End Sub

Private Sub cmdSearch_Click()
On Error GoTo errHdl

    fAddNew = True
    With frmFind
        .GetfSearch = 2
        Set .FormCall = Me
        .Show vbModal
    End With
    fAddNew = False

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - cmdSearch_Click "
End Sub
''            ----------FLEXGRID-------------
Private Sub flexPLU_Click()
On Error GoTo errHdl

    Dim sTemp() As String
    Dim ctrl As Control
    
    If fAddNew Then Exit Sub
    If fFlexClick Then Exit Sub
    fFlexClick = True
    fLoad = False
    With flexPLU
        If .Row = 0 Then .Row = 1
        ReDim Preserve sTemp(.Cols - 1)
        For i = 1 To .Cols - 1
        DoEvents
            sTemp(i) = .TextMatrix(.Row, i)
        Next i
    End With
    For Each ctrl In Me
    DoEvents
        With ctrl
            If .Tag <> "" Then
                If TypeOf ctrl Is TextBox Then
                    If .Tag = 1 Then
                        .Text = sTemp(.Tag)
                    Else
                        Select Case .Tag
                            Case 3 To 11
                                .Text = Format(sTemp(.Tag), formatNum)
                            Case Else
                                .Text = sTemp(.Tag)
                        End Select
                    End If
                ElseIf TypeOf ctrl Is ComboBox Then
                    If .ListCount <> 0 Then
                        If sTemp(.Tag) = "" Then
                            .ListIndex = 0
                        Else
                            If .Index = 0 Then 'GroupA
                                  '.Text = sTemp(.Tag)-1
                                  .ListIndex = sTemp(.Tag) - 1
                            Else: .ListIndex = sTemp(.Tag)
                            End If
                        End If
                    End If
                End If
            End If
        End With
    Next ctrl
    If flexPLU.TextMatrix(1, 0) = "" Then Exit Sub
'    gan gtri da check cho cac lstFlag
    With rsPLU
        .Find "ItemNum='" & flexPLU.TextMatrix(flexPLU.Row, 0) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If .Fields("Picture") = "00" Then
                Image1.Picture = Nothing
            Else
                If Dir(.Fields("Picture"), vbDirectory) <> "" Then
                    Image1.Picture = LoadPicture(.Fields("Picture"))
                End If
            End If
        End If
    End With
    For i = 0 To txtFlag.count - 1 Step 1
    DoEvents
        AddValueForList txtFlag(i).Text, lstFlag(i)
    Next i
    lblCode.Caption = flexPLU.TextMatrix(flexPLU.Row, 0)
    lblName.Caption = sTemp(1)
    fLoad = True
    fFlexClick = False
    
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - flexPLU_Click "
End Sub

Private Sub flexPLU_EnterCell()
On Error GoTo errHdl

    If fLoad Then flexPLU_Click

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - flexPLU_EnterCell "
End Sub

Private Sub SetDataInFlex(str As String)
On Error GoTo errHdl

    Dim irow As Integer
    Dim i As Integer

    InitFlexGrid
    With flexPLU
        ReDim Preserve arrFieldNames(.Cols - 1)
        For i = 0 To .Cols - 1
        DoEvents
            Select Case i
                Case 0: arrFieldNames(i) = "ItemNum"
                Case 1: arrFieldNames(i) = "ItemName"
                Case 2: arrFieldNames(i) = "Dept_ID"
                Case 3: arrFieldNames(i) = "Std_Price1"
                Case 4: arrFieldNames(i) = "Std_Price2"
                Case 5: arrFieldNames(i) = "Std_Price3"
                Case 6: arrFieldNames(i) = "HH_Price1"
                Case 7: arrFieldNames(i) = "HH_Price2"
                Case 8: arrFieldNames(i) = "HH_Price3"
                Case 9: arrFieldNames(i) = "EV_Price1"
                Case 10: arrFieldNames(i) = "EV_Price2"
                Case 11: arrFieldNames(i) = "EV_Price3"
                Case 12: arrFieldNames(i) = "LimitPrice"
                Case 13: arrFieldNames(i) = "Unit"
                Case 14: arrFieldNames(i) = "Minstock"
                Case 15: arrFieldNames(i) = "Modify_Number"
                Case 16 To 20: arrFieldNames(i) = "F" & (i - 15)
                Case Else: arrFieldNames(i) = ""
            End Select
        Next i
    End With
    irow = 1
    With rsVirtual_PLU
        .Sort = str
        If .RecordCount > 0 Then
            '15-02-2006: Dongle
            Dim iTempMaxPLU As Long
            iTempMaxPLU = .RecordCount
            flexPLU.Rows = iTempMaxPLU + 1
            .MoveFirst
            Do While Not .EOF
            DoEvents
                If irow > iTempMaxPLU Then Exit Do
                DataInFlex irow
                irow = irow + 1
                .MoveNext
            Loop
        End If
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - SetDataInFlex "
End Sub

Private Sub Data(str As String)
On Error GoTo errHdl

    Dim irow As Integer
    Dim i As Integer
    InitFlexGrid
    With flexPLU
        ReDim Preserve arrFieldNames(.Cols - 1)
        For i = 0 To .Cols - 1
        DoEvents
            Select Case i
                Case 0: arrFieldNames(i) = "ItemNum"
                Case 1: arrFieldNames(i) = "ItemName"
                Case 2: arrFieldNames(i) = "Dept_ID"
                Case 3: arrFieldNames(i) = "Std_Price1"
                Case 4: arrFieldNames(i) = "Std_Price2"
                Case 5: arrFieldNames(i) = "Std_Price3"
                Case 6: arrFieldNames(i) = "HH_Price1"
                Case 7: arrFieldNames(i) = "HH_Price2"
                Case 8: arrFieldNames(i) = "HH_Price3"
                Case 9: arrFieldNames(i) = "EV_Price1"
                Case 10: arrFieldNames(i) = "EV_Price2"
                Case 11: arrFieldNames(i) = "EV_Price3"
                Case 12: arrFieldNames(i) = "LimitPrice"
                Case 13: arrFieldNames(i) = "Unit"
                Case 14: arrFieldNames(i) = "Minstock"
                Case 15: arrFieldNames(i) = "Modify_Number"
                Case 16 To 20: arrFieldNames(i) = "F" & (i - 15)
                Case Else: arrFieldNames(i) = ""
            End Select
        Next i
    End With
   irow = 1
    With rsVirtual_PLU
        .Sort = str
        If .RecordCount > 0 Then
            Dim iTempMaxPLU As Long
            iTempMaxPLU = .RecordCount
            flexPLU.Rows = iTempMaxPLU + 1
            If Not .BOF Then
            DoEvents
            .MoveFirst
            .Move (iTempMaxPLU - 1)
            Do While Not .EOF
            DoEvents
                If irow > iTempMaxPLU Then Exit Do
                DataInFlex irow
                irow = irow + 1
                .MovePrevious
            Loop
            End If
        End If
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - Data "
End Sub

Private Sub InitFlexGrid()
On Error GoTo errHdl

    With flexPLU
        .AllowUserResizing = flexResizeBoth
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .Cols = rsVirtual_PLU.Fields.count - 4
        For i = 0 To .Cols - 1 Step 1
        DoEvents
            Select Case i
                Case 0:    .ColWidth(i) = 2100: .ColAlignment(i) = 4
                Case 1:    .ColWidth(i) = 3500: .ColAlignment(i) = 1
                Case 2:    .ColWidth(i) = 1000: .ColAlignment(i) = 4
                Case 3, 4, 5, 6, 7, 8, 9, 10, 11: .ColWidth(i) = 1400
                Case 12: .ColWidth(i) = 1200: .ColAlignment(i) = 4
                Case 13:    .ColWidth(i) = 1000
                Case 14: .ColWidth(i) = 1200: .ColAlignment(i) = 4
                Case 15: .ColWidth(i) = 2000: .ColAlignment(i) = 4
                Case 16 To 20: .ColWidth(i) = 800: .ColAlignment(i) = 4
                Case Else
                    .ColWidth(i) = 0
            End Select
        Next i
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - InitFlexGrid "
End Sub

Private Sub DataInFlex(ByVal irow As Integer)
On Error GoTo errHdl

    Dim sTemp As String
    
    With flexPLU
        For i = 0 To .Cols - 1
        DoEvents
            If arrFieldNames(i) = "" Then Exit For '22-02
            If Len(arrFieldNames(i)) > 5 Then
                sTemp = rsVirtual_PLU.Fields(arrFieldNames(i))
            Else
                sTemp = rsVirtual_PLU.Fields(arrFieldNames(i))
            End If
            If arrFieldNames(i) = "Std_Price1" Or arrFieldNames(i) = "Std_Price2" Then
                .TextMatrix(irow, i) = Format(sTemp, "#,##0")
            Else
                .TextMatrix(irow, i) = sTemp
            End If
        Next i
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - DataInFlex "
End Sub

Private Sub Form_Resize()
On Error GoTo errHdl

    With flexPLU
        .top = 500
        .Height = Me.ScaleHeight - 1000
        .Left = 0
        .Width = Me.ScaleWidth - picLabel.Width - 200
        
    End With
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - Form_Resize "
End Sub

Private Sub Form_Unload(Cancel As Integer)
    j = 0: j = 0: k = 0
    Set rsPLU = Nothing
    Set rsVirtual_PLU = Nothing
    Set cnData = Nothing
End Sub



Private Sub Image1_Click()
    Call cmd_Click
End Sub


'                -----------LISTFLAG-------------
Private Sub lstFlag_Click(Index As Integer)
On Error GoTo errHdl

    Dim strflag As String

    If fLoad Then ' event is called directly by clicking on list, not call by another functions or subs
        strflag = ""
        For i = 0 To lstFlag(Index).ListCount - 1
        DoEvents
            If lstFlag(Index).Selected(i) Then
                strflag = strflag & "1"
            Else: strflag = strflag & "0"
            End If
        Next i
        txtFlag(Index).Text = FillZeroForString(BinToHex(strflag), 2)
        Call UpdateData
    End If

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - lstFlag_Click "
End Sub

Private Sub cmdUpdate_Click()
    Unload Me
    frmUpdatePrice.Show vbModal
End Sub

Private Sub MyButton1_Click()
    Dim cnt As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object

    
    Dim recArray As Variant
    
    Dim strDB As String
    Dim fldCount As Integer
    Dim recCount As Long
    Dim iCol As Integer
    Dim irow As Integer
    
   strDB = "SELECT * from Inventory"
         Set rst = OpenCriticalTable(strDB, cnData)
    ' Create an instance of Excel and add a workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets("Sheet1")
  
    ' Display Excel and give user control of Excel's lifetime
   ' xlApp.Visible = False
    xlApp.UserControl = True
    
    ' Copy field names to the first row of the worksheet
    fldCount = rst.Fields.count
    For iCol = 1 To fldCount
        xlWs.Cells(1, iCol).Value = rst.Fields(iCol - 1).name
    Next
        
    ' Check version of Excel
    If Val(Mid(xlApp.Version, 1, InStr(1, xlApp.Version, ".") - 1)) > 8 Then
        'EXCEL 2000,2002,2003, or 2007: Use CopyFromRecordset
         
        ' Copy the recordset to the worksheet, starting in cell A2
        xlWs.Cells(2, 1).CopyFromRecordset rst
        'Note: CopyFromRecordset will fail if the recordset
        'contains an OLE object field or array data such
        'as hierarchical recordsets
        
    Else
        'EXCEL 97 or earlier: Use GetRows then copy array to Excel
    
        ' Copy recordset to an array
        recArray = rst.GetRows
        'Note: GetRows returns a 0-based array where the first
        'dimension contains fields and the second dimension
        'contains records. We will transpose this array so that
        'the first dimension contains records, allowing the
        'data to appears properly when copied to Excel
        
        ' Determine number of records

        recCount = UBound(recArray, 2) + 1 '+ 1 since 0-based array
        

        ' Check the array for contents that are not valid when
        ' copying the array to an Excel worksheet
        For iCol = 0 To fldCount - 1
            For irow = 0 To recCount - 1
                ' Take care of Date fields
                If IsDate(recArray(iCol, irow)) Then
                    recArray(iCol, irow) = Format(recArray(iCol, irow))
                ' Take care of OLE object fields or array fields
                ElseIf IsArray(recArray(iCol, irow)) Then
                    recArray(iCol, irow) = "Array Field"
                End If
            Next irow 'next record
        Next iCol 'next field
            
        ' Transpose and Copy the array to the worksheet,
        ' starting in cell A2
        xlWs.Cells(2, 1).Resize(recCount, fldCount).Value = _
            TransposeDim(recArray)
    End If

    ' Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    xlWb.SaveAs App.Path & "\menu.xls"

    ' Close ADO objects
    rst.Close
    'cnt.Close
    Set rst = Nothing
    'Set cnt = Nothing
    
    ' Release Excel references
    Set xlWs = Nothing
    Set xlWb = Nothing

    Set xlApp = Nothing
MsgBox "Xuat file menu hoan tat"
End Sub


Function TransposeDim(v As Variant) As Variant
' Custom Function to Transpose a 0-based array (v)
    
    Dim X As Long, Y As Long, Xupper As Long, Yupper As Long
    Dim tempArray As Variant
    
    Xupper = UBound(v, 2)
    Yupper = UBound(v, 1)
    
    ReDim tempArray(Xupper, Yupper)
    For X = 0 To Xupper
        For Y = 0 To Yupper
            tempArray(X, Y) = v(Y, X)
        Next Y
    Next X
    
    TransposeDim = tempArray


End Function


'                -----------OPTFLAG-------------
Private Sub optFlag_Click(Index As Integer)
On Error GoTo errHdl

    For i = 0 To optFlag.count - 1 Step 1
    DoEvents
        If Index = i Then
              frmList(i).Visible = True
        Else: frmList(i).Visible = False
        End If
    Next i

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - optFlag_Click "
End Sub
'                -----------TEXTBOX-------------
Private Sub SetTextNull()
On Error GoTo errHdl

    For i = 0 To txtPLU.count - 1 Step 1
    DoEvents
        txtPLU(i).Text = ""
    Next i

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - SetTextNull "
End Sub

Private Sub picBasicColor_Click(Index As Integer)
    On Error GoTo Handle
        txtPLU(10).Text = DectoHex(picBasicColor(Index).BackColor)
        Frame5.Visible = False
        UpdateData
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " picBasicColor_Click"
    
End Sub





Private Sub txtPLU_Change(Index As Integer)
    Select Case Index
        Case 1 To 9
            If Not IsNumeric(txtPLU(Index).Text) Then txtPLU(Index).Text = 0
        Case 10
            If txtPLU(Index).Text = "0" Or txtPLU(Index).Text = "00" Then txtPLU(Index).Text = "FFFF"
    End Select
End Sub

Private Sub txtPLU_DblClick(Index As Integer)
On Error GoTo Handle
    With frmKeyboard.txtInput
        
        .PasswordChar = ""
        .Text = txtPLU(Index).Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtPLU(Index).Text = .Let_Text_Input
        End With
        UpdateData
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtPLU_DblClick"
End Sub

Private Sub txtPLU_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl

    Dim iTempIndex As Integer
    If KeyAscii = 33 Then 'Or KeyAscii = 44
        KeyAscii = 0
        Exit Sub
    End If
    iTempIndex = -1
    If KeyAscii = 13 Then
        Select Case Index
            Case 0: cboPLU(0).SetFocus
            Case 1 To 10
                    iTempIndex = Index + 1
            Case 11
                    tabPLU.Tab = 1
        End Select
        If iTempIndex <> -1 Then
            With txtPLU(iTempIndex)
                .SetFocus
                .SelStart = 0
                .SelLength = 9999
            End With
        End If
        Exit Sub
    End If
    Select Case Index
        Case 1 To 9, 12
                If KeyAscii < 32 Then Exit Sub
                Select Case KeyAscii
                    Case 48 To 57, 46
                    Case Else:   KeyAscii = 0
                End Select
'        Case 12
'                Select Case KeyAscii
'                    Case 8
'                    Case 48 To 57, 46
'                    Case Else:   KeyAscii = 0
'                End Select
    End Select

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - txtPLU_KeyPress "
End Sub

Private Sub txtPLU_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    Select Case KeyCode
        Case 8 'key backspace
        Case Is < 32: Exit Sub
        Case vbKeyDown, vbKeyUp
            txtPLU(Index).SelStart = 0
            txtPLU(Index).SelLength = 9999
    End Select
    With txtPLU(Index)
        If .Text = "" Then
            Select Case Index
                Case 0: .Text = "-"
                Case 10 'Color
                    .Text = FillZeroForString("0", .MaxLength)
                Case 11
                    .Text = "Ly"
                Case 12: .Text = "0" 'MinStock
                Case 13: .Text = "-"
            End Select
        End If
    End With
    UpdateData

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - cmdXoa_Click "
End Sub

Private Sub txtPlu_LostFocus(Index As Integer)
On Error GoTo errHdl

    Dim sTemp As String
    
    With txtPLU(Index)
        Select Case Index
'            Case 1 To 9
'                    .Text = Format(CDbl("0" & .Text), "#,##0") 'Right("0000000" & .Text, .MaxLength)
            Case 1 To 9, 12
                    If .Text <> "" Then
                        If IsNumeric(.Text) Then
                            .Text = Format(CDbl(.Text), "#,##0")
                        Else: .Text = 0
                        End If
                    Else: .Text = 0
                    End If
'            Case 12
'                    If .Text = "" Then
'                        .Text = "0"
'                    ElseIf IsNumeric(.Text) Then
'                        If CDbl(.Text) < 0 Then
'                            tabPLU.Tab = 0
'                            .SetFocus
'                            .SelStart = 0
'                            .SelLength = 9999
'                        Else
'                            .Text = CDbl(.Text)
'                        End If
'                    Else: .Text = "0"
'                    End If
        End Select
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - txtPlu_LostFocus "
End Sub

Private Sub LockTxtFlag()
On Error GoTo errHdl

    For i = 0 To txtFlag.count - 1 Step 1
    DoEvents
        txtFlag(i).Locked = True
    Next i

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - LockTxtFlag "
End Sub

'           -----------UPDATE DATA-------------
Private Sub UpdateData()
On Error GoTo errHdl

    Dim strPLU() As String
    Dim i As Integer
    Dim tempPrice As Long
    
    If rsVirtual_PLU.RecordCount = 0 Then Exit Sub
    fUpdate = True
    With rsVirtual_PLU
        .MoveFirst
        .Find "ItemNum='" & flexPLU.TextMatrix(flexPLU.Row, 0) & "'"
        If Not .EOF Then
            strPLU = SetTextTemp
            If !fStatus = "default" Then !fStatus = "Update"
            For i = 1 To UBound(strPLU) - 1 Step 1
            DoEvents
                flexPLU.TextMatrix(flexPLU.Row, i) = strPLU(i)
                .Fields(arrFieldNames(i)).Value = strPLU(i)
                 
            Next i
            .Update
            flexPLU.Refresh
            lblCode.Caption = !ItemNum
            lblName.Caption = strPLU(1)
        End If
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - UpdateData "
End Sub

'
Private Function SetTextTemp()
On Error GoTo errHdl

    Dim ctrl As Control
    Dim S1() As String, S2 As String
    Dim tempKey As Integer

    ReDim Preserve S1(flexPLU.Cols - 1)
    For Each ctrl In Me
    DoEvents
        With ctrl
            If .Tag <> "" Then
                If TypeOf ctrl Is TextBox Then
                    Select Case .Tag
                        Case 3 To 11 'Price
                                S2 = CDbl("0" & .Text) 'Right("00000000" & .Text, .MaxLength)
                        Case 12
                            If .Text <> "" Then
                                S2 = .Text
                            Else
                                S2 = "&H0080FFFF&"
                            End If
                        Case 14
                                If .Text <> "" Then
                                    If IsNumeric(.Text) Then
                                        S2 = (.Text)
                                    Else: S2 = 0
                                    End If
                                Else: S2 = 0
                                End If
                        Case Else: S2 = .Text
                    End Select
                    S1(.Tag) = S2
                ElseIf TypeOf ctrl Is ComboBox Then
                    If .ListCount = 0 Then
                        S2 = "000"
                    Else
                        If .Tag = 2 Then 'GroupA
                              S2 = FillZeroForString(.ListIndex + 1, 3)
                        Else: S2 = FillZeroForString(.ListIndex, 3)
                        End If
                    End If
                    S1(.Tag) = S2
                End If
            End If
        End With
    Next ctrl
    SetTextTemp = S1

Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - SetTextTemp "
End Function

Private Sub Init_Tab_Information()
On Error GoTo errHdl

    Dim iMaxNumber As Integer
    Dim sTemp As String

    If flexPLU.TextMatrix(1, 0) = "" Then
        flexPLU.Enabled = False
        tabPLU.Enabled = False
        cmdDelete.Enabled = False
        Exit Sub
    Else
        flexPLU.Enabled = True
        tabPLU.Enabled = True
    End If

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - Init_Tab_Information "
End Sub

Private Sub InitTagForCtrl()
On Error GoTo errHdl

    For i = 0 To txtFlag.count - 1
    DoEvents
        txtFlag(i).Tag = i + 16
    Next i
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - InitTagForCtrl "
End Sub

Private Sub Add_DataUpdate_To_DB()
On Error GoTo errHdl

    Dim rsPLU As New ADODB.Recordset
    Dim sSQL As String
    Dim i As Integer
    Dim strSql As String
    Dim intCount As Integer
    
    Set rsPLU = Open_Table(cnData, "Inventory")
    If rsPLU.State = 0 Then Exit Sub
    
    With rsVirtual_PLU
        'UPDATE
        .MoveFirst
        .Filter = "fStatus='Update'"
        If Not .EOF Then .MoveFirst
        Do While Not .EOF
        DoEvents
            rsPLU.MoveFirst
            rsPLU.Find "ItemNum='" & !ItemNum & "'"
            If Not rsPLU.EOF Then
                For i = 0 To rsPLU.Fields.count - 1
                DoEvents
                
                    rsPLU.Fields(i).Value = .Fields(i).Value
                Next i
                rsPLU.Update
            End If
            .MoveNext
        Loop
        .Filter = adFilterNone
        .MoveFirst
        .Filter = "fStatus='AddNew'"
        If Not .EOF Then .MoveFirst
        Do While Not .EOF
        DoEvents
            If rsPLU.RecordCount > 0 Then
                rsPLU.MoveFirst
                rsPLU.Find "ItemNum='" & !ItemNum & "'"
                If rsPLU.EOF Then rsPLU.addNew
            Else
                rsPLU.addNew
            End If
            For i = 0 To rsPLU.Fields.count - 1
            DoEvents
                rsPLU.Fields(i).Value = .Fields(i).Value
            Next i
            rsPLU.Update
            !fStatus = "Update"
            .MoveNext
        Loop
        .Filter = adFilterNone

        'DELETE
        .MoveFirst
        .Filter = "fStatus='Delete'"
        If Not .EOF Then
            .MoveFirst
            Do While Not .EOF
            DoEvents
                rsPLU.MoveFirst
                rsPLU.Find "ItemNum='" & !ItemNum & "'"
                If Not rsPLU.EOF Then
                    sSQL = "Delete from Inventory where ItemNum='" & !ItemNum & "'"
                    cnData.Execute sSQL
                End If
                .MoveNext
            Loop
            Set rsPLU = Nothing

            Dim iInc As Integer
            Dim res As New ADODB.Recordset
    
            Set res = Open_Table(cnData, "SetMLink")
            If res.State = 0 Then GoTo 1
            If res.RecordCount = 0 Then GoTo 1
            .MoveFirst
            Do While Not .EOF
            DoEvents
                res.MoveFirst
                res.Find "PluCode='" & !ItemNum & "'"
                If Not res.EOF Then
                    cnData.Execute "Delete  from SetMLink where PLUCode='" & !PluCode & "'"
                End If
                .MoveNext
            Loop
1:          CloseRecordset res
        End If
        .Filter = adFilterNone
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - Add_DataUpdate_To_DB "
End Sub

Private Sub Hide_Button(ByVal fHide As Boolean)
On Error GoTo errHdl

    cmdSend.Enabled = Not fHide
    cmdAdd.Enabled = Not fHide
    cmdDelete.Enabled = Not fHide
    cmdSearch.Enabled = Not fHide
    cmdClose.Enabled = Not fHide

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - Hide_Button "
End Sub


Public Function fReset_MaxLength(txtText As TextBox, iIndex As Integer, ilength As Integer, svalue As String, sNumFormat As String) As String
On Error GoTo errHdl

    Dim iCount As Byte
    
    iCount = 0
    With txtText
        .MaxLength = ilength
        For i = 1 To Len(svalue)
        DoEvents
            If Mid(svalue, i, 1) = "." Or Mid(svalue, i, 1) = "," Then
                iCount = iCount + 1
            End If
        Next i
        .MaxLength = .MaxLength + iCount
        svalue = Format(fConvert_Price(svalue, formatNum), formatNum)
    End With
    fReset_MaxLength = svalue

Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - fReset_MaxLength "
End Function

Public Function fConvert_Price(svalue As String, sNumFormat As String) As String
On Error GoTo errHdl

    Dim iLen As Integer
    Dim sTemp As String
    Dim iResult As Double
    
    If svalue = "" Then GoTo 1
    If sNumFormat = "" Then GoTo 1
    sTemp = Right(sNumFormat, Len(sNumFormat) - InStr(sNumFormat, "."))
    If Len(sTemp) = Len(sNumFormat) Then
        iResult = svalue
    Else
        svalue = RemoveComma(svalue)
        iResult = CDbl(svalue) / CDbl("1" & sTemp)
    End If
1:    fConvert_Price = iResult

Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - fConvert_Price "
End Function

Public Sub Add_Flag_Items()
On Error GoTo Handle
    Dim arrFlag() As String
    Dim iCount As Integer
    arrFlag = LoadLanguage(LngFile, "#01:017:")
    iCount = lstFlag.count - 1
    For i = 0 To iCount
        lstFlag(i).FontSize = 12
        Select Case i
            Case 0
                For j = 1 To 8
                    lstFlag(i).AddItem arrFlag(j + 30)
                Next j
            Case 1
                For j = 1 To 8
                    lstFlag(i).AddItem arrFlag(j + 38)
                Next j
            Case 2
                For j = 1 To 8
                    lstFlag(i).AddItem arrFlag(j + 46)
                Next j
            Case 3
                For j = 1 To 8
                    lstFlag(i).AddItem arrFlag(j + 54)
                Next j
            Case 4
                For j = 1 To 8
                    lstFlag(i).AddItem arrFlag(j + 60)
                Next j
        End Select
    Next i
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "Add_Flag_Items"
End Sub


