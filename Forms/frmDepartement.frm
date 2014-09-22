VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmDepartement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Group A"
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12480
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog comdlgColor 
      Left            =   11400
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame frmCmd 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   6480
      TabIndex        =   17
      Top             =   6360
      Width           =   5655
      Begin prjTouchScreen.MyButton cmdAdd 
         Height          =   825
         Left            =   60
         TabIndex        =   18
         Tag             =   "L4"
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1455
         BTYPE           =   5
         TX              =   "&Thªm"
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
         MICON           =   "frmDepartement.frx":0000
         PICN            =   "frmDepartement.frx":001C
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
         Cancel          =   -1  'True
         Height          =   825
         Left            =   4200
         TabIndex        =   19
         Tag             =   "L7"
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1455
         BTYPE           =   5
         TX              =   "§ãng"
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
         MICON           =   "frmDepartement.frx":046E
         PICN            =   "frmDepartement.frx":048A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdHelp 
         Height          =   825
         Left            =   2790
         TabIndex        =   20
         Tag             =   "L6"
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1455
         BTYPE           =   5
         TX              =   "Gióp ®ì"
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
         MICON           =   "frmDepartement.frx":6724
         PICN            =   "frmDepartement.frx":6740
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdSend 
         Height          =   825
         Left            =   1410
         TabIndex        =   21
         Tag             =   "L5"
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1455
         BTYPE           =   5
         TX              =   "L­u"
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
         MICON           =   "frmDepartement.frx":6D7A
         PICN            =   "frmDepartement.frx":6D96
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
   Begin TabDlg.SSTab tabGroup 
      Height          =   4335
      Left            =   6360
      TabIndex        =   2
      Top             =   1920
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Th«ng tin Nhãm hµng"
      TabPicture(0)   =   "frmDepartement.frx":72DA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmSetup"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cê nhãm hµng"
      TabPicture(1)   =   "frmDepartement.frx":72F6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmFlag(1)"
      Tab(1).ControlCount=   1
      Begin VB.Frame frmFlag 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Index           =   1
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   5655
         Begin VB.PictureBox picList 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3420
            Index           =   1
            Left            =   60
            ScaleHeight     =   3360
            ScaleWidth      =   5460
            TabIndex        =   14
            Top             =   160
            Width           =   5520
            Begin VB.ListBox lstFlag 
               Height          =   2610
               Index           =   0
               ItemData        =   "frmDepartement.frx":7312
               Left            =   120
               List            =   "frmDepartement.frx":7314
               Style           =   1  'Checkbox
               TabIndex        =   16
               Top             =   600
               Width           =   5175
            End
            Begin VB.TextBox txtFlag 
               Alignment       =   2  'Center
               Height          =   375
               Index           =   0
               Left            =   2280
               MaxLength       =   2
               TabIndex        =   15
               Top             =   120
               Width           =   975
            End
         End
      End
      Begin VB.Frame frmSetup 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   90
         TabIndex        =   5
         Top             =   450
         Width           =   5775
         Begin prjTouchScreen.MyButton cmdMain 
            Height          =   495
            Left            =   4680
            TabIndex        =   75
            Top             =   1560
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            BTYPE           =   14
            TX              =   "..."
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
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDepartement.frx":7316
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
            Left            =   1800
            TabIndex        =   25
            Top             =   1320
            Visible         =   0   'False
            Width           =   3975
            Begin VB.PictureBox picBasicColor 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   47
               Left            =   3480
               ScaleHeight     =   225
               ScaleWidth      =   315
               TabIndex        =   73
               Tag             =   "FF"
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
               TabIndex        =   72
               Tag             =   "20"
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
               TabIndex        =   71
               Tag             =   "B6"
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
               TabIndex        =   70
               Tag             =   "2D"
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
               TabIndex        =   69
               Tag             =   "6D"
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
               TabIndex        =   68
               Tag             =   "6C"
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
               TabIndex        =   67
               Tag             =   "6C"
               Top             =   2040
               Width           =   375
            End
            Begin VB.PictureBox picBasicColor 
               BackColor       =   &H00000000&
               Height          =   285
               Index           =   40
               Left            =   120
               ScaleHeight     =   225
               ScaleWidth      =   315
               TabIndex        =   66
               Tag             =   "00"
               Top             =   2040
               Width           =   375
            End
            Begin VB.PictureBox picBasicColor 
               BackColor       =   &H00800040&
               Height          =   285
               Index           =   39
               Left            =   3480
               ScaleHeight     =   225
               ScaleWidth      =   315
               TabIndex        =   65
               Tag             =   "21"
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
               TabIndex        =   64
               Tag             =   "20"
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
               TabIndex        =   63
               Tag             =   "00"
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
               TabIndex        =   62
               Tag             =   "01"
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
               TabIndex        =   61
               Tag             =   "04"
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
               TabIndex        =   60
               Tag             =   "04"
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
               TabIndex        =   59
               Tag             =   "64"
               Top             =   1680
               Width           =   375
            End
            Begin VB.PictureBox picBasicColor 
               BackColor       =   &H00000040&
               Height          =   285
               Index           =   32
               Left            =   120
               ScaleHeight     =   225
               ScaleWidth      =   315
               TabIndex        =   58
               Tag             =   "20"
               Top             =   1680
               Width           =   375
            End
            Begin VB.PictureBox picBasicColor 
               BackColor       =   &H00FF0080&
               Height          =   285
               Index           =   31
               Left            =   3480
               ScaleHeight     =   225
               ScaleWidth      =   315
               TabIndex        =   57
               Tag             =   "63"
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
               TabIndex        =   56
               Tag             =   "61"
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
               TabIndex        =   55
               Tag             =   "01"
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
               TabIndex        =   54
               Tag             =   "03"
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
               TabIndex        =   53
               Tag             =   "0C"
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
               TabIndex        =   52
               Tag             =   "0C"
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
               TabIndex        =   51
               Tag             =   "EC"
               Top             =   1320
               Width           =   375
            End
            Begin VB.PictureBox picBasicColor 
               BackColor       =   &H00000080&
               Height          =   285
               Index           =   24
               Left            =   120
               ScaleHeight     =   225
               ScaleWidth      =   315
               TabIndex        =   50
               Tag             =   "60"
               Top             =   1320
               Width           =   375
            End
            Begin VB.PictureBox picBasicColor 
               BackColor       =   &H008000FF&
               Height          =   285
               Index           =   23
               Left            =   3480
               ScaleHeight     =   225
               ScaleWidth      =   315
               TabIndex        =   49
               Tag             =   "E1"
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
               TabIndex        =   48
               Tag             =   "60"
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
               TabIndex        =   47
               Tag             =   "6F"
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
               TabIndex        =   46
               Tag             =   "05"
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
               TabIndex        =   45
               Tag             =   "0D"
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
               TabIndex        =   44
               Tag             =   "1C"
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
               TabIndex        =   43
               Tag             =   "EC"
               Top             =   960
               Width           =   375
            End
            Begin VB.PictureBox picBasicColor 
               BackColor       =   &H00404080&
               Height          =   285
               Index           =   16
               Left            =   120
               ScaleHeight     =   225
               ScaleWidth      =   315
               TabIndex        =   42
               Tag             =   "64"
               Top             =   960
               Width           =   375
            End
            Begin VB.PictureBox picBasicColor 
               BackColor       =   &H00FF00FF&
               Height          =   285
               Index           =   15
               Left            =   3480
               ScaleHeight     =   225
               ScaleWidth      =   315
               TabIndex        =   41
               Tag             =   "E3"
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
               TabIndex        =   40
               Tag             =   "6E"
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
               TabIndex        =   39
               Tag             =   "0E"
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
               TabIndex        =   38
               Tag             =   "1F"
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
               TabIndex        =   37
               Tag             =   "1C"
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
               TabIndex        =   36
               Tag             =   "7C"
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
               TabIndex        =   35
               Tag             =   "FC"
               Top             =   600
               Width           =   375
            End
            Begin VB.PictureBox picBasicColor 
               BackColor       =   &H000000FF&
               Height          =   285
               Index           =   8
               Left            =   120
               ScaleHeight     =   225
               ScaleWidth      =   315
               TabIndex        =   34
               Tag             =   "E0"
               Top             =   600
               Width           =   375
            End
            Begin VB.PictureBox picBasicColor 
               BackColor       =   &H00FF80FF&
               Height          =   285
               Index           =   7
               Left            =   3480
               ScaleHeight     =   225
               ScaleWidth      =   315
               TabIndex        =   33
               Tag             =   "EF"
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
               TabIndex        =   32
               Tag             =   "EE"
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
               TabIndex        =   31
               Tag             =   "0F"
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
               TabIndex        =   30
               Tag             =   "7F"
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
               TabIndex        =   29
               Tag             =   "1D"
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
               TabIndex        =   28
               Tag             =   "7D"
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
               TabIndex        =   27
               Tag             =   "FD"
               Top             =   210
               Width           =   375
            End
            Begin VB.PictureBox picBasicColor 
               BackColor       =   &H008080FF&
               Height          =   285
               Index           =   0
               Left            =   120
               ScaleHeight     =   225
               ScaleWidth      =   315
               TabIndex        =   26
               Tag             =   "ED"
               Top             =   210
               Width           =   375
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
            Begin VB.Label lblMove 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label1"
               ForeColor       =   &H80000008&
               Height          =   320
               Left            =   105
               TabIndex        =   74
               Top             =   195
               Width           =   400
            End
         End
         Begin prjTouchScreen.MyButton cmdColor 
            Height          =   495
            Left            =   3480
            TabIndex        =   24
            Top             =   2640
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
            BTYPE           =   4
            TX              =   "Mµu hiÓn thÞ nhãm"
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
            BCOL            =   8421631
            BCOLO           =   8454143
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDepartement.frx":7332
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin VB.TextBox txtGroup 
            BeginProperty Font 
               Name            =   ".VnArialH"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Index           =   1
            Left            =   1440
            TabIndex        =   22
            Top             =   2640
            Width           =   1815
         End
         Begin VB.ComboBox cboMainGroup 
            BeginProperty Font 
               Name            =   ".VnArialH"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1440
            TabIndex        =   12
            Text            =   "cboMainGroup"
            Top             =   1620
            Width           =   3255
         End
         Begin VB.TextBox txtGroup 
            BeginProperty Font 
               Name            =   ".VnArialH"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Index           =   0
            Left            =   1440
            TabIndex        =   6
            Top             =   540
            Width           =   3855
         End
         Begin VB.Label lblColor 
            Caption         =   "Mµu hiÓn thÞ:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   23
            Top             =   2280
            Width           =   1725
         End
         Begin VB.Label lblGroupName 
            Caption         =   "Group &Name:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   8
            Tag             =   "L2"
            Top             =   180
            Width           =   1725
         End
         Begin VB.Label lblLink 
            Caption         =   "&Link to Main Group-A/Major Group:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   180
            TabIndex        =   7
            Tag             =   "L9"
            Top             =   1230
            Width           =   3735
         End
      End
   End
   Begin VB.PictureBox picLabel 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   6330
      ScaleHeight     =   1005
      ScaleWidth      =   5925
      TabIndex        =   1
      Top             =   180
      Width           =   5985
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Group Name"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   405
         Left            =   60
         TabIndex        =   4
         Top             =   480
         Width           =   5715
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Group No"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   405
         Left            =   60
         TabIndex        =   3
         Top             =   75
         Width           =   5715
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flexGroupA 
      Height          =   7545
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   13309
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      TextStyleFixed  =   3
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjTouchScreen.MyButton cmdSearch 
      Height          =   915
      Left            =   4320
      TabIndex        =   9
      Tag             =   "L8"
      Top             =   6600
      Visible         =   0   'False
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1614
      BTYPE           =   14
      TX              =   "&T×m kiÕm"
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
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDepartement.frx":734E
      PICN            =   "frmDepartement.frx":736A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSForms.ComboBox cboSeach 
      Height          =   765
      Left            =   210
      TabIndex        =   11
      Top             =   6720
      Visible         =   0   'False
      Width           =   3975
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "7011;1349"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   ".VnArial"
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblSeach 
      Caption         =   "&T×m kiÕm:"
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
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6420
      Visible         =   0   'False
      Width           =   1755
   End
End
Attribute VB_Name = "frmDepartement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim rsGroupA As New ADODB.Recordset
    Dim fLoad As Boolean, fUpdate As Boolean
    Dim fActivate As Boolean
    Dim fFlexClick As Boolean
    Dim arrUpdate() As Variant
    Dim i, j As Integer
    Dim DescArr() As String
    Dim addNew As Boolean
    Dim flag As Boolean



Private Sub cmdKeyboard_Click()
    frmKeyboard.Show vbModal
End Sub

Private Sub cboMainGroup_LostFocus()
    If cboMainGroup.ListIndex = 0 Then
        MsgBox "B¹n ph¶i chän nhãm chÝnh"
        cboMainGroup.SetFocus
    End If
End Sub

Private Sub cmdAdd_Click()
    If cmdAdd.Caption = DescArr(4) Then
        Call clearText
    End If
    addNew = True
    flexGroupA.Rows = flexGroupA.Rows + 1
    flexGroupA.TextMatrix(flexGroupA.Rows - 1, 0) = MaxDept_ID
    flexGroupA.TextMatrix(flexGroupA.Rows - 1, 1) = "GRPA"
    flexGroupA.TextMatrix(flexGroupA.Rows - 1, 2) = "01"
    flexGroupA.TextMatrix(flexGroupA.Rows - 1, 3) = "00"
    flexGroupA.TextMatrix(flexGroupA.Rows - 1, 4) = "3647829"
End Sub

Private Sub cmdColor_Click()
    With Frame5
        .Visible = True
        .top = 1200
        .Left = 1680
    End With
End Sub

Private Sub cmdMain_Click()
    frmMainGroup.Show vbModal
    SetCombo "MainGroup", cboMainGroup, "GroupName", flag
End Sub

Private Sub cmdSend_Click()
On Error GoTo Handle
Dim res
If addNew = True Then
    With rsGroupA
        .addNew
        .Fields("Dept_ID") = MaxDept_ID
        .Fields("Store_ID") = Store_ID
        .Fields("Description") = txtGroup(0).Text
        .Fields("MainGroup") = Right("00" & cboMainGroup.ListIndex - 1, 2)
        .Fields("F") = txtFlag(0).Text
        .Fields("ColorDept") = txtGroup(1).Text
        .Update
    End With
    'Call UpdateData
End If

If fUpdate Then
      res = MsgBox("B¹n cã muèn l­u th«ng tin thay ®æi kh«ng?", vbYesNo)
    Select Case res
        Case vbYes
            
            Add_DataUpdate_To_DB
        Case vbNo:  Exit Sub
    End Select
End If
fUpdate = False
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - Save_Change Department"
End Sub

'Private Sub flexGroupA_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim ans As Integer
'    If KeyCode = vbKeyDelete Then
'        ans = MsgBox("B¹n cã muèn xãa mÉu tin nµy kh«ng?", vbYesNo)
'        If ans = vbYes Then
'            cnData.Execute "Delete  from Departments where Dept_ID='" & flexGroupA.TextMatrix(flexGroupA.Row, 0) & "'"
'            Call UpdateData
'            Call SetDataInFlex
'        End If
'    End If
'End Sub

'           ------------ FORM ----------
Private Sub Form_Activate()
On Error GoTo errHdl

    Dim ctrl As Control
    
    If rsGroupA.State = 0 Then
        cmdClose_Click
        Exit Sub
    End If
    If fActivate Then Exit Sub
    fActivate = True
    DescArr = LoadLanguage(LngFile, "#01:012:")
    If cmdSend.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(3)
    tabGroup.TabCaption(0) = DescArr(3)
    tabGroup.TabCaption(1) = DescArr(11)
    With flexGroupA
        .TextMatrix(0, 0) = DescArr(1)
        .TextMatrix(0, 1) = DescArr(2)
        .TextMatrix(0, 2) = DescArr(10)
        .TextMatrix(0, 3) = "F"
        .TextMatrix(0, 4) = "Color"
        .ColAlignment(1) = 2
        
    End With
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    With flexGroupA
        If Shift = 2 Then
            If KeyCode = vbKeyDown Then
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    If .Row > 16 Then .TopRow = .Row - 15
                End If
                KeyCode = 0
                flexGroupA_Click
            ElseIf KeyCode = vbKeyUp Then
                If .Row > 1 Then
                    .Row = .Row - 1
                    If .Row < .TopRow Then .TopRow = .Row
                End If
                KeyCode = 0
                flexGroupA_Click
            End If
        End If
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Form_KeyDown"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl
    DescArr = LoadLanguage(LngFile, "#01:012:")
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsGroupA = OpenCriticalTable("select * from Departments order by Dept_ID ASC", cnData)
    If rsGroupA.State = 0 Then Exit Sub
    Initialize
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Form_Load"
End Sub

Private Sub Initialize()
On Error GoTo errHdl

    Dim ctrl As Control
    
    fLoad = False: fUpdate = False: fActivate = False
    ReDim Preserve arrUpdate(0)
    SetDataInFlex
    For i = 0 To txtGroup.count - 1
    DoEvents
        Select Case i
            Case 0: txtGroup(0).MaxLength = rsGroupA.Fields("Description").DefinedSize
            Case 1: txtGroup(1).Locked = True
        End Select
    Next i
     flag = True
   
    SetCombo "MainGroup", cboMainGroup, "GroupName", flag
        
    With flexGroupA
'        SetColorFlexGrid flexGroupA, 1, 1, .Cols
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
        .ScrollTrack = True
    End With
    
    Call Add_Flag_To_List
    
    flexGroupA_Click
    fLoad = True
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Initialize"
End Sub
'           ---------- COMBOBOX ---------
Private Sub cboMaingroup_Click()
On Error GoTo errHdl

    If fLoad Then UpdateData  'update dlieu tren grid & csdl
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - cboMainGroup_Click"
End Sub

Private Sub cboMainGroup_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    Dim tempIndex As Integer
    If KeyAscii = 13 Then
         tempIndex = 0
       
        With txtGroup(tempIndex)
            .SetFocus
            .SelStart = 0
            .SelLength = 9999
        End With
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - cboMainGroup_KeyPress"
End Sub
'           --------- COMMANDBUTTON --------
Private Sub cmdClose_Click()
On Error GoTo errHdl

    Dim res
    
    If Not fUpdate Then GoTo 1
      res = MsgBox("B¹n cã muèn l­u th«ng tin thay ®æi?", vbYesNoCancel)
    Select Case res
        Case vbYes
            Add_DataUpdate_To_DB
        Case vbNo:  GoTo 1
        Case vbCancel: Exit Sub
    End Select
1:
    CloseRecordset rsGroupA
    Unload Me
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - cmdClose_Click"
End Sub
'           ---------- FLEXGRID ----------
Private Sub flexGroupA_Click()
On Error GoTo errHdl

    Dim sTemp() As String
    Dim ctrl As Control
    Dim j As Integer
    fLoad = False
    If rsGroupA.RecordCount = 0 Then Exit Sub
    If fFlexClick Then Exit Sub
    fFlexClick = True
    With flexGroupA
        ReDim Preserve sTemp(.Cols)
        For i = 0 To .Cols - 1
        DoEvents
            sTemp(i) = .TextMatrix(.Row, i)
        Next i
        For Each ctrl In Me
        DoEvents
            With ctrl
                If .Tag <> "" Then
                    If TypeOf ctrl Is TextBox And .Tag <= flexGroupA.Cols - 1 Then
                        .Text = sTemp(.Tag)
                        '.BackColor = flexGroupA.TextMatrix(flexGroupA.Row, 4)
                    ElseIf TypeOf ctrl Is ComboBox Then
                        If .ListCount <> 0 Then
                            .ListIndex = sTemp(.Tag)
                        End If
                    End If
                End If
            End With
        Next ctrl
        
        For j = 0 To txtFlag.count - 1 Step 1
            DoEvents
                AddValueForList txtFlag(j).Text, lstFlag(j)
        Next j
            
        lblNo.Caption = .TextMatrix(.Row, 0)
        lblName.Caption = sTemp(1)
    End With
    fFlexClick = False
    fLoad = True
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - flexGroupA_Click"
End Sub

Private Sub flexGroupA_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 13 Then
        With txtGroup(0)
            .SetFocus
            .SelStart = 0
            .SelLength = 9999
        End With
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - flexGroupA_KeyPress"
End Sub

Private Sub flexGroupA_EnterCell()
On Error GoTo errHdl

    If fLoad Then flexGroupA_Click
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - flexGroupA_EnterCell"
End Sub

Private Sub SetDataInFlex()
On Error GoTo errHdl

    Dim irow As Integer
    Dim sTemp As String
    Set rsGroupA = Open_Table(cnData, "Departments")
    SetHeaderFlexGrid
    irow = 1
    With rsGroupA
        .Sort = "Dept_ID ASC"
        If .RecordCount > 0 Then
            flexGroupA.Rows = .RecordCount + 1
            Do While Not .EOF
            DoEvents
                For i = 0 To flexGroupA.Cols - 1
                DoEvents
                    Select Case i
                        Case 0: sTemp = "Dept_ID"
                        Case 1: sTemp = "Description": txtGroup(0).Tag = 1
                        Case 2: sTemp = "MainGroup": cboMainGroup.Tag = 2
                        Case 3: sTemp = "F": txtFlag(0).Tag = 3
                        Case 4: sTemp = "ColorDept": txtGroup(1).Tag = 4
                    End Select
                    flexGroupA.TextMatrix(irow, i) = .Fields(sTemp)
                Next i
                irow = irow + 1
                .MoveNext
            Loop
        End If
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - SetDataInFlex"
End Sub

Private Sub SetHeaderFlexGrid()
On Error GoTo errHdl

    With flexGroupA
        .Cols = rsGroupA.Fields.count - 2
        .AllowUserResizing = flexResizeBoth
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        
        For i = 0 To rsGroupA.Fields.count - 2
        DoEvents
            Select Case i
                Case 0: .ColWidth(i) = 1600: .ColAlignment(i) = 4
                Case 1: .ColWidth(i) = 3550: .ColAlignment(i) = 1
                Case 2: .ColWidth(i) = 1500: .ColAlignment(i) = 4
                Case 3: .ColWidth(i) = 1000: .ColAlignment(i) = 4
                Case 4: .ColWidth(i) = 1000: .ColAlignment(i) = 4
            End Select
        Next i
        
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - SetHeaderFlexGrid"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set cnData = Nothing
End Sub

Private Sub lstFlag_Click(Index As Integer)
On Error GoTo errHdl

    Dim strflag As String
    If fLoad Then
        strflag = ""
        For i = 0 To lstFlag(Index).ListCount - 1
        DoEvents
            If lstFlag(Index).Selected(i) Then
                  strflag = strflag & "1"
            Else: strflag = strflag & "0"
            End If
        Next i
        txtFlag(Index).Text = FillZeroForString(BinToHex(strflag), 2)
        UpdateData
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - lstFlag_Click"
End Sub

Private Sub picBasicColor_Click(Index As Integer)
    Frame5.Visible = False
    txtGroup(1).Text = picBasicColor(Index).BackColor
    UpdateData
End Sub

Private Sub txtGroup_DblClick(Index As Integer)
    On Error GoTo Handle
        With frmKeyboard
            .FormCallkeyboard = "Other"
            .txtInput.PasswordChar = ""
            .txtInput.Text = txtGroup(0).Text
            .Show vbModal
            txtGroup(0).Text = .Let_Text_Input
        End With
        
        Call UpdateData
       
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtGroup_DblClick"
End Sub

'           --------- TEXTBOX ---------
Private Sub txtGroup_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 33 Or KeyAscii = 44 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 Then
        Select Case Index
            Case 0: cboMainGroup.SetFocus
            Case 1
                    With txtGroup(0)
                        .SetFocus
                        .SelStart = 0
                        .SelLength = 9999
                    End With
        End Select
        Exit Sub
    End If
    If Index = 1 Then
        Select Case KeyAscii
            Case Is < 32: Exit Sub
            Case Else: KeyAscii = 0
        End Select
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - txtGroup_KeyPress"
End Sub

Private Sub txtGroup_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    Select Case KeyCode
        Case 8 'key backspace
        Case Is < 32, vbKeyDown, vbKeyUp: Exit Sub
    End Select
    UpdateData
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - txtGroup_KeyUp"
End Sub

Private Sub SetTextNull()
On Error GoTo errHdl

    txtGroup(0).Text = ""
    txtGroup(1).Text = ""
    lblNo.Caption = "01"
    lblName.Caption = "Group-A-01"
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - SetTextNull"
End Sub
'           ----------- UPDATE DATA ---------
Private Sub UpdateData()
On Error GoTo errHdl

    Dim sTemp() As String
    Dim i As Integer
    
    If rsGroupA.RecordCount = 0 Then Exit Sub
    fUpdate = True
    sTemp = SetTextTemp
    With flexGroupA
        For i = 1 To UBound(sTemp) - 2 Step 1
        DoEvents
            .TextMatrix(.Row, i) = sTemp(i)
        Next i
        lblNo = sTemp(0)
        lblName = sTemp(1)
    End With
    arrUpdate = Add_UpdatedData_To_Array(flexGroupA, arrUpdate)
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - UpdateData"
End Sub

Private Function SetTextTemp()
On Error GoTo errHdl

    Dim ctrl As Control
    Dim S1() As String
    
    'ReDim Preserve s1(rsGroupA.Fields.Count - 1)
    ReDim Preserve S1(rsGroupA.Fields.count - 1)
    S1(0) = flexGroupA.TextMatrix(flexGroupA.Row, 0)
    For Each ctrl In Me
    DoEvents
        With ctrl
            If .Tag <> "" Then
                If TypeOf ctrl Is TextBox And .Tag <= flexGroupA.Cols - 1 Then
                    S1(.Tag) = .Text
                ElseIf TypeOf ctrl Is ComboBox Then
                    If .ListCount <> 0 Then
                        S1(.Tag) = FillZeroForString(.ListIndex, 2)
                    End If
                End If
            End If
        End With
    Next ctrl
    SetTextTemp = S1
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - SetTextTemp"
End Function


'           --- ADD UPDATED DATA TO DATABASE ---
Private Sub Add_DataUpdate_To_DB()
On Error GoTo errHdl

    Dim sFieldName As String
    Dim i As Integer
    
    With rsGroupA
        For i = 1 To UBound(arrUpdate)
            DoEvents
            .MoveFirst
            .Find "Dept_ID='" & arrUpdate(i)(0) & "'"
            For j = 1 To .Fields.count - 3
            DoEvents
                Select Case j
                    'Case 0: sFieldName = "Dept_ID"
                    Case 1: sFieldName = "Description"
                    Case 2: sFieldName = "MainGroup"
                    Case 3: sFieldName = "F"
                    Case 4: sFieldName = "ColorDept"
                End Select
                .Fields(sFieldName) = arrUpdate(i)(j)
            Next j
            .Update
        Next i
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Add_DataUpdate_To_DB"
End Sub

Public Sub clearText()
    On Error GoTo Handle
        txtGroup(0).Text = ""
        txtGroup(0).SetFocus
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  clearText"
End Sub

Public Function MaxDept_ID() As String
On Error GoTo Handle
Dim str As String
Dim rs As New ADODB.Recordset
Dim Dept_ID As String
str = "select max(Dept_ID) as MaxDept from Departments"
Set rs = OpenCriticalTable(str, cnData)
    If rs.RecordCount > 0 Then
        Dept_ID = Right("000" & CDbl("0" & rs.Fields("MaxDept")) + 1, 3)
    Else
        Dept_ID = "001"
    End If
    MaxDept_ID = Dept_ID
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & " MaxDept_ID"
End Function

Public Sub Add_Flag_To_List()
On Error GoTo Handle
    With lstFlag(0)
        For i = 1 To 8
            .AddItem DescArr(i + 11)
        Next i
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Add_Flag_To_List"
End Sub
