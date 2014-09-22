VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmItem_Details 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6270
   ClientLeft      =   15
   ClientTop       =   -30
   ClientWidth     =   12150
   ClipControls    =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   12150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   8280
      TabIndex        =   79
      Top             =   1200
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   6376
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "M¸y in order"
      TabPicture(0)   =   "frmItem_Details.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picList(4)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "§Þnh d¹ng"
      TabPicture(1)   =   "frmItem_Details.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picList(0)"
      Tab(1).ControlCount=   1
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
         Index           =   4
         Left            =   120
         ScaleHeight     =   3075
         ScaleWidth      =   3180
         TabIndex        =   83
         Top             =   360
         Width           =   3240
         Begin VB.TextBox txtFlag 
            Alignment       =   2  'Center
            Height          =   375
            Index           =   1
            Left            =   840
            TabIndex        =   85
            Tag             =   "13"
            Text            =   "Text1"
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ListBox lstFlag 
            Height          =   3060
            Index           =   1
            Left            =   -30
            Style           =   1  'Checkbox
            TabIndex        =   84
            Top             =   0
            Width           =   3255
         End
      End
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
         Index           =   0
         Left            =   -74880
         ScaleHeight     =   3075
         ScaleWidth      =   3180
         TabIndex        =   80
         Top             =   360
         Width           =   3240
         Begin VB.TextBox txtFlag 
            Alignment       =   2  'Center
            Height          =   375
            Index           =   0
            Left            =   840
            TabIndex        =   82
            Tag             =   "13"
            Text            =   "Text1"
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ListBox lstFlag 
            Height          =   3060
            Index           =   0
            Left            =   -30
            Style           =   1  'Checkbox
            TabIndex        =   81
            Top             =   0
            Width           =   3255
         End
      End
   End
   Begin VB.TextBox txtcolor 
      Height          =   390
      Left            =   9480
      TabIndex        =   65
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame5 
      Caption         =   "Basic Colors"
      Height          =   2415
      Left            =   2040
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   3975
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   47
         Left            =   3480
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   63
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
         TabIndex        =   62
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
         TabIndex        =   61
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
         TabIndex        =   60
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
         TabIndex        =   59
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
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
         Tag             =   "04"
         Top             =   1680
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00FFC0FF&
         Height          =   285
         Index           =   33
         Left            =   600
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
         Tag             =   "61"
         Top             =   1320
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00800000&
         Height          =   285
         Index           =   29
         Left            =   2520
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   45
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
         TabIndex        =   44
         Tag             =   "03"
         Top             =   1320
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00FFFF80&
         Height          =   285
         Index           =   27
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   43
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
         TabIndex        =   42
         Tag             =   "0C"
         Top             =   1320
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H006FADE1&
         Height          =   285
         Index           =   25
         Left            =   600
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   41
         Tag             =   "EC"
         Top             =   1320
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00FFFFC0&
         Height          =   285
         Index           =   24
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
         Tag             =   "60"
         Top             =   960
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00C4D94A&
         Height          =   285
         Index           =   21
         Left            =   2520
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   37
         Tag             =   "6F"
         Top             =   960
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00A2BB28&
         Height          =   285
         Index           =   20
         Left            =   2040
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   36
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
         TabIndex        =   35
         Tag             =   "0D"
         Top             =   960
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Index           =   18
         Left            =   1080
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
         Tag             =   "0E"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00FF8080&
         Height          =   285
         Index           =   12
         Left            =   2040
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   28
         Tag             =   "1F"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   11
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   27
         Tag             =   "1C"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H0000FF00&
         Height          =   285
         Index           =   10
         Left            =   1080
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
         Tag             =   "0F"
         Top             =   210
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   4
         Left            =   2040
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   20
         Tag             =   "7F"
         Top             =   210
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H0080C0FF&
         Height          =   285
         Index           =   3
         Left            =   1560
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
         Tag             =   "FD"
         Top             =   210
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00458DCF&
         Height          =   285
         Index           =   0
         Left            =   120
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   16
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
         TabIndex        =   64
         Top             =   195
         Width           =   400
      End
   End
   Begin prjTouchScreen.MyButton cmdPrice13 
      Height          =   855
      Left            =   6600
      TabIndex        =   7
      Top             =   2145
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MyButton1"
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
      BCOL            =   16711680
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdPrice12 
      Height          =   855
      Left            =   3840
      TabIndex        =   6
      Top             =   2145
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MyButton1"
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
      BCOL            =   16711680
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdPrice11 
      Height          =   855
      Left            =   1200
      TabIndex        =   5
      Top             =   2145
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MyButton1"
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
      BCOL            =   16711680
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdGroup 
      Height          =   855
      Left            =   3840
      TabIndex        =   4
      Top             =   1245
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MyButton1"
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
      BCOL            =   16711680
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":008C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdUnit 
      Height          =   855
      Left            =   1200
      TabIndex        =   3
      Top             =   1245
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MyButton1"
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
      BCOL            =   16711680
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":00A8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdItemName 
      Height          =   1095
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   1931
      BTYPE           =   6
      TX              =   "Tªn"
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
      BCOL            =   16711680
      BCOLO           =   33023
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":00C4
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
      Height          =   1095
      Left            =   6360
      TabIndex        =   1
      Top             =   5040
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1931
      BTYPE           =   5
      TX              =   "§ãn&g"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":00E0
      PICN            =   "frmItem_Details.frx":00FC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdSave 
      Height          =   1095
      Left            =   3600
      TabIndex        =   0
      Top             =   5040
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1931
      BTYPE           =   5
      TX              =   "&L­u/ §ãng"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0736
      PICN            =   "frmItem_Details.frx":0752
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdPrice21 
      Height          =   855
      Left            =   1200
      TabIndex        =   8
      Top             =   3045
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MyButton1"
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
      BCOL            =   16711680
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0C96
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdPrice22 
      Height          =   855
      Left            =   3840
      TabIndex        =   9
      Top             =   3045
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MyButton1"
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
      BCOL            =   16711680
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0CB2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdPrice23 
      Height          =   855
      Left            =   6600
      TabIndex        =   10
      Top             =   3045
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MyButton1"
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
      BCOL            =   16711680
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0CCE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdPrice31 
      Height          =   855
      Left            =   1200
      TabIndex        =   11
      Top             =   3945
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MyButton1"
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
      BCOL            =   16711680
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0CEA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdPrice32 
      Height          =   855
      Left            =   3840
      TabIndex        =   12
      Top             =   3945
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MyButton1"
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
      BCOL            =   16711680
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0D06
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdPrice33 
      Height          =   855
      Left            =   6600
      TabIndex        =   13
      Top             =   3945
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "MyButton1"
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
      BCOL            =   16711680
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0D22
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdColor 
      Height          =   1095
      Left            =   10200
      TabIndex        =   14
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1931
      BTYPE           =   6
      TX              =   "Mµu hiÓn thÞ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArialH"
         Size            =   11.25
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
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0D3E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   1095
      Index           =   0
      Left            =   0
      TabIndex        =   66
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1931
      BTYPE           =   6
      TX              =   "Tªn mãn"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0D5A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   870
      Index           =   1
      Left            =   0
      TabIndex        =   67
      Top             =   1250
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1535
      BTYPE           =   5
      TX              =   "§VT"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0D76
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   870
      Index           =   2
      Left            =   2640
      TabIndex        =   68
      Top             =   1245
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1535
      BTYPE           =   5
      TX              =   "Nhãm"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0D92
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   870
      Index           =   3
      Left            =   0
      TabIndex        =   69
      Top             =   2150
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1535
      BTYPE           =   5
      TX              =   "Gi¸ 1:"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0DAE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   870
      Index           =   4
      Left            =   0
      TabIndex        =   70
      Top             =   3045
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1535
      BTYPE           =   5
      TX              =   "Gi¸ 4:"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0DCA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   870
      Index           =   5
      Left            =   0
      TabIndex        =   71
      Top             =   3940
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1535
      BTYPE           =   5
      TX              =   "Gi¸ 7:"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0DE6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   870
      Index           =   6
      Left            =   2640
      TabIndex        =   72
      Top             =   2145
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1535
      BTYPE           =   5
      TX              =   "Gi¸ 2:"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0E02
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   870
      Index           =   7
      Left            =   2640
      TabIndex        =   73
      Top             =   3045
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1535
      BTYPE           =   5
      TX              =   "Gi¸ 5:"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0E1E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   870
      Index           =   8
      Left            =   2640
      TabIndex        =   74
      Top             =   3945
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1535
      BTYPE           =   5
      TX              =   "Gi¸ 8:"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0E3A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   870
      Index           =   9
      Left            =   5400
      TabIndex        =   75
      Top             =   2145
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1535
      BTYPE           =   5
      TX              =   "Gi¸ 3:"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0E56
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   870
      Index           =   10
      Left            =   5400
      TabIndex        =   76
      Top             =   3045
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1535
      BTYPE           =   5
      TX              =   "Gi¸ 6:"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0E72
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   870
      Index           =   11
      Left            =   5400
      TabIndex        =   77
      Top             =   3945
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1535
      BTYPE           =   5
      TX              =   "Gi¸ 9:"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0E8E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdPrinter 
      Height          =   870
      Left            =   5400
      TabIndex        =   78
      Top             =   1245
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1535
      BTYPE           =   5
      TX              =   "M¸y In"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   12582912
      FCOLO           =   12582912
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmItem_Details.frx":0EAA
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
Attribute VB_Name = "frmItem_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsItem As New ADODB.Recordset
Dim rsDept As New ADODB.Recordset
Dim Item_Code As String
Dim fLoad As Boolean

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdColor_Click()
    With Frame5
            .Visible = True
            .top = cmdGroup.top
            .Left = cmdGroup.Left + cmdGroup.Width - Frame5.Width
        End With
End Sub

Private Sub cmdGroup_Click()
On Error GoTo Handle
   With frmDept_select
        .Show vbModal
        cmdGroup.Tag = .Return_Code
    End With
    With rsDept
        .Find "Dept_ID='" & cmdGroup.Tag & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            cmdGroup.Caption = "Thuéc nhãm: " & .Fields("Description")
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " -Form_Load"
End Sub

Private Sub cmdItemName_Click()
On Error GoTo Handle
    With frmKeyboard
        .FormCallkeyboard = "Other"
        .txtInput.PasswordChar = ""
        .txtInput.Text = cmdItemName.Caption
        .txtInput.SelStart = 0
        .txtInput.SelLength = 9999
        .Show vbModal
        cmdItemName.Caption = .Let_Text_Input
    End With
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdItemName_Click"
End Sub


Private Sub cmdPrice11_Click()
On Error GoTo Handle
    With frmPhimso
         .lblTitle.Caption = "NhËp gi¸ b¸n:"
        .FormCall = 3
        .txtQty.Text = cmdPrice11.Caption
        .txtQty.SelStart = 0
        .txtQty.SelLength = 9999
        .Show vbModal
        cmdPrice11.Caption = Format(.Return_Value, "#,##0")
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdPrice11_Click"
End Sub

Private Sub cmdPrice12_Click()
On Error GoTo Handle
    With frmPhimso
    .lblTitle.Caption = "NhËp gi¸ b¸n:"
        .FormCall = 3
        .txtQty.Text = cmdPrice12.Caption
        .txtQty.SelStart = 0
        .txtQty.SelLength = 9999
        .Show vbModal
        cmdPrice12.Caption = Format(.Return_Value, "#,##0")
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdPrice12"
End Sub

Private Sub cmdPrice13_Click()
On Error GoTo Handle
    With frmPhimso
    .lblTitle.Caption = "NhËp gi¸ b¸n:"
        .FormCall = 3
        .txtQty.Text = cmdPrice13.Caption
        .txtQty.SelStart = 0
        .txtQty.SelLength = 9999
        .Show vbModal
        cmdPrice13.Caption = Format(.Return_Value, "#,##0")
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdPrice13"
End Sub

Private Sub cmdPrice21_Click()
On Error GoTo Handle
    With frmPhimso
    .lblTitle.Caption = "NhËp gi¸ b¸n:"
        .FormCall = 3
        .txtQty.Text = cmdPrice21.Caption
        .txtQty.SelStart = 0
        .txtQty.SelLength = 9999
        .Show vbModal
        cmdPrice21.Caption = Format(.Return_Value, "#,##0")
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdPrice21"
End Sub

Private Sub cmdPrice22_Click()
On Error GoTo Handle
    With frmPhimso
    .lblTitle.Caption = "NhËp gi¸ b¸n:"
        .FormCall = 3
        .txtQty.Text = cmdPrice22.Caption
        .txtQty.SelStart = 0
        .txtQty.SelLength = 9999
        .Show vbModal
        cmdPrice22.Caption = Format(.Return_Value, "#,##0")
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdPrice22"
End Sub

Private Sub cmdPrice23_Click()
On Error GoTo Handle
    With frmPhimso
    .lblTitle.Caption = "NhËp gi¸ b¸n:"
        .FormCall = 3
        .txtQty.Text = cmdPrice23.Caption
        .txtQty.SelStart = 0
        .txtQty.SelLength = 9999
        .Show vbModal
        cmdPrice23.Caption = Format(.Return_Value, "#,##0")
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdPrice23"
End Sub

Private Sub cmdPrice31_Click()
On Error GoTo Handle
    With frmPhimso
        .FormCall = 3
        .lblTitle.Caption = "NhËp gi¸ b¸n:"
        .txtQty.Text = cmdPrice31.Caption
        .txtQty.SelStart = 0
        .txtQty.SelLength = 9999
        .Show vbModal
        cmdPrice31.Caption = Format(.Return_Value, "#,##0")
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdPrice31"

End Sub

Private Sub cmdPrice32_Click()
On Error GoTo Handle
    With frmPhimso
    .lblTitle.Caption = "NhËp gi¸ b¸n:"
        .FormCall = 3
        .txtQty.Text = cmdPrice32.Caption
        .txtQty.SelStart = 0
        .txtQty.SelLength = 9999
        .Show vbModal
        cmdPrice32.Caption = Format(.Return_Value, "#,##0")
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdPrice32"
End Sub

Private Sub cmdPrice33_Click()
On Error GoTo Handle
    With frmPhimso
        .lblTitle.Caption = "NhËp gi¸ b¸n:"
        .FormCall = 3
        .txtQty.Text = cmdPrice33.Caption
        .txtQty.SelStart = 0
        .txtQty.SelLength = 9999
        .Show vbModal
        cmdPrice33.Caption = Format(.Return_Value, "#,##0")
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdPrice33"

End Sub

Private Sub cmdSave_Click()
On Error GoTo Handle
    With rsItem
        .Find "ItemNum='" & Item_Code & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Fields("ItemName") = cmdItemName.Caption
                .Fields("Unit") = cmdUnit.Caption
                .Fields("Dept_ID") = cmdGroup.Tag
                .Fields("Std_Price1") = cmdPrice11.Caption
                .Fields("Std_Price2") = cmdPrice12.Caption
                .Fields("Std_Price3") = cmdPrice13.Caption
                .Fields("HH_Price1") = cmdPrice21.Caption
                .Fields("HH_Price2") = cmdPrice22.Caption
                .Fields("HH_Price3") = cmdPrice23.Caption
                .Fields("EV_Price1") = cmdPrice31.Caption
                .Fields("EV_Price2") = cmdPrice32.Caption
                .Fields("EV_Price3") = cmdPrice33.Caption
                .Fields("LimitPrice") = txtcolor.Text
                .Fields("F2") = txtFlag(1).Text
                .Fields("F1") = txtFlag(0).Text
                .Update
            End If
    End With
    Unload Me
Exit Sub
Handle:
    MsgBox Err.Number & " D÷ liÖu b¹n võa thay ®æi kh«ng phï hîp " & " -cmdSave_Click"

End Sub

Private Sub cmdUnit_Click()
On Error GoTo Handle
    With frmKeyboard
        .FormCallkeyboard = "Other"
        .txtInput.PasswordChar = ""
        .txtInput.Text = cmdUnit.Caption
        .txtInput.SelStart = 0
        .txtInput.SelLength = 9999
        .Show vbModal
        cmdUnit.Caption = .Let_Text_Input
    End With
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdItemName_Click"
End Sub

Private Sub color_Click(Index As Integer)
    Call cmdColor_Click
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Set rsItem = Open_Table(cnData, "Inventory")
    Set rsDept = Open_Table(cnData, "Departments")
    Call Load_Details
    Call Add_Flag_Items
    For i = o To txtFlag.count - 1
        Call AddValueForList(txtFlag(i).Text, lstFlag(i))
    Next
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " -Form_Load"
End Sub

Public Property Let Get_Item_Code(ByVal vNewValue As Variant)
    Item_Code = vNewValue
End Property


Public Sub Load_Details()
On Error GoTo Handle
    With rsItem
        If Not .EOF Then
            .Find "ItemNum='" & Item_Code & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                cmdItemName.Caption = .Fields("ItemName")
                cmdItemName.BackColor = HexToDec(.Fields("LimitPrice"))
                cmdUnit.Caption = .Fields("Unit")
                cmdGroup.Tag = .Fields("Dept_ID")
                txtFlag(1).Text = .Fields("F2")
                txtFlag(0).Text = .Fields("F1")
                With rsDept
                    .Find "Dept_ID='" & cmdGroup.Tag & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        cmdGroup.Caption = "Thuéc nhãm: " & .Fields("Description")
                    End If
                End With
                cmdPrice11.Caption = Format(.Fields("Std_Price1"), "#,##0")
                cmdPrice12.Caption = Format(.Fields("Std_Price2"), "#,##0")
                cmdPrice13.Caption = Format(.Fields("Std_Price3"), "#,##0")
                cmdPrice21.Caption = Format(.Fields("HH_Price1"), "#,##0")
                cmdPrice22.Caption = Format(.Fields("HH_Price2"), "#,##0")
                cmdPrice23.Caption = Format(.Fields("HH_Price3"), "#,##0")
                cmdPrice31.Caption = Format(.Fields("EV_Price1"), "#,##0")
                cmdPrice32.Caption = Format(.Fields("EV_Price2"), "#,##0")
                cmdPrice33.Caption = Format(.Fields("EV_Price3"), "#,##0")
                txtcolor.Text = .Fields("LimitPrice")
            End If
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - Load_Details"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Item_Code = ""
    CloseRecordset rsItem
    CloseRecordset rsDept
    Set cnData = Nothing
End Sub

Private Sub picBasicColor_Click(Index As Integer)
    On Error GoTo Handle
       
        Frame5.Visible = False
        txtcolor.Text = DectoHex(picBasicColor(Index).BackColor)
        cmdItemName.BackColor = picBasicColor(Index).BackColor
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " picBasicColor_Click"
    
End Sub

Public Sub Add_Flag_Items()
On Error GoTo Handle
    Dim arrFlag() As String
    Dim j, i As Integer
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
            End Select
    Next i
        fLoad = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "Add_Flag_Items"
End Sub

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
    End If

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - lstFlag_Click "
End Sub
