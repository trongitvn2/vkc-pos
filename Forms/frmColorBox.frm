VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmColorBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Color Box"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10275
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Table Plan Color Setting"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6000
      TabIndex        =   73
      Top             =   0
      Width           =   4215
      Begin VB.TextBox txtBackColor 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   77
         Text            =   "Background Table Plan"
         Top             =   1200
         Width           =   3855
      End
      Begin VB.TextBox txtShapecolor 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   76
         Text            =   "Shape Color"
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox txtFontColor 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   75
         Text            =   "Fontcolor "
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cboFont 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Text            =   "cboFont"
         Top             =   240
         Width           =   3855
      End
   End
   Begin prjTouchScreen.MyButton cmdReserve 
      Height          =   735
      Index           =   0
      Left            =   6060
      TabIndex        =   64
      Top             =   1740
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      BTYPE           =   4
      TX              =   "ORDERED"
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
      BCOL            =   12648447
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmColorBox.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Frame Frame2 
      Caption         =   "Basic Colors"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      Tag             =   "L1"
      Top             =   60
      Width           =   5775
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   71
         Left            =   4920
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   101
         Tag             =   "EF"
         Top             =   210
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Index           =   70
         Left            =   4920
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   100
         Tag             =   "EF"
         Top             =   2010
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00C0C000&
         Height          =   285
         Index           =   69
         Left            =   4920
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   99
         Tag             =   "EF"
         Top             =   1680
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00404080&
         Height          =   285
         Index           =   68
         Left            =   4920
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   98
         Tag             =   "EF"
         Top             =   1320
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H000000C0&
         Height          =   285
         Index           =   67
         Left            =   4920
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   97
         Tag             =   "EF"
         Top             =   960
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00004080&
         Height          =   285
         Index           =   66
         Left            =   4920
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   96
         Tag             =   "EF"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H0000C0C0&
         Height          =   285
         Index           =   65
         Left            =   4560
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   95
         Tag             =   "EF"
         Top             =   210
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00808000&
         Height          =   285
         Index           =   64
         Left            =   4560
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   94
         Tag             =   "EF"
         Top             =   2010
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00800000&
         Height          =   285
         Index           =   63
         Left            =   4560
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   93
         Tag             =   "EF"
         Top             =   1680
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00004040&
         Height          =   285
         Index           =   62
         Left            =   4560
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   92
         Tag             =   "EF"
         Top             =   1320
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00008080&
         Height          =   285
         Index           =   61
         Left            =   4560
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   91
         Tag             =   "EF"
         Top             =   960
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00008080&
         Height          =   285
         Index           =   60
         Left            =   4560
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   90
         Tag             =   "EF"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   59
         Left            =   4200
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   89
         Tag             =   "EF"
         Top             =   210
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H008080FF&
         Height          =   285
         Index           =   58
         Left            =   4200
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   88
         Tag             =   "EF"
         Top             =   2040
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Index           =   57
         Left            =   4200
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   87
         Tag             =   "EF"
         Top             =   1680
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00404040&
         Height          =   285
         Index           =   56
         Left            =   4200
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   86
         Tag             =   "EF"
         Top             =   1320
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00808080&
         Height          =   285
         Index           =   55
         Left            =   4200
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   85
         Tag             =   "EF"
         Top             =   960
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Index           =   54
         Left            =   4200
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   84
         Tag             =   "EF"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   53
         Left            =   3840
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   83
         Tag             =   "EF"
         Top             =   2040
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00004000&
         Height          =   285
         Index           =   52
         Left            =   3840
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   82
         Tag             =   "EF"
         Top             =   1680
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00008080&
         Height          =   285
         Index           =   51
         Left            =   3840
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   81
         Tag             =   "EF"
         Top             =   1320
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00004040&
         Height          =   285
         Index           =   50
         Left            =   3840
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   80
         Tag             =   "EF"
         Top             =   960
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00808000&
         Height          =   285
         Index           =   49
         Left            =   3840
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   79
         Tag             =   "EF"
         Top             =   600
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00008080&
         FillColor       =   &H00008080&
         FillStyle       =   0  'Solid
         Height          =   285
         Index           =   48
         Left            =   3840
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   78
         Tag             =   "EF"
         Top             =   210
         Width           =   375
      End
      Begin VB.PictureBox picBasicColor 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   47
         Left            =   3480
         ScaleHeight     =   225
         ScaleWidth      =   315
         TabIndex        =   58
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   55
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
         TabIndex        =   54
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
         TabIndex        =   53
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
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
         TabIndex        =   49
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
         TabIndex        =   48
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   59
         Top             =   195
         Width           =   400
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Custom Colors"
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   2490
      Width           =   5775
      Begin VB.TextBox txtSColor 
         Height          =   495
         Left            =   1080
         TabIndex        =   5
         Top             =   1830
         Width           =   4455
      End
      Begin MSComctlLib.Slider sliColor 
         Height          =   435
         Index           =   0
         Left            =   915
         TabIndex        =   2
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   767
         _Version        =   393216
         LargeChange     =   1
         Max             =   7
         SelectRange     =   -1  'True
      End
      Begin MSComctlLib.Slider sliColor 
         Height          =   435
         Index           =   1
         Left            =   915
         TabIndex        =   3
         Top             =   630
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   767
         _Version        =   393216
         LargeChange     =   1
         Max             =   7
         SelectRange     =   -1  'True
      End
      Begin MSComctlLib.Slider sliColor 
         Height          =   435
         Index           =   2
         Left            =   915
         TabIndex        =   4
         Top             =   990
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   767
         _Version        =   393216
         LargeChange     =   1
         Max             =   3
         SelectRange     =   -1  'True
      End
      Begin VB.Label lblCustom 
         Alignment       =   1  'Right Justify
         Caption         =   "Blue:"
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
         Left            =   45
         TabIndex        =   10
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblHex 
         Caption         =   "Value Hex"
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
         Left            =   1200
         TabIndex        =   60
         Top             =   1590
         Width           =   735
      End
      Begin VB.Label lblCustom 
         Alignment       =   1  'Right Justify
         Caption         =   "Green:"
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
         Left            =   45
         TabIndex        =   9
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblCustom 
         Alignment       =   1  'Right Justify
         Caption         =   "Red:"
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
         Left            =   45
         TabIndex        =   8
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lblBasic 
         Caption         =   "Sample Color:"
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
         TabIndex        =   6
         Top             =   1590
         Width           =   1215
      End
   End
   Begin VB.Frame fraButton 
      Height          =   945
      Left            =   5940
      TabIndex        =   0
      Top             =   4980
      Width           =   4215
      Begin prjTouchScreen.MyButton cmdOK 
         Height          =   705
         Left            =   30
         TabIndex        =   62
         Tag             =   "L2"
         Top             =   150
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&L­u/Tho¸t"
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
         BCOL            =   16578804
         BCOLO           =   16777152
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmColorBox.frx":001C
         PICN            =   "frmColorBox.frx":0038
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
         Height          =   705
         Left            =   1410
         TabIndex        =   63
         Tag             =   "L3"
         Top             =   150
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "Gióp ®ì"
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
         BCOL            =   16578804
         BCOLO           =   16777152
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmColorBox.frx":057C
         PICN            =   "frmColorBox.frx":0598
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdSetTime 
         Height          =   705
         Left            =   2790
         TabIndex        =   72
         Tag             =   "L4"
         Top             =   150
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Thêi gian Thanh to¸n/ Dän dÑp"
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
         BCOL            =   16578804
         BCOLO           =   16777152
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmColorBox.frx":0BD2
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
   Begin prjTouchScreen.MyButton cmdReserve 
      Height          =   735
      Index           =   2
      Left            =   6060
      TabIndex        =   65
      Top             =   2610
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      BTYPE           =   4
      TX              =   "RESERVED"
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
      BCOL            =   12648447
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmColorBox.frx":0BEE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdReserve 
      Height          =   735
      Index           =   7
      Left            =   8100
      TabIndex        =   66
      Top             =   4230
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      BTYPE           =   4
      TX              =   "VACANT"
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
      BCOL            =   12648447
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmColorBox.frx":0C0A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdReserve 
      Height          =   735
      Index           =   6
      Left            =   6060
      TabIndex        =   67
      Top             =   4230
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      BTYPE           =   4
      TX              =   "CLEANING"
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
      BCOL            =   12648447
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmColorBox.frx":0C26
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdReserve 
      Height          =   735
      Index           =   3
      Left            =   8100
      TabIndex        =   68
      Top             =   2610
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      BTYPE           =   4
      TX              =   "SEATED"
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
      BCOL            =   12648447
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmColorBox.frx":0C42
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdReserve 
      Height          =   735
      Index           =   1
      Left            =   8100
      TabIndex        =   69
      Top             =   1740
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      BTYPE           =   4
      TX              =   "MULTIPLE"
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
      BCOL            =   12648447
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmColorBox.frx":0C5E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdReserve 
      Height          =   735
      Index           =   5
      Left            =   8100
      TabIndex        =   70
      Top             =   3420
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      BTYPE           =   4
      TX              =   "PAID"
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
      BCOL            =   12648447
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmColorBox.frx":0C7A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdReserve 
      Height          =   735
      Index           =   4
      Left            =   6060
      TabIndex        =   71
      Top             =   3420
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      BTYPE           =   4
      TX              =   "SUBTOTAL"
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
      BCOL            =   12648447
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmColorBox.frx":0C96
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
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
      Left            =   1080
      TabIndex        =   61
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmColorBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim sBackColor As String
    Dim sHexcolor As String
    Dim k As Integer
    Dim rscolor As New ADODB.Recordset
    Dim strReseveColor As String
    Dim iIndex As Integer
    Dim DescArr() As String
    Dim fontclick As Boolean
    Dim ReserveClick, shapecolorclick, bkcolorclick As Boolean


Private Sub cboFont_Change()
    CurFont = cbofont.Text
End Sub

Private Sub cmdReserve_Click(Index As Integer)
ReserveClick = True
    strReseveColor = cmdReserve(Index).Caption
    txtSColor.Text = strReseveColor
    iIndex = Index
    With rscolor
        .Find "ReserveType='" & strReseveColor & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
             txtSColor.BackColor = !ReserveValue
             cmdReserve(Index).BackColor = !ReserveValue
        End If
    End With
   fontclick = False
    bkcolorclick = False
    shapecolorclick = False
End Sub

Private Sub cmdSetTime_Click()
    frmPaymentTimer.Show vbModal
End Sub

'            ------- FORM -------
Private Sub Form_Load()
On Error GoTo errHdl
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    If cnData.State <> 0 Then
        Set rscolor = Open_Table(cnData, "ColorTablePlan")
    End If
    Add_Font
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Load"
End Sub

Private Sub Form_Activate()
On Error GoTo errHdl
    Dim i As Integer
    Dim sDefaultColor As String
    Dim sDefaultHex As String
    Dim ctrl As Control
    Dim DescArr() As String
    If cmdOk.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    DescArr = LoadLanguage(LngFile, "#03:012:")
    Me.Caption = DescArr(1)
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl

    For i = 0 To cmdReserve.count - 1
        With rscolor
            .Find "ID='" & i + 1 & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                cmdReserve(i).BackColor = !ReserveValue
            End If
        End With
    Next
    SetValueForSliColor
    cbofont.Text = CurFont
    txtFontColor.BackColor = ColorFont
    txtShapecolor.BackColor = ShapeColor
    txtBackColor.BackColor = bkColor
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Activate"
End Sub
'            ------- COMMAND BUTTON -------


Private Sub cmdOK_Click()
On Error GoTo errHdl

    sBackColor = txtSColor.BackColor
    sHexcolor = txtSColor.Tag
    SaveSettingStr "SYSTEM", "Font", cbofont.Text, myIniFile
    SaveSettingStr "SYSTEM", "FontColor", txtFontColor.BackColor, myIniFile
    SaveSettingStr "SYSTEM", "ShapeColor", ShapeColor, myIniFile
    SaveSettingStr "SYSTEM", "BkColor", bkColor, myIniFile
    Unload Me
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- cmdOK_Click"
End Sub

Private Sub Form_Unload(Cancel As Integer)
fontclick = False
bkcolorclick = False
ReserveClick = False
shapecolorclick = False
ReserveClick = False
End Sub

'            ------- PICTUREBOX -------
Private Sub picBasicColor_Click(Index As Integer)
On Error GoTo errHdl

    With picBasicColor(Index)
        lblMove.Left = .Left - 15
        lblMove.top = .top - 15
        shpMove.Left = .Left - 40
        shpMove.top = .top - 50
        txtSColor.BackColor = .BackColor
        txtSColor.Tag = .Tag
        lblHex.ForeColor = .BackColor
        lblHex.Caption = .Tag
        Label2.Caption = .BackColor
        SetValueForSliColor
    End With
    With rscolor
        If strReseveColor <> "" Then
            .Find "ReserveType='" & strReseveColor & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Fields("ReserveValue") = picBasicColor(Index).BackColor
                .Update
            Else
                .addNew
                .Fields("ReserveType") = strReseveColor
                .Fields("ReserveValue") = picBasicColor(Index).BackColor
                .Update
            End If
        End If
    End With
    If ReserveClick Then
        cmdReserve(iIndex).BackColor = picBasicColor(Index).BackColor
        ReserveClick = False
    End If
    If fontclick Then
        txtFontColor.BackColor = picBasicColor(Index).BackColor
        ColorFont = txtFontColor.BackColor
        fontclick = False
    End If
    If shapecolorclick Then
        txtShapecolor.BackColor = picBasicColor(Index).BackColor
        ShapeColor = txtShapecolor.BackColor
        shapecolorclick = False
    End If
    If bkcolorclick = True Then
        txtBackColor.BackColor = picBasicColor(Index).BackColor
        bkColor = txtBackColor.BackColor
        bkcolorclick = False
    End If
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- picBasicColor_Click"
End Sub
'            ------- SLIDER -------
Private Sub sliColor_Change(Index As Integer)
On Error GoTo errHdl

    SetColor
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- sliColor_Change"
End Sub
'            ------- FUNCTION -------
Private Sub SetColor()
On Error GoTo errHdl

    With txtSColor
        .BackColor = "&H" & Right("00" & Hex(sliColor(2).Value * 85), 2) _
                          & Right("00" & Hex(sliColor(1).Value * 36), 2) _
                          & Right("00" & Hex(sliColor(0).Value * 36), 2)
        .Tag = Change_Color_To_Hex
        lblHex.ForeColor = .BackColor
        lblHex.Caption = .Tag
        Label2.Caption = .BackColor
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- SetColor"
End Sub

Public Property Get Get_Color() As String
    Get_Color = sBackColor
End Property

Public Property Get Get_HexColor() As String
    Get_HexColor = sHexcolor
End Property

Private Function Change_Color_To_Hex() As String
On Error GoTo errHdl

'   Hex(SysColor)-> BGR (vd: FF00FF) ->  RGB(Dec(R)&Dec(G)/36; Dec(B)/85)
'   Hex(Bin(R) or Bin(G) or Bin(B))
'   Red  :    X00000  (X: 3 chu so)
'   Green:    000X00  (X: 3 chu so)
'   Blue :    000000X (X: 2 chu so)
    Dim Arr(3) As String
    Dim svalue As String
    
    Arr(1) = FillZeroForString(DecToBin(sliColor(0).Value), 3) & "00000" 'red
    Arr(2) = "000" & FillZeroForString(DecToBin(sliColor(1).Value), 3) & "00"  'green
    Arr(3) = FillZeroForString(DecToBin(sliColor(2).Value), 8)  'blue
    For k = 1 To 8
    DoEvents
        svalue = svalue & Mid(Arr(1), k, 1) Or Mid(Arr(2), k, 1) Or Mid(Arr(3), k, 1)
    Next k
    svalue = BinToHex(svalue)
    Change_Color_To_Hex = svalue
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Change_Color_To_Hex"
End Function

Private Sub SetValueForSliColor()
On Error GoTo errHdl

    Dim sRed As String
    Dim sGreen As String
    Dim sBlue As String
    Dim sTemp As String
    
    sTemp = DectoHex(txtSColor.BackColor)
    sRed = HexToDec(Right(sTemp, 2))
    sGreen = HexToDec(Mid(sTemp, 3, 2))
    sBlue = HexToDec(Left(sTemp, 2))
    sliColor(0).Value = CInt(sRed) / 36
    sliColor(1).Value = CInt(sGreen) / 36
    sliColor(2).Value = CInt(sBlue) / 85
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- SetValueForSliColor"
End Sub


Private Sub txtBackColor_Click()
    fontclick = False
    bkcolorclick = True
    ReserveClick = False
    shapecolorclick = False
End Sub

Private Sub txtFontColor_Click()
    fontclick = True
    bkcolorclick = False
    ReserveClick = False
    shapecolorclick = False
End Sub

Public Sub Add_Font()
On Error GoTo Handle
    With cbofont
        .Clear
        .AddItem ".VnArial"
        .AddItem ".VnArialH"
        .AddItem ".VnArial Narrow"
        .AddItem ".VnArial NarrowH"
        .AddItem ".VnArial NarrowH"
        .AddItem ".VnTime"
        .AddItem ".VnTimeH"
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub


Private Sub txtShapecolor_Click()
    fontclick = False
    bkcolorclick = False
    ReserveClick = False
    shapecolorclick = True
End Sub
