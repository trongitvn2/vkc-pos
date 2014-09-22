VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSetupKaraoke 
   Caption         =   "Setup Karaoke price"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   240
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStd_Price 
      Caption         =   "Standard Price"
      Height          =   2895
      Left            =   120
      TabIndex        =   47
      Top             =   840
      Width           =   15255
      Begin VB.Frame Frame8 
         Caption         =   "Gi¸ 4"
         Height          =   2535
         Index           =   2
         Left            =   13680
         TabIndex        =   97
         Top             =   240
         Width           =   1455
         Begin VB.TextBox txtStd_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   100
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtStd_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   99
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtStd_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   98
            Top             =   1800
            Width           =   1215
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Gi¸ 1"
         Height          =   2535
         Left            =   9120
         TabIndex        =   93
         Top             =   240
         Width           =   1335
         Begin VB.TextBox txtStd_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   96
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox txtStd_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   95
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtStd_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   94
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Gi¸ 3"
         Height          =   2535
         Index           =   1
         Left            =   12120
         TabIndex        =   65
         Top             =   240
         Width           =   1455
         Begin VB.TextBox txtStd_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   68
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txtStd_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   67
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtStd_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Gi¸ 2"
         Height          =   2535
         Index           =   0
         Left            =   10560
         TabIndex        =   61
         Top             =   240
         Width           =   1455
         Begin VB.TextBox txtStd_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtStd_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   63
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtStd_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   62
            Top             =   1800
            Width           =   1215
         End
      End
      Begin VB.Frame fraWeek 
         Caption         =   "C¸c ngµy trong tuÇn"
         Height          =   2295
         Left            =   120
         TabIndex        =   56
         Top             =   480
         Width           =   4095
         Begin VB.ComboBox cboStd_Weekday_From 
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
            Left            =   480
            TabIndex        =   58
            Text            =   "Combo1"
            Top             =   840
            Width           =   1455
         End
         Begin VB.ComboBox cboStd_Weekday_To 
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
            Left            =   2520
            TabIndex        =   57
            Text            =   "Combo1"
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Tõ:"
            Height          =   375
            Left            =   120
            TabIndex        =   60
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label3 
            Caption         =   "§Õn:"
            Height          =   375
            Left            =   2040
            TabIndex        =   59
            Top             =   960
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Left            =   4320
         TabIndex        =   48
         Top             =   240
         Width           =   4695
         Begin MSComCtl2.DTPicker dtpStdFrom 
            Height          =   495
            Index           =   0
            Left            =   600
            TabIndex        =   1
            Top             =   360
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "H:mm:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.25
         End
         Begin MSComCtl2.DTPicker dtpStdTo 
            Height          =   495
            Index           =   0
            Left            =   2880
            TabIndex        =   2
            Top             =   360
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "H:mm:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.5826388889
         End
         Begin MSComCtl2.DTPicker dtpStdFrom 
            Height          =   495
            Index           =   1
            Left            =   600
            TabIndex        =   3
            Top             =   1080
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.25
         End
         Begin MSComCtl2.DTPicker dtpStdTo 
            Height          =   495
            Index           =   1
            Left            =   2880
            TabIndex        =   4
            Top             =   1080
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.5826388889
         End
         Begin MSComCtl2.DTPicker dtpStdFrom 
            Height          =   495
            Index           =   2
            Left            =   600
            TabIndex        =   5
            Top             =   1800
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.25
         End
         Begin MSComCtl2.DTPicker dtpStdTo 
            Height          =   495
            Index           =   2
            Left            =   2880
            TabIndex        =   6
            Top             =   1800
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.5826388889
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "§Õn:"
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
            Height          =   285
            Left            =   2400
            TabIndex        =   54
            Tag             =   "L7"
            Top             =   480
            Width           =   525
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Tõ :"
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
            Height          =   285
            Left            =   120
            TabIndex        =   53
            Tag             =   "L6"
            Top             =   480
            Width           =   405
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "§Õn:"
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
            Height          =   285
            Left            =   2400
            TabIndex        =   52
            Tag             =   "L7"
            Top             =   1200
            Width           =   525
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Tõ :"
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
            Height          =   285
            Left            =   120
            TabIndex        =   51
            Tag             =   "L6"
            Top             =   1200
            Width           =   405
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "§Õn:"
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
            Height          =   285
            Left            =   2400
            TabIndex        =   50
            Tag             =   "L7"
            Top             =   1920
            Width           =   525
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "Tõ :"
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
            Height          =   285
            Left            =   120
            TabIndex        =   49
            Tag             =   "L6"
            Top             =   1920
            Width           =   405
         End
      End
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Height          =   855
      Left            =   8880
      TabIndex        =   45
      Top             =   9960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "&§ãng"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSetupKaraoke.frx":0000
      PICN            =   "frmSetupKaraoke.frx":001C
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
      Height          =   855
      Left            =   7200
      TabIndex        =   44
      Top             =   9960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "&Gióp ®ì"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSetupKaraoke.frx":62B6
      PICN            =   "frmSetupKaraoke.frx":62D2
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
      Height          =   855
      Left            =   5520
      TabIndex        =   43
      Top             =   9960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "&L­u"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSetupKaraoke.frx":690C
      PICN            =   "frmSetupKaraoke.frx":6928
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
      Caption         =   "Holidays price"
      Height          =   2895
      Left            =   120
      TabIndex        =   32
      Top             =   6840
      Width           =   15255
      Begin VB.Frame Frame11 
         Caption         =   "Gi¸ 4"
         Height          =   2535
         Index           =   3
         Left            =   13680
         TabIndex        =   105
         Top             =   240
         Width           =   1455
         Begin VB.TextBox txtHL_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   108
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtHL_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   107
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtHL_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   106
            Top             =   1800
            Width           =   1215
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Gi¸ 3"
         Height          =   2535
         Index           =   2
         Left            =   12120
         TabIndex        =   89
         Top             =   240
         Width           =   1455
         Begin VB.TextBox txtHL_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   92
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txtHL_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   91
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtHL_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   90
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Gi¸ 2"
         Height          =   2535
         Index           =   1
         Left            =   10560
         TabIndex        =   85
         Top             =   240
         Width           =   1455
         Begin VB.TextBox txtHL_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   88
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtHL_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   87
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtHL_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   86
            Top             =   1800
            Width           =   1215
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Gi¸ 1"
         Height          =   2535
         Index           =   0
         Left            =   9120
         TabIndex        =   73
         Top             =   240
         Width           =   1335
         Begin VB.TextBox txtHL_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   76
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox txtHL_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   75
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtHL_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   74
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "C¸c ngµy lÔ trong n¨m (dd/MM)"
         Height          =   2295
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   4095
         Begin VB.TextBox txtHoliday 
            Alignment       =   2  'Center
            Height          =   495
            Left            =   120
            MaxLength       =   5
            TabIndex        =   46
            Top             =   360
            Width           =   1575
         End
         Begin prjTouchScreen.MyButton cmdAdd 
            Height          =   495
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            BTYPE           =   14
            TX              =   "Thªm v¸o list >>"
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
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmSetupKaraoke.frx":6E6C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin VB.ListBox lstHolidays 
            Height          =   1635
            Left            =   1800
            TabIndex        =   41
            Top             =   360
            Width           =   2175
         End
         Begin prjTouchScreen.MyButton cmdDelete 
            Height          =   495
            Left            =   120
            TabIndex        =   55
            Top             =   1680
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   873
            BTYPE           =   14
            TX              =   "<< Xãa khái list"
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
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmSetupKaraoke.frx":6E88
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
      Begin VB.Frame Frame6 
         Height          =   2535
         Left            =   4320
         TabIndex        =   33
         Top             =   240
         Width           =   4695
         Begin MSComCtl2.DTPicker dtpHLFrom 
            Height          =   495
            Index           =   0
            Left            =   600
            TabIndex        =   15
            Top             =   360
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.25
         End
         Begin MSComCtl2.DTPicker dtpHLTo 
            Height          =   495
            Index           =   0
            Left            =   2880
            TabIndex        =   16
            Top             =   360
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.5826388889
         End
         Begin MSComCtl2.DTPicker dtpHLFrom 
            Height          =   495
            Index           =   1
            Left            =   600
            TabIndex        =   17
            Top             =   1080
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.25
         End
         Begin MSComCtl2.DTPicker dtpHLTo 
            Height          =   495
            Index           =   1
            Left            =   2880
            TabIndex        =   18
            Top             =   1080
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.5826388889
         End
         Begin MSComCtl2.DTPicker dtpHLFrom 
            Height          =   495
            Index           =   2
            Left            =   600
            TabIndex        =   19
            Top             =   1800
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.25
         End
         Begin MSComCtl2.DTPicker dtpHLTo 
            Height          =   495
            Index           =   2
            Left            =   2880
            TabIndex        =   20
            Top             =   1800
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.5826388889
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            Caption         =   "§Õn:"
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
            Height          =   285
            Left            =   2400
            TabIndex        =   39
            Tag             =   "L7"
            Top             =   480
            Width           =   525
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Tõ :"
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
            Height          =   285
            Left            =   120
            TabIndex        =   38
            Tag             =   "L6"
            Top             =   480
            Width           =   405
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "§Õn:"
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
            Height          =   285
            Left            =   2400
            TabIndex        =   37
            Tag             =   "L7"
            Top             =   1200
            Width           =   525
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Tõ :"
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
            Height          =   285
            Left            =   120
            TabIndex        =   36
            Tag             =   "L6"
            Top             =   1200
            Width           =   405
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            Caption         =   "§Õn:"
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
            Height          =   285
            Left            =   2400
            TabIndex        =   35
            Tag             =   "L7"
            Top             =   1920
            Width           =   525
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "Tõ :"
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
            Height          =   285
            Left            =   120
            TabIndex        =   34
            Tag             =   "L6"
            Top             =   1920
            Width           =   405
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Happy Hour price"
      Height          =   2895
      Left            =   120
      TabIndex        =   21
      Top             =   3840
      Width           =   15255
      Begin VB.Frame Frame10 
         Caption         =   "Gi¸ 4"
         Height          =   2535
         Index           =   3
         Left            =   13680
         TabIndex        =   101
         Top             =   240
         Width           =   1455
         Begin VB.TextBox txtHP_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   9
            Left            =   120
            TabIndex        =   104
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtHP_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   10
            Left            =   120
            TabIndex        =   103
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtHP_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   11
            Left            =   120
            TabIndex        =   102
            Top             =   1800
            Width           =   1215
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Gi¸ 3"
         Height          =   2535
         Index           =   2
         Left            =   12120
         TabIndex        =   81
         Top             =   240
         Width           =   1455
         Begin VB.TextBox txtHP_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   8
            Left            =   120
            TabIndex        =   84
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txtHP_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   7
            Left            =   120
            TabIndex        =   83
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtHP_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   6
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Gi¸ 2"
         Height          =   2535
         Index           =   1
         Left            =   10560
         TabIndex        =   77
         Top             =   240
         Width           =   1455
         Begin VB.TextBox txtHP_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   80
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtHP_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   79
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtHP_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   5
            Left            =   120
            TabIndex        =   78
            Top             =   1800
            Width           =   1215
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Gi¸ 1"
         Height          =   2535
         Index           =   0
         Left            =   9120
         TabIndex        =   69
         Top             =   240
         Width           =   1335
         Begin VB.TextBox txtHP_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   72
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox txtHP_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   71
            Top             =   1080
            Width           =   1095
         End
         Begin VB.TextBox txtHP_Price 
            Alignment       =   1  'Right Justify
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   70
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "C¸c ngµy trong tuÇn"
         Height          =   2295
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   4095
         Begin VB.ComboBox cboH_Weekday_To 
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
            Left            =   2400
            TabIndex        =   8
            Text            =   "Combo1"
            Top             =   840
            Width           =   1575
         End
         Begin VB.ComboBox cboH_Weekday_From 
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
            Left            =   480
            TabIndex        =   7
            Text            =   "Combo1"
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label23 
            Caption         =   "Tõ:"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   375
         End
         Begin VB.Label Label22 
            Caption         =   "§Õn:"
            Height          =   375
            Left            =   1920
            TabIndex        =   30
            Top             =   960
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2535
         Left            =   4320
         TabIndex        =   22
         Top             =   240
         Width           =   4695
         Begin MSComCtl2.DTPicker dtpHFrom 
            Height          =   495
            Index           =   0
            Left            =   600
            TabIndex        =   9
            Top             =   360
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.25
         End
         Begin MSComCtl2.DTPicker dtpHTo 
            Height          =   495
            Index           =   0
            Left            =   2880
            TabIndex        =   10
            Top             =   360
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.5826388889
         End
         Begin MSComCtl2.DTPicker dtpHFrom 
            Height          =   495
            Index           =   1
            Left            =   600
            TabIndex        =   11
            Top             =   1080
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.25
         End
         Begin MSComCtl2.DTPicker dtpHTo 
            Height          =   495
            Index           =   1
            Left            =   2880
            TabIndex        =   12
            Top             =   1080
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.5826388889
         End
         Begin MSComCtl2.DTPicker dtpHFrom 
            Height          =   495
            Index           =   2
            Left            =   600
            TabIndex        =   13
            Top             =   1800
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.25
         End
         Begin MSComCtl2.DTPicker dtpHTo 
            Height          =   495
            Index           =   2
            Left            =   2880
            TabIndex        =   14
            Top             =   1800
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "hh:MM:ss"
            Format          =   17956866
            UpDown          =   -1  'True
            CurrentDate     =   38462.5826388889
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "§Õn:"
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
            Height          =   285
            Left            =   2400
            TabIndex        =   28
            Tag             =   "L7"
            Top             =   480
            Width           =   525
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            Caption         =   "Tõ :"
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
            Height          =   285
            Left            =   120
            TabIndex        =   27
            Tag             =   "L6"
            Top             =   480
            Width           =   405
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "§Õn:"
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
            Height          =   285
            Left            =   2400
            TabIndex        =   26
            Tag             =   "L7"
            Top             =   1200
            Width           =   525
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Tõ :"
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
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Tag             =   "L6"
            Top             =   1200
            Width           =   405
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            Caption         =   "§Õn:"
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
            Height          =   285
            Left            =   2400
            TabIndex        =   24
            Tag             =   "L7"
            Top             =   1920
            Width           =   525
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Tõ :"
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
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Tag             =   "L6"
            Top             =   1920
            Width           =   405
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Gi¸ giê tÝnh trªn 1 phót"
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
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "frmSetupKaraoke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsKaraoke As New ADODB.Recordset
Dim str As String

Private Sub cmdAdd_Click()
On Error GoTo Handle
If txtHoliday.Text = "" Then Exit Sub
    With lstHolidays
        .AddItem (txtHoliday.Text)
    End With
    txtHoliday.Text = ""
    txtHoliday.SetFocus
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdAdd_Click"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub AddCbo_Weekday(cbo As ComboBox)
On Error GoTo Handle
Dim i As Integer
Dim wkday As String
    With cbo
        .Clear
        For i = 1 To 7
            wkday = WeekdayName(i, False, vbSunday)
            .AddItem wkday
        Next i
        .ListIndex = 0
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " AddCbo_Weekday"
End Sub

Private Sub cmddelete_Click()
On Error GoTo Handle
    With lstHolidays
        .RemoveItem (.ListIndex)
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdDelete_Click"
End Sub

Private Sub cmdHelp_Click()
Dim strMessage As String
strMessage = "ThiÕt lËp gi¸ giê trªn 1 Phót theo c¸c ngµy ®­îc" & vbCrLf & _
             "chän trong Listdown, Cã 3 møc gi¸ theo giê vµ 3 møc" & vbCrLf & _
             "gi¸ theo khu vùc"

    MsgBox strMessage, vbInformation, "Trî gióp thiÕt lËp gi¸ giê"
End Sub

Private Sub cmdSave_Click()
On Error GoTo Handle
Dim reponse As Integer
    reponse = MsgBox("B¹n cã muèn l­u th«ng tin thay ®æi kh«ng?", vbYesNoCancel)
    Select Case reponse
        Case vbYes
            Call Update_DB
        Case vbNo
            Set rsKaraoke = Nothing
            Exit Sub
        Case vbCancel
            Exit Sub
    End Select
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    If cnData.State = 0 Then
        Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    End If
        Set rsKaraoke = Open_Table(cnData, "Setup_Karaoke")
    Call AddCbo_Weekday(cboStd_Weekday_From)
    Call AddCbo_Weekday(cboStd_Weekday_To)
    Call AddCbo_Weekday(cboH_Weekday_From)
    Call AddCbo_Weekday(cboH_Weekday_To)
    Call Initilize

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Private Sub txtHL_Price_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii < 32 And KeyAscii <> 13 Then Exit Sub
                Select Case KeyAscii
                    Case 48 To 57, 46
                    Case 13
                        txtHL_Price(Index).Text = Format(txtHL_Price(Index).Text, formatNum)
                        txtHL_Price(Index).SelStart = Len(txtHL_Price(Index))
                    Case Else:   KeyAscii = 0
                End Select
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "txtStd_Price_KeyPress"

End Sub

Private Sub txtHP_Price_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Handle
If KeyAscii < 32 And KeyAscii <> 13 Then Exit Sub
                Select Case KeyAscii
                    Case 48 To 57, 46
                    Case 13
                       txtHP_Price(Index).Text = Format(txtHP_Price(Index).Text, formatNum)
                        txtHP_Price(Index).SelStart = Len(txtHP_Price(Index))
                    Case Else:   KeyAscii = 0
                End Select
    
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "txtHL_Price_KeyPress"

End Sub

Private Sub txtStd_Price_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii < 32 And KeyAscii <> 13 Then Exit Sub
                Select Case KeyAscii
                    Case 48 To 57, 46
                    Case 13
                        txtStd_Price(Index).Text = Format(txtStd_Price(Index).Text, formatNum)
                        txtStd_Price(Index).SelStart = Len(txtStd_Price(Index))
                    Case Else:   KeyAscii = 0
                End Select
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "txtStd_Price_KeyPress"
End Sub

Public Sub Initilize()
On Error GoTo Handle
Dim i, j As Integer
    If rsKaraoke.State = 1 And rsKaraoke.RecordCount > 0 Then
        rsKaraoke.MoveFirst
    Else
        Exit Sub
    End If
    With rsKaraoke
        For i = 1 To 3
            .Find "ID='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                Select Case i
                    Case 1:
                        cboStd_Weekday_From.Text = .Fields("Weekday_from")
                        cboStd_Weekday_To.Text = .Fields("Weekday_To")
                            txtStd_Price(0).Text = .Fields("Price1")
                            txtStd_Price(1).Text = .Fields("Price2")
                            txtStd_Price(2).Text = .Fields("Price3")
                            txtStd_Price(3).Text = .Fields("Price11")
                            txtStd_Price(4).Text = .Fields("Price21")
                            txtStd_Price(5).Text = .Fields("Price31")
                            txtStd_Price(6).Text = .Fields("Price12")
                            txtStd_Price(7).Text = .Fields("Price22")
                            txtStd_Price(8).Text = .Fields("Price32")
                            txtStd_Price(9).Text = .Fields("Price13")
                            txtStd_Price(10).Text = .Fields("Price23")
                            txtStd_Price(11).Text = .Fields("Price33")
                            dtpStdFrom(0).Value = .Fields("From_Time1")
                            dtpStdTo(0).Value = .Fields("To_Time1")
                            dtpStdFrom(1).Value = .Fields("From_Time2")
                            dtpStdTo(1).Value = .Fields("To_Time2")
                            dtpStdFrom(2).Value = .Fields("From_Time3")
                            dtpStdTo(2).Value = .Fields("To_Time3")
                           
                    Case 2:
                        cboH_Weekday_From.Text = .Fields("Weekday_from")
                        cboH_Weekday_To.Text = .Fields("Weekday_To")
                            txtHP_Price(0).Text = .Fields("Price1")
                            txtHP_Price(1).Text = .Fields("Price2")
                            txtHP_Price(2).Text = .Fields("Price3")
                            txtHP_Price(3).Text = .Fields("Price11")
                            txtHP_Price(4).Text = .Fields("Price21")
                            txtHP_Price(5).Text = .Fields("Price31")
                            txtHP_Price(6).Text = .Fields("Price12")
                            txtHP_Price(7).Text = .Fields("Price22")
                            txtHP_Price(8).Text = .Fields("Price32")
                            txtHP_Price(9).Text = .Fields("Price13")
                            txtHP_Price(10).Text = .Fields("Price23")
                            txtHP_Price(11).Text = .Fields("Price33")
                            dtpHFrom(0).Value = .Fields("From_Time1")
                            dtpHTo(0).Value = .Fields("To_Time1")
                            dtpHFrom(1).Value = .Fields("From_Time2")
                            dtpHTo(1).Value = .Fields("To_Time2")
                            dtpHFrom(2).Value = .Fields("From_Time3")
                            dtpHTo(2).Value = .Fields("To_Time3")
                    Case 3:
                            txtHL_Price(0).Text = .Fields("Price1")
                            txtHL_Price(1).Text = .Fields("Price2")
                            txtHL_Price(2).Text = .Fields("Price3")
                            txtHL_Price(3).Text = .Fields("Price11")
                            txtHL_Price(4).Text = .Fields("Price21")
                            txtHL_Price(5).Text = .Fields("Price31")
                            txtHL_Price(6).Text = .Fields("Price12")
                            txtHL_Price(7).Text = .Fields("Price22")
                            txtHL_Price(8).Text = .Fields("Price32")
                            txtHL_Price(9).Text = .Fields("Price13")
                            txtHL_Price(10).Text = .Fields("Price23")
                            txtHL_Price(11).Text = .Fields("Price33")
                            dtpHLFrom(0).Value = .Fields("From_Time1")
                            dtpHLTo(0).Value = .Fields("To_Time1")
                            dtpHLFrom(1).Value = .Fields("From_Time2")
                            dtpHLTo(1).Value = .Fields("To_Time2")
                            dtpHLFrom(2).Value = .Fields("From_Time3")
                            dtpHLTo(2).Value = .Fields("To_Time3")
                            Call AddListBox(.Fields("Weekday_from"))
                End Select
            End If
        Next
    End With
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Public Sub AddListBox(str As String)
Dim plash As Integer
Dim count As Integer
Dim chuoi, tmpStr As String
Dim mang() As String
    chuoi = str
    count = 0
    Do While Len(chuoi) > 0
        plash = InStr(1, chuoi, ";", 0)
        ReDim Preserve mang(count)
        If plash <> 0 Then
            tmpStr = Mid(chuoi, 1, plash - 1)
            mang(count) = tmpStr
            If Len(chuoi) - Len(tmpStr) - 1 > 0 Then
                chuoi = Mid(chuoi, plash + 1, Len(chuoi) - Len(tmpStr) - 1)
            Else
                chuoi = ""
            End If
        Else
            mang(count) = tmpStr
            Exit Do
        End If
        count = count + 1
    Loop
        lstHolidays.Clear
    For count = 0 To UBound(mang())
        With lstHolidays
            .AddItem mang(count)
        End With
    Next
End Sub

Public Sub Update_DB()
On Error GoTo Handle
Dim i As Integer
    With rsKaraoke
        For i = 1 To 3
            .Find "ID='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                Select Case i
                    Case 1:
                        .Fields("Weekday_from") = cboStd_Weekday_From.Text
                        .Fields("Weekday_To") = cboStd_Weekday_To.Text
                        .Fields("Price1") = txtStd_Price(0).Text
                        .Fields("Price2") = txtStd_Price(1).Text
                        .Fields("Price3") = txtStd_Price(2).Text
                        .Fields("Price11") = txtStd_Price(3).Text
                        .Fields("Price21") = txtStd_Price(4).Text
                        .Fields("Price31") = txtStd_Price(5).Text
                        .Fields("Price12") = txtStd_Price(6).Text
                        .Fields("Price22") = txtStd_Price(7).Text
                        .Fields("Price32") = txtStd_Price(8).Text
                        .Fields("Price13") = txtStd_Price(9).Text
                        .Fields("Price23") = txtStd_Price(10).Text
                        .Fields("Price33") = txtStd_Price(11).Text
                        .Fields("From_Time1") = dtpStdFrom(0).Value
                        .Fields("To_Time1") = dtpStdTo(0).Value
                        .Fields("From_Time2") = dtpStdFrom(1).Value
                        .Fields("To_Time2") = dtpStdTo(1).Value
                        .Fields("From_Time3") = dtpStdFrom(2).Value
                        .Fields("To_Time3") = dtpStdTo(2).Value
                        .Update
                    Case 2:
                        .Fields("Weekday_from") = cboH_Weekday_From.Text
                        .Fields("Weekday_To") = cboH_Weekday_To.Text
                        .Fields("Price1") = txtHP_Price(0).Text
                        .Fields("Price2") = txtHP_Price(1).Text
                        .Fields("Price3") = txtHP_Price(2).Text
                        .Fields("Price11") = txtHP_Price(3).Text
                        .Fields("Price21") = txtHP_Price(4).Text
                        .Fields("Price31") = txtHP_Price(5).Text
                        .Fields("Price12") = txtHP_Price(6).Text
                        .Fields("Price22") = txtHP_Price(7).Text
                        .Fields("Price32") = txtHP_Price(8).Text
                        .Fields("Price13") = txtHP_Price(9).Text
                        .Fields("Price23") = txtHP_Price(10).Text
                        .Fields("Price33") = txtHP_Price(11).Text
                        .Fields("From_Time1") = dtpHFrom(0).Value
                        .Fields("To_Time1") = dtpHTo(0).Value
                        .Fields("From_Time2") = dtpHFrom(1).Value
                        .Fields("To_Time2") = dtpHTo(1).Value
                        .Fields("From_Time3") = dtpHFrom(2).Value
                        .Fields("To_Time3") = dtpHTo(2).Value
                        .Update
                    Case 3:
                        .Fields("Price1") = txtHL_Price(0).Text
                        .Fields("Price2") = txtHL_Price(1).Text
                        .Fields("Price3") = txtHL_Price(2).Text
                        .Fields("Price11") = txtHL_Price(3).Text
                        .Fields("Price21") = txtHL_Price(4).Text
                        .Fields("Price31") = txtHL_Price(5).Text
                        .Fields("Price12") = txtHL_Price(6).Text
                        .Fields("Price22") = txtHL_Price(7).Text
                        .Fields("Price32") = txtHL_Price(8).Text
                        .Fields("Price13") = txtHL_Price(9).Text
                        .Fields("Price23") = txtHL_Price(10).Text
                        .Fields("Price33") = txtHL_Price(11).Text
                        .Fields("From_Time1") = dtpHLFrom(0).Value
                        .Fields("To_Time1") = dtpHLTo(0).Value
                        .Fields("From_Time2") = dtpHLFrom(1).Value
                        .Fields("To_Time2") = dtpHLTo(1).Value
                        .Fields("From_Time3") = dtpHLFrom(2).Value
                        .Fields("To_Time3") = dtpHLTo(2).Value
                        .Fields("Weekday_from") = gfList_value_String
                        .Update
                        
                End Select
            End If
        Next
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Update_DB"
End Sub

Public Function gfList_value_String() As String
On Error GoTo Handle
Dim str As String
Dim i As Integer
str = ""
    With lstHolidays
        For i = 0 To .ListCount
            If str = "" Then
                str = .List(i)
            Else
                str = str & ";" & .List(i)
            End If
        Next
    End With
    gfList_value_String = str
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & "gfList_value_String"
End Function

