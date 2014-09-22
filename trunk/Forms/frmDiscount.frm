VERSION 5.00
Begin VB.Form frmDiscount 
   Caption         =   "Gi¶m gi¸"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12825
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
   ScaleHeight     =   8445
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fra 
      Height          =   8295
      Left            =   5400
      TabIndex        =   10
      Top             =   0
      Width           =   7335
      Begin prjTouchScreen.MyButton cmdClose 
         Cancel          =   -1  'True
         Height          =   1095
         Left            =   4200
         TabIndex        =   11
         Top             =   7080
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1931
         BTYPE           =   3
         TX              =   "§ãng"
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
         BCOL            =   255
         BCOLO           =   33023
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmDiscount.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdDiscount 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1720
         BTYPE           =   3
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
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8438015
         BCOLO           =   12648384
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmDiscount.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdOK 
         Height          =   1095
         Left            =   240
         TabIndex        =   23
         Top             =   7080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1931
         BTYPE           =   3
         TX              =   "L­u"
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
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmDiscount.frx":0038
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
   Begin VB.Frame fraSection 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   5175
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   4580
         Left            =   0
         ScaleHeight     =   4575
         ScaleWidth      =   5295
         TabIndex        =   8
         Top             =   120
         Width           =   5300
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   975
            Index           =   0
            Left            =   0
            TabIndex        =   14
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            BTYPE           =   3
            TX              =   "1"
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
            BCOL            =   16777215
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDiscount.frx":0054
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   705
            Left            =   15
            TabIndex        =   9
            Top             =   0
            Width           =   3220
         End
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   975
            Index           =   1
            Left            =   1080
            TabIndex        =   15
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            BTYPE           =   3
            TX              =   "2"
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
            BCOL            =   16777215
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDiscount.frx":0070
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   975
            Index           =   2
            Left            =   2160
            TabIndex        =   16
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            BTYPE           =   3
            TX              =   "3"
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
            BCOL            =   16777215
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDiscount.frx":008C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   975
            Index           =   3
            Left            =   0
            TabIndex        =   17
            Top             =   1680
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            BTYPE           =   3
            TX              =   "4"
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
            BCOL            =   16777215
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDiscount.frx":00A8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   975
            Index           =   4
            Left            =   1080
            TabIndex        =   18
            Top             =   1680
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            BTYPE           =   3
            TX              =   "5"
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
            BCOL            =   16777215
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDiscount.frx":00C4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   975
            Index           =   5
            Left            =   2160
            TabIndex        =   19
            Top             =   1680
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            BTYPE           =   3
            TX              =   "6"
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
            BCOL            =   16777215
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDiscount.frx":00E0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   975
            Index           =   6
            Left            =   0
            TabIndex        =   20
            Top             =   2640
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            BTYPE           =   3
            TX              =   "7"
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
            BCOL            =   16777215
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDiscount.frx":00FC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   975
            Index           =   7
            Left            =   1080
            TabIndex        =   21
            Top             =   2640
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            BTYPE           =   3
            TX              =   "8"
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
            BCOL            =   16777215
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDiscount.frx":0118
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   975
            Index           =   8
            Left            =   2160
            TabIndex        =   22
            Top             =   2640
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            BTYPE           =   3
            TX              =   "9"
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
            BCOL            =   16777215
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDiscount.frx":0134
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   975
            Index           =   9
            Left            =   0
            TabIndex        =   24
            Top             =   3600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            BTYPE           =   3
            TX              =   "0"
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
            BCOL            =   16777215
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDiscount.frx":0150
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   975
            Index           =   10
            Left            =   1080
            TabIndex        =   25
            Top             =   3600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            BTYPE           =   3
            TX              =   "00"
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
            BCOL            =   16777215
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDiscount.frx":016C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   975
            Index           =   11
            Left            =   2160
            TabIndex        =   26
            Top             =   3600
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   1720
            BTYPE           =   3
            TX              =   "."
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
            BCOL            =   16777215
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmDiscount.frx":0188
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   735
            Index           =   12
            Left            =   3240
            TabIndex        =   27
            Top             =   0
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1296
            BTYPE           =   3
            TX              =   "Backspace"
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
            BCOL            =   16777215
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "frmDiscount.frx":01A4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   1935
            Index           =   13
            Left            =   3240
            TabIndex        =   28
            Top             =   720
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   3413
            BTYPE           =   3
            TX              =   "Clear"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
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
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "frmDiscount.frx":01C0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdAlpha 
            Height          =   1935
            Index           =   14
            Left            =   3240
            TabIndex        =   29
            Top             =   2640
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   3413
            BTYPE           =   3
            TX              =   "Enter"
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
            BCOL            =   16777215
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   16777215
            MPTR            =   1
            MICON           =   "frmDiscount.frx":01DC
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
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   5640
         Width           =   2175
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   4800
         Width           =   3135
      End
      Begin VB.Label lblCash 
         Alignment       =   1  'Right Justify
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   735
         Left            =   1800
         TabIndex        =   7
         Top             =   6600
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "Cßn l¹i:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   6720
         Width           =   2055
      End
      Begin VB.Label lblDiscount 
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   5
         Top             =   5760
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   5400
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Label Label2 
         Caption         =   "Gi¶m:"
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
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   5760
         Width           =   1935
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Tæng céng:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   4920
         Width           =   1575
      End
   End
   Begin VB.Label Label4 
      Caption         =   "Chän h×nh thøc khuyÕn m·i hoÆc nhËp % cÇn gi¶m"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmDiscount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Total As Double
Dim Pro_Value As Integer
Dim rsPromotion As New ADODB.Recordset
Dim isOK As Boolean
Dim Other As String
Dim reason_discount As String


Private Sub cmdAlpha_Click(Index As Integer)
On Error GoTo Handle
    Select Case Index
        Case 0 To 11:
                txtQty.Text = txtQty.Text & cmdAlpha(Index).Caption
                Other = 0
        Case 13
            txtQty.Text = ""
        Case 14
            If CDbl("0" & txtQty.Text) > 100 Then
                MsgBox "Kh«ng thÓ gi¶m gi¸ h¬n 100%"
            Else
                lblDiscount.Caption = txtQty.Text
            End If
           txtQty.Text = ""
        Case 12
            If Len(txtQty) > 0 Then
              txtQty.Text = Left(txtQty, Len(txtQty) - 1)
            End If
            Other = 0
    End Select
Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdAlpha_Click" & vbCrLf
    MsgBox Err.Number & Err.Description & Me.name & " cmdAlpha_Click"
End Sub

Private Sub cmdAlpha_MouseOut(Index As Integer)
On Error GoTo Handle
    With cmdAlpha(Index)
        If Index = 12 Then
            .Font.Size = 14
        Else
            .Font.Size = 24
        End If
        .FontItalic = False
        .ForeColor = vbBlue
        .SpecialEffect = cbNone
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdAlpha_MouseOver"
End Sub

Private Sub cmdAlpha_MouseOver(Index As Integer)
On Error GoTo Handle
    With cmdAlpha(Index)
        .Font.Size = 32
        .FontItalic = True
        .ForeColor = vbRed
        .SpecialEffect = cbShadowed
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdAlpha_MouseOver"
End Sub

'Private Sub cmdAlpha_Click(Index As Integer)
' Error GoTo Handle
'    Select Case Index
'        Case 0 To 11:
'                txtQty.Text = txtQty.Text & cmdAlpha(Index).Caption
'                Other = 0
'        Case 13
'            txtQty.Text = ""
'        Case 14
'           lblDiscount.Caption = txtQty.Text
'           txtQty.Text = ""
'        Case 12
'            If Len(txtQty) > 0 Then
'              txtQty.Text = Left(txtQty, Len(txtQty) - 1)
'            End If
'            Other = 0
'    End Select
'Exit Sub
'Handle:
'    'Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdAlpha_Click" & vbCrLf
'    MsgBox Err.Number & Err.Description & Me.Name & " cmdAlpha_Click"
'End Sub

Private Sub cmdClose_Click()
    isOK = False
    Unload Me
End Sub

Public Sub LoadCommand(rs As ADODB.Recordset, strTenfield1 As String)
On Error Resume Next 'GoTo Handle
Dim Index As Integer
Dim sodong As Integer
Dim i, j As Integer
Index = 1
rs.MoveFirst
If rs.RecordCount Mod 2 > 0 Then
    sodong = rs.RecordCount / 2 + 1
Else
    sodong = rs.RecordCount / 2
End If
If rs.RecordCount > 0 Then
For i = 1 To sodong
    For j = 1 To 3
            Load cmdDiscount(Index)
            With cmdDiscount(Index)
            If i = 1 Then
                If Index Mod 4 = 0 Then
                    .Left = Fra.Left + 100
                    .top = cmdDiscount(Index - 1).top + cmdDiscount(Index - 1).Height + 200
                Else
                    .top = cmdDiscount(Index - 1).top
                    If j = 1 Then
                        .Left = 200
                    Else
                        .Left = cmdDiscount(Index - 1).Left + cmdDiscount(Index - 1).Width + 100
                    End If
                End If
            Else
                If (Index - 1) Mod 3 = 0 Then
                    .Left = 200
                    .top = cmdDiscount(Index - 1).top + cmdDiscount(Index - 1).Height + 200
                Else
                    .top = cmdDiscount(Index - 1).top
                    If j = 1 Then
                       .Left = 200
                    Else
                        .Left = cmdDiscount(Index - 1).Left + cmdDiscount(Index - 1).Width + 100
                    End If
                End If
            End If
                If Not rs.EOF Then
                    .Caption = rs.Fields("" & strTenfield1 & "") '& Chr(13) & rs.Fields("" & strTenfield2 & "")
                    .Tag = rs.Fields("Pro_Value")
                    .ToolTipText = rs.Fields("Pro_ID")
                Else
                    Exit Sub
                End If
                .Visible = True
                .Height = 1000
                .Width = 2200
        
            End With
        rs.MoveNext
        Index = Index + 1
    Next j
Next i

End If
Exit Sub
'Handle:
'    MsgBox Err.Number & Err.Description & Me.Name & "  LoadCommandSub"
End Sub

Private Sub cmdDiscount_Click(Index As Integer)
    On Error GoTo Handle
        lblDiscount.Caption = cmdDiscount(Index).Tag
        Other = cmdDiscount(Index).ToolTipText
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "cmdDiscount_Click"
End Sub


Public Property Let Get_Total(ByVal vNewValue As Variant)
    Total = vNewValue
End Property

Public Property Get Let_Value() As Variant
    Let_Value = Pro_Value
End Property

Private Sub cmdOK_Click()
On Error GoTo Handle
Dim isReason As Boolean
If txtQty.Text <> "" Then cmdAlpha_Click (14)
    If lblDiscount.Caption = "" Or CDbl(lblDiscount.Caption) > 100 Then Exit Sub
    isOK = True
    With frmPro_Reason
        .Show vbModal
        reason_discount = .Let_Reason
        isReason = .Let_OK_Cancel
    End With
    
     If isOK = False Then
        Pro_Value = 0
    Else
        If isReason = True Then
            Pro_Value = CDbl("0" & lblDiscount.Caption)
        Else
            Pro_Value = 0
        End If
    End If
    Unload Me
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdOK_Click"
End Sub

Private Sub Form_Load()
 On Error GoTo Handle
    If cnData.State = 0 Then cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    Set rsPromotion = Open_Table(cnData, "Promotion")
    TxtTotal.Text = Format(Total, "#,##0")
    lblCash.Caption = Format(CDbl("0" & TxtTotal.Text) - CDbl("0" & txtDiscount.Text), "#,###")
    
    Call LoadCommand(rsPromotion, "Pro_Name")
    
    
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Form_Load"
End Sub

Private Sub lblDiscount_Change()
On Error GoTo Handle
        txtDiscount = Format(CDbl(TxtTotal.Text) * CDbl(lblDiscount.Caption) / 100, "#,###")
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "lblDiscount_Change"
End Sub

Private Sub txtDiscount_Change()
On Error GoTo Handle
        lblCash.Caption = Format(CDbl("0" & TxtTotal.Text) - CDbl("0" & txtDiscount.Text), "#,###")
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "txtDiscount_Change"
End Sub

Public Property Get Let_OK() As Variant
    Let_OK = isOK
End Property


Public Property Get Let_Discount_Status() As Variant
    Let_Discount_Status = Other
End Property


Public Property Get Let_Reason_Discount() As Variant
    Let_Reason_Discount = reason_discount
End Property

