VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTablePlan 
   BackColor       =   &H00FFC0C0&
   Caption         =   "SO DO BAN"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   135
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
   Icon            =   "frmTablePlan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   15240
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   14160
      Top             =   10320
   End
   Begin VB.Timer BackupTimer 
      Interval        =   60000
      Left            =   8760
      Top             =   480
   End
   Begin VB.Frame fraTable 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8925
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   15255
      Begin VB.Frame fraSyn 
         BorderStyle     =   0  'None
         ForeColor       =   &H00FF0000&
         Height          =   4215
         Left            =   4200
         TabIndex        =   20
         Top             =   2160
         Visible         =   0   'False
         Width           =   7455
         Begin VB.Frame Frame2 
            Height          =   2175
            Left            =   0
            TabIndex        =   24
            Top             =   720
            Width           =   7455
            Begin VB.CheckBox chkCustomer 
               Caption         =   "Danh môc kh¸ch hµng"
               Height          =   375
               Left            =   4080
               TabIndex        =   32
               Top             =   1680
               Width           =   2535
            End
            Begin VB.CheckBox chkThuchu 
               Caption         =   "Thu Chi"
               Height          =   375
               Left            =   600
               TabIndex        =   31
               Top             =   1680
               Width           =   2895
            End
            Begin VB.CheckBox chkDMThuchi 
               Caption         =   "Danh môc kho¶n Thu, Chi"
               Height          =   375
               Left            =   600
               TabIndex        =   30
               Top             =   1200
               Width           =   3015
            End
            Begin VB.CheckBox chkVendor 
               Caption         =   "Danh môc nhµ cung cÊp"
               Height          =   375
               Left            =   4080
               TabIndex        =   29
               Top             =   1200
               Width           =   2535
            End
            Begin VB.CheckBox chkItems 
               Caption         =   "Danh môc hµng"
               Height          =   375
               Left            =   4080
               TabIndex        =   28
               Top             =   720
               Width           =   2415
            End
            Begin VB.CheckBox chkGroup 
               Caption         =   "Danh môc Nhãm hµng"
               Height          =   375
               Left            =   4080
               TabIndex        =   27
               Top             =   240
               Width           =   2415
            End
            Begin VB.CheckBox chkTablePlan 
               Caption         =   "S¬ ®å bµn"
               Height          =   375
               Left            =   600
               TabIndex        =   26
               Top             =   720
               Width           =   2895
            End
            Begin VB.CheckBox ChkSale 
               Caption         =   "D÷ liÖu b¸n hµng"
               Height          =   375
               Left            =   600
               TabIndex        =   25
               Top             =   240
               Width           =   2895
            End
         End
         Begin prjTouchScreen.MyButton cmdCancal 
            Height          =   735
            Left            =   3720
            TabIndex        =   23
            Top             =   3360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   1296
            BTYPE           =   3
            TX              =   "&Tho¸t"
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
            BCOL            =   16761024
            BCOLO           =   33023
            FCOL            =   16711680
            FCOLO           =   16777215
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmTablePlan.frx":111EA
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
            Height          =   735
            Left            =   1440
            TabIndex        =   22
            Top             =   3360
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   1296
            BTYPE           =   3
            TX              =   "§ån&g ý"
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
            BCOL            =   16761024
            BCOLO           =   33023
            FCOL            =   16711680
            FCOLO           =   16777215
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmTablePlan.frx":11206
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "®ång bé:"
            BeginProperty Font 
               Name            =   ".VnArialH"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   0
            TabIndex        =   35
            Top             =   200
            Width           =   3255
         End
         Begin VB.Label lblUnCheckAll 
            Caption         =   "Bá chän tÊt c¶"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4080
            TabIndex        =   34
            Top             =   3000
            Width           =   2175
         End
         Begin VB.Label lblCheckAll 
            Caption         =   "Chän tÊt c¶"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1440
            TabIndex        =   33
            Top             =   3000
            Width           =   1935
         End
         Begin VB.Label lblTitle 
            Alignment       =   2  'Center
            Caption         =   "..."
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   3360
            TabIndex        =   21
            Top             =   240
            Width           =   4095
         End
      End
      Begin VB.PictureBox picWait 
         BackColor       =   &H00C00000&
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4920
         ScaleHeight     =   360
         ScaleWidth      =   4905
         TabIndex        =   18
         Top             =   4440
         Visible         =   0   'False
         Width           =   4965
         Begin MSComctlLib.ProgressBar probarWait 
            Height          =   390
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   4905
            _ExtentX        =   8652
            _ExtentY        =   688
            _Version        =   393216
            Appearance      =   1
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   120
         Top             =   1920
      End
      Begin VB.Frame fraTakeOut 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3285
         Left            =   9000
         TabIndex        =   0
         Top             =   5520
         Visible         =   0   'False
         Width           =   4455
         Begin VB.Timer TextFly 
            Interval        =   1000
            Left            =   3600
            Top             =   2040
         End
         Begin prjTouchScreen.MyButton cmdNewTable 
            Height          =   915
            Left            =   270
            TabIndex        =   3
            Top             =   240
            Width           =   2805
            _ExtentX        =   4948
            _ExtentY        =   1614
            BTYPE           =   5
            TX              =   "Th«ng tin bµn míi"
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
            BCOL            =   12648447
            BCOLO           =   12648447
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmTablePlan.frx":11222
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   2
            DrawMode        =   7  'Invert
            X1              =   60
            X2              =   15345
            Y1              =   1380
            Y2              =   1395
         End
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label lblSync 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Caption         =   "§ang ®ång bé....."
         BeginProperty Font 
            Name            =   ".VnArial NarrowH"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   6600
         TabIndex        =   12
         Top             =   4800
         Visible         =   0   'False
         Width           =   4935
      End
      Begin VB.Label lblTable 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "#1"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   1635
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Height          =   1035
         Index           =   0
         Left            =   120
         Top             =   210
         Visible         =   0   'False
         Width           =   1755
      End
   End
   Begin prjTouchScreen.MyButton cmdExittoLogin 
      Height          =   1020
      Left            =   -90
      TabIndex        =   1
      Tag             =   "L2"
      Top             =   -15
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1799
      BTYPE           =   6
      TX              =   "&Tho¸t ca"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632064
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmTablePlan.frx":1123E
      UMCOL           =   0   'False
      SOFT            =   -1  'True
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdEdit 
      Height          =   1020
      Left            =   1340
      TabIndex        =   2
      Tag             =   "L3"
      Top             =   -15
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1799
      BTYPE           =   6
      TX              =   "ThiÕt lËp s¬ ®å bµn"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632064
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmTablePlan.frx":1125A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdOption 
      Height          =   1020
      Left            =   2780
      TabIndex        =   5
      Tag             =   "L4"
      Top             =   -15
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1799
      BTYPE           =   6
      TX              =   "&ChØnh söa danh môc"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632064
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmTablePlan.frx":11276
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdTReserved 
      Height          =   1020
      Left            =   4210
      TabIndex        =   6
      Tag             =   "L5"
      Top             =   -15
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1799
      BTYPE           =   6
      TX              =   "&§Æt bµn"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632064
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmTablePlan.frx":11292
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdDeliver 
      Height          =   1020
      Left            =   5660
      TabIndex        =   7
      Top             =   -15
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1799
      BTYPE           =   6
      TX              =   "Xö lý bµn lçi"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632064
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmTablePlan.frx":112AE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdTableOption 
      Height          =   225
      Left            =   9600
      TabIndex        =   8
      Tag             =   "L7"
      Top             =   840
      Visible         =   0   'False
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   397
      BTYPE           =   6
      TX              =   "  &Bµn míi  ph¸t sinh"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12640511
      BCOLO           =   33023
      FCOL            =   12582912
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmTablePlan.frx":112CA
      UMCOL           =   0   'False
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   -1  'True
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdSection 
      Height          =   1005
      Index           =   0
      Left            =   0
      TabIndex        =   9
      Top             =   10080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1773
      BTYPE           =   6
      TX              =   "Section"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArialH"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTablePlan.frx":112E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdSynchronize 
      Height          =   1020
      Left            =   7100
      TabIndex        =   11
      Top             =   -15
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   1799
      BTYPE           =   6
      TX              =   "§ång bé d÷ liÖu"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632064
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmTablePlan.frx":11302
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblSection 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   585
      Left            =   8880
      TabIndex        =   17
      Top             =   360
      Width           =   2145
   End
   Begin VB.Label lblQuay 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "QuÇy sè:1"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   8880
      TabIndex        =   16
      Tag             =   "L8"
      Top             =   0
      Width           =   2145
   End
   Begin VB.Label lblPhone 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ÑT:(8) - 8867.869 - 0918.655.887 "
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11160
      TabIndex        =   15
      Top             =   735
      Width           =   4215
   End
   Begin VB.Label lblAdd 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "565/6 Bình Thôùi, P.10, Q.11, Tp.HCM"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11160
      TabIndex        =   14
      Top             =   405
      Width           =   4215
   End
   Begin VB.Label lblCompanyname 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "phuùc thaïnh vinh"
      BeginProperty Font 
         Name            =   "VNI-Algerian"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   11160
      TabIndex        =   13
      Top             =   0
      Width           =   4215
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "ChØnh söa danh môc"
      Visible         =   0   'False
      Begin VB.Menu mnuEditTable 
         Caption         =   "S¬ ®å bµn"
      End
      Begin VB.Menu mnuGroup 
         Caption         =   "Nhãm hµng"
      End
      Begin VB.Menu mnuitem 
         Caption         =   "M· hµng "
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Lùa chän m¸y in"
      End
      Begin VB.Menu mnuReport 
         Caption         =   "B¸o c¸o b¸n hµng"
      End
      Begin VB.Menu mnuLocation 
         Caption         =   "Khu vùc"
      End
      Begin VB.Menu mnuprinterName 
         Caption         =   "Tªn m¸y in"
      End
   End
End
Attribute VB_Name = "frmTablePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Desarr() As String
Dim Drag As Boolean
Dim rsSection As New ADODB.Recordset
'Dim Sec_ID As String
Dim CountTable As Integer
Dim CountSection As Integer
Dim rsTable As New ADODB.Recordset
Dim iLoad As Boolean
Dim iLoadSection As Boolean
Dim EventCall As String
Dim StateCall As Integer
Dim TranferTable As String
Dim LocationTranfer As String
Dim BillTranfer As Double
Dim TimeSync As Double
Dim Max_Invoice_Backup, Discount As Integer
Dim isclick As Boolean
Dim Table_ID As String
Dim AmountBackup As Double
Dim arrPriterKP() As String
Dim Countdown As Integer
Dim TimerBackup As Integer
Dim Option_call As Integer

Private Sub BackupTimer_Timer()
    TimerBackup = TimerBackup + 1
    If TimerBackup = 10 Then
        Call Backup_DB
        TimerBackup = 0
    End If
End Sub

Private Sub cmdCancal_Click()
    fraSyn.Visible = False
End Sub

Private Sub cmdDeliver_Click()
On Error GoTo Handle
Dim strSQL As String
    strSQL = "delete * from Invoice_OnHold where Invoice_Number not in (Select Invoice_Number from Invoice_Totals)"
    cnData.Execute strSQL
    cnData.Execute "Update Invoice_Totals set InvoiceNotesUsed =false"
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "   cmdTakeOut_Click"
End Sub

Private Sub cmdEdit_Click()
    'Keyboard = "EditTable"
    With frmPassword
        .FormActionKey = "EditTable"
        .Show vbModal
    End With
End Sub

Private Sub cmdExittoLogin_Click()
    Set rsSection = Nothing
    Set rsTable = Nothing
    Set rsTranfer = Nothing
    Set cnData = Nothing
'    Call gsDELETE_TMP_FILE
    Unload Me
    
    With frmLogin
        .Me_State = 1
        .Show vbModal
        
    End With
End Sub

Public Sub Backup_DB()
On Error GoTo Handle
    Dim fso As New FileSystemObject
    fso.CopyFile WorkingFolder & "\Database.mdb", WorkingFolder & "\Backup", True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " Backup_DB"
End Sub


Private Sub cmdNewTable_Click()
Select Case EventCall
    Case "TakeOut"
        With frmKeyboard
            .txtInput.MaxLength = 10
            .FormCallkeyboard = "TakeOut"
            .txtInput.PasswordChar = ""
            .Show vbModal
        End With
    Case "Delivery"
        With frmFindCustomer
            .FormCall = "Delivery"
            .Show vbModal
        End With
    Case "Banphatsinh"
        With frmKeyboard
            .txtInput.MaxLength = 4
            .FormCallkeyboard = "Banphatsinh"
            .txtInput.PasswordChar = ""
            .Show vbModal
        End With
End Select
End Sub

Private Sub cmdOK_Click()
On Error GoTo Handle
Dim cnBackup As New ADODB.Connection
Dim cnOrg As New ADODB.Connection

If Dir(WorkingFolder & "\Database.mdb", vbDirectory) <> "" Then
    Set cnOrg = Get_Connection(WorkingFolder & "\Database.mdb", "100881")
End If

If Dir(BackupFolder & "\Database.mdb", vbDirectory) <> "" Then
    Set cnBackup = Get_Connection(BackupFolder & "\Database.mdb", "100881")
End If
    
    'Dong bo so do ban
    If chkTablePlan.Value = 1 Then Call gfBackup_TablePlan(cnOrg, cnBackup)
    
    'Dong bo nhom hang
    If chkGroup.Value = 1 Then Call gfBackup_Group(cnOrg, cnBackup)
    
    'Dong bo Danh sach hang
    If chkItems.Value = 1 Then Call gfBackup_Items(cnOrg, cnBackup)
    
    'Dong bo du lieu ban hang
    If ChkSale.Value = 1 Then
        If gfSynchronizeData = False Then MsgBox "§· ®ång bé xong d÷ liÖu b¸n hµng"
    End If
        
    'Dong bo Khach hang
    If chkCustomer.Value = 1 Then Call gfBackup_Customer(cnOrg, cnBackup)
    
    'Dong bo Nha cung cap
    If chkVendor.Value = 1 Then Call gfBackup_Vendor(cnOrg, cnBackup)
    
    'Dong bo Danh muc thu chi
    If chkDMThuchi.Value = 1 Then Call gfBackup_DMInOut(cnOrg, cnBackup)
'
'    'Dong bo Du lieu thu chi
    If chkThuchu.Value = 1 Then Call gfBackup_InOut(cnOrg, cnBackup)
    
    lblTitle.Caption = "§· hoµn tÊt ®ång bé d÷ liÖu !"
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " cmdOK_Click"
End Sub

Private Sub cmdOption_Click()
On Error GoTo Handle
    frmSetup.Show vbModal
Exit Sub
Handle:
MsgBox Err.Number & Err.Description
End Sub

Private Sub cmdPrint_Buffer_Click()
    StateCall = 5
End Sub

Private Sub cmdSynchronize_Click()
On Error GoTo Handle
If MsgBox("B¹n cã ch¾c ch¾n kÕt thóc ngµy, d÷ liÖu b¸n hµng sÏ bÞ xãa bá toµn bé ?", vbYesNo) = vbYes Then
    If Dir(BackupFolder & "\Database.mdb", vbDirectory) <> "" Then
        fraSyn.Visible = True
        Delay (50)
    Else
        MsgBox "Kh«ng thÓ xãa d÷ liÖu"
        
    End If
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " cmdSynchronize_Click"
End Sub

Private Sub cmdTableOption_Click()
MsgBox "Kh«ng sö dông"
Exit Sub
'Unload Me
'On Error GoTo Handle
'    EventCall = "Banphatsinh"
'    With fraTakeOut
'        .top = fraTable.top - 1200
'        .Left = fraTable.Left
'        .Visible = True
'    End With
'    lblSection.Caption = cmdTableOption.Caption
'Exit Sub
'Handle:
'MsgBox Err.Number & Err.Description & Me.Name & "   cmdTakeOut_Click"
End Sub

Private Sub cmdTReserved_Click()
On Error GoTo Handle
    If Not Check_Table_exist("Table_Reservered") Then
        Call Create_Table_Reserverd
    End If
    frmReservered.Show vbModal
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "   cmdTakeOut_Click"
End Sub

'Private Sub Command1_Click()
'    frmKitchenView.Show vbModal
'End Sub

Private Sub Form_Activate()
    On Error GoTo Handle
    Dim ctrl As Control
'    If cmdEdit.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
        Desarr = LoadLanguage(LngFile, "#01:004:")
        Me.Caption = Desarr(1)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Desarr(Mid(ctrl.Tag, 2))
        Next ctrl
    If Dir(WorkingFolder & "\Database.mdb", vbDirectory) <> "" Then
        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881")
    Else
        MsgBox "Vui lßng chän ®­êng dÉn d÷ liÖu ch­¬ng tr×nh !"
        End
    End If
    isclick = False
    'Init the Location
    
    Call Load_Section
    'Init Table on the first Location
    If Sec_ID <> "" And Sec_ID <> "TO" And Sec_ID <> "DE" And Sec_ID <> "AR" Then
        Sleep (500)
        Call LoadTable(Sec_ID)
        'cmdSection(CDbl(Sec_ID)).BackColor = vbGreen
        lblSection.Caption = cmdSection(1).Caption
    Else
        Sec_ID = Get_First_Section
        Call LoadTable(Sec_ID)
        'cmdSection(1).BackColor = vbGreen
'        lblSection.Caption = cmdSection(Sec_ID).Caption
    End If
    ' Gan font mac dinh cho nhan cong ty
        lblCompanyname.Font.Name = "VNI-Algerian"
        lblAdd.Font.Name = "VNI-Times"
        lblPhone.Font.Name = "VNI-Times"
        'If UserLevel <> 1 Then cmdSynchronize.Enabled = False
    fraTakeOut.Visible = False
    If UserLevel <> 1 Then CheckRight
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Handle
        If KeyCode = vbKeyF1 Then
            frmAboutInfor.Show vbModal
'        ElseIf KeyCode = vbKeyG Then
'            Call mnuGroup_Click
        End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " Form_KeyDown"
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    Dim ctrl As Control
    Desarr = LoadLanguage(LngFile, "#01:004:")
    Me.Caption = Desarr(1)
'    sText_org = " Cty TNHH TM & DV Phóc Th¹nh Vinh - Gi¶i ph¸p b¸n hµng chuyªn nghiÖp" & _
                " cho Nhµ hµng,Cafe,Bar,Karaoke...Hç trî KT 24/7:0918.655.887-0903.613.673-0909.118.669"
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Desarr(Mid(ctrl.Tag, 2))
    Next ctrl
    
    If Dir(WorkingFolder & "\Database.mdb", vbDirectory) <> "" Then
        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881")
    Else
        MsgBox "Vui lßng chän ®­êng dÉn d÷ liÖu cho ch­¬ng tr×nh"
        End
    End If
    isclick = False
    Discount = 0
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Form_Load"
End Sub

Private Sub cmdAddLocation_Click()
    With frmKeyboard
        .FormCallkeyboard = "Add_Section"
        .lblTitle.Caption = "Enter_Section"
        .txtInput.PasswordChar = ""
        .Show vbModal
    End With
    Call Load_Section
End Sub

Private Sub cmdDeleteLocation_Click()
    On Error GoTo Handle
        cnData.Execute "Delete * from Table_Diagram_Sections where Location_ID='" & Sec_ID & "'"
    
    Call Load_Section
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "   cmdDeleteLocation_Click "
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub cmdSection_Click(Index As Integer)
    On Error GoTo Handle
    Dim ctrl As Control
        Sec_ID = Format(cmdSection(Index).Tag, "00")
        Call LoadTable(CStr(Sec_ID))
        fraTable.Enabled = True
        iLoad = True
        lblSection.Caption = cmdSection(Index).Caption
    fraTakeOut.Visible = False
    For Each ctrl In Me
        If ctrl.Name = "cmdSection" Then
            ctrl.ForeColor = vbBlue
        End If
    Next ctrl
    cmdSection(Index).ForeColor = vbRed
    Exit Sub
    
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  cmdSection_Click "
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuEdit, 0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsSection = Nothing
    Set rsTable = Nothing
    CountTable = 0
    iLoad = False
    CountSection = 0
    iLoadSection = False
    EventCall = ""
    Discount = 0
End Sub

Private Sub fraTable_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = 2 Then
'        PopupMenu mnuEdit, 0
'    End If
End Sub

Private Sub lblCheckAll_Click()
    chkCustomer.Value = 1
    ChkSale.Value = 1
    chkDMThuchi.Value = 1
    chkGroup.Value = 1
    chkItems.Value = 1
    chkTablePlan.Value = 1
    chkThuchu.Value = 1
    chkVendor.Value = 1
    
End Sub

Private Sub lblTable_Click(Index As Integer)
On Error GoTo Handle
    Dim rsinvoice_hold As New ADODB.Recordset
    Dim rsInvoice_Total As New ADODB.Recordset
    Dim rsInvoice_Notes As New ADODB.Recordset
    Dim TimeIn_Kar As String
    Dim i As Integer
    Dim IsPrintTranfer As Boolean
    IsPrintTranfer = False
    If ArrayFlag(SF(4), 7) = 1 Then IsPrintTranfer = True
'    If isclick = True Then
'        MsgBox "Vui lßng kh«ng kÝch ®óp vµo sè bµn, BÊm OK ®Ó tiÕp tôc !C¶m ¬n"
'        Call Form_Activate
'        Exit Sub
'    End If
    isclick = True
    Discount = 0
    
    If cnData.State = 0 Then
        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881")
    End If
    If cnData.State <> 0 Then
        Set rsinvoice_hold = OpenCriticalTable("select * from Invoice_OnHold where Section_ID='" & Sec_ID & "'", cnData)
        Set rsInvoice_Total = OpenCriticalTable("select * from Invoice_Totals where Station_ID='" & Sec_ID & "'", cnData)
        Set rsInvoice_Notes = Open_Table(cnData, "Invoice_Totals_Notes")
    End If
    '=============================================================================
'    picWait.Visible = True
'    With probarWait
'        .Max = 6
'        If .Value < .Max Then
'            .Value = .Max
'        End If
'    End With
    Table_ID = Left(lblTable(Index).Caption, InStr(lblTable(Index).Caption, Chr(13)))
    Select Case StateCall
        Case 1  'Mo ban binh thuong
        
        If Sec_ID = "" Then
            MsgBox "B¹n ph¶i chän khu vùc tr­íc khi më bµn! C¶m ¬n!", vbInformation
        Else
        Table_ID = Left(lblTable(Index).Caption, InStr(lblTable(Index).Caption, Chr(13)))
        MaxInvoice = GetMaxInvoice_Number
        SaveSettingStr "SYSTEM", "MaxInvoice", MaxInvoice, myIniFile
        
        
           
        With rsinvoice_hold
            .Find "OnHoldID='" & Table_ID & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    'Khong ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                    currentBill = .Fields("Invoice_Number")
                Else
                    'Ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                    currentBill = MaxInvoice
                    .addNew
                    .Fields("Invoice_Number") = MaxInvoice
                    .Fields("OnHoldID") = Table_ID
                    .Fields("Cashier_ID") = UserID
                    .Fields("Store_ID") = Store_ID
                    .Fields("Occupied") = -1
                    .Fields("Section_ID") = Sec_ID
                    .Fields("Status") = 0
                    .Update
                End If
        End With
        
        
          With rsInvoice_Notes
            .Find "Invoice_Number='" & currentBill & "'", , adSearchForward, adBookmarkFirst
                If .EOF Then
                    ' ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                    .addNew
                    .Fields("Invoice_Number") = currentBill
                    .Fields("Store_ID") = Store_ID
                    .Fields("OpenTime") = DateDefault & Format(Now, "HH:mm:ss")
                    .Fields("ClosingTime") = "C"
                    .Update
                    .Requery
                End If
        End With
        
        
        Set rsInvoice_Total = OpenCriticalTable("select * from Invoice_Totals ", cnData)
        With rsInvoice_Total
            If rsInvoice_Total.State = 1 And .RecordCount > 0 Then rsInvoice_Total.MoveFirst
            .Find "Invoice_Number='" & currentBill & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                   
                    If .Fields("InvoiceNotesUsed") = True Then
                        MsgBox "Bµn nµy ®· ®­îc më t¹i mét m¸y kh¸c, B¹n kh«ng thÓ më bµn nµy!!!"
                         picWait.Visible = False
                        probarWait.Value = 0
                        Exit Sub
                    End If
                    If .Fields("CustNum") <> "101" Then
                        Dim rscust As New ADODB.Recordset
                        Set rscust = Open_Table(cnData, "Customer")
                            rscust.Find "CustNum='" & .Fields("CustNum") & "'", , adSearchForward, adBookmarkFirst
                            If Not rscust.EOF Then
                                CustNo(0) = .Fields("CustNum")
                                CustNo(1) = rscust!CustName & ""
                                CustNo(2) = rscust!Acct_Balance
                                Discount = CDbl("0" & rscust.Fields("Discount"))
                                .Fields("InvoiceNotesUsed") = True
                                .Update
                                .Requery
                            End If
                    Else
                        Discount = .Fields("Discount")
                    End If
                    'Discount = .Fields("Discount")
                    .Fields("InvoiceNotesUsed") = True
                    .Update
                    .Requery
                Else
                    ' ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                    .addNew
                    .Fields("Invoice_Number") = currentBill
                    .Fields("Store_ID") = Store_ID
                    .Fields("CustNum") = "101"
                    .Fields("DateTime") = DateDefault & Format(Now, "HH:mm:ss")
                    .Fields("InvoiceNotesUsed") = True
                    .Fields("Status") = "O"
                    .Fields("Station_ID") = Sec_ID
                    .Fields("Cashier_ID") = UserID
                    .Fields("Payment_MeThod") = "CA"
                    .Fields("InvType") = 0
                    .Fields("Orig_OnHoldID") = Trim(Table_ID)
                    .Fields("Tax_Rate_ID") = 0
                    .Update
                    .Requery
                End If
        End With
           
      
             With frmOrder
                .Get_Secion = Sec_ID
                .GetBill_Number = currentBill
                .Get_Table_ID = Table_ID
                .Get_Discount = Discount
                .FormCall = 2
                picWait.Visible = False
                probarWait.Value = 0
                .Show vbModal
            End With
    '    picWait.Visible = False
        End If
        currentBill = ""
        Exit Sub
   Case 2  'Chuyen ban
       Dim DesTab As String
       Dim billDes As String
               DesTab = Left(lblTable(Index).Caption, InStr(lblTable(Index).Caption, Chr(13)))
    
            Set rsinvoice_hold = OpenCriticalTable("Select * from Invoice_OnHold where Section_ID='" & Sec_ID & "'", cnData)
            With rsinvoice_hold
                .Find "OnHoldID='" & DesTab & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
    '                BillDes = .Fields("Invoice_Number")
    '                StateCall = 3
    '                lblTable_Click (Index)
                    MsgBox " Bµn nµy ®· cã, vui lßng chän chøc n¨ng gép bµn !!", vbInformation
    '
                    isclick = False
                    picWait.Visible = False
                    StateCall = 1
                    Exit Sub
                End If
            End With
            Set rsinvoice_hold = OpenCriticalTable("Select * from Invoice_OnHold where Section_ID='" & LocationTranfer & "'", cnData)
            If rsinvoice_hold.State = 1 Then rsinvoice_hold.MoveFirst
            With rsinvoice_hold
                .Find "OnHoldID='" & TranferTable & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    .Fields("OnHoldID") = DesTab
                    .Fields("Section_ID") = Sec_ID
                    .Update
                End If
            End With
            
            Set rsInvoice_Total = OpenCriticalTable("select * from Invoice_Totals ", cnData)
    
            If rsInvoice_Total.State = 1 And rsInvoice_Total.RecordCount > 0 Then rsInvoice_Total.MoveFirst
                With rsInvoice_Total
                .Find "Invoice_Number='" & BillTranfer & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        .Fields("Orig_OnHoldID") = DesTab
                        .Fields("Station_ID") = Sec_ID
                        .Update
                    End If
                End With
            'kiem tra neu co in bep thi cap nhat chuyen ban
            If IsPrintTranfer = True Then
                If Check_System_KP("02") Then
                    'Cap nhat chuyen ban trong chi tiet order
                    Call Update_Table_In_Order_Details(BillTranfer, BillTranfer, 1, TranferTable, DesTab)
                    'Cap nhat thông tin chuyên bàn vào bang chuyên gôp bàn.
                    'Kiem tra nêu bang chuyen gôp bàn chua có thì tao mói.
                    If Check_Table_exist("Tranfer_Joint_table") = False Then
                        Call Create_Table_Joint_Tranfer
                    End If
                    Call Update_Tranfer(BillTranfer, BillTranfer, LocationTranfer, Sec_ID, TranferTable, DesTab, 1)
                    For i = 1 To UBound(arrPriterKP)
                        If arrPriterKP(i) <> "" Then
                            Dim rsPrintKP As New ADODB.Recordset
                            Dim printer_Name As String
                            Set rsPrintKP = Open_Table(cnData, "Printer_Mapping")
                            With rsPrintKP
                                .Find "PrinterName='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
                                If Not .EOF Then
                                    Call Print_Tranfer_Bill(BillTranfer, .Fields("Details"))
                                End If
                            End With
                        End If
        
                    Next
        
                End If
            End If
    '        '==============================================
            frmMessage.Show vbModal
            picWait.Visible = False
            probarWait.Value = 0
            cnData.Execute "delete * from Tranfer_Joint_table"
            StateCall = 1
    Case 3  'Gop ban
            'Tim so Bill Ban dich duoc chuyen toi
            Dim DesTable As String
            Dim DesBill As String
            Dim rsFindBill As New ADODB.Recordset
            DesTable = Left(lblTable(Index).Caption, InStr(lblTable(Index).Caption, Chr(13)))
            If TranferTable = DesTable Then
                StateCall = 1
                TranferTable = ""
                Exit Sub
            End If
            Set rsFindBill = OpenCriticalTable("SELECT Invoice_OnHold.Invoice_Number, Invoice_OnHold.OnHoldID FROM Invoice_OnHold WHERE Invoice_OnHold.Section_ID='" & Sec_ID & "' and Invoice_OnHold.OnHoldID='" & DesTable & "'", cnData)
            If rsFindBill.RecordCount > 0 Then
                If Not rsFindBill.EOF Then
                    DesBill = rsFindBill.Fields("Invoice_Number")
                End If
            Else
                StateCall = 2
                lblTable_Click (Index)
                Exit Sub
            End If
            '======================================================================
            'Cap nhat so bill cua Ban dich vao danh muc hang ban voi so Bill cua ban Nguon
            Dim rsInvoice_Itemized As New ADODB.Recordset
            Dim rsmaxLine As New ADODB.Recordset
            i = 0
            Set rsmaxLine = OpenCriticalTable("select Max(Invoice_Itemized.LineNum)as MaxLine from Invoice_Itemized ", cnData)
            Set rsInvoice_Itemized = OpenCriticalTable("select * from Invoice_Itemized where Invoice_Number=" & BillTranfer, cnData)
            If rsmaxLine.RecordCount > 0 Then
                If Not rsmaxLine.EOF Then
                    i = CInt("0" & rsmaxLine.Fields("MaxLine")) + 1
                End If
            End If
                With rsInvoice_Itemized
                    Do While Not .EOF
                        .Fields("Invoice_Number") = DesBill
                        .Fields("LineNum") = i
                        .Update
                    .MoveNext
                    i = i + 1
                    Loop
                End With
            
            '======================================================================
            'Xoa ban nguon da tam tinh
            Set rsinvoice_hold = OpenCriticalTable("Select * from Invoice_OnHold where Section_ID='" & LocationTranfer & "'", cnData)
            With rsinvoice_hold
                .Find "Invoice_Number=" & BillTranfer, , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    If CDbl("0" & billDes) <> BillTranfer Then
                        .Delete adAffectCurrent
                    End If
                End If
            End With
            
            '======================================================================
            'Cap nhat gio dong ban cho ban nguon
            If rsInvoice_Notes.State = 0 Then
                Set rsInvoice_Notes = Open_Table(cnData, "Invoice_Totals_Notes")
            End If
            With rsInvoice_Notes
                .Find "Invoice_Number=" & BillTranfer, , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    .Fields("ClosingTime") = DateDefault & Format(Now, "HH:mm:ss")
                    .Update
                End If
            End With
            
            '======================================================================
            'Cap nhat tong so tien cua bill nguon sang bill dich, Cap nhat status =T
            Dim dblTotal_Org As Double
            Set rsInvoice_Total = OpenCriticalTable("select * from Invoice_Totals ", cnData)
            If rsInvoice_Total.State = 0 Then Set rsInvoice_Total = Open_Table(cnData, "Invoice_Totals")
            If rsInvoice_Total.State <> 0 Then rsInvoice_Total.MoveFirst
            With rsInvoice_Total
                .Find "Invoice_Number=" & BillTranfer, , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    dblTotal_Org = CDbl("0" & .Fields("Total_Price"))
                    .Fields("Total_Price") = 0
                    .Fields("Grand_Total") = 0
                    .Fields("Status") = "T" & DesBill & "-" & dblTotal_Org
                    .Update
                End If
            End With
            'Cap nhat Total va Grand Total vao bill dich
            '====================================================================
            If rsInvoice_Total.State <> 0 Then rsInvoice_Total.MoveFirst
            With rsInvoice_Total
                .Find "Invoice_Number=" & DesBill, , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    'dblTotal_Org = CDbl("0" & .Fields("Total_Price"))
                    .Fields("Total_Price") = .Fields("Total_Price") + dblTotal_Org
                    .Fields("Grand_Total") = .Fields("Grand_Total") + dblTotal_Org
                    .Update
                End If
            End With
    '        'Cap nhat Gop ban vao bang chi tiet order
            If IsPrintTranfer = True Then
                If Check_System_KP("02") Then
                    'Cap nhat chuyen ban trong chi tiet order
                    Call Update_Table_In_Order_Details(BillTranfer, CDbl("0" & DesBill), 2, TranferTable, DesTable)
                    'Cap nhat thông tin chuyên bàn vào bang chuyên gôp bàn.
                    'Kiem tra nêu bang chuyen gôp bàn chua có thì tao mói.
                    If Check_Table_exist("Tranfer_Joint_table") = False Then
                        Call Create_Table_Joint_Tranfer
                    End If
                    Call Update_Tranfer(BillTranfer, CDbl("0" & DesBill), LocationTranfer, Sec_ID, TranferTable, DesTable, 2)
                    For i = 1 To UBound(arrPriterKP)
                        If arrPriterKP(i) <> "" Then
                            Set rsPrintKP = Open_Table(cnData, "Printer_Mapping")
                            With rsPrintKP
                                .Find "PrinterName='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
                                If Not .EOF Then
                                    Call Print_Tranfer_Bill(CDbl("0" & DesBill), .Fields("Details"))
                                End If
                            End With
                        End If
        
                    Next
                End If
            End If
            
            frmMessage.Show vbModal
            StateCall = 1
            'Mo form Order
            With frmOrder
                .Get_Secion = Sec_ID
                .GetBill_Number = DesBill
                .Get_Table_ID = DesTable
                .Get_Discount = Discount
                .FormCall = 2
                picWait.Visible = False
                probarWait.Value = 0
    
                .Show vbModal
            End With
            
    '        picWait.Visible = False
        
    Case 4  'Chuyen mon
        If MsgBox("B¹n cã muèn chuyÓn mãn sang bµn nµy kh«ng ?", vbYesNo) = vbYes Then
            Dim TableDestination As String
            Dim BillDestination As Double
            Dim dblTrans As Double
            TableDestination = Left(lblTable(Index).Caption, InStr(lblTable(Index).Caption, Chr(13)))
            
             With rsTranfer
                If rsTranfer.State = 1 Then
                    Do While Not .EOF
                        dblTrans = dblTrans + CDbl("0" & .Fields("Amt"))
                    .MoveNext
                    Loop
                End If
            End With
            Set rsinvoice_hold = OpenCriticalTable("select * from Invoice_OnHold where Section_ID='" & Sec_ID & "'", cnData)
            rsinvoice_hold.Find "OnholdID='" & TableDestination & "'", , adSearchForward, adBookmarkFirst
            If Not rsinvoice_hold.EOF Then
                BillDestination = CDbl("0" & rsinvoice_hold.Fields("Invoice_Number"))
                'Cap nhat so bill cua Ban dich vao danh muc hang ban voi so Bill cua ban Nguon
                Set rsmaxLine = OpenCriticalTable("select Max(Invoice_Itemized.LineNum)as MaxLine from Invoice_Itemized where Invoice_Number=" & BillDestination, cnData)
                If rsmaxLine.RecordCount > 0 Then
                    If Not rsmaxLine.EOF Then
                        i = CInt("0" & rsmaxLine.Fields("MaxLine")) + 1
                    End If
                End If
                Set rsInvoice_Itemized = Open_Table(cnData, "Invoice_Itemized")
                If rsTranfer.State = 1 Then rsTranfer.MoveFirst
                With rsInvoice_Itemized
                    .addNew
                    .Fields("Invoice_Number") = BillDestination
                    .Fields("LineNum") = i
                    .Fields("ItemNum") = rsTranfer.Fields("PluNo")
                    .Fields("Quantity") = rsTranfer.Fields("Qty")
                    .Fields("PricePer") = rsTranfer.Fields("Std_Price1")
                    .Fields("DiffItemName") = rsTranfer.Fields("PluName")
                    .Fields("LineDisc") = 0
                    .Fields("Store_ID") = Store_ID
                    .Update
                End With
            ''' Cap nhat tong tien cho Ban nay
                With rsInvoice_Total
                    .Find "Invoice_Number=" & BillDestination, , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        .Fields("Total_Price") = .Fields("Total_Price") + dblTrans
                        .Update
                        .Requery
                    End If
                End With
            Else
                If lblSection.Caption = "" Then
                    MsgBox "B¹n ph¶i chän khu vùc tr­íc khi më bµn! C¶m ¬n!", vbInformation
                Else
                Table_ID = Left(lblTable(Index).Caption, InStr(lblTable(Index).Caption, Chr(13)))
                MaxInvoice = GetMaxInvoice_Number
                SaveSettingStr "SYSTEM", "MaxInvoice", MaxInvoice, myIniFile
                With rsinvoice_hold
                    .Find "OnHoldID='" & Table_ID & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            'Khong ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                            BillDestination = .Fields("Invoice_Number")
                        Else
                            ' ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                            BillDestination = MaxInvoice
                            .addNew
                            .Fields("Invoice_Number") = MaxInvoice
                            .Fields("OnHoldID") = Table_ID
                            .Fields("Cashier_ID") = UserID
                            .Fields("Store_ID") = Store_ID
                            .Fields("Occupied") = -1
                            .Fields("Section_ID") = Sec_ID
                            .Fields("Status") = 0
                            .Update
'                            .Requery
                        End If
                End With
                
                With rsInvoice_Notes
                    .Find "Invoice_Number='" & BillDestination & "'", , adSearchForward, adBookmarkFirst
                        If .EOF Then
                            ' ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                            .addNew
                            .Fields("Invoice_Number") = BillDestination
                            .Fields("Store_ID") = Store_ID
                            .Fields("OpenTime") = DateDefault & Format(Now, "HH:mm:ss")
                            .Fields("ClosingTime") = "C"
                            .Update
                        End If
                End With
                
                With rsInvoice_Total
                    .Find "Invoice_Number='" & BillDestination & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            Set rscust = Open_Table(cnData, "Customer")
                                rscust.Find "CustNum='" & .Fields("CustNum") & "'", , adSearchForward, adBookmarkFirst
                                If Not rscust.EOF Then
                                    CustNo(0) = .Fields("CustNum")
                                    CustNo(1) = rscust!CustName
                                    CustNo(2) = rscust!Acct_Balance
                                    Discount = CDbl("0" & rscust.Fields("Discount"))
                                End If
                            Discount = .Fields("Discount")
                        Else
                            ' ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                            .addNew
                            .Fields("Invoice_Number") = BillDestination
                            .Fields("Store_ID") = Store_ID
                            .Fields("CustNum") = "101"
                            .Fields("DateTime") = DateDefault & Format(Now, "HH:mm:ss")
                            .Fields("InvoiceNotesUsed") = -1
                            .Fields("Status") = "O"
                            .Fields("Station_ID") = Sec_ID
                            .Fields("Cashier_ID") = UserID
                            .Fields("Payment_MeThod") = "CA"
                            .Fields("InvType") = 0
                            .Fields("Orig_OnHoldID") = Trim(Table_ID)
                            .Fields("Tax_Rate_ID") = 0
                            .Update
                        End If
                End With
                
                Dim rsInvoice_Items As New ADODB.Recordset
                Set rsInvoice_Items = Open_Table(cnData, "Invoice_Itemized")
                i = 0
                If rsTranfer.State = 1 Then rsTranfer.MoveFirst
                With rsTranfer
                    Do While Not .EOF
                        rsInvoice_Items.addNew
                        rsInvoice_Items.Fields("Invoice_Number") = MaxInvoice
                        rsInvoice_Items.Fields("LineNum") = i
                        rsInvoice_Items.Fields("ItemNum") = rsTranfer.Fields("PluNo")
                        rsInvoice_Items.Fields("Quantity") = rsTranfer.Fields("Qty")
                        rsInvoice_Items.Fields("PricePer") = rsTranfer.Fields("Std_Price1")
                        rsInvoice_Items.Fields("DiffItemName") = rsTranfer.Fields("PluName")
                        rsInvoice_Items.Fields("LineDisc") = 0
                        rsInvoice_Items.Fields("Store_ID") = Store_ID
                        rsInvoice_Items.Fields("Kit_Description") = rsTranfer.Fields("Kit_Desc")
                        rsInvoice_Items.Update
                    .MoveNext
                    i = i + 1
                    Loop
                End With
            
            End If
            End If
             With frmOrder
                .Get_Secion = Sec_ID
                .GetBill_Number = BillDestination
                .Get_Table_ID = TableDestination
                .Get_Discount = Discount
                .FormCall = 2
                picWait.Visible = False
                probarWait.Value = 0
    
                .Show vbModal
            End With
            
    '        picWait.Visible = False
            
            Delay (500)
            frmMessage.lblTitle.Caption = "Mãn b¹n chän ®· ®­îc chuyÓn sang bµn " & TableDestination
            frmMessage.Show vbModal
            
            StateCall = 1
            
            Set rsTranfer = Nothing
    Else
        If rsInvoice_Itemized.State = 1 Then rsInvoice_Itemized.MoveFirst
        If rsInvoice_Itemized.State = 0 Then
            Set rsInvoice_Itemized = Open_Table(cnData, "Invoice_Itemized")
        End If
        If rsmaxLine.State = 0 Then
            Set rsmaxLine = OpenCriticalTable("select Max(Invoice_Itemized.LineNum)as MaxLine from Invoice_Itemized where Invoice_Number=" & BillTranfer, cnData)
        End If
            If rsmaxLine.RecordCount > 0 Then
                If Not rsmaxLine.EOF Then
                    i = CInt("0" & rsmaxLine.Fields("MaxLine")) + 1
                End If
            End If
            Set rsInvoice_Itemized = Open_Table(cnData, "Invoice_Itemized")
            If rsTranfer.State = 1 Then rsTranfer.MoveFirst
            With rsInvoice_Itemized
                Do While Not rsTranfer.EOF
                    .addNew
                    .Fields("Invoice_Number") = BillTranfer
                    .Fields("LineNum") = i
                    .Fields("ItemNum") = rsTranfer.Fields("PluNo")
                    .Fields("Quantity") = rsTranfer.Fields("Qty")
                    .Fields("PricePer") = rsTranfer.Fields("Std_Price1")
                    .Fields("DiffItemName") = rsTranfer.Fields("PluName")
                    .Fields("LineDisc") = 0
                    .Fields("Store_ID") = Store_ID
                    .Fields("Kit_Description") = rsTranfer.Fields("Kit_Desc")
                    .Update
                rsTranfer.MoveNext
                i = i + 1
                Loop
            End With
        Set rsTranfer = Nothing
        frmMessage.lblTitle.Caption = "Mãn b·n chän ®· ®­îc chuyÓn l¹i bµn gèc !"
        frmMessage.Show vbModal
        
    End If
'    Case 5
'        With rsinvoice_hold
'            .Find "OnHoldID='" & Table_ID & "'", , adSearchForward, adBookmarkFirst
'                If Not .EOF Then
'                    Call Print_Buffer(.Fields("Invoice_Number"))
'                    StateCall = 1
'                End If
'        End With
'
    End Select
    dblTotal_Org = 0
    picWait.Visible = False
    StateCall = 1
    Discount = 0
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " lblTable_Click  "
End Sub

Public Sub Load_Section()
    On Error GoTo Handle
    Dim ctrl As Control
    Dim strright As String
    
'        strright = Get_Right_Location(UserID) 'Lay quyen khu vuc
    
        Dim i, a, b As Integer
        i = 1
        a = 0
        isclick = False
        If cnData.State > 0 Then
             Set rsSection = OpenCriticalTable("select * from Table_Diagram_Sections order by Location_ID ASC", cnData)
        Else
            Exit Sub
        End If
        If rsSection.EOF Then Exit Sub
        If iLoadSection = True Then
            For Each ctrl In Me
                If TypeOf ctrl Is MyButton And ctrl.Name = "cmdSection" Then
                    a = a + 1
                End If
            Next
            For b = 1 To a - 1
                Unload cmdSection(b)
            Next
            
        End If
            Do While Not rsSection.EOF
                Load cmdSection(i)
                With cmdSection(i)
                    If i = 1 Then
                        .Left = cmdSection(i - 1).Left + 20
                    Else
                        .Left = cmdSection(i - 1).Left + cmdSection(i - 1).Width + 20
                    End If
                    .top = cmdSection(i - 1).top
                    .Visible = True
'                    If Mid(Right("000000" & HexToBin(strright), rsSection.RecordCount), i, 1) = 1 Then
'                        .Enabled = True
'                        fraTable.Enabled = True
'                        Call LoadTable(rsSection.Fields("Location_ID"))
'                        Sec_ID = rsSection.Fields("Location_ID")
'                        lblSection.Caption = rsSection.Fields("Section_ID")
'                    Else
'                        .Enabled = False
'                        fraTable.Enabled = False
'                    End If
                    fraTable.Enabled = True
                    .Caption = rsSection.Fields("Section_ID")
                    .Tag = rsSection.Fields("Location_ID")
                End With
                i = i + 1
            rsSection.MoveNext
            Loop
        CountSection = rsSection.RecordCount
        iLoadSection = True
        a = 0
        
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "   Load_Section"
End Sub

Public Sub LoadTable(Section_ID As String)
On Error GoTo Handle
Dim rscolor As New ADODB.Recordset
Dim rsSeatedColor As New ADODB.Recordset
Dim rsVacantColor As New ADODB.Recordset
Dim rsSubtotalColor As New ADODB.Recordset
Dim rsInvoice_Total As New ADODB.Recordset

Dim i, J As Integer
i = 1: J = 1
    Dim str As String
    Dim ctrl As Control
    If CountTable > 0 Then
        For J = 1 To CountTable
'            DoEvents
            Unload lblTable(J)
            Unload Shape1(J)
            Unload lblTime(J)
        Next
    End If
    If cnData.State = 0 Then
        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881")
    End If
    'Lay Bang mau
    Dim TypeColor, SeatedColor, BlankTable, SubtotalColor As String
    TypeColor = "RESERVED"
    SeatedColor = "SEATED"
    BlankTable = "VACANT"
    SubtotalColor = "SUBTOTAL"
    Set rscolor = OpenCriticalTable("Select ReserveValue from ColorTablePlan where ReserveType='" & TypeColor & "'", cnData)
    Set rsSeatedColor = OpenCriticalTable("Select ReserveValue from ColorTablePlan where ReserveType='" & SeatedColor & "'", cnData)
    Set rsVacantColor = OpenCriticalTable("Select ReserveValue from ColorTablePlan where ReserveType='" & BlankTable & "'", cnData)
    Set rsSubtotalColor = OpenCriticalTable("Select ReserveValue from ColorTablePlan where ReserveType='" & SubtotalColor & "'", cnData)
    
    str = "select * from Table_Diagram where Section_ID='" & Section_ID & "'"
    Set rsTable = OpenCriticalTable(str, cnData)
    CountTable = rsTable.RecordCount
    Dim strTableTotal As String
    Do While Not rsTable.EOF
        Load lblTime(i)
        Load lblTable(i)
'        DoEvents
        With lblTable(i)
            .Left = rsTable.Fields("XPOS")
            .top = rsTable.Fields("YPOS")
            .Height = rsTable.Fields("Height")
            .Width = rsTable.Fields("width")
            strTableTotal = "SELECT Invoice_OnHold.Invoice_Number, Invoice_Totals.Store_ID,Invoice_Totals.DateTime,Invoice_Totals.Status," & _
            "Invoice_OnHold.OnHoldID, Invoice_Totals.Grand_Total, Invoice_Totals.Total_Price, " & _
            "Invoice_Totals.Orig_OnHoldID, Invoice_OnHold.Section_ID FROM Invoice_OnHold" & _
            " INNER JOIN Invoice_Totals ON Invoice_OnHold.Invoice_Number = Invoice_Totals.Invoice_Number " & _
            " where Invoice_OnHold.OnHoldID = '" & rsTable.Fields("Table_number") & Chr(13) & "' and Invoice_OnHold.Section_ID='" & Section_ID & "'"
            Set rsInvoice_Total = OpenCriticalTable(strTableTotal, cnData)
            If rsInvoice_Total.RecordCount > 0 Then
                If CDbl(rsInvoice_Total.Fields("Grand_Total")) > 0 Then
                    If rsInvoice_Total.Fields("Status") = "P" Then
                        .Caption = rsTable.Fields("Table_Number") & Chr(13) & Format(rsInvoice_Total.Fields("Grand_Total"), formatNum)
                        .BackStyle = 1
                        .BackColor = rsSubtotalColor.Fields("ReserveValue")
                        .Font.Size = rsTable.Fields("Cost_Center_Index")
                        lblTime(i).BackColor = vbRed
                    Else
                        .Caption = rsTable.Fields("Table_Number") & Chr(13) & Format(rsInvoice_Total.Fields("Grand_Total"), formatNum)
                        .BackStyle = 1
                        .BackColor = rscolor.Fields("ReserveValue")
                        .Font.Size = rsTable.Fields("Cost_Center_Index")
                        lblTime(i).BackColor = vbRed
                    End If
                Else
                    .Caption = rsTable.Fields("Table_Number") & Chr(13) & Format(rsInvoice_Total.Fields("Grand_Total"), formatNum)
                    .BackStyle = 1
                    .BackColor = rsSeatedColor.Fields("ReserveValue")
                    .Font.Size = rsTable.Fields("Cost_Center_Index")
                    lblTime(i).BackColor = vbRed
                End If
                lblTime(i).Caption = Right(rsInvoice_Total.Fields("DateTime"), 8)
            Else
                .Caption = rsTable.Fields("Table_Number") & Chr(13)
                .BackStyle = 1
                .BackColor = rsVacantColor.Fields("ReserveValue")
                .Font.Size = rsTable.Fields("Cost_Center_Index")
                lblTime(i).Caption = ""
                lblTime(i).BackStyle = 0
            End If
            .Visible = True
        End With
        Load Shape1(i)
        With Shape1(i)
            .Left = lblTable(i).Left - 40
            .top = lblTable(i).top - 45
            .Height = lblTable(i).Height + 100
            .Width = lblTable(i).Width + 100
            .Shape = rsTable.Fields("ShapeType")
            .Visible = True
        End With
        With lblTime(i)
            .Left = lblTable(i).Left
            .top = lblTable(i).top + lblTable(i).Height - lblTime(i).Height  'fraTable.top + lblTable(i).top - 300 ' + lblTable(i).Height
            .Width = lblTable(i).Width
            .Height = 255
            .Visible = True
        End With
    rsTable.MoveNext
    i = i + 1
    Loop
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "  LoadTable"
End Sub


Public Property Let FormState(ByVal vNewValue As Variant)
    StateCall = vNewValue
End Property


Public Property Let GetTableTranfer(ByVal vNewValue As Variant)
    TranferTable = vNewValue
End Property


Public Property Let GetLocationTranfer(ByVal vNewValue As Variant)
    LocationTranfer = vNewValue
End Property


Public Property Let GetBillTranfer(ByVal vNewValue As Variant)
    BillTranfer = vNewValue
End Property


Public Sub CheckRight()
On Error GoTo Handle
    Dim res As New ADODB.Recordset
        
        Set res = LoadPasswordData
        With MyRight
            res.MoveFirst
            Do While Not res.EOF
                If StrComp(res.Fields("UserName"), userName, 1) = 0 Then
                    .FullRight = res.Fields("Right")
                    .Sodoban = RightDeCode(Left(.FullRight, 16))
                    .Danhmuc = RightDeCode(Mid(.FullRight, 33, 16))
                    Exit Do
                End If
                res.MoveNext
            Loop
            If Mid(.Sodoban, 1, 1) = 0 Then
                  cmdEdit.Enabled = False
            Else: cmdEdit.Enabled = True
            End If
            If Mid(.Danhmuc, 1, 1) = 0 Then
                  cmdOption.Enabled = False
            Else: cmdOption.Enabled = True
            End If
        End With

    CloseRecordset res
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " CheckRight"
End Sub

Private Sub lblTable_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuEdit, 0
    End If
End Sub

Private Sub lblTime_Click(Index As Integer)
    Call lblTable_Click(Index)
End Sub

Private Sub lblUnCheckAll_Click()
    chkCustomer.Value = False
    ChkSale.Value = False
    chkDMThuchi.Value = False
    chkGroup.Value = False
    chkItems.Value = False
    chkTablePlan.Value = False
    chkThuchu.Value = False
    chkVendor.Value = False
End Sub

Private Sub mnuEditTable_Click()
    frmEditTablePlan.Show vbModal
End Sub

Private Sub mnuGroup_Click()
    frmDepartement.Show vbModal
End Sub

Private Sub mnuitem_Click()
    frmItems.Show vbModal
End Sub

Private Sub mnuLocation_Click()
    frmLocationName.Show vbModal
End Sub

Private Sub mnuprint_Click()
    frmPrintDefault.Show vbModal
End Sub

Private Sub mnuprinterName_Click()
    frmPrintName.Show vbModal
End Sub

Private Sub mnuReport_Click()
    frmSetup.Show vbModal
    
End Sub

'Private Sub TextFly_Timer()
''On Error GoTo Handle
''
''    'Do While Len(sText_Fly) < Len(sText_org)
''    If sText_org = "" Then Exit Sub
''        sText_Fly = sText_Fly & Left(sText_org, 1)
''        sText_org = Right(sText_org, Len(sText_org) - 1)
''    'Loop
''    Me.Caption = sText_Fly
''    If Len(sText_Fly) = Len(" Cty TNHH TM & DV Phóc Th¹nh Vinh - Gi¶i ph¸p b¸n hµng chuyªn nghiÖp" & _
''                " cho Nhµ hµng,Cafe,Bar,Karaoke...Hç trî KT 24/7:0918.655.887-0903.613.673-0909.118.669") Then
''        Delay (3000)
''        sText_Fly = ""
''        sText_org = " Cty TNHH TM & DV Phóc Th¹nh Vinh - Gi¶i ph¸p b¸n hµng chuyªn nghiÖp" & _
''                " cho Nhµ hµng,Cafe,Bar,Karaoke...Hç trî KT 24/7:0918.655.887-0903.613.673-0909.118.669"
''
''    End If
''Exit Sub
''Handle:
''    MsgBox Err.Number & Err.Description & Me.Name & " TextFly_Timer"
'End Sub

'Private Sub Timer1_Timer()
'On Error GoTo Handle
'    TimeSync = TimeSync + 1
'    If check_Backup = False Then Exit Sub
'    If TimeSync = GetTimeSync Then
'        If gfSynchronizeData = True Then
'            TimeSync = 0
'        Else
'            MsgBox "Ch­a ®ång bé"
'            TimeSync = 0
'            lblSync.Visible = False
'
'        End If
'    End If
'Exit Sub
'Handle:
'MsgBox Err.Number & Err.Description & Me.Name & " Timer1_Timer"
'End Sub

Public Function GetTimeSync() As Double
    On Error GoTo Handle
    Dim i As Double
    Dim rsInfor As New ADODB.Recordset
    If cnData.State = 0 Then
        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881")
    End If
    Set rsInfor = Open_Table(cnData, "Setup")
    With rsInfor
        If Not rsInfor.EOF Then
            i = Hour(.Fields("TimeSync")) * 3600 + Minute(.Fields("TimeSync")) * 60
            AmountBackup = .Fields("AmountLimited")
        End If
    End With
    
    GetTimeSync = i
    Exit Function
Handle:
    GetTimeSync = 0
    MsgBox Err.Number & Err.Description & Me.Name & " GetTimeSync"
End Function

Public Function gfSynchronizeData() As Boolean
On Error GoTo Handle
    gfSynchronizeData = False
    Dim cnBackup As New ADODB.Connection
    Dim cnOrg As New ADODB.Connection
    Dim rsOrg, rsOn_Hold As New ADODB.Recordset
    lblTitle.Caption = "D÷ liÖu b¸n hµng"
    Delay (500)
    Set rsOn_Hold = OpenCriticalTable("Select * from Invoice_OnHold where Invoice_Number <>0", cnData)
    If rsOn_Hold.EOF Then
        If Dir(WorkingFolder & "\Database.mdb", vbDirectory) <> "" Then
            Set cnOrg = Get_Connection(WorkingFolder & "\Database.mdb", "100881")
        End If
        If Dir(BackupFolder & "\Database.mdb", vbDirectory) <> "" Then
            Set cnBackup = Get_Connection(BackupFolder & "\Database.mdb", "100881")
        Else
            Exit Function
        End If
        'Lay so tien gioi han backup
        Call GetTimeSync
        If cnBackup.State = 0 Then Exit Function
            Set rsOrg = OpenCriticalTable("select * from Invoice_Totals where status<>'O' and Synchronized= False ", cnOrg)
            With rsOrg
                .Sort = "Invoice_Number ASC"
                Do While Not .EOF
                    Max_Invoice_Backup = GetMax_Invoice_Backup
                    Call gfBackup_Invoice_Notes(cnBackup, cnOrg, .Fields("Invoice_Number"))
                    Call gfBackup_Invoice_Totals(cnBackup, cnOrg, .Fields("Invoice_Number"))
                    Call gfBackup_Invoice_Itemized(cnBackup, cnOrg, .Fields("Invoice_Number"))
                    Call gfBackup_Deleted_Item(cnBackup, cnOrg, .Fields("Invoice_Number"))
                    Call gfBackup_Invoice_Per(cnBackup, cnOrg, .Fields("Invoice_Number"))
                    Call gfBackup_Invoice_Kitchen_Order_Mast(cnBackup, cnOrg, .Fields("Invoice_Number"))
                    Call gfBackup_Invoice_Kitchen_Order_Items(cnBackup, cnOrg, .Fields("Invoice_Number"))
                .MoveNext
                Loop
            End With
            'Xoa tat ca du lieu sau ki backup voi Tong tien tren HD > Amount_Limited
            Call Delete_Invoice_AmountLarger(AmountBackup)
            cnOrg.Execute "delete * from Items_Deleted " 'where Invoice_Num=" & .Fields("Invoice_Number")
            cnOrg.Execute "delete * from Invoice_Totals where Invoice_Number<>0"
            cnOrg.Execute "delete * from Tranfer_Joint_table"
        gfSynchronizeData = True
    Else
        MsgBox "Vui lßng thanh to¸n hÕt bµn ®ang më tr­íc khi ®ång bé d÷ liÖu"
    End If
Exit Function
Handle:
gfSynchronizeData = False
MsgBox Err.Number & Err.Description & Me.Name & " gfSynchronizeData"
End Function

Public Sub gfBackup_Invoice_Notes(cnBackup As Connection, cnOrg As Connection, Invoice_Num As Double)
On Error GoTo Handle
    Dim rsInvoice_Note_Org As New ADODB.Recordset
    Dim rsInvoice_Note_Des As New ADODB.Recordset
'    cnBackup.BeginTrans
'    cnOrg.BeginTrans
        Set rsInvoice_Note_Org = OpenCriticalTable("Select * from Invoice_Totals_Notes where Invoice_Number=" & Invoice_Num, cnOrg)
        Set rsInvoice_Note_Des = Open_Table(cnBackup, "Invoice_Totals_notes")
        
        With rsInvoice_Note_Org
            Do While Not .EOF
                With rsInvoice_Note_Des
                    .addNew
                    .Fields("Invoice_Number") = Max_Invoice_Backup  'rsInvoice_Note_Org.Fields("Invoice_Number")
'                    .Fields("Invoice_Number") = rsInvoice_Note_Org.Fields("Invoice_Number")
                    .Fields("Store_ID") = rsInvoice_Note_Org.Fields("Store_ID")
                    .Fields("OpenTime") = rsInvoice_Note_Org.Fields("OpenTime")
                    .Fields("ClosingTime") = rsInvoice_Note_Org.Fields("ClosingTime")
                    .Fields("Total_Minute") = rsInvoice_Note_Org.Fields("Total_Minute")
                    .Fields("Karaoke_Amount") = rsInvoice_Note_Org.Fields("Karaoke_Amount")
                    .Update
                    .Requery
                End With
            .MoveNext
            Loop
        End With
'    cnOrg.Execute "Delete * from Invoice_Totals_notes where Invoice_Number=" & Invoice_Num
'    cnOrg.Execute "Delete * from Invoice_Totals where Invoice_Number=" & Invoice_Num
'    cnBackup.CommitTrans
'    cnOrg.CommitTrans
Exit Sub
Handle:
'cnBackup.RollbackTrans
'cnOrg.RollbackTrans
MsgBox Err.Number & Err.Description & Me.Name & "gfBackup_Invoice_Notes"
End Sub

Public Sub gfBackup_Invoice_Totals(cnBackup As Connection, cnOrg As Connection, Invoice_Num As Double)
On Error GoTo Handle
    Dim rsInvoice_Totals_Org As New ADODB.Recordset
    Dim rsInvoice_Totals_Des As New ADODB.Recordset
    Dim i As Integer
'    cnBackup.BeginTrans
'    cnOrg.BeginTrans
        Set rsInvoice_Totals_Org = OpenCriticalTable("Select * from Invoice_Totals where Invoice_Number=" & Invoice_Num, cnOrg)
        Set rsInvoice_Totals_Des = Open_Table(cnBackup, "Invoice_Totals")
        With rsInvoice_Totals_Org
            i = 0
            Do While Not .EOF
                With rsInvoice_Totals_Des
                    .addNew
                    .Fields("Invoice_Number") = Max_Invoice_Backup  'rsInvoice_Totals_Org.Fields("Invoice_Number")
'                    .Fields("Invoice_Number") = rsInvoice_Totals_Org.Fields("Invoice_Number")
                    .Fields("Store_ID") = rsInvoice_Totals_Org.Fields("Store_ID")
                    .Fields("CustNum") = rsInvoice_Totals_Org.Fields("CustNum")
                    .Fields("DateTime") = rsInvoice_Totals_Org.Fields("DateTime")
                    .Fields("Total_Cost") = rsInvoice_Totals_Org.Fields("Total_Cost")
                    .Fields("Discount") = rsInvoice_Totals_Org.Fields("Discount")
                    .Fields("KarDiscount") = rsInvoice_Totals_Org.Fields("KarDiscount")
                    .Fields("Total_Price") = rsInvoice_Totals_Org.Fields("Total_Price")
                    .Fields("Total_Tax1") = rsInvoice_Totals_Org.Fields("Total_Tax1")
                    .Fields("Total_Tax2") = rsInvoice_Totals_Org.Fields("Total_Tax2")
                    .Fields("Total_Tax3") = rsInvoice_Totals_Org.Fields("Total_Tax3")
                    .Fields("Grand_Total") = rsInvoice_Totals_Org.Fields("Grand_Total")
                    .Fields("Amt_Tendered") = rsInvoice_Totals_Org.Fields("Amt_Tendered")
                    .Fields("Amt_Change") = rsInvoice_Totals_Org.Fields("Amt_Change")
                    .Fields("InvoiceNotesUsed") = True
                    .Fields("Status") = rsInvoice_Totals_Org.Fields("Status")
                    .Fields("Cashier_ID") = rsInvoice_Totals_Org.Fields("Cashier_ID")
                    .Fields("Station_ID") = rsInvoice_Totals_Org.Fields("Station_ID")
                    .Fields("Payment_Method") = rsInvoice_Totals_Org.Fields("Payment_Method")
                    .Fields("Acct_Balance_Due") = rsInvoice_Totals_Org.Fields("Acct_Balance_Due")
                    
                    .Fields("InvType") = rsInvoice_Totals_Org.Fields("InvType")
                    .Fields("Orig_OnHoldID") = rsInvoice_Totals_Org.Fields("Orig_OnHoldID")
                    .Fields("Tax_Rate_ID") = rsInvoice_Totals_Org.Fields("Tax_Rate_ID")
                    .Fields("OrderMan") = rsInvoice_Totals_Org.Fields("OrderMan")
                    .Fields("Service_Charge") = rsInvoice_Totals_Org.Fields("Service_Charge")
                    .Fields("VATFee") = rsInvoice_Totals_Org.Fields("VATFee")
                    .Fields("Adjustment1") = rsInvoice_Totals_Org.Fields("Adjustment1")
                    .Fields("Adj1Rate") = rsInvoice_Totals_Org.Fields("Adj1Rate")
                    .Fields("Adjustment2") = rsInvoice_Totals_Org.Fields("Adjustment2")
                    .Fields("Adj2Rate") = rsInvoice_Totals_Org.Fields("Adj2Rate")
                    .Fields("Adjustment3") = rsInvoice_Totals_Org.Fields("Adjustment3")
                    .Fields("Adjustment4") = rsInvoice_Totals_Org.Fields("Adjustment4")
                    .Fields("AddMoney") = rsInvoice_Totals_Org.Fields("AddMoney")
                    .Fields("Synchronized") = True
                    .Fields("Personals") = rsInvoice_Totals_Org.Fields("Personals")
                    .Update
'                    .Requery
                End With
            .MoveNext
            i = i + 1
            Loop
        End With
        cnData.Execute "Update Invoice_Totals set Synchronized= Yes where Invoice_Number=" & Invoice_Num
'    cnBackup.CommitTrans
'    cnOrg.CommitTrans
Exit Sub
Handle:
'    cnBackup.RollbackTrans
'    cnOrg.RollbackTrans
    MsgBox Err.Number & Err.Description & Me.Name & "gfBackup_Invoice_Totals"
End Sub


Public Sub gfBackup_Invoice_Itemized(cnBackup As Connection, cnOrg As Connection, Invoice_Num As Double)
On Error GoTo Handle
Dim i As Integer
    Dim rsInvoice_Item_Org As New ADODB.Recordset
    Dim rsInvoice_Item_Des As New ADODB.Recordset
    
'    cnBackup.BeginTrans
'    cnOrg.BeginTrans
        Set rsInvoice_Item_Org = OpenCriticalTable("Select * from Invoice_Itemized where Invoice_Number=" & Invoice_Num, cnOrg)
        Set rsInvoice_Item_Des = Open_Table(cnBackup, "Invoice_Itemized")
        
        i = 0
        With rsInvoice_Item_Org
        
            Do While Not .EOF
                With rsInvoice_Item_Des
                    .addNew
                    .Fields("Invoice_Number") = Max_Invoice_Backup
'                    .Fields("Invoice_Number") = rsInvoice_Item_Org.Fields("Invoice_Number")
                    .Fields("LineNum") = rsInvoice_Item_Org.Fields("LineNum")
                    .Fields("ItemNum") = rsInvoice_Item_Org.Fields("ItemNum")
                    .Fields("Quantity") = rsInvoice_Item_Org.Fields("Quantity")
                    .Fields("PricePer") = rsInvoice_Item_Org.Fields("PricePer")
                    .Fields("Tax1Per") = rsInvoice_Item_Org.Fields("Tax1Per")
                    .Fields("Tax2Per") = rsInvoice_Item_Org.Fields("Tax2Per")
                    .Fields("Tax3Per") = rsInvoice_Item_Org.Fields("Tax3Per")
                    .Fields("Serial_Num") = rsInvoice_Item_Org.Fields("Serial_Num")
                    .Fields("Kit_Description") = rsInvoice_Item_Org.Fields("Kit_Description")
                    .Fields("LineDisc") = rsInvoice_Item_Org.Fields("LineDisc")
                    .Fields("DiffItemName") = rsInvoice_Item_Org.Fields("DiffItemName")
                    .Fields("Store_ID") = rsInvoice_Item_Org.Fields("Store_ID")
                    .Fields("Section_ID") = rsInvoice_Item_Org.Fields("Section_ID")
                    .Fields("Person") = rsInvoice_Item_Org.Fields("Person")
                    .Fields("Returned") = rsInvoice_Item_Org.Fields("Returned")
                    .Update
'                    .Requery
                End With
            .MoveNext
            i = i + 1
            Loop
        End With
    'cnOrg.Execute "Delete * from Invoice_Itemized where Invoice_Number=" & Invoice_Num
'    cnBackup.CommitTrans
'    cnOrg.CommitTrans
Exit Sub
Handle:
'    cnBackup.RollbackTrans
'    cnOrg.RollbackTrans
    MsgBox Err.Number & Err.Description & Me.Name & "gfBackup_Invoice_Itemized"
End Sub


Public Sub gfBackup_Deleted_Item(cnBackup As Connection, cnOrg As Connection, Invoice_Num As Double)
On Error GoTo Handle
    Dim rsInvoice_ItemDelete_Org As New ADODB.Recordset
    Dim rsInvoice_ItemDelete_Des As New ADODB.Recordset
    
'    cnBackup.BeginTrans
'    cnOrg.BeginTrans
        Set rsInvoice_ItemDelete_Org = OpenCriticalTable("Select * from Items_Deleted where Invoice_Num=" & Invoice_Num, cnOrg)
        Set rsInvoice_ItemDelete_Des = Open_Table(cnBackup, "Items_Deleted")
        With rsInvoice_ItemDelete_Org
            Do While Not .EOF
                With rsInvoice_ItemDelete_Des
                    .addNew
                    .Fields("Sec_ID") = rsInvoice_ItemDelete_Org.Fields("Sec_ID")
                    .Fields("Invoice_Num") = rsInvoice_ItemDelete_Org.Fields("Invoice_Num")
                    .Fields("Table_ID") = rsInvoice_ItemDelete_Org.Fields("Table_ID")
                    .Fields("Cashier_ID") = rsInvoice_ItemDelete_Org.Fields("Cashier_ID")
                    .Fields("PluNo") = rsInvoice_ItemDelete_Org.Fields("PluNo")
                    .Fields("Quantity") = rsInvoice_ItemDelete_Org.Fields("Quantity")
                    .Fields("Price") = rsInvoice_ItemDelete_Org.Fields("Price")
                    .Fields("Amount") = rsInvoice_ItemDelete_Org.Fields("Amount")
                    .Fields("DateTime") = rsInvoice_ItemDelete_Org.Fields("DateTime")
                    .Fields("Ordered") = rsInvoice_ItemDelete_Org.Fields("Ordered")
                    .Fields("Reason") = rsInvoice_ItemDelete_Org.Fields("Reason")
                    .Fields("PrintCount") = rsInvoice_ItemDelete_Org.Fields("PrintCount")
                    .Update
'                    .Requery
                End With
            .MoveNext
            Loop
        End With
'    cnBackup.CommitTrans
'    cnOrg.CommitTrans
Exit Sub
Handle:
'    cnBackup.RollbackTrans
'    cnOrg.RollbackTrans
    MsgBox Err.Number & Err.Description & Me.Name & "gfBackup_Deleted_Item"
End Sub
'''''''''''''''''''''''''''''''
Public Sub gfBackup_Invoice_Per(cnBackup As Connection, cnOrg As Connection, Invoice_Num As Double)
On Error GoTo Handle
    Dim rsInvoice_Per_Org As New ADODB.Recordset
    Dim rsInvoice_Per_Des As New ADODB.Recordset
    Dim i As Integer
'    cnBackup.BeginTrans
'    cnOrg.BeginTrans
        Set rsInvoice_Per_Org = OpenCriticalTable("Select * from Invoice_Totals_Person_Mapping where Invoice_Number=" & Invoice_Num, cnOrg)
        Set rsInvoice_Per_Des = Open_Table(cnBackup, "Invoice_Totals_Person_Mapping")
        i = 0
        With rsInvoice_Per_Org
            Do While Not .EOF
                With rsInvoice_Per_Des
                    .addNew
                    .Fields("Invoice_Number") = Max_Invoice_Backup '+ i 'rsInvoice_Note_Org.Fields("Invoice_Number")
'                    .Fields("Invoice_Number") = rsInvoice_Per_Org.Fields("Invoice_Number")
                    .Fields("Store_ID") = Store_ID
                    .Fields("SeatNum") = rsInvoice_Per_Org.Fields("SeatNum")
                    .Update
'                    .Requery
                End With
            .MoveNext
            i = i + 1
            Loop
        End With
    'cnOrg.Execute "Delete * from Invoice_Totals_Person_Mapping where Invoice_Number=" & Invoice_Num
'    cnBackup.CommitTrans
'    cnOrg.CommitTrans
Exit Sub
Handle:
'cnBackup.RollbackTrans
'cnOrg.RollbackTrans
MsgBox Err.Number & Err.Description & Me.Name & "gfBackup_Invoice_Notes"
End Sub


Public Function check_Backup() As Boolean
On Error GoTo Handle
    check_Backup = False
    If ArrayFlag(SF(3), 3) = 1 Then
        check_Backup = True
    End If
    
Exit Function
Handle:
check_Backup = False
MsgBox Err.Number & Err.Description & Me.Name & "check_Backup"
End Function

Public Function GetMax_Invoice_Backup() As Double
On Error GoTo Handle
Dim Max_Invoice As Double
    Dim rsmax As New ADODB.Recordset
    Dim cnmax As New ADODB.Connection
    If Dir(BackupFolder & "\Database.mdb", vbDirectory) <> "" Then
        Set cnmax = Get_Connection(BackupFolder & "\Database.mdb", "100881")
    End If
    Set rsmax = OpenCriticalTable("select Max(Invoice_Number) as maxInvoice from Invoice_Totals", cnmax)
    If rsmax.RecordCount <> 0 Then
        If Not rsmax.EOF Then
            If " " & rsmax.Fields("MaxInvoice") = " " Then
                Max_Invoice = CDbl("0" & " " & rsmax.Fields("maxInvoice")) + 1
            Else
                Max_Invoice = CDbl("0" & rsmax.Fields("maxInvoice")) + 1
            End If
        Else
            MaxInvoice = 1
        End If
    End If
    GetMax_Invoice_Backup = Max_Invoice
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " GetMax_Invoice_Backup"
End Function
'
'Public Function Check_Karaoke() As Boolean
'On Error GoTo Handle
'Dim useKaraoke As Boolean
'    Dim rsLocation As New ADODB.Recordset
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\database.mdb", "100881")
'
'    Set rsLocation = Open_Table(cnData, "Table_Diagram_Sections")
'    With rsLocation
'        If .RecordCount > 0 And Not .EOF Then .MoveFirst
'            .Find "Location_ID='" & Sec_ID & "'", , adSearchForward, adBookmarkFirst
'            If Not .EOF Then
'                If .Fields("Used_Karaoke") = True Then
'                    useKaraoke = True
'                Else
'                    useKaraoke = False
'                End If
'            End If
'
'    End With
'Check_Karaoke = useKaraoke
'
'Exit Function
'Handle:
'    MsgBox Err.Number & Err.Description & Me.Name & "  Check_Karaoke"
'    Check_Karaoke = False
'End Function

Public Sub Update_VIP_Percent()
On Error GoTo Handle
    Dim rsLocation As New ADODB.Recordset
    If cnData.State = 1 Then Set rsLocation = Open_Table(cnData, "Table_Diagram_Sections")
    With rsLocation
    
        If Not .EOF And .RecordCount > 0 Then .MoveFirst
        .Find "Location_ID=02", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If Format(Now, "HH:mm:ss") >= Format("16:59:59", "HH:mm:ss") Then
                    .Fields("PriceRate") = "33.33"
                ElseIf Format(Now, "HH:mm:ss") <= Format("06:59:59", "HH:mm:ss") Then
                    .Fields("PriceRate") = "25"
                End If
                .Update
            End If
        
    End With
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " Update_VIP_Percent"
End Sub

Public Function Get_First_Section() As String
On Error GoTo Handle
    Dim S As String
    Dim rsTable As New ADODB.Recordset
    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881")
    Set rsTable = OpenCriticalTable("Select Min(Location_ID) as minID from Table_Diagram_Sections", cnData)
    If Not rsTable.EOF And rsTable.RecordCount > 0 Then
        S = rsTable.Fields("MinID")
    Else
        S = "01"
    End If
Get_First_Section = S
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " Get_First_Section"
End Function

Public Sub Update_Table_In_Order_Details(Bill_Source As Double, Bill_Des As Double, TranferType As Integer, Optional Source_Table As String, Optional Des_Table As String)
On Error GoTo Handle
    Dim rsKitchen_Master As New ADODB.Recordset
    Dim rsKitchen_Items As New ADODB.Recordset
    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881")
    Set rsKitchen_Master = Open_Table(cnData, "Kitchen_Order_Master")
    Set rsKitchen_Items = Open_Table(cnData, "Kitchen_Order_Items")
    With rsKitchen_Master
        .Find "Invoice_Number=" & Bill_Des, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If TranferType = 2 Then
                Dim rsTempItem As New ADODB.Recordset
                Dim rsMaxLine_Items_Des As New ADODB.Recordset
                Dim MaxLine As Integer
                    Set rsMaxLine_Items_Des = OpenCriticalTable("select Max(LineNum)as MaxLine from Kitchen_Order_Items where invoice_Number=" & Bill_Des, cnData)
                    Set rsTempItem = OpenCriticalTable("select * from Kitchen_Order_Items where invoice_Number=" & Bill_Source, cnData)
                    If Not rsMaxLine_Items_Des.EOF Then
                        MaxLine = rsMaxLine_Items_Des.Fields("MaxLine") + 1
                    End If
                    With rsTempItem
                        Do While Not .EOF
                            With rsKitchen_Items
                                .addNew
                                .Fields("Invoice_Number") = Bill_Des
                                .Fields("ItemName") = rsTempItem.Fields("ItemName")
                                .Fields("Quantity") = rsTempItem.Fields("Quantity")
                                .Fields("Price") = rsTempItem.Fields("Price")
                                .Fields("LineNum") = MaxLine
                                .Fields("Kit_Desc") = rsTempItem.Fields("Kit_Desc")
                                .Fields("Printer_ID") = rsTempItem.Fields("Printer_ID")
                                .Fields("Send_KP_Date") = rsTempItem.Fields("Send_KP_Date")
                                .Fields("Send_KP_Time") = rsTempItem.Fields("Send_KP_Time")
                                .Update
                            End With
                        .MoveNext
                        MaxLine = MaxLine + 1
                        Loop
                    End With
                    cnData.Execute "delete * from Kitchen_Order_Items where Invoice_Number=" & Bill_Source
                    cnData.Execute "Delete * from Kitchen_Order_Master where Invoice_Number=" & Bill_Source
                End If
                '.Fields("Invoice_Number") = Bill_Des
                .Fields("Table_ID") = Des_Table
                .Update
                
            End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " Update_Table_In_Order_Details"
End Sub

Public Function Check_System_KP(S As String) As Boolean
On Error GoTo Handle
    Dim rsSys As New ADODB.Recordset
    Dim Used As Boolean
    Dim i As Integer
    
    Set rsSys = OpenCriticalTable("select * from SystemFlag where SF='" & S & "'", cnData)
    With rsSys
        If Not .EOF Then
                For i = 1 To 8
                    ReDim Preserve arrPriterKP(i)
                    If ArrayFlag(.Fields("Data"), i) = 1 Then
                        Used = True
                        arrPriterKP(i) = 1
                    End If
                Next i
        End If
        
    End With
    Check_System_KP = Used
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Check_System_KP"
    Check_System_KP = False
End Function

Public Sub Update_Tranfer(BillOrg As Double, billDes As Double, LocationOrg As String, LocationDes As String, TableOrg As String, TableDes As String, State As Integer)
On Error GoTo Handle
    Dim rsTranferTable As New ADODB.Recordset
    Set rsTranferTable = Open_Table(cnData, "Tranfer_Joint_table")
    With rsTranferTable
        .addNew
        .Fields("Org_bill") = BillOrg
        .Fields("Des_bill") = billDes
        .Fields("Org_Location") = LocationOrg
        .Fields("Des_Location") = LocationDes
        .Fields("Org_Table") = TableOrg
        .Fields("Des_Table") = TableDes
        .Fields("Cashier_ID") = UserID
        .Fields("DateTime") = DateDefault & Format(Now, "HH:mm:ss")
        .Fields("State") = State
        .Update
        
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "  Update_Tranfer"
End Sub

Public Sub Print_Tranfer_Bill(BillTranfer As Double, printer_Name As String)
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim iReport As CRAXDDRT.Report
    If cnData.State = 0 Then
        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881")
    End If
    SQL = "select * from Tranfer_Joint_table where Des_bill =" & BillTranfer
    Set crTranfer = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crTranfer
        .Database.AddADOCommand cnData, cmd
        .billDes.SetUnboundFieldSource "{ado.Des_bill}"
        .BillOrg.SetUnboundFieldSource "{ado.Org_bill}"
        .LocationDes.SetUnboundFieldSource "{ado.Des_Location}"
        .location.SetUnboundFieldSource "{ado.Org_Location}"
        .TableDes.SetUnboundFieldSource "{ado.Des_Table}"
        .TableOrg.SetUnboundFieldSource "{ado.Org_Table}"
        .CashierID.SetUnboundFieldSource "{ado.Cashier_ID}"
    End With
    Set iReport = crTranfer
    With frmShowSendKP
        .Report = crTranfer
        .Get_ID = "01"
        .GetPrinter = printer_Name
        .Show vbModal
    End With
Exit Sub
Handle:
Exit Sub
MsgBox Err.Number & Err.Description & Me.Name & " Print_Tranfer_Bill"
End Sub

Public Function Get_Right_Location(User_ID As String) As String
On Error GoTo Handle
Dim a As String
Dim rsCashier_Location As New ADODB.Recordset
    Set rsCashier_Location = Open_Table(cnData, "Stations")
    With rsCashier_Location
        If .RecordCount > 0 Then
            .MoveFirst
            .Find "Cashier_ID='" & User_ID & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                a = .Fields("Location")
            End If
        End If
    End With
    Get_Right_Location = a
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "  Get_Right_Location"
End Function

Public Sub Delete_Invoice_AmountLarger(ByVal amount As Double)
    On Error GoTo Handle
    Dim rsOrg As New ADODB.Recordset
    Set rsOrg = OpenCriticalTable("select * from Invoice_Totals where status='C' and Grand_Total>=" & amount & " and Synchronized= True and Invoice_Number<>0", cnData)
    If rsOrg.State <> 0 Then
        If rsOrg.RecordCount > 0 Then
            rsOrg.MoveFirst
        Else
            Exit Sub
        End If
    End If
    With rsOrg
        Do While Not .EOF
            cnData.Execute "Delete * from Items_Deleted where Invoice_Num=" & rsOrg.Fields("Invoice_Number")
            cnData.Execute "Delete * from Invoice_Itemized where Invoice_Number=" & rsOrg.Fields("Invoice_Number")
            cnData.Execute "Delete * from Invoice_Totals_Person_Mapping where Invoice_Number=" & rsOrg.Fields("Invoice_Number")
            cnData.Execute "Delete * from Kitchen_Order_Master where Invoice_Number=" & rsOrg.Fields("Invoice_Number")
            cnData.Execute "Delete * from Kitchen_Order_Items where Invoice_Number=" & rsOrg.Fields("Invoice_Number")
            cnData.Execute "Delete * from Invoice_Totals_Notes where Invoice_Number=" & rsOrg.Fields("Invoice_Number")
            cnData.Execute "Delete * from Invoice_Totals where Invoice_Number=" & rsOrg.Fields("Invoice_Number")
            cnData.Execute "Delete * from Kitchen_Order_Master where Invoice_Number=" & rsOrg.Fields("Invoice_Number")
            cnData.Execute "Delete * from Kitchen_Order_Items where Invoice_Number=" & rsOrg.Fields("Invoice_Number")
'            .Requery
        .MoveNext
        Loop
    End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " Delete_Invoice_AmountLarger"
End Sub
Public Sub gfBackup_Invoice_Kitchen_Order_Mast(cnBackup As Connection, cnOrg As Connection, Invoice_Num As Double)
On Error GoTo Handle
    Dim rsInvoice_Kitchen_Org As New ADODB.Recordset
    Dim rsInvoice_Kitchen_Des As New ADODB.Recordset
    Dim i As Integer
'    cnBackup.BeginTrans
'    cnOrg.BeginTrans
        Set rsInvoice_Kitchen_Org = OpenCriticalTable("Select * from Kitchen_Order_Master where Invoice_Number=" & Invoice_Num, cnOrg)
        Set rsInvoice_Kitchen_Des = Open_Table(cnBackup, "Kitchen_Order_Master")
        i = 0
        With rsInvoice_Kitchen_Org
            Do While Not .EOF
                With rsInvoice_Kitchen_Des
                    .addNew
                    .Fields("Invoice_Number") = Max_Invoice_Backup '+ i 'rsInvoice_Note_Org.Fields("Invoice_Number")
'                    .Fields("Invoice_Number") = rsInvoice_Per_Org.Fields("Invoice_Number")
                    .Fields("Station_ID") = rsInvoice_Kitchen_Org.Fields("Station_ID")
                    .Fields("Store_ID") = rsInvoice_Kitchen_Org.Fields("Store_ID")
                    .Fields("Cashier_ID") = rsInvoice_Kitchen_Org.Fields("Cashier_ID")
                    .Fields("Table_ID") = rsInvoice_Kitchen_Org.Fields("Table_ID")
                    .Update
'                    .Requery
                End With
            .MoveNext
            i = i + 1
            Loop
        End With
Exit Sub
Handle:
'    cnBackup.RollbackTrans
'    cnOrg.RollbackTrans
    MsgBox Err.Number & Err.Description & Me.Name & "gfBackup_Invoice_Kitchen_Order_Mast"
End Sub

Public Sub gfBackup_Invoice_Kitchen_Order_Items(cnBackup As Connection, cnOrg As Connection, Invoice_Num As Double)
On Error GoTo Handle
    Dim rsInvoice_Kitchen_Items_Org As New ADODB.Recordset
    Dim rsInvoice_Kitchen_Items_Des As New ADODB.Recordset
    Dim i As Integer
'    cnBackup.BeginTrans
'    cnOrg.BeginTrans
        Set rsInvoice_Kitchen_Items_Org = OpenCriticalTable("Select * from Kitchen_Order_Items where Invoice_Number=" & Invoice_Num, cnOrg)
        Set rsInvoice_Kitchen_Items_Des = Open_Table(cnBackup, "Kitchen_Order_Items")
        i = 0
        With rsInvoice_Kitchen_Items_Org
            Do While Not .EOF
                With rsInvoice_Kitchen_Items_Des
                    .addNew
                    .Fields("Invoice_Number") = Max_Invoice_Backup '+ i 'rsInvoice_Note_Org.Fields("Invoice_Number")
                    .Fields("ItemName") = rsInvoice_Kitchen_Items_Org.Fields("ItemName")
                    .Fields("ItemNum") = rsInvoice_Kitchen_Items_Org.Fields("ItemNum")
                    .Fields("Quantity") = rsInvoice_Kitchen_Items_Org.Fields("Quantity")
                    .Fields("Price") = rsInvoice_Kitchen_Items_Org.Fields("Price")
                    .Fields("LineNum") = rsInvoice_Kitchen_Items_Org.Fields("LineNum")
                    .Fields("Kit_Desc") = rsInvoice_Kitchen_Items_Org.Fields("Kit_Desc")
                    .Fields("Printer_ID") = rsInvoice_Kitchen_Items_Org.Fields("Printer_ID")
                    .Fields("Send_KP_Date") = rsInvoice_Kitchen_Items_Org.Fields("Send_KP_Date")
                    .Fields("Send_KP_Time") = rsInvoice_Kitchen_Items_Org.Fields("Send_KP_Time")
                    .Update
'                    .Requery
                End With
            .MoveNext
            i = i + 1
            Loop
        End With
Exit Sub
Handle:
'    cnBackup.RollbackTrans
'    cnOrg.RollbackTrans
    MsgBox Err.Number & Err.Description & Me.Name & "gfBackup_Invoice_Kitchen_Order_Items"
End Sub
Public Sub gfBackup_TablePlan(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsTable_Org As New ADODB.Recordset
    Dim rsTable_Des As New ADODB.Recordset
    
        lblTitle.Caption = "S¬ då bµn....."
        Delay (500)
        'Cap nhat Khu vuc
        Call gfBackup_Location(cnOrg, cnBackup)
        
        cnBackup.Execute "delete * from Table_Diagram"
        Set rsTable_Org = Open_Table(cnOrg, "Table_Diagram")
        Set rsTable_Des = Open_Table(cnBackup, "Table_Diagram")
        
        With rsTable_Org
            Do While Not .EOF
                With rsTable_Des
                    .addNew
                    .Fields("Store_ID") = rsTable_Org.Fields("Store_ID")
                    .Fields("Section_ID") = rsTable_Org.Fields("Section_ID")
                    .Fields("Table_Number") = rsTable_Org.Fields("Table_Number")
                    .Fields("ShapeType") = rsTable_Org.Fields("ShapeType")
                    .Fields("XPos") = rsTable_Org.Fields("XPos")
                    .Fields("YPos") = rsTable_Org.Fields("YPos")
                    .Fields("Height") = rsTable_Org.Fields("Height")
                    .Fields("Width") = rsTable_Org.Fields("Width")
                    .Fields("Cost_Center_Index") = rsTable_Org.Fields("Cost_Center_Index")
                    .Fields("NumSeats") = rsTable_Org.Fields("NumSeats")
                    .Update
                    .Requery
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
'    cnBackup.RollbackTrans
'    cnOrg.RollbackTrans
    MsgBox Err.Number & Err.Description & Me.Name & "gfBackup_Invoice_Kitchen_Order_Mast"
End Sub

Public Sub gfBackup_Location(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsLocation_Org As New ADODB.Recordset
    Dim rsLocation_Des As New ADODB.Recordset
    'Xoa Du lieu khu vuc
        cnBackup.Execute "Delete * from Table_Diagram_Sections"
        
        Set rsLocation_Org = Open_Table(cnOrg, "Table_Diagram_Sections")
        Set rsLocation_Des = Open_Table(cnBackup, "Table_Diagram_Sections")
        With rsLocation_Org
            Do While Not .EOF
                With rsLocation_Des
                    .addNew
                    .Fields("Store_ID") = rsLocation_Org.Fields("Store_ID")
                    .Fields("Location_ID") = rsLocation_Org.Fields("Location_ID")
                    .Fields("Section_ID") = rsLocation_Org.Fields("Section_ID")
                    .Fields("PriceRate") = rsLocation_Org.Fields("PriceRate")
                    .Fields("Used_Karaoke") = rsLocation_Org.Fields("Used_Karaoke")
                    .Fields("Price_Level") = rsLocation_Org.Fields("Price_Level")
                    .Update
                    .Requery
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "gfBackup_Location"
End Sub

Public Sub gfBackup_Group(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsGroup_Org As New ADODB.Recordset
    Dim rsGroup_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Nhãm hµng....."
        Delay (500)
        
        Set rsGroup_Org = Open_Table(cnOrg, "Departments")
        Set rsGroup_Des = Open_Table(cnBackup, "Departments")
        
        With rsGroup_Org
            Do While Not .EOF
                With rsGroup_Des
                    .Find "Dept_ID='" & rsGroup_Org.Fields("Dept_ID") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("Store_ID") = rsGroup_Org.Fields("Store_ID")
                            .Fields("Description") = rsGroup_Org.Fields("Description")
                            .Fields("MainGroup") = rsGroup_Org.Fields("MainGroup")
                            .Fields("F") = rsGroup_Org.Fields("F")
                            .Fields("ColorDept") = rsGroup_Org.Fields("ColorDept")
                            .Update
                        Else
                            .addNew
                            .Fields("Dept_ID") = rsGroup_Org.Fields("Dept_ID")
                            .Fields("Store_ID") = rsGroup_Org.Fields("Store_ID")
                            .Fields("Description") = rsGroup_Org.Fields("Description")
                            .Fields("MainGroup") = rsGroup_Org.Fields("MainGroup")
                            .Fields("F") = rsGroup_Org.Fields("F")
                            .Fields("ColorDept") = rsGroup_Org.Fields("ColorDept")
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "gfBackup_Group"
End Sub

Public Sub gfBackup_Items(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsItems_Org As New ADODB.Recordset
    Dim rsItems_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Danh môc hµng....."
        Delay (500)
        
        Set rsItems_Org = Open_Table(cnOrg, "Inventory")
        Set rsItems_Des = Open_Table(cnBackup, "Inventory")
        With rsItems_Org
            Do While Not .EOF
                With rsItems_Des
                    .Find "ItemNum='" & rsItems_Org.Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("ItemName") = rsItems_Org.Fields("ItemName")
                            .Fields("Dept_ID") = rsItems_Org.Fields("Dept_ID")
                            .Fields("Std_Price1") = rsItems_Org.Fields("Std_Price1")
                            .Fields("Std_Price2") = rsItems_Org.Fields("Std_Price2")
                            .Fields("Std_Price3") = rsItems_Org.Fields("Std_Price3")
                            .Fields("HH_Price1") = rsItems_Org.Fields("HH_Price1")
                            .Fields("HH_Price2") = rsItems_Org.Fields("HH_Price2")
                            .Fields("HH_Price3") = rsItems_Org.Fields("HH_Price3")
                            .Fields("EV_Price1") = rsItems_Org.Fields("EV_Price1")
                            .Fields("EV_Price2") = rsItems_Org.Fields("EV_Price2")
                            .Fields("EV_Price3") = rsItems_Org.Fields("EV_Price3")
                            '.Fields("LimitPrice") = rsItems_Org.Fields("LimitPrice")
                            .Fields("Unit") = rsItems_Org.Fields("Unit")
                            .Fields("Minstock") = rsItems_Org.Fields("Minstock")
                            .Fields("Modify_Number") = rsItems_Org.Fields("Modify_Number")
                            .Fields("F1") = rsItems_Org.Fields("F1")
                            .Fields("F2") = rsItems_Org.Fields("F2")
                            .Fields("F3") = rsItems_Org.Fields("F3")
                            .Fields("F4") = rsItems_Org.Fields("F4")
                            .Fields("F5") = rsItems_Org.Fields("F5")
                            .Fields("Date_Created") = Date
                            .Fields("Picture") = rsItems_Org.Fields("Picture")
                            .Fields("Print_On_Receipt") = rsItems_Org.Fields("Print_On_Receipt")
                            .Fields("Store_ID") = Store_ID
                            .Update
                        Else
                            .addNew
                            .Fields("ItemNum") = rsItems_Org.Fields("ItemNum")
                            .Fields("ItemName") = rsItems_Org.Fields("ItemName")
                            .Fields("Dept_ID") = rsItems_Org.Fields("Dept_ID")
                            .Fields("Std_Price1") = rsItems_Org.Fields("Std_Price1")
                            .Fields("Std_Price2") = rsItems_Org.Fields("Std_Price2")
                            .Fields("Std_Price3") = rsItems_Org.Fields("Std_Price3")
                            .Fields("HH_Price1") = rsItems_Org.Fields("HH_Price1")
                            .Fields("HH_Price2") = rsItems_Org.Fields("HH_Price2")
                            .Fields("HH_Price3") = rsItems_Org.Fields("HH_Price3")
                            .Fields("EV_Price1") = rsItems_Org.Fields("EV_Price1")
                            .Fields("EV_Price2") = rsItems_Org.Fields("EV_Price2")
                            .Fields("EV_Price3") = rsItems_Org.Fields("EV_Price3")
                            .Fields("LimitPrice") = rsItems_Org.Fields("LimitPrice")
                            .Fields("Unit") = rsItems_Org.Fields("Unit")
                            .Fields("Minstock") = rsItems_Org.Fields("Minstock")
                            .Fields("Modify_Number") = rsItems_Org.Fields("Modify_Number")
                            .Fields("F1") = rsItems_Org.Fields("F1")
                            .Fields("F2") = rsItems_Org.Fields("F2")
                            .Fields("F3") = rsItems_Org.Fields("F3")
                            .Fields("F4") = rsItems_Org.Fields("F4")
                            .Fields("F5") = rsItems_Org.Fields("F5")
                            .Fields("Date_Created") = Date
                            .Fields("Picture") = rsItems_Org.Fields("Picture")
                            .Fields("Print_On_Receipt") = rsItems_Org.Fields("Print_On_Receipt")
                            .Fields("Store_ID") = Store_ID
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "gfBackup_Items"
End Sub

Public Sub gfBackup_Customer(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsCust_Org As New ADODB.Recordset
    Dim rsCust_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Kh¸ch hµng....."
        Delay (500)
        Set rsCust_Org = Open_Table(cnOrg, "Customer")
        Set rsCust_Des = Open_Table(cnBackup, "Customer")
        
        With rsCust_Org
            Do While Not .EOF
                With rsCust_Des
                    .Find "CustNum='" & rsCust_Org.Fields("CustNum") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("CustName") = rsCust_Org.Fields("CustName")
                            .Fields("Company") = rsCust_Org.Fields("Company")
                            .Fields("Address") = rsCust_Org.Fields("Address")
                            .Fields("Phone") = rsCust_Org.Fields("Phone")
                            .Fields("Fax") = rsCust_Org.Fields("Fax")
                            .Fields("Discount") = rsCust_Org.Fields("Discount")
                            .Fields("TaxCode") = rsCust_Org.Fields("TaxCode")
                            .Fields("AccountNo") = rsCust_Org.Fields("AccountNo")
                            .Fields("Acct_Open_Date") = rsCust_Org.Fields("Acct_Open_Date")
                            .Fields("Acct_Close_Date") = rsCust_Org.Fields("Acct_Close_Date")
                            .Fields("Acct_Balance") = rsCust_Org.Fields("Acct_Balance")
                            .Fields("Cashier") = rsCust_Org.Fields("Cashier")
                            .Fields("Acct_Max_Balance") = rsCust_Org.Fields("Acct_Max_Balance")
                            .Fields("Birthday") = rsCust_Org.Fields("Birthday")
'                            .Fields("Last_Visit") = rsCust_Org.Fields("Last_Visit")
'                            .Fields("Tax_Rate_ID") = rsCust_Org.Fields("Tax_Rate_ID")
                            .Fields("Point") = rsCust_Org.Fields("Point")
                            .Update
                        Else
                            .addNew
                            .Fields("CustNum") = rsCust_Org.Fields("CustNum")
                            .Fields("CustName") = rsCust_Org.Fields("CustName")
                            .Fields("Company") = rsCust_Org.Fields("Company")
                            .Fields("Address") = rsCust_Org.Fields("Address")
                            .Fields("Phone") = rsCust_Org.Fields("Phone")
                            .Fields("Fax") = rsCust_Org.Fields("Fax")
                            .Fields("Discount") = rsCust_Org.Fields("Discount")
                            .Fields("TaxCode") = rsCust_Org.Fields("TaxCode")
                            .Fields("AccountNo") = rsCust_Org.Fields("AccountNo")
                            .Fields("Acct_Open_Date") = rsCust_Org.Fields("Acct_Open_Date")
                            .Fields("Acct_Close_Date") = rsCust_Org.Fields("Acct_Close_Date")
                            .Fields("Acct_Balance") = rsCust_Org.Fields("Acct_Balance")
                            .Fields("Cashier") = rsCust_Org.Fields("Cashier")
                            .Fields("Acct_Max_Balance") = rsCust_Org.Fields("Acct_Max_Balance")
                            .Fields("Birthday") = rsCust_Org.Fields("Birthday")
'                            .Fields("Last_Visit") = rsCust_Org.Fields("Last_Visit")
'                            .Fields("Tax_Rate_ID") = rsCust_Org.Fields("Tax_Rate_ID")
                            .Fields("Point") = rsCust_Org.Fields("Point")
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "gfBackup_Custoer"
End Sub

Public Sub gfBackup_Vendor(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsVendor_Org As New ADODB.Recordset
    Dim rsVendor_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Kh¸ch hµng....."
        Delay (500)
        Set rsVendor_Org = Open_Table(cnOrg, "Vendors")
        Set rsVendor_Des = Open_Table(cnBackup, "Vendors")
        
        With rsVendor_Org
            Do While Not .EOF
                With rsVendor_Des
                    .Find "Vendor_Number='" & rsVendor_Org.Fields("Vendor_Number") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("Vendor_Name") = rsVendor_Org.Fields("Vendor_Name")
                            .Fields("Company") = rsVendor_Org.Fields("Company")
                            .Fields("Address_1") = rsVendor_Org.Fields("Address_1")
                            .Fields("Address_2") = rsVendor_Org.Fields("Address_2")
                            .Fields("Phone") = rsVendor_Org.Fields("Phone")
                            .Fields("Fax") = rsVendor_Org.Fields("Fax")
                            .Fields("Vendor_Tax_ID") = rsVendor_Org.Fields("Vendor_Tax_ID")
                            .Fields("Vendor_AccNo") = rsVendor_Org.Fields("Vendor_AccNo")
                            .Fields("Email") = rsVendor_Org.Fields("Email")
                            .Fields("Website") = rsVendor_Org.Fields("Website")
                            .Update
                        Else
                            .addNew
                            .Fields("Vendor_Number") = rsVendor_Org.Fields("Vendor_Number")
                            .Fields("Vendor_Name") = rsVendor_Org.Fields("Vendor_Name")
                            .Fields("Company") = rsVendor_Org.Fields("Company")
                            .Fields("Address_1") = rsVendor_Org.Fields("Address_1")
                            .Fields("Address_2") = rsVendor_Org.Fields("Address_2")
                            .Fields("Phone") = rsVendor_Org.Fields("Phone")
                            .Fields("Fax") = rsVendor_Org.Fields("Fax")
                            .Fields("Vendor_Tax_ID") = rsVendor_Org.Fields("Vendor_Tax_ID")
                            .Fields("Vendor_AccNo") = rsVendor_Org.Fields("Vendor_AccNo")
                            .Fields("Email") = rsVendor_Org.Fields("Email")
                            .Fields("Website") = rsVendor_Org.Fields("Website")
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "gfBackup_Vendor"
End Sub

Public Sub gfBackup_Thu(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsThu_Org As New ADODB.Recordset
    Dim rsThu_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Danh môc Thu....."
        Delay (500)
        Set rsThu_Org = Open_Table(cnOrg, "Receipt")
        Set rsThu_Des = Open_Table(cnBackup, "Receipt")
        
        With rsThu_Org
            Do While Not .EOF
                With rsThu_Des
                    .Find "MaThu='" & rsThu_Org.Fields("MaThu") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("DienGiai") = rsThu_Org.Fields("DienGiai")
                            .Update
                        Else
                            .addNew
                            .Fields("MaThu") = rsThu_Org.Fields("MaThu")
                            .Fields("DienGiai") = rsThu_Org.Fields("DienGiai")
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "Danh muc thu"
End Sub
Public Sub gfBackup_Chi(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsChi_Org As New ADODB.Recordset
    Dim rsChi_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Danh môc Chi....."
        Delay (500)
        Set rsChi_Org = Open_Table(cnOrg, "Expense")
        Set rsChi_Des = Open_Table(cnBackup, "Expense")
        
        With rsChi_Org
            Do While Not .EOF
                With rsChi_Des
                    .Find "MaChi='" & rsChi_Org.Fields("MaChi") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("DienGiai") = rsChi_Org.Fields("DienGiai")
                            .Update
                        Else
                            .addNew
                            .Fields("MaChi") = rsChi_Org.Fields("MaChi")
                            .Fields("DienGiai") = rsChi_Org.Fields("DienGiai")
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "Danh muc chi"
End Sub

Public Sub gfBackup_Thutien(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsThu_Org As New ADODB.Recordset
    Dim rsThu_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Thu tiÒn....."
        Delay (500)
        Set rsThu_Org = Open_Table(cnOrg, "Income")
        Set rsThu_Des = Open_Table(cnBackup, "Income")
        
        With rsThu_Org
            Do While Not .EOF
                With rsThu_Des
                    .Find "ID='" & rsThu_Org.Fields("ID") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("Store_ID") = rsThu_Org.Fields("Store_ID")
                            .Fields("Cashier_ID") = rsThu_Org.Fields("Cashier_ID")
                            .Fields("DateTime") = rsThu_Org.Fields("DateTime")
                            .Fields("Customer_ID") = rsThu_Org.Fields("Customer_ID")
                            .Fields("Receipt_ID") = rsThu_Org.Fields("Receipt_ID")
                            .Fields("Reciever_Name") = rsThu_Org.Fields("Reciever_Name")
                            .Fields("Division") = rsThu_Org.Fields("Division")
                            .Fields("Amount") = rsThu_Org.Fields("Amount")
                            .Fields("Description") = rsThu_Org.Fields("Description")
                            .Fields("Payment_Method") = rsThu_Org.Fields("Payment_Method")
                            .Update
                        Else
                            .addNew
                            .Fields("ID") = rsThu_Org.Fields("ID")
                            .Fields("Store_ID") = rsThu_Org.Fields("Store_ID")
                            .Fields("Cashier_ID") = rsThu_Org.Fields("Cashier_ID")
                            .Fields("DateTime") = rsThu_Org.Fields("DateTime")
                            .Fields("Customer_ID") = rsThu_Org.Fields("Customer_ID")
                            .Fields("Receipt_ID") = rsThu_Org.Fields("Receipt_ID")
                            .Fields("Reciever_Name") = rsThu_Org.Fields("Reciever_Name")
                            .Fields("Division") = rsThu_Org.Fields("Division")
                            .Fields("Amount") = rsThu_Org.Fields("Amount")
                            .Fields("Description") = rsThu_Org.Fields("Description")
                            .Fields("Payment_Method") = rsThu_Org.Fields("Payment_Method")
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "Thu tien"
End Sub

Public Sub gfBackup_Chitien(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsChi_Org As New ADODB.Recordset
    Dim rsChi_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Chi tiÒn....."
        Delay (500)
        Set rsChi_Org = Open_Table(cnOrg, "PayOuts")
        Set rsChi_Des = Open_Table(cnBackup, "PayOuts")
        
        With rsChi_Org
            Do While Not .EOF
                With rsChi_Des
                    .Find "ID='" & rsChi_Org.Fields("ID") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("Store_ID") = rsChi_Org.Fields("Store_ID")
                            .Fields("Cashier_ID") = rsChi_Org.Fields("Cashier_ID")
                            .Fields("DateTime") = rsChi_Org.Fields("DateTime")
                            .Fields("Vendor_Number") = rsChi_Org.Fields("Vendor_Number")
                            .Fields("Amount") = rsChi_Org.Fields("Amount")
                            .Fields("Description") = rsChi_Org.Fields("Description")
                            .Fields("Payment_Method") = rsChi_Org.Fields("Payment_Method")
                            .Fields("Expense_ID") = rsChi_Org.Fields("Expense_ID")
                            .Fields("Recieve_Name") = rsChi_Org.Fields("Recieve_Name")
                            .Fields("Division") = rsChi_Org.Fields("Division")
                            .Update
                        Else
                            .addNew
                            .Fields("ID") = rsChi_Org.Fields("ID")
                            .Fields("Store_ID") = rsChi_Org.Fields("Store_ID")
                            .Fields("Cashier_ID") = rsChi_Org.Fields("Cashier_ID")
                            .Fields("DateTime") = rsChi_Org.Fields("DateTime")
                            .Fields("Vendor_Number") = rsChi_Org.Fields("Vendor_Number")
                            .Fields("Amount") = rsChi_Org.Fields("Amount")
                            .Fields("Description") = rsChi_Org.Fields("Description")
                            .Fields("Payment_Method") = rsChi_Org.Fields("Payment_Method")
                            .Fields("Expense_ID") = rsChi_Org.Fields("Expense_ID")
                            .Fields("Recieve_Name") = rsChi_Org.Fields("Recieve_Name")
                            .Fields("Division") = rsChi_Org.Fields("Division")
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "Chi tien"
End Sub


Public Sub gfBackup_DMInOut(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Call gfBackup_Chi(cnOrg, cnBackup)
    Call gfBackup_Thu(cnOrg, cnBackup)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " gfBackup_DMInOut"

End Sub


Public Sub gfBackup_InOut(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Call gfBackup_Chitien(cnOrg, cnBackup)
    Call gfBackup_Thutien(cnOrg, cnBackup)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " gfBackup_InOut"

End Sub



Private Sub TextFly_Timer()
    Countdown = Countdown + 1
    If Countdown = 300 Then
        Call Backup_DB
        Countdown = 0
    End If
End Sub

Private Sub Timer2_Timer()
'    Call LoadTable(Sec_ID)
End Sub
