VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTablePlan 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   10770
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11400
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTablePlan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10770
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame frabk 
      Height          =   11055
      Left            =   -120
      TabIndex        =   0
      Top             =   -120
      Width           =   20535
      Begin VB.PictureBox picCom 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   11160
         ScaleHeight     =   1215
         ScaleWidth      =   4455
         TabIndex        =   25
         Top             =   120
         Width           =   4455
         Begin VB.Label lblPhone 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "ÑT: 0918.655.887 (24/24) "
            BeginProperty Font 
               Name            =   "VNI-Times"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   28
            Top             =   720
            Width           =   4335
         End
         Begin VB.Label lblAdd 
            Alignment       =   2  'Center
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "565/6 Bình Thôùi, P.10, Q.11, Tp.HCM"
            BeginProperty Font 
               Name            =   "VNI-Times"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   27
            Top             =   405
            Width           =   4335
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
            Left            =   0
            TabIndex        =   26
            Top             =   0
            Width           =   4335
         End
      End
      Begin VB.Frame fraTable 
         BackColor       =   &H00808080&
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
         ForeColor       =   &H00FF8080&
         Height          =   8805
         Left            =   0
         TabIndex        =   13
         Top             =   1320
         Width           =   13245
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   120
            Top             =   1920
         End
         Begin VB.PictureBox picWait 
            BackColor       =   &H00C00000&
            Height          =   420
            Left            =   4920
            ScaleHeight     =   360
            ScaleWidth      =   4905
            TabIndex        =   14
            Top             =   4440
            Visible         =   0   'False
            Width           =   4965
            Begin MSComctlLib.ProgressBar probarWait 
               Height          =   390
               Left            =   0
               TabIndex        =   15
               Top             =   0
               Width           =   4905
               _ExtentX        =   8652
               _ExtentY        =   688
               _Version        =   393216
               Appearance      =   1
            End
         End
         Begin VB.Shape Shape1 
            BorderColor     =   &H00C0C0FF&
            BorderWidth     =   3
            Height          =   1035
            Index           =   0
            Left            =   120
            Top             =   210
            Visible         =   0   'False
            Width           =   1755
         End
         Begin VB.Label lblTable 
            Alignment       =   2  'Center
            BackColor       =   &H00FF8080&
            BackStyle       =   0  'Transparent
            Caption         =   "#1"
            BeginProperty Font 
               Name            =   ".VnArialH"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFC0&
            Height          =   975
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Visible         =   0   'False
            Width           =   1635
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblTime 
            Alignment       =   2  'Center
            Caption         =   "Label2"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   17
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1215
            Left            =   3600
            TabIndex        =   16
            Top             =   4080
            Visible         =   0   'False
            Width           =   8655
         End
      End
      Begin VB.Timer Timer2 
         Interval        =   10000
         Left            =   12480
         Top             =   10200
      End
      Begin VB.PictureBox pic1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   9975
         Left            =   13200
         ScaleHeight     =   9975
         ScaleWidth      =   2295
         TabIndex        =   3
         Top             =   1320
         Width           =   2295
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00C0C0C0&
            Height          =   4575
            Left            =   -120
            ScaleHeight     =   4515
            ScaleWidth      =   2355
            TabIndex        =   9
            Top             =   5400
            Width           =   2415
            Begin MSDataGridLib.DataGrid dtgReserve 
               Height          =   4335
               Left            =   120
               TabIndex        =   10
               Top             =   0
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   7646
               _Version        =   393216
               AllowUpdate     =   0   'False
               BackColor       =   16777215
               BorderStyle     =   0
               HeadLines       =   1
               RowHeight       =   15
               WrapCellPointer =   -1  'True
               FormatLocked    =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   ".VnArial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   ".VnArial"
                  Size            =   8.25
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
                     DividerStyle    =   5
                  EndProperty
                  BeginProperty Column01 
                  EndProperty
               EndProperty
            End
         End
         Begin prjTouchScreen.MyButton cmdCash_Open 
            Height          =   735
            Left            =   120
            TabIndex        =   4
            Top             =   3360
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1296
            BTYPE           =   6
            TX              =   "         Më kÐt          (F7)"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16744576
            BCOLO           =   12632256
            FCOL            =   16777215
            FCOLO           =   16711680
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
         Begin prjTouchScreen.MyButton cmdPrintedBill 
            Height          =   735
            Left            =   120
            TabIndex        =   5
            Top             =   2520
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1296
            BTYPE           =   6
            TX              =   " In l¹i hãa ®¬n   (F6)"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16744576
            BCOLO           =   12632256
            FCOL            =   16777215
            FCOLO           =   16711680
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
         Begin prjTouchScreen.MyButton cmdReserve 
            Height          =   735
            Left            =   120
            TabIndex        =   6
            Top             =   1680
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1296
            BTYPE           =   6
            TX              =   "          §Æt tiÖc          (F5)"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16744576
            BCOLO           =   12632256
            FCOL            =   16777215
            FCOLO           =   16711680
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
         Begin prjTouchScreen.MyButton cmdPayment 
            Height          =   735
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1296
            BTYPE           =   6
            TX              =   "      Thanh to¸n         (F4)"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16744576
            BCOLO           =   12632256
            FCOL            =   16777215
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmTablePlan.frx":1123E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdPrint_Receipt 
            Height          =   735
            Left            =   120
            TabIndex        =   8
            Top             =   0
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1296
            BTYPE           =   6
            TX              =   "      In Hãa §¬n        (F3)"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   16744576
            BCOLO           =   12632256
            FCOL            =   16777215
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmTablePlan.frx":1125A
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
            Height          =   780
            Left            =   120
            TabIndex        =   11
            Top             =   4200
            Width           =   2040
            _ExtentX        =   3598
            _ExtentY        =   1376
            BTYPE           =   6
            TX              =   "§ãng"
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
            BCOL            =   255
            BCOLO           =   255
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
         Begin VB.Label lblnote 
            BackStyle       =   0  'Transparent
            Caption         =   "DS tiÖc trong ngµy"
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
            TabIndex        =   12
            Top             =   5040
            Width           =   2055
         End
      End
      Begin VB.Timer Timer3 
         Left            =   9360
         Top             =   480
      End
      Begin VB.Timer TimerBackColor 
         Left            =   10320
         Top             =   120
      End
      Begin prjTouchScreen.MyButton cmdError 
         Height          =   975
         Left            =   6480
         TabIndex        =   1
         Tag             =   "L6"
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1720
         BTYPE           =   6
         TX              =   "MyButton1"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   12632256
         FCOL            =   16777215
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTablePlan.frx":11292
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdOption 
         Height          =   975
         Left            =   4320
         TabIndex        =   2
         Tag             =   "L4"
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1720
         BTYPE           =   6
         TX              =   "MyButton1"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   12632256
         FCOL            =   16777215
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTablePlan.frx":112AE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdTableOption 
         Height          =   225
         Left            =   9000
         TabIndex        =   19
         Tag             =   "L7"
         Top             =   960
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   397
         BTYPE           =   6
         TX              =   "  &Bµn míi  ph¸t sinh"
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
         Height          =   765
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   10200
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1349
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
      Begin prjTouchScreen.MyButton cmdExittoLogin 
         Height          =   975
         Left            =   120
         TabIndex        =   21
         Tag             =   "L2"
         Top             =   240
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1720
         BTYPE           =   6
         TX              =   "MyButton1"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   12632256
         FCOL            =   16777215
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTablePlan.frx":11302
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
         Height          =   975
         Left            =   2160
         TabIndex        =   22
         Tag             =   "L3"
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1720
         BTYPE           =   6
         TX              =   "MyButton1"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   12632256
         FCOL            =   16777215
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTablePlan.frx":1131E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
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
         Left            =   8280
         TabIndex        =   24
         Tag             =   "L8"
         Top             =   240
         Width           =   2745
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
         Left            =   8280
         TabIndex        =   23
         Top             =   720
         Width           =   2745
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Xö lý bµn tiÖc"
      Visible         =   0   'False
      Begin VB.Menu mnuSelect 
         Caption         =   "Gäi bµn"
      End
      Begin VB.Menu mnuDetails 
         Caption         =   "Xem chi tiÕt"
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
Dim rscust As New ADODB.Recordset
Dim rsinvoice_hold As New ADODB.Recordset
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
Dim Max_Invoice_Backup As Integer
Dim isclick As Boolean
Dim Table_ID As String
Dim AmountBackup As Double
Dim arrPriterKP() As String
Dim Countdown As Integer
Dim TimerBackup As Integer
Dim Option_call As Integer
Dim PRLEVEL, Sercharge, PRICE_RATE, TimeLevel As Integer
Dim Section_Select As Boolean
Dim light As Integer
Dim Reserve_Code As String
Dim Amount_Reserve As Double
Dim rsOrdered As New ADODB.Recordset
Dim TimeOrder As String
Dim kp_item As String

Private Sub cmdBufferPrint_Click()
If StateCall = 5 Then
        StateCall = 1
        TimerBackColor.Interval = 0
        cmdPrint_Receipt.BackColor = &HFFC0C0
    Else
        StateCall = 5
        TimerBackColor.Interval = 1000
    End If
End Sub


Private Sub cmdClose_Click()
    'Call Backup_DB_Store
    Set cnData = Nothing
    End
End Sub

Private Sub cmdError_Click()
On Error GoTo Handle
Dim strSql As String
    strSql = "delete  from Invoice_OnHold where Invoice_Number not in (Select Invoice_Number from Invoice_Totals)"
    If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    cnData.Execute strSql
    cnData.Execute "Update Invoice_Totals set InvoiceNotesUsed =0"
    MsgBox "Bµn lçi ®· xö lý, vui lßng chän bµn cÇn më", vbInformation
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   cmdError_Click"
End Sub

Private Sub cmdEdit_Click()
 On Error GoTo Handle
   ' Call Open_File
    ''Print #fFile, "ThiÕt lËp s¬ ®å bµn" & vbTab & Now & vbTab  & ":" & userName
 
    With frmPassword
        .FormActionKey = "EditTable"
        .Show vbModal
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & "File nµy ®· më"
    Close #fFile
End Sub

Private Sub cmdExittoLogin_Click()
    Set rsSection = Nothing
    Set rsTable = Nothing
    Set rsTranfer = Nothing
'    Set cnData = Nothing
'    Call gsDELETE_TMP_FILE
    'Print #fFile, "Tho¸t ca" & vbTab & Now & vbTab & ":" & userName
'    If Dir(WorkingFolder & "\Database.mdb" & Format(Now, "dd-MM-yyyy"), vbDirectory) <> "" Then
'        Kill WorkingFolder & "\Database.mdb" & Format(Now, "dd-MM-yyyy")
'    End If
    'Call Backup_DB_Store
    Unload Me
    With frmLogin
        .Me_State = 1
        .Show vbModal
    End With
End Sub

'Public Sub Backup_DB()
'On Error GoTo Handle
'    Dim fso As New FileSystemObject
'    fso.CopyFile WorkingFolder & "\Database.mdb", WorkingFolder & "\Database.mdb" & Format(Now, "dd-MM-yyyy"), True
'Exit Sub
'Handle:
'MsgBox Err.Number & Err.Description & Me.name & " Backup_DB"
'End Sub
'
'Public Sub Backup_DB_Store()
'On Error GoTo Handle
'    Dim fso As New FileSystemObject
'    If Dir(WorkingFolder & "\Store DB", vbDirectory) = "" Then
'        MkDir WorkingFolder & "\Store DB"
'    End If
'    fso.CopyFile WorkingFolder & "\Database.mdb", WorkingFolder & "\Store DB" & "\Database.mdb" & Format(Now, "dd-MM-yyyy"), True
'Exit Sub
'Handle:
'MsgBox Err.Number & Err.Description & Me.name & " Backup_DB"
'End Sub


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


Private Sub cmdOption_Click()
On Error GoTo Handle
    If ReceiptType = "58" Then
            frmSetup_Simple.Show vbModal
    Else
        frmSetup.Show vbModal
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description
End Sub

Private Sub cmdPrint_Buffer_Click()
    StateCall = 5
End Sub

Private Sub cmdPayment_Click()
On Error GoTo Handle
Dim ID As String
    If Not Get_Right(UserID, "payment") Then
            With frmPassword
                .FormActionKey = "Others"
                .Show vbModal
                ID = .return_Pass
                If Not .Return_right Then Exit Sub
            End With
            If Get_Right(ID, "payment") Then
                GoTo OK
            Else
                Exit Sub
            End If
    Else
        GoTo OK
    End If
OK:
  If StateCall = 6 Then
        StateCall = 1
        TimerBackColor.Interval = 0
        cmdPayment.BackColor = &HFFC0C0
        Timer3.Interval = 0
    Else
        StateCall = 6
        TimerBackColor.Interval = 1000
        cmdPayment.BackColor = &H80FF&
        Timer3.Interval = 500
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & "  - cmdPayment_Click"
End Sub

Private Sub cmdPrint_Receipt_Click()
On Error GoTo Handle
Dim ID As String
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "bufferPrint") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "bufferPrint") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
If StateCall = 5 Then
        StateCall = 1
        TimerBackColor.Interval = 0
        cmdPrint_Receipt.BackColor = &HFF8080
        Timer3.Interval = 0
    Else
        StateCall = 5
        TimerBackColor.Interval = 1000
        cmdPrint_Receipt.BackColor = &HFF8080
        Timer3.Interval = 500
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   cmdPrint_Receipt_Click"
End Sub

Private Sub cmdPrintedBill_Click()
On Error GoTo Handle
    frmPreviewBill.Show vbModal
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdPrintedBill_Click"
End Sub

Private Sub cmdReserve_Click()
    frmReservered.Show vbModal
End Sub

Private Sub cmdCash_Open_Click()
    Call OpenPrinterCashDraw(GetSettingStr("Receipt", "Receipt_DeviceName", True, myIniFile))
End Sub

Private Sub DataGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    PopupMenu mnuEdit, 0
End If
End Sub



Private Sub dtgReserve_Click()
On Error GoTo Handle
    Reserve_Code = dtgReserve.Columns(0).Value
    Table_ID = dtgReserve.Columns(4).Value
    Amount_Reserve = dtgReserve.Columns(5).Value
     Sec_ID = Val("0" & dtgReserve.Columns(6).Value)
    cmdSection_Click (Sec_ID)
    Exit Sub
Handle:
If Err.Number = 7005 Then
    MsgBox "Danh s¸ch rçng"
End If
End Sub

Private Sub dtgReserve_DblClick()
    Call mnuDetails_Click
End Sub

Private Sub dtgReserve_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuEdit, 0
    End If
End Sub

Private Sub Form_Activate()
    On Error GoTo Handle
    Dim ctrl As Control
        Desarr = LoadLanguage(LngFile, "#01:004:")
        Me.Caption = Desarr(1)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Desarr(Mid(ctrl.Tag, 2))
        Next ctrl
        'If cmdExittoLogin.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    isclick = False
    If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    'Init the Location
    Call Load_Section
    'Init Table on the first Location
    If Sec_ID <> "" And Sec_ID <> "TO" And Sec_ID <> "DE" And Sec_ID <> "AR" Then
        Sleep (500)
        Call LoadTable(Sec_ID)
        lblSection.Caption = cmdSection(1).Caption
    Else
        Sec_ID = Get_First_Section
        Call LoadTable(Sec_ID)
    End If
    ' Gan font mac dinh cho nhan cong ty
        lblCompanyname.Font.name = "VNI-Algerian"
        lblAdd.Font.name = "VNI-Times"
        lblPhone.Font.name = "VNI-Times"
        'If UserLevel <> 1 Then cmdSynchronize.Enabled = False
    'fraTakeOut.Visible = False
    fraTable.BackColor = bkColor
    frabk.BackColor = bkColor
    pic1.BackColor = bkColor
    picCom.BackColor = bkColor
    Call Load_Reserve
    If UserLevel <> 1 Then CheckRight
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Handle
        If KeyCode = vbKeyF1 Then
            frmAboutInfor.Show vbModal
        End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Form_KeyDown"
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If Shift = 2 Then
    If KeyCode = 113 Then
        MsgBox "HDD:" & MachineID.Get_Disk_Serial & Space(10) & "MAIN:" & MachineID.Get_MainSerial
    End If
Else
     If KeyCode = vbKeyF3 Then
        If cmdPrint_Receipt.Enabled = True Then
            cmdPrint_Buffer_Click
        End If
     ElseIf KeyCode = vbKeyF4 Then
        If cmdPayment.Enabled = True Then
            cmdPayment_Click
        End If
     ElseIf KeyCode = vbKeyF5 Then
        If cmdReserve.Enabled = True Then
            cmdReserve_Click
        End If
     ElseIf KeyCode = vbKeyF6 Then
        If cmdPrintedBill.Enabled = True Then
            cmdPrintedBill_Click
        End If
     ElseIf KeyCode = vbKeyF7 Then
        If cmdCash_Open.Enabled = True Then
            Call cmdCash_Open_Click
        End If
    ElseIf KeyCode = vbKeyF12 Then
        MsgBox "HDD:" & MachineID.Get_Disk_Serial & Space(10) & "MAIN:" & MachineID.Get_MainSerial
    End If
End If
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    Dim ctrl As Control
    Desarr = LoadLanguage(LngFile, "#01:004:")
    Me.Caption = Desarr(1)
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Desarr(Mid(ctrl.Tag, 2))
    Next ctrl
    Set rsinvoice_hold = Open_Table(cnData, "Invoice_OnHold")
  
    isclick = False
    isTimer = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
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
    If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
        cnData.Execute "Delete  from Table_Diagram_Sections where Location_ID='" & Sec_ID & "'"
    Call Load_Section
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   cmdDeleteLocation_Click "
End Sub

Private Sub cmdDone_Click()
    Unload Me
End Sub

Private Sub cmdSection_Click(Index As Integer)
    On Error GoTo Handle
    Dim ctrl As Control
        Section_Select = True
        Sec_ID = Format(cmdSection(Index).Tag, "00")
        Call LoadTable(CStr(Sec_ID))
        fraTable.Enabled = True
        iLoad = True
        lblSection.Caption = cmdSection(Index).Caption
'    fraTakeOut.Visible = False
    For Each ctrl In Me
        If ctrl.name = "cmdSection" Then
            ctrl.ForeColor = vbBlue
        End If
    Next ctrl
    cmdSection(Index).ForeColor = vbRed
    Call Added_Charge
    Exit Sub
    
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdSection_Click "
End Sub

Private Sub Form_Resize()
    pic1.Left = Me.Width - pic1.Width
    fraTable.Width = Me.Width - pic1.Width
    picCom.Left = Me.Width - picCom.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsSection = Nothing
    Set rsTable = Nothing
    CountTable = 0
    iLoad = False
    CountSection = 0
    iLoadSection = False
    EventCall = ""
    Section_Select = False
    kp_item = ""
End Sub






Private Sub lblTable_Click(Index As Integer)
On Error GoTo Handle
    If UserID = "131112" Or UserID = "0909419887" Then
        MsgBox "M· sè Qu¶n TrÞ kh«ng ®­îc b¸n hµng !"
    Else
        Dim rsinvoice_hold As New ADODB.Recordset
        Dim rsInvoice_Total As New ADODB.Recordset
        Dim rsInvoice_Notes As New ADODB.Recordset
        Dim i As Integer
        Dim IsPrintTranfer As Boolean
        IsPrintTranfer = False
        'Neu chua chon khu thi load muc gia
        If Section_Select = False Then Call Added_Charge
        If ArrayFlag(SF(4), 7) = 1 Then IsPrintTranfer = True
        isclick = True
'        Discount = 0
'        If cnData.State = 0 Then
'            Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'        End If
        
        If cnData.State <> 0 Then
            Set rsinvoice_hold = OpenCriticalTable("select * from Invoice_OnHold where Section_ID='" & Sec_ID & "'", cnData)
            Set rsInvoice_Total = OpenCriticalTable("select * from Invoice_Totals where Station_ID='" & Sec_ID & "'", cnData)
            Set rsInvoice_Notes = Open_Table(cnData, "Invoice_Totals_Notes")
        End If
        Table_ID = Left(lblTable(Index).Caption, InStr(lblTable(Index).Caption, Chr(13)))
       
        Select Case StateCall
            Case 1  'Mo ban binh thuong
            If Sec_ID = "" Then
                MsgBox "B¹n ph¶i chän khu vùc tr­íc khi më bµn! C¶m ¬n!", vbInformation
            Else
                Table_ID = Left(lblTable(Index).Caption, InStr(lblTable(Index).Caption, Chr(13)))
               
                'Print #fFile, "+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
                'Print #fFile, "Më bµn: " & Table_ID & vbTab & Now & vbTab & ":" & userName
                'Ghi HD xuong bang Invoice_onHold
                Call sUpdate_Invoice_Hold
                'ghi du lieu xuong Invoie_Note va Invoice_totals
                Call pUpdate_Invoice_Notes(currentBill)
                If fUpdate_Invoice_Total(currentBill, 0) Then
                    Set rsOrdered = Let_Record_Ordered(currentBill)
                    'Mo giao dien Order
                         With frmOrder
                            .Get_Secion = Sec_ID
                            .Get_Record_Ordered = rsOrdered
                            .GetBill_Number = currentBill
                            .Get_Table_ID = Table_ID
'                            .Get_Discount = Discount
                            .Get_Price_Level = PRLEVEL
                            .Get_VAT = VAT
                            .Get_Service = Sercharge
                            .Get_PriceRate = PRICE_RATE
                            .Get_TimeLevel = TimeLevel
                            .FormCall = 2
                            Picwait.Visible = False
                            probarWait.Value = 0
                            .Show vbModal
                        End With
                '    picWait.Visible = False
                End If
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
                        MsgBox " Bµn nµy ®· cã, vui lßng chän chøc n¨ng gép bµn !!", vbInformation
                        isclick = False
                        Picwait.Visible = False
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
                'Ghi Thong tin chuyen ban xuong file Log
                'Print #fFile, "ChuyÓn bµn: " & TranferTable & vbTab & "-->" & vbTab & DesTab & vbTab & Now & vbTab & ":" & userName
    
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
                            If Mid(kp_item, i, 1) = 1 Then
                                If ArrayFlag(SF(6), 5) = 1 Then
                                    If arrPriterKP(i) <> "" Then
                                        Dim prtName As String
                                        prtName = Get_Printer_Order(Sec_ID, Format(i, "00"))
                                        Call Print_Tranfer_Bill(BillTranfer, prtName, StateCall)
                                    End If
                                Else
                                    If arrPriterKP(i) <> "" Then
                                        Dim rsPrintKP As New ADODB.Recordset
                                        Dim Printer_Name As String
                                        Set rsPrintKP = Open_Table(cnData, "Printer_Mapping")
                                        With rsPrintKP
                                            .Find "PrinterName='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
                                            If Not .EOF Then
                                                Call Print_Tranfer_Bill(BillTranfer, .Fields("Details"), StateCall)
                                            End If
                                        End With
                                    End If
                                End If
                            End If
                        Next
            
                    End If
                End If
        '        '==============================================
                frmMessage.Show vbModal
                Picwait.Visible = False
                probarWait.Value = 0
                If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
                    cnData.Execute "delete from Tranfer_Joint_table"
                    StateCall = 1
                    kp_item = ""
        Case 3  'Gop ban
                'Tim so Bill Ban dich duoc chuyen toi
                Dim DesTable As String
                Dim DesBill As String
                Dim rsFindBill As New ADODB.Recordset
                If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
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
                'Print #fFile, "Gäp bµn: " & TranferTable & "(H§:" & BillTranfer & ")" & vbTab & "-->" & vbTab & DesTable & " (H§:" & DesBill & ")" & vbTab & Now & vbTab & ":" & userName
                '======================================================================
                'Cap nhat so bill cua Ban dich vao danh muc hang ban voi so Bill cua ban Nguon
                Dim rsInvoice_Itemized As New ADODB.Recordset
                Dim rsmaxLine As New ADODB.Recordset
                i = 0
                Set rsmaxLine = OpenCriticalTable("select Max(Invoice_Itemized.LineNum)as MaxLine from Invoice_Itemized where Invoice_Number='" & DesBill & "'", cnData)
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
                        .Fields("Total_Tax1") = 0
                        .Fields("Total_Tax2") = 0
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
                        .Fields("Total_tax1") = .Fields("Total_tax1") + dblTotal_Org
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
                            If Mid(kp_item, i, 1) = 1 Then
                                If arrPriterKP(i) <> "" Then
                                    Set rsPrintKP = Open_Table(cnData, "Printer_Mapping")
                                    With rsPrintKP
                                        .Find "PrinterName='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
                                        If Not .EOF Then
                                            Call Print_Tranfer_Bill(CDbl("0" & DesBill), .Fields("Details"), StateCall)
                                        End If
                                    End With
                                End If
                            End If
                        Next
                    End If
                End If
                
                frmMessage.Show vbModal
                StateCall = 1
                kp_item = ""
                Set rsOrdered = Let_Record_Ordered(DesBill)
                'Mo form Order
                With frmOrder
                    .Get_Secion = Sec_ID
                    .GetBill_Number = DesBill
                    .Get_Record_Ordered = rsOrdered
                    .Get_Table_ID = DesTable
'                    .Get_Discount = Discount
                    .FormCall = 2
                    Picwait.Visible = False
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
                With rsTranfer
                    'Print #fFile, "ChuyÓn mãn " & TranferTable & "(H§:" & BillTranfer & ")" & vbTab & "-->" & vbTab & TableDestination & " (H§:" & BillDestination & ")" & vbTab & Now & vbTab & ":" & userName
                    Do While Not .EOF
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
                        'Print #fFile, vbTab & rsTranfer.Fields("PluNo") & vbTab & rsTranfer.Fields("PluName") & vbTab & rsTranfer.Fields("Qty") & vbTab & rsTranfer.Fields("Std_Price1")
                        .MoveNext
                        Loop
                    End With
                ''' Cap nhat tong tien cho Ban nay
                If rsInvoice_Total.State = 0 Then Set rsInvoice_Total = Open_Table(cnData, "Invoice_Totals")
                    With rsInvoice_Total
                        .Find "Invoice_Number=" & BillDestination, , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("Total_Price") = .Fields("Total_Price") + dblTrans
                            .Update
'                            .Requery
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
                    If rsInvoice_Notes.State = 0 Then Set rsInvoice_Notes = Open_Table(cnData, "Invoice_Totals_Notes")
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
                    If rsInvoice_Total.State = 0 Then Set rsInvoice_Total = Open_Table(cnData, "Invoice_Totals")
                    With rsInvoice_Total
                        .Find "Invoice_Number='" & BillDestination & "'", , adSearchForward, adBookmarkFirst
                            If Not .EOF Then
                                Set rscust = Open_Table(cnData, "Customer")
                                    rscust.Find "CustNum='" & .Fields("CustNum") & "'", , adSearchForward, adBookmarkFirst
                                    If Not rscust.EOF Then
                                        CustNo(0) = .Fields("CustNum")
                                        CustNo(1) = rscust!CustName
                                        CustNo(2) = rscust!Acct_Balance
                                       ' Discount = CDbl("0" & rscust.Fields("Discount"))
                                    End If
'                                Discount = .Fields("Discount")
                            Else
                                ' ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                                .addNew
                                .Fields("Invoice_Number") = BillDestination
                                .Fields("Store_ID") = Store_ID
                                .Fields("Total_Cost") = 0
                                .Fields("Total_Price") = .Fields("Total_Price") + dblTrans
                                .Fields("Total_Tax1") = .Fields("Total_Tax1") + dblTrans
                                .Fields("Total_Tax2") = 0
                                .Fields("Total_Tax3") = 0
                                .Fields("Grand_Total") = .Fields("Grand_Total") + dblTrans
                                .Fields("CustNum") = "101"
                                .Fields("DateTime") = DateDefault & Format(Now, "HH:mm:ss")
                                .Fields("InvoiceNotesUsed") = -1
                                .Fields("Status") = "O"
                                .Fields("Station_ID") = Sec_ID
                                .Fields("Cashier_ID") = UserID
                                .Fields("Payment_MeThod") = "CA"
                                .Fields("InvType") = 0
                                .Fields("Orig_OnHoldID") = Trim(Table_ID)
    '                            .Fields("Tax_Rate_ID") = 0
                                .Update
                            End If
                    End With
                    
                    Dim rsInvoice_Items As New ADODB.Recordset
                    Set rsInvoice_Items = Open_Table(cnData, "Invoice_Itemized")
                    i = 0
                    If rsTranfer.State = 1 Then rsTranfer.MoveFirst
                    With rsTranfer
                        'Print #fFile, "ChuyÓn mãn " & TranferTable & "(H§:" & BillTranfer & ")" & vbTab & "-->" & vbTab & TableDestination & " (H§:" & BillDestination & ")" & vbTab & Now & vbTab & ":" & userName
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
                            rsInvoice_Items.Fields("Kit_Description") = Trim(rsTranfer.Fields("Kit_Desc")) & ""
                            rsInvoice_Items.Update
                            'Print #fFile, vbTab & rsTranfer.Fields("PluNo") & vbTab & rsTranfer.Fields("PluName") & vbTab & rsTranfer.Fields("Qty") & vbTab & rsTranfer.Fields("Std_Price1")
                        .MoveNext
                        i = i + 1
                        Loop
                    End With
                
                End If
                End If
                'In Phieu chuyen mon xuong bep
                'Call Print_Item_tranfer(rsTranfer)
                
                Set rsOrdered = Let_Record_Ordered(BillDestination)
                 With frmOrder
                    .Get_Secion = Sec_ID
                    .GetBill_Number = BillDestination
                    .Get_Table_ID = TableDestination
                    .Get_Record_Ordered = rsOrdered
'                    .Get_Discount = Discount
                    .FormCall = 2
                    Picwait.Visible = False
                    probarWait.Value = 0
                    .Show vbModal
                End With
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
                    'Print #fFile, "ChuyÓn mãn " & TranferTable & "(H§:" & BillTranfer & ")" & vbTab & "-->" & vbTab & TableDestination & " (H§:" & BillDestination & ")" & vbTab & Now & vbTab & ":" & userName
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
                        .Fields("Kit_Description") = Trim(rsTranfer.Fields("Kit_Desc")) & ""
                        .Update
                    'Print #fFile, vbTab & rsTranfer.Fields("PluNo") & vbTab & rsTranfer.Fields("PluName") & vbTab & rsTranfer.Fields("Qty") & vbTab & rsTranfer.Fields("Std_Price1")
                    rsTranfer.MoveNext
                    i = i + 1
                    Loop
                End With
            Set rsTranfer = Nothing
            frmMessage.lblTitle.Caption = "Mãn b·n chän ®· ®­îc chuyÓn l¹i bµn gèc !"
            frmMessage.Show vbModal
            
        End If
    Case 5
        lblSync.Visible = True
         lblSync.FontSize = 24
        lblSync.Caption = "§ang in phiÕu"
        Timer3.Interval = 0
        
        If rsinvoice_hold.State = 0 Then Set rsinvoice_hold = OpenCriticalTable("Select * from Invoice_OnHold where Section_ID='" & Sec_ID & "'", cnData)
            With rsinvoice_hold
                .Find "OnHoldID='" & Table_ID & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        Call Print_Receipt_Count(.Fields("Invoice_Number"))
                        Call Print_Receipt(.Fields("Invoice_Number"))
                        StateCall = 1
                        TimerBackColor.Interval = 0
                        cmdPrint_Receipt.BackColor = &HFF8080
                    Else
                        MsgBox "Kh«ng cã d÷ liÖu ®Ó in"
                        TimerBackColor.Interval = 0
                        cmdPrint_Receipt.BackColor = &HFF8080
                    End If
            End With
            lblSync.Visible = False
        Case 6
        lblSync.Visible = True
        lblSync.FontSize = 24
        lblSync.Caption = "§ang thanh to¸n"
        Timer3.Interval = 0
             If rsinvoice_hold.State = 0 Then Set rsinvoice_hold = OpenCriticalTable("Select * from Invoice_OnHold where Section_ID='" & Sec_ID & "'", cnData)
            With rsinvoice_hold
                .Find "OnHoldID='" & Table_ID & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        currentBill = .Fields("Invoice_Number")
                        StateCall = 1
                        TimerBackColor.Interval = 0
                        cmdPayment.BackColor = &HFF8080
                        Call Payment(currentBill)
                    Else
                        MsgBox "Kh«ng cã d÷ liÖu ®Ó thanh to¸n"
                        TimerBackColor.Interval = 0
                        cmdPayment.BackColor = &HFF8080
                    End If
            End With
            lblSync.Visible = False
            
        End Select
        dblTotal_Org = 0
        Picwait.Visible = False
        StateCall = 1
        Set rsTranfer = Nothing
'        Discount = 0
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " lblTable_Click  "
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
        Set rsSection = OpenCriticalTable("select * from Table_Diagram_Sections order by Location_ID ASC", cnData)
        If rsSection.EOF Then Exit Sub
        If iLoadSection = True Then
            For Each ctrl In Me
                If TypeOf ctrl Is MyButton And ctrl.name = "cmdSection" Then
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
    MsgBox Err.Number & Err.Description & Me.name & "   Load_Section"
End Sub

Public Sub LoadTable(Section_ID As String)
On Error GoTo Handle
Dim rscolor As New ADODB.Recordset
Dim rsSeatedColor As New ADODB.Recordset
Dim rsVacantColor As New ADODB.Recordset
Dim rsSubtotalColor As New ADODB.Recordset
Dim rsInvoice_Total As New ADODB.Recordset

Dim i, j As Integer
i = 1: j = 1
    Dim str As String
    Dim ctrl As Control
    If CountTable > 0 Then
        For j = 1 To CountTable
'            DoEvents
            Unload lblTable(j)
            Unload Shape1(j)
            Unload lblTime(j)
        Next
    End If
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
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
            " where Invoice_OnHold.OnHoldID = '" & rsTable.Fields("Table_number") & Chr(13) & "'  and Invoice_OnHold.Section_ID='" & Section_ID & "'"
        
            Set rsInvoice_Total = OpenCriticalTable(strTableTotal, cnData)
            If rsInvoice_Total.RecordCount > 0 Then
                If CDbl("0" & rsInvoice_Total.Fields("Grand_Total")) > 0 Then
                    If RTrim(rsInvoice_Total.Fields("Status")) = "P" Then
                        If IsNull(rsInvoice_Total.Fields("Grand_Total")) Then
                             .Caption = rsTable.Fields("Table_Number") & Chr(13) & Format(0, formatNum)
                        Else
                            .Caption = rsTable.Fields("Table_Number") & Chr(13) & Format(Abs(rsInvoice_Total.Fields("Grand_Total")), formatNum)
                        End If
                        .BackStyle = 1
                        .BackColor = rsSubtotalColor.Fields("ReserveValue")
                        .Font.Size = rsTable.Fields("Cost_Center_Index")
                        lblTime(i).BackColor = vbRed
                    Else
                        .Caption = rsTable.Fields("Table_Number") & Chr(13) & Format(Abs(rsInvoice_Total.Fields("Grand_Total")), formatNum)
                        .BackStyle = 1
                        .BackColor = rscolor.Fields("ReserveValue")
                        .Font.Size = rsTable.Fields("Cost_Center_Index")
                        lblTime(i).BackColor = vbRed
                    End If
                Else
                    
                    If IsNull(rsInvoice_Total.Fields("Grand_Total")) Then
                             .Caption = rsTable.Fields("Table_Number") & Chr(13) & Format(0, formatNum)
                        Else
                           .Caption = rsTable.Fields("Table_Number") & Chr(13) & Format(Abs(rsInvoice_Total.Fields("Grand_Total")), formatNum)
                        End If
                    .BackStyle = 1
                    .BackColor = rsSeatedColor.Fields("ReserveValue")
                    .Font.Size = rsTable.Fields("Cost_Center_Index")
                    lblTime(i).BackColor = vbRed
                End If
                lblTime(i).Caption = Mid(rsInvoice_Total.Fields("DateTime"), 9, 8)
                lblTime(i).ToolTipText = gfCONVERT_STRING_TO_DATE(Left(rsInvoice_Total.Fields("DateTime"), 8))
                
            Else
                .Caption = rsTable.Fields("Table_Number") & Chr(13)
                .BackStyle = 1
                .BackColor = rsVacantColor.Fields("ReserveValue")
                .Font.Size = rsTable.Fields("Cost_Center_Index")
                lblTime(i).Caption = ""
                lblTime(i).BackStyle = 0
            End If
            .ForeColor = ColorFont
            .FontName = CurFont
            .Visible = True
        End With
        Load Shape1(i)
        With Shape1(i)
            .Left = lblTable(i).Left - 40
            .top = lblTable(i).top - 45
            .Height = lblTable(i).Height + 100
            .Width = lblTable(i).Width + 100
            .Shape = rsTable.Fields("ShapeType")
            .BorderColor = ShapeColor
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
MsgBox Err.Number & Err.Description & Me.name & "  LoadTable"
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
                    .FullRight = res.Fields("UserRight")
                    .Sodoban = RightDeCode(Left(.FullRight, 16))
                    .Danhmuc = RightDeCode(Mid(.FullRight, 33, 16))
                    .Nhanvien = RightDeCode(Mid(.FullRight, 193, 64))
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
            
            If Mid(.Nhanvien, 14, 1) = 0 Then
                  cmdCash_Open.Enabled = False
            Else: cmdCash_Open.Enabled = True
            End If
        End With

    CloseRecordset res
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " CheckRight"
End Sub

Private Sub lblTime_Click(Index As Integer)
    Call lblTable_Click(Index)
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



Private Sub mnuDetails_Click()
On Error GoTo Handle
 Dim cmd As New ADODB.Command
    Dim SQL As String
    Dim iReport As New CRAXDDRT.Report
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If

     SQL = "SELECT Table_Reservered.Reservered_Code, " & _
    " CustName, Address, Phone," & _
    " Date_Reservered," & _
    " Time_Reservered, Table_Reservered.Table_ID," & _
    " Amount, Table_Reservered.Description," & _
    " Table_Reserved_Details.ItemNum," & _
    " Table_Reserved_Details.ItemName, Table_Reserved_Details.Qty," & _
    " Table_Reserved_Details.Price,Table_Reserved_Details.Qty*Table_Reserved_Details.Price as AMT" & _
    " FROM Table_Reservered LEFT JOIN Table_Reserved_Details ON" & _
    " Table_Reservered.Reservered_Code = Table_Reserved_Details.Reservered_Code" & _
    " Where Table_Reservered.Reservered_Code='" & Reserve_Code & "'" & _
    " order by Table_Reserved_Details.ItemNum"
    Set crPhieudatcho = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crPhieudatcho
        .Database.AddADOCommand cnData, cmd
        .txtSophieu.SetUnboundFieldSource "{ado.Reservered_Code}"
        .txtKhachhang.SetUnboundFieldSource "{ado.CustName}"
        .txtDiachi.SetUnboundFieldSource "{ado.Address}"
        .txtDienthoai.SetUnboundFieldSource "{ado.Phone}"
        .txtBandat.SetUnboundFieldSource "{ado.Table_ID}"
        .txtNgaytiec.SetUnboundFieldSource "{ado.Date_Reservered}"
        .txtGiotiec.SetUnboundFieldSource "{ado.Time_Reservered}"
        .txtSotien.SetUnboundFieldSource "{ado.Amount}"
        .txtDiengiai.SetUnboundFieldSource "{ado.Description}"
        .txtmahang.SetUnboundFieldSource "{ado.ItemNum}"
        .TENMON.SetUnboundFieldSource "{ado.ItemName}"
        .SL.SetUnboundFieldSource "{ado.Qty}"
        .DG.SetUnboundFieldSource "{ado.Price}"
        .TTien.SetUnboundFieldSource "{ado.AMT}"
        'canh le
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign

'        .PrintOut
    End With
    Set iReport = crPhieudatcho
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " mnuDetails_Click"
End Sub

Private Sub mnuSelect_Click()
On Error GoTo Handle
 Dim Invoice_Num As Double
 Dim rsReserve_Details As New ADODB.Recordset
    Invoice_Num = GetMaxInvoice_Number
    'Ghi d÷ liÖu xuèng Invoice_On_Hold
        If fUpdate_Invoice_OnHold(Invoice_Num, Table_ID) Then
            'Ghi th«ng tin xuèng Invoice_totals_Notes
                Call pUpdate_Invoice_Notes(Invoice_Num)
            'Ghi d÷ liÖu xuèng Invoice_Totals
                Call fUpdate_Invoice_Total(Invoice_Num, Amount_Reserve)
            'Ghi chi tiÕt order xuèng Invoice_Itemized
               Set rsReserve_Details = Get_Reserve_Details(Reserve_Code)
                With frmOrder
                    .Get_Secion = Sec_ID
                    .Get_Record_Ordered = rsReserve_Details
                    .GetBill_Number = Invoice_Num
                    .Get_Table_ID = Table_ID
                    .FormCall = 2
                    .Show vbModal
                End With
            'Xãa th«ng tin d÷ liÖu ®Æt tiÖc
                Call gfUpdate_Reserve_isUsed(Reserve_Code)
        End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description
End Sub





Public Function GetTimeSync() As Double
    On Error GoTo Handle
    Dim i As Double
    Dim rsInfor As New ADODB.Recordset
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsInfor = Open_Table(cnData, "Setup")
    With rsInfor
        If Not rsInfor.EOF Then
            i = Hour(.Fields("TimeSync")) * 3600 + Minute(.Fields("TimeSync")) * 60 + Second(.Fields("TimeSync"))
            AmountBackup = .Fields("AmountLimited")
        End If
    End With
    
    GetTimeSync = i
    Exit Function
Handle:
    GetTimeSync = 0
    MsgBox Err.Number & Err.Description & Me.name & " GetTimeSync"
End Function



Public Function check_Backup() As Boolean
On Error GoTo Handle
    check_Backup = False
    If ArrayFlag(SF(3), 3) = 1 Then
        check_Backup = True
    End If
    
Exit Function
Handle:
check_Backup = False
MsgBox Err.Number & Err.Description & Me.name & "check_Backup"
End Function

Public Function GetMax_Invoice_Backup(DateMax As String) As Double
On Error GoTo Handle
Dim Max_Invoice As Double
    Dim rsmax As New ADODB.Recordset
    Dim cnmax As New ADODB.Connection
    If Dir(BackupFolder & "\Database.mdb", vbDirectory) <> "" Then
        Set cnmax = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    End If
    Set rsmax = OpenCriticalTable("select Max(Invoice_Number) as maxInvoice from Invoice_Totals where left(Invoice_Totals.DateTime,8)='" & DateMax & "'", cnmax)
    If rsmax.RecordCount <> 0 Then
        If Not rsmax.EOF Then
            Max_Invoice = rsmax.Fields("maxInvoice") + 1
        End If
    End If
    GetMax_Invoice_Backup = Max_Invoice
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & " GetMax_Invoice_Backup"
End Function


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
MsgBox Err.Number & Err.Description & Me.name & " Update_VIP_Percent"
End Sub

Public Function Get_First_Section() As String
On Error GoTo Handle
    Dim S As String
    Dim rsTable As New ADODB.Recordset
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    Set rsTable = OpenCriticalTable("Select Min(Location_ID) as minID from Table_Diagram_Sections", cnData)
    If Not rsTable.EOF And rsTable.RecordCount > 0 Then
        S = rsTable.Fields("MinID")
    Else
        S = "01"
    End If
Get_First_Section = S
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Get_First_Section"
End Function

Public Sub Update_Table_In_Order_Details(Bill_Source As Double, Bill_Des As Double, TranferType As Integer, Optional Source_Table As String, Optional Des_Table As String)
On Error GoTo Handle
    Dim rsKitchen_Master As New ADODB.Recordset
    Dim rsKitchen_Items As New ADODB.Recordset
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
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
                                .Fields("ItemNum") = Trim(rsTempItem.Fields("ItemNum"))
                                .Fields("ItemName") = Trim(rsTempItem.Fields("ItemName"))
                                .Fields("Quantity") = rsTempItem.Fields("Quantity")
                                .Fields("Price") = rsTempItem.Fields("Price")
                                .Fields("LineNum") = MaxLine
                                .Fields("Kit_Desc") = Trim(rsTempItem.Fields("Kit_Desc"))
                                .Fields("Printer_ID") = rsTempItem.Fields("Printer_ID")
                                .Fields("Send_KP_Date") = Trim(rsTempItem.Fields("Send_KP_Date"))
                                .Fields("Send_KP_Time") = Trim(rsTempItem.Fields("Send_KP_Time"))
                                .Update
                            End With
                        .MoveNext
                        MaxLine = MaxLine + 1
                        Loop
                    End With
                    If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
                    cnData.Execute "delete  from Kitchen_Order_Items where Invoice_Number=" & Bill_Source
                    cnData.Execute "Delete  from Kitchen_Order_Master where Invoice_Number=" & Bill_Source
                End If
                '.Fields("Invoice_Number") = Bill_Des
                .Fields("Table_ID") = Des_Table
                .Update
                
            End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Update_Table_In_Order_Details"
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
    MsgBox Err.Number & Err.Description & Me.name & "  Check_System_KP"
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
MsgBox Err.Number & Err.Description & Me.name & "  Update_Tranfer"
End Sub

Public Sub Print_Tranfer_Bill(BillTranfer As Double, Printer_Name As String, state_form As Integer)
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim iReport As CRAXDDRT.Report
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = "select * from Tranfer_Joint_table where Des_bill =" & BillTranfer
    Set crTranfer = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crTranfer
        .Database.AddADOCommand cnData, cmd
        If state_form = 2 Then
            .txtTitle.SetText "PhiÕu chuyÓn bµn"
            .txtstate.SetText "ChuyÓn "
        ElseIf state_form = 3 Then
            .txtTitle.SetText "PhiÕu gép bµn"
            .txtstate.SetText "Gép"
        End If
        .billDes.SetUnboundFieldSource "{ado.Des_bill}"
        .BillOrg.SetUnboundFieldSource "{ado.Org_bill}"
        .LocationDes.SetUnboundFieldSource "{ado.Des_Location}"
        .Location.SetUnboundFieldSource "{ado.Org_Location}"
        .TableDes.SetUnboundFieldSource "{ado.Des_Table}"
        .TableOrg.SetUnboundFieldSource "{ado.Org_Table}"
        .CashierID.SetUnboundFieldSource "{ado.Cashier_ID}"
    End With
    Set iReport = crTranfer
    With frmShowSendKP
        .Report = crTranfer
        .Get_ID = "01"
        .GetPrinter = Printer_Name
        .Show vbModal
    End With
Exit Sub
Handle:
Exit Sub
MsgBox Err.Number & Err.Description & Me.name & " Print_Tranfer_Bill"
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
MsgBox Err.Number & Err.Description & Me.name & "  Get_Right_Location"
End Function



Public Sub Added_Charge()
On Error GoTo Handle
    With rsSection
        .Find "Location_ID='" & Sec_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            PRLEVEL = Val("0" & .Fields("Price_Level"))
            Sercharge = Val("0" & .Fields("Service_Charge"))
            VAT = Val("0" & .Fields("VAT"))
            PRICE_RATE = Val("0" & .Fields("PriceRate"))
            isTimer = .Fields("isTimer")
            TimeLevel = Val("0" & .Fields("TimeLevel"))
        Else
            MsgBox "Kh«ng t×m thÊy møc gi¸ cho khu vùc nµy! Vui lßng chän l¹i khu vùc ", vbInformation
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Added_Charge"
End Sub

Public Sub Update_CO_By_C(ByVal InvoiceDateTime As Double, cn As ADODB.Connection)
On Error GoTo Handle
'    Dim MinInvoice As Double
    Dim rsInvoice_Totals As ADODB.Recordset
'    MinInvoice = Get_MinInvoice(cnData, InvoiceDateTime)
Dim str As String
    str = "Select * from Invoice_totals" & _
            " where Status='CO' and Left(DateTime,8)='" & InvoiceDateTime & "'"
    Set rsInvoice_Totals = OpenCriticalTable(str, cn)
    With rsInvoice_Totals
        Do While Not .EOF
            If .Fields("Grand_total") > 0 Then
                .Fields("Status") = "C"
                .Update
            End If
        rsInvoice_Totals.MoveNext
        Loop
    End With
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Update_CO_By_C"
End Sub

Public Sub pUpdate_Invoice_Notes(ByVal BillNO As Double)
On Error GoTo Handle
Dim rsInvoice_Notes As New ADODB.Recordset
Set rsInvoice_Notes = Open_Table(cnData, "Invoice_totals_notes")
      With rsInvoice_Notes
        .Find "Invoice_Number='" & BillNO & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                ' ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                .addNew
                .Fields("Invoice_Number") = BillNO
                .Fields("Store_ID") = Store_ID
                .Fields("OpenTime") = DateDefault & TimeOrder
                .Fields("ClosingTime") = "C"
                .Update
            End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Update_Invoice_Notes"
End Sub

Public Function fUpdate_Invoice_Total(ByVal BillNO As Double, Amount As Double) As Boolean
On Error GoTo Handle
Dim rsInvoice_Totals As ADODB.Recordset
Dim isUpdate As Boolean
    isUpdate = False
       Set rsInvoice_Totals = OpenCriticalTable("select * from Invoice_Totals ", cnData)
            With rsInvoice_Totals
                If rsInvoice_Totals.State = 1 And .RecordCount > 0 Then rsInvoice_Totals.MoveFirst
                .Find "Invoice_Number='" & BillNO & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        If .Fields("InvoiceNotesUsed") = True Then
                            MsgBox "Bµn nµy ®· ®­îc më t¹i mét m¸y kh¸c, B¹n kh«ng thÓ më bµn nµy!!!"
                             Picwait.Visible = False
                            probarWait.Value = 0
                            isUpdate = False
                            Exit Function
                        End If
                        If .Fields("CustNum") <> "101" Then
                            Set rscust = Open_Table(cnData, "Customer")
                                rscust.Find "CustNum='" & .Fields("CustNum") & "'", , adSearchForward, adBookmarkFirst
                                If Not rscust.EOF Then
                                    CustNo(0) = .Fields("CustNum")
                                    CustNo(1) = rscust!CustName & ""
                                    CustNo(2) = CDbl("0" & rscust!Acct_Balance)
                                    .Fields("InvoiceNotesUsed") = True
                                    .Update
                                End If
                        Else
                        End If
                        .Fields("InvoiceNotesUsed") = True
                        .Update
                         isUpdate = True
                    Else
                        ' ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                        .addNew
                        .Fields("Invoice_Number") = BillNO
                        .Fields("Store_ID") = Store_ID
                        .Fields("CustNum") = "101"
                        .Fields("DateTime") = DateDefault & TimeOrder
                        .Fields("Total_Price") = 0
                        .Fields("Total_Cost") = 0
                        .Fields("Total_Tax1") = 0
                        .Fields("Total_Tax2") = 0
                        .Fields("Total_Tax3") = 0
                        .Fields("Grand_Total") = 0
                        .Fields("InvoiceNotesUsed") = True
                        .Fields("Status") = "O"
                        .Fields("Station_ID") = Sec_ID
                        .Fields("Cashier_ID") = UserID
                        .Fields("Payment_MeThod") = "CA"
                        .Fields("InvType") = 0
                        .Fields("Orig_OnHoldID") = Trim(Table_ID)
                        .Fields("VATFee") = VAT
                        .Fields("Reserve") = Amount
                        .Fields("Service_Charge") = Sercharge
                        .Fields("Synchronized") = "False"
                        .Update
                        isUpdate = True
                    End If
            End With
            fUpdate_Invoice_Total = isUpdate
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & " sUpdate_Invoice_Total"
End Function

Public Sub sUpdate_Invoice_Hold()
On Error GoTo Handle
Dim strSql As String
strSql = "Select * from Invoice_OnHold where Section_ID='" & Sec_ID & "'"
Set rsinvoice_hold = OpenCriticalTable(strSql, cnData)

    With rsinvoice_hold
    If .RecordCount > 0 Then .MoveFirst
        .Find "OnHoldID='" & Table_ID & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                'Khong ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                currentBill = .Fields("Invoice_Number")
            Else
                If isTimer Then
                    With frmTimeLogin
                        .GetOpen = True
                        .lblTitle.Caption = "Më bµn/Phßng"
                        .Show vbModal
                        TimeOrder = .Get_Time_In
                        If TimeOrder = "" Then MsgBox "Vui lßng chän giê më", vbInformation
                    End With
                Else
                    TimeOrder = Format(Now, "HH:mm:ss")
                End If
                'Ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                currentBill = GetMaxInvoice_Number
                SaveSettingStr "SYSTEM", "MaxInvoice", currentBill, myIniFile
                .addNew
                .Fields("Invoice_Number") = currentBill
                .Fields("OnHoldID") = Table_ID
                .Fields("Cashier_ID") = UserID
                .Fields("Store_ID") = Store_ID
                .Fields("Occupied") = -1
                .Fields("Section_ID") = Sec_ID
                .Fields("Status") = 0
                .Update
            End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " sUpdate_Invoice_Hold"
End Sub


Private Sub Timer3_Timer()
If light < 2 Then
    If StateCall = 5 Then
        cmdPrint_Receipt.BackColor = &H80FF&
    Else
        cmdPayment.BackColor = &HFF8080
    End If
    light = light + 1
Else
    If StateCall = 5 Then
        cmdPrint_Receipt.BackColor = &HFF8080
    Else
        cmdPayment.BackColor = &H80FF&
    End If
    light = 0
End If
End Sub

Public Sub Load_Reserve()
On Error GoTo Handle
    Dim strSql As String
    Dim rsReserve As New ADODB.Recordset
    
    strSql = "select Reservered_Code,CustName,Phone,Time_Reservered,Table_ID,Amount,Section_ID from Table_Reservered where Table_Reservered.Date_Reservered='" & gfCONVERT_STRING_TO_DATE(DateDefault) & "' and IsUsed=0"
    Set rsReserve = OpenCriticalTable(strSql, cnData)
    
    If rsReserve.RecordCount > 0 Then
        Set dtgReserve.DataSource = rsReserve
        Call Init_caption_Reserve
    Else
        Set dtgReserve.DataSource = Nothing
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Kh«ng load ®­îc d÷ liÖu ®Æt bµn"
End Sub

Public Sub Init_caption_Reserve()
On Error GoTo Handle
    With dtgReserve
        .Columns(0).Width = 0
        .Columns(1).Caption = "Tªn kh¸ch"
        .Columns(1).Width = 1200
        
        .Columns(2).Caption = "§T"
        .Columns(2).Width = 1000
        
        .Columns(3).Caption = "Giê tiÖc"
        .Columns(3).Width = 1200
        
        .Columns(4).Caption = "Sè bµn"
         .Columns(4).Width = 800
        .Columns(5).Width = 0
        .Columns(6).Width = 0
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Kh«ng khëi t¹o ®­îc caption danh s¸ch ®Æt bµn"
End Sub


Public Function fUpdate_Invoice_OnHold(invoice_No As Double, table_Num As String) As Boolean
On Error GoTo Handle
Dim isUpdate As Boolean
Dim strSql As String
strSql = "Select * from Invoice_OnHold where Section_ID='" & Sec_ID & "'"
Set rsinvoice_hold = OpenCriticalTable(strSql, cnData)
    With rsinvoice_hold
        .Find "OnHoldID='" & table_Num & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("Invoice_Number") = invoice_No
                .Fields("OnHoldID") = table_Num
                .Fields("Cashier_ID") = UserID
                .Fields("Store_ID") = Store_ID
                .Fields("Occupied") = -1
                .Fields("Section_ID") = Sec_ID
                .Fields("Status") = 0
                .Update
                isUpdate = True
            Else
                MsgBox " Bµn nµy ®ang cã kh¸ch nªn kh«ng thÓ gäi bµn"
                isUpdate = False
            End If
    End With
    fUpdate_Invoice_OnHold = isUpdate
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & " sUpdate_Invoice_Hold"
End Function


Public Property Get Let_Record() As Variant
    Let_Record = rsOrdered
End Property



Public Function Get_Reserve_Details(Condition As String) As ADODB.Recordset
On Error GoTo Handle
    Dim str As String
    Dim rs As New ADODB.Recordset
    str = "SELECT Table_Reservered.Reservered_Code, Table_Reservered.Table_ID, Table_Reserved_Details.ItemNum as PluNo," & _
            " Table_Reserved_Details.ItemName as PluName, Table_Reserved_Details.Qty, Table_Reserved_Details.Price as Std_Price1," & _
            " Table_Reserved_Details.LineDisc, Table_Reserved_Details.Line_Disc_Desc, Table_Reserved_Details.Kit_Desc" & _
            " FROM Table_Reservered INNER JOIN Table_Reserved_Details ON Table_Reservered.Reservered_Code" & _
            " = Table_Reserved_Details.Reservered_Code" & _
            " where Table_Reservered.Reservered_Code='" & Condition & "'"
    Set rs = OpenCriticalTable(str, cnData)
    Set Get_Reserve_Details = rs
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Get_Reserve_Details"
Set Get_Reserve_Details = Nothing
End Function

Public Sub gfUpdate_Reserve_isUsed(code As String)
On Error GoTo Handle
If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    cnData.Execute "Update Table_Reservered set IsUsed=1 where Reservered_Code='" & code & "'"
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  gfUpdate_Reserve_isUsed"
End Sub

Public Sub Payment(ByVal Bill As String)
On Error GoTo Handle
Dim Grand_Total As Double
    Dim rsInvoice_Totals As New ADODB.Recordset
    Set rsInvoice_Totals = Open_Table(cnData, "Invoice_Totals")
    With rsInvoice_Totals
        .Find "Invoice_Number=" & Bill, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            Grand_Total = .Fields("Grand_Total")
        End If
    End With
    With frmPayment
        .Get_Grand_Total = Grand_Total
        .Get_Invoice_Number = Bill
        .Show vbModal
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub


Public Property Let get_item_tranfer(ByVal vNewValue As Variant)
    kp_item = vNewValue
End Property

