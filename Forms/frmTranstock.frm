VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmTranstock 
   Caption         =   "ChuyÓn kho"
   ClientHeight    =   11055
   ClientLeft      =   165
   ClientTop       =   510
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
   ScaleHeight     =   11055
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin prjTouchScreen.MyButton cmdKeyboard 
      Height          =   375
      Left            =   5880
      TabIndex        =   31
      Top             =   4320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   1
      TX              =   "Keyboard"
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
      BCOLO           =   8438015
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTranstock.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Frame Frame4 
      Height          =   4335
      Left            =   6960
      TabIndex        =   7
      Top             =   480
      Width           =   8175
      Begin VB.ComboBox cboStock_Type 
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
         ItemData        =   "frmTranstock.frx":001C
         Left            =   1560
         List            =   "frmTranstock.frx":0026
         TabIndex        =   36
         Text            =   "Combo1"
         Top             =   2640
         Width           =   6495
      End
      Begin VB.ComboBox cboStockSource 
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         ItemData        =   "frmTranstock.frx":004D
         Left            =   1560
         List            =   "frmTranstock.frx":004F
         TabIndex        =   25
         Top             =   2040
         Width           =   2775
      End
      Begin VB.ComboBox cboStockDes 
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   23
         Top             =   2040
         Width           =   2775
      End
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   855
         Left            =   6615
         TabIndex        =   21
         Top             =   3360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "Th&o¸t"
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
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTranstock.frx":0051
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdNewDoc 
         Height          =   855
         Left            =   240
         TabIndex        =   20
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "&Thªm CT"
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
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTranstock.frx":006D
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
         Height          =   855
         Left            =   3420
         TabIndex        =   19
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "&CËp nhËt CT"
         ENAB            =   0   'False
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
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTranstock.frx":0089
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmddelete 
         Height          =   855
         Left            =   1830
         TabIndex        =   18
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "&Xãa CT"
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
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTranstock.frx":00A5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   17
         Top             =   1440
         Width           =   6495
      End
      Begin VB.TextBox txtReceiver 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   16
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtDeliver 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox txtDocNum 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpDateOut 
         Height          =   435
         Left            =   5400
         TabIndex        =   10
         Top             =   240
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   767
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
         Format          =   64094209
         UpDown          =   -1  'True
         CurrentDate     =   38594
      End
      Begin prjTouchScreen.MyButton cmdEdit 
         Height          =   855
         Left            =   5010
         TabIndex        =   35
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         BTYPE           =   1
         TX              =   "&Söa ch÷a"
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
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTranstock.frx":00C1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Lo¹i kho:"
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
         Height          =   375
         Left            =   360
         TabIndex        =   37
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Kho ®Ých:"
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
         Left            =   4320
         TabIndex        =   24
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Kho nguån:"
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
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Lý do chuyÓn:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Ng­êi nhËn:"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4320
         TabIndex        =   13
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Ng­êi chuyÓn:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Ngµy CT:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4320
         TabIndex        =   11
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Sè CT:"
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
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame fraDoc 
      Caption         =   "Chøng tõ chuyÓn kho"
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6735
      Begin MSDataGridLib.DataGrid grdTranStock 
         Height          =   3735
         Left            =   120
         TabIndex        =   34
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6588
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   21
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   12
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
   Begin VB.Frame Frame3 
      Height          =   6255
      Left            =   9000
      TabIndex        =   4
      Top             =   4800
      Width           =   6135
      Begin MSDataGridLib.DataGrid grdPluDes 
         Height          =   5895
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   10398
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
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
   Begin VB.Frame Frame2 
      Height          =   6255
      Left            =   6960
      TabIndex        =   3
      Top             =   4800
      Width           =   1935
      Begin VB.TextBox txtItemCode 
         Height          =   495
         Left            =   120
         TabIndex        =   38
         Top             =   5400
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtQtyTran 
         Alignment       =   1  'Right Justify
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
         Left            =   240
         TabIndex        =   30
         Top             =   1800
         Width           =   1455
      End
      Begin prjTouchScreen.MyButton cmdAddOne 
         Height          =   735
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1296
         BTYPE           =   1
         TX              =   ">"
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
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTranstock.frx":00DD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdRemoveOne 
         Height          =   735
         Left            =   240
         TabIndex        =   29
         Top             =   2520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1296
         BTYPE           =   1
         TX              =   "<"
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
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTranstock.frx":00F9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdSaveDoc 
         Height          =   1095
         Left            =   120
         TabIndex        =   33
         Top             =   3600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1931
         BTYPE           =   1
         TX              =   "&L­u CT ChuyÓn kho"
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
         BCOL            =   12632256
         BCOLO           =   33023
         FCOL            =   255
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTranstock.frx":0115
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
   Begin VB.TextBox txtSearch 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "M· hµng kho nguån"
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   6735
      Begin MSDataGridLib.DataGrid grdPluSource 
         Height          =   5775
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   10186
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
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
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "chuyÓn kho"
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
      Left            =   6960
      TabIndex        =   32
      Top             =   0
      Width           =   8175
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "chuyÓn kho"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   6960
      TabIndex        =   28
      Top             =   -480
      Width           =   8175
   End
   Begin VB.Label Label2 
      Caption         =   "NhËp m· hµng cÇn chuyÓn:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4365
      Width           =   2535
   End
End
Attribute VB_Name = "frmTranstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMasterOrg As New ADODB.Recordset
Dim rsMasterDes As New ADODB.Recordset
Dim rsStockOrg As New ADODB.Recordset
Dim rsStockDes As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim rsTon As New ADODB.Recordset
Dim rsTontemp As New ADODB.Recordset
Dim rsTranStock As New ADODB.Recordset
Dim Tonthang, strPLU As String
Dim isLoad As Boolean

Private Sub cboStock_Type_Change()
On Error GoTo Handle
    Call Init_Transtock
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Thay doi danh muc kho"
End Sub

Private Sub cboStock_Type_Click()
    Call cboStock_Type_Change
End Sub

Private Sub cboStockDes_Change()
    If cboStockDes.Text = cboStockSource.Text Then
        MsgBox "Kho Nguån vµ kho ®Ých kh«ng ®­îc trïng nhau", vbInformation
        If cboStockDes.ListIndex = 0 Then
            cboStockDes.ListIndex = 1
        Else
            cboStockDes.ListIndex = 0
        End If
    End If
End Sub

Private Sub cboStockDes_Click()
    Call cboStockDes_Change
End Sub

Private Sub cboStockDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdUpdate.SetFocus
End Sub

Private Sub cboStockSource_Change()
    If cboStockSource.Text = cboStockDes.Text Then
        MsgBox "Kho Nguån vµ kho ®Ých kh«ng ®­îc trïng nhau", vbInformation
        If cboStockSource.ListIndex = 0 Then
            cboStockSource.ListIndex = 1
        Else
            cboStockSource.ListIndex = 0
        End If
    End If
End Sub

Private Sub cboStockSource_Click()
    Call cboStockSource_Change
End Sub

Private Sub cboStockSource_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboStockDes.SetFocus
End Sub

Private Sub cmdAddOne_Click()
On Error GoTo Handle
    With rsTemp
        .Find "ItemNum='" & strPLU & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                With rsTontemp
                    If .RecordCount > 0 Then .MoveFirst
                    .Find "ItemNum='" & strPLU & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        rsTemp.Fields("Qty") = CDbl("0" & Abs(rsTemp.Fields("Qty"))) + CDbl("0" & Abs(txtQtyTran.Text))
                        rsTemp.Fields("CostPer") = (rsTemp.Fields("CostPer") + .Fields("CostPer")) / 2
                        rsTemp.Fields("Amount") = rsTemp.Fields("CostPer") * rsTemp.Fields("Qty")
                        rsTemp.Update
                    End If
                End With
            Else
                With rsTontemp
                    .Find "ItemNum='" & strPLU & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        rsTemp.addNew
                        rsTemp.Fields("TransDocNo") = txtDocNum.Text
                        rsTemp.Fields("TransDate") = gfCONVERT_DATE_TO_STRING(dtpDateOut.Value)
                        rsTemp.Fields("Deliver") = txtDeliver.Text
                        rsTemp.Fields("Receiver") = txtReceiver.Text
                        rsTemp.Fields("Descriptions") = txtDescription.Text
                        rsTemp.Fields("Stock_ID") = rsTontemp.Fields("Stock_ID")
                        rsTemp.Fields("StockOrg") = Format(cboStockSource.ListIndex + 1, "00")
                        rsTemp.Fields("StockDes") = Format(cboStockDes.ListIndex + 1, "00")
                        rsTemp.Fields("ItemNum") = rsTontemp.Fields("ItemNum")
                        rsTemp.Fields("ItemName") = rsTontemp.Fields("ItemName")
                        rsTemp.Fields("Qty") = CDbl("0" & txtQtyTran.Text)
                        rsTemp.Fields("CostPer") = rsTontemp.Fields("CostPer")
                        rsTemp.Fields("Amount") = CDbl("0" & txtQtyTran.Text) * rsTontemp.Fields("CostPer")
                        rsTemp.Update
                        'cap nhat rsTontemp
                        .Fields("Qty") = .Fields("Qty") - CDbl("0" & txtQtyTran.Text)
                        .Fields("Amount") = .Fields("Qty") * .Fields("CostPer")
                        .Update
                      
                    End If
                End With
            
            End If
    End With
    Call InitTranDetail(rsTemp)
    Call InitDatagrid(rsTontemp)
    txtQtyTran.Text = ""
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdAddOne_Click"
End Sub

Private Sub cmdClose_Click()
    Set rsTemp = Nothing
    Unload Me
End Sub

Private Sub cmddelete_Click()
On Error GoTo Handle
    If MsgBox("B¹n cã ch¾c ch¾n muèn xãa chøng tõ nµy kg?", vbYesNo) = vbYes Then
        With rsTranStock
            .Find "TransDoc='" & txtDocNum.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Delete adAffectCurrent
                .Requery
                cnData.Execute "Delete  from Inventory_In" & Tonthang & " where Doc_Number='" & Trim(txtDocNum.Text) & "'"
            End If
        End With
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmddelete_Click"
End Sub

Private Sub cmdEdit_Click()
On Error GoTo Handle
    Call Init_Transtock
    
Exit Sub
Handle: MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub cmdKeyboard_Click()
    With frmKeyboard
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtSearch.Text = .Let_Text_Input
    End With
End Sub

Private Sub cmdNewDoc_Click()
On Error GoTo Handle
    Call addNew
    cmdNewDoc.Enabled = False
    txtReceiver.SetFocus
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - cmdNewDoc_Click"
End Sub

Private Sub cmdRemoveOne_Click()
Exit Sub
On Error GoTo Handle
    'Cap nhat trong bang tam
    With rsTemp
        .Find "ItemNum='" & txtItemCode.Text & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Qty") = .Fields("Qty") - txtQtyTran.Text
            .Update
            If .Fields("Qty") = 0 Then .Delete adAffectCurrent
        End If
    End With
    'Cap nhat lai trong kho chuyen di
    With rsTontemp
        .Find "ItemNum='" & txtItemCode.Text & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Quantity") = .Fields("Quantity") + txtQtyTran.Text
            .Update
        Else
            .addNew
            .Fields("ItemNum") = txtItemCode.Text
            With rsTemp
                .Find "ItemNum='" & txtItemCode.Text & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        rsTontemp.Fields("ItemName") = .Fields("ItemName")
                        rsTontemp.Fields("Unit") = .Fields("Unit")
                        rsTontemp.Fields("Stock_ID") = Format(cboStock_Type.ListIndex + 1, "00")
                        rsTontemp.Fields("Qty") = txtQtyTran.Text
                        rsTontemp.Fields("CostPer") = .Fields("CostPer")
                        rsTontemp.Fields("Amount") = .Fields("Amount")
                        .Update
                    End If
            End With
            .Update
        End If
    End With
Exit Sub

Handle:
MsgBox Err.Number & Err.Description & Me.name & " ChuyÓn kho ng­îc"
End Sub

Private Sub cmdSaveDoc_Click()
On Error GoTo Handle
    'Kiem tra chung tu goc cua bang nguon
    'Neu chua co  thi them moi vao
    
    Call check_Doc_Org(txtDocNum.Text)
    
    'Xoa toan bo du lieu co trong bang xuat cua kho nguon
    If Trim(txtDocNum.Text) <> "" Then
        cnData.Execute "Delete  from Inventory_In" & Tonthang & " where Doc_Number='" & txtDocNum.Text & "'"
    
        'Cap nhat du lieu tu recordset tam vao bang kho nguon
        
        Call Update_Org_Details(txtDocNum.Text)
        'Kiem  tra chung tu dich
        'Neu chua co  thi them vao
        Call check_Doc_Des(txtDocNum.Text)
        'Xoa toan bo du lieu cua bang dich
        cnData.Execute "Delete  from Inventory_InB" & Tonthang & " where Doc_Number='" & txtDocNum.Text & "'"
        'Cap nhat noi dung chung tu chuyen kho vao trong 2 kho
        Call Update_Des_Details(txtDocNum.Text)
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - L­u chøng tõ nhËp kho cã lçi"
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo Handle
    With rsTranStock
    .Find "TransDoc='" & txtDocNum.Text & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            
            .Fields("TransDate") = gfCONVERT_DATE_TO_STRING(dtpDateOut.Value)
            .Fields("TransPerson") = txtDeliver.Text
            .Fields("ReceivePerson") = txtReceiver.Text
            .Fields("Remark") = txtDescription.Text
            .Fields("SourceStock") = Format(cboStockSource.ListIndex + 1, "00")
            .Fields("DesStock") = Format(cboStockDes.ListIndex + 1, "00")
            .Fields("Stock_ID") = Format(cboStock_Type.ListIndex + 1, "00")
            .Update
        Else
            .addNew
            .Fields("TransDoc") = txtDocNum.Text
            .Fields("TransDate") = gfCONVERT_DATE_TO_STRING(dtpDateOut.Value)
            .Fields("TransPerson") = txtDeliver.Text
            .Fields("ReceivePerson") = txtReceiver.Text
            .Fields("Remark") = txtDescription.Text
            .Fields("SourceStock") = Format(cboStockSource.ListIndex + 1, "00")
            .Fields("DesStock") = Format(cboStockDes.ListIndex + 1, "00")
            .Fields("Stock_ID") = Format(cboStock_Type.ListIndex + 1, "00")
            .Update
            .Requery
        End If
    End With
    Call Init_Transtock
    cmdNewDoc.Enabled = True
    Call InitTranDoc
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub Form_Load()
On Error GoTo Handle
        Tonthang = Format(Month(Date), "00") & Right(Format(Year(Date), "00"), 2)

    With rsTemp
        If .State = 0 Then
                .Fields.Append "TransDocNo", adVarWChar, 13
                .Fields("TransDocNo").Attributes = adColNullable
                .Fields.Append "TransDate", adVarWChar, 12
                .Fields("TransDate").Attributes = adColNullable
                .Fields.Append "Stock_ID", adVarWChar, 2
                .Fields("Stock_ID").Attributes = adColNullable
                .Fields.Append "Deliver", adVarWChar, 200
                .Fields("Deliver").Attributes = adColNullable
                .Fields.Append "Receiver", adVarWChar, 200
                .Fields("Receiver").Attributes = adColNullable
                .Fields.Append "Descriptions", adVarWChar, 200
                .Fields("Descriptions").Attributes = adColNullable
                .Fields.Append "StockOrg", adVarWChar, 2
                .Fields("StockOrg").Attributes = adColNullable
                .Fields.Append "StockDes", adVarWChar, 2
                .Fields("StockDes").Attributes = adColNullable
                .Fields.Append "ItemNum", adVarWChar, 12
                .Fields.Append "ItemName", adVarWChar, 50
                .Fields.Append "Qty", adDouble
                .Fields.Append "CostPer", adDouble
                .Fields.Append "Amount", adDouble
                .Open
            End If
            Call Get_rsTemp(txtDocNum.Text)
    End With
    
    
    Set rsMasterDes = Open_Table(cnData, "Instock_MasterB")
    Set rsMasterOrg = Open_Table(cnData, "Inventory_In_Master")
    If Check_Table_exist("Inventory_In" & Tonthang) = False Then Call CreateTable_InStock(gfCONVERT_DATE_TO_STRING(Date))
    Set rsStockOrg = Open_Table(cnData, "Inventory_In" & Tonthang)
    If Check_Table_exist("Inventory_InB" & Tonthang) = False Then Call CreateTable_InStock(gfCONVERT_DATE_TO_STRING(Date))
    Set rsStockDes = Open_Table(cnData, "Inventory_InB" & Tonthang)
    
    InitTranDoc
    Call AddStockToCombo
    Call grdTranStock_Click
    cboStock_Type.ListIndex = 1
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub
Public Sub InitDatagrid(rs As Recordset)
    On Error GoTo Handle
    With grdPluSource
    If rs.State = 0 Then Exit Sub
          Set .DataSource = rs
            .Columns(0).Caption = "M· Hµng"
            .Columns(0).Width = 1400
            .Columns(1).Caption = "Tªn hµng"
            .Columns(1).Width = 2000
            .Columns(1).Alignment = dbgLeft
            .Columns(2).Caption = "§VT"
            .Columns(2).Width = 800
            .Columns(2).Alignment = dbgLeft
            .Columns(3).Caption = "Stock_ID"
            .Columns(3).Width = 0
            .Columns(4).Caption = "S.L tån"
            .Columns(4).Width = 1000
            .Columns(4).Alignment = dbgLeft
            .Columns(5).Caption = "§¬n gi¸"
            .Columns(5).Alignment = dbgLeft
            .Columns(5).Width = 1000
            .Columns(6).Caption = "Thµnh tiÒn"
            .Columns(6).Alignment = dbgCenter
            .Columns(6).Width = 1500
       End With
       Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " InitDatagrid"
End Sub


Private Sub grdPluDes_Click()
    txtQtyTran.SetFocus
    txtItemCode.Text = grdPluDes.Columns(8).Value
End Sub

Private Sub grdPluSource_Click()
    On Error GoTo Handle
        txtQtyTran.SetFocus
        strPLU = grdPluSource.Columns(0).Value
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - grdPluSource_Click"
End Sub

Private Sub grdTranStock_Click()
On Error GoTo Handle
Set rsTemp = Nothing
    txtDocNum.Text = Trim("" & grdTranStock.Columns(0).Value)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - grdTranStock_Click"
End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cboStockSource.SetFocus

End Sub

Private Sub txtDocNum_Change()
On Error GoTo Handle
    With rsTranStock
        .Find "TransDoc='" & txtDocNum.Text & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            dtpDateOut.Value = gfCONVERT_STRING_TO_DATE(.Fields("TransDate"))
            txtReceiver.Text = .Fields("ReceivePerson")
            txtDeliver.Text = .Fields("TransPerson")
            txtDescription.Text = .Fields("Remark")
            cboStockSource.ListIndex = CInt(.Fields("SourceStock")) - 1
            cboStockDes.ListIndex = CInt(.Fields("DesStock")) - 1
            cboStock_Type.ListIndex = CInt(.Fields("Stock_ID")) - 1
        End If
    End With
    Call Get_rsTemp(txtDocNum.Text)
    Call InitTranDetail(rsTemp)
'    If MsgBox("Chøng tõ nµy ®· thay ®æi, b¹n cã muèn l­u kh«ng?", vbYesNo) = vbYes Then
'        Call cmdSaveDoc_Click
'    End If
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - txtDocNum_Change"
End Sub

Private Sub txtReceiver_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtDescription.SetFocus

End Sub

Private Sub txtSearch_Change()
On Error GoTo errHdl
Dim rsTempGrid As New ADODB.Recordset
'On Error GoTo HandlEErr
With rsTempGrid
    .ActiveConnection = cnData
    .CursorLocation = adUseClient
    .LockType = adLockOptimistic
    .CursorType = adOpenDynamic
    '.Open
End With
'        Call RecreateVStockDB
        With rsTempGrid
            If .State = adStateOpen Then .Close
            
            If InStr(1, Trim(txtSearch.Text), "*", vbTextCompare) > 0 Then GoTo 1
            
                .Open "SELECT  ItemNum, Description, Unit,Stock_ID, Quantity,CostPer,Amount FROM TonA" & Tonthang & " WHERE " & _
                 "INSTR(ItemNum,""" & txtSearch.Text & """)>0 OR INSTR(Description,""" & _
                txtSearch.Text & """)>0 OR INSTR(Unit,""" & txtSearch.Text & """)>0  " & _
                " ORDER BY ItemNum"
            
            GoTo 2
1:
                .Open "SELECT  ItemNum, Description, Unit,Stock_ID, Quantity,CostPer,Amount FROM Customer WHERE " & _
                "(INSTR(ItemNum,""" & Trim(txtSearch.Text) & """)>0 OR Description LIKE '" & _
                Left(Trim(txtSearch.Text), Len(Trim(txtSearch.Text)) - 1) & "%') OR Unit LIKE '" & _
                Left(Trim(txtSearch.Text), Len(Trim(txtSearch.Text)) - 1) & "%')" & _
                " ORDER BY ItemNum"
2:
        End With
        Call InitDatagrid(rsTempGrid)
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "txtSearch_Change"
End Sub

Public Sub AddStockToCombo()
On Error GoTo Handle
    Dim rsStock As New ADODB.Recordset
    Set rsStock = Open_Table(cnData, "Stock")
    If rsStock.State <> 0 Then
        If rsStock.RecordCount > 0 Then rsStock.MoveFirst
        With cboStockDes
            .Clear
            Do While Not rsStock.EOF
                .AddItem rsStock.Fields("StockName")
            rsStock.MoveNext
            Loop
        End With
        If rsStock.RecordCount > 0 Then rsStock.MoveFirst
        With cboStockSource
            .Clear
            Do While Not rsStock.EOF
                .AddItem rsStock.Fields("StockName")
            rsStock.MoveNext
            Loop
        End With
    End If
    cboStockSource.ListIndex = 0
    cboStockDes.ListIndex = 1
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Public Sub addNew()
On Error GoTo Handle
Dim rsMaxdoc As New ADODB.Recordset
Set rsMaxdoc = OpenCriticalTable("select Max(TransDoc)as MaxDoc from Transfer_Doc where left(TransDoc,9)='T" & gfCONVERT_DATE_TO_STRING(Date) & "'", cnData)
    If rsMaxdoc.RecordCount > 0 Then
        If Not rsMaxdoc.EOF Then
            rsMaxdoc.MoveLast
            If rsMaxdoc.Fields("MaxDoc") & "" = "" Then
                txtDocNum.Text = "T" & gfCONVERT_DATE_TO_STRING(Date) & "0001"
            Else
                txtDocNum.Text = Left(rsMaxdoc.Fields("MaxDoc"), 9) & Format(Int(Right(rsMaxdoc.Fields("MaxDoc"), 4)) + 1, "0000")
            End If
        Else
            txtDocNum.Text = "T" & gfCONVERT_DATE_TO_STRING(Date) & "0001"
        End If
    End If
    txtDeliver.Text = userName
    txtDocNum.Locked = True
    dtpDateOut.Value = Date
    txtReceiver.Locked = False
    txtDescription.Locked = False
    txtReceiver.Text = ""
    txtDescription.Text = ""
    Call AddStockToCombo
    cmdUpdate.Enabled = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - addNew"
End Sub

Public Sub InitTranDoc()
    On Error GoTo Handle
    Set rsTranStock = Open_Table(cnData, "Transfer_Doc")
    
    With grdTranStock
          Set .DataSource = rsTranStock
            .Columns(0).Caption = "Sè CT"
            .Columns(0).Width = 1800
            .Columns(1).Caption = "Ngµy CT"
            .Columns(1).Width = 1500
            .Columns(1).Alignment = dbgLeft
            .Columns(2).Caption = "Ng­êi chuyÓn"
            .Columns(2).Width = 1500
            .Columns(2).Alignment = dbgLeft
            .Columns(3).Caption = "Ng­êi nhËn"
            .Columns(3).Width = 1500
            .Columns(3).Alignment = dbgLeft
            .Columns(4).Caption = "DiÔn gi¶i"
            .Columns(4).Alignment = dbgLeft
            .Columns(4).Width = 3000
            .Columns(5).Caption = "Kho nguån"
            .Columns(5).Alignment = dbgCenter
            .Columns(5).Width = 1600
            .Columns(6).Caption = "Kho ®Ých"
            .Columns(6).Alignment = dbgCenter
            .Columns(6).Width = 1600
            .Columns(7).Width = 0
       End With
       Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " InitTranDoc"
End Sub
Public Sub InitTranDetail(rs As Recordset)
    On Error GoTo Handle
    With grdPluDes
          Set .DataSource = rs
            .Columns(0).Width = 0
            .Columns(1).Width = 0
            .Columns(2).Width = 0
            .Columns(3).Width = 0
            .Columns(4).Width = 0
            .Columns(5).Width = 0
            .Columns(6).Width = 0
            .Columns(7).Width = 0
            
            .Columns(8).Caption = "M· hµng"
            .Columns(8).Width = 1800
            .Columns(9).Caption = "Tªn hµng"
            .Columns(9).Width = 2500
            .Columns(9).Alignment = dbgLeft
            .Columns(10).Caption = "Sè l­îng"""
            .Columns(10).Width = 1000
            .Columns(10).Alignment = dbgLeft
            .Columns(11).Caption = "§¬n gi¸"
            .Columns(11).Width = 1500
            .Columns(11).Alignment = dbgLeft
            .Columns(12).Caption = "Thµnh tiÒn"
            .Columns(12).Alignment = dbgLeft
            .Columns(12).Width = 1500
       End With
       Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " InitTranDetail"
End Sub

Public Sub Get_rsTemp(DocNum As String)
On Error GoTo Handle
    Dim rsItemDoc As New ADODB.Recordset
    Dim SQL As String
    SQL = "SELECT *" & _
         " FROM Inventory_In_Master INNER JOIN Inventory_In" & Tonthang & " ON Inventory_In_Master.Doc_Number = Inventory_In" & Tonthang & _
         ".Doc_Number WHERE (((Inventory_In_Master.Doc_Number)='" & txtDocNum.Text & "'))"
    'Set rsTemp = New ADODB.Recordset
    Set rsItemDoc = OpenCriticalTable(SQL, cnData)
    With rsTemp
        If .State = 0 Then
            .Fields.Append "TransDocNo", adVarWChar, 13
            .Fields("TransDocNo").Attributes = adColNullable
            .Fields.Append "TransDate", adVarWChar, 12
            .Fields("TransDate").Attributes = adColNullable
            .Fields.Append "Stock_ID", adVarWChar, 2
            .Fields("Stock_ID").Attributes = adColNullable
            .Fields.Append "Deliver", adVarWChar, 200
            .Fields("Deliver").Attributes = adColNullable
            .Fields.Append "Receiver", adVarWChar, 200
            .Fields("Receiver").Attributes = adColNullable
            .Fields.Append "Descriptions", adVarWChar, 200
            .Fields("Descriptions").Attributes = adColNullable
            .Fields.Append "StockOrg", adVarWChar, 2
            .Fields("StockOrg").Attributes = adColNullable
            .Fields.Append "StockDes", adVarWChar, 2
            .Fields("StockDes").Attributes = adColNullable
            .Fields.Append "ItemNum", adVarWChar, 12
            .Fields.Append "ItemName", adVarWChar, 50
            .Fields.Append "Qty", adDouble
            .Fields.Append "CostPer", adDouble
            .Fields.Append "Amount", adDouble
            .Open
        End If
            
        Do While Not rsItemDoc.EOF
            .addNew
            .Fields("TransDocNo") = txtDocNum.Text
            .Fields("TransDate") = gfCONVERT_DATE_TO_STRING(dtpDateOut.Value)
            .Fields("Deliver") = txtDeliver.Text
            .Fields("Receiver") = txtReceiver.Text
            .Fields("Descriptions") = txtDescription.Text
            .Fields("Stock_ID") = rsItemDoc.Fields("Stock_ID")
            .Fields("StockOrg") = Format(cboStockSource.ListIndex + 1, "00")
            .Fields("StockDes") = Format(cboStockDes.ListIndex + 1, "00")
            .Fields("ItemNum") = rsItemDoc.Fields("ItemNum")
            .Fields("ItemName") = rsItemDoc.Fields("Description")
            .Fields("Qty") = Abs(rsItemDoc.Fields("Quantity"))
            .Fields("CostPer") = Abs(rsItemDoc.Fields("CostPer"))
            .Fields("Amount") = Abs(rsItemDoc.Fields("Amount"))
            .Update
            '.Requery
        rsItemDoc.MoveNext
        Loop
    End With
Exit Sub
Handle:
MsgBox Err.Description & Me.name & " - Get_rsTemp"
End Sub

Public Sub Init_Transtock()
On Error GoTo Handle
    Set rsTontemp = Nothing
    If cboStockSource.ListIndex = 0 Then
        Set rsTon = OpenCriticalTable("SELECT  ItemNum, Description, Unit,Stock_ID, Quantity,CostPer,Amount FROM TonA" & Tonthang & " where Stock_ID='" & Format((cboStock_Type.ListIndex + 1), "00") & "'", cnData)
    Else
        Set rsTon = OpenCriticalTable("SELECT  ItemNum, Description, Unit,Stock_ID, Quantity,CostPer,Amount FROM TonB" & Tonthang & " where Stock_ID='" & Format((cboStock_Type.ListIndex + 1), "00") & "'", cnData)
    End If
    With rsTon
        Do While Not .EOF
            With rsTontemp
                If .State = 0 Then
                    .Fields.Append "ItemNum", adVarWChar, 12
                    .Fields.Append "ItemName", adVarWChar, 50
                    .Fields("ItemName").Attributes = adColNullable
                    .Fields.Append "Unit", adVarWChar, 50
                    .Fields("Unit").Attributes = adColNullable
                    .Fields.Append "Stock_ID", adVarWChar, 2
                    .Fields("Stock_ID").Attributes = adColNullable
                    .Fields.Append "Qty", adDouble
                    .Fields.Append "CostPer", adDouble
                    .Fields.Append "Amount", adDouble
                    .Open
                End If
                .addNew
                .Fields("ItemNum") = rsTon.Fields("ItemNum")
                .Fields("ItemName") = rsTon.Fields("Description") & ""
                .Fields("Unit") = rsTon.Fields("Unit") & ""
                .Fields("Stock_ID") = rsTon.Fields("Stock_ID")
                .Fields("Qty") = rsTon.Fields("Quantity")
                .Fields("CostPer") = rsTon.Fields("CostPer")
                .Fields("Amount") = rsTon.Fields("Amount")
                .Update
            End With
        
        .MoveNext
        Loop
    End With
    
    Call InitDatagrid(rsTontemp)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Init_Transtock"
End Sub

Public Sub check_Doc_Org(ByVal strCondition As String)
On Error GoTo Handle
    With rsMasterOrg
        .Find "Doc_Number='" & strCondition & "'", , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields("Doc_Number") = txtDocNum.Text
            .Fields("Stock_ID") = Format(cboStock_Type.ListIndex + 1, "00")
            .Fields("DateTime") = gfCONVERT_DATE_TO_STRING(dtpDateOut.Value)
            .Fields("Cashier_ID") = UserID
            .Fields("Receiver_person") = txtReceiver.Text
            .Fields("iReason") = "XK"
            .Fields("Store_ID") = Store_ID
            .Fields("InOutType") = "T"
            .Update
            
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & " - Kiem tra chung tu nguon co !loi "
End Sub

Public Sub check_Doc_Des(ByVal strCondition As String)
On Error GoTo Handle
    With rsMasterDes
        .Find "Doc_Number='" & strCondition & "'", , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields("Doc_Number") = txtDocNum.Text
            .Fields("Stock_ID") = Format(cboStock_Type.ListIndex + 1, "00")
            .Fields("DateTime") = gfCONVERT_DATE_TO_STRING(dtpDateOut.Value)
            .Fields("Cashier_ID") = UserID
            .Fields("Receiver_person") = txtReceiver.Text
            .Fields("iReason") = "XK"
            .Fields("Store_ID") = Store_ID
            .Fields("InOutType") = "T"
            .Update
            
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & " - Kiem tra chung tu dich co loi "
End Sub

Public Sub Update_Org_Details(ByVal strDoc As String)
    On Error GoTo Handle
    If Trim(strDoc) <> "" Then
        cnData.Execute "Delete  from Inventory_In" & Tonthang & " where Doc_Number='" & strDoc & "'"
        With rsStockOrg
            With rsTemp
            If rsTemp.State <> 0 Then
                If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                    Do While Not .EOF
                        rsStockOrg.addNew
                        rsStockOrg.Fields("Doc_Number") = .Fields("TransDocNo")
                        rsStockOrg.Fields("DateTime") = gfCONVERT_DATE_TO_STRING(dtpDateOut.Value)
                        rsStockOrg.Fields("ItemNum") = .Fields("ItemNum")
                        rsStockOrg.Fields("Description") = .Fields("ItemName")
                        rsStockOrg.Fields("Store_ID") = Store_ID
                        rsStockOrg.Fields("Quantity") = -.Fields("Qty")
                        rsStockOrg.Fields("CostPer") = .Fields("CostPer")
                        rsStockOrg.Fields("Amount") = -.Fields("Amount")
                        rsStockOrg.Update
                        rsStockOrg.Requery
                    .MoveNext
                    Loop
                Else
                    Exit Sub
                End If
            End With
        End With
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & " - Luu chi tiet chuyen kho xuong bang nguoi"
End Sub

Public Sub Update_Des_Details(ByVal strDoc As String)
    On Error GoTo Handle
    If Trim(strDoc) <> "" Then
        cnData.Execute "Delete  from Inventory_InB" & Tonthang & " where Doc_Number='" & strDoc & "'"
        With rsStockDes
            With rsTemp
            If rsTemp.State <> 0 Then
                If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                    Do While Not .EOF
                        rsStockDes.addNew
                        rsStockDes.Fields("Doc_Number") = .Fields("TransDocNo")
                        rsStockDes.Fields("DateTime") = gfCONVERT_DATE_TO_STRING(dtpDateOut.Value)
                        rsStockDes.Fields("ItemNum") = .Fields("ItemNum")
                        rsStockDes.Fields("Description") = .Fields("ItemName")
                        rsStockDes.Fields("Store_ID") = Store_ID
                        rsStockDes.Fields("Quantity") = .Fields("Qty")
                        rsStockDes.Fields("CostPer") = .Fields("CostPer")
                        rsStockDes.Fields("Amount") = .Fields("Amount")
                        rsStockDes.Update
                        rsStockDes.Requery
                    .MoveNext
                    Loop
                Else
                    Exit Sub
                End If
            End With
        End With
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & " - Luu chi tiet chuyen kho xuong bang nguoi"
End Sub
