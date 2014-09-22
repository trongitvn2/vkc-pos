VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmReservered 
   Caption         =   "§Æt bµn"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16230
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
   ScaleWidth      =   16230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   33
      Top             =   5280
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Chi tiÕt ®Æt tiÖc"
      TabPicture(0)   =   "frmReservered.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   4215
         Left            =   120
         TabIndex        =   34
         Top             =   360
         Width           =   9015
         Begin MSDataGridLib.DataGrid DataGrid1 
            Height          =   3855
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   6800
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
   End
   Begin prjTouchScreen.MyButton cmdPrint 
      Height          =   855
      Left            =   10440
      TabIndex        =   25
      Top             =   10080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "&In phiÕu"
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
      BCOL            =   14737632
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmReservered.frx":001C
      PICN            =   "frmReservered.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   855
      Left            =   5520
      TabIndex        =   24
      Top             =   10080
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "&Hñy ®Æt"
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
      BCOL            =   14737632
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReservered.frx":04AC
      PICN            =   "frmReservered.frx":04C8
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
      Height          =   855
      Left            =   12960
      TabIndex        =   23
      Top             =   10080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "§ãn&g"
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
      BCOL            =   14737632
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmReservered.frx":0B02
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdSave 
      Height          =   855
      Left            =   3240
      TabIndex        =   9
      Top             =   10080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "&L­u"
      ENAB            =   0   'False
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
      BCOL            =   14737632
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReservered.frx":0B1E
      PICN            =   "frmReservered.frx":0B3A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdNew 
      Height          =   855
      Left            =   960
      TabIndex        =   22
      Top             =   10080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "Thªm míi ®Æt chç"
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
      BCOL            =   14737632
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReservered.frx":107E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Th«ng tin ®Æt bµn"
      ForeColor       =   &H00FF0000&
      Height          =   9855
      Left            =   9480
      TabIndex        =   11
      Top             =   120
      Width           =   6735
      Begin VB.TextBox txtSection_ID 
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   4560
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkIsUse 
         Caption         =   "§· sö dông"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4800
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox chkAll 
         Caption         =   "TÊt c¶ c¸c ngµy"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1560
         TabIndex        =   31
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2850
         Index           =   7
         Left            =   720
         TabIndex        =   30
         Top             =   6840
         Width           =   5775
      End
      Begin VB.TextBox txtText 
         Height          =   450
         Index           =   1
         Left            =   2040
         TabIndex        =   1
         Top             =   1440
         Width           =   4575
      End
      Begin prjTouchScreen.MyButton cmdOrder 
         Height          =   735
         Left            =   4680
         TabIndex        =   27
         Top             =   5280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1296
         BTYPE           =   5
         TX              =   "§Æt mãn..."
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
         BCOL            =   14737632
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmReservered.frx":109A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdTable 
         Height          =   735
         Left            =   4680
         TabIndex        =   26
         Top             =   4440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1296
         BTYPE           =   5
         TX              =   "Chän bµn..."
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
         BCOL            =   14737632
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmReservered.frx":10B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   2040
         TabIndex        =   8
         Text            =   "500,000"
         Top             =   5160
         Width           =   2535
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   5
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   4560
         Width           =   2535
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   4
         Left            =   2040
         TabIndex        =   4
         Top             =   3165
         Width           =   1215
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   3
         Left            =   2040
         TabIndex        =   3
         Top             =   2640
         Width           =   4575
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   2
         Left            =   2040
         TabIndex        =   2
         Top             =   2040
         Width           =   4575
      End
      Begin VB.TextBox txtText 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   0
         Left            =   2040
         TabIndex        =   0
         Top             =   360
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   615
         Left            =   4440
         TabIndex        =   6
         Top             =   3720
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   16711680
         CalendarTitleForeColor=   16711680
         Format          =   65011713
         UpDown          =   -1  'True
         CurrentDate     =   40157
      End
      Begin MSComCtl2.DTPicker dtpTime 
         Height          =   615
         Left            =   2040
         TabIndex        =   5
         Top             =   3720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   16711680
         CalendarTitleForeColor=   16711680
         Format          =   65011714
         UpDown          =   -1  'True
         CurrentDate     =   40157
      End
      Begin VB.Label Label10 
         Caption         =   "Ghi chó:"
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
         TabIndex        =   29
         Top             =   6480
         Width           =   975
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "M· §Æt bµn:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label lblDocso 
         Caption         =   "N¨m tr¨m ngh×n ®ång ch½n ./."
         BeginProperty Font 
            Name            =   ".VnArial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   21
         Top             =   6120
         Width           =   5175
      End
      Begin VB.Label label 
         Caption         =   "B»ng ch÷:"
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
         TabIndex        =   20
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Ngµy:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3840
         TabIndex        =   19
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Vµo lóc:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Sè tiÒn cäc:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Bµn ®Æt:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Sè kh¸ch:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "§iÖn tho¹i liªn hÖ:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "§Þa chØ:"
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
         TabIndex        =   13
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Hä tªn ng­êi ®Æt:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1815
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FlgReserered 
      Height          =   5295
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9340
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
   End
   Begin prjTouchScreen.MyButton cmdEdit 
      Height          =   855
      Left            =   7920
      TabIndex        =   32
      Top             =   10080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "Söa ch÷a phiÕu"
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
      BCOL            =   14737632
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmReservered.frx":10D2
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
Attribute VB_Name = "frmReservered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTable_Reserered As New ADODB.Recordset
Dim rsDetails As New ADODB.Recordset
Dim strSql As String
Dim isOK As Boolean
Dim rsReserve_Detail As New ADODB.Recordset

Private Sub chkAll_Click()
On Error GoTo Handle
  If chkAll.Value = False Then
    strSql = "SELECT Table_Reservered.Reservered_Code,Table_Reservered.DateTime," & _
            " Table_Reservered.CustName, Table_Reservered.CustName," & _
            " Table_Reservered.Address, Table_Reservered.Phone," & _
            " Table_Reservered.Seat_Num, Table_Reservered.Date_Reservered," & _
            " Table_Reservered.Time_Reservered, Table_Reservered.Table_ID," & _
            " Table_Reservered.Amount, Table_Reservered.Description," & _
            " Table_Reservered.Date_Reservered,Table_Reservered.Cashier_ID,Table_Reservered.Section_ID,Table_Reservered.IsUsed" & _
            " From Table_Reservered" & _
            " WHERE (((Table_Reservered.Date_Reservered)='" & Format(Date, "dd/MM/yyyy") & "'))"
    
    Set rsTable_Reserered = OpenCriticalTable(strSql, cnData)
    
    Call InitData_FLGRIDORDER(rsTable_Reserered)
    
  Else
    strSql = "SELECT Table_Reservered.Reservered_Code,Table_Reservered.DateTime," & _
            " Table_Reservered.CustName, Table_Reservered.CustName," & _
            " Table_Reservered.Address, Table_Reservered.Phone," & _
            " Table_Reservered.Seat_Num, Table_Reservered.Date_Reservered," & _
            " Table_Reservered.Time_Reservered, Table_Reservered.Table_ID," & _
            " Table_Reservered.Amount, Table_Reservered.Description,Table_Reservered.Section_ID," & _
            " Table_Reservered.Date_Reservered,Table_Reservered.Cashier_ID,Table_Reservered.IsUsed" & _
            " From Table_Reservered"
            
     Set rsTable_Reserered = OpenCriticalTable(strSql, cnData)
        
        Call InitData_FLGRIDORDER(rsTable_Reserered)
    
  End If
  
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " chkAll_Click"
End Sub

Private Sub chkIsUse_Click()
On Error GoTo Handle
    With rsTable_Reserered
        .Find "Reservered_Code='" & txtText(0).Text & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If chkIsUse.Value = 1 Then
                .Fields("IsUsed") = -1
            Else
                .Fields("IsUsed") = 0
            End If
            .Update
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub cmdCancel_Click()
On Error GoTo Handle
If MsgBox("B¹n cã ch¾c ch¾n muèn hñy phiÕu ®Æt cäc nµy kh«ng?", vbYesNo) = vbYes Then
    'Hñy phiÕu ®Æt cäc ®· l­u
    If Cancel_Reserve(txtText(0).Text) Then
    'T¹o mét phiÕu chi tiÒn ®· ®Æt cäc
        Call Payment(txtText(0).Text)
    End If
End If
Call Set_FlgReserered
Set DataGrid1.DataSource = Nothing
Exit Sub
Handle:
MsgBox "Hñy ®Æt chç :" & Err.Description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    Call UnLock_text
End Sub

Private Sub cmdNew_Click()
On Error GoTo Handle
    Init_AddNew
    Un_Lock_text
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdNew_Click"
End Sub

Private Sub cmdOrder_Click()
On Error GoTo Handle
    If Not Check_Table_exist("Table_Reserved_Details") Then
        Call Create_Table_Reserved_Details
    End If
    'kiem tra xem da khoi tao menu chi tiet chua?
    With rsDetails
        If .State = 0 Then
            .Fields.Append "Reserve_Code", adVarWChar, 20
            .Fields.Append "TableNo", adVarWChar, 20
            .Fields.Append "PluNo", adWChar, 12
            .Fields.Append "PluName", adVarWChar, 100
            .Fields.Append "Qty", adDouble
            .Fields.Append "Price", adDouble
            .Fields.Append "Amt", adDouble
            .Fields.Append "Description", adVarWChar, 50
            .Open
        End If
    End With
'goi form order
    With frmMenuSelect
        .Get_Records = rsDetails
        .Show vbModal
        Set rsDetails = .return_Recordset
        isOK = .Let_OK
    End With
    If isOK = True Then
      Set DataGrid1.DataSource = rsDetails
        With DataGrid1
            .Columns(0).Width = 0
            .Columns(1).Width = 0
            .Columns(2).Width = 0
            .Columns(3).Width = 0
        End With
    End If
    'Call FlgReserered_Click
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdOrder_Click"
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL As String

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
    " Where Table_Reservered.Reservered_Code='" & txtText(0).Text & "'" & _
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
    MsgBox Err.Number & Err.Description & Me.name & " cmdPrint_Click"
End Sub

Private Sub cmdSave_Click()
On Error GoTo Handle
    If Save_Data Then
        'Hái cã l­u kh¸ch hµng nµy vµo Danh môc kh¸ch hµng th©n thiÕt kh«ng
        If MsgBox("B¹n cã muèn l­u kh¸ch hµng nµy vµo danh môc kh¸ch hµng kh«ng?", vbYesNo) = vbYes Then
            Call Save_Customer
        End If
        ' L­u chi tiÕt ®Æt mãn
        Call save_details(rsDetails)
        'T¹o phiÕu thu tiÒn mÆt víi phiÕu ®Æt chç
        Call Save_Incom(txtText(0).Text)
         Call InitData_FLGRIDORDER(rsTable_Reserered)
        Call FlgReserered_Click
    End If
    cmdNew.Enabled = True

Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdSave_Click"
End Sub

Private Sub cmdTable_Click()
On Error GoTo Handle
    With frmTableSelect
        .Show vbModal
        txtText(5).Text = .Let_Table_Num
        txtSection_ID.Text = .Let_SectionID
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdTable_Click"
End Sub

Private Sub FlgReserered_Click()
On Error GoTo Handle
    txtText(0).Text = FlgReserered.TextMatrix(FlgReserered.Row, 0)
    Set rsDetails = Get_Recorset_By_Key(txtText(0).Text)
    Call Load_Details(rsDetails)
    Call Init_Header_Grids

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " FlgReserered_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    If Check_Table_exist("Table_Reservered") = False Then
        Call Create_Table_Reserverd
    End If
    Set rsReserve_Detail = Open_Table(cnData, "Table_Reserved_Details")
    
    strSql = "SELECT Table_Reservered.Reservered_Code,Table_Reservered.DateTime," & _
            " Table_Reservered.CustName, Table_Reservered.CustName," & _
            " Table_Reservered.Address, Table_Reservered.Phone," & _
            " Table_Reservered.Seat_Num, Table_Reservered.Date_Reservered," & _
            " Table_Reservered.Time_Reservered, Table_Reservered.Table_ID," & _
            " Table_Reservered.Amount, Table_Reservered.Description," & _
            " Table_Reservered.Date_Reservered,Table_Reservered.Cashier_ID,Table_Reservered.Section_ID,Table_Reservered.IsUsed" & _
            " From Table_Reservered" & _
            " WHERE (((Table_Reservered.Date_Reservered)='" & Format(Date, "dd/MM/yyyy") & "'))"
    Set rsTable_Reserered = OpenCriticalTable(strSql, cnData)
    
    Call Set_FlgReserered
    Call Lock_text
    Call InitData_FLGRIDORDER(rsTable_Reserered)
    If UserID = "131112" Then chkIsUse.Visible = True
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Public Sub Init_AddNew()
On Error GoTo Handle
    For i = 0 To txtText.count - 1
        txtText(i).Text = ""
    Next
    txtText(0).Text = "P§C/" & Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") & Format(Hour(time), "00") & Format(Minute(time), "00") & Format(Second(time), "00")
    dtpTime.Value = Format(Now, "HH:mm:ss")
    dtpDate.Value = Format(Date, "dd/MM/yyyy")
    cmdSave.Enabled = True
    txtText(0).SetFocus
    cmdNew.Enabled = False
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Init_AddNew"
End Sub

Private Sub txtText_Change(Index As Integer)
On Error GoTo Handle
    Select Case Index
        Case 0
            With rsTable_Reserered
                .Find "Reservered_Code='" & txtText(0).Text & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    txtText(1).Text = .Fields("CustName")
                    txtText(2).Text = .Fields("Address")
                    txtText(3).Text = .Fields("Phone")
                    txtText(4).Text = .Fields("Seat_Num")
                    dtpDate.Value = Format(.Fields("Date_Reservered"), "dd/MM/yyyy")
                    dtpTime.Value = .Fields("Time_Reservered")
                    txtText(5).Text = .Fields("Table_ID")
                    txtText(6).Text = .Fields("Amount")
                    txtText(7).Text = .Fields("Description")
                    'chkIsUse.Value = .Fields("IsUsed")
                    If .Fields("IsUsed") = True Then
                        chkIsUse.Value = 1
                    Else
                        chkIsUse.Value = 0
                    End If
                End If
            End With
        Case 6
            txtText(Index).SelStart = Len(txtText(Index).Text)
            txtText(Index).Text = Format(txtText(Index).Text, "#,##0")
            lblDocso.Caption = readnumber(CDbl("0" & txtText(Index).Text)) & " ®ång ./."
        Case Else
            txtText(Index).BackColor = &H80000005
    End Select
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtText_Change"
End Sub

Public Sub Lock_text()
On Error GoTo Handle
     For i = 0 To txtText.count - 1
        txtText(i).Locked = True
    Next
    dtpTime.Enabled = False
    dtpDate.Enabled = False
    cmdTable.Enabled = False
    cmdOrder.Enabled = False
    cmdSave.Enabled = False
    cmdCancel.Enabled = True
    cmdNew.Enabled = True
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Lock_text"
End Sub

Public Sub UnLock_text()
On Error GoTo Handle
     For i = 0 To txtText.count - 1
        txtText(i).Locked = False
    Next
    dtpTime.Enabled = True
    dtpDate.Enabled = True
    cmdTable.Enabled = True
    cmdOrder.Enabled = True
    cmdNew.Enabled = False
    cmdSave.Enabled = True
    cmdCancel.Enabled = False
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnLock_text"
End Sub

Public Sub Un_Lock_text()
On Error GoTo Handle
     For i = 0 To txtText.count - 1
        If i <> 5 Then txtText(i).Locked = False
    Next
    cmdTable.Enabled = True
    cmdOrder.Enabled = True
    dtpTime.Enabled = True
    dtpDate.Enabled = True
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Un_Lock_text"
End Sub

Public Sub Set_FlgReserered()
On Error GoTo Handle
    Dim i As Integer
        With FlgReserered
            .Cols = 9
            .Rows = 5
            .ColWidth(0) = 1200
            .ColWidth(1) = 1800
            .ColWidth(2) = 2500
            .ColWidth(3) = 1250
            .ColWidth(4) = 1250
            .ColWidth(5) = 1200
            .ColWidth(6) = 1200
            .ColWidth(7) = 1200
            .ColWidth(8) = 1000
            .TextMatrix(0, 0) = "M· phiÓu ®Æt"
            .TextMatrix(0, 1) = "Tªn KH"
            .TextMatrix(0, 2) = "Ñòa chæ"
            .TextMatrix(0, 3) = "S§T"
            .TextMatrix(0, 4) = "Ngµy"
            .TextMatrix(0, 5) = "Giê"
            .TextMatrix(0, 6) = "Bµn"
            .TextMatrix(0, 7) = "TiÒn cäc"
            .TextMatrix(0, 8) = "Sè kh¸ch"
            .ColAlignment(1) = 4
            .ColAlignment(2) = 4
            .ColAlignment(3) = 4
            .ColAlignment(4) = 4
            .ColAlignment(0) = 4
            .ColAlignment(5) = 4
            .ColAlignment(6) = 4
            .ColAlignment(7) = 4
            .ColAlignment(8) = 4
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
            .TextMatrix(1, 4) = ""
            .TextMatrix(1, 5) = ""
            .TextMatrix(1, 6) = ""
            .TextMatrix(1, 7) = ""
            .TextMatrix(1, 8) = ""
            
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Set_FlgReserered"
End Sub
Public Sub InitData_FLGRIDORDER(rs As ADODB.Recordset)
On Error GoTo Handle
        Dim incount As Integer
        If rs.EOF Then Exit Sub
        rs.MoveFirst
        With rs
            Do While Not .EOF
                incount = incount + 1
                FlgReserered.Rows = rs.RecordCount + 1
                With FlgReserered
                    .TextMatrix(incount, 0) = rs!Reservered_Code
                    .TextMatrix(incount, 1) = rs!CustName
                    .TextMatrix(incount, 2) = rs!Address
                    .TextMatrix(incount, 3) = rs!Phone
                    .TextMatrix(incount, 4) = rs!Date_Reservered
                    .TextMatrix(incount, 5) = rs!Time_Reservered
                    .TextMatrix(incount, 6) = rs!Table_ID
                    .TextMatrix(incount, 7) = rs!Amount
                    .TextMatrix(incount, 8) = rs!Seat_Num
                End With
            rs.MoveNext
            Loop
        End With
        If rs.RecordCount = 0 Then
            With FlgReserered
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
                .TextMatrix(1, 4) = ""
                .TextMatrix(1, 5) = ""
            End With
        End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - SetFLGRIDORDER"
End Sub

Public Function Save_Data() As Boolean
On Error GoTo Handle
    If Not Check_Null Then Exit Function
    With rsTable_Reserered
        .Find "Reservered_Code='" & txtText(0).Text & "'", , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields("Reservered_Code") = txtText(0).Text
            .Fields("DateTime") = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") & Format(Now, "HH:mm:ss")
            .Fields("CustName") = txtText(1).Text
            .Fields("Address") = txtText(2).Text
            .Fields("Phone") = txtText(3).Text
            .Fields("Seat_Num") = txtText(4).Text
            .Fields("Date_Reservered") = Format(dtpDate.Value, "dd/MM/yyyy")
            .Fields("Time_Reservered") = dtpTime.Value
            .Fields("Table_ID") = txtText(5).Text
            .Fields("Amount") = txtText(6).Text
            .Fields("Section_ID") = txtSection_ID.Text
            .Fields("Description") = txtText(7).Text
            .Fields("Cashier_ID") = UserID
            .Fields("IsUsed") = 0
            .Update
        Else
            .Fields("DateTime") = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") & Format(Now, "HH:mm:ss")
            .Fields("CustName") = txtText(1).Text
            .Fields("Address") = txtText(2).Text
            .Fields("Phone") = txtText(3).Text
            .Fields("Seat_Num") = txtText(4).Text
            .Fields("Date_Reservered") = Format(dtpDate.Value, "dd/MM/yyyy")
            .Fields("Time_Reservered") = dtpTime.Value
            .Fields("Table_ID") = txtText(5).Text
            .Fields("Section_ID") = txtSection_ID.Text
            .Fields("Amount") = txtText(6).Text
            .Fields("Description") = txtText(7).Text
            .Fields("Cashier_ID") = UserID
            .Fields("IsUsed") = 0
            .Update
        End If
    End With
    Save_Data = True
    Call Lock_text
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "- Save_Data"
    Save_Data = False
End Function

Public Function Check_Null() As Boolean
On Error GoTo Handle
Dim isOK As Boolean
    For i = 0 To txtText.count - 1
    If txtText(i).Text = "" Then
        Select Case i
        Case 0, 1, 3 To 6
            MsgBox "Kh«ng ®­îc ®Ó trèng !"
            txtText(i).BackColor = vbMagenta
            isOK = False
        Exit Function
        Case Else
            txtText(i).Text = " "
        End Select
    Else
        isOK = True
    End If
    Next
Check_Null = isOK
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Check_Null"
End Function

Public Sub Load_Details(ByVal rs As ADODB.Recordset)
On Error GoTo Handle
    Set rsDetails = Nothing
    With rsDetails
        If .State = 0 Then
            .Fields.Append "Reserve_Code", adVarWChar, 20
            .Fields.Append "TableNo", adVarWChar, 20
            .Fields.Append "PluNo", adWChar, 12
            .Fields.Append "PluName", adVarWChar, 100
            .Fields.Append "Qty", adDouble
            .Fields.Append "Price", adDouble
            .Fields.Append "Amt", adDouble
            .Fields.Append "Description", adVarWChar, 50
            .Open
        End If
        Do While Not rs.EOF
            .addNew
            .Fields("Reserve_Code") = rs.Fields("Reservered_Code")
            .Fields("TableNo") = rs.Fields("Table_ID")
            .Fields("PluNo") = rs.Fields("ItemNum")
            .Fields("PluName") = rs.Fields("ItemName")
            .Fields("Qty") = rs.Fields("Qty")
            .Fields("Price") = rs.Fields("Price")
            .Fields("Amt") = rs.Fields("Amt")
            .Fields("Description") = rs.Fields("Description")
            .Update
        rs.MoveNext
        Loop
    End With
    Set DataGrid1.DataSource = rsDetails
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Load_Details"
End Sub

Public Function Get_Recorset_By_Key(ByVal KEY As String) As Recordset
On Error GoTo Handle
Dim rs As New ADODB.Recordset
Dim Str_Details As String
Str_Details = "SELECT Table_Reservered.Reservered_Code, " & _
                  " Table_Reserved_Details.Table_ID, Table_Reserved_Details.ItemNum," & _
                  " Table_Reserved_Details.ItemName, Table_Reserved_Details.Qty," & _
                  " Table_Reserved_Details.Price, [Qty]*[Price] AS Amt," & _
                  " Table_Reserved_Details.Description" & _
                  " FROM Table_Reservered INNER JOIN Table_Reserved_Details ON" & _
                  " Table_Reservered.Reservered_Code = Table_Reserved_Details.Reservered_Code" & _
                  " Where Table_Reservered.Reservered_Code='" & KEY & "'" & _
                  " Order by Table_Reserved_Details.ItemNum"
Set rs = OpenCriticalTable(Str_Details, cnData)

Set Get_Recorset_By_Key = rs
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & " -Get_Recorset_By_Key"
End Function

Public Sub Init_Header_Grids()
On Error GoTo Handle
    With DataGrid1
        .Columns(0).Caption = " " 'DescArr(3)
        .Columns(0).Width = 0
        .Columns(1).Caption = " " 'DescArr(4)
        .Columns(1).Width = 0
        .Columns(2).Caption = "M· hµng" 'DescArr(6)
        .Columns(2).Width = 1200
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Caption = "Tªn hµng" 'DescArr(5)
        .Columns(3).Width = 2500
        .Columns(3).Alignment = dbgLeft
        .Columns(4).Caption = "Sè l­îng"
        .Columns(4).Width = 1000
        .Columns(4).Alignment = dbgCenter
        .Columns(5).Caption = "§¬n gi¸"
        .Columns(5).Width = 1400
        .Columns(5).Alignment = dbgCenter
        .Columns(6).Caption = "Thµnh tiÒn"
        .Columns(6).Width = 1500
        .Columns(6).Alignment = dbgRight
        .Columns(7).Caption = "Ghi chó"
        .Columns(7).Width = 1500
        .Columns(7).Alignment = dbgLeft
        
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " -Init_Header_Grids"
End Sub

Public Sub save_details(rs As ADODB.Recordset)
On Error GoTo Handle
    With rsReserve_Detail
        .Find "Reservered_Code='" & txtText(0).Text & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            cnData.Execute "Delete  from Table_Reserved_Details where Reservered_Code='" & txtText(0).Text & "'"
        End If
        If rs.State = 0 Then Exit Sub
        If rs.State <> 0 And rs.RecordCount > 0 Then rs.MoveFirst
        Do While Not rs.EOF
            .addNew
            .Fields("Reservered_Code") = txtText(0).Text
            .Fields("Table_ID") = rs.Fields("TableNo")
            .Fields("ItemNum") = rs.Fields("PluNo")
            .Fields("ItemName") = rs.Fields("PluName")
            .Fields("Qty") = rs.Fields("Qty")
            .Fields("Price") = rs.Fields("Price")
            .Fields("Description") = ""
            .Fields("KP") = rs.Fields("F3")
            .Update
        rs.MoveNext
        Loop
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " -save_details"
End Sub


Public Sub Save_Incom(sophieu As String)
On Error GoTo Handle
Dim rsPhieuthu As New ADODB.Recordset
If rsPhieuthu.State = 0 Then
'If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
   Set rsPhieuthu = OpenCriticalTable("select * from Income", cnData)
End If
    With rsPhieuthu
        .Find "ID='" & sophieu & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
              .addNew
              .Fields("ID") = sophieu
              .Fields("Store_ID") = Store_ID
              .Fields("Cashier_ID") = UserID
              .Fields("DateTime") = DateDefault
              .Fields("Receipt_ID") = "§C"
              .Fields("Customer_ID") = "101"
              .Fields("Reciever_Name") = txtText(1).Text
              .Fields("Division") = txtText(3).Text
              .Fields("Payment_Method") = "TiÒn mÆt"
              .Fields("Amount") = CDbl("0" & txtText(6).Text)
              .Fields("Description") = "Thu tiÒn ®Æt cäc " & txtText(1).Text & "  " & txtText(4).Text & " kh¸ch vµo lóc " & dtpTime.Value & " ngµy " & dtpDate.Value
              .Update
            Else
              .Fields("Store_ID") = Store_ID
              .Fields("Cashier_ID") = UserID
              .Fields("DateTime") = DateDefault
              .Fields("Receipt_ID") = "§C"
              .Fields("Customer_ID") = "101"
              .Fields("Reciever_Name") = txtText(1).Text
              .Fields("Division") = txtText(3).Text
              .Fields("Payment_Method") = "TiÒn mÆt"
              .Fields("Amount") = CDbl("0" & txtText(6).Text)
              .Fields("Description") = "Thu tiÒn ®Æt cäc " & txtText(1).Text & "  " & txtText(4).Text & " kh¸ch vµo lóc " & dtpTime.Value & " ngµy " & dtpDate.Value
            End If
        End With

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub

Public Sub Save_Customer()
On Error GoTo Handle
    With frmAddCustomer
        .txtCustName.Text = txtText(1).Text
        .txtCustAdd.Text = txtText(2).Text
        .txtCustPhone.Text = txtText(3).Text
        .txtMaxAcc.Text = 0
        .Show vbModal
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Public Function Cancel_Reserve(sophieu As String) As Boolean
On Error GoTo Handle
    With rsTable_Reserered
        .Find " Reservered_Code='" & sophieu & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Delete adAffectCurrent
        End If
    End With
    Cancel_Reserve = True
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name
Cancel_Reserve = False
End Function

Public Sub Payment(sophieu As String)
On Error GoTo Handle
Dim rsPhieuChi As New ADODB.Recordset
Dim Payout_ID As String
Payout_ID = GetMaxSophieu
If rsPhieuChi.State = 0 Then
   Set rsPhieuChi = OpenCriticalTable("select * from Payouts", cnData)
End If
    With rsPhieuChi
            .Find "ID='" & Payout_ID & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("ID") = Payout_ID
                .Fields("Store_ID") = Store_ID
                .Fields("Cashier_ID") = UserID
                .Fields("DateTime") = DateDefault
                .Fields("Expense_ID") = "T§C"
                .Fields("Vendor_Number") = "0000"
                .Fields("Recieve_Name") = txtText(1).Text
                .Fields("Division") = txtText(3).Text
                .Fields("Payment_Method") = "TiÒn mÆt"
                .Fields("Amount") = CDbl("0" & txtText(6).Text)
                .Fields("Description") = "Tr¶ tiÒn ®Æt cäc víi sè phiÕu :" & sophieu & " lóc " & dtpTime.Value & " ngµy " & dtpDate.Value
                .Update
            End If
        End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub
Public Function GetMaxSophieu() As String
On Error GoTo Handle
    Dim rsmax As New ADODB.Recordset
    Dim date_Payout As String
    date_Payout = gfCONVERT_DATE_TO_STRING(dtpDate.Value)
    
    Set rsmax = OpenCriticalTable("select max(ID) as MaxID from Payouts where Substring(DateTime,5,2)='" & Format(Month(dtpDate.Value), "00") & "'", cnData)
    If Not rsmax.EOF Then
    If "" & rsmax.Fields("maxiD") = "" Then
        GetMaxSophieu = "PC/" & Mid(date_Payout, 5, 2) & Mid(date_Payout, 3, 2) & "0001"
    Else
        GetMaxSophieu = Left(rsmax.Fields("MaxID"), Len(rsmax.Fields("MaxID")) - 4) & Right("0000" & (CDbl(Right(rsmax.Fields("MaxID"), 4)) + 1), 4)
    End If
    Else
        GetMaxSophieu = "PC/" & Mid(date_Payout, 5, 2) & Mid(date_Payout, 3, 2) & "0001"
    End If
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   GetMaxSophieu"
End Function

Private Sub txtText_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Handle
    Select Case Index
        Case 4, 6
        If KeyAscii < 32 And KeyAscii <> 13 Then Exit Sub
            Select Case KeyAscii
                Case 48 To 57, 45, 46
                Case 13
                    txtText(Index).Text = Format(txtText(Index).Text, "#,##0")
                Case Else:   KeyAscii = 0
            End Select
        Case Else
    End Select
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtQty_KeyPress"
End Sub
