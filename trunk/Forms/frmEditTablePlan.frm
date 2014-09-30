VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEditTablePlan 
   BackColor       =   &H00C0C0C0&
   Caption         =   "ChØnh söa s¬ ®å bµn"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   795
   ClientWidth     =   15705
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
   Icon            =   "frmEditTablePlan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   15705
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pic1 
      BackColor       =   &H00C0C0C0&
      Height          =   8535
      Left            =   13320
      ScaleHeight     =   8475
      ScaleWidth      =   1875
      TabIndex        =   20
      Top             =   1080
      Width           =   1935
      Begin prjTouchScreen.MyButton cmdRange 
         Height          =   735
         Left            =   120
         TabIndex        =   35
         Top             =   3360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1296
         BTYPE           =   6
         TX              =   "§Æt l¹i  tù ®éng"
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
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmEditTablePlan.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.CheckBox chkRange 
         BackColor       =   &H00C0C0C0&
         Caption         =   "S¾p xÕp l¹i"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   2880
         Width           =   1695
      End
      Begin VB.ComboBox txtFont 
         Height          =   390
         ItemData        =   "frmEditTablePlan.frx":0028
         Left            =   120
         List            =   "frmEditTablePlan.frx":002A
         TabIndex        =   30
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtHeight 
         Height          =   390
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtWidth 
         Height          =   390
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   1695
      End
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   780
         Left            =   0
         TabIndex        =   21
         Top             =   7680
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   1376
         BTYPE           =   5
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
         MICON           =   "frmEditTablePlan.frx":002C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Font size"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   1755
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "ChiÒu cao"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "ChiÒu réng"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   120
         Width           =   1515
      End
   End
   Begin VB.PictureBox fraTable 
      BackColor       =   &H80000007&
      Height          =   8535
      Left            =   0
      ScaleHeight     =   8475
      ScaleWidth      =   13185
      TabIndex        =   17
      Top             =   1080
      Width           =   13245
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
         Left            =   360
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFC0&
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Height          =   1035
         Index           =   0
         Left            =   240
         Top             =   600
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Shape Shape2 
         Height          =   1785
         Index           =   0
         Left            =   0
         Top             =   120
         Visible         =   0   'False
         Width           =   2805
      End
   End
   Begin TabDlg.SSTab TabSec 
      Height          =   855
      Left            =   0
      TabIndex        =   0
      Top             =   9600
      Width           =   15525
      _ExtentX        =   27384
      _ExtentY        =   1508
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "   "
      TabPicture(0)   =   "frmEditTablePlan.frx":0048
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdSection(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin prjTouchScreen.MyButton cmdSection 
         Height          =   885
         Index           =   0
         Left            =   30
         TabIndex        =   1
         Top             =   -75
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1561
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
         MICON           =   "frmEditTablePlan.frx":0064
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
   End
   Begin TabDlg.SSTab TabTop 
      Height          =   1095
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   20325
      _ExtentX        =   35851
      _ExtentY        =   1931
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   255
      ForeColor       =   4210752
      OLEDropMode     =   1
      TabCaption(0)   =   "   "
      TabPicture(0)   =   "frmEditTablePlan.frx":0080
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSeat"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblSection"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdDeleteLocation"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdDeleteTable"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdAddLocation"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdAddTable"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdHelp"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdDone"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "SSTab3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cboSeat"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Picture1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtxpos"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtypos"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      Begin VB.TextBox txtypos 
         Height          =   495
         Left            =   19200
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox txtxpos 
         Height          =   495
         Left            =   17280
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Height          =   15
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   7815
         TabIndex        =   16
         Top             =   1080
         Width           =   7815
      End
      Begin VB.ComboBox cboSeat 
         Height          =   390
         Left            =   15180
         TabIndex        =   3
         Text            =   "N/A"
         Top             =   510
         Width           =   1215
      End
      Begin TabDlg.SSTab SSTab3 
         Height          =   1095
         Left            =   -270
         TabIndex        =   4
         Top             =   1230
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   1931
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   520
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "frmEditTablePlan.frx":009C
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "MyButton2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "MyButton1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         Begin prjTouchScreen.MyButton MyButton1 
            Height          =   825
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   1455
            BTYPE           =   3
            TX              =   "Exit toLogin"
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
            BCOL            =   12632319
            BCOLO           =   12648447
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmEditTablePlan.frx":00B8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton MyButton2 
            Height          =   825
            Left            =   2190
            TabIndex        =   6
            Top             =   120
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   1455
            BTYPE           =   3
            TX              =   "Edit layout"
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
            BCOL            =   12632319
            BCOLO           =   12648447
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmEditTablePlan.frx":00D4
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
      Begin prjTouchScreen.MyButton cmdDone 
         Height          =   1035
         Left            =   9555
         TabIndex        =   7
         Tag             =   "L9"
         Top             =   15
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1826
         BTYPE           =   6
         TX              =   "&Hoµn tÊt"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmEditTablePlan.frx":00F0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdHelp 
         Height          =   1035
         Left            =   7605
         TabIndex        =   8
         Tag             =   "L8"
         Top             =   15
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1826
         BTYPE           =   6
         TX              =   "Gióp ®ì"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmEditTablePlan.frx":010C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAddTable 
         Height          =   1035
         Left            =   105
         TabIndex        =   9
         Tag             =   "L2"
         Top             =   15
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1826
         BTYPE           =   6
         TX              =   "&Thªm bµn míi"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmEditTablePlan.frx":0128
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAddLocation 
         Height          =   1035
         Left            =   3855
         TabIndex        =   10
         Tag             =   "L4"
         Top             =   15
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1826
         BTYPE           =   6
         TX              =   "Thªm khu vùc"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmEditTablePlan.frx":0144
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdDeleteTable 
         Height          =   1035
         Left            =   1980
         TabIndex        =   11
         Tag             =   "L3"
         Top             =   15
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1826
         BTYPE           =   6
         TX              =   "&Xãa bµn"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmEditTablePlan.frx":0160
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdDeleteLocation 
         Height          =   1035
         Left            =   5730
         TabIndex        =   12
         Tag             =   "L5"
         Top             =   15
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   1826
         BTYPE           =   6
         TX              =   "Xãa khu vùc"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16744576
         BCOLO           =   16744576
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmEditTablePlan.frx":017C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Khu vùc"
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   11520
         TabIndex        =   27
         Top             =   120
         Width           =   1995
      End
      Begin VB.Label lblSection 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
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
         Height          =   315
         Left            =   11520
         TabIndex        =   26
         Top             =   480
         Width           =   2265
      End
      Begin VB.Label Label4 
         Caption         =   "YPos"
         Height          =   375
         Left            =   18480
         TabIndex        =   25
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "XPos"
         Height          =   375
         Left            =   16560
         TabIndex        =   23
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Sè ghÕ tèi ®a"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   14910
         TabIndex        =   15
         Tag             =   "L7"
         Top             =   150
         Width           =   1845
      End
      Begin VB.Label lblSeat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "1"
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   13875
         TabIndex        =   14
         Top             =   540
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Bµn sè"
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   13890
         TabIndex        =   13
         Tag             =   "L6"
         Top             =   120
         Width           =   1275
      End
   End
   Begin prjTouchScreen.MyButton cmdPrint_Receipt 
      Height          =   780
      Left            =   13080
      TabIndex        =   19
      Top             =   240
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1376
      BTYPE           =   5
      TX              =   "In Hãa §¬n"
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
      BCOL            =   -2147483647
      BCOLO           =   -2147483647
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmEditTablePlan.frx":0198
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Menu mnuAdd 
      Caption         =   "&Thªm míi"
      Begin VB.Menu mnuLocation 
         Caption         =   "Thªm khu vùc"
      End
      Begin VB.Menu mnuTable 
         Caption         =   "Thªm bµn"
      End
      Begin VB.Menu mnuDeleteLocation 
         Caption         =   "Xãa khu vùc"
      End
      Begin VB.Menu mnuDeleteTable 
         Caption         =   "Xãa bµn"
      End
   End
   Begin VB.Menu mnuAlign 
      Caption         =   "&Canh lÒ"
      Begin VB.Menu mnuAlignLeft 
         Caption         =   "Canh t&r¸i"
      End
      Begin VB.Menu mnuAlignTop 
         Caption         =   "Canh &trªn"
      End
      Begin VB.Menu mnurefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Trî gióp"
      Begin VB.Menu mnuHelpuser 
         Caption         =   "Gióp ®ì"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Tho¸t"
      End
   End
End
Attribute VB_Name = "frmEditTablePlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Drag As Boolean
Dim rsSection As New ADODB.Recordset
Dim CountTable As Integer
Dim CountSection As Integer
Dim rsTable As New ADODB.Recordset
Dim iLoad As Boolean
Dim iLoadSection As Boolean
Dim rsInvoice_On_Holds As New ADODB.Recordset
Dim indexTable As Integer
Dim rsAlign As New ADODB.Recordset
Dim XPos, YPos As Integer
Dim tableCaption  As String


Private Sub cboSeat_Change()
On Error GoTo Handle

Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   cboSeat_Change "

End Sub

Private Sub cboSeat_Click()
On Error GoTo Handle
If Sec_ID <> "" And iLoad = True Then
   With rsTable
        .Fields("NumSeats") = CDbl(cboSeat.Text)
        .Update
        .Requery
    End With
End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   cboSeat_Change "

End Sub

Private Sub cmdAddLocation_Click()
'    With frmKeyboard
'        .FormCallkeyboard = "Add_Section"
'        .lblTitle.Caption = "Enter_Section"
'        .txtInput.PasswordChar = ""
'        .Show vbModal
'    End With
'
    frmAdd_Location.Show vbModal
    iLoad = False
End Sub

Private Sub cmdAddTable_Click()
On Error GoTo Handle
    With frmRangeTable
        .Get_Location = Sec_ID
        .Get_Width = fraTable.Width
        .Get_Height = fraTable.Height
        .Show vbModal
        
    End With
    Call LoadTable(Sec_ID)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  cmdAddTable_Click "
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDeleteLocation_Click()
    On Error GoTo Handle
    Dim ans As Integer
    Dim rsInvoice_Totals As New ADODB.Recordset
    Set rsInvoice_Totals = Open_Table(cnData, "Invoice_Totals")
    With rsInvoice_On_Holds
        .Find "Section_ID='" & Sec_ID & "'", , adSearchForward, adBookmarkFirst
        If .EOF Then
            ans = MsgBox("B¹n cã ch¾c ch¾n muèn xãa khu vùc nµy kh«ng?", vbYesNo)
            If ans = vbYes Then
                With rsInvoice_Totals
                    .Find "Station_ID='" & Sec_ID & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        If MsgBox("Khu vùc nµy ®· cã b¸n hµng, nÕu b¹n chän YES d÷ liÖu b¸n hµng sÏ tù ®äng xãa bá", vbYesNo) = vbYes Then
                            GoTo 1
                        Else
                            Exit Sub
                        End If
                    Else
1:
                        cnData.Execute "Delete from Table_Diagram_Sections where Location_ID='" & Sec_ID & "'"
                        cnData.Execute "Delete  from Table_Diagram where Section_ID='" & Sec_ID & "'"
                        cnData.Execute "Delete   from Invoice_Totals where Station_ID ='" & Sec_ID & "'"
                        Call Load_Section
                    End If
                End With
            End If
        Else
            MsgBox "Khu vùc nµy ®ang sö dông, b¹n kh«ng thÓ xãa khu vùc nµy !"
        End If
    End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdDeleteLocation_Click "
End Sub

Private Sub cmdDeleteTable_Click()
    On Error GoTo Handle
        If cnData.State = adStateClosed Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
        With rsInvoice_On_Holds
            If .State <> 0 Then
                If .RecordCount > 0 Then .MoveFirst
                .Find "OnHoldID='" & Left(lblTable(indexTable).Caption, InStr(lblTable(indexTable).Caption, Chr(13))) & "'", , adSearchForward, adBookmarkFirst
                If .EOF Then
                    If MsgBox("B¹n cã ch¾n ch¾n muèn xãa bµn nµy ?", vbYesNo) = vbYes Then
                    
                        If Sec_ID <> "" Then
                        If rsAlign.State > 0 And rsAlign.RecordCount > 0 Then rsAlign.MoveFirst
                            With rsAlign
                                Do While Not .EOF
                                Dim SQL As String
                                SQL = "delete from Table_Diagram where Section_ID='" & Sec_ID & "' and Table_Number='" & .Fields("TableName") & "'"
                                cnData.Execute SQL
                                .MoveNext
                                Loop
                            End With
                        End If
                        Call LoadTable(Sec_ID)
                    End If
                Else
                    MsgBox "Bµn nµy ®ang sö dông kh«ng thÓ xãa ®­îc !!!"
                    With rsAlign
                    .Find "TableName ='" & .Fields("TableName") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Delete adAffectCurrent
                            .Update
                        End If
                    End With
                End If
            End If
        End With
        
        indexTable = 0
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdDeleteTable_Click "
End Sub

Private Sub cmdDone_Click()
    Set rsAlign = Nothing
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Showhelp "Setup Table Plan"
End Sub

Private Sub cmdRange_Click()
    On Error GoTo Handle
        If chkRange.Value = 1 Then
            If txtWidth.Text = "" Then
                MsgBox "Vui lßng chän mét bµn ®Ó lÊy kÝch th­íc chuÈn tr­íc khi ®Æt l¹i", vbInformation
                Exit Sub
            End If
            Call Auto_Range(fraTable.Width, Sec_ID)
            Set rsAlign = Nothing
        End If
        Call LoadTable(CStr(Sec_ID))
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  - cmdRange_Click"
End Sub
Public Sub Auto_Range(Width_Layout As Integer, Location_ID As String)
    On Error GoTo Handle
        Dim rows, cols As Integer
        Dim i, j As Integer
        Dim Tablewidth, TableHeight As Integer
        Dim rsLocate As New ADODB.Recordset
        
        Tablewidth = CInt("0" & txtWidth.Text)
        TableHeight = CInt("0" & txtHeight.Text)
        
        If cnData.State = adStateClosed Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
        Set rsLocate = OpenCriticalTable("Select * from Table_Diagram where Section_ID='" & Location_ID & "' order by Table_Number", cnData)
         
         cols = Int((Width_Layout - 500) / Tablewidth)
        rows = Int(rsLocate.RecordCount / cols) + 1
        rsLocate.MoveFirst
        i = 0
        For i = 1 To rows
            For j = 1 To cols
                With rsLocate
                    If .EOF Then Exit Sub
                    If i = 1 Then
                            !YPos = i * TableHeight - TableHeight + 100
                        Else
                            !YPos = i * TableHeight - TableHeight + i * 100
                        End If
                        If j = 1 Then
                            !XPos = j * Tablewidth - Tablewidth
                        Else
                            !XPos = j * Tablewidth - Tablewidth + j * 50
                        End If
                    .Update
                    If Not .EOF Then .MoveNext
                End With
            Next j
        Next i
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Auto_Range"
End Sub


Private Sub cmdSection_Click(Index As Integer)
    On Error GoTo Handle
    Dim ctrl As Control
        Sec_ID = Format(cmdSection(Index).Tag, "00")
        cmdSection(Index).BackColor = vbGreen
        Call LoadTable(CStr(Sec_ID))
        lblSection.Caption = cmdSection(Index).Caption
        iLoad = True
        For Each ctrl In Me
        If ctrl.name = "cmdSection" Then
            ctrl.ForeColor = vbBlue
        End If
    Next ctrl
    cmdSection(Index).ForeColor = vbRed
    Set rsInvoice_On_Holds = OpenCriticalTable("select * from Invoice_OnHold where Section_ID='" & Sec_ID & "'", cnData)
    Exit Sub
    
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdSection_Click "
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    Dim DescArr() As String
    Dim ctrl As Control
    If iLoad = True Then Exit Sub
    iLoad = True
    DescArr = LoadLanguage(LngFile, "#03:014:")
    If cmdDone.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
   
    Call Load_Section
    If Sec_ID <> "" Then
        Call LoadTable(Sec_ID)
    Else
        Call LoadTable("01")
    End If
    fraTable.BackColor = bkColor
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   Form_Activate"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then Call mnurefresh_Click
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    iLoad = False
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsInvoice_On_Holds = Open_Table(cnData, "Invoice_OnHold")
    Call Load_Seat_Number
    Call Load_font
    
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub

Private Sub Form_Resize()
    pic1.Left = Me.Width - pic1.Width
    TabTop.Width = Me.Width
    fraTable.Width = Me.Width - pic1.Width
    TabSec.Width = Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Handle
    Set cnData = Nothing
    Set rsSection = Nothing
    Set rsTable = Nothing
    CountTable = 0
    CountSection = 0
    iLoadSection = False
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   Form_Unload"
End Sub

Private Sub Frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuAlign, 0
    End If
End Sub

Private Sub lblTable_Click(Index As Integer)
    On Error GoTo Handle
    Dim i As Integer
        tableCaption = Left(lblTable(Index).Caption, InStr(Replace(lblTable(Index).Caption, Chr(13) & Chr(13), Chr(13)), Chr(13)))
        tableCaption = Replace(tableCaption, Chr(13), "")
        lblSeat.Caption = tableCaption
        With rsTable
            .Find "Table_Number='" & tableCaption & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                cboSeat.Text = .Fields("NumSeats")
                txtWidth.Text = .Fields("Width")
                txtHeight.Text = .Fields("Height")
                txtFont.Text = .Fields("Cost_Center_Index")
                txtxpos.Text = .Fields("XPOS")
                txtypos.Text = .Fields("YPOS")
            End If
        End With
        indexTable = Index
        With rsAlign
            If .State = 0 Then
                    .Fields.Append "TableName", adVarWChar, 30
                    .Open
            End If
            .Find "TableName ='" & tableCaption & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    .Delete adAffectCurrent
                    lblTable(Index).BackColor = vbBlue
                Else
                    .addNew
                    .Fields("TableName") = tableCaption
                    .Update
'                    .Requery
                    lblTable(Index).BackColor = vbRed
                End If
            With rsTable
                .Find "Table_Number='" & tableCaption & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    XPos = .Fields("XPos")
                    YPos = .Fields("YPos")
                End If
            End With
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " lblTable_Click  "
End Sub

Private Sub lblTable_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Handle
    Load Shape2(1)
    Shape2(1).Visible = True
    Shape2(1).Left = Shape1(Index).Left - 70
    Shape2(1).top = Shape1(Index).top - 70
    Shape2(1).Width = Shape1(Index).Width
    Shape2(1).Height = Shape1(Index).Height
    Drag = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   lblTable_MouseDown"
End Sub

'
Private Sub lblTable_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Handle
   If Drag = True Then
        lblTable(Index).Move lblTable(Index).Left + X, lblTable(Index).top + Y
    Shape2(1).Left = lblTable(Index).Left
    Shape2(1).top = lblTable(Index).top
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   lblTable_MouseMove"
End Sub

Private Sub lblTable_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo Handle
    Drag = False
    Shape1(Index).Left = lblTable(Index).Left - 40
    Shape1(Index).top = lblTable(Index).top - 45
    Dim rsupdate As New ADODB.Recordset
    Dim tableCaption  As String
    tableCaption = Left(lblTable(Index).Caption, InStr(Replace(lblTable(Index).Caption, Chr(13) & Chr(13), Chr(13)), Chr(13)))
    tableCaption = Replace(tableCaption, Chr(13), "")
    If Sec_ID <> "" Then
    Set rsupdate = OpenCriticalTable("Select * from Table_Diagram where Section_ID='" & Sec_ID & "' and Table_Number='" & tableCaption & "'", cnData)
        If rsupdate.RecordCount > 0 Then
        DoEvents
            With rsupdate
                .Fields("XPOS") = lblTable(Index).Left
                .Fields("YPOS") = lblTable(Index).top
                .Fields("Height") = CDbl("0" & txtHeight.Text)
                .Fields("Width") = CDbl("0" & txtWidth.Text)
                 DoEvents
                .Update
                .Requery
            End With
        End If
    Else
        MsgBox "B¹n ph¶i chän khu vùc tr­íc khi chØnh söa s¬ ®å bµn"
    End If
    Shape2(1).Visible = False
    Unload Shape2(1)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  lblTable_MouseUp"
End Sub

Public Sub Load_Section()
    On Error GoTo Handle
    Dim ctrl As Control
        Dim i, a, b As Integer
        i = 1
        a = 0
        If cnData.State = adStateClosed Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
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
                        .Left = cmdSection(i - 1).Left + 80
                    Else
                        .Left = cmdSection(i - 1).Left + cmdSection(i - 1).Width + 80
                    End If
                    .top = cmdSection(i - 1).top
                    .Visible = True
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
Dim rsInvoice_Total As New ADODB.Recordset
Dim rsVacantColor As New ADODB.Recordset
Dim i, j As Integer
i = 1: j = 1
    Dim str As String
    Dim ctrl As Control
    If CountTable > 0 Then
        For j = 1 To CountTable
            Unload lblTable(j)
            Unload Shape1(j)
        Next
    End If
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    'Lay Bang mau
    Dim TypeColor, SeatedColor, BlankTable As String
    TypeColor = "RESERVED"
    SeatedColor = "SEATED"
    BlankTable = "VACANT"
    Set rscolor = OpenCriticalTable("Select ReserveValue from ColorTablePlan where ReserveType='" & TypeColor & "'", cnData)
    Set rsSeatedColor = OpenCriticalTable("Select ReserveValue from ColorTablePlan where ReserveType='" & SeatedColor & "'", cnData)
    Set rsVacantColor = OpenCriticalTable("Select ReserveValue from ColorTablePlan where ReserveType='" & BlankTable & "'", cnData)

    str = "select * from Table_Diagram where Section_ID='" & Section_ID & "'"
    Set rsTable = OpenCriticalTable(str, cnData)
    CountTable = rsTable.RecordCount
    Dim strTableTotal As String
    Do While Not rsTable.EOF
        Load lblTable(i)
        With lblTable(i)
            .Left = rsTable.Fields("XPOS")
            .top = rsTable.Fields("YPOS")
            .Height = rsTable.Fields("Height")
            .Width = rsTable.Fields("width")
            strTableTotal = "SELECT Invoice_OnHold.Invoice_Number, Invoice_Totals.Store_ID," & _
            "Invoice_OnHold.OnHoldID, Invoice_Totals.Grand_Total, Invoice_Totals.Total_Price, " & _
            "Invoice_Totals.Orig_OnHoldID, Invoice_OnHold.Section_ID FROM Invoice_OnHold" & _
            " INNER JOIN Invoice_Totals ON Invoice_OnHold.Invoice_Number = Invoice_Totals.Invoice_Number " & _
            " where Invoice_OnHold.OnHoldID = '" & rsTable.Fields("Table_number") & Chr(13) & "' and Invoice_OnHold.Section_ID='" & Section_ID & "'"
            Set rsInvoice_Total = OpenCriticalTable(strTableTotal, cnData)
            If rsInvoice_Total.RecordCount > 0 Then
                If CDbl("0" & rsInvoice_Total.Fields("Grand_Total")) > 0 Then
                    .Caption = rsTable.Fields("Table_Number") & Chr(13) & Format(rsInvoice_Total.Fields("Grand_Total"), formatNum)
                    .BackStyle = 1
                    .BackColor = rscolor.Fields("ReserveValue")
                    .FontSize = rsTable.Fields("Cost_Center_Index")
                Else
                    .Caption = rsTable.Fields("Table_Number") & Chr(13) & Format(rsInvoice_Total.Fields("Grand_Total"), formatNum)
                    .BackStyle = 1
                    .BackColor = rsSeatedColor.Fields("ReserveValue")
                    .FontSize = rsTable.Fields("Cost_Center_Index")
                End If
            Else
                .Caption = rsTable.Fields("Table_Number") & Chr(13)
                .FontSize = rsTable.Fields("Cost_Center_Index")
                .BackStyle = 1
                .BackColor = rsVacantColor.Fields("ReserveValue")
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
            .BorderColor = ShapeColor
            .Visible = True
        End With
    rsTable.MoveNext
    i = i + 1
    Loop

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  LoadTable"
End Sub

Public Sub Load_Seat_Number()
On Error GoTo Handle
    Dim i As Integer
    cboSeat.Clear
    For i = 1 To 30
        With cboSeat
            .AddItem i
        End With
    Next
    cboSeat.ListIndex = 0
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Load_Seat_Number  "
End Sub

Private Sub mnuAlignLeft_Click()
    On Error GoTo Handle
        If rsAlign.State <> 0 Then
            If rsAlign.RecordCount > 0 Then rsAlign.MoveFirst
        Else
            Exit Sub
        End If
        With rsAlign
            Do While Not rsAlign.EOF
                With rsTable
                    .Find "Table_Number='" & rsAlign.Fields("TableName") & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        .Fields("xPos") = XPos
                        .Update
                        .Requery
                    End If
                End With
            rsAlign.MoveNext
            Loop
        End With
        Call cmdSection_Click(Int(Sec_ID))
        Set rsAlign = Nothing
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " mnuAlignLeft_Click"
End Sub

Private Sub mnuAlignTop_Click()
    On Error GoTo Handle
    If rsAlign.State <> 0 Then
        If rsAlign.RecordCount > 0 Then rsAlign.MoveFirst
    Else
        Exit Sub
    End If
        With rsAlign
            Do While Not rsAlign.EOF
                With rsTable
                    .Find "Table_Number='" & rsAlign.Fields("TableName") & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        .Fields("YPos") = YPos
                        .Update
                        .Requery
                    End If
                End With
            rsAlign.MoveNext
            Loop
        End With
        Call cmdSection_Click(Int(Sec_ID))
        Set rsAlign = Nothing
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " mnuAlignLeft_Click"

End Sub

Private Sub mnuDeleteLocation_Click()
    cmdDeleteLocation_Click
End Sub

Private Sub mnuDeleteTable_Click()
    cmdDeleteTable_Click
End Sub

Private Sub mnuexit_Click()
    cmdDone_Click
End Sub

Private Sub mnuLocation_Click()
    cmdAddLocation_Click
End Sub

Private Sub mnurefresh_Click()
    Call txtWidth_KeyPress(13)
End Sub

Private Sub mnuTable_Click()
On Error GoTo Handle
    With frmRangeTable
        .Get_Location = Sec_ID
        .Show vbModal
        
    End With
    Call LoadTable(Sec_ID)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  mnuTable_Click "
End Sub

Private Sub SSTab1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
        PopupMenu mnuAlign, 0
    End If
End Sub

Private Sub txtFont_change()
On Error GoTo Handle
    Set rsTable = OpenCriticalTable("select * from Table_Diagram where Section_ID='" & Sec_ID & "'", cnData)
        With rsTable
            .Find "Table_Number='" & lblSeat.Caption & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    .Fields("Width") = CDbl("0" & txtWidth.Text)
                    .Fields("Height") = CDbl("0" & txtHeight.Text)
                    .Fields("Cost_Center_Index") = CDbl("0" & txtFont.Text)
                    .Update
    '                .Requery
                End If
        End With
'  cmdSection_Click (Int(Sec_ID))
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtFont_KeyPress"
End Sub


Private Sub txtFont_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Handle
    If KeyCode = vbKeyDown Then
    Set rsTable = OpenCriticalTable("select * from Table_Diagram where Section_ID='" & Sec_ID & "'", cnData)
        With rsTable
            .Find "Table_Number='" & tableCaption & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    .Fields("Width") = CDbl("0" & txtWidth.Text)
                    .Fields("Height") = CDbl("0" & txtHeight.Text)
                    .Fields("Cost_Center_Index") = CDbl("0" & txtFont.Text)
                    .Update
    '                .Requery
                End If
        End With
  cmdSection_Click (Int(Sec_ID))
  End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtFont_KeyPress"
End Sub

Private Sub txtHeight_DblClick()
On Error GoTo Handle
Dim i As Integer
    With frmKeyboard
        .txtInput.PasswordChar = ""
        .FormCallkeyboard = "Other"
        .Show vbModal
        For i = 10 To 38
            .cmdText(i).Enabled = False
        Next
        txtHeight.Text = .Let_Text_Input
    End With
    Call txtHeight_KeyPress(13)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtHeight_DblClick"
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        Set rsTable = OpenCriticalTable("select * from Table_Diagram where Section_ID='" & Sec_ID & "'", cnData)
            If rsAlign.State <> 0 Then
                If rsAlign.RecordCount > 0 Then rsAlign.MoveFirst
            Else
                Exit Sub
            End If
            With rsAlign
                Do While Not rsAlign.EOF
                    With rsTable
                        .Find "Table_Number='" & rsAlign.Fields("TableName") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("Width") = CDbl("0" & txtWidth.Text)
                            .Fields("Height") = CDbl("0" & txtHeight.Text)
                            .Fields("Cost_Center_Index") = CDbl("0" & txtFont.Text)
                            .Update
                            .Requery
                        End If
                    End With
                rsAlign.MoveNext
                Loop
            End With
            Call cmdSection_Click(Int(Sec_ID))
            Set rsAlign = Nothing
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtHeight_KeyPress"
End Sub

Private Sub txtWidth_DblClick()
On Error GoTo Handle
Dim i As Integer
    With frmKeyboard
        .txtInput.PasswordChar = ""
        .FormCallkeyboard = "Other"
        .Show vbModal
        For i = 10 To 38
            .cmdText(i).Enabled = False
        Next
        txtWidth.Text = .Let_Text_Input
    End With
    Call txtWidth_KeyPress(13)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtWidth_DblClick"
End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
If KeyAscii = 13 Then
    Set rsTable = OpenCriticalTable("select * from Table_Diagram where Section_ID='" & Sec_ID & "'", cnData)
        If rsAlign.State <> 0 Then
            If rsAlign.RecordCount > 0 Then rsAlign.MoveFirst
        Else
            Exit Sub
        End If
        With rsAlign
            Do While Not rsAlign.EOF
                With rsTable
                    .Find "Table_Number='" & rsAlign.Fields("TableName") & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        .Fields("Width") = CDbl("0" & txtWidth.Text)
                        .Fields("Height") = CDbl("0" & txtHeight.Text)
                        .Fields("Cost_Center_Index") = CDbl("0" & txtFont.Text)
                        .Update
                        .Requery
                    End If
                End With
            rsAlign.MoveNext
            Loop
        End With
        Call cmdSection_Click(Int(Sec_ID))
        Set rsAlign = Nothing
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtWidth_KeyPress"
End Sub

Public Sub Load_font()
On Error GoTo Handle:
Dim i As Integer
txtFont.Clear
    With txtFont
        For i = 12 To 24
            .AddItem i
        Next
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Load_font"
End Sub

Private Sub txtxpos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With rsTable
            .Find "Table_Number='" & tableCaption & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                 .Fields("XPOS") = txtxpos.Text
                .Update
            End If
        End With
    End If
    Call cmdSection_Click(Int(Sec_ID))
End Sub



Private Sub txtypos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        With rsTable
            .Find "Table_Number='" & tableCaption & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                 .Fields("YPOS") = txtypos.Text
                .Update
            End If
        End With
    End If
    Call cmdSection_Click(Int(Sec_ID))
End Sub
