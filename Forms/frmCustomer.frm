VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Th«ng tin kh¸ch hµng"
   ClientHeight    =   10530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15270
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10530
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdSearch 
      Height          =   915
      Left            =   6600
      TabIndex        =   36
      Tag             =   "L38"
      Top             =   9330
      Width           =   1425
      _extentx        =   2514
      _extenty        =   1614
      btype           =   14
      tx              =   "Search"
      enab            =   -1
      font            =   "frmCustomer.frx":0000
      coltype         =   2
      focusr          =   -1
      bcol            =   16578804
      bcolo           =   16578804
      fcol            =   16711680
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmCustomer.frx":0028
      picn            =   "frmCustomer.frx":0046
      umcol           =   -1
      soft            =   0
      picpos          =   0
      ngrey           =   0
      fx              =   0
      hand            =   0
      check           =   0
      value           =   0
   End
   Begin VB.ComboBox cboSearch 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   540
      Left            =   2280
      TabIndex        =   35
      Top             =   9510
      Width           =   4245
   End
   Begin VB.PictureBox pic1 
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   8100
      ScaleHeight     =   1050
      ScaleWidth      =   6645
      TabIndex        =   15
      Top             =   240
      Width           =   6705
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Left            =   1560
         TabIndex        =   17
         Top             =   480
         Width           =   3345
      End
      Begin VB.Label lblNumber 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "CustomerID"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   1560
         TabIndex        =   16
         Top             =   30
         Width           =   3345
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flexCustomer 
      Height          =   9135
      Left            =   30
      TabIndex        =   14
      Top             =   60
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   16113
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      GridColor       =   0
      WordWrap        =   -1  'True
      TextStyleFixed  =   3
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
   Begin TabDlg.SSTab tabCustomer 
      Height          =   6255
      Left            =   8100
      TabIndex        =   11
      Top             =   1710
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Th«ng tin kh¸ch hµng"
      TabPicture(0)   =   "frmCustomer.frx":0C9A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblCustomer(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblCustomer(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblCustomer(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblCustomer(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblCustomer(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblCustomer(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblCustomer(7)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblCustomer(8)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblCustomer(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblCustomer(10)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblCustomer(9)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dtpBirth"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCustomer(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtCustomer(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCustomer(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCustomer(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCustomer(7)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtCustomer(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtCustomer(8)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtCustomer(6)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtCustomer(2)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtCustomer(13)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cboPro"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Th«ng tin tµi kho¶n"
      TabPicture(1)   =   "frmCustomer.frx":0CB6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "NhËt ký mua hµng"
      TabPicture(2)   =   "frmCustomer.frx":0CD2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtTotals"
      Tab(2).Control(1)=   "flgCustHistorySale"
      Tab(2).Control(2)=   "Label3"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Thanh to¸n c«ng nî"
      TabPicture(3)   =   "frmCustomer.frx":0CEE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblBill"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblPayby"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblAount"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "flgPaymentHistory"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cboBill"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtAmount"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmdPayment"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "cmdPrint"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "cmdclear"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "ChÝnh s¸ch tÝch lòy ®iÓm"
      TabPicture(4)   =   "frmCustomer.frx":0D0A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdSavepoint"
      Tab(4).Control(1)=   "Frame3"
      Tab(4).Control(2)=   "Frame4"
      Tab(4).ControlCount=   3
      Begin VB.ComboBox cboPro 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2010
         TabIndex        =   1
         Text            =   "Nhãm kh¸ch hµng"
         Top             =   1920
         Width           =   4815
      End
      Begin VB.TextBox txtTotals 
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -70560
         TabIndex        =   83
         Top             =   5640
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   66
         Top             =   960
         Width           =   6255
         Begin VB.TextBox txtCustomer 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   12
            Left            =   2250
            TabIndex        =   70
            Top             =   1560
            Width           =   2025
         End
         Begin VB.TextBox txtCustomer 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   10
            Left            =   360
            TabIndex        =   68
            Top             =   3600
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.TextBox txtCustomer 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   11
            Left            =   360
            TabIndex        =   67
            Top             =   4200
            Visible         =   0   'False
            Width           =   2025
         End
         Begin prjTouchScreen.MyButton cmdclean 
            Height          =   615
            Left            =   4200
            TabIndex        =   69
            Top             =   2160
            Visible         =   0   'False
            Width           =   855
            _extentx        =   1508
            _extenty        =   1085
            btype           =   5
            tx              =   "&Reset"
            enab            =   -1
            font            =   "frmCustomer.frx":0D26
            coltype         =   2
            focusr          =   -1
            bcol            =   14215660
            bcolo           =   14215660
            fcol            =   0
            fcolo           =   0
            mcol            =   12632256
            mptr            =   1
            micon           =   "frmCustomer.frx":0D4E
            umcol           =   -1
            soft            =   0
            picpos          =   0
            ngrey           =   0
            fx              =   0
            hand            =   0
            check           =   0
            value           =   0
         End
         Begin prjTouchScreen.MyButton cmdOpenAcc 
            Height          =   555
            Left            =   3990
            TabIndex        =   71
            Top             =   300
            Width           =   1125
            _extentx        =   1984
            _extenty        =   979
            btype           =   5
            tx              =   "Open"
            enab            =   -1
            font            =   "frmCustomer.frx":0D6C
            coltype         =   2
            focusr          =   -1
            bcol            =   12632256
            bcolo           =   16777152
            fcol            =   16711680
            fcolo           =   255
            mcol            =   12632256
            mptr            =   1
            micon           =   "frmCustomer.frx":0D94
            umcol           =   -1
            soft            =   0
            picpos          =   0
            ngrey           =   0
            fx              =   0
            hand            =   0
            check           =   0
            value           =   0
         End
         Begin prjTouchScreen.MyButton cmCloseAcc 
            Height          =   555
            Left            =   3990
            TabIndex        =   72
            Top             =   900
            Width           =   1125
            _extentx        =   1984
            _extenty        =   979
            btype           =   5
            tx              =   "&Open"
            enab            =   -1
            font            =   "frmCustomer.frx":0DB2
            coltype         =   2
            focusr          =   -1
            bcol            =   12632256
            bcolo           =   16777152
            fcol            =   16711680
            fcolo           =   255
            mcol            =   12632256
            mptr            =   1
            micon           =   "frmCustomer.frx":0DDA
            umcol           =   -1
            soft            =   0
            picpos          =   0
            ngrey           =   0
            fx              =   0
            hand            =   0
            check           =   0
            value           =   0
         End
         Begin MSComCtl2.DTPicker dtpOpenAcc 
            Height          =   495
            Left            =   2190
            TabIndex        =   73
            Top             =   300
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   61276161
            UpDown          =   -1  'True
            CurrentDate     =   40553
         End
         Begin MSComCtl2.DTPicker dtpCloseAcc 
            Height          =   495
            Left            =   2190
            TabIndex        =   74
            Top             =   900
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   873
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   61276161
            UpDown          =   -1  'True
            CurrentDate     =   40553
         End
         Begin VB.Label lblCustomer 
            Alignment       =   1  'Right Justify
            Caption         =   "Open Account Date:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   11
            Left            =   120
            TabIndex        =   81
            Tag             =   "L23"
            Top             =   390
            Width           =   1875
         End
         Begin VB.Label lblCustomer 
            Alignment       =   1  'Right Justify
            Caption         =   "Closed  Account Date:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   12
            Left            =   120
            TabIndex        =   80
            Tag             =   "L24"
            Top             =   870
            Width           =   1875
         End
         Begin VB.Label lblCustomer 
            Alignment       =   2  'Center
            Caption         =   "C«ng nî hiÖn t¹i"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   480
            TabIndex        =   79
            Tag             =   "L27"
            Top             =   3000
            Width           =   1485
         End
         Begin VB.Label lblCurrentBalance 
            Alignment       =   2  'Center
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   2040
            TabIndex        =   78
            Top             =   2880
            Width           =   2385
         End
         Begin VB.Label lblCustomer 
            Alignment       =   1  'Right Justify
            Caption         =   "Nh©n viªn phô tr¸ch:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   15
            Left            =   120
            TabIndex        =   77
            Tag             =   "L25"
            Top             =   1620
            Width           =   1875
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "§iÓm tÝch lòy:"
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
            Left            =   390
            TabIndex        =   76
            Top             =   2340
            Width           =   1575
         End
         Begin VB.Label lblResult 
            Alignment       =   2  'Center
            Caption         =   "0"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   2070
            TabIndex        =   75
            Top             =   2220
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "§iÓm th­ëng"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   2055
         Left            =   -74760
         TabIndex        =   56
         Top             =   3000
         Width           =   6135
         Begin VB.TextBox txtBirthPoint 
            Alignment       =   1  'Right Justify
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
            Left            =   2160
            TabIndex        =   59
            Text            =   "0"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox txtSaleAmount 
            Alignment       =   1  'Right Justify
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
            Left            =   1920
            TabIndex        =   58
            Text            =   "0"
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox txtPointSale 
            Alignment       =   1  'Right Justify
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
            Left            =   3120
            TabIndex        =   57
            Text            =   "0"
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "®iÓm"
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
            Left            =   3480
            TabIndex        =   65
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label13 
            Caption         =   "Ngµy sinh nhËt tÆng:"
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
            Left            =   240
            TabIndex        =   63
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Doanh sè ®¹t"
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
            Left            =   360
            TabIndex        =   62
            Top             =   1020
            Width           =   1455
         End
         Begin VB.Label Label15 
            Caption         =   "Th­ëng:"
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
            Left            =   2160
            TabIndex        =   61
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label16 
            Caption         =   "®iÓm"
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
            Left            =   5040
            TabIndex        =   60
            Top             =   1500
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1815
         Left            =   -74760
         TabIndex        =   48
         Top             =   1080
         Width           =   6135
         Begin VB.OptionButton optType 
            Caption         =   "TÝch lòy ®iÓm theo mãn ¨n"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   480
            TabIndex        =   52
            Top             =   240
            Width           =   3375
         End
         Begin VB.OptionButton optType 
            Caption         =   "TÝch lòy ®iÓm theo doanh sè"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   480
            TabIndex        =   51
            Top             =   720
            Value           =   -1  'True
            Width           =   3375
         End
         Begin VB.TextBox txtAmountPoint 
            Alignment       =   1  'Right Justify
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
            Left            =   1800
            TabIndex        =   50
            Text            =   "0"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtPoint 
            Alignment       =   1  'Right Justify
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
            Left            =   3840
            TabIndex        =   49
            Text            =   "0"
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label Label10 
            Caption         =   "Sè tiÒn:"
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
            TabIndex        =   55
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label11 
            Caption         =   "="
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
            Left            =   3600
            TabIndex        =   54
            Top             =   1380
            Width           =   255
         End
         Begin VB.Label Label12 
            Caption         =   "§iÓm"
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
            Left            =   4680
            TabIndex        =   53
            Top             =   1370
            Width           =   495
         End
      End
      Begin VB.TextBox txtCustomer 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   13
         Left            =   2010
         TabIndex        =   9
         Text            =   "0"
         Top             =   4800
         Width           =   4815
      End
      Begin prjTouchScreen.MyButton cmdclear 
         Height          =   495
         Left            =   -69450
         TabIndex        =   45
         Tag             =   "L36"
         Top             =   5670
         Width           =   945
         _extentx        =   1667
         _extenty        =   873
         btype           =   5
         tx              =   "Hñy bá"
         enab            =   -1
         font            =   "frmCustomer.frx":0DF8
         coltype         =   2
         focusr          =   -1
         bcol            =   16578804
         bcolo           =   16777152
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmCustomer.frx":0E20
         picn            =   "frmCustomer.frx":0E3E
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdPrint 
         Height          =   495
         Left            =   -70560
         TabIndex        =   44
         Tag             =   "L35"
         Top             =   5670
         Width           =   1035
         _extentx        =   1826
         _extenty        =   873
         btype           =   5
         tx              =   "In TBCN"
         enab            =   -1
         font            =   "frmCustomer.frx":70DA
         coltype         =   2
         focusr          =   -1
         bcol            =   16578804
         bcolo           =   16777152
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmCustomer.frx":7102
         picn            =   "frmCustomer.frx":7120
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdPayment 
         Height          =   495
         Left            =   -70350
         TabIndex        =   43
         Tag             =   "L34"
         Top             =   5100
         Width           =   1635
         _extentx        =   2884
         _extenty        =   873
         btype           =   5
         tx              =   "Thanh to¸n"
         enab            =   -1
         font            =   "frmCustomer.frx":775C
         coltype         =   2
         focusr          =   -1
         bcol            =   16578804
         bcolo           =   16777152
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmCustomer.frx":7784
         picn            =   "frmCustomer.frx":77A2
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin VB.TextBox txtAmount 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   465
         Left            =   -72270
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   5610
         Width           =   1605
      End
      Begin VB.ComboBox cboBill 
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
         Left            =   -74850
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "TÊt c¶ Bill"
         Top             =   5670
         Width           =   2205
      End
      Begin MSFlexGridLib.MSFlexGrid flgPaymentHistory 
         Height          =   4065
         Left            =   -74910
         TabIndex        =   37
         Top             =   1020
         Width           =   6765
         _ExtentX        =   11933
         _ExtentY        =   7170
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   16711680
         BackColorFixed  =   16777215
         ForeColorFixed  =   16711680
         ForeColorSel    =   255
         BackColorBkg    =   16777215
         GridColor       =   8421504
         GridColorFixed  =   12632256
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   1
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid flgCustHistorySale 
         Height          =   4575
         Left            =   -74940
         TabIndex        =   33
         Top             =   1020
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   8070
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   -2147483641
         BackColorFixed  =   16777215
         ForeColorFixed  =   -2147483641
         ForeColorSel    =   255
         BackColorBkg    =   16777215
         GridColorFixed  =   12632256
         GridLinesFixed  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtCustomer 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   2010
         TabIndex        =   3
         Top             =   2820
         Width           =   4845
      End
      Begin VB.TextBox txtCustomer 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   4560
         TabIndex        =   25
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtCustomer 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   8
         Left            =   4920
         TabIndex        =   8
         Top             =   4290
         Width           =   1905
      End
      Begin VB.TextBox txtCustomer 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   2010
         TabIndex        =   0
         Top             =   1380
         Width           =   1485
      End
      Begin VB.TextBox txtCustomer 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   7
         Left            =   2010
         TabIndex        =   7
         Top             =   4290
         Width           =   1605
      End
      Begin VB.TextBox txtCustomer 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   4950
         TabIndex        =   6
         Top             =   3720
         Width           =   1905
      End
      Begin VB.TextBox txtCustomer 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   2010
         TabIndex        =   5
         Top             =   3750
         Width           =   1605
      End
      Begin VB.TextBox txtCustomer 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   2010
         TabIndex        =   4
         Top             =   3300
         Width           =   4845
      End
      Begin VB.TextBox txtCustomer 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   2010
         TabIndex        =   2
         Top             =   2340
         Width           =   4845
      End
      Begin prjTouchScreen.MyButton cmdSavepoint 
         Height          =   735
         Left            =   -72720
         TabIndex        =   64
         Top             =   5280
         Width           =   2055
         _extentx        =   3625
         _extenty        =   1296
         btype           =   5
         tx              =   "&L­u"
         enab            =   -1
         font            =   "frmCustomer.frx":7DDE
         coltype         =   2
         focusr          =   -1
         bcol            =   12632319
         bcolo           =   12632319
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmCustomer.frx":7E06
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin MSComCtl2.DTPicker dtpBirth 
         Height          =   495
         Left            =   4560
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   61276161
         UpDown          =   -1  'True
         CurrentDate     =   40553
      End
      Begin VB.Label Label3 
         Caption         =   "Tæng céng:"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72120
         TabIndex        =   82
         Top             =   5760
         Width           =   1455
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "C«ng nî cho phÐp"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   9
         Left            =   90
         TabIndex        =   46
         Tag             =   "L26"
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label lblAount 
         Alignment       =   2  'Center
         Caption         =   "Sè tiÒn"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -72360
         TabIndex        =   41
         Tag             =   "L33"
         Top             =   5160
         Width           =   1065
      End
      Begin VB.Label lblPayby 
         Alignment       =   2  'Center
         Caption         =   "H×nh thøc thanh to¸n"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   465
         Left            =   -73290
         TabIndex        =   40
         Tag             =   "L32"
         Top             =   5160
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label lblBill 
         Alignment       =   2  'Center
         Caption         =   "Chän hãa ®¬n thanh to¸n"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   465
         Left            =   -74910
         TabIndex        =   38
         Tag             =   "L31"
         Top             =   5160
         Width           =   1275
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "C«ng ty:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   10
         Left            =   90
         TabIndex        =   32
         Tag             =   "L4"
         Top             =   2940
         Width           =   1935
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "Date of Birth:"
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
         Index           =   6
         Left            =   3540
         TabIndex        =   26
         Tag             =   "L8"
         Top             =   1410
         Width           =   1065
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "Nhãm kh¸ch hµng:"
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
         Index           =   8
         Left            =   90
         TabIndex        =   24
         Tag             =   "L11"
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "Account No:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   3870
         TabIndex        =   23
         Tag             =   "L10"
         Top             =   4440
         Width           =   1005
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "Tax Code:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   90
         TabIndex        =   22
         Tag             =   "L9"
         Top             =   4380
         Width           =   1935
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "Fax:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4110
         TabIndex        =   21
         Tag             =   "7"
         Top             =   3870
         Width           =   765
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "Telephone:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   20
         Tag             =   "L6"
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer Code:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   3
         Left            =   90
         TabIndex        =   19
         Tag             =   "L2"
         Top             =   1140
         Width           =   1935
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "&Address:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   13
         Tag             =   "L5"
         Top             =   3390
         Width           =   1935
      End
      Begin VB.Label lblCustomer 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer &Name:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Tag             =   "L3"
         Top             =   2460
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   8100
      TabIndex        =   18
      Top             =   7980
      Width           =   6735
      Begin prjTouchScreen.MyButton cmdCustomer 
         Height          =   945
         Index           =   0
         Left            =   2400
         TabIndex        =   27
         Top             =   240
         Width           =   1845
         _extentx        =   3254
         _extenty        =   1667
         btype           =   5
         tx              =   "&L­u"
         enab            =   -1
         font            =   "frmCustomer.frx":7E24
         coltype         =   2
         focusr          =   -1
         bcol            =   14737632
         bcolo           =   12640511
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmCustomer.frx":7E4C
         picn            =   "frmCustomer.frx":7E6A
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdCustomer 
         Height          =   945
         Index           =   1
         Left            =   180
         TabIndex        =   28
         Top             =   210
         Width           =   1845
         _extentx        =   3254
         _extenty        =   1667
         btype           =   5
         tx              =   "&Thªm míi"
         enab            =   -1
         font            =   "frmCustomer.frx":83AE
         coltype         =   2
         focusr          =   -1
         bcol            =   14737632
         bcolo           =   12640511
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmCustomer.frx":83D6
         picn            =   "frmCustomer.frx":83F4
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdCustomer 
         Height          =   945
         Index           =   2
         Left            =   4620
         TabIndex        =   29
         Top             =   210
         Width           =   1845
         _extentx        =   3254
         _extenty        =   1667
         btype           =   5
         tx              =   "&Xãa"
         enab            =   -1
         font            =   "frmCustomer.frx":8848
         coltype         =   2
         focusr          =   -1
         bcol            =   14737632
         bcolo           =   12640511
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmCustomer.frx":8870
         picn            =   "frmCustomer.frx":888E
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdCustomer 
         Height          =   945
         Index           =   3
         Left            =   2400
         TabIndex        =   30
         Top             =   1290
         Width           =   1845
         _extentx        =   3254
         _extenty        =   1667
         btype           =   5
         tx              =   "&Gióp ®ì"
         enab            =   -1
         font            =   "frmCustomer.frx":8ECA
         coltype         =   2
         focusr          =   -1
         bcol            =   14737632
         bcolo           =   12640511
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmCustomer.frx":8EF2
         picn            =   "frmCustomer.frx":8F10
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdCustomer 
         Height          =   945
         Index           =   4
         Left            =   4620
         TabIndex        =   31
         Top             =   1290
         Width           =   1845
         _extentx        =   3254
         _extenty        =   1667
         btype           =   5
         tx              =   "Th&o¸t"
         enab            =   -1
         font            =   "frmCustomer.frx":954C
         coltype         =   2
         focusr          =   -1
         bcol            =   14737632
         bcolo           =   12640511
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmCustomer.frx":9574
         picn            =   "frmCustomer.frx":9592
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdCustomer 
         Height          =   945
         Index           =   5
         Left            =   180
         TabIndex        =   47
         Top             =   1320
         Width           =   1845
         _extentx        =   3254
         _extenty        =   1667
         btype           =   5
         tx              =   "In DSKH"
         enab            =   -1
         font            =   "frmCustomer.frx":F82E
         coltype         =   2
         focusr          =   -1
         bcol            =   14737632
         bcolo           =   12640511
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmCustomer.frx":F856
         picn            =   "frmCustomer.frx":F874
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
   End
   Begin VB.Label lblSearch 
      Alignment       =   1  'Right Justify
      Caption         =   "T×m kiÕm th«ng tin kh¸ch hµng:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   180
      TabIndex        =   34
      Tag             =   "L37"
      Top             =   9480
      Width           =   2025
   End
End
Attribute VB_Name = "frmCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsCustomer As New ADODB.Recordset
    Dim arrUpdate() As Variant
    Dim arrDelete() As Variant
    Dim arrAddNew() As String
    Dim fLoad As Boolean
    Dim fUpdate As Boolean
    Dim fActivate As Boolean
    Dim i, j, k As Integer
    Dim DateBalance As String
    Dim DescArr() As String

'           ----------- FORM ----------

Private Sub cboBill_Change()
    If cboBill.Text = "" Then cmdPayment.Enabled = False
End Sub

Private Sub cboPro_Change()
    UpdateData
End Sub

Private Sub cboPro_Click()
    UpdateData
End Sub

Private Sub cmCloseAcc_Click()
'    dtpOpenAcc.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    dtpCloseAcc.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    dtpCloseAcc.Enabled = True
    txtCustomer(11).Text = dtpCloseAcc.Value
    Call UpdateData
End Sub

Private Sub cmdclean_Click()
On Error GoTo Handle
        With rsCustomer
            .Find "CustNum='" & txtCustomer(0).Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Fields("Point") = 0
                .Fields("Totals") = 0
                .Fields("Acct_Balance") = 0
                .Update
            End If
        End With
        lblResult.Caption = 0
        lblCurrentBalance.Caption = 0
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub cmdCustomer_Click(Index As Integer)
On Error GoTo errHdl
    Dim res
    Select Case Index
        Case 0
            If fUpdate Then
                fUpdate = False
               arrUpdate = Add_UpdatedData_To_Array(flexCustomer, arrUpdate)
                Add_DataUpdate_To_DB
                MsgBox "CËp nhËt thµnh c«ng"
            End If
        Case 1:
'            AddNewAction
            With frmAddCustomer
                .Show vbModal
            End With
            Set rsCustomer = Open_Table(cnData, "Customer")
            fUpdate = True
            Call Initalize
            Call Init_Flex_Cust
        Case 2
            If Trim(txtCustomer(0).Text) = "101" Then
                'MsgBox "®©y lµ kh¸ch hµng mÆc ®Þnh, kh«ng thÓ xãa bá kh¸ch hµng nµy !!!"
                Exit Sub
            End If
                DeleteAction
                ControlState
        Case 3: MsgBox "Ch­a cã môc gióp ®ì, mong quý kh¸ch th«ng c¶m"
        Case 4:
            If Not fUpdate Then GoTo 1
            res = MsgBox("B¹n cã muèn l­u th«ng tin thay ®æi kh«ng ?", vbYesNoCancel)
            Select Case res
                Case vbNo:      GoTo 1
                Case vbCancel:  Exit Sub
                Case vbYes
                    Add_DataUpdate_To_DB
            End Select
1:
            CloseRecordset rsCustomer
            Unload Me
        Case 5
            Call Print_Cust_List
    End Select
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- cmdCustomer_Click"
End Sub

Private Sub cmdOpenAcc_Click()
On Error GoTo Handle
    dtpOpenAcc.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    'dtpCloseAcc.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    txtCustomer(10).Text = dtpOpenAcc.Value
    dtpOpenAcc.Enabled = True
    Call UpdateData
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name

End Sub

Private Sub cmdPayment_Click()
On Error GoTo Handle
Dim rsBill_Total As New ADODB.Recordset
            Set rsBill_Total = Open_Table(cnData, "Invoice_Totals")
    If MsgBox("B¹n muèn thanh to¸n hãa ®¬n nî nµy ?", vbYesNo) = vbYes Then
         If DateBalance = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") Then
            
            With rsBill_Total
                If Not .EOF And .RecordCount > 0 Then .MoveFirst
                .Find "Invoice_Number=" & CDbl(cboBill.Text), , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        .Fields("Payment_Method") = "CA"
                        .Update
                    End If
            End With
        Else
        'Cap nhat Phieu thu no cho Khach hang nay
            Call update_Balance(txtCustomer(0).Text, txtCustomer(1).Text, txtAmount.Text, Trim(cboBill.Text))
            With rsBill_Total
                If Not .EOF And .RecordCount > 0 Then .MoveFirst
                .Find "Invoice_Number=" & CDbl(cboBill.Text), , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        .Fields("Payment_Method") = "PA"
                        .Update
                    End If
            End With
        End If
        
        With rsCustomer
                If Not .EOF And .RecordCount > 0 Then .MoveFirst
                .Find "CustNum='" & txtCustomer(0).Text & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    .Fields("Acct_Balance") = CDbl(.Fields("Acct_Balance")) - CDbl(txtAmount.Text)
                    .Update
                End If
            End With
            Call Get_History_Sale(txtCustomer(0).Text)
    End If
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdPayment_Click"
End Sub

Private Sub dtpBirth_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    txtCustomer(6).Text = dtpBirth.Value
    UpdateData
End Sub

Private Sub dtpBirth_Change()
    txtCustomer(6).Text = dtpBirth.Value
    UpdateData
End Sub

Private Sub dtpBirth_Click()
    txtCustomer(6).Text = dtpBirth.Value
    UpdateData
End Sub

Private Sub dtpCloseAcc_Change()
    txtCustomer(11).Text = dtpCloseAcc.Value
    UpdateData
End Sub

Private Sub dtpCloseAcc_Click()
    txtCustomer(11).Text = dtpCloseAcc.Value
    UpdateData
End Sub


Private Sub dtpOpenAcc_Change()
    txtCustomer(10).Text = dtpOpenAcc.Value
    UpdateData
End Sub

Private Sub dtpOpenAcc_Click()
    txtCustomer(10).Text = dtpOpenAcc.Value
    UpdateData
End Sub

Private Sub flgPaymentHistory_Click()
On Error GoTo Handle
    cboBill.Text = flgPaymentHistory.TextMatrix(flgPaymentHistory.Row, 1)
    txtAmount.Text = flgPaymentHistory.TextMatrix(flgPaymentHistory.Row, 2)
    DateBalance = flgPaymentHistory.TextMatrix(flgPaymentHistory.Row, 0)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " flgPaymentHistory_Click"
End Sub

Public Sub Set_Promotion_type()
On Error GoTo Handle
Dim rsCust_Type As New ADODB.Recordset
Set rsCust_Type = Open_Table(cnData, "Customer_Type")
If rsCust_Type.State = 0 Then Exit Sub
If rsCust_Type.RecordCount = 0 Then Exit Sub
cboPro.Clear
With rsCust_Type
    Do While Not .EOF
       With cboPro
            .AddItem rsCust_Type.Fields("CustType_ID")
       End With
       .MoveNext
       Loop
End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Set_Promotion_type"
End Sub

Private Sub Form_Activate()
    Dim ctrl As Control
    If fActivate Then Exit Sub
    fActivate = True
    DescArr = LoadLanguage(LngFile, "#01:008:") '
    If cmdCustomer(0).Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    tabCustomer.TabCaption(0) = DescArr(22)
    tabCustomer.TabCaption(1) = DescArr(30)
    tabCustomer.TabCaption(2) = DescArr(28)
    tabCustomer.TabCaption(3) = DescArr(29)
     Call Init_Flex_Cust
    For Each ctrl In Me
    DoEvents
    'If ctrl.Tag < 18 Then
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    'End If
    Next ctrl
    Call UpdateData
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    With flexCustomer
        If Shift = 2 Then
            If KeyCode = vbKeyDown Then
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    If .Row > 23 Then .TopRow = .Row - 22
                End If
                KeyCode = 0
                flexCustomer_Click
            ElseIf KeyCode = vbKeyUp Then
                If .Row > 1 Then
                    .Row = .Row - 1
                    If .Row < .TopRow Then .TopRow = .Row
                End If
                KeyCode = 0
                flexCustomer_Click
            End If
        End If
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_KeyDown"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl
Call Set_Promotion_type
    With Me
        .WindowState = 0
    End With
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    If Check_Table_exist("Customer_Point_Sale") = False Then Create_Customer_Point_Sale
    
    If UCase(UserID) = "131112" Or UserID = "881507" Then cmdclean.Visible = True
    Set rsCustomer = Open_Table(cnData, "Customer")
    If rsCustomer.State = 0 Then Exit Sub
    Initalize
    txtCustomer(0).Locked = True
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Load"
End Sub

Private Sub Initalize()
On Error GoTo errHdl

    fLoad = False: fUpdate = False: fActivate = False
    InitArray
    SetDataInFlex
    ControlState
    With flexCustomer
'        SetColorFlexGrid flexCustomer, 1, 1, .Cols
        .Col = 1
        .Row = 1
        .ColSel = .Cols - 1
        .ScrollTrack = True
    End With
    For i = 0 To txtCustomer.count - 1
    DoEvents
        If i <> 9 Then
            txtCustomer(i).MaxLength = rsCustomer.Fields(i).DefinedSize
        End If
    Next i
    txtCustomer(8).MaxLength = 12
    txtCustomer(6).MaxLength = 12
    txtCustomer(11).MaxLength = 10
    flexCustomer_Click
    fLoad = True
    Call Get_Check_Point
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Initalize"
End Sub

'           --------- FLEXGRID ---------
Private Sub flexCustomer_Click()
On Error GoTo errHdl

    Dim ctrl As Control
    Dim j As Integer
    
    fLoad = False
    With flexCustomer
        If .TextMatrix(1, 0) = "" Then SetTextNull: Exit Sub
        If .Rows = 2 Then .Row = 1
        ControlState
        For Each ctrl In Me
        DoEvents
            If ctrl.Tag <> "" And ctrl.Tag <= .Cols Then
                If TypeOf ctrl Is TextBox Then
                    If ctrl.Tag = 9 Then
                        ctrl.Text = Format("0" & .TextMatrix(.Row, ctrl.Tag - 1), formatNum)
                    Else
                        If ctrl.Tag = 13 Then
                            ctrl.Text = Format("0" & .TextMatrix(.Row, ctrl.Tag - 1), "#,##0")
                        Else
                            ctrl.Text = .TextMatrix(.Row, ctrl.Tag - 1)
                        End If
                    End If
             End If
            
            End If
        Next ctrl
        lblNumber.Caption = .TextMatrix(.Row, 0)
        lblName.Caption = .TextMatrix(.Row, 1)
        cboPro.Text = .TextMatrix(.Row, 9)
        If txtCustomer(10).Text <> "" Then dtpOpenAcc.Value = txtCustomer(10).Text
        If txtCustomer(11).Text <> "" Then dtpCloseAcc.Value = txtCustomer(11).Text
        If txtCustomer(6).Text <> "" Then dtpBirth.Value = txtCustomer(6).Text
    End With
    fLoad = True
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- flexCustomer_Click"
End Sub

Private Sub flexCustomer_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 13 Then
        With txtCustomer(0)
            .SetFocus
            .SelStart = 0
            .SelLength = 9999
        End With
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- flexCustomer_KeyPress"
End Sub

Private Sub flexCustomer_EnterCell()
On Error GoTo errHdl

    If fLoad Then flexCustomer_Click
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- flexCustomer_EnterCell"
End Sub

Private Sub SetHeaderFlexGrid()
On Error GoTo errHdl

    With flexCustomer
        .AllowUserResizing = flexResizeBoth
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .Cols = rsCustomer.Fields.count - 2
        For i = 0 To .Cols - 1
        DoEvents
            Select Case i
                Case 0: .ColWidth(i) = 960: .ColAlignment(i) = 2
                Case 1 To 6: .ColWidth(i) = 520: .ColAlignment(i) = 2
                Case Else: .ColWidth(i) = 960: .ColAlignment(i) = 6
            End Select
        Next i
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- SetHeaderFlexGrid"
End Sub

Private Sub SetDataInFlex()
On Error GoTo errHdl

    Dim irow As Integer
    Dim sTemp As String
    
    irow = 1
    SetHeaderFlexGrid
    With rsCustomer
        If .RecordCount > 0 Then
           flexCustomer.Rows = .RecordCount + 1
           .Sort = "CustNum ASC"
           .MoveFirst
           Do While Not .EOF
           DoEvents
            For i = 0 To flexCustomer.Cols - 1
            DoEvents
                Select Case i
                    Case 0: sTemp = "CustNum"
                    Case 1: sTemp = "CustName"
                    Case 2: sTemp = "Company"
                    Case 3: sTemp = "Address"
                    Case 4: sTemp = "Phone"
                    Case 5: sTemp = "Fax"
                    Case 6: sTemp = "Birthday"
                    Case 7: sTemp = "TaxCode"
                    Case 8: sTemp = "AccountNo"
                    Case 9: sTemp = "Cust_Type"
                    Case 10: sTemp = "Acct_Open_Date"
                    Case 11: sTemp = "Acct_Close_Date"
                    Case 12: sTemp = "Cashier"
                    Case 13: sTemp = "Acct_Max_Balance"
                End Select
                If IsNull(.Fields(sTemp)) Then
                    flexCustomer.TextMatrix(irow, i) = ""
                Else
                    flexCustomer.TextMatrix(irow, i) = Trim(.Fields(sTemp))
                End If
            Next i
            irow = irow + 1
            .MoveNext
           Loop
        End If
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- SetDataInFlex"
End Sub
'           ----------- TEXTBOX ---------
Private Sub LockText(flag As Boolean)
On Error GoTo errHdl

    For i = 0 To txtCustomer.count - 1
    DoEvents
        If i <> 9 Then
            If flag Then
                txtCustomer(i).Locked = True
            Else: txtCustomer(i).Locked = False
            End If
        End If
    Next i
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- LockText"
End Sub

Private Sub optType_Click(Index As Integer)
    optType(Index).Value = True
    If optType(0).Value = True Then
    
        optType(0).Value = True
        optType(1).Value = False
        txtAmountPoint.Enabled = False
        txtPoint.Enabled = False
    Else
        optType(0).Value = False
        optType(1).Value = True
        txtAmountPoint.Enabled = True
        txtPoint.Enabled = True
    End If
End Sub

Private Sub txtAmountPoint_Change()
On Error GoTo Handle
    txtAmountPoint.Text = Format(txtAmountPoint.Text, "#,##0")
    txtAmountPoint.SelStart = Len(txtAmountPoint.Text)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtAmountPoint_Change"
End Sub

Private Sub txtAmountPoint_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
   If KeyAscii = 13 Then
        txtPoint.SetFocus
        txtPoint.SelStart = 0
        txtPoint.SelLength = 999
   End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtAmountPoint_KeyPress"
End Sub

Private Sub txtBirthPoint_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
   If KeyAscii = 13 Then
        txtSaleAmount.SetFocus
        txtSaleAmount.SelStart = 0
        txtSaleAmount.SelLength = 999
   End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtBirthPoint_KeyPress"
End Sub

Private Sub txtCustomer_Change(Index As Integer)
    If Index = 0 Then
        If txtCustomer(0).Text <> "101" Then
            Call Get_History_Sale(txtCustomer(0).Text)
            Call Get_History_Point(txtCustomer(0).Text)
        End If
    End If
    If Index = 1 Then
        If txtCustomer(Index).Text = "" Then MsgBox "Tªn kh¸ch hµng kh«ng ®­îc rçng "
    End If
   
End Sub

Private Sub txtCustomer_DblClick(Index As Integer)
    If Index <> 6 Then
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtCustomer(Index).Text = .Let_Text_Input
            Call UpdateData
        End With
    End If
End Sub

Private Sub txtCustomer_GotFocus(Index As Integer)
    On Error GoTo Handle
    txtCustomer(Index).SelStart = 0
    txtCustomer(Index).SelLength = 9999
    
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "txtCustomer_GotFocus"
End Sub

Private Sub txtCustomer_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl

    Dim tempIndex As Integer
    
    If KeyAscii = 33 Or KeyAscii = 44 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 Then
        Select Case Index
            Case 13:
                cmdCustomer(0).SetFocus
                txtCustomer(13).Text = Format(txtCustomer(13).Text, "#,##0")
            Case Else
                   tempIndex = Index + 1
                  If tempIndex <> -1 And tempIndex <> 9 And tempIndex <> 10 Then
                      With txtCustomer(tempIndex)
                          .SetFocus
                          .SelStart = 0
                          .SelLength = 9999
                      End With
                  End If
            End Select
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- txtCustomer_KeyPress"
End Sub

Private Sub txtCustomer_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    Select Case KeyCode
        Case 8 'key backspace
        Case Is < 32, vbKeyDown, vbKeyUp: Exit Sub
    End Select
    UpdateData
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- txtCustomer_KeyUp"
End Sub

Private Sub SetTextNull()
On Error GoTo errHdl

    For i = 0 To txtCustomer.count - 1
    DoEvents
        txtCustomer(i).Text = ""
    Next i
    If fLoad Then txtCustomer(0).SetFocus
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- SetTextNull"
End Sub

'           ------------ UPDATE DATA ----------
Private Sub UpdateData()
On Error GoTo errHdl

    Dim sTemp() As String
    Dim i As Integer
    
    With flexCustomer
        If .TextMatrix(1, 0) = "" Then Exit Sub
        fUpdate = True
        sTemp = SetTextTemp
        For i = 0 To UBound(sTemp) Step 1
        DoEvents
            .TextMatrix(.Row, i) = sTemp(i)
        Next i
        lblName.Caption = sTemp(1)
    End With
    arrUpdate = Add_UpdatedData_To_Array(flexCustomer, arrUpdate)
    
    Exit Sub
errHdl:
   'MsgBox "O"
End Sub

Private Function SetTextTemp()
On Error GoTo errHdl

    Dim S1() As String
    Dim i As Integer
    
    With flexCustomer
        ReDim Preserve S1(.Cols - 2)
        's1(0) = .TextMatrix(.Row, 0)
        For i = 0 To 13
            If i <> 9 Then
                If txtCustomer(i).Text = "" Then
                    If i = 6 Then
                        S1(i) = "03/05/1981"
                    ElseIf i = 10 Or i = 11 Then
                        S1(i) = Date
                    Else
                        S1(i) = 0
                    End If
                Else
                    S1(i) = txtCustomer(i).Text
                End If
            Else
                S1(9) = cboPro.Text
            End If
        Next i
    End With
    SetTextTemp = S1
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- SetTextTemp"
End Function
'           ----- ADD UPDATED DATA TO DATABASE ----
Private Sub Add_DataUpdate_To_DB()
On Error GoTo errHdl

    Dim sSQL As String, iIndex As String
    Dim fDelete As Boolean
    Dim i As Integer
    Dim sFieldName As String
    
    fDelete = True
    'add record masked delete on grid to arrUpdate
    With flexCustomer
        For i = 1 To UBound(arrDelete)
        DoEvents
            For j = 1 To .Rows - 1
            DoEvents
                If .TextMatrix(j, 0) = arrDelete(i)(0) Then
                    fDelete = False
                    arrUpdate = Add_UpdatedData_To_Array(flexCustomer, arrUpdate)
                    Exit For
                End If
            Next j
            If fDelete Then
                sSQL = "Delete from Customer where CustNum='" & arrDelete(i)(0) & "'"
                cnData.Execute sSQL
            End If
        Next i
    End With
'   Update updated data on grid to DB
    With rsCustomer
        For i = 1 To UBound(arrUpdate)
        DoEvents
            If .RecordCount = 0 Then
                .addNew
            Else
                .MoveFirst
                .Find "CustNum='" & arrUpdate(i)(0) & "'"
                If .EOF Then .addNew       'AddNew new records
            End If
            For j = 0 To flexCustomer.Cols - 2 'update old records
            DoEvents
                Select Case j
                    Case 0: sFieldName = "CustNum"
                    Case 1: sFieldName = "CustName"
                    Case 2: sFieldName = "Company"
                    Case 3: sFieldName = "Address"
                    Case 4: sFieldName = "Phone"
                    Case 5: sFieldName = "Fax"
                    Case 6: sFieldName = "Birthday"
                    Case 7: sFieldName = "TaxCode"
                    Case 8: sFieldName = "AccountNo"
                    Case 9: sFieldName = "Cust_Type"
                    Case 10: sFieldName = "Acct_Open_Date"
                    Case 11: sFieldName = "Acct_Close_Date"
                    Case 12: sFieldName = "Cashier"
                    Case 13: sFieldName = "Acct_Max_Balance"
                End Select
               If j = 13 Then
                    .Fields(sFieldName) = CDbl("0" & arrUpdate(i)(j))
                ElseIf j = 10 Or j = 11 Then
                    If arrUpdate(i)(j) = "" Then
                        .Fields(sFieldName) = Date
                    Else
                        .Fields(sFieldName) = arrUpdate(i)(j)
                    End If
                Else
                    If arrUpdate(i)(j) = "" Then
                        If j = 2 Or j = 3 Or j = 4 Or j = 5 Or j = 7 Or j = 8 Or j = 14 Then
                            .Fields(sFieldName) = 0
                        Else
                            .Fields(sFieldName) = "-"
                        End If
                    Else
                        .Fields(sFieldName) = arrUpdate(i)(j)
                    End If
                End If
            Next j
            .Update
        Next i
    End With
    
    Exit Sub
errHdl:
'    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
'    & Me.name & "- Add_DataUpdate_To_DB"
End Sub
Private Sub Get_Array_AddNew() 'append addnew records from frmAddNewCust to arrAddNew()
On Error GoTo errHdl

    Dim arrTemp() As String
    
    arrTemp = frmAddNewCust.Get_AddNewRecords
    For i = 1 To UBound(arrTemp)
    DoEvents
        ReDim Preserve arrAddNew(UBound(arrAddNew) + i)
        arrAddNew(UBound(arrAddNew)) = arrTemp(i)
    Next i
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Get_Array_AddNew"
End Sub
Private Sub Get_Array_Delete()
On Error GoTo errHdl

    Dim Arr() As String
    Dim fDelete As Boolean
    
    fDelete = False
    With flexCustomer
        For i = 1 To UBound(arrDelete)
        DoEvents
            If arrDelete(i)(0) = .TextMatrix(.Row, 0) Then _
                fDelete = True: Exit For
        Next i
        If Not fDelete Then
            ReDim Preserve Arr(.Cols - 2)
            For i = 0 To .Cols - 2
            DoEvents
                Arr(i) = .TextMatrix(.Row, i)
            Next i
            ReDim Preserve arrDelete(UBound(arrDelete) + 1)
            arrDelete(UBound(arrDelete)) = Arr()
        End If
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Get_Array_Delete"
End Sub

Private Sub ControlState()
On Error GoTo errHdl

    Dim iIndex As Byte
  
        HideControl False
        iIndex = 2
        For i = 0 To UBound(arrDelete)
        DoEvents
            If arrDelete(i)(0) <> "" Then
                cmdCustomer(2).Enabled = True
                cmdCustomer(3).Enabled = True
                Exit For
            End If
        Next i
    
    If flexCustomer.TextMatrix(1, 0) <> "" Then
        cmdCustomer(iIndex).Enabled = True
        tabCustomer.Enabled = True
        LockText False                'unlock text
    Else
        tabCustomer.Enabled = False
        LockText True
        lblNumber.Caption = "Customer No"
        lblName.Caption = "Customer Name"
    End If
    Dim j As Integer
    For j = 0 To 13 'flexCustomer.Cols
        If j <> 9 Then
            txtCustomer(j).Tag = j + 1
        End If
    Next
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- ControlState"
End Sub

Private Sub InitArray()
On Error GoTo errHdl

    Dim Arr() As String
    
    ReDim Preserve arrUpdate(0)
    ReDim Preserve arrAddNew(0)
    ReDim Preserve arrDelete(0)
    ReDim Preserve Arr(flexCustomer.Cols - 2)
    arrDelete(0) = Arr()
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- InitArray"
End Sub

Private Sub Refresh_Array(arrAdd() As String, arrDel() As Variant, fAdd As Boolean)
On Error GoTo errHdl

    If fAdd Then
        For i = 0 To UBound(arrAdd)
        DoEvents
            For j = 0 To UBound(arrDel)
            DoEvents
                If arrAdd(i) = arrDel(j)(0) Then _
                    arrDel(j)(0) = ""
            Next j
        Next i
    Else
        For i = 0 To UBound(arrDelete)
        DoEvents
            For j = 0 To UBound(arrAddNew)
            DoEvents
                If arrDelete(i)(0) = arrAddNew(j) Then _
                    arrAddNew(j) = ""
            Next j
        Next i
    End If
    fUpdate = True
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Refresh_Array"
End Sub

Private Sub HideControl(flag As Boolean)
On Error GoTo errHdl

    txtCustomer(4).Visible = Not flag
    lblCustomer(2).Visible = Not flag
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- HideControl"
End Sub
Private Sub Add_Value_To_ArrDelete(fInactive As Boolean, valueStatus As String, sColor As String)
On Error GoTo errHdl

    Dim Arr() As String
    Dim fDelete As Boolean
    Dim iInc As Integer
    
    fDelete = False
    With flexCustomer
        ReDim Preserve Arr(.Cols - 2)
        .TextMatrix(.Row, .Cols - 2) = valueStatus
        For iInc = 0 To .Cols - 2
        DoEvents
            .Col = iInc: .CellForeColor = sColor
        Next iInc
        If fInactive Then 'click cmdInactive danh dau record duoc xoa
            cmdCustomer(3).Enabled = True
            For iInc = 1 To UBound(arrDelete)
            DoEvents
                If arrDelete(iInc)(0) = .TextMatrix(.Row, 0) Then _
                    fDelete = True: Exit For
            Next iInc
            If Not fDelete Then
                ReDim Preserve arrDelete(UBound(arrDelete) + 1)
                For iInc = 0 To .Cols - 2
                DoEvents
                    Arr(iInc) = .TextMatrix(.Row, iInc)
                Next iInc
                arrDelete(UBound(arrDelete)) = Arr()
            End If
            
        Else 'click cmdInactive huy bo danh dau record duoc  xoa
            For iInc = 0 To UBound(arrDelete)
            DoEvents
                If arrDelete(iInc)(0) = .TextMatrix(.Row, 0) Then _
                    arrDelete(iInc)(0) = ""
            Next iInc
            For iInc = 0 To UBound(arrDelete)
            DoEvents
                If arrDelete(iInc)(0) <> "" Then fDelete = True
            Next iInc
            If Not fDelete Then cmdCustomer(3).Enabled = False
        End If
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Add_Value_To_ArrDelete"
End Sub
Private Sub AddNewAction()
On Error GoTo errHdl

    Dim iMaxRecord As Integer
    With frmAddNewCust
        .Show vbModal
    End With
    fUpdate = True
    UpdateData
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- AddNewAction"
End Sub
Private Sub Add_Record_To_ArrInactive()
On Error GoTo errHdl

    With flexCustomer
        If .TextMatrix(.Row, .Cols - 2) = "00" Then
              Add_Value_To_ArrDelete True, "80", vbRed
        Else: Add_Value_To_ArrDelete False, "00", vbBlack
        End If
        fUpdate = True
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Add_Record_To_ArrInactive"
End Sub

Private Sub DeleteAction()
On Error GoTo errHdl

    Dim iNumDelete As Integer
    Dim iTemp As Integer
    Dim iInc As Integer
    
    iNumDelete = 0
    If MsgBox("B¹n cã muèn xãa kh¸ch hµng nµy?", vbOKCancel) <> 1 Then Exit Sub
    With flexCustomer
'       delete row out of flex
        Get_Array_Delete
        For iInc = .Rows - 1 To 1 Step -1
        DoEvents
            If iNumDelete = UBound(arrDelete) Then Exit For
            For j = 0 To UBound(arrDelete)
            DoEvents
                If arrDelete(j)(0) = .TextMatrix(iInc, 0) Then
                    iNumDelete = iNumDelete + 1
                    If .Rows = 2 Then ' if deleted row is last row then set it Null
                        For k = 0 To .Cols - 2
                        DoEvents
                            .TextMatrix(iInc, k) = ""
                        Next k
                    Else
                        .RemoveItem iInc
'                        SetColorFlexGrid flexCustomer, iInc - 1, 1, .Cols
                    End If
                    Exit For
                End If
            Next j
        Next iInc
        If .TextMatrix(1, 0) = "" Then SetTextNull
    End With
    If UBound(arrDelete) > 0 Then _
        Refresh_Array arrAddNew, arrDelete, False
        
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- DeleteAction"
End Sub

Public Sub AddNewRecords()
On Error GoTo Handle

    Get_Array_AddNew

    If UBound(arrAddNew) > 0 Then _
        Refresh_Array arrAddNew, arrDelete, True
    flexCustomer_Click
    fUpdate = True
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " AddNewRecords"
End Sub

Public Sub Get_History_Sale(CustomerNo As String)
On Error GoTo Handle
    Dim strSale_With_Cust As String
    Dim strBalance As String
    Dim rsCustSale As New ADODB.Recordset
    Dim rsBalance As New ADODB.Recordset
    
    strSale_With_Cust = "Select Left(DateTime,8) as DateSale,Invoice_Number,Grand_Total,Payment_Method from Invoice_Totals where CustNum='" & CustomerNo & "'"
    strBalance = "Select Left(DateTime,8) as DateSale,Invoice_Number,OA_Amount,Payment_Method from Invoice_Totals where CustNum='" & CustomerNo & "' and OA_Amount>0"
    
    Set rsCustSale = OpenCriticalTable(strSale_With_Cust, cnData)
    Set rsBalance = OpenCriticalTable(strBalance, cnData)
    
    Call Set_SaleHistory_Flex(rsCustSale)
    Call Set_Balance_Flex(rsBalance)
    Call Account_Balance(CustomerNo)
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Get_History_Sale"

End Sub

Public Sub Set_SaleHistory_Flex(rs As Recordset)
On Error GoTo Handle
Dim i, j As Integer
Dim totals As Double
    With flgCustHistorySale
        .Cols = 4
        If rs.RecordCount = 0 Then
            .Rows = 2
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
        End If
        .ColWidth(0) = 1500
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 2800
        .ColAlignment(0) = 2
        .ColAlignment(1) = 4
        .ColAlignment(2) = 6
        .ColAlignment(3) = 2
        .TextMatrix(0, 0) = "Ngµy mua hµng" 'DescArr(0)
        .TextMatrix(0, 1) = "Sè hãa ®¬n" 'DescArr(1)
        .TextMatrix(0, 2) = "Sè tiÒn" 'DescArr(2)
        .TextMatrix(0, 3) = "H×nh thøc thanh to¸n" 'DescArr(3)
        If rs.RecordCount > 0 Then
            flgCustHistorySale.Rows = rs.RecordCount + 1
            Do While Not rs.EOF
                i = i + 1
                .TextMatrix(i, 0) = rs.Fields("DateSale")
                .TextMatrix(i, 1) = rs.Fields("Invoice_Number")
                .TextMatrix(i, 2) = Format(rs.Fields("Grand_Total"), "#,##0")
                totals = totals + rs.Fields("Grand_Total")
                Select Case rs.Fields("Payment_Method")
                    Case "OA"
                        .TextMatrix(i, 3) = "C«ng nî"
                    Case "CA"
                        .TextMatrix(i, 3) = "TiÒn mÆt"
                    Case "CC"
                        .TextMatrix(i, 3) = "ThÎ tÝn dông"
                    Case "CH"
                        .TextMatrix(i, 3) = "ChuyÓn kho¶n"
                End Select
            rs.MoveNext
            Loop
        End If
    End With
    txtTotals.Text = Format(totals, "#,##0")
'    If txtAmountPoint.Text <> 0 Then
'        i = Int(Totals / txtAmountPoint.Text)
'        Do Until i = 0
'            J = J + txtPoint.Text
'            i = i - 1
'        Loop
'        lblResult.Caption = J
'    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "Set_SaleHistory_Flex"
End Sub
Public Sub Set_Balance_Flex(rs As Recordset)
On Error GoTo Handle
Dim i As Integer
    With flgPaymentHistory
        .Cols = 3
        If rs.RecordCount = 0 Then
            .Rows = 2
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
        End If
        .ColWidth(0) = 1800
        .ColWidth(1) = 1800
        .ColWidth(2) = 2500
        .ColAlignment(0) = 2
        .ColAlignment(1) = 4
        .ColAlignment(2) = 6
        .TextMatrix(0, 0) = "Ngµy mua hµng" 'DescArr(0)
        .TextMatrix(0, 1) = "Sè hãa ®¬n" 'DescArr(1)
        .TextMatrix(0, 2) = "Sè tiÒn" 'DescArr(2)
        If rs.RecordCount > 0 Then
            flgPaymentHistory.Rows = rs.RecordCount + 1
            Do While Not rs.EOF
                i = i + 1
                .TextMatrix(i, 0) = rs.Fields("DateSale")
                .TextMatrix(i, 1) = rs.Fields("Invoice_Number")
                .TextMatrix(i, 2) = Format(rs.Fields("OA_Amount"), "#,##0")
            rs.MoveNext
            Loop
        End If
    End With
    Call Add_Bill_Balance(rs)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "Set_Balance_Flex"
End Sub

Public Sub Add_Bill_Balance(rs As Recordset)
On Error GoTo Handle
    Dim i As Integer
    With cboBill
        .Clear
        .AddItem "TÊt c¶", 0
        If rs.RecordCount > 0 And Not rs.EOF Then rs.MoveFirst
        Do While Not rs.EOF
            i = i + 1
            .AddItem rs.Fields(0), i
            rs.MoveNext
        Loop
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Add_Bill_Balance"

End Sub

Public Sub Account_Balance(S As String)
On Error GoTo Handle
    With rsCustomer
        If Not .EOF And .RecordCount > 0 Then .MoveFirst
        .Find "CustNum='" & S & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                lblCurrentBalance.Caption = Format(.Fields("Acct_Balance"), "#,##0")
            End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "Account_Balance"
End Sub

Public Sub ClearTextbox()
On Error GoTo Handle
    Dim i As Integer
    For i = 0 To 13
            Select Case i
            Case 6: dtpBirth.Value = gfCONVERT_STRING_TO_DATE(DateDefault) 'txtCustomer(6).Text = "10/09/1980"
            Case 9: txtCustomer(9).Text = 0
            Case 10: dtpOpenAcc.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
            Case 11: dtpCloseAcc.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
            Case Else
                txtCustomer(i).Text = ""
        End Select
    Next i
    txtCustomer(0).SetFocus
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "ClearTextbox"
End Sub

Public Sub update_Balance(cust_Num As String, Cust_Name As String, Sotien As Double, HDNo As String)
On Error GoTo Handle
    Dim rsPhieuthu As New ADODB.Recordset
    Dim sophieu As String
    
    Set rsPhieuthu = Open_Table(cnData, "Income")
    
    sophieu = GetMaxSophieuThu()
    
    With rsPhieuthu
        .Find "ID='" & sophieu & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
              .addNew
              .Fields("ID") = sophieu
              .Fields("Store_ID") = Store_ID
              .Fields("Cashier_ID") = UserID
              .Fields("DateTime") = DateDefault
              .Fields("Receipt_ID") = "CN"
              .Fields("Customer_ID") = cust_Num
              .Fields("Reciever_Name") = Cust_Name
              .Fields("Division") = " "
              .Fields("Payment_Method") = "TiÒn MÆt"
              .Fields("Amount") = Sotien
              .Fields("Description") = "Thanh to¸n tiÒn nî Hãa ®¬n sè" & HDNo
              .Update
              .Requery
            End If
        End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " update_Balance"
End Sub

Public Sub Get_History_Point(ByVal code As String)
    On Error GoTo Handle
        With rsCustomer
            .Find "CustNum='" & code & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                lblResult.Caption = .Fields("Point")
            End If
        End With
        
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Public Sub Print_Cust_List()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
    Dim iReport As CRAXDDRT.Report

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = "SELECT Customer.CustNum, Customer.CustName, Customer.Address, Customer.Phone, Customer.Birthday, Customer.TaxCode,Customer.Point,Customer.Acct_Balance" & _
          " FROM Customer order by Customer.CustName"
    Set crCustList = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crCustList
        .Database.AddADOCommand cnData, cmd
        .txtCustID.SetUnboundFieldSource "{ado.CustNum}"
        .txtCustNam.SetUnboundFieldSource "{ado.CustName}"
        .txtCustAdd.SetUnboundFieldSource "{ado.Address}"
        .txtCustPhone.SetUnboundFieldSource "{ado.Phone}"
        .txtCustDate.SetUnboundFieldSource "{ado.Birthday}"
        .txtPointSave.SetUnboundFieldSource "{ado.Point}"
        .txtAccBalance.SetUnboundFieldSource "{ado.Acct_Balance}"
        
    End With
    Set iReport = crCustList
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Print_Cust_List"

End Sub

Public Sub Get_Check_Point()
On Error GoTo Handle
    Dim rsPoint As New ADODB.Recordset
    Set rsPoint = Open_Table(cnData, "Customer_Point_Sale")
    With rsPoint
        If Not .EOF Then
            If .Fields("ID") = 1 Then
                optType(0).Value = True
                optType(1).Value = False
                txtAmountPoint.Enabled = False
                txtPoint.Enabled = False
            Else
                optType(0).Value = False
                optType(1).Value = True
                txtAmountPoint.Enabled = True
                txtPoint.Enabled = True
                txtAmountPoint.Text = Format(.Fields("Amount_Get_Point"), "#,##0")
                txtPoint.Text = .Fields("Point")
                
            End If
            txtBirthPoint.Text = .Fields("BirthPoint")
            txtSaleAmount.Text = Format(.Fields("AmountSale"), "#,##0")
            txtPointSale.Text = .Fields("PointSale")
        End If
    End With
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - Get_Check_Point"
End Sub
Private Sub cmdSavepoint_Click()
On Error GoTo Handle
    Dim rsPoint As New ADODB.Recordset
    Set rsPoint = Open_Table(cnData, "Customer_Point_Sale")
    Dim i As Integer
    i = 2
    With rsPoint
    .Find "ID=2", , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields("TypeMismatch_Name") = "TÝch lòy theo doanh sè"
        End If
       If optType(0).Value = True Then
            .Fields("ID") = 1
        Else
            .Fields("ID") = 2
        End If
        .Fields("Amount_Get_Point") = txtAmountPoint.Text
        .Fields("Point") = txtPoint.Text
        .Fields("BirthPoint") = txtBirthPoint.Text
        .Fields("AmountSale") = txtSaleAmount.Text
        .Fields("PointSale") = txtPointSale.Text
        .Update
    End With
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - Get_Check_Point"
End Sub


Private Sub txtPoint_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
   If KeyAscii = 13 Then
        txtBirthPoint.SetFocus
        txtBirthPoint.SelStart = 0
        txtBirthPoint.SelLength = 999
   End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtPoint_KeyPress"
End Sub


Private Sub txtPointSale_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
   If KeyAscii = 13 Then
        cmdSavepoint.SetFocus
        
   End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtPointSale_KeyPress"
End Sub

Private Sub txtSaleAmount_Change()
On Error GoTo Handle
    txtSaleAmount.Text = Format(txtSaleAmount.Text, "#,##0")
    txtSaleAmount.SelStart = Len(txtSaleAmount.Text)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtSaleAmount_Change"
End Sub

Private Sub txtSaleAmount_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
   If KeyAscii = 13 Then
        txtPointSale.SetFocus
        txtPointSale.SelStart = 0
        txtPointSale.SelLength = 999
   End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtSaleAmount_KeyPress"
End Sub

Public Sub Init_Flex_Cust()

With flexCustomer
        .ColWidth(0) = 1200
        .ColWidth(1) = 2000
        .ColWidth(2) = 2500
        .ColWidth(3) = 3000
        .ColWidth(4) = 1500
        .ColWidth(5) = 1200
        .ColWidth(6) = 1400
        .ColWidth(7) = 1400
        .ColWidth(8) = 1400
        .ColWidth(9) = 1400
        .ColWidth(10) = 1400
        .ColWidth(11) = 1500
        .ColWidth(12) = 1200
        .ColWidth(13) = 1200
       
        .TextMatrix(0, 0) = DescArr(2)
        .TextMatrix(0, 1) = DescArr(3)
        .TextMatrix(0, 2) = DescArr(4)
        .TextMatrix(0, 3) = DescArr(5)
        .TextMatrix(0, 4) = DescArr(6)
        .TextMatrix(0, 5) = DescArr(7)
        .TextMatrix(0, 6) = DescArr(8)
        .TextMatrix(0, 7) = DescArr(9)
        .TextMatrix(0, 8) = DescArr(10)
        .TextMatrix(0, 9) = DescArr(11)
        .TextMatrix(0, 10) = DescArr(23)
        .TextMatrix(0, 11) = DescArr(24)
        .TextMatrix(0, 12) = DescArr(25)
        .TextMatrix(0, 13) = DescArr(26)
        
    End With
End Sub
