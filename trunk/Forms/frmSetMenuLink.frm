VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSetMenuLink 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "§Þnh l­îng cho menu"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15045
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   8160
      TabIndex        =   19
      Top             =   8850
      Width           =   6135
      Begin prjTouchScreen.MyButton cmdSave 
         Height          =   855
         Left            =   120
         TabIndex        =   25
         Top             =   1050
         Width           =   1815
         _ExtentX        =   3201
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
         BCOL            =   16578804
         BCOLO           =   16777152
         FCOL            =   12582912
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSetMenuLink.frx":0000
         PICN            =   "frmSetMenuLink.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton MineralReport 
         Height          =   855
         Left            =   2040
         TabIndex        =   20
         Tag             =   "L15"
         Top             =   1050
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "In DS SetMenu"
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
         BCOLO           =   16777152
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSetMenuLink.frx":0560
         PICN            =   "frmSetMenuLink.frx":057C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAdd 
         Height          =   855
         Left            =   120
         TabIndex        =   21
         Tag             =   "L2"
         Top             =   150
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "&Thªm mãn"
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
         MICON           =   "frmSetMenuLink.frx":2D2E
         PICN            =   "frmSetMenuLink.frx":2D4A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdDel 
         Height          =   855
         Left            =   2040
         TabIndex        =   22
         Tag             =   "L3"
         Top             =   150
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "&Xãa mãn "
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
         MICON           =   "frmSetMenuLink.frx":319C
         PICN            =   "frmSetMenuLink.frx":31B8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdSearch 
         Height          =   855
         Left            =   3960
         TabIndex        =   23
         Top             =   150
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "T&×m kiÕm"
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
         MICON           =   "frmSetMenuLink.frx":37F2
         PICN            =   "frmSetMenuLink.frx":380E
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
         Height          =   855
         Left            =   3960
         TabIndex        =   24
         Tag             =   "L5"
         Top             =   1050
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "&§ãng "
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
         MICON           =   "frmSetMenuLink.frx":3E48
         PICN            =   "frmSetMenuLink.frx":3E64
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
   Begin VB.PictureBox picWait 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   1395
      ScaleHeight     =   1980
      ScaleWidth      =   4905
      TabIndex        =   10
      Top             =   3525
      Visible         =   0   'False
      Width           =   4965
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   0
         Width           =   4910
      End
      Begin MSComctlLib.ProgressBar probarWait 
         Height          =   390
         Left            =   0
         TabIndex        =   11
         Top             =   1530
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblWait 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Please Wait...."
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   1050
         TabIndex        =   14
         Top             =   675
         Width           =   2865
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1625
         Left            =   0
         TabIndex        =   13
         Top             =   315
         Width           =   4890
      End
   End
   Begin VB.PictureBox picLabel 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   7230
      ScaleHeight     =   765
      ScaleWidth      =   6600
      TabIndex        =   5
      Top             =   0
      Width           =   6660
      Begin VB.Label lblNo 
         BackColor       =   &H80000008&
         Caption         =   "Number"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblLookup 
         BackColor       =   &H80000008&
         Caption         =   "Lookup Key"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2280
         TabIndex        =   6
         Top             =   240
         Width           =   5295
      End
   End
   Begin TabDlg.SSTab tabSetMLink 
      Height          =   7860
      Left            =   7275
      TabIndex        =   0
      Top             =   960
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   13864
      _Version        =   393216
      Tabs            =   2
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
      TabCaption(0)   =   "§Þnh l­îng me nu"
      TabPicture(0)   =   "frmSetMenuLink.frx":A0FE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmLookup"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mãn ch­a ®Þnh møc"
      TabPicture(1)   =   "frmSetMenuLink.frx":A11A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dtgNot_Set"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin MSDataGridLib.DataGrid dtgNot_Set 
         Height          =   7215
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   12726
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         HeadLines       =   1
         RowHeight       =   27
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
            Size            =   14.25
            Charset         =   0
            Weight          =   700
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
      Begin VB.Frame frmLookup 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7440
         Left            =   90
         TabIndex        =   1
         Top             =   610
         Width           =   7545
         Begin prjTouchScreen.MyButton cmdMPlu 
            Height          =   375
            Left            =   5880
            TabIndex        =   26
            Top             =   540
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   661
            BTYPE           =   14
            TX              =   "..."
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
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   16711680
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmSetMenuLink.frx":A136
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdDelete 
            Height          =   675
            Left            =   6540
            TabIndex        =   18
            Top             =   2040
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1191
            BTYPE           =   14
            TX              =   ""
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
            BCOL            =   16578804
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmSetMenuLink.frx":A152
            PICN            =   "frmSetMenuLink.frx":A16E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdAddPLU 
            Height          =   615
            Left            =   5010
            TabIndex        =   17
            Tag             =   "L7"
            Top             =   1020
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1085
            BTYPE           =   14
            TX              =   "Thªm vµo ®Þnh l­îng"
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
            MICON           =   "frmSetMenuLink.frx":A7A8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdPickup 
            Height          =   615
            Left            =   3420
            TabIndex        =   16
            Tag             =   "L6"
            Top             =   1020
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1085
            BTYPE           =   14
            TX              =   "Chän nhanh"
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
            MICON           =   "frmSetMenuLink.frx":A7C4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin VB.TextBox txtStockRate 
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
            Height          =   450
            Index           =   0
            Left            =   4080
            TabIndex        =   15
            Top             =   2430
            Width           =   1605
         End
         Begin MSFlexGridLib.MSFlexGrid flexMPLU 
            Height          =   5355
            Left            =   75
            TabIndex        =   9
            Top             =   1995
            Width           =   6465
            _ExtentX        =   11404
            _ExtentY        =   9446
            _Version        =   393216
            BackColorFixed  =   -2147483643
            BackColorBkg    =   -2147483643
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
         Begin VB.ComboBox cboSelect 
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
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   540
            Width           =   5745
         End
         Begin VB.Label lblData 
            Caption         =   "PLU from Set Menu PLU:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Tag             =   "L9"
            Top             =   210
            Width           =   2205
         End
         Begin VB.Label lblData 
            Caption         =   "&Content:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   3
            Tag             =   "L10"
            Top             =   1605
            Width           =   975
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   10740
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   18944
      _Version        =   393216
      BackColorFixed  =   -2147483643
      BackColorBkg    =   -2147483643
      GridColorFixed  =   12632256
      ScrollTrack     =   -1  'True
      TextStyleFixed  =   3
      GridLinesFixed  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSetMenuLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim DescArr() As String
    Private Type subMenu
        PluCode     As String
        StockRate   As Double
        PluName     As String
    End Type
    
    Private Type Menu
        PluCode As String
        LinkPLUs() As subMenu
        LinkPLUCount As Integer
    End Type

    Dim myMenu() As Menu
     
    Dim rsSetMLink As New ADODB.Recordset
    Dim fLoad As Boolean, fUpdate As Boolean
    Dim fSearch As Boolean
    Dim fVisibleStock As Boolean
    Dim sFormat As String
    Dim i, j, k As Integer


Private Sub cmdMPlu_Click()
On Error GoTo Handle
    frmSetMPLU.Show vbModal
    Call SetComboPLU
Exit Sub
Handle:
    MsgBox Err.Description & Me.name & ""
End Sub

Private Sub cmdSave_Click()
On Error GoTo Handle
    If fUpdate = True Then
        Call Add_DataUpdate_To_DB
        Picwait.Visible = False
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdSave_Click"
End Sub

Private Sub flex_Scroll()
If flex.Row < flex.Rows - 1 Then flex.Row = flex.Row + 1
    
End Sub

Private Sub Form_Activate()
On Error GoTo errHdl

    Dim ctrl As Control
        
    If UBound(DescArr) = 0 Then
        Initialize
        DescArr = LoadLanguage(LngFile, "#01:015:")
        If cmdClose.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
        Me.Caption = DescArr(1)
        With flex
            .TextMatrix(0, 0) = DescArr(11)
            .TextMatrix(0, 1) = DescArr(12)
        End With
        With flexMPLU
            .TextMatrix(0, 0) = DescArr(11)
            .TextMatrix(0, 1) = DescArr(12)
            .TextMatrix(0, 2) = DescArr(13)
        End With
        For Each ctrl In Me
            DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
        Next ctrl
    End If
    Call Init_SMPLU_Not_Set
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - Form_Activate "
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    With flex
        If Shift = 2 Then 'xac dinh cac fim duoc click: shift,ctrl,alt
            If KeyCode = vbKeyDown Then ' chon keypreview trong from =true
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    If .Row > 26 Then .TopRow = .Row - 25
                End If
                KeyCode = 0
                flex_Click
            ElseIf KeyCode = vbKeyUp Then 'ctrl + keyup
                If .Row > 1 Then
                    .Row = .Row - 1
                    If .Row < .TopRow Then .TopRow = .Row
                End If
                KeyCode = 0
                flex_Click
            End If
        End If
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - Form_KeyDown "
End Sub

Private Sub Form_Load()
On Error GoTo errHdl

    ReDim Preserve DescArr(0)
    sFormat = "#,##0.000"
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsSetMLink = Open_Table(cnData, "SetMLink")
    With probarWait
        .Value = 0
        .Min = 0
        .Max = 1000
    End With
    With Picwait
        .top = (Me.Height - .Height) / 2
        .Left = (Me.Width - .Width) / 2
    End With
    
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - Form_Load "
End Sub

Private Sub Initialize()
On Error GoTo errHdl

    fLoad = False: fUpdate = False
    fVisibleStock = False
    ReDim Preserve myMenu(0)
    With myMenu(0)
        ReDim Preserve .LinkPLUs(0)
        .LinkPLUCount = 0
    End With
    SetDataInFlex
    With txtStockRate(0)
        .Visible = False
        .MaxLength = 7
        .Height = flexMPLU.RowHeight(1)
        .Width = flexMPLU.ColWidth(2) - 20
    End With
'    SetColorFlexGrid flexMPLU, 1, 0, flexMPLU.Cols
    With flex
'        SetColorFlexGrid flex, 1, 1, .Cols
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
        .ScrollTrack = True
    End With
    SetComboPLU
    fLoad = True
    flex_Click

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - Initialize "
End Sub

Private Sub SetComboPLU()
On Error GoTo errHdl

    Dim res As New ADODB.Recordset
    cboSelect.Clear
    Set res = Open_Table(cnData, "SetMPLU")
    If Not res Is Nothing Then
        With res
            If .RecordCount = 0 Then
                cboSelect.AddItem "-------"
                GoTo 1
            End If
            .Sort = "PLUCode ASC"
            .MoveFirst
            Do While Not .EOF
                DoEvents
                cboSelect.AddItem .Fields(0) & "  " & .Fields(1) & Space(20) & .Fields(3)
                cboSelect.ItemData(cboSelect.NewIndex) = .Fields(0)
                .MoveNext
            Loop
        End With
    End If
1:
    CloseRecordset res
    cboSelect.ListIndex = 0

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - SetComboPLU "
End Sub

'           ---------- COMMANDBUTTON -------
Private Sub cmdAdd_Click()
On Error GoTo errHdl
    Dim fFound As Boolean
    With frmPickup
        .GetfPickup = 19
        Set .FormCall = Me
        .Show vbModal
    End With
    fUpdate = True
    cboSelect.ListIndex = 0
    For j = 1 To flex.Rows - 1
        DoEvents
        fFound = False
        For k = 0 To UBound(myMenu)
            DoEvents
            If myMenu(k).PluCode <> "" Then
                If StrComp(myMenu(k).PluCode, flex.TextMatrix(j, 0), 1) = 0 Then _
                    fFound = True
            End If
        Next k
        If Not fFound Then
            ReDim Preserve myMenu(UBound(myMenu) + 1)
            myMenu(UBound(myMenu)).LinkPLUCount = 0
            myMenu(UBound(myMenu)).PluCode = flex.TextMatrix(j, 0)
            ReDim Preserve myMenu(UBound(myMenu)).LinkPLUs(0)
        End If
    Next j

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - cmdAdd_Click "
End Sub

Private Sub cmdDel_Click()
On Error GoTo errHdl

    For i = 1 To UBound(myMenu)
        DoEvents
        If StrComp(myMenu(i).PluCode, flex.TextMatrix(flex.Row, 0), 1) = 0 Then
            myMenu(i).PluCode = ""
            myMenu(i).LinkPLUCount = 0
            ReDim Preserve myMenu(i).LinkPLUs(0)
        End If
    Next i
    With flex
        If .Rows = 2 Then
            For j = 0 To .Cols - 1
                DoEvents
                .TextMatrix(1, j) = ""
            Next j
        Else
            .RemoveItem .Row
'            SetColorFlexGrid flex, .Row, 1, .Cols
        End If
        fUpdate = True
    End With
    flex_Click
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - cmdDel_Click "
End Sub

Private Sub cmdAddPlu_Click()
On Error GoTo errHdl
    With flexMPLU
        If .Rows = 1 Then .Rows = 2
        If cboSelect.ItemData(cboSelect.ListIndex) = 0 Then Exit Sub
        For i = 1 To .Rows - 1
            DoEvents
            If StrComp(.TextMatrix(i, 0), FillZeroForString(cboSelect.ItemData(cboSelect.ListIndex), 6), 1) = 0 Then
                MsgBox DescArr(14)
                Exit Sub
            End If
        Next i
        If .TextMatrix(1, 0) = "" Then
            .Row = 1
        Else
            .Rows = .Rows + 1
            .Row = .Rows - 1
'            SetColorFlexGrid flexMPLU, .Row - 1, 0, .Cols
        End If
        .TextMatrix(.Row, 0) = FillZeroForString(cboSelect.ItemData(cboSelect.ListIndex), 6)
        .TextMatrix(.Row, 1) = FindNamePLU("SetMPLU", FillZeroForString(cboSelect.ItemData(cboSelect.ListIndex), 6), "PluCode", "PluName")
        .TextMatrix(.Row, 2) = "0"
        For j = 0 To UBound(myMenu)
            DoEvents
            If myMenu(j).PluCode <> "" Then
                If StrComp(myMenu(j).PluCode, flex.TextMatrix(flex.Row, 0), 1) = 0 Then
                    myMenu(j).LinkPLUCount = myMenu(j).LinkPLUCount + 1
                    ReDim Preserve myMenu(j).LinkPLUs(myMenu(j).LinkPLUCount)
                    myMenu(j).LinkPLUs(myMenu(j).LinkPLUCount).PluCode = .TextMatrix(.Row, 0)
                    myMenu(j).LinkPLUs(myMenu(j).LinkPLUCount).PluName = .TextMatrix(.Row, 1)
                    myMenu(j).LinkPLUs(myMenu(j).LinkPLUCount).StockRate = .TextMatrix(.Row, 2)
                End If
            End If
        Next j
        .Col = 2
        flexMPLU_DblClick
    End With
    Call ps_SetPluName4MyMenu
    
    fUpdate = True

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - cmdAddPlu_Click "
End Sub

Private Sub cmdClose_Click()
On Error GoTo errHdl
    Dim res
    If Not fUpdate Then GoTo 1
    res = MsgBox("B¹n cã muèn l­u th«ng tin thay ®æi kh«ng?", vbYesNo)
    Select Case res
        Case vbYes: Add_DataUpdate_To_DB
        Case vbNo:  GoTo 1
        Case vbCancel: Exit Sub
    End Select
1:
    CloseRecordset rsSetMLink
    Unload Me

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - cmdClose_Click "
End Sub

Private Sub cmdSearch_Click()
On Error GoTo errHdl

    fSearch = True
    With frmFind
        .GetfSearch = 3
        Set .FormCall = Me
        .Show vbModal
    End With
    fSearch = False

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - cmdSearch_Click "
End Sub

Private Sub cmddelete_Click()
On Error GoTo errHdl

    With flexMPLU
        If .Rows = 2 Then
            For j = 0 To .Cols - 1
                DoEvents
                .TextMatrix(1, j) = ""
            Next j
        Else
            .RemoveItem .Row
'            SetColorFlexGrid flexMPLU, .Row, 0, .Cols
        End If
    End With
    For i = 0 To UBound(myMenu)
        DoEvents
        With myMenu(i)
            If .PluCode <> "" Then
                If StrComp(.PluCode, flex.TextMatrix(flex.Row, 0), 1) = 0 Then
                    If flexMPLU.TextMatrix(1, 0) <> "" Then
                        For k = 1 To flexMPLU.Rows - 1
                            DoEvents
                            .LinkPLUCount = k
                            ReDim Preserve .LinkPLUs(.LinkPLUCount)
                            .LinkPLUs(.LinkPLUCount).PluCode = flexMPLU.TextMatrix(k, 0)
                            .LinkPLUs(.LinkPLUCount).StockRate = RemoveComma(flexMPLU.TextMatrix(k, 2))
                        Next k
                    Else
                        .LinkPLUCount = 0
                        .LinkPLUs(0).PluCode = ""
                        .LinkPLUs(0).StockRate = 0
                    End If
                End If
            End If
        End With
    Next i
    fUpdate = True

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - cmdDelete_Click "
End Sub

Private Sub cmdPickup_Click()
On Error GoTo errHdl
    With frmPickup
        .GetfPickup = 20
        Set .FormCall = Me
        .Show vbModal
    End With
    cmdAddPLU.SetFocus

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - cmdPickup_Click "
End Sub

Private Sub flex_EnterCell()
On Error GoTo errHdl

    If fLoad Then flex_Click

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - flex_EnterCell "
End Sub

Private Sub SetDataInFlex()
On Error GoTo errHdl

    Dim res As New ADODB.Recordset
    Dim irow As Integer
    Dim fFound As Boolean
        
    SetHeaderFlexGrid
    With rsSetMLink
        If .RecordCount = 0 Then Exit Sub
        If UBound(DescArr) = 0 Then
            Picwait.Visible = True
        End If
        irow = 1
        .MoveFirst
        Do While Not .EOF
            DoEvents
            fFound = False
            For k = 0 To UBound(myMenu)
                DoEvents
                If StrComp(myMenu(k).PluCode, .Fields(0), 1) = 0 Then
                    With myMenu(k)
                        .LinkPLUCount = .LinkPLUCount + 1
                        ReDim Preserve .LinkPLUs(.LinkPLUCount)
                        .LinkPLUs(.LinkPLUCount).PluCode = rsSetMLink.Fields(1)
                        .LinkPLUs(.LinkPLUCount).StockRate = rsSetMLink.Fields(2)
                        .LinkPLUs(.LinkPLUCount).PluName = FindNamePLU("SetMPLU", rsSetMLink.Fields("SMPluCode"), "Plucode", "PluName")
                    End With
                    fFound = True
                    Exit For
                End If
            Next k
            If Not fFound Then
                ReDim Preserve myMenu(UBound(myMenu) + 1)
                With myMenu(UBound(myMenu))
                    .PluCode = rsSetMLink.Fields(0)
                    .LinkPLUCount = 1
                    ReDim Preserve .LinkPLUs(1)
                    .LinkPLUs(1).PluCode = rsSetMLink.Fields(1)
                    .LinkPLUs(1).StockRate = rsSetMLink.Fields(2)
                    .LinkPLUs(1).PluName = FindNamePLU("SetMPLU", rsSetMLink.Fields("SMPluCode"), "PluCode", "PluName")
                End With
            End If
            irow = irow + 1
            .MoveNext
            If UBound(DescArr) = 0 Then
                With probarWait
                    If .Value = .Max Then .Value = 200
                    .Value = .Value + 100
                End With
            End If
        Loop
    End With
    
    Dim arrSource() As Variant
    With res
        irow = 1
        .Open "Select distinct(PLUCode) from SetMLink", cnData, adOpenKeyset, adLockOptimistic
        ReDim Preserve arrSource(.RecordCount)
        .MoveFirst
        Do While Not .EOF
        DoEvents
            If UBound(DescArr) = 0 Then
                With probarWait
                    If .Value = .Max Then .Value = 200
                    .Value = .Value + 100
                End With
            End If
            arrSource(irow) = .Fields(0)
            irow = irow + 1
            .MoveNext
        Loop
    End With
    CloseRecordset res
    flex.Rows = UBound(arrSource) + 1
    For i = 1 To UBound(arrSource)
        DoEvents
        If UBound(DescArr) = 0 Then
            With probarWait
                If .Value = .Max Then .Value = 200
                .Value = .Value + 100
            End With
        End If
        flex.TextMatrix(i, 0) = arrSource(i)
        flex.TextMatrix(i, 1) = FindNamePLU("Inventory", CStr(arrSource(i)), "ItemNum", "ItemName")
    Next i
    Picwait.Visible = False

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - SetDataInFlex "
End Sub

Private Sub SetHeaderFlexGrid()
On Error GoTo errHdl

    With flex
        .AllowUserResizing = flexResizeBoth
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .Cols = 2
        .ColWidth(0) = 2000
        .ColWidth(1) = 4500
        .ColAlignment(0) = 4
        .ColAlignment(1) = 2
    End With
    With flexMPLU
        .AllowUserResizing = flexResizeBoth
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionFree
        .Cols = 3
        .ColWidth(0) = 1150
        .ColWidth(1) = 3015
        .ColWidth(2) = 1500
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - SetHeaderFlexGrid "
End Sub

Private Sub flexMPLU_Scroll()
On Error GoTo errHdl

    If fVisibleStock = True Then
        With flexMPLU
            .Col = 2
            txtStockRate(0).top = .top + .CellTop
            txtStockRate(0).Left = .Left + .CellLeft
        End With
    End If

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - flexMPLU_Scroll "
End Sub

Private Sub flexMPLU_DblClick()
On Error GoTo errHdl

    With flexMPLU
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        If .Col = 2 Then
            fVisibleStock = True
            txtStockRate(0).Visible = True
            txtStockRate(0).top = .top + .CellTop
            txtStockRate(0).Left = .Left + .CellLeft - 20
            txtStockRate(0).Width = .ColWidth(2)
            txtStockRate(0).Height = .RowHeight(.Row)
            txtStockRate(0).Text = .TextMatrix(.Row, 2)
            txtStockRate(0).SetFocus
            txtStockRate(0).SelStart = 0
            txtStockRate(0).SelLength = 9999
        Else
            fVisibleStock = False
            txtStockRate(0).Visible = False
        End If
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - flexMPLU_DblClick "
End Sub

Private Sub flexMPLU_Click()
On Error GoTo errHdl

    fVisibleStock = False

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - flexMPLU_Click "
End Sub

Private Sub MineralReport_Click()
  On Error GoTo errHdl
    Dim cmdMaterial As New ADODB.Command
    Dim strSql As String
    strSql = "SELECT Inventory.ItemNum, Inventory.ItemName,Inventory.Std_Price1, SetMLink.SMPLUCode," & _
            " SetMLink.StockRate/1000 as Rate, SetMPLU.PLUName," & _
            " SetMPLU.Cost, SetMPLU.Unit" & _
            " FROM SetMPLU INNER JOIN (Inventory INNER JOIN SetMLink ON " & _
            " Inventory.ItemNum = SetMLink.PLUCode) ON SetMPLU.PLUCode = SetMLink.SMPLUCode"
    
    
    With cmdMaterial
        .ActiveConnection = cnData
        .CommandType = adCmdText
        .CommandText = strSql
        .Execute
    End With
    Dim sReport As New crMaterial
    With sReport
        .Database.AddADOCommand cnData, cmdMaterial
        .PluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .PluName.SetUnboundFieldSource "{ado.ItemName}"
        .Price.SetUnboundFieldSource "{ado.Std_Price1}"
        .SMPluCode.SetUnboundFieldSource "{ado.SMPluCode}"
        .SMPluName.SetUnboundFieldSource "{ado.PLUName}"
        .SMUnit.SetUnboundFieldSource "{ado.Unit}"
        .SMStockRate.SetUnboundFieldSource "{ado.Rate}"
        .Cost.SetUnboundFieldSource "{ado.Cost}"
        .ReportTitle = DescArr(16)
        .lblItemcode.SetText DescArr(17)
        .lblItemName.SetText DescArr(18)
        .lblUnit.SetText DescArr(19)
        .lblStockRate.SetText DescArr(20)
        .lblPrice.SetText DescArr(21)
        .lblTotals.SetText DescArr(22)
        
    End With
    With frmShowReport
        .Report = sReport
        .Show vbModal, Me
    End With
    Set sReport = Nothing
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - Material_Report "
End Sub

Private Sub txtStockRate_DblClick(Index As Integer)
On Error GoTo Handle
    With frmKeyboard.txtInput
        
        .PasswordChar = ""
        .Text = txtStockRate(Index).Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtStockRate(Index).Text = .Let_Text_Input
        End With
        UpdateData
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtPLU_DblClick"
End Sub

'           -------- TEXTBOX -------
Private Sub txtStockRate_GotFocus(Index As Integer)
On Error GoTo errHdl

    txtStockRate(0).SelStart = 0
    txtStockRate(0).SelLength = 9999

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - txtStockRate_GotFocus "
End Sub

Private Sub txtStockRate_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl

    Select Case KeyAscii
        Case 13: txtStockRate(0).Visible = False
        Case 44, 46 ' "."
                For i = 1 To Len(txtStockRate(0).Text)
                    If Mid(txtStockRate(0).Text, i, 1) = "." Or _
                       Mid(txtStockRate(0).Text, i, 1) = "," Then
                        KeyAscii = 0
                        Exit For
                    End If
                Next i
                If Len(RemoveComma(txtStockRate(0).Text)) > 6 Then KeyAscii = 0
        Case 8
        Case 48 To 57
                If Len(RemoveComma(txtStockRate(0).Text)) > 6 Then KeyAscii = 0
        Case Else: KeyAscii = 0
    End Select

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - txtStockRate_KeyPress "
End Sub

Private Sub txtStockRate_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    Dim sTemp As String
                        
    sTemp = Format(txtStockRate(0).Text, sFormat)
    flexMPLU.TextMatrix(flexMPLU.Row, 2) = sTemp
    fUpdate = True
    UpdateData

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - txtStockRate_KeyUp "
End Sub

Private Sub txtStockRate_LostFocus(Index As Integer)
On Error GoTo errHdl

    txtStockRate(Index).Visible = False

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - txtStockRate_LostFocus "
End Sub
'           -------- UPDATEDATA -------
Private Sub UpdateData()
On Error GoTo errHdl

    Dim k As Integer
    Dim j As Integer
    
    For j = 0 To UBound(myMenu)
        DoEvents
        With myMenu(j)
            If .PluCode <> "" Then
                If StrComp(.PluCode, flex.TextMatrix(flex.Row, 0), 1) = 0 Then
                    For k = 0 To .LinkPLUCount
                        DoEvents
                        If .LinkPLUs(k).PluCode <> "" Then
                            If StrComp(.LinkPLUs(k).PluCode, flexMPLU.TextMatrix(flexMPLU.Row, 0), 1) = 0 Then
                                If IsNumeric(flexMPLU.TextMatrix(k, 2)) Then
                                    .LinkPLUs(k).StockRate = RemoveComma(flexMPLU.TextMatrix(k, 2))
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End With
    Next j

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - UpdateData "
End Sub

'           ----- ADD UPDATED DATA TO DATABASE ----
Private Sub Add_DataUpdate_To_DB()
On Error GoTo errHdl

    Dim k As Integer
    Dim i As Integer
    
    '16-02-2006
    Dim cnStockError As New ADODB.Connection
    Dim rsStockError As New ADODB.Recordset
    Dim fCheckStockError As Boolean
    
    fCheckStockError = True
     Picwait.Visible = True
    With rsSetMLink
        cnData.Execute "Delete  from SetMLink"
        For i = 1 To UBound(myMenu)
        DoEvents
            If myMenu(i).PluCode <> "" Then
                
                If myMenu(i).LinkPLUCount > 0 Then
                    For k = 1 To myMenu(i).LinkPLUCount
                        DoEvents
                        .addNew
                        .Fields(0) = myMenu(i).PluCode
                        .Fields(1) = myMenu(i).LinkPLUs(k).PluCode
                        .Fields(2) = myMenu(i).LinkPLUs(k).StockRate
                        .Update
                    Next k
                End If
            End If
            With probarWait
                If .Value = .Max Then .Value = 200
                .Value = .Value + 100
            End With
            Call update_Cost(myMenu(i).PluCode)
        Next i
    End With
    Set rsStockError = Nothing
    Set cnStockError = Nothing

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - Add_DataUpdate_To_DB "
End Sub

Private Function FindNamePLU(sTableName As String, sCode As String, sFieldNameCode As String, sFieldNameName As String) As String
On Error GoTo errHdl

    Dim S1 As String
    Dim res As New ADODB.Recordset
    S1 = ""
    If sCode = "" Then GoTo 1
    Set res = Open_Table(cnData, sTableName)
    If res.RecordCount = 0 Then GoTo 1
    With res
        .MoveFirst
        .Find sFieldNameCode & "='" & sCode & "'"
        If Not .EOF Then S1 = .Fields(sFieldNameName)
    End With
1:  CloseRecordset res
    FindNamePLU = S1

Exit Function
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - FindNamePLU "
End Function


Private Sub flex_Click()
On Error GoTo errHdl

    Dim lngCount, lngK, lngI As Long
    If fSearch Then Exit Sub
    fLoad = False

    'xoa luoi nguyen lieu

    flexMPLU.Rows = 1

    If flex.TextMatrix(1, 0) = "" Then Exit Sub
    For lngI = 1 To UBound(myMenu)
        'so sanh ma hang trong menu khac rong
        If (myMenu(lngI).PluCode & "" = flex.TextMatrix(flex.Row, 0)) Then
            lngCount = myMenu(lngI).LinkPLUCount
            flexMPLU.Rows = lngCount + 1
            DoEvents
            For lngK = 0 To lngCount
                If myMenu(lngI).LinkPLUs(lngK).PluCode <> "" Then
                    flexMPLU.Row = lngK
                    flexMPLU.TextMatrix(flexMPLU.Row, 0) = myMenu(lngI).LinkPLUs(lngK).PluCode
                    flexMPLU.TextMatrix(flexMPLU.Row, 1) = myMenu(lngI).LinkPLUs(lngK).PluName
                    flexMPLU.TextMatrix(flexMPLU.Row, 2) = Format(CDbl(myMenu(lngI).LinkPLUs(lngK).StockRate) / 1000, sFormat)
                End If
            Next lngK
        'het kiem tr ma hang trong menu
        End If
    Next lngI
    
    If flexMPLU.Rows = 1 Then
        flexMPLU.Rows = 2
    End If
    mySetColorFlexGrid flexMPLU, 0, flexMPLU.Cols
    fLoad = True
    'Gan caption cho lblNo, lblLokkup
    lblNo.Caption = flex.TextMatrix(flex.Row, 0)
    lblLookup.Caption = flex.TextMatrix(flex.Row, 1)
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - flex_Click "
End Sub

Private Sub mySetColorFlexGrid(FlexGrid As MSFlexGrid, _
             fCol As Integer, lCol As Integer)
            
On Error GoTo errHdl

    Dim irow As Integer
    Dim iCol As Integer
    Dim myColor As Double
    For irow = 1 To FlexGrid.Rows - 1
        
        If irow Mod 500 = 0 Then Delay 200
        
        'xac dinh mau to
        If irow Mod 2 = 0 Then
            myColor = &H80000018
        Else
            myColor = &H80000013
        End If
        
        FlexGrid.Row = irow
        
        DoEvents
        'to mau cot
        For iCol = fCol To lCol - 1
            FlexGrid.Col = iCol
            FlexGrid.CellBackColor = myColor
        Next iCol
        
    Next irow
    
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - mySetColorFlexGrid "
End Sub

'gan ten PLU cho luoi link
Private Sub ps_SetPluName4MyMenu()
On Error GoTo errHdl

    Dim intC, intJ      As Integer
    Dim sPLUCode      As String
    
    
    For intC = 0 To UBound(myMenu)
        For intJ = 0 To myMenu(intC).LinkPLUCount
            sPLUCode = myMenu(intC).LinkPLUs(intJ).PluCode
            
            myMenu(intC).LinkPLUs(intJ).PluName = FindNamePLU("SetMPLU", sPLUCode, "PluCode", "PluName")
            
        Next intJ
    Next intC
    
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - ps_SetPluName4MyMenu "
End Sub


Public Sub update_Cost(ByVal Item_Code As String)
On Error GoTo Handle
    Dim rsInventory As New ADODB.Recordset
    Dim rsMenuLink As New ADODB.Recordset
    Dim strSql_Cost As String
    strSql_Cost = "SELECT Inventory.ItemNum,  Max(SetMLink.StockRate/1000) AS Rate, Max(SetMPLU.Cost) AS MaxOfCost," & _
                  "Sum(SetMLink.StockRate/1000*SetMPLU.Cost) AS Total_Cost" & _
                  " FROM SetMPLU INNER JOIN (Inventory INNER JOIN SetMLink ON Inventory.ItemNum =" & _
                  " SetMLink.PLUCode) ON SetMPLU.PLUCode = SetMLink.SMPLUCode" & _
                  " GROUP BY Inventory.ItemNum"
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    Set rsInventory = Open_Table(cnData, "Inventory")
    
    Set rsMenuLink = OpenCriticalTable(strSql_Cost, cnData)
    
    With rsInventory
        .Find "ItemNum='" & Item_Code & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            With rsMenuLink
                If .State <> 0 And .RecordCount > 0 Then .MoveFirst
                .Find "ItemNum='" & Item_Code & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    rsInventory.Fields("Minstock") = .Fields("Total_Cost")
                    rsInventory.Update
                End If
            End With
        End If
    End With
    
Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & "update_Cost"
    MsgBox Err.Number & Err.Description & Me.name & " update_Cost"
End Sub

Public Sub Init_SMPLU_Not_Set()
On Error GoTo Handle
    Dim strSql As String
    Dim rsNotSet As New ADODB.Recordset
    
    strSql = "select ItemNum,ItemName,Unit,Std_Price1 from Inventory where ItemNum not in (select PLUCode from SetMLink) and F3<>'01'"
    
    Set rsNotSet = OpenCriticalTable(strSql, cnData)
    
    Set dtgNot_Set.DataSource = rsNotSet
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - Init_SMPLU_Not_Set"

End Sub
