VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmInstockB 
   Caption         =   "NhËp kho nhµ hµng"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInstockB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame fraInstockMaster 
      Caption         =   "Chi tiÕt chøng tõ nhËp kho"
      ForeColor       =   &H00FF0000&
      Height          =   10935
      Left            =   10320
      TabIndex        =   53
      Tag             =   "L13"
      Top             =   120
      Width           =   4815
      Begin MSDataGridLib.DataGrid Grid_Doc 
         Height          =   10455
         Left            =   120
         TabIndex        =   58
         Top             =   360
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   18441
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         HeadLines       =   1
         RowHeight       =   26
         WrapCellPointer =   -1  'True
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
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   10215
      Begin prjTouchScreen.MyButton cmdSave 
         Height          =   615
         Left            =   8640
         TabIndex        =   61
         Top             =   2520
         Width           =   1455
         _extentx        =   2566
         _extenty        =   1085
         btype           =   14
         tx              =   "L­u"
         enab            =   -1  'True
         font            =   "frmInstockB.frx":000C
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   16711680
         fcolo           =   16711680
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmInstockB.frx":0034
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin VB.Frame fraVisible 
         Height          =   3495
         Left            =   9120
         TabIndex        =   56
         Top             =   3600
         Visible         =   0   'False
         Width           =   1455
         Begin prjTouchScreen.MyButton cmdUnlock 
            Height          =   2175
            Left            =   120
            TabIndex        =   57
            Top             =   480
            Width           =   1215
            _extentx        =   2143
            _extenty        =   3836
            btype           =   14
            tx              =   "Më khãa"
            enab            =   -1  'True
            font            =   "frmInstockB.frx":0052
            coltype         =   2
            focusr          =   -1  'True
            bcol            =   14215660
            bcolo           =   14215660
            fcol            =   16711680
            fcolo           =   255
            mcol            =   12632256
            mptr            =   1
            micon           =   "frmInstockB.frx":007A
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   0
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
      End
      Begin VB.Frame fraDocument 
         ForeColor       =   &H00C00000&
         Height          =   3330
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   8415
         Begin VB.ComboBox cboHTTT 
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
            ItemData        =   "frmInstockB.frx":0098
            Left            =   6720
            List            =   "frmInstockB.frx":00A2
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   2520
            Width           =   1455
         End
         Begin prjTouchScreen.MyButton cmdVendor 
            Height          =   375
            Left            =   3600
            TabIndex        =   55
            Top             =   1600
            Width           =   495
            _extentx        =   873
            _extenty        =   661
            btype           =   14
            tx              =   "..."
            enab            =   -1  'True
            font            =   "frmInstockB.frx":00AE
            coltype         =   2
            focusr          =   -1  'True
            bcol            =   14215660
            bcolo           =   14215660
            fcol            =   16711680
            fcolo           =   16711680
            mcol            =   12632256
            mptr            =   1
            micon           =   "frmInstockB.frx":00D6
            umcol           =   -1  'True
            soft            =   0   'False
            picpos          =   0
            ngrey           =   0   'False
            fx              =   0
            hand            =   0   'False
            check           =   0   'False
            value           =   0   'False
         End
         Begin VB.ComboBox cboStock 
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
            Left            =   1680
            TabIndex        =   7
            Top             =   2560
            Width           =   3735
         End
         Begin VB.TextBox txtReason 
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
            Left            =   4080
            TabIndex        =   38
            Top             =   2050
            Width           =   4095
         End
         Begin VB.ComboBox cboReason 
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
            Left            =   1680
            TabIndex        =   6
            Top             =   2050
            Width           =   2415
         End
         Begin VB.TextBox txtDiscount 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   5
            EndProperty
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
            Left            =   7440
            TabIndex        =   4
            Text            =   "0%"
            Top             =   1200
            Width           =   735
         End
         Begin VB.TextBox txtDocNo 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   0
            Top             =   240
            Width           =   3330
         End
         Begin VB.TextBox txtUserName 
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
            Left            =   1680
            TabIndex        =   2
            Top             =   1150
            Width           =   1890
         End
         Begin VB.TextBox txtDeliveryPerson 
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
            Left            =   4680
            MaxLength       =   50
            TabIndex        =   3
            Top             =   1150
            Width           =   2175
         End
         Begin VB.TextBox txtOrgDoc 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   1680
            MaxLength       =   20
            TabIndex        =   1
            Top             =   690
            Width           =   3330
         End
         Begin VB.ComboBox cboSup 
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
            Left            =   1680
            TabIndex        =   5
            Top             =   1600
            Width           =   1935
         End
         Begin VB.TextBox txtVendorName 
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
            Left            =   4080
            TabIndex        =   37
            Top             =   1600
            Width           =   4095
         End
         Begin MSComCtl2.DTPicker dtpDateOut 
            Height          =   435
            Left            =   6240
            TabIndex        =   16
            Top             =   720
            Width           =   1965
            _ExtentX        =   3466
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
            Format          =   62849025
            UpDown          =   -1  'True
            CurrentDate     =   38594
         End
         Begin MSComCtl2.DTPicker dtpDateIn 
            Height          =   435
            Left            =   6240
            TabIndex        =   15
            Top             =   240
            Width           =   1965
            _ExtentX        =   3466
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
            Format          =   62849025
            UpDown          =   -1  'True
            CurrentDate     =   38594
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "H×nh thøc TT"
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
            Height          =   240
            Left            =   5400
            TabIndex        =   60
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label lblStock 
            Alignment       =   1  'Right Justify
            Caption         =   "Kho :"
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
            Height          =   240
            Left            =   120
            TabIndex        =   48
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label lblReason 
            Alignment       =   1  'Right Justify
            Caption         =   "Lý do nhËp :"
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
            Height          =   240
            Left            =   60
            TabIndex        =   47
            Tag             =   "L20"
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label cmdDis 
            Alignment       =   1  'Right Justify
            Caption         =   "CK:"
            Height          =   255
            Left            =   6960
            TabIndex        =   46
            Tag             =   "L19"
            Top             =   1320
            Width           =   375
         End
         Begin VB.Label lblNgayCT 
            Alignment       =   1  'Right Justify
            Caption         =   "Ngµy CT:"
            Height          =   240
            Left            =   5040
            TabIndex        =   45
            Tag             =   "L3"
            Top             =   240
            Width           =   1170
         End
         Begin VB.Label lblDocNo 
            Alignment       =   1  'Right Justify
            Caption         =   "Sè chøng tõ:"
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
            Height          =   315
            Left            =   120
            TabIndex        =   44
            Tag             =   "L2"
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label lblUser 
            Alignment       =   1  'Right Justify
            Caption         =   "Ng­êi nhËp:"
            Height          =   300
            Left            =   90
            TabIndex        =   43
            Tag             =   "L6"
            Top             =   1200
            Width           =   1545
         End
         Begin VB.Label lblDeliverPerson 
            Alignment       =   1  'Right Justify
            Caption         =   "Ng­êi giao:"
            Height          =   240
            Left            =   3600
            TabIndex        =   42
            Tag             =   "L7"
            Top             =   1200
            Width           =   1050
         End
         Begin VB.Label lblOutDate 
            Alignment       =   1  'Right Justify
            Caption         =   "Ngµy §H:"
            Height          =   240
            Left            =   5040
            TabIndex        =   41
            Tag             =   "L5"
            Top             =   720
            Width           =   1170
         End
         Begin VB.Label lblDeliverCompany 
            Alignment       =   1  'Right Justify
            Caption         =   "§¬n vÞ xuÊt:"
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
            Height          =   240
            Left            =   60
            TabIndex        =   40
            Tag             =   "L8"
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label lblOrgDoc 
            Alignment       =   1  'Right Justify
            Caption         =   "Sè chøng tõ gèc:"
            Height          =   240
            Left            =   105
            TabIndex        =   39
            Tag             =   "L4"
            Top             =   720
            Width           =   1530
         End
      End
      Begin prjTouchScreen.MyButton cmdClear 
         Height          =   615
         Left            =   8640
         TabIndex        =   35
         Top             =   3225
         Width           =   1455
         _extentx        =   2566
         _extenty        =   1085
         btype           =   14
         tx              =   "Hñy"
         enab            =   -1
         font            =   "frmInstockB.frx":00F4
         coltype         =   2
         focusr          =   -1
         bcol            =   16777215
         bcolo           =   16777215
         fcol            =   16711680
         fcolo           =   16777215
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmInstockB.frx":011C
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdAddMaster 
         Height          =   495
         Left            =   8640
         TabIndex        =   49
         Tag             =   "L14"
         Top             =   240
         Width           =   1455
         _extentx        =   3836
         _extenty        =   1296
         btype           =   14
         tx              =   "&Thªm"
         enab            =   -1
         font            =   "frmInstockB.frx":013A
         coltype         =   2
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmInstockB.frx":0162
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdDelete 
         Height          =   495
         Left            =   8640
         TabIndex        =   50
         Tag             =   "L16"
         Top             =   1360
         Width           =   1455
         _extentx        =   3836
         _extenty        =   1296
         btype           =   14
         tx              =   "&Xãa"
         enab            =   -1
         font            =   "frmInstockB.frx":0180
         coltype         =   2
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmInstockB.frx":01A8
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdUpdate 
         Height          =   495
         Left            =   8640
         TabIndex        =   8
         Tag             =   "L15"
         Top             =   800
         Width           =   1455
         _extentx        =   3836
         _extenty        =   1296
         btype           =   14
         tx              =   "&CËp nhËt"
         enab            =   -1
         font            =   "frmInstockB.frx":01C6
         coltype         =   2
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmInstockB.frx":01EE
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdLock 
         Height          =   495
         Left            =   8640
         TabIndex        =   51
         Top             =   1930
         Width           =   1455
         _extentx        =   3836
         _extenty        =   1296
         btype           =   14
         tx              =   "&Khãa chøng tõ"
         enab            =   -1
         font            =   "frmInstockB.frx":020C
         coltype         =   2
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmInstockB.frx":0234
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Caption         =   "phiÕu nhËp kho"
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
         Height          =   375
         Left            =   1800
         TabIndex        =   52
         Tag             =   "L1"
         Top             =   120
         Width           =   5535
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   7215
      Left            =   0
      TabIndex        =   17
      Top             =   3840
      Width           =   10215
      Begin VB.TextBox txtPluCode 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   5880
         Width           =   1695
      End
      Begin VB.TextBox txtPluName 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   5880
         Width           =   2895
      End
      Begin VB.TextBox txtQty 
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
         Height          =   495
         Left            =   5640
         TabIndex        =   10
         Top             =   5880
         Width           =   1215
      End
      Begin VB.TextBox txtCost 
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
         Height          =   495
         Left            =   6960
         TabIndex        =   11
         Top             =   5880
         Width           =   1455
      End
      Begin VB.TextBox txtAmt 
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
         Height          =   495
         Left            =   8520
         TabIndex        =   12
         Top             =   5880
         Width           =   1455
      End
      Begin VB.TextBox txtUnit 
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
         Height          =   495
         Left            =   4920
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   5880
         Width           =   615
      End
      Begin VB.Frame Frame2 
         Height          =   5415
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   9975
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
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
            Left            =   7440
            TabIndex        =   19
            Top             =   4850
            Width           =   2415
         End
         Begin MSDataGridLib.DataGrid griPLU 
            Height          =   4095
            Left            =   360
            TabIndex        =   20
            Top             =   480
            Visible         =   0   'False
            Width           =   8055
            _ExtentX        =   14208
            _ExtentY        =   7223
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            HeadLines       =   1
            RowHeight       =   21
            WrapCellPointer =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArialH"
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
         Begin MSDataGridLib.DataGrid Grid_Details 
            Height          =   4575
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   9735
            _ExtentX        =   17171
            _ExtentY        =   8070
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   21
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
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
         Begin VB.Label Label6 
            Caption         =   "Tæng céng:"
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
            Height          =   375
            Left            =   5880
            TabIndex        =   22
            Top             =   4920
            Width           =   1695
         End
      End
      Begin prjTouchScreen.MyButton cmdAddItem 
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Tag             =   "L36"
         Top             =   6480
         Width           =   1455
         _extentx        =   3836
         _extenty        =   1296
         btype           =   14
         tx              =   "&Thªm"
         enab            =   -1
         font            =   "frmInstockB.frx":0252
         coltype         =   2
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmInstockB.frx":027A
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdUpdateItem 
         Height          =   615
         Left            =   1800
         TabIndex        =   13
         Tag             =   "L37"
         Top             =   6480
         Width           =   1455
         _extentx        =   3836
         _extenty        =   1296
         btype           =   14
         tx              =   "&CËp nhËt"
         enab            =   -1
         font            =   "frmInstockB.frx":0298
         coltype         =   2
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmInstockB.frx":02C0
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdExit 
         Height          =   615
         Left            =   8520
         TabIndex        =   25
         Tag             =   "L18"
         Top             =   6480
         Width           =   1455
         _extentx        =   3836
         _extenty        =   1296
         btype           =   14
         tx              =   "Th&o¸t"
         enab            =   -1
         font            =   "frmInstockB.frx":02DE
         coltype         =   2
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmInstockB.frx":0306
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdDeleteItem 
         Height          =   615
         Left            =   3480
         TabIndex        =   26
         Tag             =   "L38"
         Top             =   6480
         Width           =   1455
         _extentx        =   3836
         _extenty        =   1296
         btype           =   14
         tx              =   "&Xãa"
         enab            =   -1
         font            =   "frmInstockB.frx":0324
         coltype         =   2
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmInstockB.frx":034C
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdCancel 
         Height          =   615
         Left            =   5160
         TabIndex        =   27
         Tag             =   "L39"
         Top             =   6480
         Width           =   1455
         _extentx        =   3836
         _extenty        =   1296
         btype           =   14
         tx              =   "Hñy &bá"
         enab            =   -1
         font            =   "frmInstockB.frx":036A
         coltype         =   2
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmInstockB.frx":0392
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdPreview 
         Height          =   615
         Left            =   6840
         TabIndex        =   28
         Tag             =   "L40"
         Top             =   6480
         Width           =   1455
         _extentx        =   3836
         _extenty        =   1296
         btype           =   14
         tx              =   "Xem chøng tõ"
         enab            =   -1
         font            =   "frmInstockB.frx":03B0
         coltype         =   2
         focusr          =   -1
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmInstockB.frx":03D8
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin VB.Label lbltooltip 
         Caption         =   "Press keydown to select Items"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   5640
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "M· hµng:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Tag             =   "L29"
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tªn hµng:"
         Height          =   255
         Left            =   1680
         TabIndex        =   32
         Tag             =   "L30"
         Top             =   5640
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Sè l­îng"
         Height          =   255
         Left            =   5640
         TabIndex        =   31
         Tag             =   "L31"
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "§¬n gi¸"
         Height          =   255
         Left            =   6960
         TabIndex        =   30
         Tag             =   "L33"
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Thµnh tiÒn"
         Height          =   255
         Left            =   8520
         TabIndex        =   29
         Tag             =   "L34"
         Top             =   5640
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmInstockB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DescArr() As String
Dim i As Integer
Dim rsInOut As New ADODB.Recordset
Dim rsDocument As New ADODB.Recordset
Dim rsVendor As New ADODB.Recordset
Dim rsStock_List As New ADODB.Recordset
Dim rsInstockDetail As New ADODB.Recordset
Dim Doc_DateTime As String
Dim iReport As New CRAXDDRT.Report
Dim Stock_ID As String
Dim rsInventory As New ADODB.Recordset
Dim rsPLU As New ADODB.Recordset
Dim isSave As Boolean

Private Sub cboReason_Change()
On Error GoTo Handle
        With rsInOut
            .Find "MaNX='" & Trim(cboReason.Text) & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtReason.Text = .Fields("DienGiai")
            End If
        End With
    Exit Sub
Handle:
        MsgBox Err.Number & Err.Description & _
        Me.name & " cboReason_Change"
End Sub

Private Sub cboReason_Click()
    Call cboReason_Change
End Sub

Private Sub cboReason_GotFocus()
    SendKeys "%{DOWN}"
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboHTTT.SetFocus
    End If
End Sub

Private Sub cboSup_Click()
    Call cboSup_Change
End Sub

Private Sub cboSup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboReason.SetFocus
    End If
End Sub
Private Sub cboHTTT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdUpdate.SetFocus
    End If
End Sub


Private Sub cmdAddItem_Click()
    Call Init_AddNew
End Sub

Private Sub cmdAddMaster_Click()
On Error GoTo errHdl
    'Kiem tra ngay nhap co nam trong pham vi khoa so hay ko
    Dim strInStockDate As String
    Dim AutoDocNumber As Boolean
    Dim wYear, wMonth As String
    wYear = Left(DateDefault, 4)
    
    wMonth = Mid(DateDefault, 5, 2)
    strInStockDate = wYear & Mid(DateDefault, 5, 2) & Right(DateDefault, 2)
    Call sSetText4AddNewDoc
    Call GetStock_cbo
    
    txtDocNo.Text = "NK" & gfCONVERT_DATE_TO_STRING(dtpDateIn.Value) & Hour(time) & Minute(time)
    txtOrgDoc.SetFocus
    Call GetVendorTocboSup
    cmdUpdate.Enabled = True
    cmdAddMaster.Enabled = False
    cmdDelete.Enabled = False
    isSave = False
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - cmdAdd_Click "
End Sub

Private Sub cmdCancel_Click()
Call Clear_Text
End Sub

Private Sub cmddelete_Click()
On Error GoTo Handle
If MsgBox("B¹n thùc sù muèn xãa chøng tõ nhËp kho nµy ?", vbYesNo) = vbYes Then
    With rsDocument
        .Find "Doc_Number='" & Trim(txtDocNo.Text) & "'", , adSearchForward, adBookmarkFirst
            If .Fields("iLocked") = True Then
                MsgBox "Chøng tõ nµy ®· khãa, kh«ng thÓ xãa ®­îc !", vbInformation
            Else
                If Not .EOF Then
                    cnData.Execute "Delete  from Inventory_InB" & Format(Month(dtpDateIn.Value), "00") & Right(Format(Year(dtpDateIn.Value), "00"), 2) & " where Doc_Number='" & Trim(txtDocNo.Text) & "'"
                    .Delete adAffectCurrent
                    .Requery
                End If
                Call sSetGrid_Doc
            End If
    End With
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdDelete_Click"
End Sub

Private Sub cmdDeleteItem_Click()
On Error GoTo Handle
    If MsgBox("B¹n cã muèn xãa Nguyªn liÖu nµy kh«ng?", vbYesNo, "Xãa Nguyªn liÖu") = vbYes Then
        With rsInstockDetail
            .Find "ItemNum='" & Trim(txtPluCode.Text) & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Delete adAffectCurrent
            End If
        End With
        Call GetInstock(Trim(txtDocNo.Text))
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdDelete_Click"
End Sub

Private Sub cmdExit_Click()
    If isSave = True Then
        Unload Me
    Else
        Call cmdSave_Click
        Unload Me
    End If
End Sub
Public Sub InitFlexDoc()
    With Grid_Doc
        .Columns(0).Caption = DescArr(2)
        .Columns(0).Width = 1500
        .Columns(1).Caption = DescArr(3)
        .Columns(1).Width = 1600
        .Columns(2).Caption = DescArr(4)
        .Columns(2).Width = 1600
        .Columns(3).Caption = DescArr(5)
        .Columns(3).Width = 1600
        .Columns(4).Caption = DescArr(8)
        .Columns(4).Width = 2500
        .Columns(5).Caption = DescArr(9)
        .Columns(5).Width = 4500
        .Columns(6).Caption = DescArr(7)
        .Columns(6).Width = 1700
    End With
End Sub

Private Sub cmdLock_Click()
On Error GoTo Handle
If MsgBox("B¹n cã ®ång ý khãa chøng tõ nµy?Chøng tõ nµy kh«ng thÓ nhËp thªm hµng !!!", vbYesNo, "Th«ng b¸o") = vbYes Then
    With rsDocument
        .Fields("iLocked") = True
        .Update
        cmdLock.Enabled = False
    End With
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdLock_Click"
End Sub

Private Sub CmdPreview_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
    Dim Monthtable As String
    
    Monthtable = Format(Month(dtpDateIn.Value), "00") & Right(Format(Year(dtpDateIn.Value), "00"), 2)
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = "SELECT Instock_MasterB.Doc_Number, Instock_MasterB.DateTime, Instock_MasterB.Vendor_Number," & _
        " Instock_MasterB.Org_Doc_Number, Instock_MasterB.Date_Org, Instock_MasterB.Cashier_ID, " & _
        " Instock_MasterB.Delivery_Person, Instock_MasterB.Discount," & _
        " Instock_MasterB.iReason, Inventory_InB" & Monthtable & ".ItemNum, Inventory_InB" & Monthtable & ".Description," & _
        " Inventory_InB" & Monthtable & ".Quantity, Inventory_InB" & Monthtable & ".CostPer, Inventory_InB" & Monthtable & ".Amount" & _
        " FROM Instock_MasterB INNER JOIN Inventory_InB" & Monthtable & " ON Instock_MasterB.Doc_Number" & _
        " = Inventory_InB" & Monthtable & ".Doc_Number" & _
        " where  Instock_MasterB.Doc_Number='" & txtDocNo.Text & "'" & _
        " Order by ItemNum ASC"
    Set crStockIn = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crStockIn
        .Database.AddADOCommand cnData, cmd
        .SoCT.SetUnboundFieldSource "{ado.Doc_Number}"
        .NgayCT.SetUnboundFieldSource "{ado.DateTime}"
        .CTgoc.SetUnboundFieldSource "{ado.Org_Doc_Number}"
        .NgayDH.SetUnboundFieldSource "{ado.Date_Org}"
        .Nguoigiao.SetUnboundFieldSource "{ado.Delivery_Person}"
        .Nguoinhan.SetUnboundFieldSource "{ado.Cashier_ID}"
        .Donvixuat.SetUnboundFieldSource "{ado.Vendor_Number}"
        .lydonhap.SetUnboundFieldSource "{ado.iReason}"
        .PluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .PluName.SetUnboundFieldSource "{ado.Description}"
        .Quantity.SetUnboundFieldSource "{ado.Quantity}"
        .Cost.SetUnboundFieldSource "{ado.CostPer}"
        .Amount.SetUnboundFieldSource "{ado.Amount}"
        .lblTitle.SetText "PhiÕu nhËp kho" 'DescArr(24)
    End With
    Set iReport = crStockIn
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdPreview_Click"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Handle
    If MsgBox("B¹n cã muèn l­u phiÕu nhËp kho nµy kh«ng?", vbYesNo) = vbYes Then
        If cboHTTT.ListIndex = 0 Then
            If MsgBox("B¹n cã muèn t¹o phiÕu chi cho phiÕu nhËp kho nµy kh«ng?", vbYesNo) = vbYes Then
                Call Create_PayOut(cboSup.Text)
            Else
                MsgBox "B¹n ph¶i chän h×nh thøc thanh to¸n CN ®Ó tiÕp tôc l­u phiÕu nhËp kho"
                Exit Sub
            End If
        End If
        With rsDocument
            .Find "Doc_Number='" & txtDocNo.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Fields("Payment_Method") = Trim(cboHTTT.Text)
                .Fields("Totals") = TxtTotal.Text
                .Update
            End If
        End With
    Else
         With rsDocument
            .Find "Doc_Number='" & txtDocNo.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Delete adAffectCurrent
                .Requery
            End If
        End With
        Set Grid_Doc.DataSource = Grid_Doc
    End If
    isSave = True
    Exit Sub
Handle:
    MsgBox Err.Description & " - cmdSave_Click"
End Sub

Private Sub cmdUnlock_Click()
On Error GoTo Handle
    fraVisible.Visible = False
    
    With rsDocument
        .Find "Doc_Number='" & Grid_Doc.Columns(0).Value & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            !iLocked = False
            .Update
            .Requery
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - cmdUnlock_Click"
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo Handle
    If txtDocNo.Text = "" Then
        MsgBox DescArr(22), vbInformation
        Exit Sub
    End If
    If cboSup.Text = "" Then
        MsgBox DescArr(23), vbInformation
        Exit Sub
    End If
    If cboReason.Text = "" Then
        MsgBox DescArr(24), vbInformation
        Exit Sub
    End If
    If cboStock.Text = "" Then
        MsgBox DescArr(26), vbInformation
        Exit Sub
    End If
    If fUpDateMain = False Then Exit Sub
    cmdUpdate.Enabled = False
    cmdAddMaster.Enabled = True
    Grid_Doc.Enabled = True
'    cmdDelete.Enabled = True
    Call sSetGrid_Doc
    If Check_table_In_Out(dtpDateIn.Value) = False Then Call CreateTable_InStockB(Doc_DateTime)
    cmdAddItem_Click
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  cmdUpdate_Click"
End Sub


Private Sub cmdUpdateItem_Click()
On Error GoTo Handle
    If CDbl("0" & txtQty.Text) = 0 Then
        MsgBox DescArr(22), vbInformation
        Exit Sub
    End If
    If fUpdate_In_Detail = False Then Exit Sub
    Call GetInstock(Trim(txtDocNo.Text))
    Call Enab_Disab_Command(False)
    cmdAddItem.SetFocus
    isSave = False
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdUpdate_Click"
End Sub

Private Sub cmdVendor_Click()
    frmSupplier.Show vbModal
    Call GetVendorTocboSup
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
Dim ctrl As Control
'If cmdAddMaster.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))

    Next
    'InitFlexDoc
    If UserLevel <> 1 Then CheckRight
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    DescArr = LoadLanguage(LngFile, "#02:017:")
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    isSave = True
    Set rsInOut = Open_Table(cnData, "InOutType")
    Set rsDocument = OpenCriticalTable("Select * from Instock_MasterB where Substring(Instock_MasterB.DateTime,5,2)='" & Mid(DateDefault, 5, 2) & "' and InOutType='I' ORDER BY Instock_MasterB.DateTime", cnData)
    
    Set rsVendor = Open_Table(cnData, "Vendors")
    Set rsStock_List = Open_Table(cnData, "Stock_List")
    Call sSetGrid_Doc
    If rsDocument.RecordCount = 0 Then
        Call GetVendorTocboSup
        Call GetStock_cbo
        Call InitValuefor_DTPicker
    End If
    Call Lock_Vendors_Text
    Call Grid_Doc_Click
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub
Public Sub GetVendorTocboSup()
    On Error GoTo Handle
        Dim rssup As New ADODB.Recordset
        Set rssup = Open_Table(cnData, "Vendors")
        cboSup.Clear
    Do While Not rssup.EOF
        With cboSup
            .AddItem rssup.Fields("Vendor_Number")
        End With
    rssup.MoveNext
    Loop
    Exit Sub
Handle:
        MsgBox Err.Number & Err.Description & _
        Me.name & " GetVendorTocboSup"
End Sub
Public Sub GetStock_cbo()
    On Error GoTo Handle
        Dim rsStock As New ADODB.Recordset
        Set rsStock = Open_Table(cnData, "Stock_List")
        cboStock.Clear
    Do While Not rsStock.EOF
        With cboStock
            .AddItem rsStock.Fields("Stock_Name")
        End With
    rsStock.MoveNext
    Loop
    Exit Sub
Handle:
        MsgBox Err.Number & Err.Description & _
        Me.name & " GetVendorTocboSup"
End Sub
Private Sub cboSup_Change()
   On Error GoTo Handle
        Dim rssup As New ADODB.Recordset
        Set rssup = Open_Table(cnData, "Vendors")
        With rssup
            .Find "Vendor_Number='" & cboSup.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtVendorName.Text = .Fields("Vendor_Name")
             
            End If
        End With
        
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cboSup_Change"
End Sub
Private Sub cboSup_GotFocus()
    SendKeys "%{DOWN}"
End Sub

Public Sub InitValuefor_DTPicker()
On Error GoTo Handle
    dtpDateIn.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    dtpDateOut = gfCONVERT_STRING_TO_DATE(DateDefault)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " InitValuefor_DTPicker"
End Sub

Public Sub Lock_Vendors_Text()
On Error GoTo Handle
    txtVendorName.Locked = True
  
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Lock_Vendors_Text"
End Sub

Public Sub sSetText4AddNewDoc()
On Error GoTo Handle
    txtOrgDoc.Text = ""
    txtDeliveryPerson.Text = ""
    txtDiscount.Text = 0
    txtVendorName.Text = ""
    txtUserName.Text = userName
    dtpDateIn.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    dtpDateOut.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    cboHTTT.ListIndex = 0
    Grid_Doc.Enabled = False
    Set Grid_Details.DataSource = Nothing
    
    Call GetVendorTocboSup
    Call GetInOutType
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " sSetText4AddNewDoc"
End Sub

Public Sub GetInOutType()
On Error GoTo Handle
        cboReason.Clear
    If rsInOut.RecordCount > 0 Then rsInOut.MoveFirst
    Do While Not rsInOut.EOF
        With cboReason
            .AddItem rsInOut.Fields("MaNX")
        End With
    rsInOut.MoveNext
    Loop
    Exit Sub
Handle:
        MsgBox Err.Number & Err.Description & _
        Me.name & " GetInOutType"
End Sub

Public Sub Lock4Doc(b As Boolean)
On Error GoTo Handle
    cmdAddMaster.Enabled = b
    cmdDelete.Enabled = b
    cmdLock.Enabled = b
    cmdUpdate.Enabled = Not b
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Lock4Doc"
End Sub

Private Function fUpDateMain() As Boolean
On Error GoTo errHdl
    fUpDateMain = False
    
    'Kiem tra co trung so chung tu hay khong??
    If (gfCOUNT_RECORD("select count(*) from Instock_MasterB " & _
        "where Doc_Number='" & Trim(txtDocNo.Text) & "'", cnData) > 0) Then
        MsgBox DescArr(21), vbExclamation
        txtDocNo.SetFocus
        Exit Function
    End If
    Dim strYYYYMMDD As String
    strYYYYMMDD = gfCONVERT_DATE_TO_STRING(dtpDateIn.Value)
    With rsDocument
    If .State <> 0 And .RecordCount > 0 Then .MoveFirst
    .Find "Doc_Number='" & Trim(txtDocNo.Text) & "'", , adSearchForward, adBookmarkFirst
        If rsDocument.EOF Then
            .addNew
       End If
            !DateTime = strYYYYMMDD
            !Store_ID = Store_ID
            !Doc_Number = txtDocNo.Text
            !iReason = Trim(cboReason.Text)
            !cashier_ID = UserID
            !iLocked = False
            !Org_Doc_Number = txtOrgDoc.Text
            !Date_Org = gfCONVERT_DATE_TO_STRING(dtpDateOut.Value)
            !Delivery_Person = txtDeliveryPerson.Text
            !discount = txtDiscount.Text
            !Vendor_Number = Trim(cboSup.Text)
            !Stock_ID = Trim(Right("00" & cboStock.ListIndex + 1, 2))
            !InOutType = "I"
            !Payment_Method = Trim(cboHTTT.Text)
            .Update
            .Requery
'        End If
    End With
    fUpDateMain = True
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - fUpDateMain "
End Function
Private Sub sSetGrid_Doc()
On Error GoTo errHdl
    Dim i As Integer
    Dim rsTemp_Doc As New ADODB.Recordset
    With rsTemp_Doc
            If .State = 0 Then
                .Fields.Append "Doc_Number", adVarWChar, 20
                .Fields.Append "Datetime", adVarWChar, 10
                .Fields.Append "org_Doc_Number", adVarWChar, 20
                .Fields.Append "Date_org", adVarWChar, 10
                .Fields.Append "Vendor_Name", adVarWChar, 100
                .Fields.Append "Address", adVarWChar, 255
                .Fields.Append "Delivery_Person", adVarWChar, 255
                .Open
            End If
            Do While Not rsDocument.EOF
                .addNew
                .Fields("Doc_Number") = rsDocument!Doc_Number
                .Fields("Datetime") = gfCONVERT_STRING_TO_DATE(rsDocument!DateTime)
                .Fields("org_Doc_Number") = rsDocument!Org_Doc_Number
                .Fields("Date_org") = gfCONVERT_STRING_TO_DATE(rsDocument!Date_Org)
                rsVendor.Find "Vendor_Number='" & rsDocument!Vendor_Number & "'", , adSearchForward, adBookmarkFirst
                If Not rsVendor.EOF Then
                    .Fields("Vendor_Name") = rsVendor.Fields("Vendor_Name")
                    .Fields("Address") = rsVendor.Fields("Address_1") & ";" & rsVendor.Fields("Address_2")
                Else
                    .Fields("Vendor_Name") = ""
                    .Fields("Address") = ""
                End If
                .Fields("Delivery_Person") = "" & rsDocument!Delivery_Person
                .Update
            rsDocument.MoveNext
            Loop
    End With
    If rsTemp_Doc.RecordCount > 0 Then
        Set Grid_Doc.DataSource = rsTemp_Doc
        Call InitFlexDoc
    'Call flgDocument_EnterCell
    End If
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - sSetGridInstock "
End Sub

Private Function fCount(ByVal pStrTable As String, _
                ByVal pStrCon As String, _
                ByVal pcnData As ADODB.Connection) As Long
On Error GoTo errHdl
    Dim rsTemp  As ADODB.Recordset
    Dim strSql  As String
    
    fCount = 0
    
    If pStrCon = "" Then
        strSql = "select count(*) from " & pStrTable
    Else
        strSql = "select count(*) from " & pStrTable & _
            " where " & pStrCon
    End If
    
    Set rsTemp = pcnData.Execute(strSql)
    
    If rsTemp.EOF And rsTemp.BOF Then Exit Function
    
    fCount = CLng("0" & rsTemp.Fields(0).Value)
    
Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - Set_Grid_Doc "
End Function

Private Sub Form_LostFocus()
On Error GoTo Handle
    Set rsDocument = OpenCriticalTable("Select * from Instock_MasterB where Substring(Instock_MasterB.DateTime,5,2)='" & Format(Month(dtpDateIn.Value), "00") & "'", cnData)
    Call sSetGrid_Doc
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " dtpDateIn_Change"

End Sub

Private Sub Grid_Details_Click()
 On Error GoTo Handle
    If rsInstockDetail.RecordCount = 0 Then Exit Sub
        With rsInstockDetail
            .Find "ItemNum='" & Grid_Details.Columns(0).Value & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtPluCode = .Fields("ItemNum")
                txtQty = .Fields("Quantity")
                txtCost.Text = Format(.Fields("CostPer"), formatNum)
                txtAmt = Format(.Fields("Amount"), formatNum)
            End If
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Grid_Details_Click"
End Sub

Private Sub Grid_Doc_Change()
Call Grid_Doc_Click
End Sub

Private Sub Grid_Doc_Click()
On Error GoTo Handle
    With rsDocument
    If .State <> 0 And .RecordCount > 0 Then
        .MoveFirst
    Else
        Exit Sub
    End If
    .Find "Doc_Number='" & Grid_Doc.Columns(0).Value & "'", , adSearchForward, adBookmarkFirst
    If Not .EOF Then
        txtDocNo.Text = .Fields("Doc_Number")
        dtpDateIn.Value = gfCONVERT_STRING_TO_DATE(.Fields("DateTime"))
        txtOrgDoc.Text = .Fields("Org_Doc_Number")
        dtpDateOut.Value = gfCONVERT_STRING_TO_DATE(.Fields("Date_Org"))
        txtUserName.Text = userName
        txtDeliveryPerson.Text = .Fields("Delivery_Person")
        txtDiscount.Text = .Fields("Discount")
        cboSup.Text = .Fields("Vendor_Number")
        cboReason.Text = .Fields("iReason")
        If UCase(Trim(.Fields("Payment_Method"))) = "TM" Then
            cboHTTT.ListIndex = 0
         Else
            cboHTTT.ListIndex = 1
         End If
        Stock_ID = .Fields("Stock_ID")
        If .Fields("iLocked") = True Then cmdLock.Enabled = False
        With rsStock_List
            If .RecordCount > 0 Then .MoveFirst
            .Find "ID='" & rsDocument.Fields("Stock_ID") & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                cboStock.Text = .Fields("Stock_Name")
            End If
        End With
        'Lay du lieu len List Details
        If Check_table_In_Out(dtpDateIn.Value) = False Then Call CreateTable_InStockB(dtpDateIn.Value)
        Set rsInstockDetail = OpenCriticalTable("Select * from Inventory_InB" & Format(Month(dtpDateIn.Value), "00") & Right(Format(Year(dtpDateIn.Value), "00"), 2) & " where Doc_Number='" & txtDocNo.Text & "' order by ItemNum ASC", cnData)
        Set rsInventory = Open_Table(cnData, "Inventory")
        If Stock_ID = "01" Then
            Call Get_Inventory
        Else
            Set rsPLU = Nothing
            Set rsPLU = Open_Table(cnData, "SetMPLU")
    
        End If
        
        Call GetInstock(txtDocNo.Text)
    End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Grid_Doc_Click"

End Sub

Private Sub Grid_Doc_DblClick()
On Error GoTo Handle
    If UCase(UserID) = "131112" Then
            fraVisible.Visible = True
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Grid_Doc_DblClick"
End Sub

Private Sub griPLU_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl
    If KeyAscii = 27 Then
        griPLU.Visible = False
        txtPluCode.SetFocus
    ElseIf KeyAscii = 13 Then
        With rsPLU
            If .RecordCount = 0 Then
                griPLU.Visible = False
                txtPluCode.SetFocus
                MsgBox DescArr(20), vbExclamation
                Exit Sub
            End If
            txtPluCode.Text = !PluCode
            txtPluName.Text = !PluName
            txtCost.Text = !Cost
            txtUnit = !Unit
            txtQty.SetFocus
        End With
        griPLU.Visible = False
    ElseIf KeyAscii = 9 Then
        griPLU.Visible = False
        txtPluCode.SetFocus
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - griPLU_KeyPress "
End Sub

Private Sub cmdClear_Click()
 Call Clear_Text
End Sub

Private Sub txtAmt_GotFocus()
On Error GoTo errHdl

    txtAmt.SelStart = 0
    txtAmt.SelLength = Len(txtAmt)
    
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - txtAmt_GotFocus "
End Sub


Private Sub txtAmt_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl
    If KeyAscii = 13 Then
    If cmdUpdateItem.Enabled = True Then
        cmdUpdateItem.SetFocus
    Else
        cmdAddItem.SetFocus
    End If
    ElseIf KeyAscii < &H30 Or KeyAscii > &H39 Then
        If KeyAscii <> &H8 And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") Then KeyAscii = 0
    End If
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - txtAmt_KeyPress "
End Sub

Private Sub txtAmt_LostFocus()
On Error GoTo errHdl
    If CDbl(txtQty) <> 0 Then
        txtCost.Text = Format(CDbl(txtAmt.Text) / CDbl(txtQty.Text), formatNum)
    Else
        txtCost.Text = "0"
    End If
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - txtAmt_LostFocus "
End Sub

Private Sub txtCost_GotFocus()
On Error GoTo errHdl

    txtCost.SelStart = 0
    txtCost.SelLength = Len(txtCost)
    
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - txtCost_GotFocus "
End Sub

Private Sub txtCost_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 13 Then
        txtAmt.SetFocus
    ElseIf KeyAscii < &H30 Or KeyAscii > &H39 Then
        If KeyAscii <> &H8 And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") Then KeyAscii = 0
    End If

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - txtCost_KeyPress "
End Sub

Private Sub txtCost_LostFocus()
On Error GoTo errHdl
    txtAmt.Text = Format(Round(CDbl(txtQty.Text) * CDbl(txtCost.Text), 0), formatNum)
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - txtCost_LostFocus "

End Sub

Private Sub txtDeliveryPerson_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDiscount.SetFocus
    End If
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboSup.SetFocus
    End If
End Sub

Private Sub txtDocNo_Change()
On Error GoTo Handle
    With rsDocument
    If .State <> 0 And .RecordCount > 0 Then
        .MoveFirst
    Else
        Exit Sub
    End If
    .Find "Doc_Number='" & txtDocNo.Text & "'", , adSearchForward, adBookmarkFirst
    If Not .EOF Then
        dtpDateIn.Value = gfCONVERT_STRING_TO_DATE(.Fields("DateTime"))
        txtOrgDoc.Text = .Fields("Org_Doc_Number")
        dtpDateOut.Value = gfCONVERT_STRING_TO_DATE(.Fields("Date_Org"))
        txtUserName.Text = userName
        txtDeliveryPerson.Text = .Fields("Delivery_Person")
        txtDiscount.Text = .Fields("Discount")
        cboSup.Text = .Fields("Vendor_Number")
        cboReason.Text = .Fields("iReason")
        If .Fields("iLocked") = True Then
            cmdLock.Enabled = False
        Else
            cmdLock.Enabled = True
        End If
    End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtDocNo_Change"

End Sub

Private Sub txtOrgDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtDeliveryPerson.SetFocus
    End If
End Sub

Public Sub Clear_Text()
    txtPluCode.Text = ""
    txtPluName.Text = ""
    txtUnit = ""
    txtQty.Text = 1
    txtCost = 0
    txtAmt = 0
    txtPluCode.SetFocus
    cmdAddMaster.Enabled = True
    cmdUpdate.Enabled = False
    Grid_Doc.Enabled = True
End Sub

Public Function Enab_Disab_Command(b As Boolean)
On Error GoTo Handle
    cmdAddItem.Enabled = Not b
    cmdUpdateItem.Enabled = b
    cmdCancel.Enabled = b
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Enab_Disab_Command"
End Function

Public Sub Init_AddNew()
On Error GoTo Handle
    Call Clear_Text
    Enab_Disab_Command (True)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Init_AddNew"
End Sub

Public Sub GetInstock(strDoc_Num As String)
On Error GoTo Handle
Dim rs As New ADODB.Recordset
Dim str As String
    str = "select * from Inventory_InB" & Format(Month(dtpDateIn.Value), "00") & Right(Format(Year(dtpDateIn.Value), "00"), 2) & " where Doc_Number='" & strDoc_Num & "' order by ItemNum ASC"
    Set rs = OpenCriticalTable(str, cnData)
    If Not rs.EOF Then
        Call Set_FlgIn_Detail(rs)
    Else
    Set Grid_Details.DataSource = Nothing
    End If
Call sSumTotal
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " GetInstock"
End Sub

Public Sub Set_FlgIn_Detail(rs As ADODB.Recordset)
    On Error GoTo Handle
    Dim i As Integer
    Dim rsDetails_Temp As New ADODB.Recordset
    With rsDetails_Temp
            If .State = 0 Then
                .Fields.Append "ItemNum", adVarWChar, 20
                .Fields.Append "Description", adVarWChar, 36
                .Fields.Append "Unit", adVarWChar, 20
                .Fields.Append "Quantity", adDouble
                .Fields.Append "CostPer", adDouble
                .Fields.Append "Amount", adDouble
                .Open
            End If
            Do While Not rs.EOF
                .addNew
                .Fields("ItemNum") = rs.Fields("ItemNum")
                .Fields("Description") = rs.Fields("Description")
                'Tim kiem don vi tinh ga'n vao flgDetail
                If rsPLU.State = 1 Then
                    If rsPLU.RecordCount > 0 Then rsPLU.MoveFirst
                End If
                rsPLU.Find "PLUCode='" & rs.Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                If Not rsPLU.EOF Then
                    .Fields("Unit") = rsPLU.Fields("Unit")
                End If
                .Fields("Quantity") = rs.Fields("Quantity")
                .Fields("CostPer") = rs.Fields("CostPer")
                .Fields("Amount") = rs.Fields("Amount")
            rs.MoveNext
            Loop
    End With
    Set Grid_Details.DataSource = rsDetails_Temp
    Call InitFlexDetail
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Set_FlgIn_Detail"
End Sub

Public Sub InitFlexDetail()
On Error GoTo Handle
    With Grid_Details
        .Columns(0).Caption = "M· hµng" 'DescArr(3)
        .Columns(0).Width = 1200
        .Columns(1).Caption = "Tªn hµng" 'DescArr(4)
        .Columns(1).Width = 2600
        .Columns(2).Caption = "§VT" 'DescArr(6)
        .Columns(2).Width = 950
        .Columns(2).Alignment = dbgCenter
        .Columns(3).Caption = "Sè l­îng" 'DescArr(5)
        .Columns(3).Width = 1450
        .Columns(3).Alignment = dbgCenter
        .Columns(4).Caption = "§¬n gi¸" 'DescArr(7)
        .Columns(4).Width = 1500
        .Columns(4).NumberFormat = formatNum
        .Columns(4).Alignment = dbgRight
        .Columns(5).Caption = "Thµnh tiÒn" 'DescArr(8)
        .Columns(5).Width = 1600
        .Columns(5).NumberFormat = formatNum
        .Columns(5).Alignment = dbgRight
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  InitFlexDetail"
End Sub

'*********************************************************
'Chuc nang  :them chi tiet
'Tham so vao:khong
'Tham so ra :khong
'Nguoi tao  :Khac Can 18/02/2008
'Nguoi sua  :
'*********************************************************
Private Function fUpdate_In_Detail() As Boolean
On Error GoTo errHdl
    Dim strSql          As String
    Dim dblQty, dblAmt As Double
    
    If txtPluCode.Text = "" Then Exit Function
    fUpdate_In_Detail = False
    
    strSql = "select count(*) from Inventory_InB" & Format(Month(dtpDateIn.Value), "00") & Right(Format(Year(dtpDateIn.Value), "00"), 2) & " where Doc_Number='" & _
        Trim(txtDocNo.Text) & "' and ItemNum='" & _
        Trim(txtPluCode.Text) & "'"
    
    'kiem tra co detail nay chua?'chua co thi them moi
    If gfCOUNT_RECORD(strSql, cnData) = 0 Then
        With rsInstockDetail
            .addNew
            !Doc_Number = txtDocNo.Text
            !DateTime = Format(Year(dtpDateIn.Value), "0000") & Format(Month(dtpDateIn.Value), "00") & Format(Day(dtpDateIn.Value), "00")
            !ItemNum = txtPluCode.Text
            !Description = txtPluName.Text
            !Store_ID = Store_ID
            !Quantity = CDbl("0" & txtQty.Text)
            !CostPer = CDbl("0" & txtCost.Text)
            !Amount = CDbl("0" & txtAmt.Text)
            .Update
            .Requery
        End With
    Else
        'hoi neu detail co roi thi update k?
        OKCancel = MsgBox(DescArr(21) & vbCrLf & vbCrLf & _
            "Th«ng b¸o", vbYesNoCancel)
            dblQty = CDbl("0" & txtQty.Text)
            dblAmt = CDbl("0" & txtAmt)
        If OKCancel = vbYes Then
            rsInstockDetail.Find "ItemNum='" & Trim(txtPluCode.Text) & "'", , adSearchForward, adBookmarkFirst
            With rsInstockDetail
                !Quantity = !Quantity + dblQty
                !Amount = dblAmt + !Amount
                !CostPer = !Amount / !Quantity
                .Update
            End With
            rsInstockDetail.Find "ItemNum='" & Trim(txtPluCode) & "'", , adSearchForward, adBookmarkFirst
        ElseIf OKCancel = vbNo Then
            rsInstockDetail.Find "ItemNum='" & Trim(txtPluCode) & "'", , adSearchForward, adBookmarkFirst
            With rsInstockDetail
                    !Quantity = dblQty
                    !CostPer = dblAmt / dblQty
                    !Amount = dblAmt
                    .Update
            End With
            rsInstockDetail.Find "ItemNum='" & Trim(txtPluCode) & "'", , adSearchForward, adBookmarkFirst
        Else
            Exit Function
        End If
    End If
    fUpdate_In_Detail = True
    If Stock_ID = "02" Then
        Call Update_Material_Cost(txtPluCode.Text, CDbl("0" & txtCost.Text))
    ElseIf Stock_ID = "01" Then
        Call Update_Inventory_Cost(txtPluCode.Text, CDbl("0" & txtCost.Text))
    
    End If
Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - fUpdate_In_Detail "
End Function

Public Sub Update_Material_Cost(strCode As String, dblValue As Double)
On Error GoTo Handle
Dim rsMaterial As New ADODB.Recordset
Set rsMaterial = Open_Table(cnData, "SetMPlu")
With rsMaterial
    .Find "Plucode='" & strCode & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Cost") = dblValue
            .Update
            .Requery
        End If
End With

Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Public Sub Update_Inventory_Cost(strCode As String, dblValue As Double)
On Error GoTo Handle

With rsInventory
    .Find "ItemNum='" & strCode & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Minstock") = dblValue
            .Update
            .Requery
        End If
End With

Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub


Private Sub txtPluCode_Change()
On Error GoTo Handle
If rsPLU.State = 1 And rsPLU.RecordCount > 0 Then rsPLU.MoveFirst
    With rsPLU
        .Find "PLUCODE='" & Trim(txtPluCode.Text) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtPluName = .Fields("PluName")
            txtUnit = .Fields("Unit")
        End If

    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtPluCode_Change"
End Sub

Private Sub txtPluCode_GotFocus()
    lbltooltip.Visible = True
    lbltooltip.Caption = "Press keydown to select Items..."
End Sub

Private Sub txtPluCode_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl
    If KeyAscii = 13 Then
        If txtPluCode.Text = "" Then
            If cmdCancel.Enabled Then cmdCancel.SetFocus
        Else
            With rsPLU
                If Not .BOF And .RecordCount > 0 Then .MoveFirst
                .Find "PluCode='" & Right("000000" & txtPluCode.Text, .Fields("PluCode").DefinedSize) & "'", , adSearchForward, adBookmarkFirst
                If .EOF Then
                    MsgBox DescArr(19), vbCritical
                    txtPluCode.SelStart = 0
                    txtPluCode.SelLength = 9999
                    Exit Sub
                Else
                    txtPluCode.Text = !PluCode
                    txtPluName.Text = !PluName
                    txtUnit = !Unit
                    txtQty.SetFocus
                End If
            End With
        End If
    End If
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - txtPLUCode_KeyPress "
End Sub

Public Sub Get_Inventory()
On Error GoTo Handle
Set rsPLU = Nothing
    With rsPLU
        If .State = 0 Then
                .Fields.Append "PluCode", adVarWChar, 20
                .Fields.Append "PluName", adVarWChar, 100
                .Fields.Append "Unit", adVarWChar, 10
                .Fields.Append "Price11", adVarWChar, 10
                .Fields.Append "Price12", adVarWChar, 10
                .Fields.Append "Price13", adVarWChar, 10
                .Fields.Append "Price21", adVarWChar, 10
                .Fields.Append "Price22", adVarWChar, 10
                .Fields.Append "Price23", adVarWChar, 10
                .Fields.Append "Price31", adVarWChar, 10
                .Fields.Append "Price32", adVarWChar, 10
                .Fields.Append "Price33", adVarWChar, 10
                .Open
            End If
        If rsInventory.State = 1 And rsInventory.RecordCount > 0 Then rsInventory.MoveFirst
        Do While Not rsInventory.EOF
            If ArrayFlag(rsInventory.Fields("F3"), 8) = 1 Then
                .addNew
                .Fields("PluCode") = rsInventory!ItemNum
                .Fields("PluName") = rsInventory!ItemName
                .Fields("Unit") = rsInventory!Unit
                .Fields("Price11") = rsInventory!Std_Price1
                .Fields("Price12") = rsInventory!Std_Price2
                .Fields("Price13") = rsInventory!Std_Price3
                .Fields("Price21") = rsInventory!HH_Price1
                .Fields("Price22") = rsInventory!HH_Price2
                .Fields("Price23") = rsInventory!HH_Price3
                .Fields("Price31") = rsInventory!EV_Price1
                .Fields("Price32") = rsInventory!EV_Price2
                .Fields("Price33") = rsInventory!EV_Price3
                .Update
            End If
        rsInventory.MoveNext
        Loop
    End With
Exit Sub
Handle:
 MsgBox Err.Number & Err.Description & Me.name & "Get_Inventory"
End Sub

Private Sub txtPluCode_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl
    If KeyCode = vbKeyDown Then
        If Stock_ID = "02" Then
            With rsPLU
                If .State = adStateOpen Then .Close
                If InStr(1, txtPluCode.Text, "*", vbTextCompare) > 0 Then
                    .Open "SELECT  PluCode, PluName,Cost, Unit FROM SetMPLU WHERE PluCode LIKE '" & fill_search(txtPluCode.Text) & "%' OR PluName LIKE '" & _
                    fill_search(Trim(txtPluCode.Text)) & "%'  ORDER BY PluCode asc"
                Else
                    .Open "SELECT  PluCode, PluName, Cost, Unit FROM SetMPLU WHERE PluCode LIKE '" & fill_search(txtPluCode.Text) & "%' OR  PluName LIKE '" & fill_search(txtPluCode.Text) & "%' AND lTRIM(PluName)<> '' ORDER BY PluCode ASC"
                End If
            End With
        Else
            Set rsPLU = OpenCriticalTable("Select ItemNum,ItemName,Unit from Inventory", cnData)
            With rsPLU
                If .State = adStateOpen Then .Close
                If InStr(1, txtPluCode.Text, "*", vbTextCompare) > 0 Then
                    .Open "SELECT  ItemNum as PluCode, ItemName as PluName,Std_Price1 as Cost, Unit FROM Inventory WHERE ItemNum, LIKE '" & fill_search(txtPluCode.Text) & "%' OR ItemName LIKE '" & _
                    fill_search(txtPluCode.Text) & "%'  ORDER BY ItemNum asc"
                Else
                    .Open "SELECT  ItemNum as Plucode, ItemName as PluName,Std_Price1 as Cost, Unit FROM Inventory WHERE ItemNum LIKE '" & fill_search(Trim(txtPluCode.Text)) & "' OR ItemName LIKE '" & fill_search(txtPluCode.Text) & "%' AND LTRIM(ItemName)<> '' ORDER BY ItemNum ASC"
                End If
            End With
        End If
        With griPLU
            Set .DataSource = rsPLU
            .Columns(0).Caption = "M· hµng"
            .Columns(0).Width = 1500
            .Columns(1).Caption = "Tªn hµng"
            .Columns(1).Width = 3500
            .Columns(1).Alignment = dbgLeft
            .Columns(2).Caption = "Gi¸"
            .Columns(2).Alignment = dbgRight
            .Columns(2).Width = 1000
            .Columns(3).Caption = "§VT"
            .Columns(3).Alignment = dbgRight
            .Columns(3).Width = 1000
            .Visible = True
            .SetFocus
            .top = Grid_Details.top + 100
            .Left = Grid_Details.Left + 100
        End With
    End If
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - txtPluCode_KeyUp "
End Sub

Private Sub txtPluCode_LostFocus()
    lbltooltip.Visible = False
End Sub

Private Sub txtQty_GotFocus()
On Error GoTo errHdl
    txtQty.SelStart = 0
    txtQty.SelLength = 9999
    
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - txtQty_GotFocus "
End Sub


Public Function Check_table_In_Out(Date_Check As String) As Boolean
On Error GoTo Handle
    Dim cat As New ADOX.Catalog
    Check_table_In_Out = False
    cat.ActiveConnection = myProvider
    
    For i = 1 To cat.Tables.count - 1
        If cat.Tables(i).name = "Inventory_InB" & Format(Month(Date_Check), "00") & Right(Format(Year(Date_Check), "00"), 2) Then
            Check_table_In_Out = True
        End If
    Next

Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Check_table_In_Out"
    Check_table_In_Out = False
End Function

Private Sub txtQty_KeyPress(KeyAscii As Integer)
   On Error GoTo errHdl

    If KeyAscii = 13 Then
        txtCost.SetFocus
    ElseIf KeyAscii < &H30 Or KeyAscii > &H39 Then
        If KeyAscii <> &H8 And KeyAscii <> Asc(".") And KeyAscii <> Asc(",") Then KeyAscii = 0
    End If

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - txtQty_KeyPress "
End Sub
Private Sub sSumTotal()
On Error GoTo errHdl
    Dim rsSumDoc As New ADODB.Recordset
    Set rsSumDoc = OpenCriticalTable("Select sum(Amount) as SumDoc from Inventory_InB" & Format(Month(dtpDateIn.Value), "00") & Right(Format(Year(dtpDateIn.Value), "00"), 2) & " where Doc_number='" & txtDocNo.Text & "'", cnData)
        
        If Not IsNull(rsSumDoc!SumDoc) Then
            TxtTotal.Text = Format(Round(rsSumDoc!SumDoc, 0), formatNum)
        Else
            TxtTotal.Text = 0
        End If
        
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - sSumTotal "
        
End Sub



Private Sub txtQty_LostFocus()
On Error GoTo Handle
    txtAmt.Text = Format(Round(CDbl(txtQty.Text) * CDbl(txtCost.Text), 0), formatNum)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & ""
End Sub
Public Sub CheckRight()
On Error GoTo Handle
    Dim res As New ADODB.Recordset
        
        Set res = LoadPasswordData
        With MyRight
            res.MoveFirst
            Do While Not res.EOF
                If StrComp(res.Fields("UserName"), userName, 1) = 0 Then
                    .FullRight = res.Fields("UserRight")
                    .kho = RightDeCode(Mid(.FullRight, 449, 64))
                    Exit Do
                End If
                res.MoveNext
            Loop
            If Mid(.kho, 4, 1) = 0 Then
                  cmdAddMaster.Enabled = False
            Else: cmdAddMaster.Enabled = True
            End If
            If Mid(.kho, 5, 1) = 0 Then
                  cmdDelete.Enabled = False
            Else: cmdDelete.Enabled = True
            End If
            
            If Mid(.kho, 6, 1) = 0 Then
                  cmdLock.Enabled = False
            Else: cmdLock.Enabled = True
            End If
            If Mid(.kho, 7, 1) = 0 Then
                  cmdAddItem.Enabled = False
            Else: cmdAddItem.Enabled = True
            End If
            If Mid(.kho, 8, 1) = 0 Then
                  cmdDeleteItem.Enabled = False
            Else: cmdDeleteItem.Enabled = True
            End If
        End With

    CloseRecordset res
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " CheckRight"
End Sub

Public Sub Create_PayOut(CustomerID As String)
On Error GoTo Handle
Dim rsPhieuChi As New ADODB.Recordset
Set rsPhieuChi = Open_Table(cnData, "Payouts")
    With rsPhieuChi
            .Find "ID='" & txtDocNo.Text & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("ID") = txtDocNo.Text
                .Fields("Store_ID") = Store_ID
                .Fields("Cashier_ID") = UserID
                .Fields("DateTime") = gfCONVERT_DATE_TO_STRING(Date)
                .Fields("Expense_ID") = "MH"
                .Fields("Vendor_Number") = CustomerID
                .Fields("Recieve_Name") = txtDeliveryPerson.Text
                .Fields("Division") = ""
                .Fields("Payment_Method") = cboHTTT.Text
                .Fields("Amount") = TxtTotal.Text
                .Fields("Description") = "Chi tiÒn phiÕu nhËp kho sè:" & txtDocNo.Text
                .Update
            Else
                .Fields("Amount") = TxtTotal.Text
                .Update
            End If
        End With
        
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  - Create_PayOut"

End Sub
