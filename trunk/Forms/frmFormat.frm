VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmFormat 
   Caption         =   "§Þnh d¹ng"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   5625
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   970
      TabCaption(0)   =   "§Þnh d¹ng sè"
      TabPicture(0)   =   "frmFormat.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Setup m¸y in"
      TabPicture(1)   =   "frmFormat.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame5"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Customer Display Port Setting"
      TabPicture(2)   =   "frmFormat.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "S¾p xÕp"
      TabPicture(3)   =   "frmFormat.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame5 
         Caption         =   "Lo¹i m¸y in"
         Height          =   2415
         Left            =   240
         TabIndex        =   47
         Top             =   2880
         Width           =   8175
         Begin VB.Frame Frame7 
            Caption         =   "In Order"
            Height          =   1935
            Left            =   4200
            TabIndex        =   49
            Top             =   360
            Width           =   3855
            Begin VB.OptionButton optOrder 
               Caption         =   "Khæ giÊy 58mm"
               Height          =   495
               Index           =   1
               Left            =   240
               TabIndex        =   54
               Top             =   960
               Width           =   3375
            End
            Begin VB.OptionButton optOrder 
               Caption         =   "Khæ giÊy 80mm"
               Height          =   495
               Index           =   0
               Left            =   240
               TabIndex        =   53
               Top             =   360
               Width           =   3375
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "In Bill"
            Height          =   1935
            Left            =   120
            TabIndex        =   48
            Top             =   360
            Width           =   3975
            Begin VB.OptionButton optReceipt 
               Caption         =   "Khæ giÊy 75"
               Height          =   375
               Index           =   3
               Left            =   240
               TabIndex        =   55
               Top             =   960
               Width           =   3615
            End
            Begin VB.OptionButton optReceipt 
               Caption         =   "Khæ giÊy A5"
               Height          =   375
               Index           =   2
               Left            =   240
               TabIndex        =   52
               Top             =   1320
               Width           =   3615
            End
            Begin VB.OptionButton optReceipt 
               Caption         =   "Khæ giÊy 58mm"
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   51
               Top             =   600
               Width           =   3615
            End
            Begin VB.OptionButton optReceipt 
               Caption         =   "Khæ giÊy 80mm"
               Height          =   495
               Index           =   0
               Left            =   240
               TabIndex        =   50
               Top             =   240
               Width           =   3615
            End
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Canh lÒ giÊy in"
         Height          =   1935
         Left            =   240
         TabIndex        =   38
         Top             =   840
         Width           =   8175
         Begin VB.TextBox txtTop 
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
            Height          =   495
            Left            =   1080
            TabIndex        =   42
            Text            =   "Top"
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtBottom 
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
            Height          =   495
            Left            =   1080
            TabIndex        =   41
            Text            =   "Bottom"
            Top             =   1080
            Width           =   2415
         End
         Begin VB.TextBox txtLeft 
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
            Height          =   495
            Left            =   5280
            TabIndex        =   40
            Text            =   "Left"
            Top             =   480
            Width           =   2415
         End
         Begin VB.TextBox txtRight 
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
            Height          =   495
            Left            =   5280
            TabIndex        =   39
            Text            =   "Right"
            Top             =   1080
            Width           =   2415
         End
         Begin VB.Label lblTop 
            Alignment       =   1  'Right Justify
            Caption         =   "Trªn:"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   480
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Tr¸i:"
            Height          =   255
            Left            =   4440
            TabIndex        =   45
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "D­íi:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Ph¶i:"
            Height          =   255
            Left            =   4440
            TabIndex        =   43
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "S¾p xÕp theo thø tù ­u tiªn tõ 1 ®Õn ..."
         Height          =   4575
         Left            =   -74880
         TabIndex        =   33
         Top             =   720
         Width           =   8295
         Begin VB.ComboBox cbo2 
            BeginProperty Font 
               Name            =   ".VnArial NarrowH"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            ItemData        =   "frmFormat.frx":0070
            Left            =   2400
            List            =   "frmFormat.frx":007D
            TabIndex        =   37
            Top             =   1440
            Width           =   3975
         End
         Begin VB.ComboBox cbo1 
            BeginProperty Font 
               Name            =   ".VnArial NarrowH"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            ItemData        =   "frmFormat.frx":009F
            Left            =   2400
            List            =   "frmFormat.frx":00AC
            TabIndex        =   35
            Top             =   600
            Width           =   3975
         End
         Begin VB.Label Label11 
            Caption         =   "¦u tiªn 2:"
            Height          =   495
            Left            =   1080
            TabIndex        =   36
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "¦u tiªn 1:"
            Height          =   495
            Left            =   1080
            TabIndex        =   34
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Height          =   4455
         Left            =   -74880
         TabIndex        =   22
         Top             =   720
         Width           =   8175
         Begin VB.ComboBox cboComport 
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
            ItemData        =   "frmFormat.frx":00CE
            Left            =   2760
            List            =   "frmFormat.frx":00F0
            TabIndex        =   27
            Text            =   "Combo1"
            Top             =   600
            Width           =   1815
         End
         Begin VB.ComboBox cboBaud 
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
            ItemData        =   "frmFormat.frx":0131
            Left            =   2760
            List            =   "frmFormat.frx":014A
            TabIndex        =   26
            Text            =   "Combo2"
            Top             =   1320
            Width           =   2295
         End
         Begin VB.TextBox txtDatabits 
            Height          =   495
            Left            =   2760
            Locked          =   -1  'True
            TabIndex        =   25
            Text            =   "8"
            Top             =   2040
            Width           =   2295
         End
         Begin VB.ComboBox cboParity 
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
            ItemData        =   "frmFormat.frx":017C
            Left            =   2760
            List            =   "frmFormat.frx":0189
            TabIndex        =   24
            Text            =   "Combo3"
            Top             =   2760
            Width           =   2295
         End
         Begin VB.TextBox txtStopbit 
            Height          =   495
            Left            =   2760
            TabIndex        =   23
            Text            =   "1"
            Top             =   3360
            Width           =   2295
         End
         Begin VB.Label Label5 
            Caption         =   "Com Port:"
            Height          =   375
            Left            =   960
            TabIndex        =   32
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label6 
            Caption         =   "Baud Rate"
            Height          =   495
            Left            =   960
            TabIndex        =   31
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Data bits"
            Height          =   495
            Left            =   960
            TabIndex        =   30
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label8 
            Caption         =   "Parity"
            Height          =   495
            Left            =   960
            TabIndex        =   29
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label9 
            Caption         =   "Stop bit"
            Height          =   495
            Left            =   960
            TabIndex        =   28
            Top             =   3360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   4
         Top             =   600
         Width           =   8175
         Begin VB.TextBox txtSymbol 
            Height          =   495
            Left            =   4320
            TabIndex        =   10
            Top             =   4080
            Width           =   2175
         End
         Begin VB.TextBox txtQtySepa 
            Height          =   405
            Left            =   4320
            MaxLength       =   1
            TabIndex        =   9
            Top             =   600
            Width           =   780
         End
         Begin VB.TextBox txtQtyDigits 
            Height          =   405
            Left            =   4320
            MaxLength       =   1
            TabIndex        =   8
            Top             =   2385
            Width           =   1020
         End
         Begin VB.TextBox txtAmtSepa 
            Height          =   405
            Left            =   4320
            MaxLength       =   1
            TabIndex        =   7
            Top             =   1185
            Width           =   1020
         End
         Begin VB.TextBox txtGroupDigits 
            Height          =   405
            Left            =   4320
            MaxLength       =   1
            TabIndex        =   6
            Top             =   1785
            Width           =   1020
         End
         Begin VB.TextBox txtAmtDigits 
            Height          =   405
            Left            =   4320
            MaxLength       =   1
            TabIndex        =   5
            Top             =   2985
            Width           =   1020
         End
         Begin MSComCtl2.UpDown udGroupDigits 
            Height          =   375
            Left            =   5400
            TabIndex        =   11
            Top             =   1785
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udQtyDigits 
            Height          =   375
            Left            =   5400
            TabIndex        =   12
            Top             =   2385
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown udAmtDigits 
            Height          =   375
            Left            =   5400
            TabIndex        =   13
            Top             =   2985
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Currency Symbol:"
            Height          =   480
            Left            =   120
            TabIndex        =   21
            Top             =   4080
            Width           =   4185
         End
         Begin MSForms.ComboBox cboFormat 
            Height          =   495
            Left            =   4320
            TabIndex        =   20
            Top             =   3480
            Width           =   2175
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "3836;873"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   ".VnArial"
            FontHeight      =   225
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label lblFormat 
            Alignment       =   1  'Right Justify
            Caption         =   "§Þnh d¹ng:"
            Height          =   480
            Left            =   120
            TabIndex        =   19
            Top             =   3600
            Width           =   4185
         End
         Begin VB.Label lblQtySepa 
            Alignment       =   1  'Right Justify
            Caption         =   "DÊu ph©n c¸ch thËp ph©n cho sè l­îng:"
            Height          =   480
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   4185
         End
         Begin VB.Label lblQtyDigits 
            Alignment       =   1  'Right Justify
            Caption         =   "Sè ch÷ sè hµng thËp ph©n sè l­îng:"
            Height          =   480
            Left            =   120
            TabIndex        =   17
            Top             =   2385
            Width           =   4185
         End
         Begin VB.Label lblAmtSepa 
            Alignment       =   1  'Right Justify
            Caption         =   "DÊu ph©n c¸ch hµng ngh×n:"
            Height          =   480
            Left            =   120
            TabIndex        =   16
            Top             =   1185
            Width           =   4185
         End
         Begin VB.Label lblGroupDigits 
            Alignment       =   1  'Right Justify
            Caption         =   "Sè ch÷ sè cho mçi nhãm sè:"
            Height          =   480
            Left            =   120
            TabIndex        =   15
            Top             =   1785
            Width           =   4185
         End
         Begin VB.Label lblAmtDigits 
            Alignment       =   1  'Right Justify
            Caption         =   "Sè ch÷ sè hµng thËp ph©n gi¸ trÞ:"
            Height          =   480
            Left            =   120
            TabIndex        =   14
            Top             =   2985
            Width           =   4185
         End
      End
   End
   Begin prjTouchScreen.MyButton cmdHelp 
      Height          =   1095
      Left            =   8760
      TabIndex        =   2
      Top             =   3840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      BTYPE           =   14
      TX              =   "&Gióp ®ì"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFormat.frx":019E
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
      Height          =   1095
      Left            =   8760
      TabIndex        =   1
      Top             =   2520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      BTYPE           =   14
      TX              =   "&Hñy bá"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFormat.frx":01BA
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
      Left            =   8760
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1931
      BTYPE           =   14
      TX              =   "&§ång ý"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFormat.frx":01D6
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
Attribute VB_Name = "frmFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ComportID, Baudrate, Databit, Stopbit As Integer
Dim Parity, setting, Comport As String
Dim str1, str2 As String


Private Sub cboBaud_Change()
    Select Case cboBaud.ListIndex
        Case 1
            Baudrate = 2400
        Case 2
           Baudrate = 4800
        Case 3
           Baudrate = 9600
        Case 4
           Baudrate = 14400
        Case 5
           Baudrate = 19200
        Case 6
           Baudrate = 38400
        Case 7
           Baudrate = 57600
    End Select
End Sub

Private Sub cboBaud_Click()
    Call cboBaud_Change
End Sub

Private Sub cboComport_Change()
    Select Case cboComport.ListIndex
        Case 1
            Comport = "COM1"
            ComportID = 1
        Case 2
            Comport = "COM2"
            ComportID = 2
        Case 3
            Comport = "COM3"
            ComportID = 3
        Case 4
            Comport = "COM4"
            ComportID = 4
        Case 5
            Comport = "COM5"
            ComportID = 5
        Case 6
            Comport = "COM6"
            ComportID = 6
        Case 7
            Comport = "COM7"
            ComportID = 7
        Case 8
            Comport = "COM8"
            ComportID = 8
        Case 9
            Comport = "COM9"
            ComportID = 9
        Case 10
            Comport = "COM10"
            ComportID = 10
    End Select
End Sub

Private Sub cboComport_Click()
    ComportID = cboComport.ListIndex + 1
End Sub

Private Sub cboParity_Change()
    Select Case cboParity.ListIndex
        Case -1
            Parity = "N"
        Case 0
            Parity = "O"
        Case 1
            Parity = "E"
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    DigitGroupMark = txtAmtSepa.Text
    DecimalMark = txtQtySepa.Text
    DigitsGroup = CInt(txtGroupDigits.Text)
    DecimalQtyNumber = CInt(txtQtyDigits.Text)
    DecimalAmtNumber = CInt(txtAmtDigits.Text)
    formatNum = cboFormat.Text
    TopAlign = txtTop.Text
    BottomAlign = txtBottom.Text
    LeftAlign = txtLeft.Text
    RightAlign = txtRight.Text
    
    setting = Baudrate & "," & Parity & "," & Databit & "," & Stopbit
    'May in bill
    If optReceipt(0).Value = True Then
        ReceiptType = "80"
    ElseIf optReceipt(1).Value = True Then
        ReceiptType = "58"
    ElseIf optReceipt(3).Value = True Then
        ReceiptType = "75"
    Else
        ReceiptType = "A5"
    End If
    'May in order
     If optOrder(0).Value = True Then
        OrderType = "80"
   Else
        OrderType = "58"
    End If
    
    
    SaveSettingStr "NUMBER", "Digit Group Symbol", DigitGroupMark, myIniFile
    SaveSettingStr "NUMBER", "Decimal Symbol", DecimalMark, myIniFile
    SaveSettingStr "NUMBER", "Digit Group", CStr(DigitsGroup), myIniFile
    SaveSettingStr "NUMBER", "Quantity Decimal", CStr(DecimalQtyNumber), myIniFile
    SaveSettingStr "NUMBER", "Amount Decimal", CStr(DecimalAmtNumber), myIniFile
    SaveSettingStr "NUMBER", "FormatNum", formatNum, myIniFile
    SaveSettingStr "NUMBER", "CurrencySymbol", CurrencySymbol, myIniFile
    'Luu Com Setting to .ini file
    SaveSettingStr "Properties", "ComPort", cboComport.Text, myIniFile
    SaveSettingStr "Properties", "ComPortNumber", ComportID, myIniFile
    SaveSettingStr "Properties", "Setting", setting, myIniFile
    SaveSettingStr "Properties", "Baudrate", Baudrate, myIniFile
    SaveSettingStr "Properties", "Data bits", Databit, myIniFile
    SaveSettingStr "Properties", "Parity", Parity, myIniFile
    SaveSettingStr "Properties", "Stop bit", Stopbit, myIniFile
    
    'Canh le
    SaveSettingStr "ALIGN", "Top", TopAlign, myIniFile
    SaveSettingStr "ALIGN", "Bottom", BottomAlign, myIniFile
    SaveSettingStr "ALIGN", "Left", LeftAlign, myIniFile
    SaveSettingStr "ALIGN", "Right", RightAlign, myIniFile
    'Save lo¹i m¸y in
     SaveSettingStr "PRINTER", "Receipt_Type", ReceiptType, myIniFile
    SaveSettingStr "PRINTER", "Order_Type", OrderType, myIniFile
      
    'loc dieu kien
    Select Case cbo1.ListIndex
        Case 0
            str1 = "Dept_ID"
        Case 1
            str1 = "ItemNum"
        Case 2
            str1 = "ItemName"
    End Select
    
    Select Case cbo2.ListIndex
        Case 0
            str2 = "ItemNum"
        Case 1
            str2 = "ItemName"
        Case 2
            str2 = "Dept_ID"
    End Select
    
    SaveSettingStr "SORT", "Sort_by", str1 & "," & str2, myIniFile
    
    Unload Me
End Sub



Private Sub Form_Load()
    Dim DescArr() As String
    
    DescArr = LoadLanguage(LngFile, "#03:006:")
    If cmdOK.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(9)
    cmdHelp.Caption = DescArr(1)
    cmdOK.Caption = DescArr(2)
    cmdCancel.Caption = DescArr(3)
    
    Call AddValuetoformat
    ' Get Setting comport
    ComportID = GetSettingStr("Properties", "ComPortNumber", 1, myIniFile)
    Comport = GetSettingStr("Properties", "ComPort", "COM1", myIniFile)
    Baudrate = GetSettingStr("Properties", "BaudRate", 2400, myIniFile)
    setting = GetSettingStr("Properties", "Setting", "", myIniFile)
    Databit = GetSettingStr("Properties", "Data s", 8, myIniFile)
    Stopbit = GetSettingStr("Properties", "Stop bit", 1, myIniFile)
    Parity = GetSettingStr("Properties", "Parity", "N", myIniFile)
    'LÊy lo¹i m¸y in
    ReceiptType = GetSettingStr("PRINTER", "Receipt_Type", 80, myIniFile)
    OrderType = GetSettingStr("PRINTER", "Order_Type", 80, myIniFile)
    'Sap xep
    Sort_By = GetSettingStr("SORT", "Sort_by", "Dept_ID,ItemNum", myIniFile)
    
    lblQtySepa.Caption = DescArr(4)
    lblAmtSepa.Caption = DescArr(5)
    lblGroupDigits.Caption = DescArr(6)
    lblQtyDigits.Caption = DescArr(7)
    lblAmtDigits.Caption = DescArr(8)
    lblFormat.Caption = DescArr(9)
    
    txtAmtSepa.Text = DigitGroupMark
    txtQtySepa.Text = DecimalMark
    txtGroupDigits.Text = DigitsGroup
    txtQtyDigits.Text = DecimalQtyNumber
    txtAmtDigits.Text = DecimalAmtNumber
    txtSymbol.Text = CurrencySymbol
    cboFormat.Text = formatNum
    'Canh le
    txtTop.Text = TopAlign
    txtBottom.Text = BottomAlign
    txtLeft.Text = LeftAlign
    txtRight.Text = RightAlign
    ' Lay loai may in
    If ReceiptType = "80" Then
        optReceipt(0).Value = True
    ElseIf ReceiptType = "58" Then
        optReceipt(1).Value = True
    ElseIf ReceiptType = "75" Then
        optReceipt(3).Value = True
    Else
        optReceipt(2).Value = True
    End If
    '''In order
    If OrderType = "80" Then
        optOrder(0).Value = True
    Else
        optOrder(1).Value = True
    End If
    
    'Load Communication
    cboComport.ListIndex = ComportID - 1
    cboBaud.Text = Baudrate
    txtDatabits.Text = 8
    cboParity.Text = "None"
    txtStopbit.Text = 1
    
    str1 = Mid(Sort_By, 1, InStr(Sort_By, ",") - 1)
    str2 = Right(Sort_By, Len(Sort_By) - InStr(Sort_By, ","))
    
    cbo1.Text = str1
    cbo2.Text = str2
End Sub

Private Sub txtAmtDigits_GotFocus()
    txtAmtDigits.SelStart = 0
    txtAmtDigits.SelLength = 1
End Sub

Private Sub txtAmtDigits_KeyPress(KeyAscii As Integer)
    If KeyAscii < &H30 Or KeyAscii > &H39 Then
        If KeyAscii <> &H8 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAmtSepa_GotFocus()
    txtAmtSepa.SelStart = 0
    txtAmtSepa.SelLength = 1
End Sub

Private Sub txtBottom_Change()
    If Not IsNumeric(txtBottom.Text) Then txtBottom.Text = 0
End Sub

Private Sub txtBottom_DblClick()
    With frmKeyboard
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtBottom.Text = .Let_Text_Input
    End With
End Sub

Private Sub txtGroupDigits_GotFocus()
    txtGroupDigits.SelStart = 0
    txtGroupDigits.SelLength = 1
End Sub

Private Sub txtGroupDigits_KeyPress(KeyAscii As Integer)
    If KeyAscii < &H30 Or KeyAscii > &H39 Then
        If KeyAscii <> &H8 Then KeyAscii = 0
    End If
End Sub

Private Sub txtLeft_Change()
    If Not IsNumeric(txtLeft.Text) Then txtLeft.Text = 0
End Sub

Private Sub txtLeft_DblClick()
    With frmKeyboard
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtLeft.Text = .Let_Text_Input
    End With
End Sub
Private Sub txtRight_Change()
    If Not IsNumeric(txtRight.Text) Then txtRight.Text = 0
End Sub

Private Sub txtRight_DblClick()
    With frmKeyboard
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtRight.Text = .Let_Text_Input
    End With
End Sub

Private Sub txtQtyDigits_GotFocus()
    txtQtyDigits.SelStart = 0
    txtQtyDigits.SelLength = 1
End Sub

Private Sub txtQtyDigits_KeyPress(KeyAscii As Integer)
    If KeyAscii < &H30 Or KeyAscii > &H39 Then
        If KeyAscii <> &H8 Then KeyAscii = 0
    End If
End Sub

Private Sub txtQtySepa_GotFocus()
    txtQtySepa.SelStart = 0
    txtQtySepa.SelLength = 1
End Sub


Private Sub txtTop_Change()
    If Not IsNumeric(txtTop.Text) Then txtTop.Text = 0
End Sub

Private Sub txtTop_DblClick()
    With frmKeyboard
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtTop.Text = .Let_Text_Input
    End With
End Sub

Private Sub udAmtDigits_DownClick()
    If txtAmtDigits.Text <> "" Then
        If CInt(txtAmtDigits.Text) > 0 Then
            txtAmtDigits.Text = CInt(txtAmtDigits.Text) - 1
        End If
    Else
        txtAmtDigits.Text = 0
    End If
End Sub

Private Sub udAmtDigits_UpClick()
    If txtAmtDigits.Text <> "" Then
        If CInt(txtAmtDigits.Text) < 9 Then
            txtAmtDigits.Text = CInt(txtAmtDigits.Text) + 1
        End If
    Else
        txtAmtDigits.Text = 0
    End If
End Sub

Private Sub udGroupDigits_DownClick()
    If txtGroupDigits.Text <> "" Then
        If CInt(txtGroupDigits.Text) > 0 Then
            txtGroupDigits.Text = CInt(txtGroupDigits.Text) - 1
        End If
    Else
        txtGroupDigits.Text = 0
    End If
End Sub

Private Sub udGroupDigits_UpClick()
    If txtGroupDigits.Text <> "" Then
        If CInt(txtGroupDigits.Text) < 9 Then
            txtGroupDigits.Text = CInt(txtGroupDigits.Text) + 1
        End If
    Else
        txtGroupDigits.Text = 0
    End If
End Sub

Private Sub udQtyDigits_DownClick()
    If txtQtyDigits.Text <> "" Then
        If CInt(txtQtyDigits.Text) > 0 Then
            txtQtyDigits.Text = CInt(txtQtyDigits.Text) - 1
        End If
    Else
        txtQtyDigits.Text = 0
    End If
End Sub

Private Sub udQtyDigits_UpClick()
    If txtQtyDigits.Text <> "" Then
        If CInt(txtQtyDigits.Text) < 9 Then
            txtQtyDigits.Text = CInt(txtQtyDigits.Text) + 1
        End If
    Else
        txtQtyDigits.Text = 0
    End If
End Sub


Public Sub AddValuetoformat()
On Error GoTo Handle
    With cboFormat
        .Clear
        .AddItem "#,##0"
        .AddItem "#,##0.0"
        .AddItem "#,##0.00"
        .AddItem "#,##0.000"
        .AddItem "#.##0"
        .AddItem "#.##0,0"
        .AddItem "#.##0,00"
        .AddItem "#.##0,000"
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " AddValue"
End Sub
