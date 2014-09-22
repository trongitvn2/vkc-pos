VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCredit_Card_infor 
   Caption         =   "Credit Card Information"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15060
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial Narrow"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11055
   ScaleWidth      =   15060
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   5655
      Left            =   10560
      TabIndex        =   30
      Top             =   1200
      Width           =   4455
      Begin MSComCtl2.DTPicker DtpCard_Expire 
         Height          =   495
         Left            =   1440
         TabIndex        =   47
         Top             =   4320
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   61014017
         CurrentDate     =   41212
      End
      Begin VB.OptionButton optType 
         Height          =   270
         Index           =   3
         Left            =   3720
         TabIndex        =   46
         Top             =   3000
         Width           =   255
      End
      Begin VB.OptionButton optType 
         Height          =   270
         Index           =   2
         Left            =   2640
         TabIndex        =   45
         Top             =   3000
         Width           =   255
      End
      Begin VB.OptionButton optType 
         Height          =   270
         Index           =   1
         Left            =   1560
         TabIndex        =   44
         Top             =   3000
         Width           =   255
      End
      Begin VB.OptionButton optType 
         Height          =   270
         Index           =   0
         Left            =   360
         TabIndex        =   43
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox txtAdd 
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
         Height          =   435
         Left            =   240
         TabIndex        =   4
         Top             =   5160
         Width           =   4095
      End
      Begin VB.TextBox txtCard_Code 
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
         Height          =   435
         Left            =   1560
         TabIndex        =   2
         Top             =   1920
         Width           =   2775
      End
      Begin VB.TextBox txtCardType 
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
         Height          =   435
         Left            =   1560
         TabIndex        =   3
         Top             =   2400
         Width           =   2775
      End
      Begin VB.TextBox txtAccName 
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
         Height          =   435
         Left            =   1560
         TabIndex        =   1
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox txtTrancode 
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
         Height          =   435
         Left            =   1560
         TabIndex        =   0
         Top             =   960
         Width           =   2775
      End
      Begin VB.Image cardtype 
         Height          =   840
         Index           =   3
         Left            =   3360
         Picture         =   "frmCredit_Card_infor.frx":0000
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Image cardtype 
         Height          =   840
         Index           =   2
         Left            =   2265
         Picture         =   "frmCredit_Card_infor.frx":6BDD
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Image cardtype 
         Height          =   840
         Index           =   1
         Left            =   1155
         Picture         =   "frmCredit_Card_infor.frx":E7D2
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Image cardtype 
         Height          =   840
         Index           =   0
         Left            =   60
         Picture         =   "frmCredit_Card_infor.frx":17142
         Stretch         =   -1  'True
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Label Label7 
         Caption         =   "Sè giao dÞch:"
         Height          =   375
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "§Þa chØ chñ thÎ:"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Ngµy hÕt h¹n:"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "M· sè thÎ:"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Lo¹i thÎ:"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Tªn  chñ thÎ:"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Th«ng tin chñ thÎ"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   3975
      End
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Height          =   975
      Left            =   12840
      TabIndex        =   6
      Top             =   7320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "§ãng"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCredit_Card_infor.frx":1DA78
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdOk 
      Height          =   975
      Left            =   10680
      TabIndex        =   5
      Top             =   7320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "§ång ý"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCredit_Card_infor.frx":1DA94
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
      Height          =   7575
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   10335
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   26
         Left            =   2760
         TabIndex        =   42
         Top             =   7080
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   25
         Left            =   720
         TabIndex        =   41
         Top             =   7080
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   24
         Left            =   8880
         TabIndex        =   40
         Top             =   5880
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   23
         Left            =   6840
         TabIndex        =   39
         Top             =   5880
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   22
         Left            =   4800
         TabIndex        =   38
         Top             =   5880
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   21
         Left            =   2760
         TabIndex        =   29
         Top             =   5880
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   6
         Left            =   2760
         TabIndex        =   28
         Top             =   2280
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   20
         Left            =   720
         TabIndex        =   27
         Top             =   5880
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   19
         Left            =   8880
         TabIndex        =   26
         Top             =   4680
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   18
         Left            =   6840
         TabIndex        =   25
         Top             =   4680
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   17
         Left            =   4800
         TabIndex        =   24
         Top             =   4680
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   16
         Left            =   2760
         TabIndex        =   23
         Top             =   4680
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   15
         Left            =   720
         TabIndex        =   22
         Top             =   4680
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   14
         Left            =   8880
         TabIndex        =   21
         Top             =   3480
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   9
         Left            =   8880
         TabIndex        =   20
         Top             =   2280
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   4
         Left            =   8880
         TabIndex        =   19
         Top             =   1080
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   13
         Left            =   6840
         TabIndex        =   18
         Top             =   3480
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   12
         Left            =   4800
         TabIndex        =   17
         Top             =   3480
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   11
         Left            =   2760
         TabIndex        =   16
         Top             =   3480
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   8
         Left            =   6840
         TabIndex        =   15
         Top             =   2280
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   7
         Left            =   4800
         TabIndex        =   14
         Top             =   2280
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   3
         Left            =   6840
         TabIndex        =   13
         Top             =   1080
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   2
         Left            =   4800
         TabIndex        =   12
         Top             =   1080
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   1
         Left            =   2760
         TabIndex        =   11
         Top             =   1080
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   10
         Left            =   720
         TabIndex        =   10
         Top             =   3480
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   9
         Top             =   2280
         Width           =   375
      End
      Begin VB.OptionButton optBank 
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   8
         Top             =   1080
         Value           =   -1  'True
         Width           =   375
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   26
         Left            =   2160
         Picture         =   "frmCredit_Card_infor.frx":1DAB0
         Stretch         =   -1  'True
         Top             =   6240
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   25
         Left            =   120
         Picture         =   "frmCredit_Card_infor.frx":1F2F9
         Stretch         =   -1  'True
         Top             =   6240
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   24
         Left            =   8280
         Picture         =   "frmCredit_Card_infor.frx":2C5DD
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   23
         Left            =   6240
         Picture         =   "frmCredit_Card_infor.frx":2E209
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   22
         Left            =   4200
         Picture         =   "frmCredit_Card_infor.frx":2F30C
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   21
         Left            =   2160
         Picture         =   "frmCredit_Card_infor.frx":30D2E
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   6
         Left            =   2160
         Picture         =   "frmCredit_Card_infor.frx":316B5
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   20
         Left            =   120
         Picture         =   "frmCredit_Card_infor.frx":332FD
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   19
         Left            =   8280
         Picture         =   "frmCredit_Card_infor.frx":34154
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   18
         Left            =   6240
         Picture         =   "frmCredit_Card_infor.frx":35BFF
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   17
         Left            =   4200
         Picture         =   "frmCredit_Card_infor.frx":376BD
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   16
         Left            =   2160
         Picture         =   "frmCredit_Card_infor.frx":39E7D
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   15
         Left            =   120
         Picture         =   "frmCredit_Card_infor.frx":3AFBA
         Stretch         =   -1  'True
         Top             =   3840
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   14
         Left            =   8280
         Picture         =   "frmCredit_Card_infor.frx":3BE3D
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   9
         Left            =   8280
         Picture         =   "frmCredit_Card_infor.frx":4598A
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   4
         Left            =   8280
         Picture         =   "frmCredit_Card_infor.frx":46CBD
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   13
         Left            =   6240
         Picture         =   "frmCredit_Card_infor.frx":48512
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   12
         Left            =   4200
         Picture         =   "frmCredit_Card_infor.frx":49D35
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   11
         Left            =   2160
         Picture         =   "frmCredit_Card_infor.frx":4B937
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   8
         Left            =   6240
         Picture         =   "frmCredit_Card_infor.frx":4C50A
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   7
         Left            =   4200
         Picture         =   "frmCredit_Card_infor.frx":4E0FB
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   3
         Left            =   6240
         Picture         =   "frmCredit_Card_infor.frx":4FFE5
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   2
         Left            =   4200
         Picture         =   "frmCredit_Card_infor.frx":51A37
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   1
         Left            =   2160
         Picture         =   "frmCredit_Card_infor.frx":542B3
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   10
         Left            =   120
         Picture         =   "frmCredit_Card_infor.frx":56361
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   5
         Left            =   120
         Picture         =   "frmCredit_Card_infor.frx":59531
         Stretch         =   -1  'True
         Top             =   1440
         Width           =   1860
      End
      Begin VB.Image imgBank 
         Height          =   780
         Index           =   0
         Left            =   120
         Picture         =   "frmCredit_Card_infor.frx":5A193
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1860
      End
   End
End
Attribute VB_Name = "frmCredit_Card_infor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iSelect As Integer
Dim Banking_Name As String
Dim isOK As Boolean
Dim rsCredit As New ADODB.Recordset

Private Sub cardtype_Click(Index As Integer)
    optType(Index).Value = True
End Sub

Private Sub cmdClose_Click()
    isOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo Handle
    If Not Check_Null Then
        isOK = True
        Call Save_Records
        Unload Me
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " cmdOK_Click"
End Sub

Public Sub Get_BankingName(i As Integer)
On Error GoTo Handle
Select Case i
    Case 0
        Banking_Name = "Ng©n hµng An B×nh"
    Case 1
        Banking_Name = "Ng©n hµng ¸ Ch©u"
    Case 2
        Banking_Name = "Ng©n hµng n«ng nghiÖp vµ ph¸t triÓn n«ng th«n"
    Case 3
        Banking_Name = "Ng©n hµng TMCP §«ng ¸"
    Case 4
        Banking_Name = "Ng©n hµng TMCP §«ng Nam ¸"
    Case 5
        Banking_Name = "Ng©n hµng XuÊt NhËp KhÈu ViÖt Nam"
    Case 6
        Banking_Name = "Ng©n hµng TMCP §¹i ¸"
    Case 7
        Banking_Name = "Ng©n hµng TMCP Ph¸t triÒn TP.HCM"
     Case 8
        Banking_Name = "Ng©n hµng TMCP Qu©n §éi"
    Case 9
        Banking_Name = "Ng©n hµng TMCP Sµi Gßn - Hµ Néi"
     Case 10
        Banking_Name = "Ng©n hµng TMCP Kiªn Long"
     Case 11
        Banking_Name = "Ng©n hµng Ph¸t TriÓn Mª K«ng"
     Case 12
        Banking_Name = "Ng©n hµng Sµi Gßn Th­¬ng TÝn"
     Case 13
        Banking_Name = "Ng©n hµng TMCP Sµi Gßn"
     Case 14
        Banking_Name = "Ng©n hµng TMCP §¹i TÝn"
     Case 15
        Banking_Name = "Ng©n hµng Ngo¹i Th­¬ng ViÖt Nam"
     Case 16
        Banking_Name = "Ng©n hµng TMCP ViÖt Nam Th­¬ng TÝn"
     Case 17
        Banking_Name = "Ng©n hµng TMCP B¶n ViÖt"
    Case 18
        Banking_Name = "Ng©n hµng TMCP C«ng th­¬ng ViÖt Nam"
    Case 19
        Banking_Name = "Ng©n hµng TMCP B¶o ViÖt"
     Case 20
        Banking_Name = "Ng©n hµng TMCP B¾c ¸"
     Case 21
        Banking_Name = "Ng©n hµng TMCP DÇu khÝ"
    Case 22
        Banking_Name = "Ng©n hµng TMCP ViÖt ¸"
    Case 23
        Banking_Name = "Ng©n hµngTMCP Hµng H¶i ViÖt Nam"
    Case 24
        Banking_Name = "Ng©n hµngTMCP Quèc TÕ ViÖt Nam"
    Case 25
        Banking_Name = "Ng©n hµng ViÖt Nam ThÞnh V­îng"
    Case 26
        Banking_Name = "Ng©n hµng TMCP kü th­¬ng ViÖt Nam"
End Select
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " "
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    With rsCredit
        If .State = 0 Then
            .Fields.Append "Transaction_Code", adVarWChar, 20
            .Fields.Append "Banking_Name", adVarWChar, 150
            .Fields.Append "Account_Name", adVarWChar, 150
            .Fields.Append "Card_Code", adVarWChar, 30
            .Fields.Append "Card_Type", adVarWChar, 50
            .Fields.Append "Card_Expired", adVarWChar, 20
            .Fields.Append "Account_Add", adVarWChar, 250
            .Open
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "  Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsCredit_Payment = Nothing
End Sub

Private Sub imgBank_Click(Index As Integer)
    optBank(Index).Value = True
End Sub

Private Sub optBank_Click(Index As Integer)
On Error GoTo Handle
    optBank(Index).Value = True
    Call Get_BankingName(Index)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "optBank_Click "
End Sub

Public Property Get Let_OK() As Variant
    Let_OK = isOK
End Property


Public Function Check_Null() As Boolean
On Error GoTo Handle
    Dim isTextNull As Boolean
    If txtTrancode.Text = "" Then
        Call Message_Null(1)
        isTextNull = True
        GoTo 1
    Else
        isTextNull = False
    End If
    If txtAccName.Text = "" Then
        Call Message_Null(2)
        isTextNull = True
        GoTo 1
    Else
        isTextNull = False
    End If
    
    If txtCardType.Text = "" Then
        Call Message_Null(3)
        isTextNull = True
        GoTo 1
    Else
        isTextNull = False
    End If
    
    If txtCard_Code.Text = "" Then
        Call Message_Null(4)
        isTextNull = True
        GoTo 1
    Else
        isTextNull = False
    End If
    
    If IsDate(DtpCard_Expire.Value) = False Then
        Call Message_Null(5)
        isTextNull = True
        GoTo 1
    Else
        isTextNull = False
    End If
1:    Check_Null = isTextNull
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " Check_Null"

End Function

Public Sub Message_Null(ByVal number_err As Integer)
On Error GoTo Handle
Dim message As String
    Select Case number_err
        Case 1:
            message = "Sè giao dÞch cña ng©n hµng kh«ng ®­îc rçng"
        Case 2:
            message = "Tªn chñ thÎ kh«ng ®­îc rçng"
        Case 3:
            message = "Lo¹i thÎ kh«ng ®­îc rçng"
        Case 4:
            message = "M· sè thÎ kh«ng ®­îc rçng"
        Case 5:
            message = "Sai ®Þnh d¹ng ngµy (dd/mm/yyyy)"
    End Select
    MsgBox message
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " Message_Null"

End Sub

Private Sub optType_Click(Index As Integer)
On Error GoTo Handle
    optType(Index).Value = True
    txtCardType.Text = Get_CardType(Index)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "optType_Click "

End Sub

Public Function Get_CardType(i As Integer) As String
On Error GoTo Handle
Dim Card_type As String
Select Case i
    Case 0
        Card_type = "VISA CARD"
    Case 1
        Card_type = "MASTER CARD"
    Case 2
        Card_type = "AMERICAN EXPRESS"
    Case 3
        Card_type = "DINNER CLUD CARD"
End Select
Get_CardType = Card_type
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.Name
End Function


Private Sub txtAccName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtCard_Code.SetFocus
End Sub

Private Sub txtCard_Code_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtCardType.SetFocus

End Sub

Private Sub txtCard_Expire_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtAdd.SetFocus
End Sub

Private Sub txtCardType_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then DtpCard_Expire.SetFocus
End Sub

Private Sub txtTrancode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtAccName.SetFocus
End Sub

Public Sub Save_Records()
On Error GoTo Handle
    With rsCredit
        If .State <> 0 Then
            .addNew
            .Fields("Transaction_Code") = txtTrancode.Text
            .Fields("Banking_Name") = Banking_Name
            .Fields("Account_Name") = txtAccName.Text
            .Fields("Card_Code") = txtCard_Code.Text
            .Fields("Card_Type") = txtCardType.Text
            .Fields("Card_Expired") = gfCONVERT_DATE_TO_STRING(DtpCard_Expire.Value)
            .Fields("Account_Add") = txtAdd.Text
            .Update
        End If
        End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "Save_Records "

End Sub

Public Property Get Let_Records() As Variant
   Set Let_Records = rsCredit
End Property
