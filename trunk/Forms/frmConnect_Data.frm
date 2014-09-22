VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmConnect_Data 
   Caption         =   "KÕt nèi d÷ liÖu"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9300
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnect_Data.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3915
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   1320
      Width           =   3255
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Height          =   855
      Left            =   7800
      TabIndex        =   9
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConnect_Data.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdSave 
      Height          =   855
      Left            =   6000
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "L­u"
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConnect_Data.frx":0028
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
      Left            =   4200
      TabIndex        =   7
      Top             =   2760
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "Thªm "
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmConnect_Data.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.TextBox txtpass 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   6000
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1920
      Width           =   3255
   End
   Begin VB.TextBox txtDBName 
      Height          =   495
      Left            =   6000
      TabIndex        =   2
      Top             =   720
      Width           =   3255
   End
   Begin VB.TextBox txtservername 
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin MSFlexGridLib.MSFlexGrid flgDB 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6588
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Tªn ®¨ng nhËp:"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4080
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "MËt khÈu:"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4080
      TabIndex        =   6
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Tªn CSDL:"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4080
      TabIndex        =   5
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tªn m¸y chñ:"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmConnect_Data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
On Error GoTo Handle
    cmdAdd.Enabled = False
    txtservername.Text = ""
    txtDBName.Text = ""
    txtUserName.Text = ""
    txtpass.Text = ""
    cmdSave.Enabled = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdAdd_Click"
End Sub

Private Sub cmdClose_Click()
If cnData.State = 0 Then
    End
Else
    Unload Me
End If
End Sub

Private Sub cmdSave_Click()
On Error GoTo Handle
    SaveSettingStr "SYSTEM", "ServerName", txtservername.Text, myIniFile
    SaveSettingStr "SYSTEM", "DatabaseName", txtDBName.Text, myIniFile
    SaveSettingStr "SYSTEM", "Password", En_Decryption.MalgoEncrypt(txtpass.Text, 10), myIniFile
    SaveSettingStr "SYSTEM", "UserLogin", txtUserName.Text, myIniFile
    cmdAdd.Enabled = True
    
    If ServerName <> txtservername.Text Then
        ServerName = GetSettingStr("SYSTEM", "ServerName", "", myIniFile)
        DataBaseName = GetSettingStr("SYSTEM", "DatabaseName", "", myIniFile)
        UserLog = GetSettingStr("SYSTEM", "UserLogin", "", myIniFile)
        DB_Password = GetSettingStr("SYSTEM", "Password", "", myIniFile)
        DB_Password = En_Decryption.MalgoDecrypt(DB_Password, 10)
    End If
   Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    Unload Me
Exit Sub
Handle:
        MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub flgDB_Click()
    txtservername.Text = flgDB.TextMatrix(1, 0)
    txtDBName.Text = flgDB.TextMatrix(1, 1)
    txtUserName.Text = flgDB.TextMatrix(1, 2)
    txtpass.Text = DB_Password
End Sub

Private Sub flgDB_EnterCell()
    txtservername.Text = flgDB.TextMatrix(1, 0)
    txtDBName.Text = flgDB.TextMatrix(1, 1)
    txtUserName.Text = flgDB.TextMatrix(1, 2)
    txtpass.Text = DB_Password
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    With flgDB
        .Cols = 3
        .Rows = 2
        .TextMatrix(0, 0) = "Tªn Server"
        .TextMatrix(0, 1) = "Tªn CSDL"
        .TextMatrix(0, 2) = "Tªn ®¨ng nhËp"
        .TextMatrix(1, 0) = ServerName
        .TextMatrix(1, 1) = DataBaseName
        .TextMatrix(1, 2) = UserLog
        txtservername.Text = flgDB.TextMatrix(1, 0)
        txtDBName.Text = flgDB.TextMatrix(1, 1)
        txtUserName.Text = flgDB.TextMatrix(1, 2)
        txtpass.Text = DB_Password
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub

