VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDemo 
   Caption         =   "§¨ng ký"
   ClientHeight    =   4680
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7815
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
   Icon            =   "frmDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "KÝch ho¹t b¶n quyÒn"
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   7575
      Begin VB.TextBox txtDate 
         Height          =   495
         Left            =   1680
         TabIndex        =   9
         Top             =   1440
         Width           =   615
      End
      Begin VB.OptionButton optRegister 
         Caption         =   "KÝch ho¹t b¶n quyÒn ®Çy ®ñ (Full Version)"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   6735
      End
      Begin VB.OptionButton OptTrial 
         Caption         =   "Dïng thö "
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton optOneYear 
         Caption         =   "KÝch ho¹t b¶n quyÒn 1 n¨m (License 1 year)"
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   6735
      End
      Begin VB.Label Label3 
         Caption         =   "ngµy"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
   End
   Begin MSComctlLib.StatusBar stt 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   4305
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   16404
            MinWidth        =   16404
         EndProperty
      EndProperty
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   735
      Left            =   4680
      TabIndex        =   3
      Top             =   3480
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      BTYPE           =   7
      TX              =   "&Hñy bá"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDemo.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdNext 
      Height          =   735
      Left            =   1560
      TabIndex        =   2
      Top             =   3480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1296
      BTYPE           =   7
      TX              =   "&TiÕp tôc"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDemo.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "phuùc thaïnh vinh"
      BeginProperty Font 
         Name            =   "VNI-Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   7575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Chaøo möøng baïn ñeán vôùi phaàn meàm baùn haøng VKC - POS"
      BeginProperty Font 
         Name            =   "VNI-Ariston"
         Size            =   14.25
         Charset         =   0
         Weight          =   800
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdNext_Click()
    If optRegister.Value = True Then
        frmLicense.Show vbModal
    ElseIf optOneYear.Value = True Then
        frmLicense.Show vbModal
    Else
        If Val("0" & txtDate.Text) > 30 Or txtDate.Text = "" Then Exit Sub
        Dim Path_Direction, S As String
        Dim fTemp As Integer
            Path_Direction = "C:\Windows\System32"
            S = gfCONVERT_DATE_TO_STRING(Date + Val("0" & txtDate.Text) - 1)
            If Dir(Path_Direction & "\sysrt.dll", vbDirectory) = "" Then
                fTemp = FreeFile
                Open Path_Direction & "\sysrt.dll" For Output As #fTemp
                Print #fTemp, En_Decryption.MalgoEncrypt(S, 5)
                Close #fTemp
                frmLogin.Show vbModal
            End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    stt.Panels(1).Text = "Mäi th¾c m¾c xin vui lßng liªn hÖ: 0918.655.887 (24/7)"
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " Form_Load"
End Sub

Private Sub OptTrial_Click()
    txtDate.SetFocus
End Sub
