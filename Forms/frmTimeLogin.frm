VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmTimeLogin 
   Caption         =   "Giê më"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
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
   Icon            =   "frmTimeLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton cmdOK 
      Height          =   1095
      Left            =   2520
      TabIndex        =   4
      Top             =   2520
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1931
      BTYPE           =   4
      TX              =   "§ång ý"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmTimeLogin.frx":000C
      PICN            =   "frmTimeLogin.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   735
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   1296
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm:ss"
      Format          =   50855938
      UpDown          =   -1  'True
      CurrentDate     =   38462.25
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Height          =   1095
      Left            =   1080
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1931
      BTYPE           =   4
      TX              =   "Hñy"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmTimeLogin.frx":0662
      PICN            =   "frmTimeLogin.frx":067E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "27/04/2008"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   1800
      Width           =   3615
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ngµy:"
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
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Giê më phßng"
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
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label lblTimeIn 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Giê vµo:"
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
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "frmTimeLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TimeIn As String
Dim isOpened, isOK As Boolean

Private Sub cmdClose_Click()
    isOK = False
    Unload Me
End Sub

Private Sub dtpTime_Change()
    If isOpened = False Then
        If Format(dtpTime.Value, "HH:mm:ss") < Format(Now, "HH:mm:ss") Then
            dtpTime.Value = Format(Now, "HH:mm:ss")
            Exit Sub
        End If
    Else
        If Format(dtpTime.Value, "HH:mm:ss") > Format(Now, "HH:mm:ss") Then
            dtpTime.Value = Format(Now, "HH:mm:ss")
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    lblDate.Caption = Format(Now, "dd/MM/yyyy")
    dtpTime.Value = Format(Now, "HH:mm:ss")
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " Form_Load"
End Sub

Public Property Get Get_Time_In() As Variant
    Get_Time_In = TimeIn
End Property

Private Sub cmdOK_Click()
    On Error GoTo Handle
        isOK = True
        TimeIn = dtpTime.Value
        Unload Me
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " cmdOK_Click"
End Sub

Public Property Let GetOpen(ByVal vNewValue As Variant)
    isOpened = vNewValue
End Property

Private Sub Form_Unload(Cancel As Integer)
    If Not isOK Then
        TimeIn = ""
    End If
End Sub
