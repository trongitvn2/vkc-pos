VERSION 5.00
Begin VB.Form frmShowEmp_ID 
   Caption         =   " "
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton MyButton2 
      Cancel          =   -1  'True
      Height          =   975
      Left            =   3720
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Hñy"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmShowEmp_ID.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   975
      Left            =   1440
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "§ång ý"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmShowEmp_ID.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblDateTime 
      Caption         =   "Label3"
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
      Left            =   2640
      TabIndex        =   5
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label lblDes 
      Caption         =   "Vµo ca lóc:"
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
      Left            =   480
      TabIndex        =   4
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label lblEmp_ID 
      Caption         =   "Label3"
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
      Left            =   5400
      TabIndex        =   3
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Tªn :"
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
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "M· sè:"
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
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblEmp_Name 
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
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmShowEmp_ID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isOK As Boolean
Dim EmpID As String
Dim EmpName As String

Private Sub Form_Activate()
    lblEmp_ID.Caption = EmpID
    lblEmp_Name.Caption = EmpName
    lblDatetime.Caption = Now
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
        lblEmp_ID.Caption = EmpID
        lblEmp_Name.Caption = EmpName
        lblDatetime.Caption = Now
    Exit Sub
    
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - Form_Load"
End Sub


Public Property Let Emp_ID(ByVal vNewValue As Variant)
    EmpID = vNewValue
End Property

Public Property Let Emp_Name(ByVal vNewValue As Variant)
    EmpName = vNewValue
End Property

Private Sub MyButton1_Click()
    isOK = True
    Unload Me
End Sub

Private Sub MyButton2_Click()
    isOK = False
    Unload Me
End Sub

Public Property Get Let_OK() As Variant
    Let_OK = isOK
End Property

