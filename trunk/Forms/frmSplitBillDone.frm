VERSION 5.00
Begin VB.Form frmSplitBillDone 
   Caption         =   "KÕt thóc hãa ®¬n"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
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
   ScaleHeight     =   3525
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdHold 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2355
      BTYPE           =   4
      TX              =   "&T¹m tÝnh"
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
      BCOL            =   12648447
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSplitBillDone.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdPrintCheck 
      Height          =   1335
      Left            =   2640
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2355
      BTYPE           =   4
      TX              =   "&In hãa ®¬n t¹m"
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
      BCOL            =   12648447
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSplitBillDone.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdCash 
      Height          =   1335
      Left            =   5160
      TabIndex        =   3
      Top             =   1320
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2355
      BTYPE           =   4
      TX              =   "Thanh t&o¸n"
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
      BCOL            =   12648447
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSplitBillDone.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "                            Hãa ®¬n ®­îc chia!                              B¹n muèn lµm g×?"
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
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "frmSplitBillDone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bill_Master As Integer

Private Sub cmdHold_Click()
    Unload Me
End Sub

Private Sub cmdPrintCheck_Click()
    Unload Me
    With frmBillDone
        .GetBill_Master = Bill_Master
        .Show vbModal
    End With
    Unload Me
End Sub

Public Property Let Get_Master_Bill(ByVal vNewValue As Variant)
    Bill_Master = vNewValue
End Property
Public Property Get Get_Master_Bill() As Variant
    Get_Master_Bill = Bill_Master
End Property

