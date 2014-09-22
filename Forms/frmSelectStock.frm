VERSION 5.00
Begin VB.Form frmSelectStock 
   Caption         =   "Lùa chän kho"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
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
   Icon            =   "frmSelectStock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5265
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   735
      Left            =   3840
      TabIndex        =   2
      Top             =   2520
      Width           =   1335
      _extentx        =   2355
      _extenty        =   1296
      btype           =   1
      tx              =   "&Tho¸t"
      enab            =   -1  'True
      font            =   "frmSelectStock.frx":000C
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   14737632
      bcolo           =   14737632
      fcol            =   255
      fcolo           =   255
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmSelectStock.frx":0034
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdStock 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
      _extentx        =   4260
      _extenty        =   1720
      btype           =   1
      tx              =   "Kho &ChÝnh"
      enab            =   -1  'True
      font            =   "frmSelectStock.frx":0052
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   14737632
      bcolo           =   33023
      fcol            =   16711680
      fcolo           =   16777215
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmSelectStock.frx":007A
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdStockRes 
      Height          =   975
      Left            =   2760
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
      _extentx        =   4260
      _extenty        =   1720
      btype           =   1
      tx              =   "Kho &Nhµ hµng"
      enab            =   -1  'True
      font            =   "frmSelectStock.frx":0098
      coltype         =   2
      focusr          =   -1  'True
      bcol            =   14737632
      bcolo           =   33023
      fcol            =   16711680
      fcolo           =   16777215
      mcol            =   12632256
      mptr            =   1
      micon           =   "frmSelectStock.frx":00C0
      umcol           =   -1  'True
      soft            =   0   'False
      picpos          =   0
      ngrey           =   0   'False
      fx              =   0
      hand            =   0   'False
      check           =   0   'False
      value           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lùa chän kho"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmSelectStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim State As String

Public Property Let Let_State(ByVal vNewValue As Variant)
    State = vNewValue
End Property

Private Sub cmdStock_Click()
    On Error GoTo Handle
        Select Case State
            Case "IN"
                frmInstockMaster.Show vbModal
            Case "OUT"
                frmOutstockMaster.Show vbModal
        End Select
        Unload Me
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  cmdStock_Click"
End Sub

Private Sub cmdStockRes_Click()
On Error GoTo Handle
        Select Case State
            Case "IN"
                frmInstockB.Show vbModal
            Case "OUT"
                frmOustockB.Show vbModal
        End Select
        Unload Me
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  cmdStock_Click"
End Sub

Private Sub MyButton1_Click()
    State = ""
    Unload Me
End Sub
