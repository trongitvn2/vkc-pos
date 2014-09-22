VERSION 5.00
Begin VB.Form frmChange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin prjTouchScreen.MyButton cmdOk 
      Height          =   2535
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   4471
      BTYPE           =   14
      TX              =   "&OK"
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
      BCOL            =   16777215
      BCOLO           =   16711680
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmChange.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.TextBox txtChange 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   765
      Left            =   1800
      TabIndex        =   3
      Top             =   1860
      Width           =   3615
   End
   Begin VB.TextBox txtTender 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   765
      Left            =   1800
      TabIndex        =   2
      Top             =   1050
      Width           =   3615
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   765
      Left            =   1800
      TabIndex        =   1
      Top             =   180
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Thèi l¹i :"
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
      Height          =   585
      Left            =   90
      TabIndex        =   6
      Top             =   1950
      Width           =   1635
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Thanh to¸n :"
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
      Height          =   585
      Left            =   90
      TabIndex        =   5
      Top             =   1110
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tæng céng:"
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
      Height          =   585
      Left            =   90
      TabIndex        =   4
      Top             =   270
      Width           =   1635
   End
End
Attribute VB_Name = "frmChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Total As Double
Dim Tender_Amt As Double
Dim BillNO As Double
Dim i As Integer
Dim isActived As Boolean

Public Property Let GetTotal(ByVal vNewValue As Variant)
    Total = vNewValue
'   Total = CInt(vNewValue / 1000)
End Property

Public Property Let GetTender_Amt(ByVal vNewValue As Variant)
    Tender_Amt = vNewValue
End Property

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
If isActived = True Then Exit Sub
isActived = True
    If cmdOK.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    isActived = False
        i = 0
        TxtTotal.Text = Format(Total, formatNum)
        txtTender.Text = Format(Tender_Amt, formatNum)
        txtChange.Text = Format(CDbl(txtTender.Text) - CDbl(TxtTotal.Text), formatNum)
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "   Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Total = 0
    Tender_Amt = 0
End Sub


Private Sub Timer1_Timer()
    i = i + 1
    If i = 5 Then Call cmdOK_Click
End Sub

