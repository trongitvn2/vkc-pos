VERSION 5.00
Begin VB.Form frmChangeBill 
   Caption         =   "Change"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   2775
   ScaleWidth      =   7635
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   3
      Top             =   180
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
      TabIndex        =   1
      Top             =   1860
      Width           =   3615
   End
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
      BTYPE           =   5
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
      BCOL            =   16711680
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmChangeBill.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
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
      TabIndex        =   6
      Top             =   270
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
      TabIndex        =   4
      Top             =   1950
      Width           =   1635
   End
End
Attribute VB_Name = "frmChangeBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BillNO As Double
Dim i As Integer

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
If isActived = True Then Exit Sub
isActived = True
    If cmdOk.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    i = 0
    isActived = False
        Call get_Change(BillNO)
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "   Form_Load"
End Sub

Private Sub Timer1_Timer()
    i = i + 1
    If i = 5 Then Call cmdOK_Click
End Sub

Public Property Let Let_Bill(ByVal vNewValue As Variant)
    BillNO = vNewValue
End Property

Public Sub get_Change(Bill As Double)
    On Error GoTo Handle
        Dim rsTotal As New ADODB.Recordset
        Set rsTotal = Open_Table(cnData, "Invoice_Totals")
        With rsTotal
            .Find "Invoice_Number='" & Bill & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtTotal.Text = Format(.Fields("Grand_Total"), formatNum)
                txtTender.Text = Format(.Fields("Amt_Tendered"), formatNum)
                txtChange.Text = Format(CDbl(txtTender.Text) - CDbl(txtTotal.Text), formatNum)
            End If
            
        End With
    
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & " - get_Change"
End Sub
