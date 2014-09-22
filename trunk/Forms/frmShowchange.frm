VERSION 5.00
Begin VB.Form frmShowchange 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   615
      Left            =   1680
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&OK"
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
      MICON           =   "frmShowchange.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   600
      Top             =   2400
   End
   Begin VB.TextBox txtChange 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1920
      TabIndex        =   5
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtTender 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtTotals 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Thèi l¹i:"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Kh¸ch tr¶:"
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tæng céng:"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmShowchange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim Invoice_Num As Integer

Private Sub Form_Load()
On Error GoTo Handle
    Dim rsInvoice_Total As New ADODB.Recordset
    Set rsInvoice_Total = Open_Table(cnData, "Invoice_Totals")
    i = 0
    With rsInvoice_Total
        .Find "Invoice_Number=" & Invoice_Num, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtTotals.Text = Format(.Fields("Grand_Total"), "#,##0")
            txtTender.Text = Format(.Fields("Amt_Tendered"), "#,##0")
            txtChange.Text = Format(CDbl(txtTender.Text) - CDbl(txtTotals.Text), "#,##0")
        End If
    End With
    'In bill tinh tien
    With frmShowBillSale
        .GetBill = Invoice_Num
        .Show vbModal
    End With
Exit Sub
Handle:
Exit Sub
    MsgBox Err.Number & Err.Description & Me.Name & " Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Invoice_Num = 0
    i = 0
End Sub

Private Sub MyButton1_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    i = i + 1
    If i >= 7 Then Unload Me
End Sub


Public Property Let Get_Bill(ByVal vNewValue As Variant)
    Invoice_Num = vNewValue
End Property
