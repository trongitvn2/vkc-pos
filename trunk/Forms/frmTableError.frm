VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmTableError 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bµn lçi"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTableError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   9255
      Begin VB.ListBox lstTable 
         Height          =   3480
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   240
         Width           =   9015
      End
   End
   Begin prjTouchScreen.MyButton cmdExit 
      Height          =   975
      Left            =   4800
      TabIndex        =   0
      Top             =   5160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1720
      BTYPE           =   14
      TX              =   "&Tho¸t"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   12582912
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTableError.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmProcess 
      Height          =   975
      Left            =   2640
      TabIndex        =   2
      Top             =   5160
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1720
      BTYPE           =   14
      TX              =   "&Xö lý "
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   12582912
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTableError.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSForms.CheckBox CheckBox1 
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   4080
      Width           =   2175
      BackColor       =   -2147483633
      ForeColor       =   12582912
      DisplayStyle    =   4
      Size            =   "3836;661"
      Value           =   "0"
      Caption         =   "Bá chän tÊt c¶"
      FontName        =   ".VnArial"
      FontEffects     =   1073741826
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblProcessing 
      Caption         =   "§ang xö lý...."
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSForms.CheckBox chkCheckAll 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   1695
      BackColor       =   -2147483633
      ForeColor       =   12582912
      DisplayStyle    =   4
      Size            =   "2990;661"
      Value           =   "0"
      Caption         =   "Chän tÊt c¶"
      FontName        =   ".VnArial"
      FontEffects     =   1073741826
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmTableError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsInvoice_Totals As New ADODB.Recordset
Dim Arr() As String

Private Sub CheckBox1_Click()
On Error GoTo Handle
Dim i As Integer
    For i = 0 To lstTable.ListCount - 1
        lstTable.Selected(i) = False
    Next
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " chkCheckAll_Click"

End Sub

Private Sub chkCheckAll_Click()
On Error GoTo Handle
Dim i As Integer
CheckBox1.Value = False
    For i = 0 To lstTable.ListCount - 1
        lstTable.Selected(i) = True
    Next
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " chkCheckAll_Click"
End Sub

Private Sub cmdExit_Click()
    Set rsInvoice_Totals = Nothing
    
    Unload Me
End Sub

Private Sub cmProcess_Click()
On Error GoTo Handle
    lblProcessing.Visible = True
   cnData.Execute "Update Invoice_Totals set InvoiceNotesUsed =false"
    'MsgBox "Hoµn tÊt"
    lblProcessing.Caption = "Hoµn tÊt"
    'lblProcessing.Visible = False
    Unload Me
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & ""
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Set rsInvoice_Totals = OpenCriticalTable("select * from Invoice_Totals where InvoiceNotesUsed=true and Status <>'C'", cnData)
    lstTable.Clear
    Dim i As Integer
    With lstTable
        Do While Not rsInvoice_Totals.EOF
            .AddItem "Bµn :" & rsInvoice_Totals.Fields("Orig_OnHoldID")
           ReDim Preserve Arr(rsInvoice_Totals.RecordCount)
           Arr(i) = rsInvoice_Totals.Fields("Orig_OnHoldID")
        rsInvoice_Totals.MoveNext
        i = i + 1
        Loop
    End With
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & ""
End Sub
