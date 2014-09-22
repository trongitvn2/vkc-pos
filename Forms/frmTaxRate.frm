VERSION 5.00
Begin VB.Form frmTaxRate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "T˚ l÷ thu’"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
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
   Icon            =   "frmTaxRate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdOK 
      Height          =   795
      Left            =   1590
      TabIndex        =   16
      Tag             =   "L8"
      Top             =   3420
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1402
      BTYPE           =   14
      TX              =   "&ßÂng ˝"
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
      BCOL            =   16578804
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTaxRate.frx":000C
      PICN            =   "frmTaxRate.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Frame fraTaxRate 
      Caption         =   "T˚ l÷ thu’"
      Height          =   3255
      Left            =   30
      TabIndex        =   0
      Tag             =   "L1"
      Top             =   30
      Width           =   6495
      Begin VB.TextBox txtDescription4 
         Height          =   495
         Left            =   1680
         TabIndex        =   20
         Top             =   2460
         Width           =   1905
      End
      Begin VB.TextBox txtOldRate4 
         Height          =   495
         Left            =   3600
         TabIndex        =   19
         Top             =   2460
         Width           =   1305
      End
      Begin VB.TextBox txtNewRate4 
         Height          =   495
         Left            =   4920
         TabIndex        =   18
         Top             =   2460
         Width           =   1365
      End
      Begin VB.TextBox txtNewRate3 
         Height          =   495
         Left            =   4920
         TabIndex        =   15
         Top             =   1890
         Width           =   1365
      End
      Begin VB.TextBox txtOldRate3 
         Height          =   495
         Left            =   3600
         TabIndex        =   14
         Top             =   1890
         Width           =   1365
      End
      Begin VB.TextBox txtDescription3 
         Height          =   495
         Left            =   1680
         TabIndex        =   13
         Top             =   1890
         Width           =   1995
      End
      Begin VB.TextBox txtNewRate2 
         Height          =   495
         Left            =   4920
         TabIndex        =   12
         Top             =   1320
         Width           =   1365
      End
      Begin VB.TextBox txtOldRate2 
         Height          =   495
         Left            =   3600
         TabIndex        =   11
         Top             =   1320
         Width           =   1365
      End
      Begin VB.TextBox txtDescription2 
         Height          =   495
         Left            =   1680
         TabIndex        =   10
         Top             =   1320
         Width           =   1995
      End
      Begin VB.TextBox txtNewRate1 
         Height          =   495
         Left            =   4920
         TabIndex        =   9
         Top             =   750
         Width           =   1365
      End
      Begin VB.TextBox txtOldRate1 
         Height          =   495
         Left            =   3600
         TabIndex        =   8
         Top             =   750
         Width           =   1365
      End
      Begin VB.TextBox txtDescription1 
         Height          =   495
         Left            =   1680
         TabIndex        =   7
         Top             =   750
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Thu’ su t 4:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   90
         TabIndex        =   21
         Tag             =   "L10"
         Top             =   2550
         Width           =   1455
      End
      Begin VB.Label lblNewRate 
         Alignment       =   2  'Center
         Caption         =   "T˚ l÷ mÌi"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4860
         TabIndex        =   6
         Tag             =   "L7"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblOldRate 
         Alignment       =   2  'Center
         Caption         =   "T˚ l÷ cÚ"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3270
         TabIndex        =   5
         Tag             =   "L6"
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label lblDescription 
         Alignment       =   2  'Center
         Caption         =   "Di‘n gi∂i"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1770
         TabIndex        =   4
         Tag             =   "L5"
         Top             =   330
         Width           =   1335
      End
      Begin VB.Label lblTax3 
         Alignment       =   1  'Right Justify
         Caption         =   "Thu’ su t 3:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   90
         TabIndex        =   3
         Tag             =   "L4"
         Top             =   1980
         Width           =   1455
      End
      Begin VB.Label lblTax2 
         Alignment       =   1  'Right Justify
         Caption         =   "Thu’ su t 2:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   90
         TabIndex        =   2
         Tag             =   "L3"
         Top             =   1380
         Width           =   1455
      End
      Begin VB.Label lblTax1 
         Alignment       =   1  'Right Justify
         Caption         =   "Thu’ su t 1:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   120
         TabIndex        =   1
         Tag             =   "L2"
         Top             =   780
         Width           =   1455
      End
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   795
      Left            =   3360
      TabIndex        =   17
      Tag             =   "L9"
      Top             =   3420
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1402
      BTYPE           =   14
      TX              =   "&Tho∏t"
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
      BCOL            =   16578804
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTaxRate.frx":0662
      PICN            =   "frmTaxRate.frx":067E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
End
Attribute VB_Name = "frmTaxRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Desarr() As String
Dim rsTaxrate As New ADODB.Recordset

Private Sub cmdCancel_Click()
On Error GoTo Handle
    Unload Me
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Private Sub cmdOK_Click()
    On Error GoTo Handle
    With rsTaxrate
        .Fields("Tax1_Name") = txtDescription1.Text
        .Fields("Tax2_Name") = txtDescription2.Text
        .Fields("Tax3_Name") = txtDescription3.Text
        .Fields("Tax4_Name") = txtDescription4.Text
        .Fields("Tax1_Rate") = Val(txtNewRate1.Text)
        .Fields("Tax2_Rate") = Val(txtNewRate2.Text)
        .Fields("Tax3_Rate") = Val(txtNewRate3.Text)
        .Fields("Tax4_Rate") = Val(txtNewRate4.Text)
        .Update
        .Requery
    End With
    MsgBox "Hoµn t t!", vbInformation
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Load"

End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    Dim ctrl As Control
    If cmdCancel.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Desarr = LoadLanguage(LngFile, "#01:010:")
        Me.Caption = Desarr(1)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Desarr(Mid(ctrl.Tag, 2))
        Next ctrl
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Load"

End Sub

Private Sub Form_Load()
On Error GoTo Handle
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsTaxrate = OpenCriticalTable("select * from Tax_Rate", cnData)
    Call LoadText
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Public Sub LoadText()
On Error GoTo Handle
    If rsTaxrate.RecordCount > 0 Then
        With rsTaxrate
            txtDescription1.Text = .Fields("Tax1_Name")
            txtDescription2.Text = .Fields("Tax2_Name")
            txtDescription3.Text = .Fields("Tax3_Name")
            txtDescription4.Text = .Fields("Tax4_Name")
            txtOldRate1.Text = .Fields("Tax1_Rate") & "%"
            txtOldRate2.Text = .Fields("Tax2_Rate") & "%"
            txtOldRate3.Text = .Fields("Tax3_Rate") & "%"
            txtOldRate4.Text = .Fields("Tax4_Rate") & "%"
            txtNewRate1.Text = .Fields("Tax1_Rate") & "%"
            txtNewRate2.Text = .Fields("Tax2_Rate") & "%"
            txtNewRate3.Text = .Fields("Tax3_Rate") & "%"
            txtNewRate4.Text = .Fields("Tax4_Rate") & "%"
            
        End With
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " LoadText"
End Sub
