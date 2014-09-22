VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAddCustomer 
   Caption         =   "Thªm míi kh¸ch hµng"
   ClientHeight    =   6345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11475
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
   ScaleHeight     =   6345
   ScaleWidth      =   11475
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   11295
      Begin VB.Frame Frame2 
         Height          =   4095
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   11055
         Begin VB.TextBox txtMaxAcc 
            Alignment       =   1  'Right Justify
            Height          =   495
            Left            =   9000
            TabIndex        =   7
            Text            =   "0"
            Top             =   2640
            Width           =   1935
         End
         Begin VB.ComboBox cboPro 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   1
            Text            =   "Combo1"
            Top             =   600
            Width           =   2895
         End
         Begin VB.TextBox txtCustNum 
            Height          =   495
            Left            =   2520
            TabIndex        =   0
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtCustName 
            Height          =   495
            Left            =   2520
            TabIndex        =   2
            Top             =   1200
            Width           =   4815
         End
         Begin VB.TextBox txtCustAdd 
            Height          =   495
            Left            =   2520
            TabIndex        =   4
            Top             =   1920
            Width           =   8415
         End
         Begin VB.TextBox txtCustPhone 
            Height          =   495
            Left            =   2520
            TabIndex        =   5
            Top             =   2640
            Width           =   1815
         End
         Begin VB.TextBox txtCustAcc 
            Height          =   495
            Left            =   5280
            TabIndex        =   6
            Top             =   2640
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker dtpBirthday 
            Height          =   495
            Left            =   9120
            TabIndex        =   3
            Top             =   1200
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   63963137
            UpDown          =   -1  'True
            CurrentDate     =   40553
         End
         Begin MSComCtl2.DTPicker dtpOpenAcc 
            Height          =   495
            Left            =   2520
            TabIndex        =   8
            Top             =   3360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   63963137
            UpDown          =   -1  'True
            CurrentDate     =   39448
         End
         Begin MSComCtl2.DTPicker dtpExp 
            Height          =   495
            Left            =   6840
            TabIndex        =   9
            Top             =   3360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   873
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   63963137
            UpDown          =   -1  'True
            CurrentDate     =   39448
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Ngµy hÕt h¹n:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   25
            Top             =   3360
            Width           =   1935
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Ngµy më TK:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   24
            Top             =   3480
            Width           =   2175
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Giíi h¹n c«ng nî:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7560
            TabIndex        =   23
            Top             =   2640
            Width           =   1455
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "Nhãm kh¸ch hµng"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   615
            Left            =   5760
            TabIndex        =   22
            Top             =   600
            Width           =   2055
         End
         Begin MSForms.CheckBox CheckBox1 
            Height          =   495
            Left            =   4920
            TabIndex        =   11
            Top             =   600
            Width           =   1095
            BackColor       =   -2147483633
            ForeColor       =   16711680
            DisplayStyle    =   4
            Size            =   "1931;873"
            Value           =   "0"
            Caption         =   "(*****)"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Ngµy Sinh nhËt:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7440
            TabIndex        =   20
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "M· kh¸ch hµng (*):"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   240
            TabIndex        =   19
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Tªn kh¸ch hµng(*):"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   240
            TabIndex        =   18
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "§Þa chØ:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   17
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "§iÖn tho¹i:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   16
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label Label5 
            Caption         =   "Sè TK:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4560
            TabIndex        =   15
            Top             =   2760
            Width           =   735
         End
      End
      Begin prjTouchScreen.MyButton cmdClose 
         Cancel          =   -1  'True
         Height          =   855
         Left            =   6360
         TabIndex        =   12
         Top             =   5280
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "&§ãng"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   11.25
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
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAddCustomer.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdCreate 
         Height          =   855
         Left            =   2760
         TabIndex        =   10
         Top             =   5280
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "T¹o míi"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   11.25
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
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAddCustomer.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "T¹o míi kh¸ch hµng"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   2280
         TabIndex        =   21
         Top             =   240
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmAddCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rscust As New ADODB.Recordset

Private Sub CheckBox1_Click()
    If CheckBox1.Value = True Then
        txtCustNum.PasswordChar = "*"
    Else
        txtCustNum.PasswordChar = ""
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()
    On Error GoTo Handle
        With rscust
            .Find "custNum='" & TrimSpecialChar(txtCustNum.Text) & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("custNum") = TrimSpecialChar(txtCustNum.Text)
                .Fields("CustName") = "" & txtCustName.Text
                .Fields("Address") = "" & txtCustAdd.Text
                .Fields("Phone") = "" & txtCustPhone.Text
                .Fields("AccountNo") = "" & txtCustAcc.Text
                .Fields("Cust_Type") = cboPro.Text
                .Fields("Acct_Open_Date") = dtpOpenAcc.Value
                .Fields("Acct_Close_Date") = dtpExp.Value
                .Fields("Acct_Max_Balance") = CDbl("0" & txtCustAcc.Text)
                .Fields("Birthday") = dtpBirthday.Value
                .Update
            End If
        End With
        MsgBox "DA TAO MOI KHACH HANG"
        Unload Me
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdCreate_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Call Set_Promotion_type
    If cnData.State <> 0 Then Set rscust = Open_Table(cnData, "Customer")
    dtpOpenAcc.Value = Date
    dtpExp.Value = Date
    CheckBox1.Value = True
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - Form_Load"
End Sub



Private Sub txtCustAcc_DblClick()
On Error GoTo Handle
    With frmKeyboard
        .txtInput.PasswordChar = ""
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtCustAcc.Text = .Let_Text_Input
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - txtCustAcc_DblClick"

End Sub

Private Sub txtCustAdd_DblClick()
On Error GoTo Handle
    With frmKeyboard
        .txtInput.PasswordChar = ""
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtCustAdd.Text = .Let_Text_Input
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - txtCustAdd_DblClick"

End Sub

Private Sub txtCustName_DblClick()
On Error GoTo Handle
    With frmKeyboard
        .txtInput.PasswordChar = ""
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtCustName.Text = .Let_Text_Input
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - txtCustName_DblClick"
End Sub

Private Sub txtCustNum_Change()
    On Error GoTo Handle
        If CheckBox1.Value = True Then txtCustNum.PasswordChar = "*"
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtCustNum_Change"
End Sub


Public Sub Set_Promotion_type()
On Error GoTo Handle
Dim rsCust_Type As New ADODB.Recordset
Set rsCust_Type = Open_Table(cnData, "Customer_Type")
If rsCust_Type.State = 0 Then Exit Sub
If rsCust_Type.RecordCount = 0 Then Exit Sub
cboPro.Clear
With rsCust_Type
    Do While Not .EOF
       With cboPro
            .AddItem rsCust_Type.Fields("CustType_ID")
       End With
       .MoveNext
       Loop
End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Set_Promotion_type"
End Sub

Private Sub txtCustNum_DblClick()
On Error GoTo Handle
    With frmKeyboard
        .txtInput.PasswordChar = "*"
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtCustNum.Text = .Let_Text_Input
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - txtCustNum_DblClick"
End Sub

Private Sub txtCustNum_LostFocus()
On Error GoTo Handle
    If Len(txtCustNum.Text) < 4 Then
        MsgBox " M· kh¸ch hµng Ýt nhÊt 4 ký tù"
        txtCustNum.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtCustNum_LostFocus"
End Sub

Private Sub txtCustPhone_DblClick()
On Error GoTo Handle
    With frmKeyboard
        .txtInput.PasswordChar = ""
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtCustPhone.Text = .Let_Text_Input
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - txtCustPhone_DblClick"

End Sub

Private Sub txtMaxAcc_DblClick()
On Error GoTo Handle
    With frmPhimso
         .lblTitle.Caption = "NhËp giíi h¹n c«ng nî:"
        .FormCall = 3
        .Show vbModal
        txtMaxAcc = .Return_Value
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "- txtMaxAcc_DblClick"

End Sub
