VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmJobcode 
   Caption         =   "B¶ng c«ng viÖc"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11400
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
   ScaleHeight     =   7560
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmCmd 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1965
      Left            =   6330
      TabIndex        =   3
      Top             =   5490
      Width           =   4980
      Begin prjTouchScreen.MyButton cmdThem 
         Height          =   705
         Left            =   60
         TabIndex        =   4
         Tag             =   "L4"
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Thªm"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
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
         MICON           =   "frmJobcode.frx":0000
         PICN            =   "frmJobcode.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdCapnhat 
         Height          =   705
         Left            =   1680
         TabIndex        =   5
         Tag             =   "L5"
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&CËp nhËt"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
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
         MICON           =   "frmJobcode.frx":046E
         PICN            =   "frmJobcode.frx":048A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdXoa 
         Height          =   705
         Left            =   3330
         TabIndex        =   6
         Tag             =   "L6"
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Xãa"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
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
         MICON           =   "frmJobcode.frx":09CE
         PICN            =   "frmJobcode.frx":09EA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdHelp 
         Height          =   705
         Left            =   900
         TabIndex        =   7
         Tag             =   "L10"
         Top             =   1020
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Gióp ®ì"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
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
         MICON           =   "frmJobcode.frx":1024
         PICN            =   "frmJobcode.frx":1040
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   705
         Left            =   2640
         TabIndex        =   8
         Tag             =   "L8"
         Top             =   1020
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&§ãng"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
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
         MICON           =   "frmJobcode.frx":167A
         PICN            =   "frmJobcode.frx":1696
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   6450
      ScaleHeight     =   645
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   180
      Width           =   4935
      Begin VB.Label lblName 
         BackColor       =   &H80000008&
         Caption         =   "Jobcode_Name"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   330
         Width           =   2295
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000008&
         Caption         =   "JobCode_ID"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   45
         Width           =   1695
      End
   End
   Begin TabDlg.SSTab tabGroup 
      Height          =   3765
      Left            =   6420
      TabIndex        =   9
      Top             =   1380
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   6641
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "B¶ng c«ng viÖc"
      TabPicture(0)   =   "frmJobcode.frx":7930
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblExpensesNo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblExpensesName"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cboSalary_ID"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtJobcode_ID"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtJobcode_Name"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtSalaryName"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.TextBox txtSalaryName 
         Height          =   495
         Left            =   2040
         TabIndex        =   17
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox txtJobcode_Name 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   540
         TabIndex        =   11
         Tag             =   "1"
         Top             =   1770
         Width           =   3735
      End
      Begin VB.TextBox txtJobcode_ID 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   570
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   10
         Tag             =   "1"
         Top             =   810
         Width           =   1425
      End
      Begin MSForms.ComboBox cboSalary_ID 
         Height          =   495
         Left            =   600
         TabIndex        =   16
         Top             =   2760
         Width           =   1335
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2355;873"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label lblExpensesName 
         Caption         =   "Tªn C«ng viÖc"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   14
         Tag             =   "L3"
         Top             =   1350
         Width           =   1875
      End
      Begin VB.Label lblExpensesNo 
         Caption         =   "M· c«ng viÖc"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   210
         TabIndex        =   13
         Tag             =   "L2"
         Top             =   390
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "B¶ng l­¬ng"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Tag             =   "L11"
         Top             =   2400
         Width           =   1875
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flgJobcode 
      Height          =   7425
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   13097
      _Version        =   393216
      Cols            =   3
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      TextStyleFixed  =   3
      GridLinesFixed  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmJobcode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsJobcode As New ADODB.Recordset
Dim rsSalary As New ADODB.Recordset
Dim strPath As String
Dim DescArr() As String
Dim wYear As Double
Dim wMonth As Integer

Private Sub cboSalary_ID_Change()
    Call cboSalary_ID_Click
End Sub

Private Sub cboSalary_ID_Click()
On Error GoTo Handle
    With rsSalary
        .Find "Salary_ID='" & Trim(cboSalary_ID.Text) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtSalaryName.Text = .Fields("Salary_Name")
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cboSalary_ID_Click"
End Sub

Private Sub cmdCapnhat_Click()
    If Check_condition = False Then Exit Sub
    Call UpdateDatabase
    Call LoadControl
    If cmdThem.Enabled = True Then
        cmdThem.SetFocus
    Else
        cmdThem.Enabled = True
        cmdThem.SetFocus
    End If
End Sub

Private Sub cmdClose_Click()
    Set rsJobcode = Nothing
    Unload Me
End Sub

Private Sub cmdThem_Click()
On Error GoTo Handle
    If cmdThem.Caption = DescArr(4) Then
        Call UnlockText
        Call DeleteTextbox
        Call Set_Salary
    ElseIf cmdThem.Caption = DescArr(7) Then
        Call UnlockText
    End If
Exit Sub
Handle:
MsgBox Err.Number & " " & Err.Description & " " & _
            Me.name & " " & "cmdThem _click"

End Sub
Private Sub DeleteTextbox()
    On Error GoTo Handle
        cmdThem.Caption = DescArr(7)
        txtJobcode_ID.Text = GetMax_ID("JobCode", "JobCode_ID")
       txtJobcode_Name.Text = ""
       cboSalary_ID.Text = ""
       txtSalaryName.Text = ""
        txtJobcode_Name.SetFocus
        
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "DeleteTextbox"
End Sub
Private Sub UpdateDatabase()
    On Error GoTo Handle
        With rsJobcode
            .Find "Jobcode_ID='" & txtJobcode_ID.Text & _
                "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("Jobcode_ID") = txtJobcode_ID.Text
                .Fields("Jobcode_Name") = txtJobcode_Name.Text
                .Fields("Salary_ID") = Trim(cboSalary_ID.Text)
                .Update
                .Requery
            Else
                MsgBox DescArr(9), vbOKOnly
                Call DeleteTextbox
                
            End If
        End With
        Call setflgJobcode
        cmdThem.Caption = DescArr(4)
        
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "UpdateDatabase"
End Sub


Private Sub cmdXoa_Click()

    On Error GoTo Handle
    Dim ans As Integer
    ans = MsgBox("B¹n cã ch¾c ch¨n muèn xãa danh môc nµy kh«ng?", vbYesNo)
    If ans = vbYes Then
        With rsJobcode
            .Find "Jobcode_ID='" & flgJobcode.TextMatrix(flgJobcode.Row, 0) & _
                    "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Or .BOF Then
                .Delete adAffectCurrent
                .MoveNext
                .Requery
            End If
            Call Form_Load
        End With
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "cmdXoa_Click"

End Sub

Private Sub flgJobcode_EnterCell()
    On Error GoTo Handle
    With rsJobcode
        .Find "Jobcode_ID='" & flgJobcode.TextMatrix(flgJobcode.Row, 0) & _
                "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtJobcode_ID.Text = !JobCode_ID
            txtJobcode_Name.Text = !Jobcode_Name
            cboSalary_ID.Text = !Salary_ID
            lblNo.Caption = !JobCode_ID
            lblName.Caption = !Jobcode_Name
        End If
    End With
    Exit Sub
    
Handle:
    MsgBox Err.Number & " " & Err.Description & vbCrLf _
    & Me.name & " flgJobcode_EnterCell"
End Sub

Private Sub Form_Activate()
    Dim ctrl As Control
    If cmdThem.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    DescArr = LoadLanguage(LngFile, "#03:003:")
    Me.Caption = DescArr(1)
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    Dim strPath As String
    Dim str As String
    DescArr = LoadLanguage(LngFile, "#03:003:")
    str = "Select * from Jobcode"
    Set rsJobcode = OpenCriticalTable(str, cnData)
    Set rsSalary = Open_Table(cnData, "Salary")
    Call setflgJobcode
    Call LockText
    Exit Sub
Handle:
    MsgBox Err.Number & " " & _
    Err.Description & Me.name & "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Handle
    Set rsJobcode = Nothing
Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & "" & _
                        Me.name & " " & "form_Unload"
End Sub
Private Sub setflgJobcode()
On Error GoTo errHdl
    Dim intCount    As Integer
    With flgJobcode
        .Font = ".vnArial"
        .ColWidth(0) = 1500
        .ColWidth(1) = 5500
        .ColWidth(2) = 1500
        .TextMatrix(0, 0) = DescArr(2)
        .TextMatrix(0, 1) = DescArr(3)
        .TextMatrix(0, 2) = DescArr(11)
    End With
    
    If rsJobcode Is Nothing Then Exit Sub
    If rsJobcode.State = 0 Then Exit Sub
    
    If rsJobcode.EOF And rsJobcode.BOF Then
        With flgJobcode
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
        End With
        Exit Sub
    End If
   flgJobcode.Rows = rsJobcode.RecordCount + 1
    intCount = 0
    Do While Not rsJobcode.EOF
        intCount = intCount + 1
        flgJobcode.TextMatrix(intCount, 0) = rsJobcode!JobCode_ID
        flgJobcode.TextMatrix(intCount, 1) = rsJobcode!Jobcode_Name
        flgJobcode.TextMatrix(intCount, 2) = rsJobcode!Salary_ID
        rsJobcode.MoveNext
        
    Loop
'    SetColorFlexGrid flgJobcode, 1, 1, flgJobcode.Cols

    Call flgJobcode_EnterCell
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - setflgJobcode "
End Sub
Private Sub LoadControl()
    On Error GoTo Handle
    
    With rsJobcode
        .Find "Jobcode_ID='" & !JobCode_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtJobcode_ID.Text = !JobCode_ID
           txtJobcode_Name.Text = !Jobcode_Name
           cboSalary_ID.Text = !Salary_ID
            .Requery
        End If
    End With
    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " LoadControl"
End Sub

Public Sub UnlockText()
    On Error GoTo Handle
        txtJobcode_ID.Locked = False
        txtJobcode_Name.Locked = False
        cboSalary_ID.Locked = False
        txtSalaryName.Locked = True
        cmdCapnhat.Enabled = True
        txtJobcode_ID.SetFocus
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub
Public Sub LockText()
    On Error GoTo Handle
        txtJobcode_ID.Locked = True
        txtJobcode_Name.Locked = True
        txtSalaryName.Locked = True
        cboSalary_ID.Locked = True
        cmdCapnhat.Enabled = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub

Private Sub txtJobcode_Name_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtJobcode_Name.Text = .Let_Text_Input
        End With

    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtJobcode_ID_DblClick "

End Sub

Private Sub txtJobcode_Name_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        cboSalary_ID.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtJobcode_Name_KeyPress"

End Sub

Private Sub txtJobcode_ID_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            frmKeyboard.Show vbModal
            txtJobcode_ID.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtJobcode_ID_DblClick "

End Sub

Private Sub txtJobcode_ID_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        txtJobcode_Name.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtJobcode_ID_KeyPress"
End Sub


Public Sub Set_Salary()
On Error GoTo Handle
    With cboSalary_ID
        .Clear
        If rsSalary.State = 1 And rsSalary.RecordCount > 0 Then
            rsSalary.MoveFirst
        Else
            MsgBox "Ch­a cã b¶ng l­¬ng cho c«ng viÖc"
            Exit Sub
        End If
        Do While Not rsSalary.EOF
            .AddItem rsSalary.Fields("Salary_ID")
        rsSalary.MoveNext
        Loop
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Set_Salary"
End Sub

Public Function Check_condition() As Boolean
On Error GoTo Handle
    If txtJobcode_ID.Text = "" Then
        MsgBox "M· c«ng viÖc kh«ng ®­îc bá trèng"
        Check_condition = False
    Else
        Check_condition = True
    End If
    If cboSalary_ID.Text = "" Then
        MsgBox "B¹n ph¶i nèi c«ng viÖc víi b¶ng l­¬ng"
        Check_condition = False
    Else
        Check_condition = True
    End If
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Check_condition"

End Function
