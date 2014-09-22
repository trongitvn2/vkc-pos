VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDept 
   Caption         =   "Phßng ban"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
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
   ScaleHeight     =   7080
   ScaleWidth      =   9960
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
      Left            =   4890
      TabIndex        =   3
      Top             =   5010
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
         MICON           =   "frmDept.frx":0000
         PICN            =   "frmDept.frx":001C
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
         MICON           =   "frmDept.frx":046E
         PICN            =   "frmDept.frx":048A
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
         MICON           =   "frmDept.frx":09CE
         PICN            =   "frmDept.frx":09EA
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
         MICON           =   "frmDept.frx":1024
         PICN            =   "frmDept.frx":1040
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
         MICON           =   "frmDept.frx":167A
         PICN            =   "frmDept.frx":1696
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
      Left            =   4890
      ScaleHeight     =   645
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   60
      Width           =   4935
      Begin VB.Label lblName 
         BackColor       =   &H80000008&
         Caption         =   "Dept_ Name"
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
         Caption         =   "Dept_ID"
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
      Height          =   2805
      Left            =   4860
      TabIndex        =   9
      Top             =   1380
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   4948
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
      TabCaption(0)   =   "Phßng ban"
      TabPicture(0)   =   "frmDept.frx":7930
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblExpensesNo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblExpensesName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtDeptID"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtDeptName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox txtDeptName 
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
         Left            =   180
         TabIndex        =   11
         Tag             =   "1"
         Top             =   2010
         Width           =   3735
      End
      Begin VB.TextBox txtDeptID 
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
         Left            =   210
         MaxLength       =   8
         TabIndex        =   10
         Tag             =   "1"
         Top             =   1050
         Width           =   1425
      End
      Begin VB.Label lblExpensesName 
         Caption         =   "Tªn Phßng ban"
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
         TabIndex        =   13
         Tag             =   "L3"
         Top             =   1590
         Width           =   1875
      End
      Begin VB.Label lblExpensesNo 
         Caption         =   "M· PB"
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
         TabIndex        =   12
         Tag             =   "L2"
         Top             =   630
         Width           =   1755
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flgPhongban 
      Height          =   6945
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   12250
      _Version        =   393216
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
Attribute VB_Name = "frmDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsPhongban As New ADODB.Recordset
Dim strPath As String
Dim DescArr() As String
Dim wYear As Double
Dim wMonth As Integer

Private Sub cmdCapnhat_Click()
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
    Set rsPhongban = Nothing
    Unload Me
End Sub

Private Sub cmdThem_Click()
On Error GoTo Handle
    If cmdThem.Caption = DescArr(4) Then
        Call UnlockText
        Call DeleteTextbox
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
        txtDeptID.Text = ""
       txtDeptName.Text = ""
        txtDeptID.SetFocus
        
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "DeleteTextbox"
End Sub
Private Sub UpdateDatabase()
    On Error GoTo Handle
        With rsPhongban
            .Find "Dept_ID='" & txtDeptID.Text & _
                "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("Dept_ID") = txtDeptID.Text
                .Fields("Dept_Name") = txtDeptName.Text
                .Update
                .Requery
            Else
                MsgBox DescArr(9), vbOKOnly
                Call DeleteTextbox
                
            End If
        End With
        Call setflgPhongban
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
        With rsPhongban
            .Find "Dept_ID='" & flgPhongban.TextMatrix(flgPhongban.Row, 0) & _
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

Private Sub flgPhongban_EnterCell()
    On Error GoTo Handle
    With rsPhongban
        .Find "Dept_ID='" & flgPhongban.TextMatrix(flgPhongban.Row, 0) & _
                "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtDeptID.Text = !Dept_ID
            txtDeptName.Text = !Dept_Name
            lblNo.Caption = !Dept_ID
            lblName.Caption = !Dept_Name
        End If
    End With
    Exit Sub
    
Handle:
    MsgBox Err.Number & " " & Err.Description & vbCrLf _
    & Me.name & " flgPhongban_EnterCell"
End Sub

Private Sub Form_Activate()
    Dim ctrl As Control
    If cmdThem.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    DescArr = LoadLanguage(LngFile, "#03:001:")
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
    DescArr = LoadLanguage(LngFile, "#03:001:")
    str = "Select * from Company_Dept"
    Set rsPhongban = OpenCriticalTable(str, cnData)
    Call setflgPhongban
    Call LockText
    Exit Sub
Handle:
    MsgBox Err.Number & " " & _
    Err.Description & Me.name & "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Handle
    Set rsPhongban = Nothing
Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & "" & _
                        Me.name & " " & "form_Unload"
End Sub
Private Sub setflgPhongban()
On Error GoTo errHdl
    Dim intCount    As Integer
    With flgPhongban
        .Font = ".vnArial"
        .ColWidth(0) = 2500
        .ColWidth(1) = 7500
        .TextMatrix(0, 0) = DescArr(2)
        .TextMatrix(0, 1) = DescArr(3)
    End With
    
    If rsPhongban Is Nothing Then Exit Sub
    If rsPhongban.State = 0 Then Exit Sub
    
    If rsPhongban.EOF And rsPhongban.BOF Then
        With flgPhongban
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
        End With
        Exit Sub
    End If
   flgPhongban.Rows = rsPhongban.RecordCount + 1
    intCount = 0
    Do While Not rsPhongban.EOF
        intCount = intCount + 1
        flgPhongban.TextMatrix(intCount, 0) = rsPhongban!Dept_ID
        flgPhongban.TextMatrix(intCount, 1) = rsPhongban!Dept_Name
        rsPhongban.MoveNext
        
    Loop
'    SetColorFlexGrid flgPhongban, 1, 1, flgPhongban.Cols

    Call flgPhongban_EnterCell
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - setflgPhongban "
End Sub
Private Sub LoadControl()
    On Error GoTo Handle
    
    With rsPhongban
        .Find "Dept_ID='" & !Dept_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtDeptID.Text = !Dept_ID
           txtDeptName.Text = !Dept_Name
            .Requery
        End If
    End With
    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " LoadControl"
End Sub

Public Sub UnlockText()
    On Error GoTo Handle
        txtDeptID.Locked = False
        txtDeptName.Locked = False
        cmdCapnhat.Enabled = True
        txtDeptID.SetFocus
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub
Public Sub LockText()
    On Error GoTo Handle
        txtDeptID.Locked = True
        txtDeptName.Locked = True
        cmdCapnhat.Enabled = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub

Private Sub txtDeptName_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtDeptName.Text = .Let_Text_Input
       End With

    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtDept_ID_DblClick "

End Sub

Private Sub txtDeptName_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        cmdCapnhat.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtDeptName_KeyPress"

End Sub

Private Sub txtDeptID_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtDeptID.Text = .Let_Text_Input
       End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtDept_ID_DblClick "

End Sub

Private Sub txtDeptID_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        txtDeptName.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtDeptID_KeyPress"
End Sub




