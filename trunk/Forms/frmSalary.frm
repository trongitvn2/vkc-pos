VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSalary 
   Caption         =   "B¶ng l­¬ng"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11370
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
   ScaleHeight     =   7635
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   6
      Top             =   180
      Width           =   4935
      Begin VB.Label lblNo 
         BackColor       =   &H80000008&
         Caption         =   "Salary_ID"
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
         TabIndex        =   8
         Top             =   45
         Width           =   1695
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000008&
         Caption         =   "Salary_Name"
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
         TabIndex        =   7
         Top             =   330
         Width           =   2295
      End
   End
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
      TabIndex        =   0
      Top             =   5490
      Width           =   4980
      Begin prjTouchScreen.MyButton cmdThem 
         Height          =   705
         Left            =   60
         TabIndex        =   1
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
         MICON           =   "frmSalary.frx":0000
         PICN            =   "frmSalary.frx":001C
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
         TabIndex        =   2
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
         MICON           =   "frmSalary.frx":046E
         PICN            =   "frmSalary.frx":048A
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
         TabIndex        =   3
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
         MICON           =   "frmSalary.frx":09CE
         PICN            =   "frmSalary.frx":09EA
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
         TabIndex        =   4
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
         MICON           =   "frmSalary.frx":1024
         PICN            =   "frmSalary.frx":1040
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
         TabIndex        =   5
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
         MICON           =   "frmSalary.frx":167A
         PICN            =   "frmSalary.frx":1696
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
      TabCaption(0)   =   "B¶ng l­¬ng"
      TabPicture(0)   =   "frmSalary.frx":7930
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblExpensesName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblExpensesNo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtSalary_Name"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtSalary_ID"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtValue"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox txtValue 
         Alignment       =   1  'Right Justify
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
         Left            =   480
         TabIndex        =   15
         Tag             =   "1"
         Top             =   2820
         Width           =   2175
      End
      Begin VB.TextBox txtSalary_ID 
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
         TabIndex        =   11
         Tag             =   "1"
         Top             =   810
         Width           =   1425
      End
      Begin VB.TextBox txtSalary_Name 
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
         TabIndex        =   10
         Tag             =   "1"
         Top             =   1770
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Møc l­¬ng"
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
         TabIndex        =   16
         Tag             =   "L11"
         Top             =   2400
         Width           =   1875
      End
      Begin VB.Label lblExpensesNo 
         Caption         =   "M·"
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
      Begin VB.Label lblExpensesName 
         Caption         =   "Tªn b¶ng l­¬ng"
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
         TabIndex        =   12
         Tag             =   "L3"
         Top             =   1350
         Width           =   1875
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flgSalary 
      Height          =   7425
      Left            =   0
      TabIndex        =   14
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
Attribute VB_Name = "frmSalary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsSalary As New ADODB.Recordset
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
    Set rsSalary = Nothing
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
        txtSalary_ID.Text = GetMax_ID("Salary", "Salary_ID")
       txtSalary_Name.Text = ""
       txtValue.Text = ""
        txtSalary_ID.SetFocus
        
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "DeleteTextbox"
End Sub
Private Sub UpdateDatabase()
    On Error GoTo Handle
        With rsSalary
            .Find "Salary_ID='" & txtSalary_ID.Text & _
                "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("Salary_ID") = txtSalary_ID.Text
                .Fields("Salary_Name") = txtSalary_Name.Text
                .Fields("Salary") = txtValue.Text
                .Update
                .Requery
            Else
                MsgBox DescArr(9), vbOKOnly
                Call DeleteTextbox
                
            End If
        End With
        Call setflgSalary
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
        With rsSalary
            .Find "Salary_ID='" & flgSalary.TextMatrix(flgSalary.Row, 0) & _
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

Private Sub flgSalary_EnterCell()
    On Error GoTo Handle
    With rsSalary
        .Find "Salary_ID='" & flgSalary.TextMatrix(flgSalary.Row, 0) & _
                "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtSalary_ID.Text = !Salary_ID
            txtSalary_Name.Text = !Salary_Name
            txtValue.Text = Format(!Salary, formatNum)
            lblNo.Caption = !Salary_ID
            lblName.Caption = !Salary_Name
        End If
    End With
    Exit Sub
    
Handle:
    MsgBox Err.Number & " " & Err.Description & vbCrLf _
    & Me.name & " flgSalary_EnterCell"
End Sub

Private Sub Form_Activate()
    Dim ctrl As Control
    If cmdThem.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    DescArr = LoadLanguage(LngFile, "#03:002:")
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
    DescArr = LoadLanguage(LngFile, "#03:002:")
    str = "Select * from Salary"
    Set rsSalary = OpenCriticalTable(str, cnData)
    Call setflgSalary
    Call LockText
    Exit Sub
Handle:
    MsgBox Err.Number & " " & _
    Err.Description & Me.name & "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Handle
    Set rsSalary = Nothing
Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & "" & _
                        Me.name & " " & "form_Unload"
End Sub
Private Sub setflgSalary()
On Error GoTo errHdl
    Dim intCount    As Integer
    With flgSalary
        .Font = ".vnArial"
        .ColWidth(0) = 1500
        .ColWidth(1) = 5500
        .ColWidth(2) = 1500
        .TextMatrix(0, 0) = DescArr(2)
        .TextMatrix(0, 1) = DescArr(3)
        .TextMatrix(0, 2) = DescArr(11)
    End With
    
    If rsSalary Is Nothing Then Exit Sub
    If rsSalary.State = 0 Then Exit Sub
    
    If rsSalary.EOF And rsSalary.BOF Then
        With flgSalary
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
        End With
        Exit Sub
    End If
   flgSalary.Rows = rsSalary.RecordCount + 1
    intCount = 0
    Do While Not rsSalary.EOF
        intCount = intCount + 1
        flgSalary.TextMatrix(intCount, 0) = rsSalary!Salary_ID
        flgSalary.TextMatrix(intCount, 1) = rsSalary!Salary_Name
        flgSalary.TextMatrix(intCount, 2) = Format(rsSalary!Salary, formatNum)
        rsSalary.MoveNext
        
    Loop
'    SetColorFlexGrid flgSalary, 1, 1, flgSalary.Cols

    Call flgSalary_EnterCell
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - setflgSalary "
End Sub
Private Sub LoadControl()
    On Error GoTo Handle
    
    With rsSalary
        .Find "Salary_ID='" & !Salary_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtSalary_ID.Text = !Salary_ID
           txtSalary_Name.Text = !Salary_Name
           txtValue.Text = Format(!Salary, formatNum)
            .Requery
        End If
    End With
    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " LoadControl"
End Sub

Public Sub UnlockText()
    On Error GoTo Handle
        txtSalary_ID.Locked = False
        txtSalary_Name.Locked = False
        txtValue.Locked = False
        cmdCapnhat.Enabled = True
        txtSalary_ID.SetFocus
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub
Public Sub LockText()
    On Error GoTo Handle
        txtSalary_ID.Locked = True
        txtSalary_Name.Locked = True
        txtValue.Locked = True
        cmdCapnhat.Enabled = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub

Private Sub txtSalary_Name_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtSalary_Name.Text = .Let_Text_Input
        End With

    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtSalary_ID_DblClick "

End Sub

Private Sub txtSalary_Name_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        txtValue.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtSalary_Name_KeyPress"

End Sub

Private Sub txtSalary_ID_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtSalary_ID.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtSalary_ID_DblClick "

End Sub

Private Sub txtSalary_ID_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        txtSalary_Name.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtSalary_ID_KeyPress"
End Sub

Private Sub txtValue_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtValue.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtValue_DblClick "

End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        If cmdCapnhat.Enabled = True Then
            cmdCapnhat.SetFocus
        Else
            cmdThem.SetFocus
        End If
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtValue_KeyPress"

End Sub

Private Sub txtValue_LostFocus()
    txtValue.Text = Format(txtValue.Text, formatNum)
End Sub
