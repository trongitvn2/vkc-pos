VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmExpenses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Danh môc kho¶n chi"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExpenses.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9885
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
      TabIndex        =   9
      Top             =   5040
      Width           =   4980
      Begin prjTouchScreen.MyButton cmdThem 
         Height          =   705
         Left            =   60
         TabIndex        =   10
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
         MICON           =   "frmExpenses.frx":000C
         PICN            =   "frmExpenses.frx":0028
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
         TabIndex        =   11
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
         MICON           =   "frmExpenses.frx":047A
         PICN            =   "frmExpenses.frx":0496
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
         TabIndex        =   12
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
         MICON           =   "frmExpenses.frx":09DA
         PICN            =   "frmExpenses.frx":09F6
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
         TabIndex        =   13
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
         MICON           =   "frmExpenses.frx":1030
         PICN            =   "frmExpenses.frx":104C
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
         TabIndex        =   14
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
         MICON           =   "frmExpenses.frx":1686
         PICN            =   "frmExpenses.frx":16A2
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
      Top             =   90
      Width           =   4935
      Begin VB.Label lblName 
         BackColor       =   &H80000008&
         Caption         =   "Group Name"
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
         TabIndex        =   4
         Top             =   330
         Width           =   2295
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000008&
         Caption         =   "Group No"
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
         TabIndex        =   3
         Top             =   45
         Width           =   1695
      End
   End
   Begin TabDlg.SSTab tabGroup 
      Height          =   2805
      Left            =   4860
      TabIndex        =   5
      Top             =   1410
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
      TabCaption(0)   =   "Kho¶n chi"
      TabPicture(0)   =   "frmExpenses.frx":793C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblExpensesNo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblExpensesName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtMaChi"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtDienGiai"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox txtDienGiai 
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
         TabIndex        =   2
         Tag             =   "1"
         Top             =   1890
         Width           =   3735
      End
      Begin VB.TextBox txtMaChi 
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
         TabIndex        =   1
         Tag             =   "1"
         Top             =   1050
         Width           =   1425
      End
      Begin VB.Label lblExpensesName 
         Caption         =   "DiÔn gi¶i Chi"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Tag             =   "L15"
         Top             =   1590
         Width           =   1575
      End
      Begin VB.Label lblExpensesNo 
         Caption         =   "M· phiÕu Chi"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   6
         Tag             =   "L14"
         Top             =   750
         Width           =   1545
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flgKhoanChi 
      Height          =   6945
      Left            =   0
      TabIndex        =   8
      Top             =   30
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
Attribute VB_Name = "frmExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsKhoanchi As New ADODB.Recordset
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
    Set rsKhoanchi = Nothing
    Unload Me
End Sub

Private Sub cmdThem_Click()
On Error GoTo Handle
    If cmdThem.Caption = "&Thªm" Then
        Call UnlockText
        Call DeleteTextbox
    ElseIf cmdThem.Caption = DescArr(9) Then
        Call UnlockText
    End If
Exit Sub
Handle:
MsgBox Err.Number & " " & Err.Description & " " & _
            Me.name & " " & "cmdThem _click"

End Sub
Private Sub DeleteTextbox()
    On Error GoTo Handle
        cmdThem.Caption = DescArr(9)
        txtMaChi.Text = ""
       txtDiengiai.Text = ""
        txtMaChi.SetFocus
        
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "DeleteTextbox"
End Sub
Private Sub UpdateDatabase()
    On Error GoTo Handle
        With rsKhoanchi
            .Find "MaChi='" & txtMaChi.Text & _
                "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("MaChi") = txtMaChi.Text
                .Fields("DienGiai") = txtDiengiai.Text
                .Update
                .Requery
            Else
                MsgBox "MaChi ®· tån t¹i, vui lßng kiÓm tra l¹i hoÆc ®æi m· kh¸c!", vbOKOnly
                Call DeleteTextbox
                
            End If
        End With
        Call setflgKhoanChi
        cmdThem.Caption = "&Thªm" 'DescArr(4)
        
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "UpdateDatabase"
End Sub


Private Sub cmdXoa_Click()

    On Error GoTo Handle
    Dim ans As Integer
    If txtMaChi.Text <> "T§C" Then
        ans = MsgBox("B¹n cã ch¾c ch¨n muèn xãa danh môc nµy kh«ng?", vbYesNo)
        If ans = vbYes Then
            With rsKhoanchi
                .Find "MaChi='" & flgKhoanChi.TextMatrix(flgKhoanChi.Row, 0) & _
                        "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Or .BOF Then
                    .Delete adAffectCurrent
                    .MoveNext
                    .Requery
                End If
                Call Form_Load
            End With
        End If
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "cmdXoa_Click"

End Sub

Private Sub flgKhoanChi_EnterCell()
    On Error GoTo Handle
    With rsKhoanchi
        .Find "MaChi='" & flgKhoanChi.TextMatrix(flgKhoanChi.Row, 0) & _
                "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtMaChi.Text = !maChi
            txtDiengiai.Text = !DienGiai
            lblNo.Caption = !maChi
            lblName.Caption = !DienGiai
        End If
    End With
    Exit Sub
    
Handle:
    MsgBox Err.Number & " " & Err.Description & vbCrLf _
    & Me.name & " flgKhoanChi_EnterCell"
End Sub

Private Sub Form_Activate()
    Dim ctrl As Control
    If cmdThem.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    DescArr = LoadLanguage(LngFile, "#01:009:")
    Me.Caption = DescArr(13)
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    Dim strPath As String
    Dim str As String
    DescArr = LoadLanguage(LngFile, "#01:009:")
    str = "Select * from Expense"
    Set rsKhoanchi = OpenCriticalTable(str, cnData)
    Call setflgKhoanChi
    Call LockText
    Exit Sub
Handle:
    MsgBox Err.Number & " " & _
    Err.Description & Me.name & "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Handle
    Set rsKhoanchi = Nothing
Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & "" & _
                        Me.name & " " & "form_Unload"
End Sub
Private Sub setflgKhoanChi()
On Error GoTo errHdl
    Dim intCount    As Integer
    With flgKhoanChi
        .Font = ".vnArial"
        .ColWidth(0) = 2500
        .ColWidth(1) = 7500
        .TextMatrix(0, 0) = DescArr(14)
        .TextMatrix(0, 1) = DescArr(15)
    End With
    
    If rsKhoanchi Is Nothing Then Exit Sub
    If rsKhoanchi.State = 0 Then Exit Sub
    
    If rsKhoanchi.EOF And rsKhoanchi.BOF Then
        With flgKhoanChi
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
        End With
        Exit Sub
    End If
   flgKhoanChi.Rows = rsKhoanchi.RecordCount + 1
    intCount = 0
    Do While Not rsKhoanchi.EOF
        intCount = intCount + 1
        flgKhoanChi.TextMatrix(intCount, 0) = rsKhoanchi!maChi
        flgKhoanChi.TextMatrix(intCount, 1) = rsKhoanchi!DienGiai
        rsKhoanchi.MoveNext
        
    Loop
'    SetColorFlexGrid flgKhoanChi, 1, 1, flgKhoanChi.Cols

    Call flgKhoanChi_EnterCell
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - setflgKhoanChi "
End Sub
Private Sub LoadControl()
    On Error GoTo Handle
    
    With rsKhoanchi
        .Find "MaChi='" & !maChi & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtMaChi.Text = !maChi
           txtDiengiai.Text = !DienGiai
            .Requery
        End If
    End With
    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " LoadControl"
End Sub

Public Sub UnlockText()
    On Error GoTo Handle
        txtMaChi.Locked = False
        txtDiengiai.Locked = False
        cmdCapnhat.Enabled = True
        txtMaChi.SetFocus
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub
Public Sub LockText()
    On Error GoTo Handle
        txtMaChi.Locked = True
        txtDiengiai.Locked = True
        cmdCapnhat.Enabled = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub

Private Sub txtDienGiai_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtDiengiai.Text = .Let_Text_Input
       End With

    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtMaThu_DblClick "

End Sub

Private Sub txtDiengiai_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        cmdCapnhat.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtDienGiai_KeyPress"

End Sub

Private Sub txtMaChi_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtMaChi.Text = .Let_Text_Input
       End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtMaThu_DblClick "

End Sub

Private Sub txtMaChi_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        txtDiengiai.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtMaChi_KeyPress"
End Sub
