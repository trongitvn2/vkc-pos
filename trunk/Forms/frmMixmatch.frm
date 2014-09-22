VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMixmatch 
   Caption         =   "Danh s¸ch gi¶m %"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11115
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
   ScaleHeight     =   7110
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   2205
      Left            =   6090
      TabIndex        =   3
      Top             =   4770
      Width           =   4980
      Begin prjTouchScreen.MyButton cmdThem 
         Height          =   945
         Left            =   60
         TabIndex        =   4
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "&Thªm"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMixmatch.frx":0000
         PICN            =   "frmMixmatch.frx":001C
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
         Height          =   945
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "&CËp nhËt"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMixmatch.frx":046E
         PICN            =   "frmMixmatch.frx":048A
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
         Height          =   945
         Left            =   3330
         TabIndex        =   6
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "&Xãa"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMixmatch.frx":09CE
         PICN            =   "frmMixmatch.frx":09EA
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
         Height          =   945
         Left            =   1680
         TabIndex        =   7
         Top             =   1260
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "&Gióp ®ì"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMixmatch.frx":1024
         PICN            =   "frmMixmatch.frx":1040
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
         Cancel          =   -1  'True
         Height          =   945
         Left            =   3330
         TabIndex        =   8
         Top             =   1260
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "&§ãng"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMixmatch.frx":167A
         PICN            =   "frmMixmatch.frx":1696
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdEdit 
         Height          =   945
         Left            =   60
         TabIndex        =   17
         Top             =   1260
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "Söa ch÷a"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMixmatch.frx":7930
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
      Left            =   6090
      ScaleHeight     =   645
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   60
      Width           =   4935
      Begin VB.Label lblName 
         BackColor       =   &H80000008&
         Caption         =   "Group Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   330
         Width           =   2295
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000008&
         Caption         =   "Group No"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   45
         Width           =   1695
      End
   End
   Begin TabDlg.SSTab tabGroup 
      Height          =   3285
      Left            =   6060
      TabIndex        =   9
      Top             =   1380
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   5794
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
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
      TabCaption(0)   =   "ChiÕt khÊu"
      TabPicture(0)   =   "frmMixmatch.frx":794C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblExpensesNo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblExpensesName"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtMaCK"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtDienGiai"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtValue"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      Begin VB.TextBox txtValue 
         Height          =   495
         Left            =   200
         TabIndex        =   15
         Tag             =   "1"
         Top             =   2610
         Width           =   2175
      End
      Begin VB.TextBox txtDienGiai 
         Height          =   495
         Left            =   180
         TabIndex        =   11
         Tag             =   "1"
         Top             =   1770
         Width           =   3735
      End
      Begin VB.TextBox txtMaCK 
         Height          =   495
         Left            =   210
         MaxLength       =   8
         TabIndex        =   10
         Tag             =   "1"
         Top             =   810
         Width           =   1905
      End
      Begin VB.Label Label1 
         Caption         =   "Gi¸ trÞ"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Tag             =   "L15"
         Top             =   2280
         Width           =   1875
      End
      Begin VB.Label lblExpensesName 
         Caption         =   "Tªn chiÕt khÊu"
         Height          =   375
         Left            =   180
         TabIndex        =   13
         Tag             =   "L15"
         Top             =   1440
         Width           =   1875
      End
      Begin VB.Label lblExpensesNo 
         Caption         =   "M· chiÕt khÊu"
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Tag             =   "L14"
         Top             =   480
         Width           =   1755
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flgCK 
      Height          =   6945
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   12250
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
Attribute VB_Name = "frmMixmatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsMixmatch As New ADODB.Recordset
Dim strPath As String
Dim DescArr() As String
Dim wYear As Double
Dim wMonth As Integer

Private Sub cmdCapnhat_Click()
    Call UpdateDatabase
    Call LoadControl
End Sub

Private Sub cmdClose_Click()
    Set rsMixmatch = Nothing
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    Call UnlockText
End Sub

Private Sub cmdThem_Click()
On Error GoTo Handle
    
    Call UnlockText
    Call DeleteTextbox

Exit Sub
Handle:
MsgBox Err.Number & " " & Err.Description & " " & _
            Me.name & " " & "cmdThem _click"

End Sub
Private Sub DeleteTextbox()
    On Error GoTo Handle
        'cmdThem.Caption = DescArr(9)
        txtMaCK.Text = ""
        txtDiengiai.Text = ""
        txtValue.Text = ""
        txtMaCK.SetFocus
        
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "DeleteTextbox"
End Sub
Private Sub UpdateDatabase()
    On Error GoTo Handle
        With rsMixmatch
            .Find "Pro_ID='" & txtMaCK.Text & _
                "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("Pro_ID") = Format(txtMaCK.Text, "00")
                .Fields("Pro_Name") = txtDiengiai.Text
                .Fields("Pro_Value") = txtValue.Text
                .Update
                .Requery
            Else
                .Fields("Pro_ID") = Format(txtMaCK.Text, "00")
                .Fields("Pro_Name") = txtDiengiai.Text
                .Fields("Pro_Value") = txtValue.Text
                .Update
                .Requery
            End If
        End With
        Call setflgCK
        cmdThem.Caption = "&Thªm" 'DescArr(4)
        
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
        With rsMixmatch
            .Find "Pro_ID='" & flgCK.TextMatrix(flgCK.Row, 0) & _
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

Private Sub flgCK_EnterCell()
    On Error GoTo Handle
    With rsMixmatch
        .Find "Pro_ID='" & flgCK.TextMatrix(flgCK.Row, 0) & _
                "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtMaCK.Text = !Pro_ID
            txtDiengiai.Text = !Pro_Name
            txtValue.Text = !Pro_Value
            lblNo.Caption = !Pro_ID
            lblName.Caption = !Pro_Name
        End If
    End With
    Exit Sub
    
Handle:
    MsgBox Err.Number & " " & Err.Description & vbCrLf _
    & Me.name & " flgCK_EnterCell"
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    Dim str As String
    str = "Select * from Promotion"

    Set rsMixmatch = OpenCriticalTable(str, cnData)
    Call setflgCK
    Call LockText
    Exit Sub
Handle:
    MsgBox Err.Number & " " & _
    Err.Description & Me.name & "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Handle
    Set rsMixmatch = Nothing
Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & "" & _
                        Me.name & " " & "form_Unload"
End Sub
Private Sub setflgCK()
On Error GoTo errHdl
    Dim intCount    As Integer
    With flgCK
        .Font = ".vnArial"
        .ColWidth(0) = 1000
        .ColWidth(1) = 3500
        .ColWidth(2) = 1500
        .TextMatrix(0, 0) = "M· CK" 'DescArr(14)
        .TextMatrix(0, 1) = "Tªn CK" 'DescArr(15)
        .TextMatrix(0, 2) = "GÝa trÞ"
    End With
    
    If rsMixmatch Is Nothing Then Exit Sub
    If rsMixmatch.State = 0 Then Exit Sub
    
    If rsMixmatch.EOF And rsMixmatch.BOF Then
        With flgCK
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
        End With
        Exit Sub
    End If
   flgCK.Rows = rsMixmatch.RecordCount + 1
    intCount = 0
    Do While Not rsMixmatch.EOF
        intCount = intCount + 1
        flgCK.TextMatrix(intCount, 0) = rsMixmatch!Pro_ID
        flgCK.TextMatrix(intCount, 1) = rsMixmatch!Pro_Name
        flgCK.TextMatrix(intCount, 2) = rsMixmatch!Pro_Value
        rsMixmatch.MoveNext
        
    Loop
'    SetColorFlexGrid flgCK, 1, 1, flgCK.Cols

    Call flgCK_EnterCell
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - setflgCK "
End Sub
Private Sub LoadControl()
    On Error GoTo Handle
    
    With rsMixmatch
        .Find "Pro_ID='" & !Pro_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtMaCK.Text = !Pro_ID
            txtDiengiai.Text = !Pro_Name
            txtValue.Text = !Pro_Value
            .Requery
        End If
    End With
    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " LoadControl"
End Sub

Public Sub UnlockText()
    On Error GoTo Handle
        txtMaCK.Locked = False
        txtDiengiai.Locked = False
        txtValue.Locked = False
        cmdCapnhat.Enabled = True
        txtMaCK.SetFocus
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub
Public Sub LockText()
    On Error GoTo Handle
        txtMaCK.Locked = True
        txtDiengiai.Locked = True
        txtValue.Locked = True
        'cmdCapnhat.Enabled = False
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

Private Sub txtValue_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        cmdCapnhat.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtValue_KeyPress"

End Sub

Private Sub txtMaCK_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtMaCK.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtMaThu_DblClick "

End Sub

Private Sub txtMaCK_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        txtDiengiai.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtMaCK_KeyPress"
End Sub
Private Sub txtDiengiai_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        txtValue.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtDiengiai_KeyPress"
End Sub

Private Sub txtValue_LostFocus()
    If CDbl("0" & txtValue.Text) > 100 Then
        MsgBox "Gi¸ trÞ gi¶m kh«ng ®­îc lín h¬n 100"
        txtValue.SetFocus
        txtValue.SelStart = 0
        txtValue.SelLength = 999
        cmdCapnhat.Enabled = False
    Else
        cmdCapnhat.Enabled = True
    End If
End Sub
