VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmDiscount_reason_list 
   Caption         =   " "
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11190
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
   ScaleHeight     =   7170
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      TabIndex        =   7
      Top             =   60
      Width           =   4935
      Begin VB.Label lblNo 
         BackColor       =   &H80000008&
         Caption         =   "Group No"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   9
         Top             =   45
         Width           =   1695
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000008&
         Caption         =   "Group Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   8
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
      Height          =   2325
      Left            =   6090
      TabIndex        =   0
      Top             =   4680
      Width           =   4980
      Begin prjTouchScreen.MyButton cmdThem 
         Height          =   945
         Left            =   60
         TabIndex        =   1
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
         MICON           =   "frmDiscount_reason_list.frx":0000
         PICN            =   "frmDiscount_reason_list.frx":001C
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
         TabIndex        =   2
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
         MICON           =   "frmDiscount_reason_list.frx":046E
         PICN            =   "frmDiscount_reason_list.frx":048A
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
         TabIndex        =   3
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
         MICON           =   "frmDiscount_reason_list.frx":09CE
         PICN            =   "frmDiscount_reason_list.frx":09EA
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
         TabIndex        =   4
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
         MICON           =   "frmDiscount_reason_list.frx":1024
         PICN            =   "frmDiscount_reason_list.frx":1040
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
         TabIndex        =   5
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
         MICON           =   "frmDiscount_reason_list.frx":167A
         PICN            =   "frmDiscount_reason_list.frx":1696
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
         TabIndex        =   6
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
         MICON           =   "frmDiscount_reason_list.frx":7930
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
      Height          =   3285
      Left            =   6060
      TabIndex        =   10
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
      TabCaption(0)   =   "Lý do chiÕt khÊu"
      TabPicture(0)   =   "frmDiscount_reason_list.frx":794C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblExpensesName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblExpensesNo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtDienGiai"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtMaCK"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.TextBox txtMaCK 
         Height          =   495
         Left            =   210
         MaxLength       =   8
         TabIndex        =   12
         Tag             =   "1"
         Top             =   810
         Width           =   1905
      End
      Begin VB.TextBox txtDienGiai 
         Height          =   495
         Left            =   180
         TabIndex        =   11
         Tag             =   "1"
         Top             =   1770
         Width           =   3735
      End
      Begin VB.Label lblExpensesNo 
         Caption         =   "M·sè"
         Height          =   255
         Left            =   210
         TabIndex        =   14
         Tag             =   "L14"
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label lblExpensesName 
         Caption         =   "DiÔn gi¶i"
         Height          =   375
         Left            =   180
         TabIndex        =   13
         Tag             =   "L15"
         Top             =   1440
         Width           =   1875
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flgCK 
      Height          =   6945
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
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
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmDiscount_reason_list"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsReasonList As New ADODB.Recordset
Dim strPath As String
Dim DescArr() As String

Private Sub cmdCapnhat_Click()
    Call UpdateDatabase
    Call LoadControl
End Sub

Private Sub cmdClose_Click()
    Set rsReasonList = Nothing
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
        txtMaCK.SetFocus
        
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "DeleteTextbox"
End Sub
Private Sub UpdateDatabase()
    On Error GoTo Handle
        With rsReasonList
            .Find "ID='" & txtMaCK.Text & _
                "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("ID") = Format(txtMaCK.Text, "00")
                .Fields("Pro_Desc") = txtDiengiai.Text
                .Update
                .Requery
            Else
                .Fields("ID") = Format(txtMaCK.Text, "00")
                .Fields("Pro_Desc") = txtDiengiai.Text
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
        With rsReasonList
            .Find "ID='" & flgCK.TextMatrix(flgCK.Row, 0) & _
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
    With rsReasonList
        .Find "ID='" & flgCK.TextMatrix(flgCK.Row, 0) & _
                "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtMaCK.Text = !ID
            txtDiengiai.Text = !Pro_Desc
            lblNo.Caption = !ID
            lblName.Caption = !Pro_Desc
        End If
    End With
    Exit Sub
    
Handle:
    MsgBox Err.Number & " " & Err.Description & vbCrLf _
    & Me.name & " flgCK_EnterCell"
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    Dim strPath As String
    Dim str As String
    
    If Check_Table_exist("Promotion_Reason") = False Then Call Create_Promotion_Reason
    'Set rsReason = Open_Table(cnData, "Promotion_Reason")
    
    str = "Select * from Promotion_Reason"
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(strPath, "100881administrator")
'    End If
    Set rsReasonList = OpenCriticalTable(str, cnData)
    Call setflgCK
    Call LockText
    Exit Sub
Handle:
    MsgBox Err.Number & " " & _
    Err.Description & Me.name & "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Handle
    Set rsReasonList = Nothing
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
        .ColWidth(1) = 6000
        .TextMatrix(0, 0) = "M·" 'DescArr(14)
        .TextMatrix(0, 1) = "DiÔn gi¶i chiÕt khÊu" 'DescArr(15)
    End With
    
    If rsReasonList Is Nothing Then Exit Sub
    If rsReasonList.State = 0 Then Exit Sub
    
    If rsReasonList.EOF And rsReasonList.BOF Then
        With flgCK
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
        End With
        Exit Sub
    End If
   flgCK.Rows = rsReasonList.RecordCount + 1
    intCount = 0
    Do While Not rsReasonList.EOF
        intCount = intCount + 1
        flgCK.TextMatrix(intCount, 0) = rsReasonList!ID
        flgCK.TextMatrix(intCount, 1) = rsReasonList!Pro_Desc
        rsReasonList.MoveNext
        
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
    
    With rsReasonList
        .Find "ID='" & !ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtMaCK.Text = !ID
            txtDiengiai.Text = !Pro_Desc
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
        cmdCapnhat.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtDiengiai_KeyPress"
End Sub






