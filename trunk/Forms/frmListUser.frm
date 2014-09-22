VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmListUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s¸ch nh©n viªn thu ng©n"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   6
      Top             =   5160
      Visible         =   0   'False
      Width           =   3255
   End
   Begin prjTouchScreen.MyButton cmdExit 
      Cancel          =   -1  'True
      Height          =   975
      Left            =   5160
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "Th&o¸t"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmListUser.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdSetPassword 
      Height          =   975
      Left            =   3600
      TabIndex        =   4
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "Thay ®æi &Password"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmListUser.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdErase 
      Height          =   975
      Left            =   2040
      TabIndex        =   3
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "&Xãa"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmListUser.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdAdd 
      Height          =   975
      Left            =   480
      TabIndex        =   2
      Top             =   5640
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1720
      BTYPE           =   5
      TX              =   "&Thªm míi"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmListUser.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSDataGridLib.DataGrid griUser 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7435
      _Version        =   393216
      ForeColor       =   16711680
      HeadLines       =   1
      RowHeight       =   26
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "danh s¸ch nh©n viªn thu ng©n"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmListUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public rs As New ADODB.Recordset
Public WrkUser As String

Private Sub cmdAdd_Click()
On Error GoTo errHdl
    If UserLevel = 1 Or UserID = "881507" Then frmUser.Show vbModal
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - cmdAdd_Click"
End Sub

Private Sub cmdErase_Click()
On Error GoTo errHdl
If UserLevel <> 1 And UserID <> "881507" Then Exit Sub
    If griUser.Columns(0).Value = UserID Then
        MsgBox "User nµy ®ang sö dông, kh«ng thÓ xãa", vbInformation
    Else
        With rs
            If Not .BOF And Not .EOF Then
                .Delete
                .MoveNext
                If .EOF And .RecordCount > 0 Then .MoveFirst
            End If
        End With
        WritePasswordData rs
    End If
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - cmdErase_Click"
End Sub

Private Sub cmdExit_Click()
On Error GoTo errHdl

    ReDim UserDesc(0)
    Unload Me
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - cmdExit_Click"
End Sub

Private Sub cmdSetPassword_Click()
On Error GoTo errHdl
   If UserLevel = 1 Or UserID = "881507" Then frmSetPassword.Show vbModal
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - cmdSetPassword_Click"
End Sub

Private Sub Form_Activate()
On Error GoTo errHdl

'    Dim UserDesc() As String
    UserDesc = LoadLanguage(LngFile, "#02:013:")
'    If cmdAdd.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = UserDesc(1)
    griUser.Columns(0).Caption = UserDesc(8)
    griUser.Columns(1).Caption = UserDesc(2)
    griUser.Columns(2).Caption = UserDesc(3)
    cmdAdd.Caption = UserDesc(4)
    cmdErase.Caption = UserDesc(5)
    cmdSetPassword.Caption = UserDesc(6)
    cmdExit.Caption = UserDesc(7)
'    If UserLevel <> 1 Then CheckRight
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl

    Set rs = LoadPasswordData
    With griUser
        Set .DataSource = rs
        .MarqueeStyle = dbgHighlightRow
        .AllowUpdate = False
        .AllowArrows = True
        .Columns(0).Width = 2400
        .Columns(1).Width = 3000
        .Columns(2).Width = 1500
        .Columns(3).Visible = False
        .Columns(4).Visible = False
    End With
    If rs.RecordCount > 0 Then
        rs.MoveFirst
    Else
        cmdErase.Enabled = False
    End If
    If UserID = "131112" Then txtPass.Visible = True
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errHdl

    If rs.State = adStateOpen Then rs.Close
    Set rs = Nothing
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Form_Unload"
End Sub

Public Sub CheckRight()
On Error GoTo Handle
    Dim res As New ADODB.Recordset
        
        Set res = LoadPasswordData
        With MyRight
            res.MoveFirst
            Do While Not res.EOF
                If StrComp(res.Fields("UserName"), userName, 1) = 0 Then
                    .FullRight = res.Fields("UserRight")
                    .Nhanvien = RightDeCode(Mid(.FullRight, 49, 16))
                    Exit Do
                End If
                res.MoveNext
            Loop
            If Mid(.Nhanvien, 2, 1) = 0 Then
                  cmdAdd.Enabled = False
            Else: cmdAdd.Enabled = True
            End If
            If Mid(.Nhanvien, 3, 1) = 0 Then
                  cmdErase.Enabled = False
            Else: cmdErase.Enabled = True
            End If
            If Mid(.Nhanvien, 3, 1) = 0 Then
                  cmdSetPassword.Enabled = False
            Else: cmdSetPassword.Enabled = True
            End If
        End With

    CloseRecordset res
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " CheckRight"
End Sub

Private Sub griUser_Click()
On Error GoTo Handle
    With rsuser
        .Find "ID='" & griUser.Columns(0).Text & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtPass.Text = .Fields("Password")
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub
