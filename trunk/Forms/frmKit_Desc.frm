VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmKit_Desc 
   Caption         =   "Chÿ d…n b’p"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10770
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
   Icon            =   "frmKit_Desc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11055
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   11055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      Begin prjTouchScreen.MyButton cmdKeyboard 
         Height          =   735
         Left            =   9480
         TabIndex        =   3
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         BTYPE           =   1
         TX              =   ""
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
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmKit_Desc.frx":000C
         PICN            =   "frmKit_Desc.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid flgKit_Desc 
         Height          =   10095
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   17806
         _Version        =   393216
         BackColorFixed  =   -2147483628
         BackColorBkg    =   16777215
         WordWrap        =   -1  'True
         TextStyleFixed  =   1
         FocusRect       =   0
         HighLight       =   2
         GridLinesFixed  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial NarrowH"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtKit_Desc 
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   8175
      End
      Begin prjTouchScreen.MyButton cmdClear 
         Height          =   735
         Left            =   8400
         TabIndex        =   4
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1296
         BTYPE           =   1
         TX              =   "CLR"
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
         BCOL            =   8438015
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmKit_Desc.frx":047A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdOK 
         Height          =   1335
         Left            =   8400
         TabIndex        =   5
         Top             =   1080
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2355
         BTYPE           =   4
         TX              =   "&ßÂng ˝"
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmKit_Desc.frx":0496
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   1335
         Left            =   8400
         TabIndex        =   6
         Top             =   2640
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   2355
         BTYPE           =   4
         TX              =   "&Tho∏t"
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
         BCOL            =   14737632
         BCOLO           =   14737632
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmKit_Desc.frx":04B2
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
End
Attribute VB_Name = "frmKit_Desc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCookingMessage As New ADODB.Recordset
Dim strKit_Desc As String
Dim isOK As Boolean

Private Sub cmdCancel_Click()
    isOK = False
    Unload Me
End Sub

Private Sub cmdclear_Click()
    txtKit_Desc.Text = ""
End Sub

Private Sub cmdKeyboard_Click()
On Error GoTo Handle:
        With frmKeyboard
            .FormCallkeyboard = "Other"
            .txtInput.PasswordChar = ""
            .txtInput.Text = txtKit_Desc.Text
            .txtInput.SelStart = 0
            .txtInput.SelLength = 9999
            .Show vbModal
            txtKit_Desc.Text = .Let_Text_Input
            
        End With
        Call cmdOK_Click
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdKeyboard_Click "
End Sub

Private Sub cmdOK_Click()
On Error GoTo Handle
    isOK = True
    If txtKit_Desc.Text = "" Then
        MsgBox "vui lﬂng ch‰n chÿ d…n !"
    Else
        strKit_Desc = txtKit_Desc.Text
    End If
    Unload Me
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdOK_Click"
End Sub

Private Sub flgKit_Desc_Click()
    txtKit_Desc.Text = txtKit_Desc.Text & " " & flgKit_Desc.TextMatrix(flgKit_Desc.Row, 1)
    txtKit_Desc.SelStart = Len(txtKit_Desc.Text)
End Sub

Private Sub flgKit_Desc_DblClick()
On Error GoTo Handle
    strKit_Desc = txtKit_Desc.Text
    Unload Me
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " flgKit_Desc_DblClick"
End Sub

Private Sub Form_Activate()
    On Error GoTo Handle
        strKit_Desc = ""
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    Set rsCookingMessage = Open_Table(cnData, "CookingInstruction")
    Call setflgKit_Desc
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub
Private Sub setflgKit_Desc()
On Error GoTo errHdl
    Dim intCount    As Integer
    With flgKit_Desc
        .Font = ".vnArialH"
        .ColWidth(0) = 800
        .ColWidth(1) = 11500
        .TextMatrix(0, 0) = "STT"
        .TextMatrix(0, 1) = "         Chÿ d…n bar, b’p"
    End With
    
    If rsCookingMessage Is Nothing Then Exit Sub
    If rsCookingMessage.State = 0 Then Exit Sub
    
    If rsCookingMessage.EOF And rsCookingMessage.BOF Then
        With flgKit_Desc
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
        End With
        Exit Sub
    End If
   flgKit_Desc.Rows = rsCookingMessage.RecordCount + 1
    intCount = 0
    Do While Not rsCookingMessage.EOF
        intCount = intCount + 1
        flgKit_Desc.TextMatrix(intCount, 0) = rsCookingMessage!NO
        flgKit_Desc.TextMatrix(intCount, 1) = rsCookingMessage!Data
        rsCookingMessage.MoveNext
    Loop
'    SetColorFlexGrid flgKit_Desc, 1, 1, flgKit_Desc.Cols

'    Call flgKit_Desc_Click
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - setflgKit_Desc "
End Sub


Public Property Get Get_Kit_Desc() As Variant
    Get_Kit_Desc = strKit_Desc
End Property


Private Sub txtKit_Desc_DblClick()
On Error GoTo Handle:
        With frmKeyboard
            .FormCallkeyboard = "Other"
            .txtInput.PasswordChar = ""
            .Show vbModal
            txtKit_Desc.Text = .Let_Text_Input
        End With
    Call cmdOK_Click
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtKit_Desc_DblClick "

End Sub

Private Sub txtKit_Desc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOK_Click
    End If
End Sub

Public Property Get Let_Kit_Des() As Variant
    Let_Kit_Des = strKit_Desc
End Property



Public Property Get Let_OK() As Variant
    Let_OK = isOK
End Property


