VERSION 5.00
Begin VB.Form frmReason 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enter Delete Reason  for this Item"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   FillColor       =   &H00FFC0C0&
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReason.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "L2"
   Begin VB.ListBox SelectList 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      ItemData        =   "frmReason.frx":000C
      Left            =   120
      List            =   "frmReason.frx":001F
      TabIndex        =   3
      Top             =   840
      Width           =   7695
   End
   Begin prjTouchScreen.MyButton cmdKeyBoard 
      Height          =   615
      Left            =   5760
      TabIndex        =   2
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   ""
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReason.frx":0055
      PICN            =   "frmReason.frx":0071
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdOK 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   6960
      TabIndex        =   1
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "&OK"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReason.frx":04C3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.TextBox txtReason 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      HideSelection   =   0   'False
      Left            =   120
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim str As String

Private Sub cmdKeyboard_Click()
    With frmKeyboard
        .txtInput.Text = txtReason.Text
        .txtInput.SelStart = 0
        .txtInput.SelLength = Len(.txtInput.Text)
        .txtInput.PasswordChar = ""
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtReason.Text = .Let_Text_Input
        cmdOK_Click
    End With

End Sub

Private Sub cmdOK_Click()
    If Trim(txtReason.Text) = "" Then
        MsgBox "B¹n ph¶i nhËp lý do xãa mãn ®· order !!!", vbOKOnly
    Else
        str = txtReason.Text
        Unload Me
    End If
End Sub

Public Property Get GetReason() As Variant
    GetReason = str
End Property

Private Sub Form_Activate()
On Error GoTo Handle
    Dim ctrl As Control
    Dim DescArr() As String
    If cmdOK.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
        DescArr = LoadLanguage(LngFile, "#03:026:")
        Me.Caption = DescArr(1)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
        Next ctrl
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Form_Load"

End Sub

Private Sub Form_Load()
    str = ""
End Sub

Private Sub SelectList_Click()
    txtReason.Text = SelectList.Text
    
End Sub

Private Sub SelectList_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        Call cmdOK_Click
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " - SelectList_KeyPress"
End Sub

Private Sub txtReason_DblClick()
    With frmKeyboard
        .txtInput.Text = txtReason.Text
        .txtInput.SelStart = 0
        .txtInput.SelLength = Len(.txtInput.Text)
        .txtInput.PasswordChar = ""
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtReason.Text = .Let_Text_Input
        cmdOK_Click
    End With
    
   
End Sub

Private Sub txtReason_KeyDown(KeyCode As Integer, Shift As Integer)
    SelectList.SetFocus
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmdOK_Click
    Else
        txtReason.Text = txtReason.Text & Chr(KeyAscii)
    End If
    
End Sub
