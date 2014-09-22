VERSION 5.00
Begin VB.Form frmUser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Thªm míi nh©n viªn thu ng©n"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
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
   ScaleHeight     =   4320
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtID 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      MaxLength       =   2
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   1215
      Left            =   6360
      TabIndex        =   5
      Tag             =   "L7"
      Top             =   2280
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2143
      BTYPE           =   5
      TX              =   "&Tho¸t"
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
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUser.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdDone 
      Height          =   1215
      Left            =   6360
      TabIndex        =   4
      Tag             =   "L6"
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2143
      BTYPE           =   5
      TX              =   "&§ång ý"
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
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUser.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.TextBox txtRePassword 
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
      Left            =   2760
      TabIndex        =   1
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtPassword 
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
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.ComboBox cboLevel 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2880
      Width           =   3015
   End
   Begin VB.TextBox txtUserName 
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
      Left            =   2760
      TabIndex        =   2
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label lblCheck 
      Caption         =   "KiÓm tra"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5040
      TabIndex        =   15
      Tag             =   "L8"
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label lblNoteMatch 
      Alignment       =   2  'Center
      Caption         =   "chó ý"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   720
      TabIndex        =   14
      Top             =   1800
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.Label lblNoteID 
      Alignment       =   2  'Center
      Caption         =   "chó ý"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   600
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label Label1 
      Caption         =   "L­u ý: CÊp 1 lµ cÊp cao nhÊt ®­îc sö dông mäi tÝnh n¨ng cña phÇn mÒm, kh«ng ph©n quyÒn."
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Tag             =   "L5"
      Top             =   3720
      Width           =   8175
   End
   Begin VB.Label lblID 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "M· sè ®¨ng nhËp:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   5640
      TabIndex        =   11
      Tag             =   "L1"
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblRePassword 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NhËp l¹i mËt khÈu:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Tag             =   "L2"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label lblPassword 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MËt khÈu:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Tag             =   "L1"
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label lblLevel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ph©n cÊp qu¶n lý:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Tag             =   "L4"
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label lblUserName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tªn ®¨ng nhËp:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Tag             =   "L3"
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "frmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim DescArr() As String
Dim blnDbClick As Boolean
Dim isClose As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDone_Click()
Dim IDUSER As String
    IDUSER = TrimSpecialChar(txtPassword.Text)
    If IDUSER <> TrimSpecialChar(txtRePassword.Text) Then
        MsgBox "X¸c nhËn mËt khÈu kh«ng ®óng !"
        txtPassword.Text = ""
        txtRePassword.Text = ""
        txtPassword.SetFocus
        Exit Sub
    Else
        With frmListUser.rs
            If Not .BOF And .RecordCount > 0 Then .MoveFirst
            If Len(IDUSER) <= 2 Then
                MsgBox "M· sè ®¨ng nhËp Ýt nhÊt 3 ký tù"
            Else
                .Find "ID='" & Left(IDUSER, 2) & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    MsgBox "M· sè ®· tån t¹i"
                    Exit Sub
                    txtPassword.SetFocus
                    txtPassword.SelStart = 0
                    txtPassword.SelLength = 9999
                Else
                    .addNew
                    '!ID = txtID.Text
                    !ID = Left(IDUSER, 2)
                    !userName = txtUserName.Text
                    !UserLevel = cboLevel.Text
                    !Password = Right(IDUSER, Len(IDUSER) - 2)
                    .Update
                    WritePasswordData frmListUser.rs
                    Unload Me
                End If
            End If
        End With
    End If
End Sub

Private Sub Form_Activate()
    On Error GoTo handle
    Dim ctrl As Control
    If cmdCancel.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    DescArr = LoadLanguage(LngFile, "#02:015:")
        Me.Caption = DescArr(9)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
        Next ctrl
    Exit Sub
handle:
    MsgBox Err.Number & Err.Description & Me.Name & " Form_Activate "
End Sub

Private Sub Form_Load()
On Error GoTo handle
Dim i As Integer
    isClose = True
    
    For i = 1 To 8 Step 1
        cboLevel.AddItem i, i - 1
    Next
    cboLevel.ListIndex = 0
    txtUserName.Text = ""
    txtPassword.Text = ""
    txtPassword.PasswordChar = "*"
    txtRePassword.Text = ""
    txtRePassword.PasswordChar = "*"
    cmdDone.Enabled = False
Exit Sub
handle:
    MsgBox Err.Number & Err.Description & Me.Name & " Form_Load"
End Sub

Private Sub lblCheck_Click()
On Error GoTo handle
    With frmListUser.rs
        If Not .BOF And .RecordCount > 0 Then .MoveFirst
            If txtPassword.Text <> "" Then
                .Find "ID='" & Left(Trim(txtPassword.Text), 2) & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        lblNoteID.Visible = True
                        lblNoteID.Caption = "M· sè nµy ®· tån t¹i"
                        lblNoteID.ForeColor = vbRed
                        txtPassword.SetFocus
                        txtPassword.SelStart = 0
                        txtPassword.SelLength = 9999
                    
                    Else
                        If Not IsNumeric(Left(txtPassword.Text, 2)) Then
                            lblNoteID.ForeColor = vbRed
                            lblNoteID.Visible = True
                            lblNoteID.Caption = "2 ký tù ®Çu tiªn b¾t buéc ph¶i lµ ký sè !" & vbCrLf & " tõ 00 ®Õn 99"
                        Else
                            lblNoteID.Visible = True
                            lblNoteID.ForeColor = vbGreen
                            lblNoteID.Caption = "B¹n cã thÓ sö dông m· sè nµy"
                        End If
                    End If
            Else
                MsgBox "M· sè ®¨ng nhËp kh«ng ®­îc rçng, Ýt nhÊt ph¶i cã 3 ký tù", vbInformation
            End If
    End With
Exit Sub
handle:
    MsgBox Err.Number & Err.Description & Me.Name & " lblCheck_Click"
End Sub

Private Sub txtID_DblClick()
On Error GoTo handle
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtID.Text = .Let_Text_Input
           
        End With
    Exit Sub
handle:
    MsgBox Err.Number & Err.Description & Me.Name

End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
On Error GoTo handle
    If KeyAscii < 32 Then Exit Sub
    Select Case KeyAscii
        Case 48 To 57, 46
        Case 13
            txtUserName.SetFocus
        Case Else:   KeyAscii = 0
    End Select
Exit Sub
handle:
MsgBox Err.Number & Err.Description & Me.Name & " txtQty_KeyPress"
    
End Sub

Private Sub txtPassword_Change()
    If txtUserName.Text <> "" And txtPassword.Text <> "" And txtRePassword <> "" Then
        cmdDone.Enabled = True
    Else
        cmdDone.Enabled = False
    End If
    If txtPassword.Text <> "" Then
        txtRePassword.Enabled = True
    Else
        txtRePassword.Enabled = False
    End If
End Sub

Private Sub txtPassword_DblClick()
On Error GoTo handle
    blnDbClick = True
        With frmKeyboard
            .txtInput.PasswordChar = "*"
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtPassword.Text = .Let_Text_Input
        End With
    Exit Sub
handle:
    MsgBox Err.Number & Err.Description & Me.Name

End Sub

Private Sub txtPassword_LostFocus()
    On Error GoTo handle
    If blnDbClick Or isClose Then Exit Sub
        If Len(txtPassword.Text) <= 2 Then
            lblNoteID.Visible = True
            lblNoteID.Caption = "M· sè ®¨ng nhËp Ýt nhÊt 3 ký tù"
            If Not IsNumeric(Left(txtPassword.Text, 2)) Then
                MsgBox "2 ký tù ®Çu tiªn b¾t buéc ph¶i dïng ký sè !", vbInformation
            End If
            txtPassword.SetFocus
        Else
            lblNoteID.Caption = "*****"
            lblNoteID.ForeColor = vbGreen
        End If
    Exit Sub
handle:
MsgBox Err.Number & Err.Description & Me.Name & " txtPassword_LostFocus"
End Sub

Private Sub txtRePassword_Change()
    On Error GoTo handle
        If Trim(txtRePassword.Text) <> Trim(txtPassword.Text) Then
            lblNoteMatch.Visible = True
            lblNoteMatch.Caption = " NhËp l¹i mËt m· ®¨ng nhËp kh«ng trïng khíp!"
            lblNoteMatch.ForeColor = vbRed
        Else
            lblNoteMatch.Caption = " *****"
            lblNoteMatch.ForeColor = vbGreen
        End If
    Exit Sub
handle:
MsgBox Err.Number & Err.Description & Me.Name & " txtRePassword_Change"
End Sub

Private Sub txtRePassword_DblClick()
On Error GoTo handle
    blnDbClick = True
        With frmKeyboard
            .txtInput.PasswordChar = "*"
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtRePassword.Text = .Let_Text_Input
        End With
    Exit Sub
handle:
    MsgBox Err.Number & Err.Description & Me.Name


End Sub

Private Sub txtRePassword_LostFocus()
On Error GoTo handle
If blnDbClick Or isClose Then Exit Sub
        If Trim(txtRePassword.Text) <> Trim(txtPassword.Text) Then
            lblNoteMatch.Visible = True
            lblNoteMatch.Caption = " NhËp l¹i mËt m· ®¨ng nhËp kh«ng trïng khíp!"
            lblNoteMatch.ForeColor = vbRed
            txtRePassword.SetFocus
        Else
            lblNoteMatch.Caption = " *****"
            lblNoteMatch.ForeColor = vbGreen
        End If
    Exit Sub
handle:
MsgBox Err.Number & Err.Description & Me.Name & " txtRePassword_LostFocus"
End Sub

Private Sub txtUserName_Change()
    If txtUserName.Text <> "" And txtPassword.Text <> "" And txtRePassword <> "" Then
        cmdDone.Enabled = True
    Else
        cmdDone.Enabled = False
    End If
End Sub

Private Sub txtUserName_DblClick()
On Error GoTo handle
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtUserName.Text = .Let_Text_Input
        End With
    Exit Sub
handle:
    MsgBox Err.Number & Err.Description & Me.Name

End Sub
