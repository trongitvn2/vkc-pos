VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "§æi mËt khÈu"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   855
      Left            =   5280
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "&Tho¸t"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmChangePassword.frx":0000
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
      Height          =   855
      Left            =   5280
      TabIndex        =   4
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "&OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmChangePassword.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.TextBox txtConfirmPassword 
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
      Left            =   2640
      TabIndex        =   3
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtNewPassword 
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
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtOldPassword 
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
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblmessage 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label1"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   2280
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.Label lblConfirmPassword 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X¸c nhËn mËt khÈu míi:"
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
      Left            =   -45
      TabIndex        =   7
      Top             =   1830
      Width           =   2655
   End
   Begin VB.Label lblNewPassword 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MËt khÈu míi:"
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
      Left            =   -45
      TabIndex        =   6
      Top             =   1110
      Width           =   2655
   End
   Begin VB.Label lblOldPassword 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MËt khÈu cò:"
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
      Left            =   -45
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsChange As ADODB.Recordset
Dim Pass_reset As String

Private Sub cmdCancel_Click()
On Error GoTo errHdl

    Unload Me
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- cmdCancel_Click"
End Sub

Private Sub cmdDone_Click()
On Error GoTo errHdl
    With rsChange
    If .State = adStateOpen And .RecordCount > 0 Then .MoveFirst
    Do While Not .EOF
        If Left(txtNewPassword.Text, 2) = .Fields("ID") Then
            MsgBox "2 ký tù ®Çu tiªn cña mËt khÈu míi ®· ®­îc sö dông, vui lßng thay ®æi 2 ký tù ®Çu tiªn"
            Exit Sub
        End If
    .MoveNext
    Loop
        If Not .BOF And .RecordCount > 0 Then .MoveFirst
        
        .Find "ID='" & Left(Pass_reset, 2) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If Pass_reset = txtOldPassword.Text Then
                If txtNewPassword.Text = txtConfirmPassword.Text Then
                    !ID = Left(txtNewPassword, 2)
                    !Password = Mid(txtNewPassword.Text, 3, Len(txtNewPassword.Text) - 2)
                    .Update
                    
                    Call SavePasswordData(rsChange)
                    MsgBox "§æi mËt khÈu thµnh c«ng."
                    Unload Me
                Else
                    MsgBox "X¸c nhËn mËt khÈu kh«ng ®óng!"
                    txtOldPassword.Text = ""
                    txtNewPassword.Text = ""
                    txtConfirmPassword.Text = ""
                    txtOldPassword.SetFocus
                End If
            Else
                MsgBox "Sai mËt khÈu hiÖn hµnh!"
                txtOldPassword.Text = ""
                txtNewPassword.Text = ""
                txtConfirmPassword.Text = ""
                txtOldPassword.SetFocus
            End If
        Else
            MsgBox "Lçi ®¨ng nhËp tªn ng­êi sö dông."
            Unload Me
        End If
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- cmdDone_Click"
End Sub

Private Sub Form_Activate()
On Error GoTo errHdl

    Dim DescArr() As String
        
    DescArr = LoadLanguage(LngFile, "#02:014:")
    If cmdCancel.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    lblOldPassword.Caption = DescArr(2)
    lblNewPassword.Caption = DescArr(3)
    lblConfirmPassword.Caption = DescArr(4)
    cmdDone.Caption = DescArr(5)
    cmdCancel.Caption = DescArr(6)
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl
    
    Set rsChange = LoadPasswordData
    cmdDone.Enabled = False
    txtOldPassword.PasswordChar = "*"
    txtNewPassword.PasswordChar = "*"
    txtConfirmPassword.PasswordChar = "*"
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsChange = Nothing
End Sub

Private Sub txtConfirmPassword_Change()
On Error GoTo errHdl

    If txtOldPassword.Text <> "" And txtNewPassword.Text <> "" And txtConfirmPassword.Text <> "" Then
        If txtConfirmPassword.Text = txtNewPassword.Text Then
            cmdDone.Enabled = True
            lblmessage.Visible = False
        Else
            lblmessage.Caption = "X¸c nhËn mËt khÈu kh«ng ®óng !"
            lblmessage.Visible = True
            cmdDone.Enabled = False
        End If
    Else
        cmdDone.Enabled = False
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- txtConfirmPassword_Change"
End Sub

Private Sub txtConfirmPassword_DblClick()
On Error GoTo Handle
        With frmKeyboard
            .txtInput.PasswordChar = "*"
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtConfirmPassword.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name

End Sub

Private Sub txtNewPassword_Change()
On Error GoTo errHdl

    If txtOldPassword.Text <> "" And txtNewPassword.Text <> "" And txtConfirmPassword.Text <> "" Then
        cmdDone.Enabled = True
    Else
        cmdDone.Enabled = False
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- txtNewPassword_Change"
End Sub

Private Sub txtNewPassword_DblClick()
On Error GoTo Handle
        With frmKeyboard
            .txtInput.PasswordChar = "*"
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtNewPassword.Text = .Let_Text_Input
           
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name

End Sub

Private Sub txtOldPassword_Change()
On Error GoTo errHdl

    If txtOldPassword.Text <> "" And txtNewPassword.Text <> "" And txtConfirmPassword.Text <> "" Then
        cmdDone.Enabled = True
    Else
        cmdDone.Enabled = False
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- txtOldPassword_Change"
End Sub

Private Sub txtOldPassword_DblClick()
On Error GoTo Handle
        With frmKeyboard
            .txtInput.PasswordChar = "*"
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtOldPassword.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name

End Sub
Public Property Let Let_Pass_Call(ByVal vNewValue As Variant)
    Pass_reset = vNewValue
End Property
