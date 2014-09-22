VERSION 5.00
Begin VB.Form frmSetPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "§Æt mÆt khÈu"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
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
   ScaleHeight     =   2310
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   735
      Left            =   5040
      TabIndex        =   5
      Tag             =   "L5"
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
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
      MCOL            =   16711680
      MPTR            =   1
      MICON           =   "frmSetPassword.frx":0000
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
      Height          =   735
      Left            =   5040
      TabIndex        =   4
      Tag             =   "L4"
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1296
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
      MCOL            =   16711680
      MPTR            =   1
      MICON           =   "frmSetPassword.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
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
      Left            =   2880
      TabIndex        =   1
      Top             =   240
      Width           =   2175
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
      Left            =   2880
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.Label lblNewPassword 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   480
      TabIndex        =   0
      Tag             =   "L2"
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblConfirmPassword 
      Alignment       =   1  'Right Justify
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
      Left            =   0
      TabIndex        =   2
      Tag             =   "L3"
      Top             =   1080
      Width           =   2895
   End
End
Attribute VB_Name = "frmSetPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()
On Error GoTo errHdl

    Unload Me

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdDone_Click()
On Error GoTo errHdl

    If txtNewPassword.Text <> txtConfirmPassword.Text Then
        MsgBox "X¸c nhËn mËt khÈu kh«ng ®óng."
        txtNewPassword.Text = ""
        txtConfirmPassword.Text = ""
        Exit Sub
    Else
        
            If Not .EOF And Not .BOF Then
                !Password = txtNewPassword.Text
                .Update
                WritePasswordData frmListUser.rs
            End If
        
        Unload Me
    End If

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo errHdl

    Dim ctrl As Control
    Dim DescArr() As String
    If cmdDone.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    DescArr = LoadLanguage(LngFile, "#03:027:")
    Me.Caption = DescArr(1)
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl

    cmdDone.Enabled = False
    txtNewPassword.PasswordChar = "*"
    txtConfirmPassword.PasswordChar = "*"

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub txtConfirmPassword_Change()
On Error GoTo errHdl

    If txtNewPassword.Text <> "" And txtConfirmPassword.Text <> "" Then
        cmdDone.Enabled = True
    Else
        cmdDone.Enabled = False
    End If
    
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
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

    If txtNewPassword.Text <> "" And txtConfirmPassword.Text <> "" Then
        cmdDone.Enabled = True
    Else
        cmdDone.Enabled = False
    End If

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
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
