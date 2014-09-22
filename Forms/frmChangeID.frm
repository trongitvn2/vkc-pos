VERSION 5.00
Begin VB.Form frmChangeID 
   Caption         =   "§æi m· sè vµo ca"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4920
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
   ScaleHeight     =   2565
   ScaleWidth      =   4920
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNewID 
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
      Left            =   2205
      TabIndex        =   3
      Top             =   240
      Width           =   2415
   End
   Begin VB.TextBox txtConfirmID 
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
      Left            =   2205
      TabIndex        =   2
      Top             =   840
      Width           =   2415
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   945
      Left            =   2730
      TabIndex        =   0
      Top             =   1530
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1667
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
      MICON           =   "frmChangeID.frx":0000
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
      Height          =   945
      Left            =   1050
      TabIndex        =   1
      Top             =   1530
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1667
      BTYPE           =   14
      TX              =   "&§ång ý"
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
      MICON           =   "frmChangeID.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblID 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "NhËp m· sè míi:"
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
      Left            =   360
      TabIndex        =   5
      Top             =   390
      Width           =   1695
   End
   Begin VB.Label lblConfirmPassword 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "X¸c nhËn m· sè"
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
      Left            =   -120
      TabIndex        =   4
      Top             =   990
      Width           =   2175
   End
End
Attribute VB_Name = "frmChangeID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsChange As ADODB.Recordset

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
Dim rsInvoice As New ADODB.Recordset
    Dim SQL As String
    
        SQL = "Select * from Invoice_Totals where Cashier_ID='" & UserID & "' and left(DateTime,8)='" & DateDefault & "'"
    
    Set rsInvoice = OpenCriticalTable(SQL, cnData)
    If rsInvoice.State = 1 Then
        If rsInvoice.RecordCount > 0 Then
            rsInvoice.MoveFirst
        Else
            Exit Sub
        End If
    End If
    With rsInvoice
        Do While Not .EOF
            .Fields("Cashier_ID") = txtNewID.Text
            .Update
'            .Requery
            
        .MoveNext
        Loop
    End With
    With rsChange
        If Not .BOF And .RecordCount > 0 Then .MoveFirst
        .Find "ID='" & UserID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If txtNewID.Text = txtConfirmID.Text Then
                !ID = txtNewID.Text
                .Update
               Call SavePasswordData(rsChange)
                MsgBox "§æi m· s« thµnh c«ng."
                Unload Me
            Else
                MsgBox "X¸c nhËn ID kh«ng ®óng!"
                txtNewID.Text = ""
                txtConfirmID.Text = ""
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
    If cmdDone.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Dim ctrl As Control
    DescArr = LoadLanguage(LngFile, "#03:009:")
    Me.Caption = DescArr(1)
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl

    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl
    
    Set rsChange = LoadPasswordData
    cmdDone.Enabled = False
    txtNewID.PasswordChar = "*"
    txtConfirmID.PasswordChar = "*"
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsChange = Nothing
End Sub

Private Sub txtConfirmID_Change()
On Error GoTo errHdl

    If txtNewID.Text <> "" And txtConfirmID.Text <> "" Then
        cmdDone.Enabled = True
    Else
        cmdDone.Enabled = False
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- txtConfirmID_Change"
End Sub

Private Sub txtConfirmID_DblClick()
On Error GoTo Handle
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtConfirmID.Text = .Let_Text_Input
           
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub txtConfirmID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call cmdDone_Click
End If
End Sub

Private Sub txtNewID_Change()
On Error GoTo errHdl

    If txtNewID.Text <> "" And txtConfirmID.Text <> "" Then
        cmdDone.Enabled = True
    Else
        cmdDone.Enabled = False
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- txtNewPassword_Change"
End Sub


Private Sub txtNewID_DblClick()
    On Error GoTo Handle
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtNewID.Text = .Let_Text_Input
           
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub txtNewID_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtConfirmID.SetFocus
End If
End Sub
