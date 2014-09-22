VERSION 5.00
Begin VB.Form frmEnterClerkCode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Clerk Code"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2775
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
   ScaleHeight     =   1515
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton cmdOK 
      Height          =   645
      Left            =   90
      TabIndex        =   2
      Top             =   840
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1138
      BTYPE           =   14
      TX              =   "§ång ý"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
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
      MICON           =   "frmEnterClerkCode.frx":0000
      PICN            =   "frmEnterClerkCode.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   270
      TabIndex        =   1
      Top             =   375
      Width           =   2040
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   645
      Left            =   1290
      TabIndex        =   3
      Top             =   840
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1138
      BTYPE           =   14
      TX              =   "Tho¸t"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
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
      MICON           =   "frmEnterClerkCode.frx":0656
      PICN            =   "frmEnterClerkCode.frx":0672
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblCode 
      Alignment       =   2  'Center
      Caption         =   "New Clerk Code:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   0
      Tag             =   "L2"
      Top             =   75
      Width           =   1965
   End
End
Attribute VB_Name = "frmEnterClerkCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Public FormCall As Object
    Dim iCode As Integer
    Dim i As Integer
'            --------- FORM --------
Private Sub Form_Activate()
On Error GoTo errHdl

    Dim DescArr() As String
    Dim Ctrl As Control
        
    DescArr = LoadLanguage(LngFile, "#03:027:")
    If cmdOK.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    For Each Ctrl In Me
    DoEvents
        If Left(Ctrl.Tag, 1) = "L" Then Ctrl.Caption = DescArr(Mid(Ctrl.Tag, 2))
    Next Ctrl
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl

    With Me
        .Height = 1830
        .Width = 2430
        .WindowState = 0
    End With
    Dim res As New ADODB.Recordset
    Set res = Open_Table(cnData, "Clerk")
    If res.State = 0 Then Exit Sub
    For i = 0 To res.Fields.Count - 1
    DoEvents
        If res.Fields(i).Name = "ClerkCode" Then
            txtCode.MaxLength = res.Fields(i).DefinedSize
            Exit For
        End If
    Next i
    CloseRecordset res
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Form_Load"
End Sub
'            -------- COMMANDBUTTON --------
Private Sub cmdCancel_Click()
On Error GoTo errHdl

    FormCall.Get_Code = 0
    Unload Me
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - cmdCancel_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo errHdl

    Dim i As Integer
    Dim fFind As Boolean
    
    If txtCode.Text = "" Then
        txtCode.Text = "0"
        GoTo 1
    End If
    fFind = False
    With FormCall.flexClerk
        For i = 1 To .Rows - 1
        DoEvents
            If .TextMatrix(i, 2) <> "" Then
                If CInt(.TextMatrix(i, 2)) = CInt(txtCode.Text) Then
                    fFind = True
                    Exit For
                End If
            End If
        Next i
    End With
    If fFind Then
        txtCode.SetFocus
        txtCode.SelStart = 0
        txtCode.SelLength = 9999
    Else
1:
        FormCall.Get_Code = CInt(txtCode.Text)
        Unload Me
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - cmdOK_Click"
End Sub
'            --------- TEXTBOX --------
Private Sub txtCode_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    Select Case KeyAscii
        Case 48 To 57
        Case Is < 32
        Case Else: KeyAscii = 0
    End Select
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - txtCode_KeyPress"
End Sub
