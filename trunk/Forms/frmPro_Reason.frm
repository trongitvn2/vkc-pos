VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPro_Reason 
   Caption         =   "Lý do gi¶m %"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12930
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
   ScaleHeight     =   9120
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton cmdKeyboard 
      Height          =   1335
      Left            =   10200
      TabIndex        =   3
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2355
      BTYPE           =   3
      TX              =   "Bµn phÝm"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPro_Reason.frx":0000
      PICN            =   "frmPro_Reason.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   1335
      Left            =   10200
      TabIndex        =   2
      Top             =   3600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2355
      BTYPE           =   3
      TX              =   "Hñy bá"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPro_Reason.frx":046E
      PICN            =   "frmPro_Reason.frx":048A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   9015
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9975
      Begin MSFlexGridLib.MSFlexGrid flgReason 
         Height          =   8655
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   15266
         _Version        =   393216
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin prjTouchScreen.MyButton cmdOK 
      Height          =   1335
      Left            =   10200
      TabIndex        =   0
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2355
      BTYPE           =   3
      TX              =   "§ång ý"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPro_Reason.frx":0AC4
      PICN            =   "frmPro_Reason.frx":0AE0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdAddNew 
      Cancel          =   -1  'True
      Height          =   1335
      Left            =   10200
      TabIndex        =   5
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2355
      BTYPE           =   3
      TX              =   "T¹o míi"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12648384
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPro_Reason.frx":111A
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
Attribute VB_Name = "frmPro_Reason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim reason As String
Dim OK_Cancel As Boolean
Dim rsReason As New ADODB.Recordset

Private Sub cmdAddNew_Click()
    frmDiscount_reason_list.Show vbModal
    Call Init_List
End Sub

Private Sub cmdCancel_Click()
    OK_Cancel = False
    Unload Me
End Sub

Private Sub cmdKeyboard_Click()
    With frmKeyboard
        .txtInput.PasswordChar = ""
        .txtInput.SelStart = 0
        .txtInput.SelLength = 9999
        .FormCallkeyboard = "Other"
        .Show vbModal
        reason = .Let_Text_Input
        cmdOK_Click
    End With
End Sub

Private Sub cmdOK_Click()
If Trim(reason) <> "" Then
    OK_Cancel = True
Else
    OKCancel = False
End If
    Unload Me
End Sub

Public Property Get Let_Reason() As Variant
    Let_Reason = reason
End Property

Public Property Get Let_OK_Cancel() As Variant
    Let_OK_Cancel = OK_Cancel
End Property

Private Sub flgReason_Click()
    On Error GoTo Handle
        reason = flgReason.TextMatrix(flgReason.Row, 1)
        flgReason.SelectionMode = flexSelectionByRow
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " flgReason_Click"
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    
    Call Init_List
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Public Sub Init_List()
Dim i As Integer
If Check_Table_exist("Promotion_Reason") = False Then Call Create_Promotion_Reason
    Set rsReason = Open_Table(cnData, "Promotion_Reason")
i = 1
    With flgReason
        .Cols = 2
        .Rows = 1
        .ColWidth(0) = 0
        .ColWidth(1) = 9500
        .TextMatrix(0, 1) = "Lý do"
        .ColAlignment(1) = 2
        .TextMatrix(0, 1) = ""
    End With
        With rsReason
            flgReason.Rows = rsReason.RecordCount + 1
            Do While Not .EOF
                flgReason.TextMatrix(i, 1) = .Fields("Pro_Desc")
            .MoveNext
            i = i + 1
            Loop
        End With
End Sub
