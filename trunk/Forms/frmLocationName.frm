VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmLocationName 
   Caption         =   "LocationName"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   ClipControls    =   0   'False
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
   ScaleHeight     =   5505
   ScaleWidth      =   10215
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   5760
      ScaleHeight     =   855
      ScaleWidth      =   4395
      TabIndex        =   4
      Top             =   0
      Width           =   4455
      Begin VB.Label lblNo 
         BackColor       =   &H80000008&
         Caption         =   "Group No"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   45
         Width           =   3135
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000008&
         Caption         =   "Group Name"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   600
         TabIndex        =   5
         Top             =   450
         Width           =   3135
      End
   End
   Begin VB.Frame frmCmd 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   5880
      TabIndex        =   0
      Top             =   4320
      Width           =   4140
      Begin prjTouchScreen.MyButton cmdClose 
         Cancel          =   -1  'True
         Height          =   735
         Left            =   2760
         TabIndex        =   1
         Tag             =   "L5"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "&Tho¸t"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmLocationName.frx":0000
         PICN            =   "frmLocationName.frx":001C
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
         Height          =   735
         Left            =   1440
         TabIndex        =   2
         Tag             =   "L4"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "&Gióp ®ì"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmLocationName.frx":62B6
         PICN            =   "frmLocationName.frx":62D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdSave 
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Tag             =   "L3"
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "&L­u"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmLocationName.frx":690C
         PICN            =   "frmLocationName.frx":6928
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
   Begin TabDlg.SSTab tab1 
      Height          =   3015
      Left            =   5880
      TabIndex        =   7
      Top             =   1080
      Width           =   4245
      _ExtentX        =   7488
      _ExtentY        =   5318
      _Version        =   393216
      TabHeight       =   882
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tªn khu"
      TabPicture(0)   =   "frmLocationName.frx":6E6C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblGroupName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblService"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtData"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtVAT"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtService"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Gi¸ mãn"
      TabPicture(1)   =   "frmLocationName.frx":6E88
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Gi¸ giê"
      TabPicture(2)   =   "frmLocationName.frx":6EA4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraTimeLevel"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "chkused"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.CheckBox chkused 
         Caption         =   "TÝnh tiÒn giê"
         Height          =   495
         Left            =   -74760
         TabIndex        =   21
         Top             =   960
         Width           =   2775
      End
      Begin VB.Frame fraTimeLevel 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   -74880
         TabIndex        =   20
         Top             =   1080
         Width           =   3975
         Begin VB.OptionButton optTime 
            Caption         =   "Møc gi¸ 4"
            Height          =   375
            Index           =   3
            Left            =   2400
            TabIndex        =   25
            Top             =   1080
            Width           =   1455
         End
         Begin VB.OptionButton optTime 
            Caption         =   "Møc gi¸ 3"
            Height          =   375
            Index           =   2
            Left            =   2400
            TabIndex        =   24
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton optTime 
            Caption         =   "Møc gi¸ 2"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   23
            Top             =   1080
            Width           =   2295
         End
         Begin VB.OptionButton optTime 
            Caption         =   "Møc gi¸ 1"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   22
            Top             =   480
            Value           =   -1  'True
            Width           =   2295
         End
      End
      Begin VB.TextBox txtService 
         Alignment       =   1  'Right Justify
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
         Left            =   2040
         TabIndex        =   18
         Text            =   "0"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox txtVAT 
         Alignment       =   1  'Right Justify
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
         Left            =   360
         TabIndex        =   17
         Text            =   "0"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   11
         Top             =   600
         Width           =   4095
         Begin VB.Frame Frame2 
            Height          =   2055
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   3855
            Begin VB.OptionButton OptPrice 
               Caption         =   "Møc gi¸  3"
               Height          =   495
               Index           =   2
               Left            =   480
               TabIndex        =   15
               Top             =   1440
               Width           =   2415
            End
            Begin VB.OptionButton OptPrice 
               Caption         =   "Møc gi¸  2"
               Height          =   375
               Index           =   1
               Left            =   480
               TabIndex        =   14
               Top             =   840
               Width           =   2415
            End
            Begin VB.OptionButton OptPrice 
               Caption         =   "Møc gi¸  1"
               Height          =   495
               Index           =   0
               Left            =   480
               TabIndex        =   13
               Top             =   120
               Width           =   2415
            End
         End
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         MaxLength       =   30
         TabIndex        =   8
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label lblService 
         Caption         =   "PhÝ phôc vô:"
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "VAT:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblGroupName 
         Caption         =   "&Data:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Tag             =   "L2"
         Top             =   720
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   5535
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   9763
      _Version        =   393216
      BackColorBkg    =   16777215
      TextStyleFixed  =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArialH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmLocationName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim res As New ADODB.Recordset
    Dim fLoad As Boolean, fUpdate As Boolean
    Dim fClick As Boolean
    Dim fActivate As Boolean
    Dim arrUpdate() As Variant

Private Sub chkKaraoke_Click()
    Call UpdateData
End Sub

Private Sub chkused_Click()
    On Error GoTo Handle
        If chkused.Value = 1 Then
            fraTimeLevel.Enabled = True
            UpdateData
        Else
            fraTimeLevel.Enabled = False
            UpdateData
        End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "chkused_Click"
End Sub

Private Sub cmdSave_Click()
On Error GoTo errHdl
        fUpdate = False
        Add_DataUpdate_To_DB
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - cmdSave_Click"
End Sub
'           ------------ FORM ----------
Private Sub Form_Activate()
On Error GoTo errHdl

    Dim DescArr() As String
    Dim ctrl As Control

    DescArr = LoadLanguage(LngFile, "#03:031:")
    If cmdSave.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    If res.State = 0 Then
        cmdClose_Click
        Exit Sub
    End If
    If fActivate Then Exit Sub
    fActivate = True
    Me.Caption = DescArr(1)
    tab1.TabCaption(0) = DescArr(7)
    flex.TextMatrix(0, 0) = "Store"
    flex.TextMatrix(0, 1) = "Location"
    flex.TextMatrix(0, 2) = DescArr(7)
    flex.TextMatrix(0, 3) = "% VAT"
    flex.TextMatrix(0, 4) = "Møc gi¸"
    flex.TextMatrix(0, 5) = "PhÝ phôc vô"
    flex.TextMatrix(0, 6) = "TÝnh giê"
    flex.TextMatrix(0, 7) = "Møc giê"
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then
            ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
        End If
    Next ctrl
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    With flex
        If Shift = 2 Then   'dai dien cho cac fim ctrl,shift,alt
            If KeyCode = vbKeyDown Then 'to hop fim ctrl+keydown duoc click
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    If .Row > 19 Then .TopRow = .Row - 18
                End If
                KeyCode = 0
                flex_Click
            ElseIf KeyCode = vbKeyUp Then
                If .Row > 1 Then
                    .Row = .Row - 1
                    If .Row < .TopRow Then .TopRow = .Row
                End If
                KeyCode = 0
                flex_Click
            End If
        End If
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Form_KeyDown"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl

    Dim sTableName As String
'    With Me
'        .Height = 5355
'        .Width = 8805
'        .WindowState = 0
'    End With
    Set res = Open_Table(cnData, "Table_Diagram_Sections")
    If res.State = 0 Then Exit Sub
    Initalize
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Form_Load"
End Sub

Private Sub Initalize()
On Error GoTo errHdl

    fLoad = False: fUpdate = False: fActivate = False ': flagkeydown = False
    ReDim Preserve arrUpdate(0)
    SetDataInFlex
'    Call SetColorFlexGrid(flex, 1, 1, 2)
    With flex
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
        .ScrollTrack = True
    End With
    txtData.MaxLength = res.Fields("Section_ID").DefinedSize
    flex_Click
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Initalize"
End Sub
'           ------------ COMMANDBUTTON ----------
Private Sub cmdClose_Click()
On Error GoTo errHdl

    Dim response
    If Not fUpdate Then GoTo 1
    response = MsgBox("B¹n cã muèn l­u th«ng tin thay ®æi?", vbYesNoCancel)
    Select Case response
        Case vbNo:     GoTo 1
        Case vbCancel: Exit Sub
        Case vbYes
            Add_DataUpdate_To_DB
    End Select
1:
    CloseRecordset res
    Unload Me
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - cmdClose_Click"
End Sub
'           ----------- FLEXGRID ---------
Private Sub flex_Click()
On Error GoTo errHdl

    fLoad = False ' ko vao ham chkdouble_click
    With flex
        If res.RecordCount > 0 Then
            lblNo.Caption = .TextMatrix(.Row, 1)
            lblName.Caption = .TextMatrix(.Row, 2)
            txtData.Text = .TextMatrix(.Row, 2)
            txtVAT.Text = .TextMatrix(.Row, 3)
            OptPrice(CInt(.TextMatrix(.Row, 4)) - 1).Value = True
            txtService.Text = .TextMatrix(.Row, 5)
            If .TextMatrix(.Row, 6) = 1 Then
                chkused.Value = 1
                optTime(.TextMatrix(.Row, 7) - 1).Value = True
            Else
                 chkused.Value = 0
                'optTime(.TextMatrix(.Row, 7) - 1).Value = False
            End If
        Else
            SetTextNull
        End If
    End With
    fLoad = True
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - flex_Click"
End Sub

Private Sub flex_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 13 Then
        With txtData
            .SetFocus
            .SelStart = 0
            .SelLength = 9999
        End With
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - flex_KeyPress"
End Sub

Private Sub flex_EnterCell()
On Error GoTo errHdl

    If fLoad Then flex_Click
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - flex_EnterCell"
End Sub

Private Sub SetDataInFlex()
On Error GoTo errHdl

    Dim irow As Integer
    irow = 1
    SetHeaderFlexGrid
    With res
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                DoEvents
                flex.Rows = .RecordCount + 1
                flex.TextMatrix(irow, 0) = !Store_ID
                flex.TextMatrix(irow, 1) = !Location_ID
                flex.TextMatrix(irow, 2) = !Section_ID
                flex.TextMatrix(irow, 3) = !VAT
                flex.TextMatrix(irow, 4) = !Price_Level
                flex.TextMatrix(irow, 5) = !service_Charge
                If !isTimer = True Then
                    chkused.Value = 1
                    flex.TextMatrix(irow, 6) = 1
                Else
                    chkused.Value = 0
                    flex.TextMatrix(irow, 6) = 0
                End If
                flex.TextMatrix(irow, 7) = !TimeLevel
                irow = irow + 1
                .MoveNext
            Loop
        End If
    End With
    flex.ColSel = flex.Cols - 1
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - SetDataInFlex"
End Sub

Private Sub SetHeaderFlexGrid()
On Error GoTo errHdl

    With flex
        .Cols = res.Fields.count - 1
        .AllowUserResizing = flexResizeBoth
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .ColWidth(0) = 900: .ColAlignment(0) = 4
        .ColWidth(1) = 1200: .ColAlignment(1) = 1
        .ColWidth(2) = 900: .ColAlignment(1) = 4
        .ColWidth(3) = 1000: .ColAlignment(1) = 4
        .ColWidth(4) = 1000: .ColAlignment(1) = 4
        .ColWidth(5) = 1000: .ColAlignment(1) = 4
        .ColWidth(6) = 1000: .ColAlignment(1) = 4
        .ColWidth(7) = 1000: .ColAlignment(1) = 4
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - SetHeaderFlexGrid"
End Sub

Private Sub OptPrice_Click(Index As Integer)
    Call UpdateData
End Sub

Private Sub optTime_Click(Index As Integer)
    UpdateData
End Sub

Private Sub txtData_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        
        .PasswordChar = ""
        .Text = txtData.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtData.Text = .Let_Text_Input
        End With
        UpdateData
        cmdSave_Click
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtData_DblClick"
End Sub

'           ---------- TEXTBOX --------
Private Sub txtData_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    Dim iSelStart As Integer
    Dim Temp As String
    
    If KeyAscii = 33 Or KeyAscii = 44 Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case KeyAscii
        Case 8: Exit Sub 'key backspace
        Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight: Exit Sub
    End Select
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - txtData_KeyPress"
End Sub

Private Sub SetTextNull()
On Error GoTo errHdl

    txtData.Text = ""
    lblNo.Caption = "Location_ID"
    lblName.Caption = "LocationName 1"
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - SetTextNull"
End Sub

Private Sub txtData_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    Select Case KeyCode
        Case 8 'key backspace
        Case Is < 32, vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight: Exit Sub
    End Select
1:  UpdateData

    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - txtData_KeyUp"
End Sub
'           ---------- UPDATE DATA --------
Private Sub UpdateData()
On Error GoTo errHdl

    Dim strName As String
    Dim i, VATFee, Service As Integer
    
    If res.RecordCount = 0 Then Exit Sub
    If fClick Then Exit Sub
    fClick = True
    fUpdate = True
    strName = txtData.Text
    VATFee = txtVAT.Text
    Service = txtService.Text
        With flex
        .TextMatrix(.Row, 2) = strName
        .TextMatrix(.Row, 3) = VATFee
        .TextMatrix(.Row, 5) = Service
        If chkused.Value = 1 Then
            .TextMatrix(.Row, 6) = 1
        Else
            .TextMatrix(.Row, 6) = 0
        End If
        For i = 0 To 2
            If OptPrice(i).Value = True Then .TextMatrix(.Row, 4) = i + 1
        Next
        
        For i = 0 To 3
            If optTime(i).Value = True Then .TextMatrix(.Row, 7) = i + 1
        Next
        
        lblName.Caption = strName
    End With
    arrUpdate = Add_UpdatedData_To_Array(flex, arrUpdate)
    fClick = False
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - UpdateData"
End Sub

'           ----- ADD UPDATED DATA TO DATABASE ----
Private Sub Add_DataUpdate_To_DB()
On Error GoTo errHdl

    Dim i, j As Integer

    With res
        For i = 1 To UBound(arrUpdate)
            DoEvents
            .MoveFirst
            .Find "Location_ID=" & arrUpdate(i)(1)
            'For j = 2 To .Fields.Count - 1
            DoEvents
                .Fields(2) = arrUpdate(i)(2)
                '.Fields(3) = arrUpdate(i)(3)
                .Fields(4) = arrUpdate(i)(3)
                .Fields(5) = arrUpdate(i)(4)
                .Fields(6) = arrUpdate(i)(5)
                .Fields(7) = arrUpdate(i)(6)
                .Fields(8) = arrUpdate(i)(7)
            'Next j
            .Update
        Next i
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Add_DataUpdate_To_DB"
End Sub

Private Sub txtService_DblClick()
On Error GoTo Handle
        With frmPhimso
            .lblTitle.Caption = "NhËp % phÝ phôc vô"
            .FormCall = 3
            .Show vbModal
            txtService.Text = .Return_Value
            Call UpdateData
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtService_DblClick"
End Sub

Private Sub txtVAT_Change()
On Error GoTo Handle
        UpdateData
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtVAT_Change"
End Sub

Private Sub txtVAT_DblClick()
On Error GoTo Handle
        With frmPhimso
            .lblTitle.Caption = "NhËp møc thuÕ VAT"
            .FormCall = 3
             .Show vbModal
            txtVAT.Text = .Return_Value
            Call UpdateData
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtVAT_DblClick"
End Sub

Private Sub txtVAT_KeyPress(KeyAscii As Integer)
    On Error GoTo Handle
        UpdateData
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtVAT_KeyPress"
End Sub

Private Sub txtVAT_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error GoTo Handle
        UpdateData
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''
Private Sub txtService_Change()
On Error GoTo Handle
        UpdateData
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtService_Change"
End Sub

Private Sub txtService_KeyPress(KeyAscii As Integer)
    On Error GoTo Handle
        UpdateData
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtService_KeyPress"
End Sub

Private Sub txtService_KeyUp(KeyCode As Integer, Shift As Integer)
     On Error GoTo Handle
        UpdateData
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name
    
End Sub
