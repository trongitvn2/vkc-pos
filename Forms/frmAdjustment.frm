VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdjustment 
   Caption         =   "Tû lÖ t¨ng gi¶m"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8835
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
   Icon            =   "frmAdjustment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4020
      ScaleHeight     =   915
      ScaleWidth      =   4665
      TabIndex        =   0
      Top             =   150
      Width           =   4725
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   45
         Width           =   4635
      End
      Begin VB.Label lblValue 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Rate"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   390
         Width           =   4635
      End
   End
   Begin TabDlg.SSTab tabGroup 
      Height          =   3705
      Left            =   4050
      TabIndex        =   3
      Top             =   1170
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   6535
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Th«ng tin t¨ng gi¶m"
      TabPicture(0)   =   "frmAdjustment.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblGroupName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblRate"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmCmd"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtRate"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox txtRate 
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
         Left            =   600
         TabIndex        =   12
         Tag             =   "2"
         Top             =   1920
         Width           =   3735
      End
      Begin VB.TextBox txtName 
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
         Left            =   600
         TabIndex        =   8
         Tag             =   "1"
         Top             =   1020
         Width           =   3735
      End
      Begin VB.Frame frmCmd 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   60
         TabIndex        =   4
         Top             =   2460
         Width           =   4530
         Begin prjTouchScreen.MyButton cmdClose 
            Height          =   765
            Left            =   3030
            TabIndex        =   5
            Top             =   180
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1349
            BTYPE           =   14
            TX              =   "§ãng"
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
            BCOL            =   16578804
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmAdjustment.frx":0028
            PICN            =   "frmAdjustment.frx":0044
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
            Height          =   765
            Left            =   1590
            TabIndex        =   6
            Top             =   180
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1349
            BTYPE           =   14
            TX              =   "Gióp ®ì"
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
            BCOL            =   16578804
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmAdjustment.frx":62DE
            PICN            =   "frmAdjustment.frx":62FA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin prjTouchScreen.MyButton cmdSend 
            Height          =   765
            Left            =   150
            TabIndex        =   7
            Top             =   180
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   1349
            BTYPE           =   14
            TX              =   "L­u"
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
            BCOL            =   16578804
            BCOLO           =   16777152
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmAdjustment.frx":6934
            PICN            =   "frmAdjustment.frx":6950
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
      Begin VB.Label lblRate 
         Caption         =   "Tû lÖ:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Tag             =   "L3"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label lblGroupName 
         Caption         =   "Tªn:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Tag             =   "L2"
         Top             =   600
         Width           =   1455
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flexAdj 
      Height          =   4905
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   8652
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      GridLinesFixed  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArialH"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim rsAdj As New ADODB.Recordset
    Dim fLoad As Boolean
    Dim fUpdate As Boolean
    Dim fActivate As Boolean
    Dim fFlexClick As Boolean
    Dim arrUpdate() As Variant
    Dim fClick As Boolean
    Dim i, j As Integer
    ''''

Private Sub cmdKeyboard_Click()
    frmKeyboard.Show vbModal
End Sub

Private Sub cmdSend_Click()
Dim res As Integer
    
    If fUpdate Then
    res = MsgBox("B¹n cã muèn l­u th«ng tin thay ®æi ?", vbYesNo)
    If res = vbYes Then
        fUpdate = False
        Add_DataUpdate_To_DB
    End If
    End If
End Sub


'           ----------- FORM -----------
Private Sub Form_Activate()
    Dim DescArr() As String
    Dim ctrl As Control
    
    If rsAdj.State = 0 Then
        cmdClose_Click
        Exit Sub
    End If
    If fActivate Then Exit Sub
    fActivate = True
    DescArr = LoadLanguage(LngFile, "#03:029:")
    If cmdSend.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(9)
    tabGroup.TabCaption(0) = DescArr(4)
    With flexAdj
        .TextMatrix(0, 0) = DescArr(1)
        .TextMatrix(0, 1) = DescArr(2)
        .ColAlignment(1) = 2
        .TextMatrix(0, 2) = DescArr(3)
        .ColAlignment(2) = 4
    End With
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    fClick = False
    With flexAdj
        If Shift = 2 Then   'dai dien cho cac fim ctrl,shift,alt
            If KeyCode = vbKeyDown Then 'to hop fim ctrl+keydown duoc click
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    If .Row > 16 Then .TopRow = .Row - 15
                End If
                KeyCode = 0
                flexAdj_Click
            ElseIf KeyCode = vbKeyUp Then
                If .Row > 1 Then
                    .Row = .Row - 1
                    If .Row < .TopRow Then .TopRow = .Row
                End If
                KeyCode = 0
                flexAdj_Click
            End If
        End If
    End With
End Sub

Private Sub Form_Load()
Dim i As Integer
    With Me
        .Height = 5600
        .Width = 9000
        .WindowState = 0
    End With
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsAdj = Open_Table(cnData, "Adjustment")
    If rsAdj.EOF Then
    For i = 1 To 5
        With rsAdj
            .addNew
            .Fields("AdjNo") = i
            .Fields("AdjName") = "Adjustment " & i
            .Fields("AdjRate") = 0
            .Update
        End With
    Next i
    End If
    If rsAdj.State = 0 Then Exit Sub
    Initalize
End Sub

Private Sub Initalize()
    fLoad = False: fUpdate = False: fActivate = False
    ReDim Preserve arrUpdate(0)
    SetDataInFlex
    With flexAdj
'        SetColorFlexGrid flexAdj, 1, 1, .Cols
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
        .ScrollTrack = True
    End With
    txtName.MaxLength = rsAdj.Fields("AdjName").DefinedSize
    flexAdj_Click
    fLoad = True
End Sub
'           ------------ COMMANDBUTTON ------------
Private Sub cmdClose_Click()
    Dim res
    
    If Not fUpdate Then GoTo 1
    res = MsgBox("B¹n cã muèn l­u th«ng tin thay ®æi?", vbYesNoCancel)
    Select Case res
        Case vbYes
            Add_DataUpdate_To_DB
        Case vbNo:  GoTo 1
        Case vbCancel: Exit Sub
    End Select
1:
    CloseRecordset rsAdj
    Unload Me
End Sub
'           ------------ FLEXGRID ----------
Private Sub flexAdj_Click()
    Dim sTemp() As String
    Dim ctrl As Control
    Dim i As Integer
    fLoad = False
    If rsAdj.RecordCount = 0 Then SetTextNull: Exit Sub
    If fFlexClick Then Exit Sub
    fFlexClick = True
    With flexAdj
        ReDim Preserve sTemp(.Cols - 1)
        For i = 1 To .Cols - 1
        DoEvents
            sTemp(i) = .TextMatrix(.Row, i)
        Next i
        For Each ctrl In Me
        DoEvents
            With ctrl
                If TypeOf ctrl Is TextBox And ctrl.Tag <= flexAdj.Cols - 1 Then
                    .Text = sTemp(.Tag)
                End If
            End With
        Next ctrl
        
        lblName.Caption = .TextMatrix(.Row, 1)
        lblValue.Caption = .TextMatrix(.Row, 2) & "%"
    End With
    fFlexClick = False
    fLoad = True
End Sub

Private Sub flexAdj_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With txtName
            .SetFocus
            .SelStart = 0
            .SelLength = 9999
        End With
    End If
End Sub

Private Sub flexAdj_EnterCell()
    If fLoad Then flexAdj_Click
End Sub

Private Sub SetDataInFlex()
    Dim irow As Integer
    Dim sTemp As String
    Dim i As Integer
    irow = 1
    SetHeaderFlexGrid
    With rsAdj
        If .RecordCount > 0 Then
            flexAdj.Rows = .RecordCount + 1
            Do While Not .EOF
            DoEvents
                For i = 0 To flexAdj.Cols - 1
                DoEvents
                    Select Case i
                        Case 0: sTemp = "AdjNo"
                        Case 1: sTemp = "AdjName"
                        Case 2: sTemp = "AdjRate"
                    End Select
                    flexAdj.TextMatrix(irow, i) = .Fields(sTemp)
                Next i
                irow = irow + 1
                .MoveNext
            Loop
        End If
    End With
    flexAdj.ColSel = flexAdj.Cols - 1
End Sub

Private Sub SetHeaderFlexGrid()
    With flexAdj
        .Cols = rsAdj.Fields.count
        .AllowUserResizing = flexResizeBoth
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .ColWidth(0) = 700
        .ColAlignment(0) = 4
        .ColWidth(1) = 2450
        .ColWidth(2) = 1100
        .ColAlignment(2) = 4
        
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    fClick = False
End Sub

Private Sub txtName_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        .PasswordChar = ""
        .Text = txtName.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            frmKeyboard.Show vbModal
            txtName.Text = .Let_Text_Input
       End With
        UpdateData
    Exit Sub
    
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

'           ------------- TEXTBOX -----------
Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 33 Or KeyAscii = 44 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 Then cmdSend.SetFocus
End Sub

Private Sub SetTextNull()
    txtName.Text = ""
    txtRate.Text = 0
    lblName.Caption = "Adjustment1"
    lblValue.Caption = "0"
End Sub
Private Sub txtName_KeyUp(KeyCode As Integer, Shift As Integer)
    If fClick Then Exit Sub
    fClick = True
    Select Case KeyCode
        Case 27: fClick = False
        Case 8 'key backspace
        Case Is < 32, vbKeyDown, vbKeyUp: Exit Sub
    End Select
    UpdateData
    Delay (100)
    fClick = False
End Sub
'           ----------- UPDATE DATA ----------
Private Sub UpdateData()
    Dim strName, strRate As String
    
    If rsAdj.RecordCount = 0 Then Exit Sub
    fUpdate = True
    strName = txtName.Text
    strRate = txtRate.Text
    With flexAdj
        .TextMatrix(.Row, 1) = strName
        .TextMatrix(.Row, 2) = strRate
        lblName.Caption = strName
        lblValue.Caption = strRate & "%"
    End With
    arrUpdate = Add_UpdatedData_To_Array(flexAdj, arrUpdate)
End Sub

'           --- ADD UPDATED DATA TO DATABASE ---
Private Sub Add_DataUpdate_To_DB()
    Dim sTemp As String
    Dim i As Integer
    
    With rsAdj
        For i = 1 To UBound(arrUpdate)
        DoEvents
            .MoveFirst
            .Find "AdjNo=" & arrUpdate(i)(0)
            For j = 0 To .Fields.count - 1
            DoEvents
                Select Case j
                    Case 0: sTemp = "AdjNo"
                    Case 1: sTemp = "AdjName"
                    Case 2: sTemp = "AdjRate"
                End Select
                .Fields(sTemp) = arrUpdate(i)(j)
            Next j
            .Update
        Next i
    End With
End Sub

Private Sub txtRate_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        .PasswordChar = ""
        .Text = txtRate.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .FormCallkeyboard = "Other"
            frmKeyboard.Show vbModal
            txtRate.Text = .Let_Text_Input
        End With
        UpdateData
    Exit Sub
    
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub txtRate_KeyUp(KeyCode As Integer, Shift As Integer)
If fClick Then Exit Sub
    fClick = True
    Select Case KeyCode
        Case 8 'key backspace
        Case Is < 32, vbKeyDown, vbKeyUp: Exit Sub
    End Select
    UpdateData
    fClick = False
End Sub
