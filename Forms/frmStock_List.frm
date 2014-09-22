VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStock_List 
   Caption         =   "Danh môc kho"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
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
   ScaleHeight     =   5055
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4080
      TabIndex        =   3
      Top             =   3840
      Width           =   4620
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   735
         Left            =   3120
         TabIndex        =   4
         Tag             =   "L5"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmStock_List.frx":0000
         PICN            =   "frmStock_List.frx":001C
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
         Left            =   1680
         TabIndex        =   5
         Tag             =   "L4"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         MICON           =   "frmStock_List.frx":62B6
         PICN            =   "frmStock_List.frx":62D2
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
         Left            =   240
         TabIndex        =   6
         Tag             =   "L3"
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
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
         MICON           =   "frmStock_List.frx":690C
         PICN            =   "frmStock_List.frx":6928
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
      Height          =   675
      Left            =   3960
      ScaleHeight     =   615
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.Label lblName 
         BackColor       =   &H80000008&
         Caption         =   "Group Name"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   330
         Width           =   1695
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000008&
         Caption         =   "Group No"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1200
         TabIndex        =   1
         Top             =   45
         Width           =   1695
      End
   End
   Begin TabDlg.SSTab tab1 
      Height          =   2655
      Left            =   4080
      TabIndex        =   7
      Top             =   1080
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   4683
      _Version        =   393216
      Tabs            =   1
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
      TabCaption(0)   =   "Set up"
      TabPicture(0)   =   "frmStock_List.frx":6E6C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblGroupName"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtData"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
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
         TabIndex        =   8
         Top             =   960
         Width           =   3735
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
      Height          =   4935
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   8705
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
Attribute VB_Name = "frmStock_List"
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

Private Sub cmdSave_Click()
On Error GoTo errHdl
        fUpdate = False
        Add_DataUpdate_To_DB
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - cmdSave_Click"
End Sub
'           ------------ FORM ----------
Private Sub Form_Activate()
On Error GoTo errHdl

    Dim DescArr() As String
    Dim ctrl As Control

    DescArr = LoadLanguage(LngFile, "#04:004:")
    If cmdSave.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    If res.State = 0 Then
        cmdClose_Click
        Exit Sub
    End If
    If fActivate Then Exit Sub
    fActivate = True
    Me.Caption = DescArr(1)
    tab1.TabCaption(0) = DescArr(7)
    flex.TextMatrix(0, 0) = DescArr(6)
    flex.TextMatrix(0, 1) = DescArr(7)
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then
            ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
        End If
    Next ctrl
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Form_Activate"
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
    & Me.Name & " - Form_KeyDown"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl

    Dim sTableName As String
    With Me
        .Height = 5355
        .Width = 8805
        .WindowState = 0
    End With
    Set res = Open_Table(cnData, "Stock_List")
    If res.State = 0 Then Exit Sub
    Initalize
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Form_Load"
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
    txtData.MaxLength = res.Fields("Stock_Name").DefinedSize
    flex_Click
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Initalize"
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
    & Me.Name & " - cmdClose_Click"
End Sub
'           ----------- FLEXGRID ---------
Private Sub flex_Click()
On Error GoTo errHdl

    fLoad = False ' ko vao ham chkdouble_click
    With flex
        If res.RecordCount > 0 Then
            lblNo.Caption = .TextMatrix(.Row, 0)
            lblName.Caption = .TextMatrix(.Row, 1)
            txtData.Text = .TextMatrix(.Row, 1)
        Else
            SetTextNull
        End If
    End With
    fLoad = True
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - flex_Click"
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
    & Me.Name & " - flex_KeyPress"
End Sub

Private Sub flex_EnterCell()
On Error GoTo errHdl

    If fLoad Then flex_Click
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - flex_EnterCell"
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
                flex.TextMatrix(irow, 0) = !ID
                flex.TextMatrix(irow, 1) = !Stock_Name
                irow = irow + 1
                .MoveNext
            Loop
        End If
    End With
    flex.ColSel = flex.Cols - 1
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - SetDataInFlex"
End Sub

Private Sub SetHeaderFlexGrid()
On Error GoTo errHdl

    With flex
        .Cols = res.Fields.count
        .AllowUserResizing = flexResizeBoth
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .ColWidth(0) = 400: .ColAlignment(0) = 4
        .ColWidth(1) = 3210: .ColAlignment(1) = 1
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - SetHeaderFlexGrid"
End Sub
'           ---------- TEXTBOX --------
Private Sub txtData_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    Dim iSelStart As Integer
    Dim temp As String
    
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
    & Me.Name & " - txtData_KeyPress"
End Sub

Private Sub SetTextNull()
On Error GoTo errHdl

    txtData.Text = ""
    lblNo.Caption = "01"
    lblName.Caption = "Tªn kho"
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - SetTextNull"
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
    & Me.Name & " - txtData_KeyUp"
End Sub
'           ---------- UPDATE DATA --------
Private Sub UpdateData()
On Error GoTo errHdl

    Dim strName As String
        
    If res.RecordCount = 0 Then Exit Sub
    If fClick Then Exit Sub
    fClick = True
    fUpdate = True
    strName = txtData.Text
    With flex
        .TextMatrix(.Row, 1) = strName
        lblName.Caption = strName
    End With
    arrUpdate = Add_UpdatedData_To_Array(flex, arrUpdate)
    fClick = False
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - UpdateData"
End Sub

'           ----- ADD UPDATED DATA TO DATABASE ----
Private Sub Add_DataUpdate_To_DB()
On Error GoTo errHdl

    Dim i, J As Integer

    With res
        For i = 1 To UBound(arrUpdate)
            DoEvents
            .MoveFirst
            .Find "ID=" & arrUpdate(i)(0)
            For J = 0 To .Fields.count - 1
            DoEvents
                .Fields(J) = arrUpdate(i)(J)
            Next J
            .Update
        Next i
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Add_DataUpdate_To_DB"
End Sub


