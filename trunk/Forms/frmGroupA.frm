VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmGroupA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Group A"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9960
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tabGroup 
      Height          =   4095
      Left            =   4800
      TabIndex        =   5
      Top             =   2040
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Set up"
      TabPicture(0)   =   "frmGroupA.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "frmSetup"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Flag"
      TabPicture(1)   =   "frmGroupA.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picFlag"
      Tab(1).Control(1)=   "frmFlag(0)"
      Tab(1).Control(2)=   "frmFlag(1)"
      Tab(1).ControlCount=   3
      Begin VB.Frame frmFlag 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Index           =   1
         Left            =   -74880
         TabIndex        =   18
         Top             =   960
         Width           =   5535
         Begin VB.PictureBox picList 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2820
            Index           =   1
            Left            =   60
            ScaleHeight     =   2760
            ScaleWidth      =   5340
            TabIndex        =   19
            Top             =   160
            Width           =   5400
            Begin VB.TextBox txtFlag 
               Alignment       =   2  'Center
               Height          =   375
               Index           =   1
               Left            =   2280
               MaxLength       =   2
               TabIndex        =   21
               Top             =   120
               Width           =   975
            End
            Begin VB.ListBox lstFlag 
               Height          =   2085
               Index           =   1
               ItemData        =   "frmGroupA.frx":0038
               Left            =   120
               List            =   "frmGroupA.frx":003A
               Style           =   1  'Checkbox
               TabIndex        =   20
               Top             =   600
               Width           =   5175
            End
         End
      End
      Begin VB.Frame frmFlag 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Index           =   0
         Left            =   -74880
         TabIndex        =   14
         Top             =   960
         Width           =   5535
         Begin VB.PictureBox picList 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2820
            Index           =   0
            Left            =   60
            ScaleHeight     =   2760
            ScaleWidth      =   5340
            TabIndex        =   15
            Top             =   160
            Width           =   5400
            Begin VB.ListBox lstFlag 
               Height          =   2085
               Index           =   0
               ItemData        =   "frmGroupA.frx":003C
               Left            =   120
               List            =   "frmGroupA.frx":003E
               Style           =   1  'Checkbox
               TabIndex        =   17
               Top             =   600
               Width           =   5055
            End
            Begin VB.TextBox txtFlag 
               Height          =   375
               Index           =   0
               Left            =   2280
               MaxLength       =   2
               TabIndex        =   16
               Top             =   120
               Width           =   1095
            End
         End
      End
      Begin VB.PictureBox picFlag 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -72840
         ScaleHeight     =   435
         ScaleWidth      =   1605
         TabIndex        =   11
         Top             =   480
         Width           =   1660
         Begin VB.OptionButton optFlag 
            BackColor       =   &H80000016&
            Caption         =   "GF-1"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   50
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   30
            Width           =   735
         End
         Begin VB.OptionButton optFlag 
            BackColor       =   &H80000016&
            Caption         =   "GF-2"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   840
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   30
            Width           =   735
         End
      End
      Begin VB.Frame frmSetup 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   110
         TabIndex        =   8
         Top             =   310
         Width           =   4875
         Begin VB.TextBox txtGroup 
            Height          =   330
            Index           =   1
            Left            =   1440
            TabIndex        =   23
            Top             =   1920
            Width           =   1815
         End
         Begin VB.ComboBox cboMainGroup 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtGroup 
            Height          =   330
            Index           =   0
            Left            =   1440
            TabIndex        =   1
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label lblInvent 
            Alignment       =   1  'Right Justify
            Caption         =   "&Inventory Constant:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Tag             =   "L11"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label lblLink 
            Caption         =   "&Link to Main Group-A/Major Group:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Tag             =   "L10"
            Top             =   840
            Width           =   3615
         End
         Begin VB.Label lblGroupName 
            Alignment       =   1  'Right Justify
            Caption         =   "Group &Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Tag             =   "L9"
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox picLabel 
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
      Height          =   735
      Left            =   4800
      ScaleHeight     =   675
      ScaleWidth      =   6795
      TabIndex        =   3
      Top             =   0
      Width           =   6855
      Begin VB.Label lblName 
         BackColor       =   &H80000008&
         Caption         =   "Group Name"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label lblNo 
         BackColor       =   &H80000008&
         Caption         =   "Group No"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   1560
         TabIndex        =   6
         Top             =   75
         Width           =   1815
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flexGroupA 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   10821
      _Version        =   393216
      BackColorBkg    =   16777215
      TextStyleFixed  =   3
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Height          =   1000
      Left            =   4800
      TabIndex        =   4
      Top             =   600
      Width           =   4065
      Begin prjLPVECR.MyButton cmdSend 
         Height          =   735
         Left            =   75
         TabIndex        =   24
         Tag             =   "L4"
         Top             =   180
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "Save"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   8.25
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
         MICON           =   "frmGroupA.frx":0040
         PICN            =   "frmGroupA.frx":005C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjLPVECR.MyButton cmdHelp 
         Height          =   735
         Left            =   1380
         TabIndex        =   25
         Tag             =   "L5"
         Top             =   180
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "Help"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   8.25
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
         MICON           =   "frmGroupA.frx":05A0
         PICN            =   "frmGroupA.frx":05BC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjLPVECR.MyButton cmdClose 
         Cancel          =   -1  'True
         Height          =   735
         Left            =   2700
         TabIndex        =   26
         Tag             =   "L6"
         Top             =   180
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1296
         BTYPE           =   14
         TX              =   "Close"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   8.25
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
         MICON           =   "frmGroupA.frx":0BF6
         PICN            =   "frmGroupA.frx":0C12
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
   End
End
Attribute VB_Name = "frmGroupA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsGroupA As New ADODB.Recordset
    Dim fLoad As Boolean, fUpdate As Boolean
    Dim fActivate As Boolean
    Dim fFlexClick As Boolean
    Dim arrUpdate() As Variant

Private Sub cmdSend_Click()
Dim res
If fUpdate Then
      res = MsgBox(arrMessage(2) & " kh«ng ?", vbYesNo)
    Select Case res
        Case vbYes
            Add_DataUpdate_To_DB
        Case vbNo:  Exit Sub
    End Select
End If
End Sub

'           ------------ FORM ----------
Private Sub Form_Activate()
On Error GoTo errHdl

    Dim DescArr() As String
    Dim Ctrl As Control
    
    If rsGroupA.State = 0 Then
        MsgBox arrMessage(40), myInformation, myClose, arrMessage(1)
        cmdClose_Click
        Exit Sub
    End If
    If fActivate Then Exit Sub
    fActivate = True
    DescArr = LoadLanguage(LngFile, "#03:007:")
    If cmdSend.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    tabGroup.TabCaption(0) = DescArr(7)
    tabGroup.TabCaption(1) = DescArr(8)
    With flexGroupA
        .TextMatrix(0, 0) = DescArr(12)
        .TextMatrix(0, 1) = DescArr(13)
        .TextMatrix(0, 2) = "Mµu s¾c"
        .ColAlignment(1) = 2
        
    End With
    For Each Ctrl In Me
    DoEvents
        If Left(Ctrl.Tag, 1) = "L" Then Ctrl.Caption = DescArr(Mid(Ctrl.Tag, 2))
    Next Ctrl
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    With flexGroupA
        If Shift = 2 Then
            If KeyCode = vbKeyDown Then
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    If .Row > 23 Then .TopRow = .Row - 22
                End If
                KeyCode = 0
                flexGroupA_Click
            ElseIf KeyCode = vbKeyUp Then
                If .Row > 1 Then
                    .Row = .Row - 1
                    If .Row < .TopRow Then .TopRow = .Row
                End If
                KeyCode = 0
                flexGroupA_Click
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
If cnnDataGeneral.State = 0 Then
    Set cnnDataGeneral = Get_Connection(WorkingFolder & "\Data\Database\Maindata.mdb", "100881")
End If
    Set rsGroupA = Open_Table(cnnDataGeneral, "GroupA")
    If rsGroupA.State = 0 Then Exit Sub
    Initialize
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Form_Load"
End Sub

Private Sub Initialize()
On Error GoTo errHdl

    Dim Ctrl As Control
    Dim flag As Boolean
    
    fLoad = False: fUpdate = False: fActivate = False
    ReDim Preserve arrUpdate(0)
    SetDataInFlex
    For i = 0 To txtGroup.Count - 1
    DoEvents
        Select Case i
            Case 0: txtGroup(0).MaxLength = rsGroupA.Fields("GroupName").DefinedSize
            
        End Select
    Next i
     flag = True
   
    SetCombo "MainGroup", cboMainGroup, "MainGroupName", flag
    HideControl flag
        
    With flexGroupA
        SetColorFlexGrid flexGroupA, 1, 1, .Cols
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
        .ScrollTrack = True
    End With
    optFlag(0).Value = True
    flexGroupA_Click
    fLoad = True
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Initialize"
End Sub
'           ---------- COMBOBOX ---------
Private Sub cboMaingroup_Click()
On Error GoTo errHdl

    If fLoad Then UpdateData  'update dlieu tren grid & csdl
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - cboMainGroup_Click"
End Sub

Private Sub cboMainGroup_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    Dim tempIndex As Integer
    If KeyAscii = 13 Then
         tempIndex = 0
       
        With txtGroup(tempIndex)
            .SetFocus
            .SelStart = 0
            .SelLength = 9999
        End With
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - cboMainGroup_KeyPress"
End Sub
'           --------- COMMANDBUTTON --------
Private Sub cmdClose_Click()
On Error GoTo errHdl

    Dim res
    
    If Not fUpdate Then GoTo 1
      res = MsgBox(arrMessage(2), vbYesNoCancel)
    Select Case res
        Case vbYes
            Add_DataUpdate_To_DB
        Case vbNo:  GoTo 1
        Case vbCancel: Exit Sub
    End Select
1:
    CloseRecordset rsGroupA
    Unload Me
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - cmdClose_Click"
End Sub
'           ---------- FLEXGRID ----------
Private Sub flexGroupA_Click()
On Error GoTo errHdl

    Dim sTemp() As String
    Dim Ctrl As Control
    
    fLoad = False
    If rsGroupA.RecordCount = 0 Then Exit Sub
    If fFlexClick Then Exit Sub
    fFlexClick = True
    With flexGroupA
        ReDim Preserve sTemp(.Cols - 1)
        For i = 1 To .Cols - 1
        DoEvents
            sTemp(i) = .TextMatrix(.Row, i)
        Next i
        For Each Ctrl In Me
        DoEvents
            With Ctrl
                If .Tag <> "" Then
                    If TypeOf Ctrl Is TextBox And .Tag <= flexGroupA.Cols - 1 Then
                        .Text = sTemp(.Tag)
                    ElseIf TypeOf Ctrl Is ComboBox Then
                        If .ListCount <> 0 Then
                            .ListIndex = sTemp(.Tag)
                        End If
                    End If
                End If
            End With
        Next Ctrl
        lblNo.Caption = .TextMatrix(.Row, 0)
        lblName.Caption = sTemp(1)
    End With
    fFlexClick = False
    fLoad = True
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - flexGroupA_Click"
End Sub

Private Sub flexGroupA_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 13 Then
        With txtGroup(0)
            .SetFocus
            .SelStart = 0
            .SelLength = 9999
        End With
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - flexGroupA_KeyPress"
End Sub

Private Sub flexGroupA_EnterCell()
On Error GoTo errHdl

    If fLoad Then flexGroupA_Click
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - flexGroupA_EnterCell"
End Sub

Private Sub SetDataInFlex()
On Error GoTo errHdl

    Dim irow As Integer
    Dim sTemp As String
    
    SetHeaderFlexGrid
    irow = 1
    With rsGroupA
        If .RecordCount > 0 Then
            flexGroupA.Rows = .RecordCount + 1
            Do While Not .EOF
            DoEvents
                For i = 0 To flexGroupA.Cols - 1
                DoEvents
                    Select Case i
                        Case 0: sTemp = "GroupNo"
                        Case 1: sTemp = "GroupName": txtGroup(0).Tag = 1
                        Case 2: sTemp = "LinkMainGroup": cboMainGroup.Tag = 2
                        Case 3: sTemp = "F1": txtFlag(0).Tag = 3
                        Case 4: sTemp = "F2": txtFlag(1).Tag = 4
                        Case 5: sTemp = "InventConstant": txtGroup(1).Tag = 5
                    End Select
                    flexGroupA.TextMatrix(irow, i) = .Fields(sTemp)
                Next i
                irow = irow + 1
                .MoveNext
            Loop
        End If
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - SetDataInFlex"
End Sub

Private Sub SetHeaderFlexGrid()
On Error GoTo errHdl

    With flexGroupA
        .Cols = rsGroupA.Fields.Count
        .AllowUserResizing = flexResizeBoth
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        
        For i = 0 To rsGroupA.Fields.Count - 1
        DoEvents
            Select Case i
                Case 0: .ColWidth(i) = 375: .ColAlignment(i) = 4
                Case 1: .ColWidth(i) = 2550: .ColAlignment(i) = 1
                Case 2: .ColWidth(i) = 1200: .ColAlignment(i) = 4
            End Select
        Next i
        
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - SetHeaderFlexGrid"
End Sub
'           --------- TEXTBOX ---------
Private Sub txtGroup_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 33 Or KeyAscii = 44 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 Then
        Select Case Index
            Case 0: cboMainGroup.SetFocus
            Case 1
                    With txtGroup(0)
                        .SetFocus
                        .SelStart = 0
                        .SelLength = 9999
                    End With
        End Select
        Exit Sub
    End If
    If Index = 1 Then
        Select Case KeyAscii
            Case Is < 32: Exit Sub
            Case 48 To 57
                If Len(RemoveComma(txtGroup(Index).Text)) > rsGroupA.Fields("InventConstant").DefinedSize - 1 Then _
                    KeyAscii = 0
            Case Else: KeyAscii = 0
        End Select
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - txtGroup_KeyPress"
End Sub

Private Sub txtGroup_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    Select Case KeyCode
        Case 8 'key backspace
        Case Is < 32, vbKeyDown, vbKeyUp: Exit Sub
    End Select
    UpdateData
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - txtGroup_KeyUp"
End Sub

Private Sub SetTextNull()
On Error GoTo errHdl

    txtGroup(0).Text = ""
    txtGroup(1).Text = ""
    lblNo.Caption = "01"
    lblName.Caption = "Group-A-01"
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - SetTextNull"
End Sub
'           ---------- LISTBOX ---------
Private Sub lstFlag_Click(Index As Integer)
On Error GoTo errHdl

    Dim strflag As String
    If fLoad Then
        strflag = ""
        For i = 0 To lstFlag(Index).ListCount - 1
        DoEvents
            If lstFlag(Index).Selected(i) Then
                  strflag = strflag & "1"
            Else: strflag = strflag & "0"
            End If
        Next i
        txtFlag(Index).Text = FillZeroForString(BinToHex(strflag), 2)
        UpdateData
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - lstFlag_Click"
End Sub
'           ---------- OPTIONBUTTON ---------
Private Sub optFlag_Click(Index As Integer)
On Error GoTo errHdl

    For i = 0 To optFlag.Count - 1
    DoEvents
        If Index = i Then
              frmFlag(i).Visible = True
        Else: frmFlag(i).Visible = False
        End If
    Next i
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - optFlag_Click"
End Sub
'           ----------- UPDATE DATA ---------
Private Sub UpdateData()
On Error GoTo errHdl

    Dim sTemp() As String
    Dim i As Integer
    
    If rsGroupA.RecordCount = 0 Then Exit Sub
    fUpdate = True
    sTemp = SetTextTemp
    With flexGroupA
        For i = 1 To UBound(sTemp) Step 1
        DoEvents
            .TextMatrix(.Row, i) = sTemp(i)
        Next i
        lblNo = sTemp(0)
        lblName = sTemp(1)
    End With
    arrUpdate = Add_UpdatedData_To_Array(flexGroupA, arrUpdate)
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - UpdateData"
End Sub

Private Function SetTextTemp()
On Error GoTo errHdl

    Dim Ctrl As Control
    Dim s1() As String
    
    ReDim Preserve s1(rsGroupA.Fields.Count - 1)
    s1(0) = flexGroupA.TextMatrix(flexGroupA.Row, 0)
    For Each Ctrl In Me
    DoEvents
        With Ctrl
            If .Tag <> "" Then
                If TypeOf Ctrl Is TextBox And .Tag <= flexGroupA.Cols - 1 Then
                    s1(.Tag) = .Text
                ElseIf TypeOf Ctrl Is ComboBox Then
                    If .ListCount <> 0 Then
                        s1(.Tag) = FillZeroForString(.ListIndex, 2)
                    End If
                End If
            End If
        End With
    Next Ctrl
    SetTextTemp = s1
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - SetTextTemp"
End Function

Private Sub HideControl(flag As Boolean)
On Error GoTo errHdl

    lblInvent.Visible = Not flag
    txtGroup(1).Visible = Not flag 'txtinventoryConstant
    tabGroup.TabVisible(1) = Not flag
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - HideControl"
End Sub
'           --- ADD UPDATED DATA TO FILE .SED ---

'           --- ADD UPDATED DATA TO DATABASE ---
Private Sub Add_DataUpdate_To_DB()
On Error GoTo errHdl

    Dim sFieldName As String
    Dim i As Integer
    
    With rsGroupA
        For i = 1 To UBound(arrUpdate)
            DoEvents
            .MoveFirst
            .Find "GroupNo=" & arrUpdate(i)(0)
            For j = 0 To .Fields.Count - 1
            DoEvents
                Select Case j
                    Case 0: sFieldName = "GroupNo"
                    Case 1: sFieldName = "GroupName"
                    Case 2: sFieldName = "LinkMainGroup"
                    Case 3: sFieldName = "InventConstant"
                End Select
                .Fields(sFieldName) = arrUpdate(i)(j)
            Next j
            .Update
        Next i
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Add_DataUpdate_To_DB"
End Sub
