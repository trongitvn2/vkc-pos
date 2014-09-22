VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMedia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Media Flag"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11760
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
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic3 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      ScaleHeight     =   675
      ScaleWidth      =   5955
      TabIndex        =   2
      Top             =   0
      Width           =   6015
      Begin VB.Label lblMediaName 
         BackColor       =   &H80000008&
         Caption         =   "MediaName"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblMediaNo 
         BackColor       =   &H80000008&
         Caption         =   "MediaNumber"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   0
         Width           =   3375
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   5850
      TabIndex        =   1
      Top             =   5850
      Width           =   5850
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   825
         Left            =   4020
         TabIndex        =   15
         Tag             =   "L9"
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1455
         BTYPE           =   14
         TX              =   "&Close"
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
         BCOLO           =   12648384
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMedia.frx":0000
         PICN            =   "frmMedia.frx":001C
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
         Height          =   825
         Left            =   2130
         TabIndex        =   14
         Tag             =   "L8"
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1455
         BTYPE           =   14
         TX              =   "Help"
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
         MICON           =   "frmMedia.frx":62B6
         PICN            =   "frmMedia.frx":62D2
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
         Height          =   825
         Left            =   210
         TabIndex        =   13
         Tag             =   "L7"
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1455
         BTYPE           =   14
         TX              =   "Save"
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
         MICON           =   "frmMedia.frx":690C
         PICN            =   "frmMedia.frx":6928
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
   Begin MSFlexGridLib.MSFlexGrid flexMedia 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   13573
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      GridLinesFixed  =   1
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab tabMedia 
      Height          =   4935
      Left            =   5820
      TabIndex        =   5
      Top             =   810
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
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
      TabCaption(0)   =   "Tû gi¸"
      TabPicture(0)   =   "frmMedia.frx":6E6C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cê ®iÒu khiÓn"
      TabPicture(1)   =   "frmMedia.frx":6E88
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frmFlag(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame frmFlag 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3615
         Index           =   0
         Left            =   -74880
         TabIndex        =   16
         Top             =   720
         Width           =   5535
         Begin VB.PictureBox Pic2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3255
            Left            =   60
            ScaleHeight     =   3195
            ScaleWidth      =   5340
            TabIndex        =   17
            Top             =   180
            Width           =   5400
            Begin VB.TextBox txtFlag 
               Alignment       =   2  'Center
               Height          =   375
               Left            =   2400
               TabIndex        =   19
               Tag             =   "4"
               Text            =   "00"
               Top             =   120
               Width           =   735
            End
            Begin VB.ListBox lstFlag 
               DataMember      =   "0"
               Height          =   2535
               Index           =   0
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   18
               Top             =   600
               Width           =   5175
            End
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   5535
         Begin VB.TextBox txtMedia 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   2
            Left            =   1365
            TabIndex        =   12
            Tag             =   "3"
            Top             =   2280
            Width           =   1935
         End
         Begin VB.TextBox txtMedia 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   1365
            TabIndex        =   11
            Tag             =   "2"
            Top             =   1125
            Width           =   1575
         End
         Begin VB.TextBox txtMedia 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   1365
            TabIndex        =   10
            Tag             =   "1"
            Top             =   270
            Width           =   3660
         End
         Begin VB.Label lblMedia 
            Caption         =   "Foreign Currency Exchange &Rate:"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   9
            Tag             =   "L6"
            Top             =   1950
            Width           =   3255
         End
         Begin VB.Label lblMedia 
            Alignment       =   1  'Right Justify
            Caption         =   "S&ymbol:"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Tag             =   "L5"
            Top             =   1260
            Width           =   1215
         End
         Begin VB.Label lblMedia 
            Alignment       =   1  'Right Justify
            Caption         =   "Media &Name:"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Tag             =   "L4"
            Top             =   405
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim rsMedia As New ADODB.Recordset
    Dim fLoad As Boolean
    Dim fUpdate As Boolean
    Dim fActivate As Boolean
    Dim fFlexClick As Boolean
    Dim strTmp As String
    Dim sFormat As String
    Dim arrUpdate() As Variant
    Dim i, j As Integer
    Dim DescArr() As String

'           ------------- FORM ----------
Private Sub Form_Activate()
    Dim ctrl As Control
    
    If fActivate Then Exit Sub
    fActivate = True
    If cmdSave.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    tabMedia.TabCaption(0) = DescArr(2)
    tabMedia.TabCaption(1) = DescArr(3)
    With flexMedia
        .TextMatrix(0, 0) = "STT"
        .TextMatrix(0, 1) = DescArr(4)
        .TextMatrix(0, 2) = DescArr(5)
        .TextMatrix(0, 3) = DescArr(6)
        .TextMatrix(0, 4) = "F"
    End With
    'Gan cac gia tri co tien te vao trong list box
    
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    With flexMedia
        If Shift = 2 Then 'xac dinh cac fim duoc click: shift,ctrl,alt
            If KeyCode = vbKeyDown Then ' chon keypreview trong from =true
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    If .Row > 26 Then .TopRow = .Row - 25
                End If
                KeyCode = 0
                flexMedia_Click
            ElseIf KeyCode = vbKeyUp Then 'ctrl + keyup
                If .Row > 1 Then
                    .Row = .Row - 1
                    If .Row < .TopRow Then .TopRow = .Row
                End If
                KeyCode = 0
                flexMedia_Click
            End If
        End If
    End With
End Sub

Private Sub Form_Load()
    sFormat = "#,##0.000"
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\database.mdb", "100881administrator")
'    End If
        DescArr = LoadLanguage(LngFile, "#02:016:")

    Set rsMedia = Open_Table(cnData, "Media")
    If rsMedia.State = 0 Then Exit Sub
    Initialize
End Sub

Private Sub Initialize()
    fLoad = False: fUpdate = False: fActivate = False
    ReDim Preserve arrUpdate(0)
    SetDataInFlexGrid 'khoi tao & gan dlieu cho flexgrid
    With flexMedia
'        SetColorFlexGrid flexMedia, 1, 1, .Cols
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
        .ScrollTrack = True
    End With
    With rsMedia
        txtMedia(0).MaxLength = .Fields("MediaName").DefinedSize
        txtMedia(1).MaxLength = .Fields("Symbol").DefinedSize
        txtMedia(2).MaxLength = .Fields("FCRate").DefinedSize
        txtMedia(2).Alignment = 1
    End With
    lstFlag(0).Clear
    For i = 10 To 17
        lstFlag(0).AddItem DescArr(i)
    Next i
    SetTextNull
    LockTxtFlag
    flexMedia_Click
    fLoad = True
End Sub
'           ----------- CHECKBOX -----------
Private Sub chkDouble_Click()
    txtMedia(0).SetFocus
    txtMedia(0).SelStart = Len(txtMedia(0).Text)
End Sub
'           ------------ COMMANDBUTTON ------------
Private Sub cmdClose_Click()
    Dim res
    
    If Not fUpdate Then GoTo 1
    res = MsgBox("B¹n cã muèn l­u th«ng tin thay ®æi kh«ng?", vbYesNoCancel)
    Select Case res
        Case vbYes
                    Add_DataUpdate_To_DB
        Case vbNo:  GoTo 1
        Case vbCancel: Exit Sub
    End Select
1:
    CloseRecordset rsMedia
    Unload Me
End Sub

Private Sub cmdSave_Click()
Dim ans As Integer
ans = MsgBox("B¹n cã muèn l­u th«ng tin thay ®æi kh«ng?", vbYesNo)
    If ans = vbYes Then
        fUpdate = False
        Add_DataUpdate_To_DB
    End If
End Sub
'           ----------- FLEXGRID ----------
Private Sub flexMedia_Click()
    Dim ctrl As Control
    
    fLoad = False ' ko vao ham chkdouble_click
    If rsMedia.RecordCount = 0 Then Exit Sub
    If fFlexClick Then Exit Sub
    fFlexClick = True
    For Each ctrl In Me
    DoEvents
        With ctrl
            If .Tag <> "" Then
                If TypeOf ctrl Is TextBox Then
                    If .Tag = 3 Then
                        .Text = Reset_MaxLength(ctrl, .Index, rsMedia("FCRate").DefinedSize, flexMedia.TextMatrix(flexMedia.Row, .Tag), sFormat)
                    Else
                        .Text = flexMedia.TextMatrix(flexMedia.Row, .Tag)
                    End If
                End If
            End If
        End With
    Next ctrl
    'gan gtri da check vao txtflag
    AddValueForList txtFlag.Text, lstFlag(0)
    lblMediaNo.Caption = flexMedia.TextMatrix(flexMedia.Row, 0)
    lblMediaName.Caption = flexMedia.TextMatrix(flexMedia.Row, 1)
    fFlexClick = False
    fLoad = True
End Sub

Private Sub flexMedia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        With txtMedia(0)
            .SetFocus
            .SelStart = 0
            .SelLength = 9999
        End With
    End If
End Sub

Private Sub flexMedia_EnterCell()
    If fLoad Then flexMedia_Click
End Sub

Private Sub SetDataInFlexGrid()
    Dim irow As Integer
    
    SetHeaderForFlexGrid
    irow = 1
    With rsMedia
        If .RecordCount > 0 Then
            flexMedia.Rows = .RecordCount + 1
            .MoveFirst
            Do While Not .EOF
                DoEvents
                DataInFlex irow ' them dulieu vao tung cell trong grid
                irow = irow + 1
                .MoveNext
            Loop
        End If
    End With
End Sub

Private Sub SetHeaderForFlexGrid()
    With flexMedia
        .AllowUserResizing = flexResizeBoth
        .FocusRect = flexFocusNone 'xac dinh focus retangle cho o duoc chon
        .Cols = 5
        .Row = 0
        For i = 0 To .Cols - 1 Step 1
        DoEvents
            Select Case i
                Case 0
                        .ColWidth(i) = 520
                        .ColAlignment(i) = 4
                Case 1: .ColWidth(i) = 1470
                Case 2: .ColWidth(i) = 1300
                Case 3: .ColWidth(i) = 2500
                Case Else
                        .ColWidth(i) = 600
                        .ColAlignment(i) = 4
            End Select
        Next i
    End With
End Sub

Private Sub DataInFlex(ByVal irow As Integer)
    Dim sFieldName As String
    
    With flexMedia
        For i = 0 To .Cols - 1 Step 1
        DoEvents
            Select Case i
                Case 0: sFieldName = "MediaID"
                Case 1: sFieldName = "MediaName": txtMedia(0).Tag = 1
                Case 2: sFieldName = "Symbol": txtMedia(1).Tag = 2
                Case 3: sFieldName = "FCRate": txtMedia(2).Tag = 3
                Case Else: sFieldName = "F": txtFlag.Tag = 4
            End Select
            If sFieldName = "FCRate" Then
                sFieldName = Format(rsMedia.Fields(sFieldName), sFormat)
                'sFieldName = Format(sFieldName, sFormat)
            Else
                sFieldName = rsMedia.Fields(sFieldName)
            End If
            .TextMatrix(irow, i) = sFieldName
        Next i
    End With
End Sub

'           ------------ TEXTBOX ---------
Private Sub txtMedia_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strEnd As String
    Dim tempIndex As Integer
    
    If KeyAscii = 33 Or KeyAscii = 44 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 Then
        Select Case Index
            Case 2:    tempIndex = 0
            Case Else: tempIndex = Index + 1
        End Select
        With txtMedia(tempIndex)
            .SetFocus
            .SelStart = 0
            .SelLength = 9999
        End With
        Exit Sub
    End If
    Select Case KeyAscii
        Case Is < 32
        Case 48 To 57  ', 44, 46 -> ".",","
            If Index = 2 Then
                If Len(RemoveComma(txtMedia(Index).Text)) > rsMedia.Fields("FCRate").DefinedSize - 1 Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
        Case Else
            If Index = 2 Then KeyAscii = 0
    End Select
    
    Select Case KeyAscii
        Case 8: Exit Sub 'key backspace
        Case vbKeyDown, vbKeyUp, vbKeyLeft, vbKeyRight: Exit Sub
    End Select
End Sub

Private Sub txtMedia_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 8 'key backspace
        Case Is < 32, vbKeyDown, vbKeyUp: Exit Sub
    End Select
    UpdateData
End Sub

Private Sub txtMedia_LostFocus(Index As Integer)
    Dim sTemp As String
    
    If Index = 2 Then
        With txtMedia(Index)
            sTemp = Format(.Text, sFormat)
            sTemp = Format(sTemp, sFormat)
            .Text = Reset_MaxLength(txtMedia(Index), Index, rsMedia.Fields("FCRate").DefinedSize, sTemp, sFormat)
        End With
    End If
End Sub

Private Sub LockTxtFlag()
     txtFlag.Locked = True
End Sub

Private Sub SetTextNull()
    For i = 0 To txtMedia.count - 1
        txtMedia(i).Text = ""
    Next i
End Sub
'           ----------- LISTBOX ----------
Private Sub lstFlag_Click(Index As Integer)
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
        txtFlag.Text = FillZeroForString(BinToHex(strflag), 2)
        UpdateData
    End If
End Sub
'           ------------ UPDATE DATA --------
Private Sub UpdateData()
    Dim sTemp() As String
    Dim i As Integer
    
    If rsMedia.RecordCount = 0 Then Exit Sub
    fUpdate = True
    sTemp = SetTextTemp
    With flexMedia
        For i = 0 To .Cols - 1 Step 1
            .TextMatrix(.Row, i) = sTemp(i)
        Next i
        .Refresh
    End With
    lblMediaNo.Caption = sTemp(0)
    lblMediaName.Caption = sTemp(1)
    arrUpdate = Add_UpdatedData_To_Array(flexMedia, arrUpdate)
End Sub

Private Function SetTextTemp()
    Dim ctrl As Control
    Dim S1() As String
    Dim S2 As String
    
    ReDim Preserve S1(rsMedia.Fields.count - 1)
    S1(0) = flexMedia.TextMatrix(flexMedia.Row, 0)
    For Each ctrl In Me
    DoEvents
        With ctrl
            If .Tag <> "" Then
                If TypeOf ctrl Is TextBox Then
                    If .Tag = 3 Then
                        S2 = Format(.Text, sFormat)
                        S1(.Tag) = Format(S2, sFormat)
                    Else: S1(.Tag) = .Text
                    End If
                End If
            End If
        End With
    Next ctrl
    SetTextTemp = S1
End Function

'           ----- ADD UPDATED DATA TO DATABASE ----
Private Sub Add_DataUpdate_To_DB()
    Dim sFieldName As String
    Dim i As Integer
    With rsMedia
        For i = 1 To UBound(arrUpdate)
            DoEvents
            .MoveFirst
            .Find "MediaID='" & arrUpdate(i)(0) & "'"
            If .EOF Then Exit Sub
            For j = 0 To flexMedia.Cols - 1
                DoEvents
                Select Case j
                    Case 0: sFieldName = "MediaID"
                    Case 1: sFieldName = "MediaName"
                    Case 2: sFieldName = "Symbol"
                    Case 3: sFieldName = "FCRate"
                    Case Else: sFieldName = "F"
                End Select
                If sFieldName = "FCRate" Then
                    '.Fields(sFieldName) = RemoveComma(Trim(arrUpdate(i)(j)))
                    .Fields(sFieldName) = arrUpdate(i)(j)

                Else
                
                    .Fields(sFieldName) = arrUpdate(i)(j)
                End If
            Next j
            .Update
        Next i
    End With
End Sub
