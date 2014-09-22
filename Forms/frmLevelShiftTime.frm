VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmLevelShiftTime 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Level Shift Time"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9660
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
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
      Height          =   660
      Left            =   4920
      ScaleHeight     =   600
      ScaleWidth      =   4755
      TabIndex        =   1
      Top             =   0
      Width           =   4815
      Begin VB.Label lblNo 
         BackColor       =   &H80000008&
         Caption         =   "Number-01"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1200
         TabIndex        =   3
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000008&
         Caption         =   "00:00"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1200
         TabIndex        =   2
         Top             =   285
         Width           =   3120
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
      Height          =   1245
      Left            =   5040
      TabIndex        =   0
      Top             =   3600
      Width           =   4500
      Begin prjTouchScreen.MyButton cmdSave 
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
         _extentx        =   2355
         _extenty        =   1508
         btype           =   14
         tx              =   "&L­u"
         enab            =   -1  'True
         font            =   "frmLevelShiftTime.frx":0000
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   0
         micon           =   "frmLevelShiftTime.frx":0028
         picn            =   "frmLevelShiftTime.frx":0046
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdHelp 
         Height          =   855
         Left            =   1560
         TabIndex        =   11
         Top             =   240
         Width           =   1335
         _extentx        =   2355
         _extenty        =   1508
         btype           =   14
         tx              =   "&Gióp ®ì"
         enab            =   -1  'True
         font            =   "frmLevelShiftTime.frx":058A
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   0
         micon           =   "frmLevelShiftTime.frx":05B2
         picn            =   "frmLevelShiftTime.frx":05D0
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   855
         Left            =   3000
         TabIndex        =   10
         Top             =   240
         Width           =   1335
         _extentx        =   2355
         _extenty        =   1508
         btype           =   14
         tx              =   "&§ãng"
         enab            =   -1  'True
         font            =   "frmLevelShiftTime.frx":0C0C
         coltype         =   2
         focusr          =   -1  'True
         bcol            =   14215660
         bcolo           =   14215660
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   0
         micon           =   "frmLevelShiftTime.frx":0C34
         picn            =   "frmLevelShiftTime.frx":0C52
         umcol           =   -1  'True
         soft            =   0   'False
         picpos          =   0
         ngrey           =   0   'False
         fx              =   0
         hand            =   0   'False
         check           =   0   'False
         value           =   0   'False
      End
   End
   Begin TabDlg.SSTab tab1 
      Height          =   2655
      Left            =   5040
      TabIndex        =   4
      Top             =   840
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
      TabPicture(0)   =   "frmLevelShiftTime.frx":6EEE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTime(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTime(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dtpTime(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dtpTime(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin MSComCtl2.DTPicker dtpTime 
         Height          =   315
         Index           =   0
         Left            =   1425
         TabIndex        =   8
         Top             =   675
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   20774915
         UpDown          =   -1  'True
         CurrentDate     =   38610
      End
      Begin MSComCtl2.DTPicker dtpTime 
         Height          =   315
         Index           =   1
         Left            =   1425
         TabIndex        =   9
         Top             =   1350
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "HH:mm"
         Format          =   20774915
         UpDown          =   -1  'True
         CurrentDate     =   38610
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         Caption         =   "&End Time:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Tag             =   "L8"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         Caption         =   "&Start Time:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Tag             =   "L7"
         Top             =   795
         Width           =   1095
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   7095
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   12515
      _Version        =   393216
      BackColorBkg    =   16777215
      TextStyleFixed  =   3
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
End
Attribute VB_Name = "frmLevelShiftTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===FORMS:MENU LEVEL SHIFT TIME, PRICE LEVEL SHIFT TIME ====
Option Explicit
    Dim arrData() As Variant, arrUpdate() As Variant
    Dim fUpdate As Boolean, fLoad As Boolean
    Dim fActivate As Boolean
    Dim softID As String
    Private Type LevelTime
        Week As String
        Menu As String
        Price As String
    End Type

Private Sub cmdSave_Click()
On Error GoTo errHdl

    Dim sNameFile As String
    Dim spathFile As String
    Dim flagNet As Boolean
    Dim iEcrID As Integer
    Dim iSite As Integer
    Dim iNet As Integer
                
    iSite = frmNetworkSelect.Site_Position
    iNet = frmNetworkSelect.Net_Position
    If Left(softID, 3) <> "ECR" Then
        flagNet = True
        sNameFile = "Network"
    Else
        flagNet = False
        sNameFile = Left(softID, InStr(softID, vbTab) - 1)
        With mySite(iSite).sNetwork(iNet)
            For k = 0 To .EcrCount - 1 Step 1
            DoEvents
                If StrComp("ECR" & FillZeroForString(CStr(.nECR(k).ID), 3), sNameFile, 1) = 0 Then
                    iEcrID = k
                    Exit For
                End If
            Next k
        End With
    End If
    spathFile = spath_DB & "\Data\" & sNameFile
    If fUpdate Then
        Add_DataUpdate_To_File spathFile
    End If
    If Dir(spathFile & ".Sed") <> "" Then
        If flagNet Then
            SendAllData iSite, iNet, 0, NetSelect, spathFile & ".Sed"
        Else
            SendAllData iSite, iNet, iEcrID, EcrSelect, spathFile & ".Sed"
        End If
        fUpdate = False
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - cmdSave_Click"
End Sub
'           ---------- FORM ----------
Private Sub Form_Activate()
On Error GoTo errHdl

    Dim DescArr() As String
    Dim ctrl As Control
    Dim iCount As Byte
    
    If fActivate Then Exit Sub
    fActivate = True
    DescArr = LoadLanguage(LngFile, "#04:001:")
    If cmdSave.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    iCount = 0
    For i = 1 To UBound(DescArr)
    DoEvents
        Select Case i
            Case 6
                tab1.TabCaption(0) = DescArr(i)
            Case 9, 10, 13, 14, 15
                flex.TextMatrix(0, iCount) = DescArr(i)
                iCount = iCount + 1
            Case 11
                If Right(softID, 4) = "5302" Then
                    flex.TextMatrix(0, iCount) = DescArr(i)
                    iCount = iCount + 1
                End If
            Case 12
                If Right(softID, 4) = "5301" Then
                    flex.TextMatrix(0, iCount) = DescArr(i)
                    iCount = iCount + 1
                End If
        End Select
    Next
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
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
                    If .Row < 20 Then .TopRow = .Row - 19
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

    With Me
        .Height = 5085 ' 5250
        .Width = 9750 ' 9780
        .WindowState = 0
    End With
    Initalize
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Form_Load"
End Sub

Private Sub Initalize()
On Error GoTo errHdl

    fLoad = False: fUpdate = False: fActivate = False
    ReDim Preserve arrUpdate(0)
    ReDim Preserve arrData(0)
    SetDataFromEDF
    
    If UBound(arrData) = 0 Then Exit Sub
    SetDataInFlex
    With flex
        SetColorFlexGrid flex, 1, 1, .Cols
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
        .ScrollTrack = True
    End With
    fLoad = True
    flex_Click
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Initalize"
End Sub
'           ----------- COMMANDBUTTON ----------
Private Sub cmdClose_Click()
On Error GoTo errHdl

    Dim sNameFile As String, spathFile As String
    Dim sTemp As String
    Dim res
        
    res = MsgBox(arrMessage(2) & " '" & sTemp & "'?", vbYesNoCancel)
    Select Case res
        Case vbNo:     GoTo 1
        Case vbYes:    Add_DataUpdate_To_File spathFile
        Case vbCancel: Exit Sub
    End Select
1:  Unload Me

    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - cmdClose_Click"
End Sub
'           ----------- DATETIMEPICKER -----------
Private Sub dtpTime_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 13 Then
        Select Case Index
            Case 0: dtpTime(1).SetFocus
            Case 1: dtpTime(0).SetFocus
        End Select
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - dtpTime_KeyPress"
End Sub

Private Sub dtpTime_Change(Index As Integer)
On Error GoTo errHdl
    UpdateData
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - dtpTime_Change"
End Sub
'           ----------- FLEXGRID -----------
Private Sub flex_Click()
On Error GoTo errHdl

    fLoad = False
    With flex
        If .TextMatrix(.Row, 0) = "" Then SetTextNull: Exit Sub
        lblNo.Caption = .TextMatrix(.Row, 0)
        lblName.Caption = .TextMatrix(.Row, 1) & " " & _
                          .TextMatrix(.Row, 2) & " " & _
                          .TextMatrix(.Row, 3)
        dtpTime(0).Value = .TextMatrix(.Row, .Cols - 2)
        dtpTime(1).Value = .TextMatrix(.Row, .Cols - 1)
    End With
    fLoad = True
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - flex_Click"
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

    Dim arrRow() As LevelTime
    Dim arrMenu() As String
    Dim arrWeek(1 To 7) As String
    Dim arrPrice(1 To 3) As String
    Dim iCount As Integer
    
    arrWeek(1) = "Sunday": arrWeek(2) = "Monday"
    arrWeek(3) = "Tuesday": arrWeek(4) = "Wednesday"
    arrWeek(5) = "Thursday": arrWeek(6) = "Friday"
    arrWeek(7) = "Saturday"
    arrPrice(1) = "Price 1": arrPrice(2) = "Price 2"
    arrPrice(3) = "Price 3"
        ReDim Preserve arrMenu(4)
        ReDim Preserve arrRow(84)
        arrMenu(1) = "Lunch Menu": arrMenu(2) = "DinnerMenu"
        arrMenu(3) = "Holiday 1": arrMenu(4) = "Holiday 2"
    iCount = 1
    For i = 1 To UBound(arrWeek)
    DoEvents
        For j = 1 To UBound(arrMenu)
        DoEvents
            For k = 1 To UBound(arrPrice)
            DoEvents
                arrRow(iCount).Week = arrWeek(i)
                arrRow(iCount).Menu = arrMenu(j)
                arrRow(iCount).Price = arrPrice(k)
                iCount = iCount + 1
            Next k
        Next j
    Next i
    SetHeaderFlexGrid
    With flex
        .Rows = UBound(arrRow) + 1
        For i = 1 To .Rows - 1
        DoEvents
            .TextMatrix(i, 0) = arrData(i)(0)
            .TextMatrix(i, 1) = arrRow(i).Week
            .TextMatrix(i, 2) = arrRow(i).Menu
            .TextMatrix(i, 3) = arrRow(i).Price
            .TextMatrix(i, 4) = ChangeToHour(Trim(arrData(i)(1)))
            .TextMatrix(i, 5) = ChangeToHour(Trim(arrData(i)(2)))
        Next i
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - SetDataInFlex"
End Sub

Private Sub SetHeaderFlexGrid()
On Error GoTo errHdl

    With flex
        .Cols = 6
        .AllowUserResizing = flexResizeBoth
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        For i = 0 To .Cols - 1
        DoEvents
            Select Case i
            Case 0: .ColWidth(i) = 315: .ColAlignment(0) = 4
            Case 1: .ColWidth(i) = 810
            Case 2: .ColWidth(i) = 1110
            Case 3: .ColWidth(i) = 735
            Case Else
                    .ColWidth(i) = 795
            End Select
        Next i
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - SetHeaderFlexGrid"
End Sub

Private Sub SetTextNull()
On Error GoTo errHdl
    lblNo.Caption = "Number-01"
    lblName.Caption = "Sunday Lunch Menu Price 1"
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - SetTextNull"
End Sub
'           ----------- UPDATE DATA --------
Private Sub UpdateData()
On Error GoTo errHdl
    
    If flex.TextMatrix(1, 0) = "" Then Exit Sub
    fUpdate = True
    With flex
        .TextMatrix(.Row, .Cols - 2) = Format(dtpTime(0).Value, "HH:mm")
        .TextMatrix(.Row, .Cols - 1) = Format(dtpTime(1).Value, "HH:mm")
        lblName.Caption = .TextMatrix(.Row, 1) & " " _
                        & .TextMatrix(.Row, 2) & " " _
                        & .TextMatrix(.Row, 3)
    End With
    Add_UpdatedData_To_Array
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - UpdateData"
End Sub


Private Sub Add_UpdatedData_To_Array()
On Error GoTo errHdl

    Dim arrFlex(1) As String
    Dim flag As Boolean
    Dim sNo As String

    flag = False
    With flex
        sNo = .TextMatrix(.Row, 0)
        If UBound(arrUpdate) < 1 Then
1:
            ReDim Preserve arrUpdate(UBound(arrUpdate) + 1)
            arrFlex(0) = ChangeHourToString(.TextMatrix(.Row, .Cols - 2))
            arrFlex(1) = ChangeHourToString(.TextMatrix(.Row, .Cols - 1))
            arrUpdate(UBound(arrUpdate)) = arrFlex()
            Exit Sub
        End If

        For i = 1 To UBound(arrUpdate)
        DoEvents
            If InStr(arrUpdate(i)(0), sNo) <> 0 Then
                arrFlex(0) = ChangeHourToString(.TextMatrix(.Row, .Cols - 2))
                arrFlex(1) = ChangeHourToString(.TextMatrix(.Row, .Cols - 1))
                arrUpdate(i) = arrFlex()
                flag = True
                Exit For
            End If
        Next i
        If Not flag Then GoTo 1
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Add_UpdatedData_To_Array"
End Sub
