VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmClerk 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s¸ch nh©n viªn"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11790
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
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
   ScaleHeight     =   7290
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton cmdSend 
      Height          =   765
      Left            =   5640
      TabIndex        =   22
      Top             =   6270
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1349
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
      BCOL            =   16578804
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmClerk.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.PictureBox picLabel 
      BackColor       =   &H80000008&
      Height          =   705
      Left            =   5520
      ScaleHeight     =   645
      ScaleWidth      =   6195
      TabIndex        =   12
      Top             =   0
      Width           =   6255
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Clerk Name"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1560
         TabIndex        =   14
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Clerk Code"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1560
         TabIndex        =   13
         Top             =   0
         Width           =   2895
      End
   End
   Begin TabDlg.SSTab tabClerk 
      Height          =   5055
      Left            =   5580
      TabIndex        =   3
      Top             =   1020
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "Th«ng tin nh©n viªn"
      TabPicture(0)   =   "frmClerk.frx":001C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TimerHint"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtHint"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmSetup"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Flag"
      TabPicture(1)   =   "frmClerk.frx":0038
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "frmFlag(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "frmFlag(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "picFlag"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Timer TimerHint 
         Interval        =   1000
         Left            =   -69510
         Top             =   2580
      End
      Begin VB.TextBox txtHint 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   -73350
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   2175
         Width           =   2640
      End
      Begin VB.PictureBox picFlag 
         BackColor       =   &H8000000E&
         Height          =   495
         Left            =   2280
         ScaleHeight     =   435
         ScaleWidth      =   1605
         TabIndex        =   18
         Top             =   480
         Width           =   1660
         Begin VB.OptionButton optFlag 
            BackColor       =   &H80000016&
            Caption         =   "CF-2"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   840
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   30
            Width           =   735
         End
         Begin VB.OptionButton optFlag 
            BackColor       =   &H80000016&
            Caption         =   "CF-1"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   50
            MaskColor       =   &H00FF0000&
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   30
            Width           =   735
         End
      End
      Begin VB.Frame frmSetup 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   5445
         Begin VB.TextBox txtClerk 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   1200
            MaxLength       =   30
            TabIndex        =   1
            Tag             =   "1"
            Top             =   450
            Width           =   3975
         End
         Begin VB.TextBox txtClerk 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   2
            Tag             =   "2"
            Top             =   1425
            Width           =   975
         End
         Begin VB.Label lblClerkName 
            Caption         =   "Clerk/Cashier &Name:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Tag             =   "L20"
            Top             =   240
            Width           =   3405
         End
         Begin VB.Label lblClerkCode 
            Caption         =   "&Secret Clerk/Cashier Code:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Tag             =   "L21"
            Top             =   1200
            Width           =   2085
         End
      End
      Begin VB.Frame frmFlag 
         Height          =   3375
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   5895
         Begin VB.PictureBox picList 
            Height          =   3135
            Index           =   1
            Left            =   60
            ScaleHeight     =   3075
            ScaleWidth      =   5700
            TabIndex        =   9
            Top             =   160
            Width           =   5760
            Begin VB.TextBox txtFlag 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   1
               Left            =   2400
               MaxLength       =   2
               TabIndex        =   11
               Tag             =   "4"
               Top             =   120
               Width           =   735
            End
            Begin VB.ListBox lstFlag 
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2085
               Index           =   1
               ItemData        =   "frmClerk.frx":0054
               Left            =   120
               List            =   "frmClerk.frx":0056
               Style           =   1  'Checkbox
               TabIndex        =   10
               Top             =   600
               Width           =   5415
            End
         End
      End
      Begin VB.Frame frmFlag 
         Height          =   3375
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   5895
         Begin VB.PictureBox picList 
            Height          =   3135
            Index           =   0
            Left            =   60
            ScaleHeight     =   3075
            ScaleWidth      =   5700
            TabIndex        =   5
            Top             =   160
            Width           =   5760
            Begin VB.ListBox lstFlag 
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2310
               Index           =   0
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   7
               Top             =   600
               Width           =   5415
            End
            Begin VB.TextBox txtFlag 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Index           =   0
               Left            =   2400
               TabIndex        =   6
               Tag             =   "3"
               Text            =   "Text1"
               Top             =   120
               Width           =   735
            End
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flexClerk 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   12726
      _Version        =   393216
      BackColorBkg    =   16777215
      TextStyleFixed  =   3
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
   Begin prjTouchScreen.MyButton cmdSearch 
      Height          =   765
      Left            =   7170
      TabIndex        =   23
      Top             =   6270
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1349
      BTYPE           =   14
      TX              =   "T×m kiÕm"
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
      MICON           =   "frmClerk.frx":0058
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
      Left            =   8700
      TabIndex        =   24
      Top             =   6270
      Width           =   1455
      _ExtentX        =   2566
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
      MICON           =   "frmClerk.frx":0074
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Height          =   765
      Left            =   10230
      TabIndex        =   25
      Top             =   6270
      Width           =   1455
      _ExtentX        =   2566
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
      MICON           =   "frmClerk.frx":0090
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
Attribute VB_Name = "frmClerk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim rsClerk As New ADODB.Recordset
    Dim fLoad As Boolean
    Dim fUpdate As Boolean
    Dim fActivate As Boolean
    Dim fSearch As Boolean
    Dim fFlexClick As Boolean
    Dim iResultCode As Integer
    Dim sTime As Double
    Dim arrUpdate() As Variant
    Dim i, j As Integer

Private Sub cmdSend_Click()
On Error GoTo errHdl

    If fUpdate Then
        fUpdate = False
        Add_DataUpdate_To_DB
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- cmdSend_Click"
End Sub
'           ---------- FORM -----------
Private Sub Form_Activate()
On Error GoTo errHdl

    Dim DescArr() As String
    Dim ctrl As Control
    If fActivate Then Exit Sub
    fActivate = True
    DescArr = LoadLanguage(LngFile, "#03:004:")
    If cmdSend.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(11)
    tabClerk.TabCaption(0) = DescArr(16)
    tabClerk.TabCaption(1) = DescArr(17)
    With flexClerk
        .TextMatrix(0, 0) = DescArr(37)
        .TextMatrix(0, 1) = DescArr(38)
        .TextMatrix(0, 2) = DescArr(39)
        .TextMatrix(0, 3) = DescArr(55)
        .TextMatrix(0, 4) = DescArr(56)
    End With
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    With flexClerk
        If Shift = 2 Then
            If KeyCode = vbKeyDown Then
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    If .Row > 26 Then .TopRow = .Row - 25
                End If
                KeyCode = 0
                flexClerk_Click
            ElseIf KeyCode = vbKeyUp Then
                If .Row > 1 Then
                    .Row = .Row - 1
                    If .Row < .TopRow Then .TopRow = .Row
                End If
                KeyCode = 0
                flexClerk_Click
            End If
        End If
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Form_KeyDown"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl

    Dim ctrl As Control
    
    txtHint.Visible = False
    txtHint.Text = "§ang bËt Caplock" 'arrMessage(46) & " " & vbCrLf & arrMessage(47)
    With Me
        .Width = 11880
        .Height = 7260
        .WindowState = 0
    End With
    Set rsClerk = Open_Table(cnData, "Clerk")
    If rsClerk.State = 0 Then Exit Sub
    Initialize
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Form_Load"
End Sub

Private Sub Initialize()
On Error GoTo errHdl

    Dim sFieldName As String
    
    fLoad = False: fUpdate = False: fActivate = False
    ReDim Preserve arrUpdate(0)
    SetDataInFlex
    With flexClerk
        SetColorFlexGrid flexClerk, 1, 1, .Cols
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
        .ScrollTrack = True
    End With
    With rsClerk
        If Not rsClerk Is Nothing Then
            For i = 0 To txtClerk.Count - 1
            DoEvents
                Select Case i
                    Case 0: sFieldName = "ClerkName"
                    Case 1: sFieldName = "ClerkCode"
                End Select
                txtClerk(i).MaxLength = .Fields(sFieldName).DefinedSize
            Next i
        End If
    End With
    flexClerk_Click
    fLoad = True
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Initialize"
End Sub
'           --------- COMMAND BUTTON ---------
Private Sub cmdClose_Click()
On Error GoTo errHdl

    Dim res
    
    If Not fUpdate Then GoTo 1
    res = "B¹n cã muèn l­u th«ng tin thay ®æi?" 'MyMessage.ShowMsg(arrMessage(2) & frmMainData.tvwGeneralData.Nodes(1).Text & " ?", myQuestion, myYesNoCancel, arrMessage(1))
    Select Case res
        Case vbYes
                    Add_DataUpdate_To_DB
        Case vbNo:  GoTo 1
        Case vbCancel: Exit Sub
    End Select
1:
    CloseRecordset rsClerk
    Unload Me
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- cmdClose_Click"
End Sub

Private Sub cmdSearch_Click()
On Error GoTo errHdl

    fSearch = True
    With frmFind
        .GetfSearch = 1 'xac dinh form sd frmFind la frmClerk
        Set .FormCall = Me
        .Show vbModal
    End With
    fSearch = False
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- cmdSearch_Click"
End Sub
'           -------- FLEXGRID ---------
Private Sub flexClerk_EnterCell()
On Error GoTo errHdl

    If fLoad Then flexClerk_Click
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- flexClerk_EnterCell"
End Sub

Private Sub flexClerk_Click()
On Error GoTo errHdl

    Dim sTemp() As String
    Dim ctrl As Control
    
    If fSearch Then Exit Sub
    If fFlexClick Then Exit Sub
    fFlexClick = True
    fLoad = False
    With flexClerk
        ReDim Preserve sTemp(.Cols - 1)
        For i = 1 To .Cols - 1
        DoEvents
            sTemp(i) = .TextMatrix(.Row, i)
        Next i
        For Each ctrl In Me
        DoEvents
            If ctrl.Tag <> "" And ctrl.Tag <= .Cols - 1 Then
                If TypeOf ctrl Is TextBox Then _
                    ctrl.Text = sTemp(ctrl.Tag)
            End If
        Next ctrl
        lblID.Caption = .TextMatrix(.Row, 0)
        lblName.Caption = sTemp(1)
            'gan gtri da check vao txtflag
            For j = 0 To txtFlag.Count - 1 Step 1
            DoEvents
                AddValueForList txtFlag(j).Text, lstFlag(j)
            Next j
    End With
    fLoad = True
    fFlexClick = False
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- flexClerk_Click"
End Sub

Private Sub flexClerk_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 13 Then
        With txtClerk(0)
            .SetFocus
            .SelStart = 0
            .SelLength = 9999
        End With
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- flexClerk_KeyPress"
End Sub

Private Sub SetDataInFlex()
On Error GoTo errHdl

    Dim irow As Integer
    Dim sTemp As String
    
    InitFlexGrid
    irow = 1
    With rsClerk
        If .RecordCount > 0 Then
            flexClerk.Rows = .RecordCount + 1
            Do While Not .EOF
            DoEvents
                For i = 0 To flexClerk.Cols - 1
                DoEvents
                    Select Case i
                        Case 0: sTemp = "ClerkID"
                        Case 1: sTemp = "ClerkName"
                        Case 2: sTemp = "ClerkCode"
                        Case 3 To 5: sTemp = "F" & (i - 2)
                    End Select
                    sTemp = .Fields(sTemp)
                    flexClerk.TextMatrix(irow, i) = sTemp
                Next i
                irow = irow + 1
                .MoveNext
            Loop
        End If
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- SetDataInFlex"
End Sub

Private Sub InitFlexGrid()
On Error GoTo errHdl

    With flexClerk
        .AllowUserResizing = flexResizeBoth
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .Cols = rsClerk.Fields.Count
        For i = 0 To .Cols - 1 Step 1
        DoEvents
            Select Case i
                Case 0: .ColWidth(i) = 315: .ColAlignment(i) = 4
                Case 1: .ColWidth(i) = 3210
                Case 2
                    .ColWidth(i) = 825: .ColAlignment(i) = 4
                Case Else
                        .ColWidth(i) = 405: .ColAlignment(i) = 4
            End Select
        Next i
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- InitFlexGrid"
End Sub
'           --------- LISTBOX ---------
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
    & Me.Name & "- lstFlag_Click"
End Sub
'           ------- OPTION BUTTON --------
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
    & Me.Name & "- optFlag_Click"
End Sub
'           -------- TEXTBOX ----------
Private Sub LockTextFlag()
On Error GoTo errHdl

    For i = 0 To txtFlag.Count - 1
    DoEvents
        txtFlag(i).Locked = True
    Next i
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- LockTextFlag"
End Sub

Private Sub TimerHint_Timer()
On Error GoTo errHdl

    If Timer - sTime > 3 Then
        txtHint.Visible = False
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- TimerHint_Timer"
End Sub

Private Sub txtClerk_DblClick(Index As Integer)
On Error GoTo errHdl

    If Index = 1 Then
        With frmEnterClerkCode
            Set .FormCall = Me
            .Show vbModal, Me
        End With
        If iResultCode <> CInt(txtClerk(1).Text) And iResultCode <> 0 Then
            txtClerk(1).Text = FillZeroForString(CStr(iResultCode), rsClerk.Fields("ClerkCode").DefinedSize)
            UpdateData
        End If
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- txtClerk_DblClick"
End Sub

Private Sub txtClerk_GotFocus(Index As Integer)
On Error GoTo errHdl

    If Index = 1 Then
        txtHint.Visible = True
        sTime = Timer
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- txtClerk_GotFocus"
End Sub

Private Sub txtClerk_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl

    Dim tempIndex As Integer
    If KeyAscii = 33 Or KeyAscii = 44 Then
        KeyAscii = 0
        Exit Sub
    End If
    If Index = 1 And KeyAscii <> 13 Then KeyAscii = 0
    Select Case KeyAscii
        Case 13
            If Index = 0 Then
                tempIndex = 1
            Else
                tempIndex = 0
            End If
            With txtClerk(tempIndex)
                .SetFocus
                .SelStart = 0
                .SelLength = 9999
            End With
    End Select
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- txtClerk_KeyPress"
End Sub

Private Sub txtClerk_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    Select Case KeyCode
        Case 8 'key backspace
        Case Is < 32, vbKeyDown, vbKeyUp: Exit Sub
    End Select
    UpdateData
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- txtClerk_KeyUp"
End Sub
'           ---------- UPDATE DATA ----------
Private Sub UpdateData()
On Error GoTo errHdl

    Dim sTemp() As String
    Dim i As Integer
    
    If rsClerk.RecordCount = 0 Then Exit Sub
    fUpdate = True
    sTemp = SetTextTemp
    With flexClerk
        For i = 0 To .Cols - 1 Step 1
        DoEvents
            .TextMatrix(.Row, i) = sTemp(i)
        Next i
        lblID.Caption = sTemp(0)
        lblName.Caption = sTemp(1)
    End With
    arrUpdate = Add_UpdatedData_To_Array(flexClerk, arrUpdate)
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- UpdateData"
End Sub

Private Function SetTextTemp()
On Error GoTo errHdl

    Dim s1() As String
    Dim ctrl As Control
    
    With flexClerk
        ReDim Preserve s1(rsClerk.Fields.Count - 1)
        s1(0) = .TextMatrix(.Row, 0)
        For Each ctrl In Me
        DoEvents
            If ctrl.Tag <> "" Then
                If TypeOf ctrl Is TextBox And ctrl.Tag <= .Cols - 1 Then
                    s1(ctrl.Tag) = ctrl.Text
                End If
            End If
        Next ctrl
    End With
    SetTextTemp = s1
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- SetTextTemp"
End Function
'           ----- ADD UPDATED DATA TO DATABASE ----
Private Sub Add_DataUpdate_To_DB()
On Error GoTo errHdl

    Dim sFieldName As String
    Dim i As Integer
    
    With rsClerk
        For i = 1 To UBound(arrUpdate)
            DoEvents
            .MoveFirst
            .Find "ClerkID=" & arrUpdate(i)(0)
            For j = 0 To flexClerk.Cols - 1
            DoEvents
                Select Case j
                    Case 0: sFieldName = "ClerkID"
                    Case 1: sFieldName = "ClerkName"
                    Case 2: sFieldName = "ClerkCode"
                    Case 3 To 5: sFieldName = "F" & (j - 2)
                End Select
                .Fields(sFieldName) = arrUpdate(i)(j)
            Next j
            .Update
        Next i
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Add_DataUpdate_To_DB"
End Sub

Public Property Let Get_Code(ByVal vNewValue As Integer)
    iResultCode = vNewValue
End Property
