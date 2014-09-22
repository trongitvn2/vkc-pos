VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Më tËp tin"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6735
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
   ScaleHeight     =   3780
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdNetwork 
      Height          =   615
      Left            =   5520
      TabIndex        =   12
      Tag             =   "L7"
      Top             =   3120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "M¹ng LAN"
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
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOpen.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   645
      Left            =   5550
      TabIndex        =   11
      Tag             =   "L6"
      Top             =   990
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1138
      BTYPE           =   3
      TX              =   "&Hñy bá"
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
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOpen.frx":001C
      PICN            =   "frmOpen.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdOK 
      Height          =   585
      Left            =   5520
      TabIndex        =   10
      Tag             =   "L5"
      Top             =   240
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1032
      BTYPE           =   3
      TX              =   "&§ång ý"
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
      MICON           =   "frmOpen.frx":62D2
      PICN            =   "frmOpen.frx":62EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComDlg.CommonDialog cmndir 
      Left            =   6000
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox cboFileType 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.DriveListBox drvDriver 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2640
      TabIndex        =   3
      Top             =   3360
      Width           =   2655
   End
   Begin VB.DirListBox dFolder 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   2640
      TabIndex        =   2
      Top             =   840
      Width           =   2655
   End
   Begin VB.FileListBox fFile 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox txtFileName 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblDriver 
      Caption         =   "§Üa:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Tag             =   "L4"
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label lblFileType 
      Caption         =   "D¹ng tËp tin:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Tag             =   "L3"
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label lblFolderSelect 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   7
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblFolder 
      Caption         =   "Th­ &môc"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Tag             =   "L2"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblFileName 
      Caption         =   "&TËp tin"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Tag             =   "L1"
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim fNetwork As Boolean
Dim fso As New FileSystemObject
Dim LastDrive As String
Dim FName As String
Dim ListType() As String
Dim DefaultFileType As String
Public fOK As Byte
Dim DescArr() As String

Private Sub cboFileType_Change()
   ' fFile.Pattern = ListType(cboFileType.ListIndex)
End Sub

Private Sub cboFileType_Click()
    Call cboFileType_Change
End Sub

Private Sub cmdCancel_Click()
    FName = ""
    fOK = 0
    Unload Me
End Sub

Private Sub cmdNetwork_Click()
    fNetwork = True
    With cmndir
        .Filter = "Site group(*pgm)|*.mdb"
        .DefaultExt = "*.mdb"
        .ShowOpen
        '.InitDir = "My Computer"
        txtFileName.Text = .FileName
        txtFileName.Text = Left(.FileName, InStrRev(.FileName, "\") - 1)
    End With
    If txtFileName.Text <> "" Then cmdOK.Enabled = True
End Sub
Private Sub cmdOK_Click()
If fNetwork = True Then
    FName = txtFileName.Text
    DefaultSite = FName
    fOK = 1
    SaveSettingStr "Default Site", "Default Site", DefaultSite, myIniFile
    Unload Me
    MsgBox "Vui lßng khëi ®éng l¹i øng dông tr­íc khi ch¹y !!"
    End
Else
    CurDir = lblFolderSelect.Caption
    FName = txtFileName.Text
    If InStr(FName, "\\") Then
        FName = Replace(FName, "\\", "\")
    End If
    DefaultSite = FName
    Call CheckSitePath(DefaultSite)
    fOK = 1
    SaveSettingStr "Default Site", "Default Site", DefaultSite, myIniFile
    Unload Me
    MsgBox "Vui lßng khëi ®éng l¹i øng dông tr­íc khi ch¹y !!"
    End
End If
End Sub

Private Sub dFolder_Change()
    fFile.Path = dFolder.Path
    lblFolderSelect = dFolder.Path
    txtFileName.Text = dFolder.Path
End Sub

Private Sub dFolder_Click()
    Call dFolder_Change
End Sub

Private Sub drvDriver_Change()
    Dim d
    Set d = fso.GetDrive(Left(drvDriver.Drive, InStr(drvDriver.Drive, ":")))
    If d.DriveType = 1 And Not d.IsReady Then
        drvDriver.Drive = LastDrive
        MsgBox "®Üa mÒm ch­a s½n sµng.", vbExclamation
        Exit Sub
    ElseIf d.DriveType = 4 And Not d.IsReady Then
        drvDriver.Drive = LastDrive
        MsgBox "CD-ROM ch­a s½n sµng.", vbExclamation, "TouchScreen Utilities Software!!!"
        Exit Sub
    End If
    LastDrive = drvDriver.Drive
    dFolder.Path = drvDriver.Drive
End Sub

Private Sub fFile_Click()
    txtFileName.Text = fFile.FileName
End Sub

Private Sub fFile_DblClick()
    Call fFile_Click
    'Call cmdOK_Click
End Sub

Private Sub Form_Activate()
    Dim DescArr() As String
    Dim ctrl As Control
    If LngFile <> "" And Dir(LngFile) <> "" Then
       DescArr = LoadLanguage(LngFile, "#01:002:")
        If cmdCancel.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
        Me.Caption = "Më tËp tin"
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    End If
    If cboFileType.ListCount > 0 Then
        cboFileType.ListIndex = 0
    Else
        If DefaultFileType <> "" Then fFile.Pattern = DefaultFileType
    End If
End Sub

Private Sub Form_Load()
Dim ctrl As Control
    If Len(Trim(CurDir)) <> 0 Then
        drvDriver.Drive = Left(CurDir, 3)
        dFolder.Path = CurDir
    End If
    txtFileName.Text = ""
    cmdOK.Enabled = False
    lblFolderSelect.Caption = dFolder.Path
'Set Language
DescArr = LoadLanguage(LngFile, "#01:002:")
    Me.Caption = "Më tËp tin"
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl


'Set Default Data
    LastDrive = drvDriver.Drive
End Sub

Private Sub txtFileName_Change()
    If Trim(txtFileName.Text) <> "" Then
        If Dir(txtFileName.Text, vbDirectory) <> "" Then
            cmdOK.Enabled = True
        Else
            cmdOK.Enabled = False
        End If
    ElseIf UCase(Trim(Right(txtFileName.Text, 4))) = ".mdb" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
        Call cboFileType_Change
    End If
End Sub

Private Sub txtFileName_Click()
    Call txtFileName_Change
End Sub

Private Sub txtFileName_GotFocus()
    txtFileName.SelStart = 0
    txtFileName.SelLength = 9999
End Sub

Private Sub txtFileName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        fFile.SetFocus
        If fFile.ListCount > 0 Then fFile.ListIndex = 0
    End If
End Sub

Public Property Let FileType(ByVal vNewValue As Variant)
    Dim cnt As Integer
    cnt = 0
    Do While True
        DoEvents
        If InStr(vNewValue, "|") Then
            cboFileType.AddItem Left(vNewValue, InStr(vNewValue, "|") - 1), cnt
            ReDim Preserve ListType(cnt)
            vNewValue = Right(vNewValue, Len(vNewValue) - InStr(vNewValue, "|"))
            If InStr(vNewValue, ";") Then
                ListType(cnt) = Left(vNewValue, InStr(vNewValue, ";") - 1)
                vNewValue = Right(vNewValue, Len(vNewValue) - InStr(vNewValue, ";"))
            Else
                ListType(cnt) = vNewValue
                vNewValue = ""
            End If
            cnt = cnt + 1
        Else
            Exit Do
        End If
    Loop
End Property

Public Property Get FileName() As Variant
    FileName = FName
End Property

Public Property Let DefaultType(ByVal vNewValue As Variant)
    DefaultFileType = vNewValue
End Property
