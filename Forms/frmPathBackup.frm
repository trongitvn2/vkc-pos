VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPathBackup 
   Caption         =   "§­êng dÉn sao chÐp d÷ liÖu"
   ClientHeight    =   4140
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6870
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
   ScaleHeight     =   4140
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmndir 
      Left            =   5640
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFileName 
      Height          =   435
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   2415
   End
   Begin VB.FileListBox fFile 
      Height          =   2115
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   2415
   End
   Begin VB.DirListBox dFolder 
      Height          =   2130
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   2655
   End
   Begin VB.DriveListBox drvDriver 
      Height          =   345
      Left            =   2640
      TabIndex        =   4
      Top             =   3480
      Width           =   2655
   End
   Begin VB.ComboBox cboFileType 
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3480
      Width           =   2295
   End
   Begin prjTouchScreen.MyButton cmdNetwork 
      Height          =   615
      Left            =   5430
      TabIndex        =   0
      Top             =   3270
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "&M¹ng LAN"
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
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPathBackup.frx":0000
      PICN            =   "frmPathBackup.frx":001C
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
      Height          =   675
      Left            =   5490
      TabIndex        =   1
      Top             =   1200
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   1191
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
      MICON           =   "frmPathBackup.frx":046E
      PICN            =   "frmPathBackup.frx":048A
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
      Height          =   645
      Left            =   5460
      TabIndex        =   2
      Top             =   300
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1138
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
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPathBackup.frx":6724
      PICN            =   "frmPathBackup.frx":6740
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblFileName 
      Caption         =   "&TËp tin"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label lblFolder 
      Caption         =   "Th­ &môc"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2640
      TabIndex        =   11
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lblFolderSelect 
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label lblFileType 
      Caption         =   "D¹ng tËp tin:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label lblDriver 
      Caption         =   "§Üa:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   3240
      Width           =   2175
   End
End
Attribute VB_Name = "frmPathBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim fso As New FileSystemObject
Dim LastDrive As String
Dim FName As String
Dim ListType() As String
Dim DefaultFileType As String
Dim iFileTypeValue As Long
Dim isLoaded As Boolean
Dim DescArr() As String

Private Sub cboFileType_Change()

    fFile.Pattern = ListType(cboFileType.ListIndex)
End Sub

Private Sub cboFileType_Click()
    Call cboFileType_Change
End Sub

Private Sub cmdCancel_Click()
    FName = ""
    Unload Me
End Sub

Private Sub cmdNetwork_Click()
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
    On Error GoTo Handle
        Dim OKCancel As Integer
        txtFileName.Text = txtFileName.Text
    FName = txtFileName.Text
    If InStr(FName, "\\") Then
        FName = Replace(FName, "\\", "\")
    End If
        DefaultSite = FName
    If Dir(FName & "\Database.mdb", vbDirectory) = "" Then
            Dim fso As New FileSystemObject
            fso.CopyFile WorkingFolder & "\Database.mdb", FName & "\Database.mdb"
        DefaultSite = FName
        SaveSettingStr "SYSTEM", "Backup Site", DefaultSite, myIniFile
    
        iFileTypeValue = cboFileType.ListIndex
        Unload Me
    End If
    SaveSettingStr "SYSTEM", "Backup Site", DefaultSite, myIniFile
     Unload Me
     MsgBox "Vui lßng khëi ®éng l¹i øng dông tr­íc khi sö dông"
     End
Exit Sub
Handle:
    MsgBox Err.Description & " cmdOK_Click"
End Sub

Private Sub dFolder_Change()
    fFile.Path = dFolder.Path
    lblFolderSelect = dFolder.Path
End Sub

Private Sub dFolder_Click()
    Call dFolder_Change
End Sub

Private Sub drvDriver_Change()
    Dim d
    Set d = fso.GetDrive(Left(drvDriver.Drive, InStr(drvDriver.Drive, ":")))
    If d.DriveType = 4 Then
        drvDriver.Drive = LastDrive
        Exit Sub
    End If
    LastDrive = drvDriver.Drive
    dFolder.Path = drvDriver.Drive
End Sub

Private Sub fFile_Click()
    txtFileName.Text = fFile.FileName
End Sub

Private Sub Form_Activate()
Dim ctrl As Control
    If isLoaded Then Exit Sub
    isLoaded = True
    If cboFileType.ListCount > 0 Then
        cboFileType.ListIndex = 0
    Else
        If DefaultFileType <> "" Then fFile.Pattern = DefaultFileType
    End If
    LastDrive = drvDriver.Drive
        Me.Caption = "Më tËp tin"
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
End Sub

Private Sub Form_Load()
Dim ctrl As Control

    isLoaded = False
    If Len(Trim(CurDir)) <> 0 Then
        drvDriver.Drive = Left(CurDir, 3)
        dFolder.Path = CurDir
    End If
    txtFileName.Text = ""
    cmdOK.Enabled = False
    lblFolderSelect.Caption = dFolder.Path
    DescArr = LoadLanguage(LngFile, "#01:002:")
    If cmdCancel.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
        Me.Caption = "Më tËp tin"
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl

End Sub
Private Sub txtFileName_Change()
    If Trim(txtFileName.Text) <> "" Then
        cmdOK.Enabled = True
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Sub txtFileName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        fFile.SetFocus
        If fFile.ListCount > 0 Then fFile.ListIndex = 0
    End If
End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("/") Or KeyAscii = Asc("\") Or _
        KeyAscii = Asc(":") Or KeyAscii = Asc("?") Or _
        KeyAscii = Asc("""") Or KeyAscii = Asc("<") Or _
        KeyAscii = Asc(">") Or KeyAscii = Asc("|") Then
        KeyAscii = 0
    End If
End Sub

Public Property Get FileTypeValue() As Long
    FileTypeValue = iFileTypeValue
End Property

Public Property Let FileType(ByVal vNewValue As String)
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



