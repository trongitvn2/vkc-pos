VERSION 5.00
Begin VB.Form frmSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "cmdCancel"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
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
   ScaleHeight     =   3855
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdNetwork 
      Height          =   615
      Left            =   5430
      TabIndex        =   12
      Top             =   3150
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
      MICON           =   "frmSave.frx":0000
      PICN            =   "frmSave.frx":001C
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
      TabIndex        =   11
      Top             =   1080
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
      MICON           =   "frmSave.frx":046E
      PICN            =   "frmSave.frx":048A
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
      TabIndex        =   10
      Top             =   180
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
      MICON           =   "frmSave.frx":6724
      PICN            =   "frmSave.frx":6740
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
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
      Width           =   2415
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
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2415
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
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmSave"
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
Dim Descarr() As String

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

Private Sub cmdOK_Click()
    
        Dim OKCancel As Integer
        txtFileName.Text = txtFileName.Text & _
        Mid(ListType(cboFileType.ListIndex), _
        InStr(ListType(cboFileType.ListIndex), "."), _
        Len(ListType(cboFileType.ListIndex)) - _
        InStr(ListType(cboFileType.ListIndex), ".") + 1)
    FName = lblFolderSelect & "\" & Left(txtFileName.Text, Len(txtFileName.Text) - 4)
    If InStr(FName, "\\") Then
        FName = Replace(FName, "\\", "\")
    End If
        iFileTypeValue = cboFileType.ListIndex
        If Dir(RemoveExtFile(FName) & "\Data", vbDirectory) = "" Then
           MkDir (RemoveExtFile(FName))
'           MkDir (RemoveExtFile(FName) & "\Data")
        End If
    If Dir(FName, vbDirectory) <> "" Then
        OKCancel = MsgBox(FName & " " & "B¹n cã muèn thay thÕ tËp tin ®ã kh«ng?" & vbCrLf, vbOKCancel)
        If OKCancel = 1 Then
            Kill RemoveExtFile(FName)
            Dim fso As New FileSystemObject
            fso.CopyFile App.Path & "\Data\Database.mdb", RemoveExtFile(FName) & "\Data\Database.mdb"
            Unload Me
        Else
            FName = ""
        End If
    Else
        If Dir(RemoveExtFile(FName), vbDirectory) <> "" Then
            fso.DeleteFolder RemoveExtFile(FName)
            fso.CopyFile App.Path & "\Data\Databse.mdb", RemoveExtFile(FName) & "\Data"
        Else
            
            fso.CopyFile App.Path & "\Data\Databse.mdb", RemoveExtFile(FName) & "\Data"
        End If
    DefaultSite = FName
    SaveSettingStr "Default Site", "Defaul tSite", DefaultSite, myIniFile

    iFileTypeValue = cboFileType.ListIndex
    Unload Me
    End If
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
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Descarr(Mid(ctrl.Tag, 2))
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
 
    Descarr = LoadLanguage(LngFile, "#01:002:")
    If cmdCancel.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
        Me.Caption = "Më tËp tin"
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Descarr(Mid(ctrl.Tag, 2))
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

