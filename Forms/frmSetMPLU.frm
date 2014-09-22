VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSetMPLU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Danh s¸ch nguyªn liÖu"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15015
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
   ScaleHeight     =   11145
   ScaleWidth      =   15015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   1455
      Left            =   9990
      ScaleHeight     =   1395
      ScaleWidth      =   5115
      TabIndex        =   10
      Top             =   60
      Width           =   5175
      Begin VB.Label lblNo 
         BackColor       =   &H80000008&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   120
         Width           =   5535
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000008&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   5655
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flex 
      Height          =   11115
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   19606
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      TextStyleFixed  =   3
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TabDlg.SSTab tabSetMPLU 
      Height          =   8295
      Left            =   9930
      TabIndex        =   13
      Tag             =   "L7"
      Top             =   2640
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
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
      TabCaption(0)   =   "Th«ng tin nguyªn liÖu"
      TabPicture(0)   =   "frmSetMPLU.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdPrint"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdClose"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdHelp"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSearch"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDelete"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdAdd"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frmSetup"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
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
         Height          =   4035
         Left            =   110
         TabIndex        =   14
         Top             =   585
         Width           =   4725
         Begin VB.TextBox txtPLU 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   2
            Left            =   1200
            MaxLength       =   8
            TabIndex        =   3
            Tag             =   "2"
            Top             =   2040
            Width           =   2265
         End
         Begin VB.TextBox txtPLU 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   4
            Left            =   3570
            TabIndex        =   5
            Tag             =   "4"
            Top             =   2850
            Width           =   795
         End
         Begin VB.TextBox txtPLU 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   0
            Left            =   1200
            TabIndex        =   1
            Tag             =   "0"
            Top             =   375
            Width           =   2280
         End
         Begin VB.TextBox txtPLU 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   3
            Left            =   1200
            TabIndex        =   4
            Tag             =   "3"
            Top             =   2835
            Width           =   945
         End
         Begin VB.TextBox txtPLU 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Index           =   1
            Left            =   1200
            TabIndex        =   2
            Tag             =   "1"
            Top             =   1200
            Width           =   3120
         End
         Begin VB.Label lblPLU 
            Alignment       =   1  'Right Justify
            Caption         =   "Cost:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   20
            Tag             =   "L24"
            Top             =   2115
            Width           =   645
         End
         Begin VB.Label lblMinstock 
            Caption         =   "Kho tèi thiÓu:"
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
            Left            =   2250
            TabIndex        =   19
            Tag             =   "L15"
            Top             =   2940
            Width           =   1335
         End
         Begin VB.Label lblPLU 
            Alignment       =   1  'Right Justify
            Caption         =   "PLU &Code:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   150
            TabIndex        =   17
            Tag             =   "L12"
            Top             =   450
            Width           =   1065
         End
         Begin VB.Label lblPLU 
            Alignment       =   1  'Right Justify
            Caption         =   "Unit:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   525
            TabIndex        =   16
            Tag             =   "L14"
            Top             =   3030
            Width           =   645
         End
         Begin VB.Label lblPLU 
            Alignment       =   1  'Right Justify
            Caption         =   "PLU &Name:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   15
            Tag             =   "L13"
            Top             =   1275
            Width           =   1065
         End
      End
      Begin prjTouchScreen.MyButton cmdAdd 
         Height          =   1065
         Left            =   60
         TabIndex        =   6
         Top             =   4770
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1879
         BTYPE           =   14
         TX              =   "Thªm míi"
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
         MICON           =   "frmSetMPLU.frx":001C
         PICN            =   "frmSetMPLU.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdDelete 
         Height          =   1065
         Left            =   1680
         TabIndex        =   7
         Top             =   4770
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1879
         BTYPE           =   14
         TX              =   "&Xãa"
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
         MICON           =   "frmSetMPLU.frx":048A
         PICN            =   "frmSetMPLU.frx":04A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdSearch 
         Height          =   1065
         Left            =   3360
         TabIndex        =   8
         Top             =   4770
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1879
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
         MICON           =   "frmSetMPLU.frx":0AE0
         PICN            =   "frmSetMPLU.frx":0AFC
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
         Height          =   1065
         Left            =   1680
         TabIndex        =   18
         Top             =   6120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1879
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
         MICON           =   "frmSetMPLU.frx":1136
         PICN            =   "frmSetMPLU.frx":1152
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
         Height          =   1065
         Left            =   3360
         TabIndex        =   9
         Top             =   6090
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1879
         BTYPE           =   14
         TX              =   "&§ãng"
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
         MICON           =   "frmSetMPLU.frx":178C
         PICN            =   "frmSetMPLU.frx":17A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdPrint 
         Height          =   1065
         Left            =   60
         TabIndex        =   21
         Top             =   6120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1879
         BTYPE           =   14
         TX              =   "In DS Nguyªn liÖu"
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
         MICON           =   "frmSetMPLU.frx":7A42
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
End
Attribute VB_Name = "frmSetMPLU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim rsSetMPLU As New ADODB.Recordset
    Dim arrUpdate() As Variant
    Dim arrDelete() As String, arrAddNew() As String
    Dim arrTemp() As String, ECRName As String
    Dim fLoad As Boolean, fUpdate As Boolean
    Dim fActivate As Boolean, fAddNew As Boolean
    Dim i, j As Integer

Private Sub cmdPrint_Click()
On Error GoTo Handle
    Dim SQL As String
    Dim cmd As New ADODB.Command
    Dim iReport As CRAXDDRT.Report
    SQL = "SELECT SetMPLU.PLUCode, SetMPLU.PLUName, SetMPLU.Unit, SetMPLU.Cost, SetMPLU.MinStock" & _
               " FROM SetMPLU"
    
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set crsetPLUList = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crsetPLUList
        .Database.AddADOCommand cnData, cmd
        .txtPluCode.SetUnboundFieldSource "{ado.PLUCode}"
        .txtPluName.SetUnboundFieldSource "{ado.PLUName}"
        .txtUnit.SetUnboundFieldSource "{ado.Unit}"
        .txtPrice1.SetUnboundFieldSource "{ado.Cost}"
        .txtMinstock.SetUnboundFieldSource "{ado.MinStock}"
        With .txtPrice1
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    Set iReport = crsetPLUList
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With

Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & " - cmdPrint_Click"
End Sub

'             -------------- FORM -------------
Private Sub Form_Activate()
On Error GoTo errHdl

    Dim DescArr() As String
    Dim ctrl As Control
        
    If fActivate Then Exit Sub
    fActivate = True
    DescArr = LoadLanguage(LngFile, "#01:013:")
    If cmdClose.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    With flex
        .TextMatrix(0, 0) = DescArr(12) 'Ma
        .TextMatrix(0, 1) = DescArr(13) 'Ten
        .TextMatrix(0, 2) = DescArr(24) 'Gia
        .TextMatrix(0, 3) = DescArr(14) 'DVT
        .TextMatrix(0, 4) = DescArr(15) 'Ton toi thieu
    End With
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & " Form_Activate"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    With flex
        If Shift = 2 Then 'xac dinh cac fim duoc click: shift,ctrl,alt
            If KeyCode = vbKeyDown Then ' chon keypreview trong from =true
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                    If .Row > 26 Then .TopRow = .Row - 25
                End If
                KeyCode = 0
                flex_Click
            ElseIf KeyCode = vbKeyUp Then 'ctrl + keyup
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
    MsgBox Err.Number & " - " & Err.Description & "Form_KeyDown"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl

'    Me.Height = 7800
'    Me.Width = 11500
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsSetMPLU = Open_Table(cnData, "SetMPLU")
    If Not rsSetMPLU Is Nothing Then
        With rsSetMPLU
            For i = 0 To txtPLU.count - 4
            DoEvents
                txtPLU(i).MaxLength = .Fields(i).DefinedSize
            Next i
        End With
    End If
    ReDim Preserve arrUpdate(0)
    ReDim Preserve arrAddNew(0)
    ReDim Preserve arrDelete(0)
    Initialize

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & " Form_Load"
End Sub

Private Sub Initialize()
On Error GoTo errHdl

    fLoad = False: fUpdate = False: fActivate = False
    txtPLU(0).Enabled = False
    SetDataInFlex
    With flex
'        SetColorFlexGrid flex, 1, 0, .Cols
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
        .ScrollTrack = True
    End With
    fLoad = True
    flex_Click

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & "Initialize"
End Sub
'            ----------FLEXGRID-------------
Private Sub flex_Click()
On Error GoTo errHdl

    Dim sTemp() As String
    Dim ctrl As Control
                    
    If fAddNew Then Exit Sub
    fLoad = False
    With flex
        If .TextMatrix(1, 0) = "" Then Exit Sub
        ReDim Preserve sTemp(.Cols - 1)
        For i = 0 To .Cols - 1
        DoEvents
            sTemp(i) = .TextMatrix(.Row, i)
        Next i
    End With
    For Each ctrl In Me
    DoEvents
        With ctrl
            If .Tag <> "" Then
                If TypeOf ctrl Is TextBox Then
                    Select Case .Index
                    
                    Case 2
                        .Text = Format(sTemp(.Tag), "#,##0")
                    Case Else
                        .Text = sTemp(.Tag)
                    End Select
                    
                ElseIf TypeOf ctrl Is ComboBox Then
                    .ListIndex = sTemp(.Tag) - 1
                End If
            End If
        End With
    Next ctrl
    lblNo.Caption = sTemp(0)
    lblName.Caption = sTemp(1)
    arrTemp = SetTextTemp
    fLoad = True

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & "flex_Click"
End Sub

Private Sub flex_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    With txtPLU(1)
        .SetFocus
        .SelStart = 0
        .SelLength = 9999
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & "flex_KeyPress"
End Sub

Private Sub flex_EnterCell()
On Error GoTo errHdl

    If fLoad Then flex_Click

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub SetDataInFlex()
On Error GoTo errHdl

    Dim irow As Integer

    SetHeaderFlexGrid
    irow = 1
    With rsSetMPLU
        .Requery
        .Sort = "PLUCode ASC"
        If .RecordCount > 0 Then
            flex.Rows = .RecordCount + 1
            .MoveFirst
            Do While Not .EOF
            DoEvents
                For i = 0 To flex.Cols - 1
                DoEvents
                    If IsNull(.Fields(i)) Then
                        flex.TextMatrix(irow, i) = "2"
                        .Fields(i) = "2"
                    Else
                        flex.TextMatrix(irow, i) = .Fields(i)
                        
                    End If
                Next i
                irow = irow + 1
                .MoveNext
            Loop
        End If
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & "SetDataInFlex"
End Sub

Private Sub SetHeaderFlexGrid()
On Error GoTo errHdl

    Dim fFound As Boolean
    
    fFound = False
    For i = 0 To rsSetMPLU.Fields.count - 1
    DoEvents
        If rsSetMPLU.Fields(i).name = "Cost" Then
            fFound = True
            Exit For
        End If
    Next i
    If Not fFound Then
        cnData.Execute "ALTER TABLE SetMPLU " _
                             & "ADD COLUMN Cost Double;"
    End If
    rsSetMPLU.Requery
    With flex
        .AllowUserResizing = flexResizeBoth
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionByRow
        .Cols = rsSetMPLU.Fields.count
        .ColWidth(0) = 1200
        .ColWidth(1) = 4200
        .ColWidth(2) = 1200: .ColAlignment(2) = 4
        .ColWidth(3) = 1200: .ColAlignment(3) = 4
        .ColWidth(4) = 1200: .ColAlignment(3) = 4
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & " SetHeaderFlexGrid"
End Sub

'             -------------- COMMAND BUTTON -------------
Private Sub cmdAdd_Click()
On Error GoTo errHdl

    fAddNew = True
    frmAddMPLU.Show vbModal, Me
    Get_Array_AddNew
    If UBound(arrAddNew) > 0 Then
        fUpdate = True
        For i = 0 To UBound(arrAddNew)
        DoEvents
            For j = 0 To UBound(arrDelete)
            DoEvents
                If arrAddNew(i) = arrDelete(j) Then
                    arrDelete(j) = ""
                End If
            Next j
        Next i
    End If
    fAddNew = False
    flex_Click

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & "cmdAdd_Click"
End Sub

Private Sub Get_Array_AddNew() 'append addnew records from frmAddNewPLU to arrAddNew()
On Error GoTo errHdl

    Dim arrTemp() As String
    Dim iTemp As Integer
    
    arrTemp = frmAddMPLU.Get_AddNewRecords
    iTemp = UBound(arrTemp)
    ReDim Preserve arrAddNew(UBound(arrAddNew) + iTemp)
    For i = 1 To UBound(arrTemp)
    DoEvents
        arrAddNew(UBound(arrAddNew) - iTemp + i) = arrTemp(i)
    Next i

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & "Get_Array_AddNew"
End Sub

Private Sub Get_Array_Delete() 'append deleted records from frmDeletePLU to arrDelete()
On Error GoTo errHdl

    Dim arrTemp() As String
    Dim iTemp As Integer
    
    arrTemp = frmDeleteMPLU.Get_DeleteRecords
    iTemp = UBound(arrTemp)
    ReDim Preserve arrDelete(UBound(arrDelete) + iTemp)
    For i = 1 To UBound(arrTemp)
    DoEvents
        arrDelete(UBound(arrDelete) - iTemp + i) = arrTemp(i)
    Next i

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & " Get_Array_Delete"
End Sub

Private Sub cmddelete_Click()
On Error GoTo errHdl

    fAddNew = True
    frmDeleteMPLU.Show vbModal, Me
    Get_Array_Delete
    If UBound(arrDelete) > 0 Then
        fUpdate = True
        For i = 0 To UBound(arrDelete)
        DoEvents
            For j = 0 To UBound(arrAddNew)
            DoEvents
                If arrDelete(i) = arrAddNew(j) Then
                    arrAddNew(j) = ""
                End If
            Next j
        Next i
    End If
    fAddNew = False

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & "cmdDelete_Click"
End Sub

Private Sub cmdClose_Click()
On Error GoTo errHdl

    Dim res
    
    If Not fUpdate Then GoTo 1
    res = MsgBox("B¹n cã muèn l­u th«ng tin thay ®æi kh«ng?", vbYesNo)
    Select Case res
        Case vbYes: Add_DataUpdate_To_DB
        Case vbNo:  GoTo 1
        Case vbCancel: Exit Sub
    End Select
1:
    CloseRecordset rsSetMPLU
    Unload Me

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdSearch_Click()
On Error GoTo errHdl

    fAddNew = True
    With frmFind
        .GetfSearch = 3
        Set .FormCall = Me
        .Show vbModal
    End With
    fAddNew = False

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub txtPLU_DblClick(Index As Integer)
    On Error GoTo Handle
    With frmKeyboard.txtInput
        
        .PasswordChar = ""
        .Text = txtPLU(Index).Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtPLU(Index).Text = .Let_Text_Input
        End With
        UpdateData
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtPLU_DblClick"
End Sub

'               ------- TEXTBOX --------
Private Sub txtPLU_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 13 Then
        Select Case Index
            Case 1
                    With txtPLU(2)
                        .SetFocus
                        .SelStart = 0
                        .SelLength = 9999
                    End With
            Case 2
                    With txtPLU(3)
                        .SetFocus
                        .SelStart = 0
                        .SelLength = 9999
                    End With
            Case 3
                    With txtPLU(4)
                        .SetFocus
                        .SelStart = 0
                        .SelLength = 9999
                    End With
            Case 4
                    With flex
                        .SetFocus
                        If .Row < .Rows - 1 Then .Row = .Row + 1
                    End With
        End Select
    End If

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & "txtPLU_KeyPress"
End Sub

Private Sub txtPLU_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    Select Case KeyCode
        Case 8 'key backspace
        Case Is < 32, vbKeyDown, vbKeyUp: Exit Sub
    End Select
    UpdateData

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

'            ----------- COMBOBOX ----------
Private Sub cboGroup_Click()
On Error GoTo errHdl

    If fLoad Then UpdateData

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cboGroup_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 13 Then
        With txtPLU(3)
            .SetFocus
            .SelStart = 0
            .SelLength = 9999
        End With
    End If

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

'           -----------UPDATE DATA-------------
Private Sub UpdateData()
On Error GoTo errHdl

    Dim strPLU() As String
    Dim i As Integer
    
    fUpdate = True
    strPLU = SetTextTemp
    With flex
        For i = 1 To UBound(strPLU) Step 1
        DoEvents
            .TextMatrix(.Row, i) = strPLU(i)
        Next i
        .Refresh
        lblNo.Caption = .TextMatrix(.Row, 0)
        lblName.Caption = strPLU(1)
    End With
    arrUpdate = Add_UpdatedData_To_Array(flex, arrUpdate)

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & "UpdateData"
End Sub

Private Function SetTextTemp()
On Error GoTo errHdl

    Dim ctrl As Control
    Dim S1() As String

    ReDim Preserve S1(flex.Cols - 1)
    For Each ctrl In Me
    DoEvents
        With ctrl
            If .Tag <> "" Then
                If TypeOf ctrl Is TextBox Then
                    If .Index = 0 Then
                        S1(.Tag) = FillZeroForString(.Text, .MaxLength)
                    Else
                        S1(.Tag) = .Text
                    End If
                ElseIf TypeOf ctrl Is ComboBox Then
                    S1(.Tag) = FillZeroForString(.ListIndex + 1, 2)
                End If
            End If
        End With
    Next ctrl
    SetTextTemp = S1

Exit Function
errHdl:
    MsgBox Err.Number & " - " & Err.Description & "SetTextTemp"
End Function

'           ----- ADD UPDATED DATA TO DATABASE ----
Private Sub Add_DataUpdate_To_DB()
On Error GoTo errHdl

    Dim sSQL As String
    Dim k As Integer, L As Integer
    Dim i As Integer
    
    With flex
        For k = 1 To UBound(arrAddNew)
            DoEvents
            For L = 1 To .Rows - 1
            DoEvents
                If .TextMatrix(L, 0) = arrAddNew(k) Then
                    .Row = L
                    arrUpdate = Add_UpdatedData_To_Array(flex, arrUpdate)
                End If
            Next L
        Next k
    End With
    With rsSetMPLU
        For i = 1 To UBound(arrUpdate)
            DoEvents
            If .RecordCount = 0 Then
                .addNew
            Else
                .MoveFirst
                .Find "PLUCode='" & arrUpdate(i)(0) & "'"
                If .EOF Then .addNew 'addNew rows duoc them moi
            End If
            For j = 0 To flex.Cols - 1
            DoEvents
                .Fields(j) = arrUpdate(i)(j)
            Next j
            .Update
            '.Requery
        Next i
        For i = 1 To UBound(arrDelete)
            DoEvents
            sSQL = "Delete from SetMPLU where PLUCode='" & arrDelete(i) & "'"
            cnData.Execute sSQL
        Next i
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & "Add_DataUpdate_To_DB"
End Sub

