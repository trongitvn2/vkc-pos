VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRightSelection 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QuyÒn sö dông"
   ClientHeight    =   9825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15150
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9825
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.CheckBox Check1 
         Height          =   375
         Left            =   11880
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Caption         =   "Danh s¸ch ng­êi dïng"
         ForeColor       =   &H00FF0000&
         Height          =   8895
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   6375
         Begin MSDataGridLib.DataGrid dtgUser 
            Height          =   8055
            Left            =   120
            TabIndex        =   20
            Top             =   600
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   14208
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   26
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
      End
      Begin VB.TextBox txtID 
         Height          =   390
         Left            =   13560
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtUserName 
         Height          =   495
         Left            =   8640
         TabIndex        =   11
         Top             =   2040
         Width           =   3135
      End
      Begin VB.TextBox txtRetypeCode 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   8640
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox txtUserCode 
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   8640
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   720
         Width           =   3135
      End
      Begin VB.ComboBox cboLevel 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   12120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2040
         Width           =   2055
      End
      Begin MSComctlLib.TreeView tvwRightAccess 
         Height          =   5535
         Left            =   6600
         TabIndex        =   1
         Top             =   4080
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   9763
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "VNI-Times"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjTouchScreen.MyButton cmddelete 
         Height          =   615
         Left            =   9360
         TabIndex        =   3
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "Xãa"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRightSelection.frx":0000
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
         Cancel          =   -1  'True
         Height          =   615
         Left            =   13680
         TabIndex        =   4
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "&Tho¸t"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRightSelection.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdUpdate 
         Height          =   615
         Left            =   7920
         TabIndex        =   5
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "CËp nhËt"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRightSelection.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAdd 
         Height          =   615
         Left            =   6480
         TabIndex        =   15
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "Thªm "
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRightSelection.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdPass 
         Height          =   615
         Left            =   10800
         TabIndex        =   16
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "MËt khÈu"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRightSelection.frx":0070
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
         Height          =   615
         Left            =   12240
         TabIndex        =   18
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1085
         BTYPE           =   5
         TX              =   "L­u"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRightSelection.frx":008C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.Label lblmatch 
         Caption         =   "( X¸c nhËn ch­a trïng khíp)"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   12000
         TabIndex        =   21
         Top             =   1200
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "CÊp ®é ng­êi dïng"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   12000
         TabIndex        =   14
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Ph©n quyÒn "
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   6600
         TabIndex        =   13
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Tªn ng­êi dïng:"
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6720
         TabIndex        =   12
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "X¸c nhËn m·:"
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6720
         TabIndex        =   10
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label lblUserCode 
         Caption         =   "M· ®¨ng nhËp:"
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   6720
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "danh s¸ch ng­êi dïng"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   1080
         TabIndex        =   2
         Top             =   120
         Width           =   10455
      End
   End
End
Attribute VB_Name = "frmRightSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Private Type Node_Tree
        ID As String
        Desc As String
    End Type
    Dim DescArrTree() As Node_Tree
    
    Dim rsRight As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim fLoad As Boolean
    Dim iupdate As Boolean

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        txtuserCode.PasswordChar = ""
    Else
        txtuserCode.PasswordChar = "*"
    End If
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo Handle
    If UserLevel <> 1 Then Exit Sub
        Call Init_AddNew
        Lock_text (False)
        iupdate = True
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdAdd_Click"
End Sub

Private Sub cmddelete_Click()
    On Error GoTo Handle
    If UserLevel <> 1 Then Exit Sub
        DoEvents
        If MsgBox("B¹n cã ch¾c ch¾n muèn xãa ng­êi dïng tªn: " & txtUserName.Text & " kh«ng?", vbYesNo) = vbYes Then
            With rsTemp
            .Find "ID='" & txtID.Text & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    .Delete adAffectCurrent
                End If
                Delay (100)
            End With
            Set dtgUser.DataSource = rsTemp
            iupdate = True
        End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmddelete_Click"
End Sub

Private Sub cmdPass_Click()
    With frmChangePassword
        .Let_Pass_Call = txtuserCode.Text
        .Show vbModal
    End With
End Sub

Private Sub cmdSave_Click()
On Error GoTo Handle
    Call SavePasswordData(rsTemp)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdSave_Click"
End Sub

Private Sub dtgUser_Click()
On Error GoTo Handle
    txtID.Text = Left(Trim(dtgUser.Columns(0).Text), 2)
    With rsTemp
    .Find "ID='" & txtID.Text & "'", , adSearchForward, adBookmarkFirst
    If Not .EOF Then
        txtuserCode.Text = txtID.Text & .Fields("Password")
        txtRetypeCode.Text = txtuserCode.Text
        txtUserName.Text = .Fields("UserName")
        cboLevel.ListIndex = Val(.Fields("UserLevel")) - 1
    End If
End With

    tvwRightAccess.Enabled = True
    InitRightOnTreeView
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " dtgUser_Click"
End Sub

'            ------------ FORM -----------
Private Sub Form_Load()
    fLoad = False
    ReDim Preserve DescArrTree(0)
    Dim i As Integer
    For i = 1 To 5 Step 1
        cboLevel.AddItem i, i - 1
    Next
    
     Set rsRight = Open_Table(cnData, "User_Login")
    With rsTemp
        If .State = 0 Then
            .Fields.Append "ID", adVarWChar, 10
            .Fields.Append "UserName", adVarWChar, 50
            .Fields.Append "UserLevel", adVarWChar, 1
            .Fields.Append "Password", adVarWChar, 250
            .Fields.Append "UserRight", adVarWChar, 250
            .Open
        End If
        With rsRight
        If .RecordCount > 0 Then .MoveFirst
            Do While Not .EOF
                rsTemp.addNew
                rsTemp.Fields("ID") = .Fields("ID")
                rsTemp.Fields("UserName") = .Fields("UserName")
                rsTemp.Fields("UserLevel") = .Fields("UserLevel")
                rsTemp.Fields("Password") = En_Decryption.MalgoDecrypt(Trim(.Fields("Password")), 10)
                rsTemp.Fields("UserRight") = .Fields("UserRight")
                rsTemp.Update
            .MoveNext
            Loop
        End With
    End With
    init_dtguser
    InitTvwRight
    Lock_text (True)
    fLoad = True
End Sub

Private Sub Form_Activate()
    Dim DescArr() As String
    DescArr = LoadLanguage(LngFile, "#04:002:")
    If UserID = "131112" Then Check1.Visible = True
End Sub
'            ------------ COMBOBOX -----------
Private Sub cboUser_Click()
    If fLoad Then InitRightOnTreeView
End Sub

Private Sub Form_Unload(Cancel As Integer)
CloseRecordset rsTemp
End Sub

'            ------------ TREEVIEW -----------
Private Sub tvwRightAccess_NodeCheck(ByVal Node As MSComctlLib.Node)
    SelectChild Node
    AddRightForUser
    iupdate = True
    cmdUpdate_Click
End Sub
'            ------------ COMMANDBUTTON -----------
Private Sub cmdCancel_Click()
If iupdate Then
    Call cmdSave_Click
End If
    CloseRecordset rsRight
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo Handle
If iupdate Then
    With rsTemp
        .Find "ID='" & txtID.Text & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
'            If MsgBox("M· ng­êi dïng ®· cã trong hÖ thèng, B¹n cã muèn thay ®æi th«ng tin kh«ng?", vbYesNo) = vbYes Then
                .Fields("ID") = txtID.Text
                .Fields("UserName") = txtUserName.Text
                .Fields("UserLevel") = cboLevel.Text
                .Fields("Password") = Mid(txtuserCode.Text, 3, Len(txtuserCode.Text) - 2)
                .Fields("UserRight") = ValueRightCode
                .Update
'            End If
        Else
            .addNew
            .Fields("ID") = txtID.Text
            .Fields("UserName") = txtUserName.Text
            .Fields("UserLevel") = cboLevel.Text
            .Fields("Password") = Mid(txtuserCode.Text, 3, Len(txtuserCode.Text) - 2)
            .Fields("UserRight") = ValueRightCode
            .Update
        End If
    End With
    cmdSave.Enabled = True
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " o k"
End Sub
'            ------------ OTHER FUNCTIONS -----------
Private Sub SelectChild(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
                
    SelectParent Node
    If Node.Children = 0 Then Exit Sub
    i = Node.Child.FirstSibling.Index
    tvwRightAccess.Nodes(i).Checked = Node.Checked
    SelectChild tvwRightAccess.Nodes(i)
    While i <> Node.Child.LastSibling.Index
        DoEvents
        i = tvwRightAccess.Nodes(i).Next.Index
        tvwRightAccess.Nodes(i).Checked = Node.Checked
        SelectChild tvwRightAccess.Nodes(i)
    Wend
End Sub

Private Sub SelectParent(ByVal Node As MSComctlLib.Node)
    Dim tempParentNode As MSComctlLib.Node
    Dim i As Integer
    Dim fFound As Boolean
    
    If Node.Parent Is Nothing Then Exit Sub
    If Node.Checked = True Then
        Node.Parent.Checked = True
        SelectParent Node.Parent
    Else
        fFound = False
        Set tempParentNode = Node.Parent
        i = tempParentNode.Child.FirstSibling.Index
        If tvwRightAccess.Nodes(i).Checked = True Then
            fFound = True
        Else
            Do While i <> tempParentNode.Child.LastSibling.Index
                DoEvents
                i = tvwRightAccess.Nodes(i).Next.Index
                If tvwRightAccess.Nodes(i).Checked = True Then
                    fFound = True
                    Exit Do
                End If
            Loop
        End If
        If Not fFound Then
            tempParentNode.Checked = False
            SelectParent tempParentNode
        End If
    End If
End Sub

Private Sub InitTvwRight()
    Dim iLen  As Byte
    Dim i As Integer
    Call Read_Desc(LngFile, "#04:003:")
    'If cmdOK.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    With tvwRightAccess
        .Nodes.Add , , "MF", "Cua so chinh"
        .Nodes("MF").Expanded = True
        For i = 1 To UBound(DescArrTree)
        DoEvents
            iLen = Len(DescArrTree(i).ID)
            Select Case iLen
                Case 3: .Nodes.Add "MF", tvwChild, DescArrTree(i).ID, DescArrTree(i).Desc
                Case 5
                        .Nodes.Add Left(DescArrTree(i).ID, 3), tvwChild, DescArrTree(i).ID, DescArrTree(i).Desc
                        .Nodes(Left(DescArrTree(i).ID, 3)).Expanded = True
                Case 7: .Nodes.Add Left(DescArrTree(i).ID, 5), tvwChild, DescArrTree(i).ID, DescArrTree(i).Desc
                Case 9: .Nodes.Add Left(DescArrTree(i).ID, 7), tvwChild, DescArrTree(i).ID, DescArrTree(i).Desc
            End Select
        Next i
    End With
    InitRightOnTreeView
End Sub

Private Sub Read_Desc(ByVal LngFile As String, ByVal DescPos As String)
    Dim hFile As Integer
    Dim tmpStr As String
    Dim lFound As Boolean
    lFound = False
    hFile = FreeFile
    Open LngFile For Input As #hFile
    Do While Not EOF(hFile)
        DoEvents
        Line Input #hFile, tmpStr
        If Left(tmpStr, Len(DescPos)) = DescPos Then
            lFound = True
            tmpStr = Right(tmpStr, Len(tmpStr) - Len(DescPos))
            ReDim Preserve DescArrTree(UBound(DescArrTree) + 1)
            DescArrTree(UBound(DescArrTree)).ID = "B" & Left(tmpStr, InStr(tmpStr, ":") - 1)
            DescArrTree(UBound(DescArrTree)).Desc = Right(tmpStr, Len(tmpStr) - InStr(tmpStr, ":"))
        Else
            If lFound And Left(tmpStr, 4) = "#000" Then Exit Do
        End If
    Loop
    Close #hFile
End Sub



Private Function ValueRightCode() As String
    Dim str_Right As String
    Dim sTempRight As String
    Dim tempKey As String
    Dim sTemp As String
    Dim iCount As Integer
    Dim i As Integer
    
    iCount = 0
    str_Right = "": sTemp = "": tempKey = ""
    With tvwRightAccess
        For i = 1 To .Nodes.count
        DoEvents
            Select Case Left(.Nodes(i).KEY, 3)
                Case "B01", "B02", "B03", "B04", "B05", "B06", "B07", "B08", "B09", "B10"
                        If tempKey <> Left(.Nodes(i).KEY, 3) And tempKey <> "" Then
                            If iCount < 8 Then
                                sTemp = Left(sTemp & "00000000", 8)
                                sTemp = FillZeroForString(BinToHex(sTemp), 2)
                                sTemp = addBlankValue(sTempRight & sTemp, 64)
                                str_Right = str_Right & sTemp
                                sTempRight = ""
                                sTemp = ""
                                iCount = 0
                            End If
                        End If
                        iCount = iCount + 1
                        If .Nodes(i).Checked Then
                            sTemp = sTemp & "1"
                        Else
                            sTemp = sTemp & "0"
                        End If
                        If iCount = 8 Then
                            sTempRight = sTempRight & FillZeroForString(BinToHex(sTemp), 2)
                            iCount = 0
                            sTemp = ""
                        End If
                        tempKey = Left(.Nodes(i).KEY, 3)
            End Select
        Next i
    End With
    If iCount < 8 Then
        If iCount <> 0 Then
            sTemp = Left(sTemp & "00000000", 8)
            sTemp = FillZeroForString(BinToHex(sTemp), 2)
            str_Right = str_Right & addBlankValue(sTempRight & sTemp, 64)
        Else
            str_Right = str_Right & addBlankValue(sTempRight, 64)
        End If
    End If
    ValueRightCode = str_Right
End Function

Private Function addBlankValue(S1 As String, iLen As Integer) As String
    Do While Len(S1) < 64
    DoEvents
        S1 = S1 & "-1"
    Loop
    addBlankValue = S1
End Function

Private Sub InitRightOnTreeView()
    Dim tempIndex As Integer
    Dim sFullRight As String
    Dim tmpRight As sRight
    Dim i As Integer
        
    tmpRight = MyRight
    With tmpRight
        If rsRight.RecordCount = 0 Then Exit Sub
            rsRight.MoveFirst
        Do While Not rsRight.EOF
        DoEvents
            If StrComp(Trim(rsRight.Fields("ID")), txtID.Text, 1) = 0 Then
                .FullRight = rsRight.Fields("UserRight")
                .Sodoban = RightDeCode(Left(.FullRight, 64))
                .Banhang = RightDeCode(Mid(.FullRight, 65, 64))
                .Danhmuc = RightDeCode(Mid(.FullRight, 129, 64))
                .Nhanvien = RightDeCode(Mid(.FullRight, 193, 64))
                .Caidathethong = RightDeCode(Mid(.FullRight, 257, 64))
                .Caidatdanhmuc = RightDeCode(Mid(.FullRight, 321, 64))
                .Baocao = RightDeCode(Mid(.FullRight, 385, 64))
                .kho = RightDeCode(Mid(.FullRight, 449, 64))
                .Thuchi = RightDeCode(Mid(.FullRight, 513, 64))
                .Suaten = RightDeCode(Mid(.FullRight, 577, 64))
                Exit Do
            End If
            rsRight.MoveNext
        Loop
    End With
    
    tempIndex = 0
    With tvwRightAccess
        For i = 2 To .Nodes.count
        DoEvents
            If Len(.Nodes(i).KEY) = 3 Then
                tempIndex = 0
            End If
            Select Case Left(.Nodes(i).KEY, 3)
                Case "B01": sFullRight = Trim(tmpRight.Sodoban)
                Case "B02": sFullRight = Trim(tmpRight.Banhang)
                Case "B03": sFullRight = Trim(tmpRight.Danhmuc)
                Case "B04": sFullRight = Trim(tmpRight.Nhanvien)
                Case "B05": sFullRight = Trim(tmpRight.Caidathethong)
                Case "B06": sFullRight = Trim(tmpRight.Caidatdanhmuc)
                Case "B07": sFullRight = Trim(tmpRight.Baocao)
                Case "B08": sFullRight = Trim(tmpRight.kho)
                Case "B09": sFullRight = Trim(tmpRight.Thuchi)
                Case "B10": sFullRight = Trim(tmpRight.Suaten)
                
            End Select
            tempIndex = tempIndex + 1
            If Mid(sFullRight, tempIndex, 1) = 1 Then
                  .Nodes(i).Checked = True
            Else
                .Nodes(i).Checked = False
            End If
            .Nodes(i).Expanded = False
        Next i
        
    End With
End Sub

Private Function RightDeCode(S1 As String) As String
    Dim sResult As String
    Dim i As Integer
    
    sResult = ""
    If S1 = "" Then GoTo 1
    For i = 1 To Len(S1) Step 2
    DoEvents
        If Mid(S1, i, 2) <> "-1" And Mid(S1, 1, 2) <> "  " Then
            sResult = sResult & FillZeroForString(HexToBin(Mid(S1, i, 2)), 8)
        End If
    Next i
1:  RightDeCode = sResult
End Function

Private Sub AddRightForUser()
    With rsTemp
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
        DoEvents
            If .Fields("ID") = txtID.Text Then
                .Fields("UserRight") = ValueRightCode
                Exit Do
            End If
            .MoveNext
        Loop
    End With
End Sub

Public Sub Init_AddNew()
On Error GoTo Handle
Dim i As Integer
    For i = 1 To 5 Step 1
        cboLevel.AddItem i, i - 1
    Next
    cboLevel.ListIndex = 0
    txtUserName.Text = ""
    txtuserCode.Text = ""
    txtuserCode.PasswordChar = "*"
    txtRetypeCode.Text = ""
    txtRetypeCode.PasswordChar = "*"
    tvwRightAccess.Enabled = True
    lblmatch.Visible = True
    txtuserCode.SetFocus
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Init_AddNew"
End Sub

Private Sub txtID_Change()
On Error GoTo Handle

InitRightOnTreeView
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - txtID_Change"

End Sub

Private Sub txtRetypeCode_Change()
If StrComp(txtuserCode.Text, txtRetypeCode.Text) <> 0 Then
   lblmatch.Caption = "X¸c nhËn mËt khÈu kh«ng ®óng!"
   lblmatch.ForeColor = vbRed
Else
    lblmatch.Caption = "OK!"
   lblmatch.ForeColor = vbBlue
End If
End Sub

Private Sub txtRetypeCode_LostFocus()
If StrComp(txtuserCode.Text, txtRetypeCode.Text) <> 0 Then
    MsgBox "X¸c nhËn mËt khÈu kh«ng ®óng!"
    txtRetypeCode.SelStart = 0
    txtRetypeCode.SelLength = 9999
    txtRetypeCode.SetFocus
Else
    lblmatch.Caption = "OK!"
   lblmatch.ForeColor = vbBlue
End If
End Sub

Private Sub txtUserCode_Change()
On Error GoTo Handle
    txtID.Text = Left(txtuserCode.Text, 2)
    'InitRightOnTreeView
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - txtUserCode_Change"
End Sub

Public Sub init_dtguser()
On Error GoTo Handle
    Set dtgUser.DataSource = rsTemp
    With dtgUser
        .Columns(0).Caption = "M· ng­êi dïng"
        .Columns(0).Width = 1500
        .Columns(1).Caption = "Tªn ng­êi dïng"
        .Columns(1).Width = 3500
        .Columns(2).Caption = "CÊp ®é User"
        .Columns(2).Width = 2000
        .Columns(3).Caption = "MËt khÈu"
        .Columns(3).Width = 0
        .Columns(4).Width = 0
        
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub txtUserCode_LostFocus()
If Len(txtuserCode.Text) <= 2 Then
    MsgBox "M· §¨ng nhËp chøa Ýt nhÊt 3 ký tù"
    txtuserCode.SelStart = 0
    txtuserCode.SelLength = 9999
    txtuserCode.SetFocus
End If
End Sub


Public Sub Lock_text(F As Boolean)
txtuserCode.Locked = F
txtRetypeCode.Locked = F
txtUserName.Locked = F
cboLevel.Enabled = Not F
cmdUpdate.Enabled = Not F
'cmdSave.Enabled = Not f
End Sub
