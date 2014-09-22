VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStockType 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Danh môc kho"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10515
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
   ScaleHeight     =   6345
   ScaleWidth      =   10515
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6375
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   10455
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
         Left            =   5520
         ScaleHeight     =   615
         ScaleWidth      =   4755
         TabIndex        =   5
         Top             =   240
         Width           =   4815
         Begin VB.Label lblName 
            BackColor       =   &H80000008&
            Caption         =   "Group Name"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   330
            Width           =   4455
         End
         Begin VB.Label lblNo 
            BackColor       =   &H80000008&
            Caption         =   "Group No"
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   45
            Width           =   4455
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
         Height          =   2085
         Left            =   5640
         TabIndex        =   1
         Top             =   4080
         Width           =   4620
         Begin prjTouchScreen.MyButton cmdClose 
            Height          =   855
            Left            =   1620
            TabIndex        =   2
            Tag             =   "L5"
            Top             =   1200
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1508
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
            MICON           =   "frmStockType.frx":0000
            PICN            =   "frmStockType.frx":001C
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
            Height          =   855
            Left            =   120
            TabIndex        =   3
            Tag             =   "L4"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1508
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
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmStockType.frx":62B6
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
            Height          =   855
            Left            =   1620
            TabIndex        =   4
            Tag             =   "L3"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1508
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
            MICON           =   "frmStockType.frx":62D2
            PICN            =   "frmStockType.frx":62EE
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
            Height          =   855
            Left            =   3120
            TabIndex        =   14
            Tag             =   "L4"
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1508
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
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   16711680
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmStockType.frx":6832
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
      Begin TabDlg.SSTab tab1 
         Height          =   2655
         Left            =   5640
         TabIndex        =   8
         Top             =   1320
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
         TabPicture(0)   =   "frmStockType.frx":684E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblGroupName"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblCode"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtStockName"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtStockID"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         Begin VB.TextBox txtStockID 
            Height          =   495
            Left            =   360
            TabIndex        =   12
            Top             =   1080
            Width           =   3615
         End
         Begin VB.TextBox txtStockName 
            BeginProperty Font 
               Name            =   ".VnArialH"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   9
            Top             =   2040
            Width           =   3975
         End
         Begin VB.Label lblCode 
            Caption         =   "M· kho:"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Tag             =   "L2"
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label lblGroupName 
            Caption         =   "&Data:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Tag             =   "L2"
            Top             =   1680
            Width           =   1095
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flgStockList 
         Height          =   6015
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   10610
         _Version        =   393216
         BackColorBkg    =   16777215
         TextStyleFixed  =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmStockType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsStockList As New ADODB.Recordset
Dim strPath As String
Dim DescArr() As String


Private Sub cmdSave_Click()
    Call UpdateDatabase
    Call LoadControl
    If cmdAdd.Enabled = True Then
        cmdAdd.SetFocus
    Else
        cmdAdd.Enabled = True
        cmdAdd.SetFocus
    End If
End Sub

Private Sub cmdClose_Click()
    Set rsStockList = Nothing
    Unload Me
End Sub

Private Sub cmdAdd_Click()
On Error GoTo Handle
    If cmdAdd.Caption = "Thªm míi" Then
        Call UnlockText
        Call DeleteTextbox
        txtStockID.SetFocus
    ElseIf cmdAdd.Caption = "&Söa" Then
        Call UnlockText
    End If
Exit Sub
Handle:
MsgBox Err.Number & " " & Err.Description & " " & _
            Me.name & " " & "cmdAdd _click"

End Sub
Private Sub DeleteTextbox()
    On Error GoTo Handle
        cmdAdd.Caption = "Söa"
        txtStockID.Text = ""
       txtStockName.Text = ""
        txtStockID.SetFocus
        
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "DeleteTextbox"
End Sub
Private Sub UpdateDatabase()
    On Error GoTo Handle
        With rsStockList
            .Find "StockID='" & txtStockID.Text & _
                "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("StockID") = txtStockID.Text
                .Fields("StockName") = txtStockName.Text
                .Update
                .Requery
            Else
                MsgBox "StockID ®· tån t¹i, vui lßng kiÓm tra l¹i hoÆc ®æi m· kh¸c!", vbOKOnly
                Call DeleteTextbox
                
            End If
        End With
        Call setflgStockList
        cmdAdd.Caption = "Thªm míi" 'DescArr(4)
        
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "UpdateDatabase"
End Sub


Private Sub cmddelete_Click()

    On Error GoTo Handle
    Dim ans As Integer
    ans = MsgBox("B¹n cã ch¾c ch¨n muèn xãa danh môc nµy kh«ng?", vbYesNo)
    If ans = vbYes Then
        With rsStockList
            .Find "StockID='" & flgStockList.TextMatrix(flgStockList.Row, 0) & _
                    "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Or .BOF Then
                .Delete adAffectCurrent
                .MoveNext
                .Requery
            End If
            Call Form_Load
        End With
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "cmdDelete_Click"

End Sub

Private Sub flgStockList_EnterCell()
    On Error GoTo Handle
    With rsStockList
        .Find "StockID='" & flgStockList.TextMatrix(flgStockList.Row, 0) & _
                "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtStockID.Text = !stockID
            txtStockName.Text = !StockName
            lblNo.Caption = !stockID
            lblName.Caption = !StockName
        End If
    End With
    Exit Sub
    
Handle:
    MsgBox Err.Number & " " & Err.Description & vbCrLf _
    & Me.name & " flgStockList_EnterCell"
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    
    Dim str As String
    str = "Select * from Stock"
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(strPath, "100881administrator")
'    End If
    Set rsStockList = OpenCriticalTable(str, cnData)
    Call setflgStockList
    Call LockText
    Exit Sub
Handle:
    MsgBox Err.Number & " " & _
    Err.Description & Me.name & "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Handle
    Set rsStockList = Nothing
Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & "" & _
                        Me.name & " " & "form_Unload"
End Sub
Private Sub setflgStockList()
On Error GoTo errHdl
    Dim intCount    As Integer
    With flgStockList
        .Font = ".vnArial"
        .ColWidth(0) = 2500
        .ColWidth(1) = 7500
        .TextMatrix(0, 0) = "M· kho"
        .TextMatrix(0, 1) = "Tªn kho"
    End With
    
    If rsStockList Is Nothing Then Exit Sub
    If rsStockList.State = 0 Then Exit Sub
    
    If rsStockList.EOF And rsStockList.BOF Then
        With flgStockList
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
        End With
        Exit Sub
    End If
   flgStockList.Rows = rsStockList.RecordCount + 1
    intCount = 0
    Do While Not rsStockList.EOF
        intCount = intCount + 1
        flgStockList.TextMatrix(intCount, 0) = rsStockList!stockID
        flgStockList.TextMatrix(intCount, 1) = rsStockList!StockName
        rsStockList.MoveNext
        
    Loop
'    SetColorFlexGrid flgStockList, 1, 1, flgStockList.Cols

    Call flgStockList_EnterCell
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - setflgStockList "
End Sub
Private Sub LoadControl()
    On Error GoTo Handle
    
    With rsStockList
        .Find "StockID='" & !stockID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtStockID.Text = !stockID
           txtStockName.Text = !StockName
            .Requery
        End If
    End With
    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " LoadControl"
End Sub

Public Sub UnlockText()
    On Error GoTo Handle
        txtStockID.Locked = False
        txtStockName.Locked = False
        cmdSave.Enabled = True
        txtStockID.SetFocus
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub
Public Sub LockText()
    On Error GoTo Handle
        txtStockID.Locked = True
        txtStockName.Locked = True
        cmdSave.Enabled = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub

Private Sub txtStockName_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtStockName.Text = .Let_Text_Input
        End With

    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtStockID_DblClick "

End Sub

Private Sub txtStockName_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        cmdSave.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtStockName_KeyPress"

End Sub

Private Sub txtStockID_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtStockID.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtStockID_DblClick "

End Sub

Private Sub txtStockID_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        txtStockName.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtStockID_KeyPress"
End Sub




