VERSION 5.00
Begin VB.Form frmUpdatePrice 
   Caption         =   "CËp nhËt gi¸"
   ClientHeight    =   5385
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   13155
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpdatePrice.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   13155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Gi¸ Everning"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   8760
      TabIndex        =   14
      Top             =   2280
      Width           =   4095
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   8
         Left            =   1320
         TabIndex        =   23
         Text            =   "0"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   7
         Left            =   1320
         TabIndex        =   22
         Text            =   "0"
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   6
         Left            =   1320
         TabIndex        =   21
         Text            =   "0"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Gi¸ 3:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   600
         TabIndex        =   32
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Gi¸ 2:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   600
         TabIndex        =   31
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Gi¸ 1:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   600
         TabIndex        =   30
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gi¸ chuÈn"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4560
      TabIndex        =   12
      Top             =   2280
      Width           =   4095
      Begin VB.Frame Frame2 
         Caption         =   "Gi¸ Happy Hour"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   4095
         Begin VB.TextBox txtPrice 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   5
            Left            =   1320
            TabIndex        =   20
            Text            =   "0"
            Top             =   1560
            Width           =   2655
         End
         Begin VB.TextBox txtPrice 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   4
            Left            =   1320
            TabIndex        =   19
            Text            =   "0"
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txtPrice 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   1320
            TabIndex        =   18
            Text            =   "0"
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "Gi¸ 3:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   600
            TabIndex        =   29
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Gi¸ 2:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   600
            TabIndex        =   28
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "Gi¸ 1:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   600
            TabIndex        =   27
            Top             =   480
            Width           =   735
         End
      End
   End
   Begin VB.Frame fraPrice 
      Caption         =   "Gi¸ chuÈn"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   360
      TabIndex        =   11
      Top             =   2280
      Width           =   4095
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   1320
         TabIndex        =   17
         Text            =   "0"
         Top             =   1560
         Width           =   2655
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   1320
         TabIndex        =   16
         Text            =   "0"
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   1320
         TabIndex        =   15
         Text            =   "0"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Gi¸ 3:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   26
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Gi¸ 2:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   25
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Gi¸ 1:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   24
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.OptionButton optGroup 
      Caption         =   "Theo nhãm hµng"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   480
      Width           =   3135
   End
   Begin VB.Frame fraGroup 
      Height          =   1455
      Left            =   7680
      TabIndex        =   8
      Top             =   720
      Width           =   5175
      Begin VB.ComboBox cboGroup 
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
         Left            =   120
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   480
         Width           =   4935
      End
   End
   Begin VB.OptionButton optCode 
      Caption         =   "Theo d·y m· hµng"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Value           =   -1  'True
      Width           =   3255
   End
   Begin VB.Frame fraCode 
      Height          =   1455
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   7095
      Begin VB.ComboBox cboTo 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1320
         TabIndex        =   7
         Text            =   "Combo3"
         Top             =   840
         Width           =   5775
      End
      Begin VB.ComboBox cboFrom 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1320
         TabIndex        =   6
         Text            =   "Combo2"
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label2 
         Caption         =   "§Õn:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "M· hµng"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Height          =   735
      Left            =   7560
      TabIndex        =   1
      Top             =   4560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      BTYPE           =   2
      TX              =   "§ãng"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdatePrice.frx":000C
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
      Height          =   735
      Left            =   4440
      TabIndex        =   0
      Top             =   4560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
      BTYPE           =   2
      TX              =   "CËp nhËt"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdatePrice.frx":0028
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
Attribute VB_Name = "frmUpdatePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsInventory As New ADODB.Recordset
Dim rsDepartment As New ADODB.Recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo Handle
Dim StartCode, EndCode, DeptID As String
  If Len(cboFrom.Text) > 10 Then
    StartCode = Trim(Left(cboFrom.Text, InStr(cboFrom.Text, Space(10)) - 1))
Else
    StartCode = Trim(Val(cboFrom.Text))
End If

If Len(cboTo.Text) > 10 Then
    EndCode = Trim(Left(cboTo.Text, InStr(cboTo.Text, Space(10)) - 1))
Else
     EndCode = Trim(Val(cboTo.Text))
End If
    DeptID = Trim(Left(cboGroup.Text, InStr(cboGroup.Text, Space(10)) - 1))
    
    If optCode.Value = True Then
        Call Update_Price_ByCode(StartCode, EndCode)
    ElseIf optGroup.Value = True Then
        Call Update_Price_ByGroup(DeptID)
    End If
    MsgBox "CËp nhËt thµnh c«ng!"
    'frmItems.Show vbModal
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  - cmdUpdate_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    If cnData.State = 0 Then Exit Sub
    'Khoi tao records
    Set rsInventory = Open_Table(cnData, "Inventory")
    Set rsDepartment = Open_Table(cnData, "Departments")
    
    'Gan du lieu vao cac combo
    Call InitCombo("Departments", cboGroup, "Dept_ID", "Description", False)
    Call InitCombo("Inventory", cboFrom, "ItemNum", "ItemName", False)
    Call InitCombo("Inventory", cboTo, "ItemNum", "ItemName", False)
    'Disable combo Group
    fraGroup.Enabled = False
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  - Form_Load"
End Sub

Private Sub optCode_Click()
    If optCode.Value = True Then
        fraCode.Enabled = True
        fraGroup.Enabled = False
    Else
        fraCode.Enabled = False
        fraGroup.Enabled = True
    End If
End Sub

Private Sub optGroup_Click()
     If optGroup.Value = True Then
        fraCode.Enabled = False
        fraGroup.Enabled = True
    Else
        fraCode.Enabled = True
        fraGroup.Enabled = False
    End If
End Sub

Public Sub InitCombo(ByVal sTableName As String, ByVal cbo As ComboBox, ByVal sFieldName As String, ByVal sFieldName1 As String, ByVal fEmpty As Boolean)
On Error GoTo errHdl

    Dim res As New ADODB.Recordset
    
    If sTableName = "" Then Exit Sub
    cbo.Clear
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set res = Open_Table(cnData, sTableName)
    If res.State = 0 Then Exit Sub
    If fEmpty = True Then cbo.AddItem "-------"
    With res
        If .RecordCount = 0 Then
            cbo.AddItem "-------"
            GoTo 1
        End If
        .MoveFirst
        Do While Not .EOF
        DoEvents
            cbo.AddItem res.Fields(sFieldName) & Space(10) & res.Fields(sFieldName1)
            .MoveNext
        Loop
    End With
1:
    CloseRecordset res
    cbo.ListIndex = 0
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "InitCombo - InitCombo"
End Sub

Private Sub txtPrice_Change(Index As Integer)
On Error GoTo Handle
    txtPrice(Index).Text = Format(txtPrice(Index).Text, "#,##0")
    txtPrice(Index).SelStart = Len(txtPrice(Index).Text)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtPrice_Change"
End Sub

Private Sub txtPrice_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        txtPrice(Index).Text = Format(txtPrice(Index).Text, "#,##0")
        If Index <= 7 Then
            txtPrice(Index + 1).SetFocus
        Else
            cmdUpdate.SetFocus
        End If
        
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtPrice_Change"
End Sub

Public Sub Update_Price_ByCode(ByVal Start_Code As String, ByVal End_Code As String)
On Error GoTo Handle
Dim i As Integer
    If rsInventory.State <> 0 Then
        If rsInventory.RecordCount > 0 Then rsInventory.MoveFirst
    Else
        Exit Sub
    End If
    For i = Val(Start_Code) To Val(End_Code)
        With rsInventory
            .Find "ItemNum='" & Right("000000000000" & i, 12) & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Fields("Std_Price1") = txtPrice(0).Text
                .Fields("Std_Price2") = txtPrice(1).Text
                .Fields("Std_Price3") = txtPrice(2).Text
                
                'HH-Price
                .Fields("HH_Price1") = txtPrice(3).Text
                .Fields("HH_Price2") = txtPrice(4).Text
                .Fields("HH_Price3") = txtPrice(5).Text
                'Everning Price
                .Fields("EV_Price1") = txtPrice(6).Text
                .Fields("EV_Price2") = txtPrice(7).Text
                .Fields("EV_Price3") = txtPrice(8).Text
                .Update
            End If
        End With
    Next
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  - Update_Price_ByCode "
End Sub

Public Sub Update_Price_ByGroup(ByVal Dept_ID As String)
On Error GoTo Handle
    Dim strSql As String
    strSql = "Select * from Inventory Where Dept_ID='" & Dept_ID & "'"
    Set rsInventory = OpenCriticalTable(strSql, cnData)
    With rsInventory
        Do While Not rsInventory.EOF
            .Fields("Std_Price1") = txtPrice(0).Text
                .Fields("Std_Price2") = txtPrice(1).Text
                .Fields("Std_Price3") = txtPrice(2).Text
                
                'HH-Price
                .Fields("HH_Price1") = txtPrice(3).Text
                .Fields("HH_Price2") = txtPrice(4).Text
                .Fields("HH_Price3") = txtPrice(5).Text
                'Everning Price
                .Fields("EV_Price1") = txtPrice(6).Text
                .Fields("EV_Price2") = txtPrice(7).Text
                .Fields("EV_Price3") = txtPrice(8).Text
                .Update
        .MoveNext
        Loop
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  - Update_Price_ByGroup "
End Sub
