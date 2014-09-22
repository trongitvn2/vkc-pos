VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmCustom_type 
   Caption         =   "Danh môc nhãm kh¸ch hµng"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12075
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
   Icon            =   "frmCustom_type.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   12075
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNote 
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   3360
      Width           =   3615
   End
   Begin VB.TextBox txtPro_Value 
      Height          =   390
      Left            =   2640
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   6360
      TabIndex        =   13
      Top             =   0
      Width           =   5655
      Begin MSFlexGridLib.MSFlexGrid flgCust_Type 
         Height          =   4455
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   7858
         _Version        =   393216
      End
   End
   Begin prjTouchScreen.MyButton cmdthoat 
      Cancel          =   -1  'True
      Height          =   855
      Left            =   4800
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      BTYPE           =   2
      TX              =   "§ãn&g"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCustom_type.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdXoa 
      Height          =   855
      Left            =   3240
      TabIndex        =   7
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      BTYPE           =   2
      TX              =   "&Xãa"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCustom_type.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdLuu 
      Height          =   855
      Left            =   1680
      TabIndex        =   6
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      BTYPE           =   2
      TX              =   "&CËp nhËt"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCustom_type.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdThem 
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   3960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      BTYPE           =   2
      TX              =   "&Thªm míi"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCustom_type.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.ComboBox cboProm_Type 
      Height          =   390
      Left            =   2640
      TabIndex        =   2
      Text            =   "Chän h×nh thøc khuyÕn m·i"
      Top             =   2160
      Width           =   3615
   End
   Begin VB.TextBox txtTennhom 
      Height          =   390
      Left            =   2640
      TabIndex        =   1
      Top             =   1560
      Width           =   3615
   End
   Begin VB.TextBox txtmanhom 
      Height          =   390
      Left            =   2640
      MaxLength       =   20
      TabIndex        =   0
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Ghi chó"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   16
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Gi¸ trÞ khuyÕn m·i"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label lblHinhthuc 
      Alignment       =   1  'Right Justify
      Caption         =   "H×nh thøc khuyÕn m·i"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label lblTennhom 
      Alignment       =   1  'Right Justify
      Caption         =   "Tªn nhãm:"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label lblManhom 
      Alignment       =   1  'Right Justify
      Caption         =   "M· nhãm:"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Danh môc nhãm kh¸ch hµng th©n thiÕt"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmCustom_type"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCustomer_Type As New ADODB.Recordset
Dim strSql As String

Private Sub cmdLuu_Click()
On Error GoTo Handle
Dim ans As Integer
    If allow_Save Then
        With rsCustomer_Type
            .Find "CustType_ID='" & Trim(txtmanhom.Text) & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("CustType_ID") = txtmanhom.Text
                .Fields("CustType_Name") = txtTennhom.Text
                .Fields("Promotion") = cboProm_Type.ListIndex
                .Fields("Pro_Value") = CDbl("0" & txtPro_Value.Text)
                .Fields("Note") = txtNote.Text
                .Update
            Else
                ans = MsgBox("M· nhãm nµy ®· tån t¹i råi, b¹n cã muèn l­u th«ng tin thay ®æi kh«ng?", vbYesNo)
                If ans = vbYes Then
                    .Fields("CustType_ID") = txtmanhom.Text
                    .Fields("CustType_Name") = txtTennhom.Text
                    .Fields("Promotion") = cboProm_Type.ListIndex
                    .Fields("Pro_Value") = CDbl("0" & txtPro_Value.Text)
                    .Fields("Note") = txtNote.Text
                    .Update
                Else
                    Call Init_Add_new
                    cmdThem.Enabled = True
                    cmdLuu.Enabled = False
                End If
            End If
        End With
        Call Set_Data_In_Flex
    End If
    cmdThem.Enabled = True
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & ""
End Sub

Private Sub cmdThem_Click()
    Call Init_Add_new
End Sub

Private Sub cmdthoat_Click()
    Unload Me
End Sub

Public Sub Set_flgCust_Type()
    On Error GoTo Handle
        With flgCust_Type
            .Cols = 4
            .Rows = 2
            .ColWidth(0) = 1200
            .ColWidth(1) = 2200
            .ColWidth(2) = 2500
            .ColWidth(3) = 1000
            .TextMatrix(0, 0) = "M· nhãm"
            .TextMatrix(0, 1) = "Tªn nhãm"
            .TextMatrix(0, 2) = "ChÝnh s¸ch khuyÕn m·i"
            .TextMatrix(0, 3) = "Gi¸ trÞ khuyÕn m·i"
            .ColAlignment(0) = 2
            .ColAlignment(1) = 2
            .ColAlignment(2) = 2
            .ColAlignment(3) = 6
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Set_flgCust_Type"
End Sub

Private Sub Set_Data_In_Flex()
On Error GoTo Handle
Dim incount As Integer
Dim rs As New ADODB.Recordset
Set rs = Open_Table(cnData, "Customer_Type")
    With rs
    If .State = 0 Then Exit Sub
    If .RecordCount = 0 Then Call Set_flgCust_Type
        Do While Not .EOF
             incount = incount + 1
                flgCust_Type.Rows = rs.RecordCount + 1
                With flgCust_Type
                    .TextMatrix(incount, 0) = rs!CustType_ID
                    .TextMatrix(incount, 1) = rs!CustType_Name
                    If rs.Fields("Promotion") = 0 Then
                        .TextMatrix(incount, 2) = ""
                    ElseIf rs.Fields("Promotion") = 1 Then
                        .TextMatrix(incount, 2) = "Gi¶m tæng Hãa ®¬n " & rs!Pro_Value & "%"
                    ElseIf rs.Fields("Promotion") = 2 Then
                        .TextMatrix(incount, 2) = "Gi¶m thøc ¨n " & rs!Pro_Value & "%"
                    ElseIf rs.Fields("Promotion") = 3 Then
                        .TextMatrix(incount, 2) = "Gi¶m thøc uèng " & rs!Pro_Value & "%"
                    End If
                    .TextMatrix(incount, 3) = rs!Pro_Value
                End With
        .MoveNext
        Loop
    End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Set_Data_In_Flex"
End Sub

Private Sub cmdXoa_Click()
On Error GoTo Handle
    With rsCustomer_Type
        .Find "CustType_ID='" & flgCust_Type.TextMatrix(flgCust_Type.Row, 0) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Delete adAffectCurrent
        End If
        Call Set_Data_In_Flex
    End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  cmdXoa_Click"
End Sub

Private Sub flgCust_Type_Click()
On Error GoTo Handle
    With rsCustomer_Type
        .Find "CustType_ID='" & flgCust_Type.TextMatrix(flgCust_Type.Row, 0) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtmanhom.Text = .Fields("CustType_ID")
            txtTennhom.Text = .Fields("CustType_Name")
            txtPro_Value.Text = .Fields("Pro_Value")
            cboProm_Type.ListIndex = .Fields("Promotion")
            txtNote.Text = .Fields("Note") & ""
        End If
    End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  flgCust_Type_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Set rsCustomer_Type = Open_Table(cnData, "Customer_Type")
    Call Set_Promotion_type
    Call Set_flgCust_Type
    Call Set_Data_In_Flex
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Form_Load"
End Sub

Public Sub Set_Promotion_type()
On Error GoTo Handle
   With cboProm_Type
        .Clear
        .AddItem "Kh«ng cã h×nh thøc khuyÕn m·i", 0
        .AddItem "Gi¶m tæng Hãa ®¬n", 1
        .AddItem "Gi¶m thøc ¨n", 2
        .AddItem "Gi¶m thøc uèng", 3
   End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Set_Promotion_type"
End Sub

Public Sub Init_Add_new()
On Error GoTo Handle
    cmdThem.Enabled = False
    txtmanhom.Text = ""
    txtTennhom.Text = ""
    cboProm_Type.ListIndex = 0
    txtPro_Value.Text = 0
    txtmanhom.SetFocus
    cmdLuu.Enabled = True
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Init_Add_new"
End Sub

Public Function allow_Save() As Boolean
On Error GoTo Handle
Dim isallow As Boolean
    If txtmanhom.Text = "" Then
        MsgBox "M· nhãm kh«ng ®­îc rçng"
        isallow = False
    ElseIf txtTennhom.Text = "" Then
        MsgBox "Tªn nhãm kh«ng ®­îc rçng"
        isallow = False
    Else
        isallow = True
    End If
    allow_Save = isallow
    Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  allow_Save"
End Function
