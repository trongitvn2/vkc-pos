VERSION 5.00
Begin VB.Form frmAddNewCust 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Customer"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10140
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1"
   Begin VB.ComboBox cboPro 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   360
      Width           =   2535
   End
   Begin prjTouchScreen.MyButton cmdCreate 
      Height          =   885
      Left            =   2610
      TabIndex        =   10
      Top             =   4110
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1561
      BTYPE           =   14
      TX              =   "T¹o míi"
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
      BCOL            =   16578804
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAddNewCust.frx":0000
      PICN            =   "frmAddNewCust.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.TextBox txtNewCust 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1890
      MaxLength       =   50
      TabIndex        =   3
      Tag             =   "1"
      Top             =   1530
      Width           =   8235
   End
   Begin VB.TextBox txtNewCust 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   4560
      MaxLength       =   15
      TabIndex        =   6
      Tag             =   "4"
      Top             =   2850
      Width           =   2325
   End
   Begin VB.TextBox txtNewCust 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   1920
      MaxLength       =   13
      TabIndex        =   8
      Tag             =   "6"
      Top             =   3480
      Width           =   2805
   End
   Begin VB.TextBox txtNewCust 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   6360
      TabIndex        =   9
      Tag             =   "8"
      Top             =   3480
      Width           =   3765
   End
   Begin VB.TextBox txtNewCust 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   8010
      MaxLength       =   12
      TabIndex        =   7
      Tag             =   "5"
      Top             =   2880
      Width           =   2085
   End
   Begin VB.TextBox txtNewCust 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1890
      MaxLength       =   15
      TabIndex        =   5
      Tag             =   "3"
      Top             =   2850
      Width           =   2085
   End
   Begin VB.TextBox txtNewCust 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1890
      MaxLength       =   255
      TabIndex        =   4
      Tag             =   "2"
      Top             =   2190
      Width           =   8235
   End
   Begin VB.TextBox txtNewCust 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1890
      MaxLength       =   100
      TabIndex        =   2
      Tag             =   "1"
      Top             =   900
      Width           =   8235
   End
   Begin VB.TextBox txtNewCust 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1890
      MaxLength       =   12
      TabIndex        =   0
      Tag             =   "0"
      Top             =   360
      Width           =   3030
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   885
      Left            =   5970
      TabIndex        =   11
      Top             =   4110
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1561
      BTYPE           =   14
      TX              =   "&Tho¸t"
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
      BCOL            =   16578804
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAddNewCust.frx":0656
      PICN            =   "frmAddNewCust.frx":0672
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
      Height          =   885
      Left            =   4290
      TabIndex        =   22
      Top             =   4110
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1561
      BTYPE           =   14
      TX              =   " Gióp ®ì"
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
      BCOL            =   16578804
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAddNewCust.frx":690C
      PICN            =   "frmAddNewCust.frx":6928
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblNewCust 
      Caption         =   "(Ýt nhÊt 4 ký tù)"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   10
      Left            =   4920
      TabIndex        =   23
      Tag             =   "L2"
      Top             =   480
      Width           =   1755
   End
   Begin VB.Label lblNewCust 
      Alignment       =   1  'Right Justify
      Caption         =   "Tªn c«ng ty:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   9
      Left            =   0
      TabIndex        =   21
      Tag             =   "L4"
      Top             =   1590
      Width           =   1875
   End
   Begin VB.Label lblNewCust 
      Alignment       =   1  'Right Justify
      Caption         =   "Balance:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   8
      Left            =   6480
      TabIndex        =   20
      Tag             =   "L11"
      Top             =   360
      Width           =   1125
   End
   Begin VB.Label lblNewCust 
      Alignment       =   1  'Right Justify
      Caption         =   "Discount:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   4800
      TabIndex        =   19
      Tag             =   "L10"
      Top             =   3480
      Width           =   1515
   End
   Begin VB.Label lblNewCust 
      Alignment       =   1  'Right Justify
      Caption         =   "Account No:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   6
      Left            =   660
      TabIndex        =   18
      Tag             =   "L9"
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblNewCust 
      Alignment       =   1  'Right Justify
      Caption         =   "TaxCode:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   6960
      TabIndex        =   17
      Tag             =   "L8"
      Top             =   2910
      Width           =   1035
   End
   Begin VB.Label lblNewCust 
      Alignment       =   1  'Right Justify
      Caption         =   "Fax:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   3900
      TabIndex        =   16
      Tag             =   "L7"
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label lblNewCust 
      Alignment       =   1  'Right Justify
      Caption         =   "Phone:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   15
      Tag             =   "L6"
      Top             =   2880
      Width           =   1875
   End
   Begin VB.Label lblNewCust 
      Alignment       =   1  'Right Justify
      Caption         =   "Address:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   14
      Tag             =   "L5"
      Top             =   2190
      Width           =   1875
   End
   Begin VB.Label lblNewCust 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer &Name:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   13
      Tag             =   "L3"
      Top             =   960
      Width           =   1875
   End
   Begin VB.Label lblNewCust 
      Alignment       =   1  'Right Justify
      Caption         =   "N&umber:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   12
      Tag             =   "L2"
      Top             =   420
      Width           =   1875
   End
End
Attribute VB_Name = "frmAddNewCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim res As New ADODB.Recordset
    Dim arrAddNew() As String
    Dim i, j As Integer
'           ------------- FORM -----------
Private Sub Form_Activate()
On Error GoTo errHdl
    Dim DescArr() As String
    Dim ctrl As Control
    
    DescArr = LoadLanguage(LngFile, "#01:008:")
    If cmdCreate.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(21)
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Activate"
End Sub

Public Sub Set_Promotion_type()
On Error GoTo Handle
Dim rsCust_Type As New ADODB.Recordset
Set rsCust_Type = Open_Table(cnData, "Customer_Type")
If rsCust_Type.State = 0 Then Exit Sub
If rsCust_Type.RecordCount = 0 Then Exit Sub
cboPro.Clear
With rsCust_Type
    Do While Not .EOF
       With cboPro
            .AddItem rsCust_Type.Fields("CustType_ID")
       End With
       .MoveNext
       Loop
End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Set_Promotion_type"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl
    Call Set_Promotion_type
    Set res = Open_Table(cnData, "Customer")
    ReDim Preserve arrAddNew(0)
    With Me
        .WindowState = 0
    End With
'    With res
'        For i = 0 To 9
'        DoEvents
'            txtNewCust(i).MaxLength = .Fields(i).DefinedSize
'        Next i
'    End With

    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Load"
End Sub

Private Sub cboCustomer_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 13 Then
        With txtNewCust(2)
            .SetFocus
            .SelStart = 0
            .SelLength = 9999
        End With
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- cboCustomer_KeyPress"
End Sub
'           ---------- COMMANDBUTTON -----------
Private Sub cmdCancel_Click()
On Error GoTo errHdl
    CloseRecordset res
    Unload Me
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- cmdCancel_Click"
End Sub

Private Sub cmdCreate_Click()
On Error GoTo errHdl
    Dim sTemp() As String
    Dim irow As Integer
    
    If CheckTextNull Then
        MsgBox "Th«ng tin kh«ng ®­îc rèng", vbInformation
        Exit Sub
    End If
    sTemp = SetTextTemp
    With frmCustomer.flexCustomer
        For i = 1 To .Rows - 1
        DoEvents
            If .TextMatrix(i, 0) = sTemp(0) Then
                txtNewCust(0).Text = ""
                txtNewCust(0).SetFocus
                Exit Sub
            End If
        Next i
        ReDim Preserve arrAddNew(UBound(arrAddNew) + 1)
        arrAddNew(UBound(arrAddNew)) = txtNewCust(0).Text
        If .TextMatrix(1, 0) <> "" Then
              .Rows = .Rows + 1
              irow = .Rows - 1
        Else: irow = 1
        End If
        For i = 0 To .Cols - 1
        DoEvents
            .TextMatrix(irow, i) = sTemp(i)
        Next i
'        SetColorFlexGrid frmCustomer.flexCustomer, irow, 1, .Cols
        .Refresh
    End With
    
    If MsgBox("B¹n cã muèn thªm tiÕp kh«ng?", vbYesNo) = 6 Then
        SetTextNull
        txtNewCust(0).SetFocus
    Else: Unload Me
    Call frmCustomer.AddNewRecords
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- cmdCancel_Click"
End Sub
'           ------------- TEXTBOX -----------
Private Sub SetTextNull()
On Error GoTo errHdl
    For i = 0 To txtNewCust.count - 1
    DoEvents
        txtNewCust(i).Text = ""
    Next i
    txtNewCust(0).SetFocus
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- SetTextNull"
End Sub

Private Sub txtNewCust_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl

    Dim tempIndex As Integer
    
    If KeyAscii = 33 Or KeyAscii = 44 Then
        KeyAscii = 0
        Exit Sub
    End If
    If KeyAscii = 13 Then
    
        Select Case Index
        Case 8: ' GoTo 1
            cmdCreate.SetFocus
        Case Else:
           tempIndex = Index + 1
        End Select
        If tempIndex <> -1 Then
            With txtNewCust(tempIndex)
                .SetFocus
                .SelStart = 0
                .SelLength = 9999
            End With
        End If
    End If
'1: cmdCreate.SetFocus
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- txtNewCust_KeyPress"
End Sub

Private Function CheckTextNull() As Boolean
On Error GoTo errHdl

    For i = 0 To txtNewCust.count - 1
    DoEvents
        If txtNewCust(0).Text = "" Or txtNewCust(1).Text = "" Then
            CheckTextNull = True
            Exit Function
        End If
    Next i
    CheckTextNull = False
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- CheckTextNull"
End Function

Private Function SetTextTemp()
On Error GoTo errHdl
    Dim S1() As String
    
    ReDim Preserve S1(res.Fields.count - 1)
    For i = 0 To 7
        S1(i) = txtNewCust(i).Text
    Next
        S1(9) = cboPro.Text
    SetTextTemp = S1
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- SetTextTemp"
End Function

Public Function Get_AddNewRecords()
On Error GoTo errHdl

    Dim Arr() As String
    
    ReDim Preserve Arr(UBound(arrAddNew))
    For i = 1 To UBound(arrAddNew)
    DoEvents
        Arr(i) = arrAddNew(i)
    Next i
    Get_AddNewRecords = Arr
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Get_AddNewRecords"
End Function

Private Sub txtNewCust_LostFocus(Index As Integer)
On Error GoTo Handle
    If Len(txtNewCust(0).Text) < 4 Then
        MsgBox "M· kh¸ch hµng tèi thiÓu 4 ký tù", vbInformation
        txtNewCust(0).SetFocus
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name
End Sub
