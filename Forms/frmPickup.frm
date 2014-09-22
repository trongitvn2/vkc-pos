VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPickup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "T×m kÝm th«ng tin"
   ClientHeight    =   10275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11385
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10275
   ScaleWidth      =   11385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton cmdPickup 
      Height          =   1155
      Left            =   9360
      TabIndex        =   8
      Top             =   2880
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   2037
      BTYPE           =   14
      TX              =   "Chän"
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmPickup.frx":0000
      PICN            =   "frmPickup.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdFindPre 
      Height          =   675
      Left            =   2760
      TabIndex        =   5
      Top             =   1350
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   1191
      BTYPE           =   14
      TX              =   ""
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmPickup.frx":0E3E
      PICN            =   "frmPickup.frx":0E5A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdFindNext 
      Height          =   675
      Left            =   4200
      TabIndex        =   7
      Top             =   1350
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1191
      BTYPE           =   14
      TX              =   ""
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
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmPickup.frx":0FE9
      PICN            =   "frmPickup.frx":1005
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid flexPickup 
      Height          =   7935
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   13996
      _Version        =   393216
      Cols            =   3
      BackColorFixed  =   -2147483643
      BackColorBkg    =   -2147483643
      GridColorFixed  =   12632256
      TextStyleFixed  =   3
      GridLinesFixed  =   1
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
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1800
      TabIndex        =   0
      Top             =   630
      Width           =   6855
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8685
      TabIndex        =   1
      Top             =   0
      Width           =   8745
      Begin VB.Label lblName 
         BackColor       =   &H80000008&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   2775
      End
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Height          =   1155
      Left            =   9360
      TabIndex        =   6
      Top             =   4200
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   2037
      BTYPE           =   14
      TX              =   "&§ãng"
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
      MICON           =   "frmPickup.frx":1193
      PICN            =   "frmPickup.frx":11AF
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblFind 
      Alignment       =   1  'Right Justify
      Caption         =   "&Chuçi cÇn t×m:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Tag             =   "L2"
      Top             =   840
      Width           =   1485
   End
End
Attribute VB_Name = "frmPickup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim fPickup As Byte, iCol As Integer
    Dim sCurValue As String
    Dim sHexCode As String
    Public FormCall As Object
    Dim i, j As Integer


Public Property Let GetfPickup(ByVal vNewValue As Variant)
    fPickup = vNewValue
End Property

Public Property Let GetCurrentValue(ByVal vCurValue As String)
    sCurValue = vCurValue
End Property

Public Property Let GetHexCode(ByVal vNewValue As Variant)
    sHexCode = vNewValue
End Property
'           --------------- FORM -------------
Private Sub Form_Activate()
    Dim DescArr() As String
    Dim ctrl As Control
    Dim iCount As Byte
    
    iCount = 0
    DescArr = LoadLanguage(LngFile, "#01:016:")
    If cmdPickup.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    For i = 1 To UBound(DescArr)
        DoEvents
        Select Case i
            Case 5, 6, 7
                flexPickup.TextMatrix(0, iCount) = DescArr(i)
                iCount = iCount + 1
            Case Else
        End Select
    Next
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    With flexPickup
        '.SetFocus
        .Row = 1
        .Col = 1
    End With
    'New
    If sCurValue <> "" Then
        If InStr(1, sCurValue, "CurRow", vbTextCompare) <> 0 Then
            flexPickup.Row = Mid(sCurValue, 7)
            If flexPickup.Row = 0 Then flexPickup.Row = 1
            flexPickup.TopRow = flexPickup.Row
        Else
            txtFind.Text = sCurValue
            cmdFindNext_Click
            'txtFind.Text = ""
        End If
    End If
    'End
End Sub

Private Sub Form_Load()
    Initialize
End Sub

Private Sub Initialize()
    SetDataInFlex
'    SetColorFlexGrid flexPickup, 1, 0, iCol
    With flexPickup
        .Row = 1
        .Col = 1
        .ColSel = .Cols - 1
        .ScrollTrack = True
    End With
End Sub

Private Sub txtFind_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        .PasswordChar = ""
        .Text = txtFind.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtFind.Text = .Let_Text_Input
        End With
        txtFind.SetFocus
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtFind_DblClick"
End Sub

'           -------------- TEXTBOX -----------
Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdFindNext_Click
End Sub
'           ------------ COMMANDBUTTON ---------
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFindNext_Click()
    Dim irow As Integer, iCol As Integer
    Dim strfind As String
    Dim found As Boolean
    Dim i As Integer
    Dim j As Integer
    
    found = False
    strfind = Trim(txtFind.Text)
    With flexPickup
        If strfind = "" Then Exit Sub
        irow = .Row: iCol = .Col
        If .Row = .Rows - 1 Then Exit Sub
        If .Col = .Cols - 1 Then
              .Row = .Row + 1
              If .Row = .Rows - 1 Then
                    Exit Sub
              Else: .Col = 0
              End If
        Else: .Col = .Col + 1
        End If
        For i = .Row To .Rows - 1
        DoEvents
            For j = .Col To .Cols - 1
            DoEvents
                If (InStr(1, .TextMatrix(i, j), strfind, 1)) <> 0 Then
                    .Row = i: .Col = j
                    .TopRow = .Row
                    found = True
                    Exit Sub
                End If
            Next j
            .Col = 0 'New
        Next i
        If Not found Then
            .Row = irow
            .Col = iCol
        End If
    End With
End Sub

Private Sub cmdFindPre_Click()
    Dim irow As Integer, iCol As Integer
    Dim strfind As String
    Dim found As Boolean
    Dim i As Integer
    Dim j As Integer
    
    found = False
    strfind = Trim(txtFind.Text)
    With flexPickup
        If strfind = "" Then Exit Sub
        irow = .Row: iCol = .Col
        If .Row = 1 Then Exit Sub
        If .Col = 0 Then
              .Row = .Row - 1
              If .Row = 1 Then
                    Exit Sub
              Else: .Col = .Cols - 1
              End If
        Else: .Col = .Col - 1
        End If
        For i = .Row To 1 Step -1
        DoEvents
            For j = .Col To 0 Step -1
            DoEvents
                If (InStr(1, .TextMatrix(i, j), strfind, 1)) <> 0 Then
                    .Row = i: .Col = j
                    .TopRow = .Row
                    found = True
                    Exit Sub
                End If
            Next j
            .Col = .Cols - 1 'New
        Next i
        If Not found Then
            .Row = irow
            .Col = iCol
        End If
    End With
End Sub

Private Sub cmdPickup_Click()
    Dim sCode As String
    Dim sName As String
    Dim fDuplicate As Boolean
    
    If flexPickup.TextMatrix(1, 0) = "" Then GoTo 1
    If fPickup = 19 Then 'frmSetMLink
        fDuplicate = False
        sCode = flexPickup.TextMatrix(flexPickup.Row, 1)
        sName = flexPickup.TextMatrix(flexPickup.Row, 2)
        With FormCall.flex
            If .TextMatrix(1, 0) = "" Then
                .TextMatrix(1, 0) = sCode
                .TextMatrix(1, 1) = sName
            Else
                For i = 1 To .Rows - 1
                DoEvents
                    If StrComp(.TextMatrix(i, 0), sCode, 1) = 0 Then
                        fDuplicate = True
                        Exit For
                    End If
                Next i
                If Not fDuplicate Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = sCode
                    .TextMatrix(.Rows - 1, 1) = sName
'                    Call SetColorFlexGrid(FormCall.flex, .Rows - 2, 1, .Cols)
                End If
            End If
        End With
'        Unload Me
        Exit Sub
    Else
        PickupValue
    End If
'    If InStr(1, FormCall.Name, "frmItems", 1) Then
'        FormCall.cmdAddPlu_Click
'        Exit Sub
'    End If
    If fPickup <> 18 Then GoTo 1
1:  Unload Me
End Sub

Private Sub PickupValue()
    Dim iBeginArrange As Integer
    Dim iEndArrange As Integer
    Dim sNo As String
    Dim flag As Boolean

    iBeginArrange = 90:  iEndArrange = 99
     
    With FormCall
'        sNo = flexPickup.TextMatrix(flexPickup.Row, 1)
        Select Case fPickup
            Case 10: sNo = flexPickup.TextMatrix(flexPickup.Row, 1)
            Case Else: sNo = flexPickup.TextMatrix(flexPickup.Row, 0)
        End Select
        
        Select Case fPickup
            Case 1:  .cboPLU(0).ListIndex = Val(sNo) - 1 'GA
            Case 10: .txtPLU(13).Text = sNo
            Case 20: .cboSelect.ListIndex = sNo - 1
        End Select
    End With
End Sub
'           ------------- FLEXGRID -----------
Private Sub flexPickup_Click()
    lblName.Caption = flexPickup.TextMatrix(flexPickup.Row, 1)
End Sub

Private Sub flexPickup_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdPickup_Click
End Sub

Private Sub flexPickup_DblClick()
    cmdPickup_Click
End Sub

Private Sub SetDataInFlex()
On Error GoTo Handle
       iCol = 3
    SetHeaderFlexGrid iCol
    Select Case fPickup
        Case 10: DataInFlex "Inventory", "ItemNum", "ItemName"
        Case 19: DataInFlex "Inventory", "ItemNum", "ItemName"
        Case 20: DataInFlex "SetMPLU", "ItemNum", "ItemName"
        Case 1: DataInFlex "Departments", "Dept_ID", "Description"
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " SetDataInFlex"
    End Select
End Sub

Private Sub SetHeaderFlexGrid(ByVal iCol As Integer)
    With flexPickup
        .Cols = iCol
        .AllowUserResizing = flexResizeBoth
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionFree
        .ColWidth(0) = 800
        .ColAlignment(1) = 4
        .ColAlignment(2) = 4
        .ColWidth(1) = 2500
        .ColWidth(2) = 5000

    End With
End Sub

Private Sub DataInFlex(sTableName As String, ByVal Field1 As String, ByVal Field2 As String)
    Dim res As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim irow As Integer
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Select Case fPickup
        Case 10: Set rs = OpenCriticalTable("select  ItemNum,ItemName,F4 from Inventory ", cnData)
            With res
            If .State = 0 Then
                .Fields.Append "ItemNum", adVarWChar, 20
                .Fields.Append "ItemName", adVarWChar, 50
                .Open
            End If
            rs.MoveFirst
            Do While Not rs.EOF
                If Mid(Right("00000000" & HexToBin(rs.Fields("F4")), 8), 3, 1) = 1 Then
                    .addNew
                    .Fields("ItemNum") = rs.Fields("ItemNum")
                    .Fields("ItemName") = rs.Fields("ItemName")
                    .Update
                End If
                rs.MoveNext
            Loop
        End With
        Case 19
            Set res = OpenCriticalTable("Select * from Inventory where ItemNum not in (Select PluCode from SetMLink) ", cnData)
        Case Else: Set res = Open_Table(cnData, sTableName)
    End Select
    
    With res
        If .RecordCount = 0 Then Exit Sub
        flexPickup.Rows = .RecordCount + 1
        .Sort = .Fields(0).name & " ASC"
        irow = 1
        .MoveFirst
        Do While Not .EOF
        DoEvents
            Select Case fPickup
                Case 0, 19 'grid 3 cot
                    flexPickup.TextMatrix(irow, 0) = irow
                    flexPickup.TextMatrix(irow, 1) = .Fields("ItemNum")
                    flexPickup.TextMatrix(irow, 2) = .Fields("ItemName")
                Case 20
                    flexPickup.TextMatrix(irow, 0) = irow
                    flexPickup.TextMatrix(irow, 1) = .Fields("PluCode")
                    flexPickup.TextMatrix(irow, 2) = .Fields("PluName")
                Case Else      'grid 2 cot
                    flexPickup.TextMatrix(irow, 0) = irow
                    flexPickup.TextMatrix(irow, 1) = .Fields(Field1)
                    flexPickup.TextMatrix(irow, 2) = .Fields(Field2)
            End Select
            irow = irow + 1
            .MoveNext
        Loop
    End With
    CloseRecordset res
End Sub

Private Sub DataInFlexByCombo(ByVal cbo As ComboBox, ByVal ilength As Integer, ByVal fItemData As Boolean)
    With cbo
        flexPickup.Rows = .ListCount + 1
        For i = 0 To .ListCount - 1 Step 1
        DoEvents
            If fItemData Then
                flexPickup.TextMatrix(i + 1, 0) = Hex(FillZeroForString(.ItemData(i), ilength))
            Else
                flexPickup.TextMatrix(i + 1, 0) = Right("00" & i, ilength)
            End If
            flexPickup.TextMatrix(i + 1, 1) = .List(i)
        Next i
    End With
End Sub

Private Sub DataInFlexByCombo_3Col(ByVal cbo As ComboBox, ByVal ilength As Integer)
    With cbo
        flexPickup.Rows = .ListCount + 1
        For i = 0 To .ListCount - 1 Step 1
        DoEvents
            flexPickup.TextMatrix(i + 1, 0) = FillZeroForString(CStr(i + 1), ilength)
            flexPickup.TextMatrix(i + 1, 1) = .ItemData(i)
            flexPickup.TextMatrix(i + 1, 2) = .List(i)
        Next i
    End With
End Sub

