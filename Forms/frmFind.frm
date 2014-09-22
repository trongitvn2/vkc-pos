VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
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
   ScaleHeight     =   2475
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton cmdKeyboard 
      Height          =   795
      Left            =   3660
      TabIndex        =   6
      Top             =   210
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1402
      BTYPE           =   5
      TX              =   "Key board"
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
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFind.frx":0000
      PICN            =   "frmFind.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdFindNext 
      Height          =   855
      Left            =   240
      TabIndex        =   3
      Tag             =   "L3"
      Top             =   1530
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "T×m vÒ tr­íc"
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
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFind.frx":046E
      PICN            =   "frmFind.frx":048A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.CheckBox chkFind 
      Caption         =   "Match &case"
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
      Left            =   240
      TabIndex        =   2
      Tag             =   "L2"
      Top             =   1140
      Width           =   1695
   End
   Begin VB.ComboBox cboFind 
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
      Left            =   150
      TabIndex        =   0
      Top             =   480
      Width           =   3435
   End
   Begin prjTouchScreen.MyButton cmdFindPre 
      Height          =   855
      Left            =   1920
      TabIndex        =   4
      Tag             =   "L4"
      Top             =   1530
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "T×m vÒ sau"
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
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmFind.frx":05EA
      PICN            =   "frmFind.frx":0606
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   855
      Left            =   3600
      TabIndex        =   5
      Tag             =   "L5"
      Top             =   1530
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "Th&o¸t"
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
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmFind.frx":0794
      PICN            =   "frmFind.frx":07B0
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
      Caption         =   "Fi&nd what:"
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
      Left            =   150
      TabIndex        =   1
      Tag             =   "L1"
      Top             =   180
      Width           =   3450
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Public FormCall As Object
    Dim fSearch As Byte
    Dim flex As MSFlexGrid
    Dim i, J As Integer


Public Property Let GetfSearch(ByVal vNewValue As Variant)
    fSearch = vNewValue
End Property

Private Sub cboFind_KeyPress(KeyAscii As Integer)
    On Error GoTo handle
        If KeyAscii = 13 Then Call cmdFindNext_Click
    Exit Sub
handle:
    MsgBox Err.Number & Err.Description & Me.Name & " cboFind_KeyPress"
End Sub

Private Sub cmdKeyboard_Click()
    With frmKeyboard
        .FormCallkeyboard = "Other"
        .txtInput.PasswordChar = ""
        .Show vbModal
        cboFind.Text = .Let_Text_Input
    End With
    
End Sub

'           ------------ FORM ---------
Private Sub Form_Activate()
On Error GoTo errHdl

    Dim DescArr() As String
    Dim ctrl As Control
    
    DescArr = LoadLanguage(LngFile, "#02:001:")
    If cmdCancel.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = "Search....."
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl
'    With Me
'        .Height = 1680 '1575
'        .Width = 4035 '4185
'        .WindowState = 0
'    End With
    Set flex = SelectFlex
    InitValueForCboFind
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - Form_Load"
End Sub
'           ------------ COMBOBOX ---------
Private Sub cboFind_LostFocus()
On Error GoTo errHdl

    cboFind.AddItem cboFind.Text
    If InStr(1, str_Search, ";" & cboFind.Text & ";", 1) = 0 Then
        str_Search = str_Search & cboFind.Text & ";"
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - cboFind_LostFocus"
End Sub
'           ---------- COMMANDBUTTON ---------
Private Sub cmdCancel_Click()
On Error GoTo errHdl

    Unload Me
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - cmdCancel_Click"
End Sub

Private Sub cmdFindNext_Click()
On Error GoTo errHdl
    Dim irow As Integer, iCol As Integer
    Dim strfind As String
    Dim found As Boolean

    found = False
    strfind = Trim(cboFind.Text)
    With flex
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
            For J = .Col To .Cols - 1
            DoEvents
                If chkFind.Value = 1 Then
'                    If (StrComp(.TextMatrix(i, j), strfind, 0)) <> -1 Then
                    If (InStr(1, .TextMatrix(i, J), strfind, vbBinaryCompare)) <> 0 Then
                        .Row = i: .Col = J
                        .TopRow = .Row
                        found = True
                        Exit Sub
                    End If
                Else
                    If (InStr(1, .TextMatrix(i, J), strfind, 1)) <> 0 Then
                        .Row = i: .Col = J
                        .TopRow = .Row
                        found = True
                        Exit Sub
                    End If
                End If
            Next J
            .Col = 0
        Next i
        If Not found Then
            .Row = irow
            .Col = iCol
        End If
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - cmdFindNext_Click"
End Sub

Private Sub cmdFindPre_Click()
On Error GoTo errHdl

    Dim irow As Integer, iCol As Integer
    Dim strfind As String
    Dim found As Boolean
    Dim fAdd As Boolean
    
    found = False: fAdd = True
    For i = 0 To cboFind.ListCount - 1
    DoEvents
        If StrComp(cboFind.Text, cboFind.List(i), 1) = 0 Then
            fAdd = False
            Exit For
        End If
    Next i
    If fAdd Then cboFind.AddItem cboFind.Text
    strfind = Trim(cboFind.Text)
    With flex
        If strfind = "" Then Exit Sub
        irow = .Row: iCol = .Col
        If .Row = 1 And .Col = 0 Then Exit Sub
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
            For J = .Col To 0 Step -1
            DoEvents
                If chkFind.Value = 1 Then
'                    If (StrComp(.TextMatrix(i, j), strfind, 0)) <> -1 Then
                    If (InStr(1, .TextMatrix(i, J), strfind, vbBinaryCompare)) <> 0 Then
                        .Row = i: .Col = J
'                        .TopRow = .Row
                        If .Row < .TopRow Then .TopRow = .Row
                        found = True
                        Exit Sub
                    End If
                Else
                    If (InStr(1, .TextMatrix(i, J), strfind, 1)) <> 0 Then
                        .Row = i: .Col = J
'                        .TopRow = .Row
                        If .Row < .TopRow Then .TopRow = .Row
                        found = True
                        Exit Sub
                    End If
                End If
            Next J
            .Col = .Cols - 1
        Next i
        If Not found Then
            .Row = irow
            .Col = iCol
        End If
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - cmdFindPre_Click"
End Sub

Private Function SelectFlex() As MSFlexGrid
On Error GoTo errHdl

    Dim flex As MSFlexGrid
    If fSearch < 0 Then Exit Function
    If fSearch = 1 Then 'tim kiem trong form Clerk
        Set flex = FormCall.flexClerk
    ElseIf fSearch = 2 Then
        Set flex = FormCall.flexPLU
    ElseIf fSearch = 3 Then
        Set flex = FormCall.flex
    End If
    Set SelectFlex = flex
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - SelectFlex"
End Function

Private Sub InitValueForCboFind()
On Error GoTo errHdl

    Dim sTemp As String
    
    If str_Search = "" Then
        str_Search = ";"
        Exit Sub
    End If
    cboFind.Clear
    sTemp = str_Search
    Do While Len(sTemp) > 0
    DoEvents
        cboFind.AddItem Left(sTemp, InStr(1, sTemp, ";", 1) - 1)
        sTemp = Mid(sTemp, InStr(1, sTemp, ";", 1) + 1)
    Loop
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & " - InitValueForCboFind"
End Sub
