VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeleteMPLU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Xãa danh môc nguyªn liÖu"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7320
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
   ScaleHeight     =   3195
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCreatePLU 
      Caption         =   "Delete"
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
      Height          =   1380
      Left            =   0
      TabIndex        =   2
      Tag             =   "L17"
      Top             =   120
      Width           =   5610
      Begin VB.OptionButton optDeletePLU 
         Caption         =   "&Range of PLUs"
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
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Tag             =   "L19"
         Top             =   810
         Visible         =   0   'False
         Width           =   2580
      End
      Begin VB.OptionButton optDeletePLU 
         Caption         =   "&Single PLU"
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
         Height          =   330
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Tag             =   "L18"
         Top             =   285
         Width           =   2580
      End
   End
   Begin VB.ComboBox cboPLUCode 
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
      Index           =   0
      Left            =   1635
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1635
      Width           =   3975
   End
   Begin VB.ComboBox cboPLUCode 
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
      Index           =   1
      Left            =   1635
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   2205
      Width           =   3975
   End
   Begin MSComctlLib.ProgressBar probar 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   2805
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin prjTouchScreen.MyButton cmdDelete 
      Height          =   735
      Left            =   5670
      TabIndex        =   8
      Tag             =   "L22"
      Top             =   540
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "§ång ý"
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
      MICON           =   "frmDeleteMPLU.frx":0000
      PICN            =   "frmDeleteMPLU.frx":001C
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
      Height          =   735
      Left            =   5670
      TabIndex        =   9
      Tag             =   "L23"
      Top             =   1470
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "Tho¸t"
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
      MICON           =   "frmDeleteMPLU.frx":0656
      PICN            =   "frmDeleteMPLU.frx":0672
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblEndPLU 
      Alignment       =   1  'Right Justify
      Caption         =   "End PLU Code:"
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
      Height          =   315
      Left            =   15
      TabIndex        =   7
      Tag             =   "L21"
      Top             =   2250
      Width           =   1575
   End
   Begin VB.Label lblBeginPLU 
      Alignment       =   1  'Right Justify
      Caption         =   "Start PLU Code:"
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
      Height          =   375
      Left            =   -105
      TabIndex        =   6
      Tag             =   "20"
      Top             =   1665
      Width           =   1695
   End
End
Attribute VB_Name = "frmDeleteMPLU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim res As New ADODB.Recordset
    Dim istartPLU As Double
    Dim iendPLU As Double
    Dim arrDelete() As String
    Dim i As Integer
'           ---------- FORM ---------
Private Sub Form_Activate()
On Error GoTo errHdl

    Dim DescArr() As String
    Dim ctrl As Control
    
    DescArr = LoadLanguage(LngFile, "#01:013:")
    If cmdDelete.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(17)
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next
    optDeletePLU(0).SetFocus
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl

    ReDim Preserve arrDelete(0)
    With Me
        .WindowState = 0
    End With
    EndPLU_status False
    optDeletePLU(0).Value = True
    probar.Visible = False
    InitCombo
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Form_Load"
End Sub

Private Sub SetComboPLU(cbo As ComboBox)
On Error GoTo errHdl

    cbo.Clear
    With frmSetMPLU.flex
        If .TextMatrix(1, 0) = "" Then Exit Sub
        For i = 1 To .Rows - 1
        DoEvents
            cbo.AddItem .TextMatrix(i, 0) & "   " & .TextMatrix(i, 1)
            cbo.ItemData(cbo.NewIndex) = i
        Next i
        cbo.ListIndex = .Row - 1
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- SetComboPLU"
End Sub
'           ----------- COMMANDBUTTON --------
Private Sub cmdCancel_Click()
On Error GoTo errHdl

    Unload Me
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- cmdCancel_Click"
End Sub

Private Sub cmdDelete_Click()
On Error GoTo errHdl
    HideControl True
'    istartPLU = FillZeroForString(Trim(str(cboPLUCode(0).Text)), 6)
'    iendPLU = FillZeroForString(Trim(str(cboPLUCode(1).Text)), 6)
    If optDeletePLU(0).Value = True Or _
      (optDeletePLU(1).Value = True And istartPLU = iendPLU) Then
        DeleteOnePLU
        GoTo 1
    End If
    If optDeletePLU(1).Value = True Then
        DeleteMultiPLU
        probar.Visible = False
    End If
1:
    InitCombo
    HideControl False
    If frmSetMPLU.flex.TextMatrix(1, 0) = "" Then Unload Me
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- cmdDelete_Click"
End Sub
'           ---------- OPTIONBUTTON ----------
Private Sub optDeletePLU_Click(Index As Integer)
On Error GoTo errHdl

    If Index = 1 Then
          EndPLU_status True
    Else: EndPLU_status False
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- optDeletePLU_Click"
End Sub

Private Sub optDeletePLU_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl

    Select Case KeyAscii
        Case 13: cboPLUCode(0).SetFocus
        Case vbKeyDown, vbKeyUp
                If Index = 0 Then
                      optDeletePLU(1).SetFocus
                Else: optDeletePLU(0).SetFocus
                End If
    End Select
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- optDeletePLU_KeyPress"
End Sub
'           ---------- COMBOBOX ---------
Private Sub cboPLUCode_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 13 Then
        Select Case Index
            Case 0
                    If optDeletePLU(1).Value = True Then
                          cboPLUCode(1).SetFocus
                    Else: cmdDelete.SetFocus
                    End If
            Case 1: cmdDelete.SetFocus
        End Select
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- cboPLUCode_KeyPress"
End Sub
'           ---------- FUNCTIONS DELETE --------
Private Sub EndPLU_status(flag As Boolean) 'an hoac hien lblEndPlu & txtEndPlu
On Error GoTo errHdl
        
    If flag Then
        lblBeginPLU.top = fraCreatePLU.top + fraCreatePLU.Height + 500
        cboPLUCode(0).top = lblBeginPLU.top
        lblEndPLU.top = lblBeginPLU.top + 500
        cboPLUCode(1).top = lblBeginPLU.top + 500
    Else
        lblBeginPLU.top = lblEndPLU.top
        cboPLUCode(0).top = cboPLUCode(1).top
    End If
    lblEndPLU.Visible = flag
    cboPLUCode(1).Visible = flag
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- EndPLU_status"
End Sub

Private Function CheckPLUCode() As Boolean
On Error GoTo errHdl

    Dim sValuePLUCode As String
    Dim flag As Boolean
    
    flag = True
    sValuePLUCode = cboPLUCode(0).Text
    istartPLU = CDbl(Left(sValuePLUCode, InStr(sValuePLUCode, "   ") - 1))
    sValuePLUCode = cboPLUCode(1).Text
    If optDeletePLU(1).Value Then
        iendPLU = CDbl(Left(sValuePLUCode, InStr(sValuePLUCode, "   ") - 1))
        If istartPLU > iendPLU Then
            flag = False
        End If
    End If
    CheckPLUCode = flag
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- CheckPLUCode"
End Function

Private Sub DeleteOnePLU()
On Error GoTo errHdl

    Dim sPLU As String
    Dim i As Integer
    
    sPLU = Left(cboPLUCode(0).Text, 6) 'FillZeroForString(Trim(str(istartPLU)), 6)
'    sPLU = Left(istartPLU, 6)
    If MsgBox("B¹n cã muèn xãa nguyªn liÖu nµy kh«ng?", vbOKCancel) = 1 Then
        With frmSetMPLU.flex
            For i = 1 To .Rows - 1
            DoEvents
                If .TextMatrix(i, 0) = sPLU Then
                    If .Rows = 2 Then ' if deleted row is last row then set it Null
                        Delete_Last_Row i
                    Else
                        .RemoveItem i
'                        SetColorFlexGrid frmSetMPLU.flex, i - 1, 1, .Cols
                    End If
                    ReDim Preserve arrDelete(UBound(arrDelete) + 1)
                    arrDelete(UBound(arrDelete)) = sPLU
                    Exit For
                End If
            Next i
        End With
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- DeleteOnePLU"
End Sub

Private Sub DeleteMultiPLU()
On Error GoTo errHdl

    Dim iTemp As Integer
    Dim i As Integer
    Dim J As Integer
    
    If MsgBox("B¹n cã muèn xãa mét d·y nguyªn liÖu nµy?", vbOKCancel) = 1 Then
        InitProgressBar
        With frmSetMPLU.flex
            For i = iendPLU To istartPLU Step -1
            DoEvents
                For J = .Rows - 1 To 1 Step -1
                DoEvents
                    If .TextMatrix(J, 0) = i Then
                        If .Rows = 2 Then ' if deleted row is last row then set it Null
                            Delete_Last_Row i
                            iTemp = i
                        Else
                            If i = istartPLU Then iTemp = J
                            .RemoveItem J
                        End If
                        ReDim Preserve arrDelete(UBound(arrDelete) + 1)
                        arrDelete(UBound(arrDelete)) = FillZeroForString(Trim(str(i)), 6)
                    End If
                Next J
                probar.Value = iendPLU - i
            Next i
'            SetColorFlexGrid frmSetMPLU.flex, iTemp, 1, .Cols
        End With
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- DeleteMultiPLU"
End Sub
'               --------- OTHER FUNCTIONS ------
Private Sub HideControl(fHide As Boolean)
On Error GoTo errHdl

    cmdDelete.Enabled = Not fHide
    cmdCancel.Enabled = Not fHide
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- HideControl"
End Sub

Private Sub InitProgressBar()
On Error GoTo errHdl

    With probar
        .Visible = True
        .Min = 0
        .Max = iendPLU - istartPLU
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- InitProgressBar"
End Sub

Public Function Get_DeleteRecords()
On Error GoTo errHdl

    Dim Arr() As String
    
    ReDim Preserve Arr(UBound(arrDelete))
    For i = 1 To UBound(arrDelete)
    DoEvents
        Arr(i) = arrDelete(i)
    Next i
    Get_DeleteRecords = Arr
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Get_DeleteRecords"
End Function

Private Sub InitCombo()
On Error GoTo errHdl

    SetComboPLU cboPLUCode(0)
    SetComboPLU cboPLUCode(1)
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- InitCombo"
End Sub

Private Sub Delete_Last_Row(irow As Integer)
On Error GoTo errHdl

    Dim k As Byte
    
    With frmSetMPLU.flex
        For k = 0 To .Cols - 1
        DoEvents
            .TextMatrix(irow, k) = ""
        Next k
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Delete_Last_Row"
End Sub
