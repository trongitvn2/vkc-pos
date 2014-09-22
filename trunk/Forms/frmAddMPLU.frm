VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddMPLU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Menu PLU"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6390
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
   ScaleHeight     =   2850
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton cmdCreate 
      Height          =   735
      Left            =   4740
      TabIndex        =   8
      Top             =   240
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1296
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
      MICON           =   "frmAddMPLU.frx":0000
      PICN            =   "frmAddMPLU.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.TextBox txtPLUCode 
      Height          =   420
      Index           =   1
      Left            =   1650
      TabIndex        =   7
      Top             =   1680
      Width           =   3030
   End
   Begin VB.TextBox txtPLUCode 
      Height          =   420
      Index           =   0
      Left            =   1650
      TabIndex        =   6
      Top             =   1230
      Width           =   3030
   End
   Begin VB.Frame fraCreatePLU 
      Caption         =   "Create New PLU"
      Height          =   1125
      Left            =   0
      TabIndex        =   1
      Tag             =   "L16"
      Top             =   60
      Width           =   4680
      Begin VB.OptionButton optCreatePLU 
         Caption         =   "&Single PLU"
         Height          =   300
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Tag             =   "L8"
         Top             =   285
         Width           =   2775
      End
      Begin VB.OptionButton optCreatePLU 
         Caption         =   "&Range of PLUs"
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Tag             =   "L9"
         Top             =   615
         Width           =   3015
      End
   End
   Begin MSComctlLib.ProgressBar probar 
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   2400
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   735
      Left            =   4740
      TabIndex        =   9
      Top             =   1110
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1296
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
      MICON           =   "frmAddMPLU.frx":0656
      PICN            =   "frmAddMPLU.frx":0672
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblBeginPLU 
      Alignment       =   1  'Right Justify
      Caption         =   "New PLU Code:"
      Height          =   255
      Left            =   75
      TabIndex        =   5
      Tag             =   "L10"
      Top             =   1230
      Width           =   1575
   End
   Begin VB.Label lblEndPLU 
      Alignment       =   1  'Right Justify
      Caption         =   "End PLU Code:"
      Height          =   255
      Left            =   45
      TabIndex        =   4
      Tag             =   "L11"
      Top             =   1665
      Width           =   1575
   End
End
Attribute VB_Name = "frmAddMPLU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim res As New ADODB.Recordset
    Dim arrPLU() As String
    Dim arrAddNew() As String
    Dim fcheckplucode As Byte
    Dim i, j As Integer
'           ---------- FORM -----------
Private Sub Form_Activate()
On Error GoTo errHdl

    Dim DescArr() As String
    Dim ctrl As Control
    
    DescArr = LoadLanguage(LngFile, "#01:013:")
    If cmdCreate.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(16)
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    optCreatePLU(0).SetFocus
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set res = Open_Table(cnData, "SetMPLU")
    txtPluCode(0).MaxLength = res.Fields(0).DefinedSize
    txtPluCode(1).MaxLength = txtPluCode(0).MaxLength
    ReDim Preserve arrAddNew(0)
    With Me
        .Height = 3000
        .Width = 7400
        .WindowState = 0
    End With
    EndPLU_status False
    optCreatePLU(0).Value = True
    probar.Visible = False
    txtPluCode(0).Text = Get_Max_MPLUCODE
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Load"

End Sub
'           ----------- COMMAND BUTTON ----------
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()
On Error GoTo errHdl
    HideControl True
    If CheckPLUCode Then
        If optCreatePLU(0).Value = True Then
            AddNewOnePLU
        ElseIf optCreatePLU(1).Value = True Then
            AddNewMultiPLU
        End If
        probar.Visible = False
   
    End If
    HideControl False

    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- cmdCreate_Click"

End Sub
'           ----------- TEXTBOX -----------
Private Sub txtPluCode_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl
    Select Case KeyAscii
        Case 13
            Select Case Index
                Case 0
                        If optCreatePLU(1).Value = True Then
                            With txtPluCode(1)
                                .SetFocus
                                .SelStart = 0
                                .SelLength = 9999
                            End With
                        Else
                            cmdCreate.SetFocus
                        End If
                Case 1: cmdCreate.SetFocus
            End Select
        Case Is < 32, 48 To 57
        Case Else
            KeyAscii = 0
    End Select
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- txtPLUCode_KeyPress"
End Sub
'           ---------- OPTION BUTTON ----------
Private Sub optCreatePLU_Click(Index As Integer)
On Error GoTo errHdl

    If Index = 1 Then
        EndPLU_status True
    Else: EndPLU_status False
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- optCreatePLU_Click"
End Sub

Private Sub optCreatePLU_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl
    Select Case KeyAscii
        Case 13
                With txtPluCode(0)
                    .SetFocus
                    .SelStart = 0
                    .SelLength = 9999
                End With
        Case vbKeyDown, vbKeyUp
                If Index = 0 Then
                      optCreatePLU(1).SetFocus
                Else: optCreatePLU(0).SetFocus
                End If
    End Select
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- optCreatePLU_KeyPress"
End Sub

Private Sub EndPLU_status(flag As Boolean) 'an hoac hien lblEndPlu & txtEndPlu
On Error GoTo errHdl

    If flag Then
        lblBeginPLU.top = 1170
        txtPluCode(0).top = 1050
    Else
        lblBeginPLU.top = lblEndPLU.top
        txtPluCode(0).top = txtPluCode(1).top
    End If
    lblEndPLU.Visible = flag
    txtPluCode(1).Visible = flag
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- EndPLU_status"
End Sub
'           ----------- FUNCTIONS ADD NEW RECORD ---------
Private Function CheckPLUCode() As Boolean
On Error GoTo errHdl
    Dim sPluCode1 As String, sPluCode2 As String
    
    sPluCode1 = Trim(txtPluCode(0).Text)
    sPluCode2 = Trim(txtPluCode(1).Text)
    fcheckplucode = 0
    If sPluCode1 = "" Then fcheckplucode = 1: GoTo 1
    If CDbl(sPluCode1) < 0 Then fcheckplucode = 3: GoTo 1
    If frmSetMPLU.flex.Rows - 1 > 999999 Then fcheckplucode = 4: GoTo 1
    
    If optCreatePLU(1).Value Then
        If sPluCode2 = "" Then
            fcheckplucode = 1
            GoTo 1
        ElseIf CDbl(sPluCode1) > CDbl(sPluCode2) Then
            fcheckplucode = 2
            GoTo 1
        ElseIf CDbl(sPluCode1) < 0 Or CDbl(sPluCode2) < 0 Then
            fcheckplucode = 3
            GoTo 1
        End If
    End If
1:  If fcheckplucode <> 0 Then
        CheckPLUCode = False
        Exit Function
    Else: CheckPLUCode = True
    End If
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- CheckPLUCode"
End Function

Private Sub AddNewOnePLU()
On Error GoTo errHdl
    Dim sPLU As String
    Dim iLen As Integer
    
    sPLU = txtPluCode(0).Text
    iLen = res.Fields("PLUCode").DefinedSize
    If Len(sPLU) <= iLen Then
        sPLU = FillZeroForString(sPLU, iLen)
        With frmSetMPLU.flex
            For i = 1 To .Rows - 1
                DoEvents
                If .TextMatrix(i, 0) = sPLU Then
                    MsgBox "M· hµng nµy ®· cã trong danh môc", vbInformation
                 Exit Sub
                End If
            Next
        End With
    End If
    AddDataToGrid sPLU
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- AddNewOnePLU"
End Sub

Private Sub AddNewMultiPLU()
On Error GoTo errHdl
    Dim istartPLU As Double
    Dim iendPLU As Double
    Dim iLen As Integer
    Dim iInc As Integer
    Dim arrTemp() As String
    Dim sPLU As String
    Dim i As Double
    Dim j As Double
    
    istartPLU = CDbl(txtPluCode(0).Text)
    iendPLU = CDbl(txtPluCode(1).Text)
    If iendPLU > 999999 Then iendPLU = 999999
    If istartPLU = iendPLU Then _
        AddNewOnePLU: Exit Sub
    
    iLen = res.Fields("PluCode").DefinedSize
    
    
    ReDim Preserve arrTemp(iendPLU - istartPLU + 1)
    iInc = 0
    With frmSetMPLU.flex
        For i = istartPLU To iendPLU
            DoEvents
            iInc = iInc + 1
            For j = 1 To .Rows - 1
                DoEvents
                arrTemp(iInc) = FillZeroForString(Trim(CStr(i)), 6)
                If .TextMatrix(j, 0) = arrTemp(iInc) Then
                    MsgBox "M· hµng"
                    Exit Sub
                End If
            Next j
        Next i
        iInc = 0
        InitProgressBar istartPLU, iendPLU
        For j = istartPLU To iendPLU
            DoEvents
            iInc = iInc + 1
            sPLU = arrTemp(iInc)
            AddDataToGrid sPLU
            probar.Value = probar.Max - (iendPLU - j)
        Next j
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- AddNewMultiPLU"
End Sub

Private Sub AddDataToGrid(sPLU As String)
On Error GoTo errHdl
    Dim irow As Integer
        
    ReDim Preserve arrAddNew(UBound(arrAddNew) + 1)
    arrAddNew(UBound(arrAddNew)) = sPLU
        
    With frmSetMPLU.flex
        If .TextMatrix(1, 0) <> "" Then
              .Rows = .Rows + 1
              irow = .Rows - 1
        Else: irow = 1
        End If
        .TextMatrix(irow, 0) = sPLU
        .TextMatrix(irow, 1) = "PLU-NAME " & CDbl(sPLU)
        .TextMatrix(irow, 2) = "0"
        .TextMatrix(irow, 3) = "Kg"
        .TextMatrix(irow, 4) = "2"
'        SetColorFlexGrid frmSetMPLU.flex, irow, 1, .Cols
        .Refresh
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- AddDataToGrid"
End Sub
'           ---------- OTHER FUNCTIONS ----------
Private Sub HideControl(fHide As Boolean)
On Error GoTo errHdl
    cmdCancel.Enabled = Not fHide
    cmdCreate.Enabled = Not fHide
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- AddDataToGrid"
End Sub

Private Sub InitProgressBar(ByVal iBegin As Double, ByVal iEnd As Double)
On Error GoTo errHdl
    With probar
        .Visible = True
        .Value = 0
        .Min = 0
        .Max = iEnd - iBegin
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- InitProgressBar"
End Sub

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
    & Me.name & "- InitProgressBar"
End Function
Public Function Get_Max_MPLUCODE() As String
On Error GoTo Handle
    Dim rsmax As New ADODB.Recordset
    Set rsmax = OpenCriticalTable("select max(PLUCode) as Max_PLU from SetMPLU", cnData)
    If Not rsmax.EOF Then
        If "" & rsmax.Fields("Max_PLU") = "" Then
            Get_Max_MPLUCODE = "1"
        Else
            Get_Max_MPLUCODE = rsmax.Fields("Max_PLU") + 1
        End If
    Else
        Get_Max_MPLUCODE = rsmax.Fields("Max_PLU") + 1
    End If
Exit Function
Handle:
MsgBox Err.Number & Err.Description & "  mdlGlobal " & "   Get_Max_MPLUCODE"

End Function

