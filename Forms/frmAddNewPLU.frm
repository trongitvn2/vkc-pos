VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddNewPLU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create PLU"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7980
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
   ScaleHeight     =   4275
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4950
      Top             =   1275
   End
   Begin VB.TextBox txtHint 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1050
      TabIndex        =   9
      Top             =   2760
      Width           =   3615
   End
   Begin MSComctlLib.ProgressBar probar 
      Height          =   375
      Left            =   4140
      TabIndex        =   0
      Top             =   3690
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame fraCreatePLU 
      Caption         =   "Create New PLU"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1830
      Left            =   120
      TabIndex        =   4
      Tag             =   "L3"
      Top             =   0
      Width           =   5850
      Begin VB.OptionButton optCreatePLU 
         Caption         =   "&Single PLU"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Tag             =   "L4"
         Top             =   570
         Width           =   2775
      End
      Begin VB.OptionButton optCreatePLU 
         Caption         =   "&Range of PLUs"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Tag             =   "L5"
         Top             =   1170
         Width           =   2775
      End
   End
   Begin VB.CheckBox chkCreatePLU 
      Caption         =   "Create &having selected PLU's preset data."
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   15
      TabIndex        =   3
      Tag             =   "L8"
      Top             =   3660
      Width           =   4005
   End
   Begin VB.TextBox txtPLUCode 
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
      Index           =   0
      Left            =   1800
      MaxLength       =   14
      TabIndex        =   2
      Top             =   2160
      Width           =   3825
   End
   Begin VB.TextBox txtPLUCode 
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
      Index           =   1
      Left            =   1830
      MaxLength       =   14
      TabIndex        =   1
      Top             =   2910
      Width           =   3825
   End
   Begin prjTouchScreen.MyButton cmdCreate 
      Height          =   795
      Left            =   6255
      TabIndex        =   10
      Tag             =   "L9"
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1402
      BTYPE           =   14
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
      BCOL            =   16777215
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAddNewPLU.frx":0000
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
      Height          =   795
      Left            =   6255
      TabIndex        =   11
      Tag             =   "L10"
      Top             =   1350
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1402
      BTYPE           =   14
      TX              =   "&Gióp ®ì"
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
      BCOL            =   16777215
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAddNewPLU.frx":001C
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
      Height          =   795
      Left            =   6270
      TabIndex        =   12
      Tag             =   "L11"
      Top             =   2250
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1402
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
      BCOL            =   16777215
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAddNewPLU.frx":0038
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
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   -90
      TabIndex        =   8
      Tag             =   "L6"
      Top             =   2190
      Width           =   1815
   End
   Begin VB.Label lblEndPLU 
      Alignment       =   1  'Right Justify
      Caption         =   "End PLU Code:"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   30
      TabIndex        =   7
      Tag             =   "L7"
      Top             =   2985
      Width           =   1575
   End
End
Attribute VB_Name = "frmAddNewPLU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim res As New ADODB.Recordset
    Public FormCall As Object
    Dim array_PLUCodes As String
    Dim str_NewPLUs As String
    Dim fcheckplucode As Byte
    Dim iMaxPLU As Double
    Dim sTime As Double
    Dim i, J As Integer
    Dim DescArr() As String

Public Property Let Get_MaxPLU(ByVal vNewValue As Variant)
    iMaxPLU = vNewValue
End Property

Public Property Let Get_CurPLUs(ByVal vNewValue As String)
    array_PLUCodes = vNewValue
End Property
'           ---------- FORM -----------
Private Sub Form_Activate()
On Error GoTo errHdl

    
    Dim ctrl As Control
    
    DescArr = LoadLanguage(LngFile, "#02:011:")
    If cmdCreate.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    txtHint.Text = DescArr(2)
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    'optCreatePLU(0).SetFocus
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl
    Dim strSQL  As String
    Dim intCount As Integer
    
    str_NewPLUs = ""
        
    Set res = Open_Table(cnData, "Inventory")
    If res.State = 1 Then res.Close
    
    Set res = Open_Table(cnData, "Inventory")
    
    txtHint.Visible = False
    
    EndPLU_status False
    optCreatePLU(0).Value = True
    probar.Visible = False
    txtPLUCode(0).Text = Get_Max_PLUCODE
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CloseRecordset res
End Sub
'           ----------- COMMAND BUTTON ----------
Private Sub cmdCancel_Click()
    CloseRecordset res
    Unload Me
End Sub

Private Sub cmdCreate_Click()
On Error GoTo errHdl

    Dim arrTempPLU() As String
    
    HideControl True
    If CheckPLUCode Then
        arrTempPLU = SetTextTemp
        If optCreatePLU(0).Value = True Then
            AddNewOnePLU arrTempPLU
        ElseIf optCreatePLU(1).Value = True Then
            AddNewMultiPLU arrTempPLU
        End If
        probar.Visible = False
    End If
    HideControl False
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- cmdCreate_Click"
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
    & Me.Name & "- optCreatePLU_Click"
End Sub

Private Sub optCreatePLU_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl

    Select Case KeyAscii
        Case 13
                With txtPLUCode(0)
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
    & Me.Name & "- optCreatePLU_KeyPress"
End Sub

Private Sub EndPLU_status(flag As Boolean) 'an hoac hien lblEndPlu & txtEndPlu
On Error GoTo errHdl
    If flag Then
        lblBeginPLU.top = 2100
        txtPLUCode(0).top = 2100
    Else
        lblBeginPLU.top = 2100
        txtPLUCode(0).top = 2100
        lblEndPLU.top = lblBeginPLU.top + 600
        txtPLUCode(1).top = txtPLUCode(0).top + 600
    End If
    lblEndPLU.Visible = flag
    txtPLUCode(1).Visible = flag
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- EndPLU_status"
End Sub
'           ---------- CHECKBOX -----------
Private Sub chkCreatePLU_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii = 13 Then cmdCreate.SetFocus
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- chkCreatePLU_KeyPress"
End Sub

Private Sub chkCreatePLU_GotFocus()
On Error GoTo errHdl

    sTime = Timer
    txtHint.Visible = True
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- chkCreatePLU_GotFocus"
End Sub

Private Sub chkCreatePLU_LostFocus()
On Error GoTo errHdl

    txtHint.Visible = False
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- chkCreatePLU_GotFocus"
End Sub
'           ----------- FUNCTIONS ADD NEW RECORDs ---------
Private Function CheckPLUCode() As Boolean
On Error GoTo errHdl

    Dim sPluCode1 As String
    Dim sPluCode2 As String
        
    sPluCode1 = Trim(txtPLUCode(0).Text)
    sPluCode2 = Trim(txtPLUCode(1).Text)
    fcheckplucode = 0
    If sPluCode1 = "" Then fcheckplucode = 1: GoTo 1
    If CDbl(sPluCode1) < 0 Then fcheckplucode = 3: GoTo 1
    'If FormCall.flexPLU.Rows - 1 > iMaxPLU Then fcheckplucode = 4: GoTo 1
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
    & Me.Name & "- CheckPLUCode"
End Function

Private Sub AddNewOnePLU(Arr() As String)
On Error GoTo errHdl

    Dim sPLU As String
    Dim iLen As Integer
    Dim strPluCodeTemp As String
    
    sPLU = txtPLUCode(0).Text
    
    strPluCodeTemp = sPLU
    iLen = Len(strPluCodeTemp)
    
    If Len(sPLU) <= iLen Then
        'khong can sua FillZeroForString cho nay
        sPLU = FillZeroForString(sPLU, 12)
'         sPLU = FillZeroForString(sPLU, 20)
        If InStr(1, array_PLUCodes, ";" & sPLU & ";", vbBinaryCompare) <> 0 Then
            'MsgBox arrMessage(13), myInformation, myClose, arrMessage(1)
            MsgBox "M· hµng nµy ®· tån t¹i råi !Vui lßng t¹o l¹i m· kh¸c", vbInformation
            Exit Sub
        End If
        '---------------
        Dim iBegin As Integer
        Dim iEnd As Integer
        Dim iExprice As Integer
        
        array_PLUCodes = array_PLUCodes & sPLU & ";"
        AddDataToGrid sPLU, Arr, iBegin, iEnd, iExprice
'        SetColorFlexGrid FormCall.flexPLU, FormCall.flexPLU.Rows - 1, 1, FormCall.flexPLU.Cols
    MsgBox DescArr(12), vbOKOnly
    End If
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- AddNewOnePLU"
End Sub

Private Sub AddNewMultiPLU(Arr() As String)
On Error GoTo errHdl

    Dim istartPLU As String, iendPLU As String
    Dim arrayNewPLU() As String
    Dim sPLU As String
    Dim i As Double, J As Double
    Dim iLen As Integer
    
    istartPLU = (txtPLUCode(0).Text)
    iendPLU = (txtPLUCode(1).Text)
    arrayNewPLU = GetArrayPLU(istartPLU, iendPLU)
'    If (FormCall.flexPLU.Rows - 1) + UBound(arrayNewPLU) + 1 > iMaxPLU Then
'        MsgBox "abc" & iMaxPLU, vbExclamation
'        'MsgBox arrMessage(11) & iMaxPLU & arrMessage(12), vbExclamation, myClose, arrMessage(1)
'        Exit Sub
'    End If
    If StrComp(istartPLU, iendPLU, 1) = 0 Then
        AddNewOnePLU Arr
        Exit Sub
    End If
    
    Dim strPluCodeTemp As String
    
    strPluCodeTemp = txtPLUCode(0).Text
    iLen = Len(strPluCodeTemp)
    
    For i = 0 To UBound(arrayNewPLU)
    DoEvents
        If InStr(1, array_PLUCodes, ";" & arrayNewPLU(i) & ";", vbBinaryCompare) <> 0 Then
            MsgBox "M· hµng [" & arrayNewPLU(i) & "] ®· tån t¹i råi!", vbInformation
            Exit Sub
        End If
    Next i
    '---------------
    Dim iBegin As Integer
    Dim iEnd As Integer
    Dim iExprice As Integer
    With probar
        InitProgressBar
        For J = 0 To UBound(arrayNewPLU)
        DoEvents
            sPLU = arrayNewPLU(J)
            array_PLUCodes = array_PLUCodes & sPLU & ";" '16/03/06
            If J Mod 500 = 0 Then Delay 200
            AddDataToGrid sPLU, Arr, iBegin, iEnd, iExprice
'            probar.value = probar.Max - (iendPLU - j)
            Me.Caption = "Ma hang thu " & J
            If .Value = .Max Then .Value = 300
            .Value = .Value + 100
        Next J
        .Visible = False
    End With
'    SetColorFlexGrid FormCall.flexPLU, irow + 1, 1, FormCall.flexPLU.Cols
    MsgBox DescArr(12), vbInformation
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- AddNewMultiPLU"
End Sub

Private Function AddDefaultValueForArr(ByVal iPos_GA As Integer, ByVal fSaveAs As Boolean)
On Error GoTo errHdl
    Dim res As New ADODB.Recordset
    Dim S1() As String
    Dim i As Integer
    Dim strSQL As String
    Dim intCount As Integer
       
    Set res = Open_Table(cnData, "Inventory")
    If res.State = 0 Then Exit Function
    With res
        ReDim Preserve S1(.Fields.count - 4)
        
        If fSaveAs Then
            With FormCall.flexPLU
                If .TextMatrix(1, 0) = "" Then
                    S1 = DataInFlex(FormCall.flexPLU, True)
                Else
                    For i = 1 To UBound(S1) ' - 1
                        S1(i) = .TextMatrix(.Row, i)
                    Next i
                End If
            End With
        Else
          S1 = DataInFlex(FormCall.flexPLU, True)
         
        End If
    End With
    CloseRecordset res
    AddDefaultValueForArr = S1
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- AddDefaultValueForArr"
End Function

Private Sub AddDataToGrid(ByVal sPLU As String, sTemp() As String, _
                          ByVal iBegin As Integer, iEnd As Integer, iExprice As Integer)
On Error GoTo errHdl
    Dim irow As Integer
    str_NewPLUs = str_NewPLUs & sPLU & ";"
    With FormCall.flexPLU
        If .TextMatrix(1, 0) <> "" Then
              .Rows = .Rows + 1
              irow = .Rows - 1
        Else: irow = 1
        End If
        .TextMatrix(irow, 0) = sPLU
        Dim iTemp As Integer
        
        iTemp = UBound(sTemp)
        If chkCreatePLU.Value = False Then
            iTemp = UBound(sTemp) - 1
        End If
        For i = 1 To iTemp 'UBound(sTemp) - 1
        DoEvents
            Select Case i
                Case 1
                        If chkCreatePLU.Value = False Then
                            .TextMatrix(irow, i) = sPLU
                        Else
                            If sTemp(i) = "PLU-NAME" Then
                                .TextMatrix(irow, i) = sTemp(i) & " " & sPLU
                            Else
                                .TextMatrix(irow, i) = sTemp(i)
                            End If
                        End If
                Case iBegin To iEnd
                    .TextMatrix(irow, i) = Format( _
                        Val(sTemp(i)), formatNum)
                    
                Case iExprice
                    .TextMatrix(irow, i) = sTemp(i)
                           
                Case Else: .TextMatrix(irow, i) = sTemp(i)
            End Select
        Next i
        .Refresh
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- AddDataToGrid"
End Sub

Private Function SetTextTemp()
On Error GoTo errHdl
    Dim sTemp() As String
    
    If chkCreatePLU = False Then
         sTemp = AddDefaultValueForArr(5, False)
      
    Else
        sTemp = AddDefaultValueForArr(0, True) 'gtri duoc tao theo mahang dc chon t/ung
    End If
    SetTextTemp = sTemp
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- SetTextTemp"
End Function

Private Sub Timer1_Timer()
On Error GoTo errHdl

    If Timer - sTime > 3 Then
        txtHint.Visible = False
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Timer1_Timer"
End Sub

Private Sub txtPlucode_DblClick(Index As Integer)
On Error GoTo Handle
    With frmKeyboard
        .txtInput.PasswordChar = ""
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtPLUCode(Index).Text = .Let_Text_Input
       
    End With
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  txtPlucode_DblClick "
End Sub

'           ----------- TEXTBOX -----------
Private Sub txtPluCode_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo errHdl

    Select Case KeyAscii
        Case 13
            Select Case Index
                Case 0
                        If optCreatePLU(1).Value = True Then
                            With txtPLUCode(1)
                                .SetFocus
                                .SelStart = 0
                                .SelLength = 9999
                            End With
                        Else
                            chkCreatePLU.SetFocus
                        End If
                Case 1
                        chkCreatePLU.SetFocus
            End Select
        Case Is < 32, 48 To 57, 44, 46
        Case Else: KeyAscii = 0
    End Select
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Timer1_Timer"
End Sub
'           ---------- OTHER FUNCTIONS ----------
Private Sub HideControl(fHide As Boolean)
On Error GoTo errHdl

    cmdHelp.Enabled = Not fHide
    cmdCancel.Enabled = Not fHide
    cmdCreate.Enabled = Not fHide
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- HideControl"
End Sub

Private Sub InitProgressBar()
On Error GoTo errHdl
    With probar
        .Visible = True
        .Value = 0
        .Min = 0
'        .Max = iEnd - iBegin
        .Max = 1000
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- InitProgressBar"
End Sub

Public Function Get_AddNewRecords()
On Error GoTo errHdl

    Get_AddNewRecords = str_NewPLUs
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- Get_AddNewRecords"
End Function
'--------New-------
Private Function DataInFlex(flex As MSFlexGrid, fCopy As Boolean)
On Error GoTo errHdl

    Dim sTemp As String
    Dim sResult() As String
        
    With flex
        res.MoveFirst
        ReDim Preserve sResult(.Cols)
        For i = 0 To .Cols - 1
        DoEvents
            Select Case i
                Case 1
                        If fCopy Then
                            sResult(i) = "PLU-NAME"
                        End If
                        sTemp = ""
                Case 2: sTemp = "Dept_ID" 'GroupA
                        sResult(i) = "01"
                        sTemp = ""
                'Price
                Case 3, 4, 5: sTemp = "Std_Price" & (i - 2)
                Case 6, 7, 8: sTemp = "HH_Price" & (i - 5)
                Case 9, 10, 11: sTemp = "EV_Price" & (i - 8)
                Case 12: sTemp = "LimitPrice"
                        sResult(i) = "00"
                Case 13: sTemp = "Unit"
                        sResult(i) = "C¸i"
                        'sResult(i) = "0"
                        sTemp = ""
                Case 14: sTemp = "MinStock"
                        sTemp = ""
                Case 15: sTemp = "Modify_Number"
                        sTemp = ""
                Case 16 To 20: sTemp = "F" & (i - 15)
                            sResult(i) = "10"
                Case 21: sTemp = "Picture"
                        sResult(i) = "a"
                        sTemp = ""
                
            End Select
            If sTemp <> "" Then
                sResult(i) = FillZeroForString("0", res.Fields(sTemp).DefinedSize)
            End If
        Next i
    End With
    DataInFlex = sResult
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- DataInFlex"
End Function


'---------- ATHU
Private Function GetArrayPLU(ByVal istartPLU As String, ByVal iendPLU As String) As String()
On Error GoTo errHdl

    Dim fHPLUCode As Double
    Dim eHPLUCode As Double
    Dim fLPLUCode As Double
    Dim eLPLUCode As Double
    Dim d1 As Double
    Dim d2 As Double
    Dim arrPLU() As String
    Dim iCount As Double
    Dim iLen As Integer
    
    iLen = res.Fields("ItemNum").DefinedSize
    fHPLUCode = Left(Right("00000000000000000000" & istartPLU, iLen), iLen / 2)
    eHPLUCode = Left(Right("00000000000000000000" & iendPLU, iLen), iLen / 2)
    iCount = -1
    For d1 = fHPLUCode To eHPLUCode Step 1
        DoEvents
        If d1 = fHPLUCode Then
            fLPLUCode = CDbl(Right("000000000" & istartPLU, iLen / 2))
        Else
            fLPLUCode = 1
        End If
        If d1 = eHPLUCode Then
            eLPLUCode = CDbl(Right("000000000" & iendPLU, iLen / 2))
        Else
            eLPLUCode = CDbl(FillZeroForString("1", (iLen / 2) + 1)) ' 1000000000
        End If
        
        For d2 = fLPLUCode To eLPLUCode Step 1
            DoEvents
            If d2 = CDbl(FillZeroForString("1", (iLen / 2) + 1)) Then
                iCount = iCount + 1
                ReDim Preserve arrPLU(iCount)
                arrPLU(iCount) = Right("000000000" & d1 + 1, iLen / 2) & "000000000"
            Else
                iCount = iCount + 1
                ReDim Preserve arrPLU(iCount)
                arrPLU(iCount) = Right("0000000000" & d1, 6) & Right("000000000" & d2, 6)
            End If
        Next d2
    Next d1
    GetArrayPLU = arrPLU
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.Name & "- GetArrayPLU"
End Function

Public Function Get_Max_PLUCODE() As String
On Error GoTo Handle
    Dim rsmax As New ADODB.Recordset
    Set rsmax = OpenCriticalTable("select max(ItemNum) as Max_PLU from Inventory", cnData)
    If Not rsmax.EOF Then
        If "" & rsmax.Fields("Max_PLU") = "" Then
            Get_Max_PLUCODE = "1"
        Else
            Get_Max_PLUCODE = rsmax.Fields("Max_PLU") + 1
        End If
    Else
        Get_Max_PLUCODE = rsmax.Fields("Max_PLU") + 1
    End If
Exit Function
Handle:
MsgBox Err.Number & Err.Description & "  mdlGlobal " & "   Get_Max_PLUCODE"

End Function
