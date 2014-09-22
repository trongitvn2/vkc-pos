VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDeletePLU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete PLU"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8085
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
   ScaleHeight     =   3585
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2130
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2295
      Width           =   4245
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
      Left            =   2130
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1545
      Width           =   4245
   End
   Begin VB.Frame fraCreatePLU 
      Caption         =   "Xãa thùc ®¬n"
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
      Height          =   1110
      Left            =   0
      TabIndex        =   5
      Tag             =   "L1"
      Top             =   150
      Width           =   5775
      Begin VB.OptionButton optDeletePLU 
         Caption         =   "1 mÆt hµng"
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
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Tag             =   "L2"
         Top             =   225
         Width           =   2580
      End
      Begin VB.OptionButton optDeletePLU 
         Caption         =   "&D·y c¸c mÆt hµng"
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
         TabIndex        =   1
         Tag             =   "L3"
         Top             =   660
         Width           =   2985
      End
   End
   Begin MSComctlLib.ProgressBar probar 
      Height          =   435
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   767
      _Version        =   393216
      Appearance      =   1
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   855
      Left            =   6480
      TabIndex        =   8
      Tag             =   "L7"
      Top             =   1380
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "&Tho¸t"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmDeletePLU.frx":0000
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
      Left            =   6480
      TabIndex        =   9
      Tag             =   "L6"
      Top             =   390
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "&Xãa"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmDeletePLU.frx":001C
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
      Caption         =   "M· hµng b¾t ®Çu:"
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
      Height          =   255
      Left            =   180
      TabIndex        =   7
      Tag             =   "L4"
      Top             =   1665
      Width           =   1935
   End
   Begin VB.Label lblEndPLU 
      Alignment       =   1  'Right Justify
      Caption         =   "M· hµng kÕt thóc:"
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
      Height          =   375
      Left            =   180
      TabIndex        =   6
      Tag             =   "L5"
      Top             =   2250
      Width           =   1935
   End
End
Attribute VB_Name = "frmDeletePLU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim res As New ADODB.Recordset
    Dim arrDelete() As String
    Dim fcheckplucode As Byte
    Public FormCall As Object
    Dim strPath As String 'Tam Them vao ngay 26/12/2006
    Dim flag As Byte 'Tam then ngay 06/01/2006
'           ---------- FORM ---------
Private Sub Form_Activate()
On Error GoTo errHdl

    Dim DescArr() As String
    Dim ctrl As Control
    
    DescArr = LoadLanguage(LngFile, "#02:012:")
    If cmdDelete.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    optDeletePLU(0).SetFocus
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl
    Set res = Open_Table(cnData, "Inventory")
    If res.State = 0 Then Exit Sub
    ReDim Preserve arrDelete(0)
'    With Me
'        .Height = 2625
'        .Width = 7695
'        .WindowState = 0
'    End With
    EndPLU_status False
    optDeletePLU(0).Value = True
    probar.Visible = False
    InitCombo
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errHdl

    CloseRecordset res
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Unload"
End Sub

Private Sub SetComboPLU(cbo As ComboBox)
On Error GoTo errHdl
Dim i As Integer
    cbo.Clear
    With FormCall.flexPLU
        If .TextMatrix(1, 0) = "" Then Exit Sub
        For i = 1 To .Rows - 1
        DoEvents
            cbo.AddItem .TextMatrix(i, 0) & "   " & .TextMatrix(i, 1)
        Next i
        cbo.ListIndex = .Row - 1
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- SetComboPLU"
End Sub
'           ----------- COMMANDBUTTON --------
Private Sub cmdCancel_Click()
On Error GoTo errHdl

    CloseRecordset res
    Unload Me
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- cmdCancel_Click"
End Sub

Private Sub cmddelete_Click()
On Error GoTo errHdl

    Dim istartIndex As Integer
    Dim iendIndex As Integer
    
    istartIndex = cboPLUCode(0).ListIndex
    iendIndex = cboPLUCode(1).ListIndex
    HideControl True
    If optDeletePLU(0).Value = True Or _
      (optDeletePLU(1).Value = True And istartIndex = iendIndex) Then
        
        DeleteOnePLU
        GoTo 1
    End If
    If optDeletePLU(1).Value = True Then
        DeleteMultiPLU
    End If
1:
    InitCombo
    HideControl False
    If FormCall.flexPLU.TextMatrix(1, 0) = "" Then _
        cmdCancel_Click
        
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- cmdDelete_Click"
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
    & Me.name & "- cboPLUCode_KeyPress"
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
    & Me.name & "- optDeletePLU_Click"
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
    & Me.name & "- optDeletePLU_KeyPress"
End Sub
'           ---------- FUNCTIONS DELETE --------
Private Sub EndPLU_status(flag As Boolean) 'an hoac hien lblEndPlu & txtEndPlu
On Error GoTo errHdl

    If flag Then
        lblEndPLU.Visible = True
        cboPLUCode(1).Visible = True
        lblBeginPLU.top = 1500
        cboPLUCode(0).top = 1500
    Else
        lblBeginPLU.top = lblEndPLU.top
        cboPLUCode(0).top = cboPLUCode(1).top
        lblEndPLU.Visible = False
        cboPLUCode(1).Visible = False
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- EndPLU_status"
End Sub

Private Function CheckPLUCode() As Boolean
On Error GoTo errHdl

    Dim istartIndex As Integer
    Dim iendIndex As Integer
    
    istartIndex = cboPLUCode(0).ListIndex
    fcheckplucode = 0
        
    If optDeletePLU(1).Value Then
        iendIndex = cboPLUCode(1).ListIndex
        If istartIndex > iendIndex Then
            fcheckplucode = 2
            GoTo 1
        End If
    End If
1:     If fcheckplucode <> 0 Then
            CheckPLUCode = False
            Exit Function
        Else
            CheckPLUCode = True
        End If
        
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- CheckPLUCode"
End Function

Private Sub DeleteOnePLU()
On Error GoTo errHdl

    Dim sPLU As String
    Dim i As Double
    Dim irow As Integer
    'Tam Them ngay 08/01/2007
    Dim strPluCodeNew As String
    Dim blnRet As Boolean
    Dim rsInvoice_Itemized As New ADODB.Recordset
    Set rsInvoice_Itemized = Open_Table(cnData, "Invoice_Itemized")
    
    sPLU = Left(cboPLUCode(0).Text, InStr(cboPLUCode(0).Text, "   ") - 1)

    With rsInvoice_Itemized
        .Find "ItemNum='" & sPLU & "'", , adSearchForward, adBookmarkFirst
        If .EOF Then
            If MsgBox("B¹n cã ch¾c ch¾n muèn xãa mÆt hµng nµy ?", vbOKCancel) = vbCancel Then Exit Sub
            With FormCall.flexPLU
                For i = 1 To .Rows - 1
                DoEvents
                    If InStr(1, sPLU, .TextMatrix(i, 0), 1) <> 0 Then
                        If blnRet = True Then
                            .TextMatrix(i, 0) = strPluCodeNew
                            
                            If .Rows = 2 Then
                                Delete_Last_Row
                            Else
                                .RemoveItem i
                            End If
        '                    SetColorFlexGrid FormCall.flexPLU, 1, 1, .Cols
                            'end h
                        
                                
                        End If
                                             
                        ReDim Preserve arrDelete(UBound(arrDelete) + 1)
                        arrDelete(UBound(arrDelete)) = .TextMatrix(i, 0)
                        If .Rows = 2 Then
                            Delete_Last_Row
                            irow = 1
                        Else
                            .RemoveItem i
                            
                            irow = i - 2 '.Row
                            
                        End If
                        Exit For
                    End If
                Next i
        '        SetColorFlexGrid FormCall.flexPLU, 1, 1, .Cols
            End With
        Else
        MsgBox "M· hµng nµy ®ang sö dông", vbInformation
        End If
    End With
            
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- DeleteOnePLU"
End Sub

Private Sub DeleteMultiPLU()
On Error GoTo errHdl

    Dim strDelete As String
    Dim iNumDelete As Double
    Dim iCountDelete As Double
    Dim i As Double
    Dim iLen As Integer
    Dim iPaintRow As Integer
    Dim blnRet As Boolean
    
    blnRet = True
    
    strDelete = ""
    iNumDelete = 0: iCountDelete = 0: iPaintRow = 0
    iLen = 12 'res.Fields("ItemNum").DefinedSize
    If MsgBox("B¹n cã ch¾c ch¾n muèn xãa kh«ng ?", vbOKCancel) = vbCancel Then Exit Sub
    For i = cboPLUCode(0).ListIndex To cboPLUCode(1).ListIndex Step 1
    DoEvents
        iNumDelete = iNumDelete + 1
        strDelete = strDelete & Left(cboPLUCode(0).List(i), InStr(cboPLUCode(0).List(i), "   ") - 1) & ";"
    Next i
    InitProgressBar
    With FormCall.flexPLU
        For i = 1 To .Rows - 1
        DoEvents
            If iNumDelete = iCountDelete Then Exit For
            If InStr(1, strDelete, .TextMatrix(i, 0) & ";", 1) <> 0 Then
                .TextMatrix(i, 0) = strDelete
                If .Rows = 2 Then
                    Delete_Last_Row
                Else
                    .RemoveItem i
'                    i = i - 1
                End If
                iCountDelete = iCountDelete + 1
                If iPaintRow = 0 Then iPaintRow = i
                iCountDelete = iCountDelete + 1
                ReDim Preserve arrDelete(UBound(arrDelete) + 1)
                arrDelete(UBound(arrDelete)) = .TextMatrix(i, 0)
                If .Rows = 2 Then
                    Delete_Last_Row
                Else
                    .RemoveItem i 'iCountDelete 'i
                    i = i - 1
                End If
            End If
2:          If probar.Value = probar.Max Then
                probar.Value = 300
            End If
            probar.Value = probar.Value + 100
        Next i
        probar.Visible = False
'        SetColorFlexGrid FormCall.flexPLU, iPaintRow, 1, .Cols
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- DeleteMultiPLU"
End Sub
'               --------- OTHER FUNCTIONS ------
Private Sub HideControl(fHide As Boolean)
On Error GoTo errHdl

    cmdDelete.Enabled = Not fHide
    cmdCancel.Enabled = Not fHide
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - HideControl"
End Sub

Private Sub InitProgressBar()
On Error GoTo errHdl
    With probar
        .Visible = True
        .Min = 0
        .Max = 1000
'        .Max = iendPLU - istartPLU
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - InitProgressBar"
End Sub

Public Function Get_DeleteRecords()
On Error GoTo errHdl

    Get_DeleteRecords = arrDelete
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Get_DeleteRecords"
End Function

Private Sub InitCombo()
On Error GoTo errHdl

    SetComboPLU cboPLUCode(0)
    SetComboPLU cboPLUCode(1)
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - InitCombo"
End Sub

Private Sub Delete_Last_Row()
On Error GoTo errHdl

    Dim k As Byte
    
    With FormCall.flexPLU
        For k = 0 To .Cols - 1
        DoEvents
            .TextMatrix(1, k) = ""
        Next k
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Delete_Last_Row"
End Sub
