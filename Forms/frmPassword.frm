VERSION 5.00
Begin VB.Form frmPassword 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
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
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Tag             =   "L3"
      Top             =   1350
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "&HÒy"
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
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPassword.frx":000C
      PICN            =   "frmPassword.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   150
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   4725
   End
   Begin prjTouchScreen.MyButton cmdKeyboard 
      Height          =   855
      Left            =   3360
      TabIndex        =   3
      Tag             =   "L4"
      Top             =   1350
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "&Bµn ph›m"
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
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPassword.frx":0662
      PICN            =   "frmPassword.frx":067E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdOK 
      Height          =   855
      Left            =   1440
      TabIndex        =   4
      Tag             =   "L5"
      Top             =   1350
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "&ßÂng ˝"
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
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPassword.frx":0AD0
      PICN            =   "frmPassword.frx":0AEC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NhÀp mÀt kh»u cÒa ng≠Íi qu∂n l˝"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   150
      TabIndex        =   1
      Tag             =   "L2"
      Top             =   150
      Width           =   4665
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Desarr() As String
Dim ActionForm As String
Dim strIDText As String
Dim isOK As Boolean
Private Sub cmdCancel_Click()
    isOK = False
    Unload Me
End Sub

Private Sub cmdKeyboard_Click()
    On Error GoTo Handle
        Select Case ActionForm
            Case "SetColor"
                With frmKeyboard
                    .FormCallkeyboard = "SetColor"
                    .txtInput.PasswordChar = "*"
                    .Show vbModal
                End With
            Case "EditTable"
                With frmKeyboard
                    .FormCallkeyboard = "EditTable"
                    .txtInput.PasswordChar = "*"
                    .Show vbModal
                End With
            Case "SystemFlag"
                With frmKeyboard
                    .FormCallkeyboard = "SystemFlag"
                    .txtInput.PasswordChar = "*"
                    .Show vbModal
                End With
            Case "SaleDelete"
                With frmKeyboard
                    .FormCallkeyboard = "DeleteSale"
                    .txtInput.PasswordChar = "*"
                    .Show vbModal
                End With
            Case "Select_Station"
                With frmKeyboard
                    .FormCallkeyboard = "Select_Station"
                    .Show vbModal
                End With
            Case "Employee"
                With frmKeyboard
                    .FormCallkeyboard = "Employee"
                    .Show vbModal
                End With
            Case "SaleReport"
                With frmKeyboard
                    .FormCallkeyboard = "SaleReport"
                    .txtInput.PasswordChar = "*"
                    .Show vbModal
                End With
            Case Else
                With frmKeyboard
                    .FormCallkeyboard = "Other"
                    .txtInput.PasswordChar = "*"
                    .Show vbModal
                    strIDText = .Let_Text_Input
                End With
        End Select
        Unload Me

    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "cmdKeyboard_Click"
End Sub

Private Sub cmdOK_Click()
    On Error GoTo Handle
        If txtPass.Text = "" Then
            MsgBox "Vui lﬂng nhÀp mÀt kh»u qu∂n trﬁ!", vbInformation
            Exit Sub
        End If
        isOK = True
        strIDText = txtPass.Text
        If UCase(txtPass.Text) = "131112" Then
            Unload Me
            Select Case ActionForm
                Case "SaleDelete"
                     frmDeleteSaleData.Show vbModal
                Case "Employee"
                     frmemployee.Show vbModal
                Case "SaleReport"
                    ' Set cnData = Get_Connection(BackupFolder & "\Database.mdb", "100881administrator")
                     With frmSetup
                         .Show vbModal
                     End With
                 
                Case "Select_Station"
                     frmSelect_Station.Show vbModal
                Case "EditTable"
                         frmEditTablePlan.Show vbModal
                Case "SetColor"
                     frmColorBox.Show vbModal
                Case "SystemFlag"
                     frmSystemFlag.Show vbModal
            End Select
           
        Else
        
            If UCase(Mid(txtPass.Text, 3, Len(txtPass.Text) - 2)) = UCase(UserPass) Or txtPass.Text = "131112" Then
                Unload Me
                Select Case ActionForm
                    Case "EditTable"
                        frmEditTablePlan.Show vbModal
                    Case "SetColor"
                        frmColorBox.Show vbModal
                    Case "SystemFlag"
                        frmSystemFlag.Show vbModal
                    
                End Select
    '        Else
    '            MsgBox "Sai mÀt kh»u!vui lﬂng nhÀp lπi", vbInformation
            End If
       
        End If
        Unload Me
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdOK_Click "
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    Dim ctrl As Control
    If cmdOK.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
        Desarr = LoadLanguage(LngFile, "#01:005:")
        Me.Caption = Desarr(1)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Desarr(Mid(ctrl.Tag, 2))
        Next ctrl
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Dim ctrl As Control
        Desarr = LoadLanguage(LngFile, "#01:005:")
        Me.Caption = Desarr(1)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Desarr(Mid(ctrl.Tag, 2))
        Next ctrl
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If isOK = False Then
        strIDText = ""
    End If
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    On Error GoTo Handle
        If KeyAscii = 13 Then
            'If UCase(txtPass.Text) = UCase(UserPass) Then
                Call cmdOK_Click
            'End If
        End If
        
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   txtPass_KeyPress"
End Sub


Public Property Let FormActionKey(ByVal vNewValue As Variant)
    ActionForm = vNewValue
End Property

Public Property Get return_Pass() As Variant
    return_Pass = strIDText
End Property

Public Property Get Return_right() As Variant
    Return_right = isOK
End Property
