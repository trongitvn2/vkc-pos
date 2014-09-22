VERSION 5.00
Begin VB.Form frmLicense 
   Caption         =   "License"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
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
   Icon            =   "frmLicense.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdexit 
      Height          =   735
      Left            =   5040
      TabIndex        =   11
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLicense.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdOK 
      Height          =   735
      Left            =   1920
      TabIndex        =   10
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "&Më khãa"
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
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLicense.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.TextBox txtKey5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      MaxLength       =   5
      TabIndex        =   9
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtKey4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      MaxLength       =   5
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtKey3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      MaxLength       =   5
      TabIndex        =   7
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtKey2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   5
      TabIndex        =   6
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtKey1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaxLength       =   5
      TabIndex        =   5
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox txtSystemCode 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   5895
   End
   Begin VB.TextBox txtuserCode 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   4215
   End
   Begin prjTouchScreen.MyButton cmdHelp 
      Height          =   735
      Left            =   3480
      TabIndex        =   12
      Top             =   3240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLicense.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Vui lßng nhËp m· sö dông vµo 5 « bªn d­íi sau ®ã bÊm ""Më khãa"" ®Ó ®­îc sö dông phÇn mÒm b¶n quyÒn"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   8295
   End
   Begin VB.Label lblLicense 
      Alignment       =   1  'Right Justify
      Caption         =   "M· sö dông:"
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
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label lblSystemCode 
      Alignment       =   1  'Right Justify
      Caption         =   "M· m¸y:"
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
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lbluser 
      Alignment       =   1  'Right Justify
      Caption         =   "M· ng­êi dïng:"
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
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
End
Attribute VB_Name = "frmLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim User_Code, System_Code As String
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    frmLicen_Help.Show vbModal
End Sub

Private Sub cmdOK_Click()
Dim countLogin As Integer
Dim sCode, S, Check_ID As String
Dim fTemp As Integer
Dim Path_Direction, Validate As String
Dim S1, S2, s3, s4, s5, s6, s7, s8, s9, s10 As String
    
Path_Direction = "C:\Windows\System32"

If Dir(Path_Direction & "\KernelSys.sys", vbDirectory) <> "" Then
    Kill Path_Direction & "\KernelSys.sys"
    fTemp = FreeFile
    Open Path_Direction & "\KernelSys.sys" For Input As #fTemp
        DoEvents
        Line Input #fTemp, S
    Close #fTem
End If

    Check_ID = UCase(ProcessID) '& Mac_ID
    
    sCode = txtuserCode.Text & txtSystemCode.Text
    
    S1 = Mid(sCode, 36, 1) & Mid(sCode, 8, 1) & Mid(sCode, 25, 1) & Mid(sCode, 6, 1) & Mid(sCode, 23, 1)
    S2 = Mid(sCode, 7, 1) & Right(sCode, 1) & Mid(sCode, 13, 1) & Mid(sCode, 36, 1) & Mid(sCode, 25, 1)
    s3 = Mid(sCode, 3, 1) & Mid(sCode, 15, 1) & Mid(sCode, 11, 1) & Mid(sCode, 39, 1) & Mid(sCode, 24, 1)
    s4 = Mid(sCode, 16, 1) & Mid(sCode, 2, 1) & Mid(sCode, 9, 1) & Mid(sCode, 5, 1) & Mid(sCode, 26, 1)
    s5 = Mid(sCode, 14, 1) & Mid(sCode, 2, 1) & Mid(sCode, 1, 1) & Mid(sCode, 32, 1) & Mid(sCode, 13, 1)
    
   s6 = Mid(sCode, 5, 1) & Mid(sCode, 2, 1) & Mid(sCode, 20, 1) & Mid(sCode, 15, 1) & Mid(sCode, 26, 1)
    s7 = Mid(sCode, 7, 1) & Right(sCode, 1) & Mid(sCode, 19, 1) & Mid(sCode, 12, 1) & Mid(sCode, 3, 1)
    s8 = Mid(sCode, 18, 1) & Mid(sCode, 24, 1) & Mid(sCode, 11, 1) & Mid(sCode, 17, 1) & Mid(sCode, 5, 1)
   s9 = Mid(sCode, 21, 1) & Mid(sCode, 16, 1) & Mid(sCode, 9, 1) & Mid(sCode, 24, 1) & Mid(sCode, 18, 1)
    s10 = Mid(sCode, 26, 1) & Mid(sCode, 33, 1) & Mid(sCode, 13, 1) & Mid(sCode, 8, 1) & Mid(sCode, 10, 1)

S = txtKey1.Text & txtKey2.Text & txtKey3.Text & txtKey4.Text & txtKey5.Text

    If S = S1 & S2 & s3 & s4 & s5 Then
        Validate = En_Decryption.MalgoEncrypt(Check_ID, 15)
        fTemp = FreeFile
        Open Path_Direction & "\KernelSys.sys" For Output As #fTemp
            Print #fTemp, Validate
        Close #fTemp
    ElseIf S = s6 & s7 & s8 & s9 & s10 Then
        'Validate = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date) + 30, "00")
        Validate = Get_LifeTime
        If Validate = "" Then
           Validate = gfCONVERT_DATE_TO_STRING(Date + 365)
        End If
        Validate = En_Decryption.MalgoEncrypt(Validate, 5)
        fTemp = FreeFile
        Open Path_Direction & "\systrial.dll" For Output As #fTemp
            Print #fTemp, Validate
        Close #fTemp
        Call Update_Lifetime(Validate)
    Else
        MsgBox "B¹n nhËp sai License ! Vui lßng liªn hÖ : 0918.655.887 ®Ó ®­îc cÊp m· sö dông !!!"
        Exit Sub
    End If
    Unload Me
    frmLogin.Show vbModal
End Sub

Private Sub Form_Activate()
On Error GoTo errHdl

    Dim DescArr() As String
    Dim ctrl As Control
    
    DescArr = LoadLanguage(LngFile, "#03:019:")
    If cmdOk.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Activate"
End Sub

Private Sub Form_Load()
    Call Load_user_ID
End Sub

Public Sub Load_user_ID()
On Error GoTo Handle
    
    User_Code = Let_UserCode
    System_Code = Let_SystemCode
    
    txtuserCode.Text = User_Code
    txtSystemCode.Text = System_Code
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Load_user_ID"
End Sub

Private Sub txtKey1_Change()
If Len(txtKey1.Text) = txtKey1.MaxLength Then txtKey2.SetFocus
End Sub

Private Sub txtKey2_Change()
If Len(txtKey2.Text) = txtKey2.MaxLength Then txtKey3.SetFocus

End Sub

Private Sub txtKey3_Change()
    If Len(txtKey3.Text) = txtKey3.MaxLength Then txtKey4.SetFocus
End Sub

Private Sub txtKey4_Change()
    If Len(txtKey4.Text) = txtKey4.MaxLength Then txtKey5.SetFocus

End Sub

Private Sub txtKey5_Change()
    If Len(txtKey5.Text) = txtKey5.MaxLength Then cmdOk.SetFocus

End Sub
'Lay Code used
Public Function GetKey(ByVal strCode As String, i As Integer) As String
On Error GoTo Handle
Dim strResult As String
GetKey = ""
Do While Len(strResult) < 5
    
    strResult = strResult & Mid(strCode, i, 1)
    i = i + 3
Loop
GetKey = strResult
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  GetKey"
End Function


Public Function Let_UserCode() As String
On Error GoTo Handle
    Dim sUser_Code As String
'    Dim strMix As String
'    strMix = ProcessID & Mac_ID
'    sUser_Code = Mid(strMix, 27, 1) & Mid(strMix, 11, 1) & Mid(strMix, 8, 1) & Mid(strMix, 1, 1) & _
'                Mid(strMix, 1, 1) & Mid(strMix, 23, 1) & Mid(strMix, 2, 1) & Mid(strMix, 14, 1) & _
'                Mid(strMix, 25, 1) & Mid(strMix, 9, 1) & Mid(strMix, 18, 1) & Mid(strMix, 22, 1) & _
'                Mid(strMix, 10, 1) & Mid(strMix, 20, 1) & Mid(strMix, 7, 1)
'
 sUser_Code = Left(ProcessID, 5) & Right(Mac_ID, 5) & Mid(ProcessID, 9, 3) & Left(Mac_ID, 2)
Let_UserCode = sUser_Code
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Let_UserCode"

End Function
Public Function Let_SystemCode() As String
On Error GoTo Handle
    Dim sSystem_Code As String
'    Dim strMix As String
'    strMix = ProcessID & Mac_ID
'    sSystem_Code = Mid(strMix, 1, 1) & Mid(strMix, 6, 1) & Mid(strMix, 3, 1) & Mid(strMix, 8, 1) & _
'                Mid(strMix, 10, 1) & Mid(strMix, 7, 1) & Mid(strMix, 14, 1) & Mid(strMix, 3, 1) & _
'                Mid(strMix, 13, 1) & Mid(strMix, 19, 1) & Mid(strMix, 10, 1) & Mid(strMix, 5, 1) & _
'                Mid(strMix, 20, 1) & Mid(strMix, 9, 1) & Mid(strMix, 16, 1) & Mid(strMix, 12, 1) & _
'                Mid(strMix, 23, 1) & Mid(strMix, 11, 1) & Mid(strMix, 9, 1) & Mid(strMix, 6, 1) & _
'                Mid(strMix, 22, 1) & Mid(strMix, 2, 1) & Mid(strMix, 8, 1) & Mid(strMix, 7, 1) & _
'                Mid(strMix, 4, 1) & Mid(strMix, 24, 1)
sSystem_Code = Mid(ProcessID, 3, 5) & Left(ProcessID, 2) & Right(ProcessID, 4) & Mid(ProcessID, 10, 3) & Mac_ID
Let_SystemCode = sSystem_Code
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Let_SystemCode"

End Function



Public Function Get_LifeTime() As String
On Error GoTo Handle
Dim rTime As String
Dim rsrTime As New ADODB.Recordset
If cnData.State = adStateClosed Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
Set rsrTime = Open_Table(cnData, "DateLock")
With rsrTime
    If .RecordCount > 0 Then
        rTime = En_Decryption.MalgoDecrypt(.Fields("LifeTime"), 5)
    Else
        rTime = ""
    End If
End With
Get_LifeTime = rTime
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & "Get_LifeTime"
End Function
Public Sub Update_Lifetime(S As String)
On Error GoTo Handle
Dim rsrTime As New ADODB.Recordset
If cnData.State = adStateClosed Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
Set rsrTime = Open_Table(cnData, "DateLock")
With rsrTime
    If .RecordCount = 0 Then
        .addNew
        .Fields("LifeTime") = S
        .Update
    Else
        .Fields("LifeTime") = S
        .Update
    End If
End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "Update_Lifetime"
End Sub
