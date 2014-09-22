VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPaymentTimer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timer"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
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
   Icon            =   "frmPaymentTimer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdOK 
      Height          =   765
      Left            =   1800
      TabIndex        =   1
      Tag             =   "L2"
      Top             =   4980
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1349
      BTYPE           =   14
      TX              =   "&OK"
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
      BCOL            =   8438015
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPaymentTimer.frx":000C
      PICN            =   "frmPaymentTimer.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8281
      _Version        =   393216
      TabOrientation  =   2
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Payment Time"
      TabPicture(0)   =   "frmPaymentTimer.frx":0662
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FraPaymentTime"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cleaning Time"
      TabPicture(1)   =   "frmPaymentTimer.frx":067E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4575
         Left            =   -74580
         TabIndex        =   4
         Top             =   30
         Width           =   6105
         Begin VB.TextBox txtValueClean 
            Height          =   345
            Left            =   2400
            TabIndex        =   7
            Top             =   150
            Width           =   765
         End
         Begin VB.ListBox lstCleaningTime 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2460
            ItemData        =   "frmPaymentTimer.frx":069A
            Left            =   120
            List            =   "frmPaymentTimer.frx":069C
            Style           =   1  'Checkbox
            TabIndex        =   5
            Top             =   540
            Width           =   5895
         End
         Begin VB.Label lblDescription 
            Caption         =   $"frmPaymentTimer.frx":069E
            Height          =   1425
            Left            =   90
            TabIndex        =   9
            Top             =   3030
            Width           =   5955
         End
      End
      Begin VB.Frame FraPaymentTime 
         Height          =   4575
         Left            =   420
         TabIndex        =   3
         Top             =   30
         Width           =   6045
         Begin VB.TextBox txtValuePayment 
            Height          =   345
            Left            =   2310
            TabIndex        =   8
            Top             =   150
            Width           =   765
         End
         Begin VB.ListBox lstPaymentTime 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2460
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   6
            Top             =   510
            Width           =   5775
         End
         Begin VB.Label lblDecriptionPay 
            Caption         =   $"frmPaymentTimer.frx":07DC
            Height          =   1485
            Left            =   30
            TabIndex        =   10
            Top             =   3060
            Width           =   5925
         End
      End
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   765
      Left            =   3420
      TabIndex        =   2
      Tag             =   "L3"
      Top             =   4980
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1349
      BTYPE           =   14
      TX              =   "&Cancel"
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
      BCOL            =   8438015
      BCOLO           =   16777152
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPaymentTimer.frx":0913
      PICN            =   "frmPaymentTimer.frx":092F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
End
Attribute VB_Name = "frmPaymentTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstimer As New ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Call UpdateCleaningTime
    Call UpdatePaymentTime
    Unload Me
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    Dim ctrl As Control
    Dim DescArr() As String
    If cmdOK.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
        DescArr = LoadLanguage(LngFile, "#03:020:")
        Me.Caption = DescArr(1)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
        Next ctrl
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"

End Sub

Private Sub Form_Load()
Dim i As Integer
On Error GoTo Handle
'If cnData.State = 0 Then
'    Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'End If
If cnData.State <> 0 Then
    Set rstimer = Open_Table(cnData, "colorTablePlan")
End If
    With lstPaymentTime
        For i = 1 To 8
            .AddItem "Payment Time - " & i * 2 & " Min"
        Next i
    End With
    With lstCleaningTime
        For i = 1 To 8
            .AddItem "Cleaning Time - " & i * 2 & " Min"
        Next i
    End With
    With rstimer
        .Find "ReserveType='" & "PAID" & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtValuePayment.Text = .Fields("ValueTime")
        End If
    End With
    With rstimer
        .Find "ReserveType='" & "CLEANING" & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtValueClean.Text = .Fields("ValueTime")
        End If
    End With
    Call AddValueForList(txtValueClean.Text, lstCleaningTime)
    Call AddValueForList(txtValuePayment.Text, lstPaymentTime)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rstimer = Nothing
End Sub

Private Sub lstCleaningTime_Click()
Dim i As Integer
Dim str As String
    On Error GoTo Handle
    For i = 0 To lstCleaningTime.ListCount - 1
        If i <> lstCleaningTime.ListIndex Then
            If lstCleaningTime.Selected(i) Then lstCleaningTime.Selected(i) = False
        End If
    Next i
    For i = 0 To lstCleaningTime.ListCount - 1
        If lstCleaningTime.Selected(i) Then
            str = str & "1"
        Else: str = str & "0"
        End If
    Next
            txtValueClean.Text = FillZeroForString(BinToHex(str), 2)
    Exit Sub
Handle:
        MsgBox Err.Number & Err.Description & Me.name & "lstCleaningTime_Click"
End Sub

Private Sub lstPaymentTime_Click()
Dim i As Integer
Dim str As String
    On Error GoTo Handle
    For i = 0 To lstPaymentTime.ListCount - 1
        If i <> lstPaymentTime.ListIndex Then
            If lstPaymentTime.Selected(i) Then lstPaymentTime.Selected(i) = False
        End If
    Next i
    For i = 0 To lstPaymentTime.ListCount - 1
        If lstPaymentTime.Selected(i) Then
            str = str & "1"
        Else: str = str & "0"
        End If
    Next
            txtValuePayment.Text = FillZeroForString(BinToHex(str), 2)
    Exit Sub
Handle:
        MsgBox Err.Number & Err.Description & Me.name & "lstPaymentTime_Click"
End Sub

Public Sub UpdatePaymentTime()
On Error GoTo Handle
    With rstimer
        .Find "ReserveType='" & "PAID" & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("ValueTime") = txtValuePayment.Text
            .Update
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  UpdatePaymentTime"
End Sub

Public Sub UpdateCleaningTime()
On Error GoTo Handle
    With rstimer
        .Find "ReserveType='" & "CLEANING" & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("ValueTime") = txtValueClean.Text
            .Update
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  UpdateCleaningTime"
End Sub

Public Sub AddValueForList(ByVal str1 As String, ByVal lst As ListBox)
On Error GoTo errHdl

    Dim strBin As String
    Dim k As Integer
    
    strBin = HexToBin(str1)
    strBin = FillZeroForString(strBin, 8)
    For k = 0 To Len(strBin) - 1 Step 1
    DoEvents
        If Mid(strBin, k + 1, 1) = 1 Then
            lst.Selected(k) = True
        Else
            lst.Selected(k) = False
        End If
    Next k
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlFunctionPublic - AddValueForList"
End Sub

