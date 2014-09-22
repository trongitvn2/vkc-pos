VERSION 5.00
Begin VB.Form frmClerkAttenden 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ChÊm c«ng nh©n viªn"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
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
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   30
      Top             =   2190
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   705
      Left            =   150
      TabIndex        =   2
      Tag             =   "L4"
      Top             =   1710
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   1244
      BTYPE           =   4
      TX              =   "Tho¸t"
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
      BCOL            =   12640511
      BCOLO           =   16777088
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmClerkAttenden.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdLogout 
      Height          =   1335
      Left            =   2370
      TabIndex        =   1
      Tag             =   "L3"
      Top             =   270
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   2355
      BTYPE           =   4
      TX              =   "&Ra ca"
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
      MICON           =   "frmClerkAttenden.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdLogin 
      Height          =   1335
      Left            =   150
      TabIndex        =   0
      Tag             =   "L2"
      Top             =   270
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   2355
      BTYPE           =   4
      TX              =   "&Vµo ca"
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
      MICON           =   "frmClerkAttenden.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblDateTime 
      Alignment       =   2  'Center
      Caption         =   "DateTime"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   645
      Left            =   180
      TabIndex        =   3
      Top             =   2520
      Width           =   4305
   End
End
Attribute VB_Name = "frmClerkAttenden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DescArr() As String
Dim rsTime_Clock As ADODB.Recordset
Dim Clerk_ID As String
Dim DateLog As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLogin_Click()
On Error GoTo Handle
    With rsTime_Clock
        If CheckTimeOut(Clerk_ID, "I") = False Then
            .addNew
            .Fields("ID") = MaxIDTime
            .Fields("Cashier_ID") = Clerk_ID
            .Fields("StartDateTime") = DateLog & Format(Now, "HH:mm:ss")
            .Fields("EndDateTime") = ""
            .Fields("NumMinutes") = 0
            .Fields("Status") = "I"
            .Update
            .Requery
'            Call Print_LogIn(Clerk_ID)
        Else
            MsgBox " B¹n ch­a ra ca do ®ã kh«ng thÓ vµo ca!"
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdLogin_Click"
End Sub

Private Sub cmdLogout_Click()
On Error GoTo Handle
    With rsTime_Clock
        If CheckTimeOut(Clerk_ID, "O") = True Then
            Dim rslog As New ADODB.Recordset
            Set rslog = OpenCriticalTable("Select * from Time_Clock where Cashier_ID='" & Clerk_ID & "'", cnData)
            rslog.Find "Status= 'I'", , adSearchForward, adBookmarkFirst
            If Not rslog.EOF Then
                .Fields("EndDateTime") = DateLog & Format(Now, "HH:mm:ss")
                .Fields("NumMinutes") = 0
                .Fields("Status") = "O"
                .Update
                .Requery
            End If
'            Call Print_LogIn(Clerk_ID)
        Else
            MsgBox " B¹n ch­a vµo ca do ®ã kh«ng thÓ ra ca!"
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdLogin_Click"
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    If cmdCancel.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Activate"
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
        DescArr = LoadLanguage(LngFile, "#03:010:")
        lblDatetime.Caption = Date & "  --  " & Format(Now, "HH:mm:ss")
'        If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\database.mdb", "100881administrator")
        Set rsTime_Clock = Open_Table(cnData, "Time_Clock")
        DateLog = Format(Day(Date), "00") & "/" & Format(Month(Date), "00") & "/" & Format(Year(Date), "0000")
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Private Sub Timer1_Timer()
On Error GoTo Handle
        lblDatetime.Caption = Date & "  --  " & Format(Now, "HH:mm:ss")
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Timer1_Timer"
End Sub

Public Property Get Get_ClerkID() As Variant
    Get_ClerkID = Clerk_ID
End Property

Public Property Let Get_ClerkID(ByVal vNewValue As Variant)
    Clerk_ID = vNewValue
End Property

Public Function CheckTimeOut(CashierID As String, InOut As String) As Boolean
On Error GoTo Handle
Dim rsisLogOut As New ADODB.Recordset
Dim isLogOut As Boolean
    Set rsisLogOut = OpenCriticalTable("Select * from Time_Clock where Status='" & InOut & "'", cnData)
    With rsisLogOut
        .Find "Cashier_ID='" & CashierID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            isLogOut = True
        End If
    End With
    CheckTimeOut = isLogOut
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " CheckTimeOut"
End Function

'Public Sub Print_LogIn(Clerk_ID As String)
'On Error GoTo Handle
'    Dim cmd As New ADODB.Command
'    Dim SQL As String
'    SQL = "select Cashier_ID,StartDateTime,EndDateTime from Time_Clock where Cashier_ID='" & Clerk_ID & "' and ID=" & MaxID(Clerk_ID)
'    Set crClerkLogin = Nothing
'        cmd.ActiveConnection = cnData
'        cmd.CommandText = SQL
'        cmd.Execute
'    With crClerkLogin
'        .Database.AddADOCommand cnData, cmd
'        .txtclerkID.SetUnboundFieldSource "{ado.Cashier_ID}"
'        .txtLoginDate.SetUnboundFieldSource "{ado.StartDateTime}"
'        .txtLogOutDate.SetUnboundFieldSource "{ado.EndDateTime}"
'    End With
'    With frmShow_Report_80
'        .Report = crClerkLogin
'        .Show vbModal
'
'    End With
'    Unload frmShow_Report_80
'    Unload Me
'Exit Sub
'Handle:
'MsgBox Err.Number & Err.Description & Me.Name & " Print_LogIn"
'End Sub

Public Function MaxID(Clerk As String) As Double
On Error GoTo Handle
    Dim rsmax As New ADODB.Recordset
    Dim maxClerk As String
    Set rsmax = OpenCriticalTable("select Max(ID) as MaxID from Time_Clock where Cashier_ID='" & Clerk & "'", cnData)
    If rsmax.RecordCount > 0 Then
        maxClerk = rsmax.Fields("MaxID")
    End If
    
    If maxClerk = 0 Then
        MaxID = 1
    Else
        MaxID = maxClerk
    End If
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & "MaxID"

End Function
Public Function MaxIDTime() As Double
On Error GoTo Handle
    Dim rsmax As New ADODB.Recordset
    Dim MaxIDClock As Double
    Set rsmax = OpenCriticalTable("select Max(ID) as MaxID from Time_Clock ", cnData)
    If rsmax.RecordCount > 0 Then
        If CDbl("0" & rsmax.Fields("MaxID")) = 0 Then
            MaxIDClock = 1
        Else
            MaxIDClock = rsmax.Fields("MaxID") + 1
        End If
    End If
        MaxIDTime = MaxIDClock
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & "MaxID"

End Function
