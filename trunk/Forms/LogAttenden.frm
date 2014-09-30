VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmClerkLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "§¨ng nhËp chÊm c«ng"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
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
   ScaleHeight     =   3870
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   120
      Top             =   600
   End
   Begin prjTouchScreen.MyButton cmdIn 
      Height          =   855
      Left            =   6480
      TabIndex        =   7
      Top             =   960
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   1
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
      BCOL            =   16711680
      BCOLO           =   255
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "LogAttenden.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.TextBox txtPass 
      Height          =   555
      IMEMode         =   3  'DISABLE
      Left            =   1920
      TabIndex        =   3
      Top             =   2040
      Width           =   4365
   End
   Begin VB.TextBox txtID 
      Height          =   555
      IMEMode         =   3  'DISABLE
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   4365
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   1005
      Left            =   2040
      TabIndex        =   4
      Tag             =   "L4"
      Top             =   2670
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   1773
      BTYPE           =   1
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "LogAttenden.frx":001C
      PICN            =   "LogAttenden.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdKeyboard 
      Height          =   1005
      Left            =   4290
      TabIndex        =   5
      Tag             =   "L5"
      Top             =   2670
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   1773
      BTYPE           =   1
      TX              =   "&Bµn phÝm"
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
      BCOL            =   14737632
      BCOLO           =   16711680
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "LogAttenden.frx":62D2
      PICN            =   "LogAttenden.frx":62EE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdOut 
      Height          =   855
      Left            =   6480
      TabIndex        =   8
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      BTYPE           =   1
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
      BCOL            =   33023
      BCOLO           =   33023
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "LogAttenden.frx":6740
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpTime 
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   240
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm:ss"
      Format          =   64225282
      UpDown          =   -1  'True
      CurrentDate     =   38462.5826388889
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "chÊm c«ng nh©n viªn"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label lblPass 
      Alignment       =   1  'Right Justify
      Caption         =   "Tªn Nh©n viªn:"
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   0
      TabIndex        =   2
      Tag             =   "L3"
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label lblCashierID 
      Alignment       =   1  'Right Justify
      Caption         =   "M· sè NV:"
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Width           =   1575
   End
End
Attribute VB_Name = "frmClerkLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DescArr() As String
Dim rsEmployee As New ADODB.Recordset
Dim rsInOut As New ADODB.Recordset
Dim rsInOutPrint As New ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdIn_Click()
On Error GoTo Handle
If txtID.Text = "" Then Exit Sub
    With rsInOut
        .addNew
        .Fields("Cash_ID") = txtID.Text
        .Fields("DateTime") = Format(Now, "HH:mm:ss") & DateDefault
        .Fields("TimeType") = "I"
        .Fields("InOutRight") = True
        .Update
        .Requery
    End With
    'Xoa du lieu trong ban in phieu
    cnData.Execute "Delete  from InOutAttendent"
    'Ghi nhan du lieu moi xuong bang cham cong
    With rsInOutPrint
        .addNew
        .Fields("Cash_ID") = txtID.Text
        With rsEmployee
            .Find "Cashier_ID='" & txtID.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                rsInOutPrint.Fields("Cash_Name") = .Fields("EmpName")
            End If
        End With
        .Fields("DateInOut") = Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000")
        .Fields("TimeInOut") = Format(Now, "HH:mm:ss")
        .Fields("InOutRight") = "I"
        .Update
        .Requery
    End With
Call Print_LogIn(Trim(txtID.Text))
txtID.Text = ""
txtpass.Text = ""
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - cmdIn_Click"
End Sub

Private Sub cmdKeyboard_Click()
On Error GoTo Handle
    With frmKeyboard
        .txtInput.PasswordChar = ""
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtID.Text = .Let_Text_Input
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & " - cmdKeyboard_Click"
End Sub

Private Sub cmdOut_Click()
On Error GoTo Handle
If txtID.Text = "" Then Exit Sub
Dim rsCheckInOut As New ADODB.Recordset
    Set rsCheckInOut = OpenCriticalTable("Select Cash_ID, InOutRight from Attendent where  right(DateTime,8)='" & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000") & "'", cnData)
    If rsCheckInOut.State <> 0 And rsCheckInOut.RecordCount > 0 Then rsCheckInOut.MoveFirst
    With rsCheckInOut
        .Find "Cash_ID='" & txtID.Text & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If .Fields("InOutRight") = False Then
               MsgBox "B¹n ®· ra ca råi !!"
            End If
        Else
            MsgBox "B¹n ch­a vµo ca nªn kh«ng thÓ ra ca"
            Exit Sub
        End If
        
    End With
    With rsInOut
        .addNew
        .Fields("Cash_ID") = txtID.Text
        .Fields("DateTime") = Format(Now, "HH:mm:ss") & Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000")
        .Fields("TimeType") = "O"
        .Fields("InOutRight") = False
        .Update
        .Requery
    End With
    'Xoa du lieu trong ban in phieu
    cnData.Execute "Delete  from InOutAttendent"
    'Ghi nhan du lieu moi xuong bang cham cong
    With rsInOutPrint
        .addNew
        .Fields("Cash_ID") = txtID.Text
        With rsEmployee
            .Find "Cashier_ID='" & txtID.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                rsInOutPrint.Fields("Cash_Name") = .Fields("EmpName")
            End If
        End With
        .Fields("DateInOut") = Format(Day(Date), "00") & Format(Month(Date), "00") & Format(Year(Date), "0000")
        .Fields("TimeInOut") = Format(Now, "HH:mm:ss")
        .Fields("InOutRight") = "O"
        .Update
        .Requery
    End With
Call Print_LogIn(Trim(txtID.Text))
txtID.Text = ""
txtpass.Text = ""
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - cmdOut_Click"

End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    dtpTime.Value = Format(Now, "HH:mm:ss")
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    Set rsEmployee = Open_Table(cnData, "Employee")
    Set rsInOut = Open_Table(cnData, "Attendent")
    Set rsInOutPrint = Open_Table(cnData, "InOutAttendent")
    DescArr = LoadLanguage(LngFile, "#03:011:")
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Private Sub Timer1_Timer()
    dtpTime.Value = Format(Now, "HH:mm:ss")
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        With rsEmployee
            .Find "Cashier_ID='" & Trim(txtID.Text) & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtpass.Text = .Fields("EmpName")
            End If
        End With
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtID_KeyPress"
End Sub
Public Sub Print_LogIn(Clerk_ID As String)
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL As String
    SQL = "select Cash_ID,Cash_Name,DateInOut,TimeInOut,InOutRight from InOutAttendent"
    Set crClerkLogin = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crClerkLogin
        .Database.AddADOCommand cnData, cmd
        .CashID.SetUnboundFieldSource "{ado.Cash_ID}"
        .CashName.SetUnboundFieldSource "{ado.Cash_Name}"
        .LoginDate.SetUnboundFieldSource "{ado.DateInOut}"
        .LogInTime.SetUnboundFieldSource "{ado.TimeInOut}"
        .InOutType.SetUnboundFieldSource "{ado.InOutRight}"
    End With
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = crClerkLogin
        .Show vbModal
    End With
    Unload frmShow_Report_80
Exit Sub
Handle:
    Exit Sub
MsgBox Err.Number & Err.Description & Me.name & " Print_LogIn"
Unload frmShow_Report_80
End Sub
