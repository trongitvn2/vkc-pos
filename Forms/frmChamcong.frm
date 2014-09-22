VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmChamcong 
   Caption         =   "ChÊm c«ng nh©n viªn"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14520
   ClipControls    =   0   'False
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
   ScaleHeight     =   10950
   ScaleWidth      =   14520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "ChÊm c«ng"
      Height          =   3855
      Left            =   8280
      TabIndex        =   26
      Top             =   4080
      Width           =   6135
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3375
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   5953
         _Version        =   393216
         ForeColor       =   16711680
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtID 
      Height          =   495
      Left            =   11280
      TabIndex        =   25
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid flgNhanvien 
      Height          =   6975
      Left            =   0
      TabIndex        =   24
      Top             =   960
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   12303
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3015
      Left            =   120
      TabIndex        =   9
      Top             =   8040
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   5318
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Th«ng tin nh©n viªn"
      TabPicture(0)   =   "frmChamcong.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2535
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   13935
         Begin VB.TextBox txtCMND 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   390
            Left            =   7920
            TabIndex        =   23
            Top             =   1440
            Width           =   4095
         End
         Begin VB.TextBox txtDienthoai 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   390
            Left            =   2160
            TabIndex        =   22
            Top             =   1440
            Width           =   3735
         End
         Begin VB.Frame Frame2 
            Caption         =   "H×nh 4 x 6"
            Height          =   2175
            Left            =   12120
            TabIndex        =   21
            Top             =   240
            Width           =   1695
            Begin VB.Image Image1 
               Height          =   1695
               Left            =   120
               Stretch         =   -1  'True
               Top             =   360
               Width           =   1455
            End
         End
         Begin VB.TextBox txtDiachi 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   390
            Left            =   7080
            TabIndex        =   20
            Top             =   960
            Width           =   4935
         End
         Begin VB.TextBox txtNgaysinh 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   390
            Left            =   1440
            TabIndex        =   19
            Top             =   960
            Width           =   4455
         End
         Begin VB.TextBox txtBophan 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   390
            Left            =   7080
            TabIndex        =   18
            Top             =   360
            Width           =   3975
         End
         Begin VB.TextBox txtTenNV 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   390
            Left            =   1440
            TabIndex        =   17
            Top             =   360
            Width           =   4455
         End
         Begin VB.Label Label10 
            Caption         =   "CMND/Passport:"
            Height          =   375
            Left            =   6000
            TabIndex        =   16
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label9 
            Caption         =   "§iÖn tho¹i liªn hÖ:"
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   1440
            Width           =   1935
         End
         Begin VB.Label Label8 
            Caption         =   "§Þa chØ:"
            Height          =   375
            Left            =   6000
            TabIndex        =   14
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label7 
            Caption         =   "Ngµy sinh:"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Bé phËn"
            Height          =   375
            Left            =   6000
            TabIndex        =   12
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Hä tªn:"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1095
         End
      End
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Height          =   975
      Left            =   12840
      TabIndex        =   8
      Top             =   3000
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "§ãng"
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
      BCOL            =   255
      BCOLO           =   255
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmChamcong.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdSave 
      Height          =   975
      Left            =   8520
      TabIndex        =   7
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Vµo ca"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmChamcong.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpLogIn_Time 
      Height          =   495
      Left            =   9720
      TabIndex        =   5
      Top             =   1680
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      Format          =   108724226
      CurrentDate     =   41050
   End
   Begin MSComCtl2.DTPicker dtpDateLog 
      Height          =   495
      Left            =   9840
      TabIndex        =   1
      Top             =   960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      Format          =   108724225
      CurrentDate     =   41050
   End
   Begin MSComCtl2.DTPicker dtpLogOut_Time 
      Height          =   495
      Left            =   9720
      TabIndex        =   6
      Top             =   2280
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      _Version        =   393216
      Format          =   108724226
      CurrentDate     =   41050
   End
   Begin prjTouchScreen.MyButton cmdTinhcong 
      Height          =   975
      Left            =   11160
      TabIndex        =   28
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "TÝnh c«ng"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmChamcong.frx":0054
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
      Height          =   975
      Left            =   9840
      TabIndex        =   29
      Top             =   3000
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Ra ca"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmChamcong.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label Label4 
      Caption         =   "Giê ra:"
      Height          =   375
      Left            =   8400
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Giê vµo:"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Ngµy:"
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ChÊm c«ng nh©n viªn"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "frmChamcong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsNhanvien As New ADODB.Recordset
Dim strNhanvien As String
Dim rsAtt As New ADODB.Recordset
Dim stratt As String
Dim strMonth As String

Private Sub cmdClose_Click()
    CloseRecordset rsNhanvien
    CloseRecordset rsAtt
    Unload Me
End Sub

Private Sub cmdOut_Click()
On Error GoTo Handle
    Dim SQL As String
    Dim rsSave As New ADODB.Recordset
    SQL = "Select * from Att" & strMonth & " where Date_Log='" & gfCONVERT_DATE_TO_STRING(dtpDateLog.Value) & "'"
    Set rsSave = OpenCriticalTable(SQL, cnData)
    If rsSave.RecordCount > 0 Then
            rsSave.MoveFirst
    End If
    With rsSave
        .Find "Emp_ID='" & txtID.Text & "'", , adSearchForward, adBookmarkFirst
        
        If Not .EOF Then
            .Fields("LogOut_Time") = Format(dtpLogOut_Time.Value, "HH:mm:ss")
            .Update
        Else
            .addNew
            .Fields("Emp_ID") = txtID.Text
            .Fields("Date_Log") = gfCONVERT_DATE_TO_STRING(dtpDateLog.Value)
            .Fields("LogIn_Time") = ""
            .Fields("LogOut_Time") = Format(dtpLogOut_Time.Value, "HH:mm:ss")
            .Update
            .Requery
        End If
    End With
   
    Call view_Att
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & " cmdOut_Click"

End Sub

Private Sub cmdSave_Click()
On Error GoTo Handle
    Dim SQL As String
    Dim rsSave As New ADODB.Recordset
    SQL = "Select * from Att" & strMonth & " where Date_Log='" & gfCONVERT_DATE_TO_STRING(dtpDateLog.Value) & "'"
    Set rsSave = OpenCriticalTable(SQL, cnData)
    If rsSave.RecordCount > 0 Then
            rsSave.MoveFirst
    End If
    With rsSave
        .Find "Emp_ID='" & txtID.Text & "'", , adSearchForward, adBookmarkFirst
        
        If Not .EOF Then
            .Fields("LogIn_Time") = Format(dtpLogIn_Time.Value, "HH:mm:ss")
            .Update
        Else
            .addNew
            .Fields("Emp_ID") = txtID.Text
            .Fields("Date_Log") = gfCONVERT_DATE_TO_STRING(dtpDateLog.Value)
            .Fields("LogIn_Time") = Format(dtpLogIn_Time.Value, "HH:mm:ss")
            .Fields("LogOut_Time") = ""
            .Update
            .Requery
        End If
    End With
    MsgBox "Vµo ca lóc:" & Format(dtpLogIn_Time.Value, "HH:mm:ss")
    
    Call view_Att
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & " cmdSave_Click"
End Sub

Private Sub cmdTinhcong_Click()
On Error GoTo Handle
Dim strDate As String

Dim rsngaycong As New ADODB.Recordset
Dim strNgaycong As String
strDate = Format(Month(Date), "00") & Format(Year(Date), "0000")
    If Check_Table_exist("Ngaycong" & strDate) = False Then
        Call create_tblNgaycong(strDate)
    Else
        Call add_emp(strDate)
        Set rsngaycong = Open_Table(cnData, "Ngaycong" & strDate)
        Call Calculate(strDate)
    End If
    With frmTonghopChamcong
       .Show vbModal
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name
End Sub

Private Sub dtpDateLog_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    Set rsAtt = Nothing
    Set rsAtt = OpenCriticalTable("Select * from Att" & strMonth & " where Date_Log='" & gfCONVERT_DATE_TO_STRING(dtpDateLog.Value) & "'", cnData)
    Call view_Att
End Sub

Private Sub dtpDateLog_Change()
Set rsAtt = Nothing
    Set rsAtt = OpenCriticalTable("Select * from Att" & strMonth & " where Date_Log='" & gfCONVERT_DATE_TO_STRING(dtpDateLog.Value) & "'", cnData)
    Call view_Att
End Sub

Private Sub flgNhanvien_Click()
On Error GoTo Handle
    txtID.Text = flgNhanvien.TextMatrix(flgNhanvien.Row, 0)
    With rsNhanvien
        .Find "Cashier_ID='" & txtID.Text & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtTenNV.Text = rsNhanvien.Fields("EmpName")
            txtBophan.Text = rsNhanvien.Fields("Dept_Name")
            txtNgaysinh.Text = rsNhanvien.Fields("Birthday")
            txtDiachi.Text = rsNhanvien.Fields("Address")
            txtDienthoai.Text = rsNhanvien.Fields("Phone")
            If Dir(rsNhanvien.Fields("Picture"), vbDirectory) <> "" Then
                Image1.Picture = LoadPicture(rsNhanvien.Fields("Picture"))
            End If
        End If
    End With
    With rsAtt
        .Find "Emp_ID='" & txtID.Text & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            dtpDateLog.Value = gfCONVERT_STRING_TO_DATE(.Fields("Date_Log"))
            dtpLogIn_Time.Value = .Fields("LogIn_Time")
            dtpLogOut_Time.Value = .Fields("LogIn_Time")
        End If
        
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    strMonth = Format(Month(Date), "00") & Format(Year(Date), "0000")
    dtpDateLog.Value = Date
    dtpLogIn_Time.Value = Format(Now, "HH:mm:ss")
    dtpLogOut_Time.Value = Format(Now, "HH:mm:ss")
    If Not Check_Table_exist("Att" & strMonth) Then
        Call Create_Table_Att(strMonth)
        cnData.Execute "ALTER TABLE Att" & strMonth & " ADD PRIMARY KEY (Emp_ID,Date_Log)"
    End If
    strNhanvien = "SELECT Employee.Cashier_ID,Employee.Address,Employee.Phone,Employee.Birthday,Employee.Picture, Employee.EmpName, Company_Dept.Dept_Name, Work_Shift.Shift_Name" & _
                  " FROM Work_Shift INNER JOIN (Employee INNER JOIN Company_Dept ON Employee.Dept_ID = Company_Dept.Dept_ID) ON Work_Shift.Shift_ID = Employee.Shift" & _
                  " ORDER BY Work_Shift.Shift_Name,Employee.Cashier_ID"
    Set rsNhanvien = OpenCriticalTable(strNhanvien, cnData)
    Call Set_flgNhanvien
    Call SetFLGRIDNhanVien(rsNhanvien)
    'Xu ly luu du lieu
    Set rsAtt = OpenCriticalTable("Select * from Att" & strMonth & " where Date_Log='" & gfCONVERT_DATE_TO_STRING(dtpDateLog.Value) & "'", cnData)
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " Form_Load"
End Sub

Public Sub Set_flgNhanvien()
    On Error GoTo Handle
        With flgNhanvien
            .Cols = 4
            .Rows = 10
            .ColWidth(0) = 1500
            .ColWidth(1) = 3200
            .ColWidth(2) = 2000
            .TextMatrix(0, 0) = "M· Nh©n viªn"
            .TextMatrix(0, 1) = "Tªn Nh©n Viªn"
            .TextMatrix(0, 2) = "Bé phËn"
            .TextMatrix(0, 3) = "Ca"
            .ColAlignment(1) = 2
            .ColAlignment(2) = 4
            .ColAlignment(3) = 4
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
            
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Set_flgNhanvien"
End Sub

Public Sub SetFLGRIDNhanVien(rs As ADODB.Recordset)
On Error GoTo Handle
        Dim incount As Integer
        rs.MoveFirst
        With rs
            Do While Not .EOF
                incount = incount + 1
                flgNhanvien.Rows = rs.RecordCount + 1
                With flgNhanvien
                    .TextMatrix(incount, 0) = rs!cashier_ID
                    .TextMatrix(incount, 1) = rs!EmpName
                    .TextMatrix(incount, 2) = rs!Dept_Name
                    .TextMatrix(incount, 3) = rs!Shift_Name
                End With
            rs.MoveNext
            Loop
        End With
        If rs.RecordCount = 0 Then
            With flgNhanvien
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
            End With
        End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "SetFLGRIDNhanVien"
End Sub

Public Sub view_Att()
On Error GoTo Handle
    Dim str As String
    Dim rsview As New ADODB.Recordset
    str = "SELECT Att" & strMonth & ".Emp_ID, Employee.EmpName, Att" & strMonth & ".LogIn_Time," & _
          "Att" & strMonth & ".LogOut_Time" & _
          " FROM Employee INNER JOIN" & _
          " Att" & strMonth & " ON Employee.Cashier_ID = Att" & strMonth & ".Emp_ID" & _
          " where Att" & strMonth & ".Date_Log='" & gfCONVERT_DATE_TO_STRING(dtpDateLog.Value) & "'"
    Set rsview = OpenCriticalTable(str, cnData)
    DataGrid1.DefColWidth = 1400
    Set DataGrid1.DataSource = rsview
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " view_Att"
End Sub

Public Sub add_emp(ByVal strDate As String)
On Error GoTo Handle
    Dim i
    Dim rsemp As New ADODB.Recordset
    Dim rsngaycong As New ADODB.Recordset
    Set rsemp = Open_Table(cnData, "Employee")
    Set rsngaycong = Open_Table(cnData, "Ngaycong" & strDate)
    If rsemp.State <> 0 Then
        If rsemp.RecordCount > 0 Then rsemp.MoveFirst
    End If
    Do While Not rsemp.EOF
        With rsngaycong
            .Find "Emp_ID='" & rsemp.Fields("Cashier_ID") & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("Emp_ID") = rsemp.Fields("Cashier_ID")
                .Fields("Emp_Name") = rsemp.Fields("EmpName")
                For i = 1 To 31
                    .Fields(Format(i, "00") & "In") = ""
                    .Fields(Format(i, "00") & "Out") = ""
                Next i
                .Update
            End If
        End With
    rsemp.MoveNext
    Loop
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "add_emp"
End Sub

Public Sub Calculate(ByVal strDate As String)
On Error GoTo Handle
    Dim rsattendent As New ADODB.Recordset
    Dim rsworkday As New ADODB.Recordset
    Dim strattendent As String
    Dim i As Double
    Dim iday, iIn, iOut As String
    Set rsworkday = Open_Table(cnData, "Ngaycong" & strDate)
    For i = 1 To 31
        iday = Format(i, "00")
        iIn = Format(i, "00") & "In"
        iOut = Format(i, "00") & "Out"
        strattendent = "select * from Att" & strDate & " where right(Date_Log,2)='" & iday & "'"
        Set rsattendent = OpenCriticalTable(strattendent, cnData)
            Do While Not rsattendent.EOF
                With rsworkday
                    .Find "Emp_ID='" & rsattendent.Fields("Emp_ID") & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        .Fields(iIn) = rsattendent.Fields("LogIn_Time")
                        .Fields(iOut) = rsattendent.Fields("LogOut_Time") '.Fields "(" & Format(i, "00") & "In)" =
                        .Update
                    End If
                End With
            rsattendent.MoveNext
            Loop
    Next
    
Exit Sub
Handle:
MsgBox Err.Description & " - Calculate"

End Sub
