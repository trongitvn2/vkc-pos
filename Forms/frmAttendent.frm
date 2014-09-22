VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAttendent 
   Caption         =   "Cham cong nhan vien"
   ClientHeight    =   10995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAttendent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboShift 
      Height          =   390
      Left            =   12240
      TabIndex        =   5
      Text            =   "Ca lµm viÖc"
      Top             =   120
      Width           =   2415
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9375
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   16536
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "VNI-Times"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Chaám coâng nhaân vieân"
      TabPicture(0)   =   "frmAttendent.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdOut"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "flgNhanvien"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdIn"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Toång hôïp ngaøy coâng"
      TabPicture(1)   =   "frmAttendent.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblDenngay"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblFromdate"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdReport"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdExport"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "dtpToDate"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "dtpFromDate"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "dtgatt"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdGetAtt"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin prjTouchScreen.MyButton cmdGetAtt 
         Height          =   855
         Left            =   -61560
         TabIndex        =   11
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
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
         BCOL            =   14737632
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAttendent.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin MSDataGridLib.DataGrid dtgatt 
         Height          =   7935
         Left            =   -74880
         TabIndex        =   6
         Top             =   1320
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   13996
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   26
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
            Size            =   14.25
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
      Begin prjTouchScreen.MyButton cmdIn 
         Height          =   1095
         Left            =   12960
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1931
         BTYPE           =   5
         TX              =   "Vµo ca"
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
         BCOL            =   14737632
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAttendent.frx":0060
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid flgNhanvien 
         Height          =   8655
         Left            =   240
         TabIndex        =   2
         Top             =   540
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   15266
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial NarrowH"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin prjTouchScreen.MyButton cmdOut 
         Height          =   1095
         Left            =   12960
         TabIndex        =   4
         Top             =   2880
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1931
         BTYPE           =   5
         TX              =   "Ra ca"
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
         BCOL            =   14737632
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAttendent.frx":007C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   495
         Left            =   -66780
         TabIndex        =   7
         Top             =   630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   90832897
         UpDown          =   -1  'True
         CurrentDate     =   40858
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   495
         Left            =   -63540
         TabIndex        =   8
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   90832897
         UpDown          =   -1  'True
         CurrentDate     =   40858
      End
      Begin prjTouchScreen.MyButton cmdExport 
         Height          =   855
         Left            =   -61560
         TabIndex        =   12
         Top             =   3840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         BTYPE           =   3
         TX              =   "XuÊt sang Excel"
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
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAttendent.frx":0098
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdReport 
         Height          =   855
         Left            =   -61560
         TabIndex        =   13
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         BTYPE           =   3
         TX              =   "B¸o c¸o"
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
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAttendent.frx":00B4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.Label lblFromdate 
         Caption         =   "Tõ ngµy:"
         Height          =   375
         Left            =   -68040
         TabIndex        =   10
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label lblDenngay 
         Caption         =   "§Õn ngµy:"
         Height          =   375
         Left            =   -65100
         TabIndex        =   9
         Top             =   690
         Width           =   1305
      End
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   13320
      TabIndex        =   14
      Top             =   10320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1296
      BTYPE           =   4
      TX              =   "§ãn&g"
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
      BCOLO           =   16777215
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmAttendent.frx":00D0
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
      Caption         =   "CHAÁM COÂNG NHAÂN VIEÂN"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   21.75
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
      Top             =   0
      Width           =   17175
   End
End
Attribute VB_Name = "frmAttendent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsNhanvien As New ADODB.Recordset
Dim rsAtt As New ADODB.Recordset
Dim Emp_ID, Emp_Name As String

Private Sub cboShift_Change()
On Error GoTo Handle
    Set rsNhanvien = OpenCriticalTable("SELECT Employee.Cashier_ID, Employee.EmpName, Company_Dept.Dept_Name, Work_Shift.Shift_Name" & _
                                        " FROM Work_Shift INNER JOIN (Company_Dept INNER JOIN Employee ON Company_Dept.Dept_ID = " & _
                                        " Employee.Dept_ID) ON Work_Shift.Shift_ID = Employee.Shift" & _
                                        " Where Work_Shift.Shift_ID='" & Format(cboShift.ItemData(cboShift.ListIndex), "00") & "'", cnData)
    Call InitFlex(rsNhanvien)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " - cboShift_Change"
End Sub

Private Sub cboShift_Click()
On Error GoTo Handle
    Call cboShift_Change
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " - cboShift_Click"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdGetAtt_Click()
On Error GoTo Handle
    Dim strSQL As String
    Dim rsAttendent As New ADODB.Recordset
    strSQL = "SELECT Employee.Cashier_ID, Employee.EmpName, Company_Dept.Dept_Name, Attendent.DateTimeIn," & _
             "Attendent.DateTimeOut" & _
             " FROM (Employee INNER JOIN Attendent ON Employee.Cashier_ID = Attendent.Cash_ID)" & _
             " INNER JOIN Company_Dept ON Employee.Dept_ID = Company_Dept.Dept_ID" & _
             " Where left(Attendent.DateTimeIn,10)>='" & dtpFromDate.Value & "' and left(Attendent.DateTimeIn,10)<='" & dtpToDate.Value & "'"
    Set rsAttendent = OpenCriticalTable(strSQL, cnData)
    Set dtgatt.DataSource = rsAttendent
    Call SetHeaderDatagrid
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " - cmdGetAtt_Click"
End Sub

Private Sub cmdIn_Click()
On Error GoTo Handle
Dim isSave As Boolean
If Emp_ID = "" Then
    MsgBox "Vui lßng chän m· sè Nh©n viªn cÇn vµo ca!", vbInformation
Exit Sub
End If
    With frmShowEmp_ID
        .lblDes = "Vµo ca lóc:"
        .Emp_ID = Emp_ID
        .Emp_Name = Emp_Name
        .Show vbModal
        isSave = .Let_OK
    End With
    If isSave = True Then
        With rsAtt
            .Find "Cash_ID='" & Emp_ID & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If Left(.Fields("DateTimeIn"), 10) = Format(Now, "dd/MM/yyyy") Then
                     MsgBox Emp_Name & " ®· vµo ca h«m nay"
                Else
                    .addNew
                    .Fields("Cash_ID") = Emp_ID
                    .Fields("DateTimeIn") = Now
                    .Fields("TimeType") = "I"
                    .Fields("InOutRight") = 1
                    .Update
                End If
            Else
                .addNew
                .Fields("Cash_ID") = Emp_ID
                .Fields("DateTimeIn") = Now
                .Fields("TimeType") = "I"
                .Fields("InOutRight") = 1
                .Update
            
            End If
            
        End With
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " - cmdIn_Click"
End Sub

Private Sub cmdOut_Click()
On Error GoTo Handle
Dim isSave As Boolean
If Emp_ID = "" Then
    MsgBox "Vui lßng chän m· sè Nh©n viªn cÇn ra ca!", vbInformation
Exit Sub
End If
    With frmShowEmp_ID
        .lblDes = "Ra ca lóc:"
        .Emp_ID = Emp_ID
        .Emp_Name = Emp_Name
        .Show vbModal
        isSave = .Let_OK
    End With
    If isSave = True Then
        With rsAtt
            .Find "Cash_ID='" & Emp_ID & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If Left(.Fields("DateTimeIn"), 10) = Format(Now, "dd/MM/yyyy") Then
                    If .Fields("DateTimeOut") <> "" Then
                        MsgBox Emp_Name & " ®· vµo ca h«m nay"
                    Else
                        .Fields("DateTimeOut") = Now
                        .Update
                    End If
                Else
                    .addNew
                    .Fields("Cash_ID") = Emp_ID
                    .Fields("DateTimeOut") = Now
                    .Fields("TimeType") = "O"
                    .Fields("InOutRight") = 1
                    .Update
                End If
            Else
                .addNew
                .Fields("Cash_ID") = Emp_ID
                .Fields("DateTimeOut") = Now
                .Fields("TimeType") = "O"
                .Fields("InOutRight") = 1
                .Update
            End If
        End With
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " - cmdOut_Click"
End Sub

Private Sub flgNhanvien_Click()
On Error GoTo Handle
    Emp_ID = flgNhanvien.TextMatrix(flgNhanvien.Row, 0)
    Emp_Name = flgNhanvien.TextMatrix(flgNhanvien.Row, 1)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " - flgNhanvien_Click"
End Sub

Private Sub flgNhanvien_DblClick()
On Error GoTo Handle
    Emp_ID = flgNhanvien.TextMatrix(flgNhanvien.Row, 0)
    Emp_Name = flgNhanvien.TextMatrix(flgNhanvien.Row, 1)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " - flgNhanvien_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    dtpFromDate.Value = Date
    dtpToDate.Value = Date
    If Check_Table_exist("Attendent") = False Then
        Call Create_Attendent
    End If
    If cnData.State <> 0 Then
        Set rsNhanvien = OpenCriticalTable("SELECT Employee.Cashier_ID, Employee.EmpName, Company_Dept.Dept_Name, Work_Shift.Shift_Name" & _
                                        " FROM Work_Shift INNER JOIN (Company_Dept INNER JOIN Employee ON Company_Dept.Dept_ID = Employee.Dept_ID) ON Work_Shift.Shift_ID = Employee.Shift", cnData)
        Set rsAtt = Open_Table(cnData, "Attendent")
    End If
    Call LoadNhanvien
    Call InitFlex(rsNhanvien)
    Call Add_Shift_To_Combo
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " - Form_Load"
End Sub

Public Sub LoadNhanvien()
On Error GoTo Handle
    With flgNhanvien
            .Cols = 4
            .Rows = 3
            .ColWidth(0) = 1500
            .ColWidth(1) = 4200
            .ColWidth(2) = 3500
            .ColWidth(3) = 3500
            .TextMatrix(0, 0) = "M· nh©n viªn"
            .TextMatrix(0, 1) = "Tªn nh©n viªn"
            .TextMatrix(0, 2) = "B¶ng c«ng viÖc"
            .TextMatrix(0, 3) = "Ca lµm viÖc"
            .ColAlignment(0) = 2
            .ColAlignment(1) = 4
            .ColAlignment(2) = 4
            .ColAlignment(3) = 4
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
           
        End With
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "- LoadNhanvien"
End Sub

Public Sub InitFlex(rs As ADODB.Recordset)
On Error GoTo Handle
    Dim incount As Integer
        rs.MoveFirst
        With rs
            Do While Not .EOF
                incount = incount + 1
                flgNhanvien.Rows = rs.RecordCount + 1
                With flgNhanvien
                    .TextMatrix(incount, 0) = rs.Fields(0)
                    .TextMatrix(incount, 1) = rs.Fields(1)
                    .TextMatrix(incount, 2) = rs.Fields(2)
                    .TextMatrix(incount, 3) = rs.Fields(3)
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
    MsgBox Err.Number & Err.Description & Me.Name & " - InitFlex"

End Sub

Public Sub SetHeaderDatagrid()
On Error GoTo Handle
    With dtgatt
        .Columns(0).Caption = "M· nh©n viªn"
        .Columns(0).Width = 1500
        .Columns(1).Caption = "Tªn nh©n viªn"
        .Columns(1).Width = 3000
        .Columns(2).Caption = "Bé phËn"
        .Columns(2).Width = 1400
        .Columns(3).Caption = "Giê vµo"
        .Columns(3).Width = 3000
        .Columns(4).Caption = "Giê ra"
        .Columns(4).Width = 3000
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " - SetHeaderDatagrid"
End Sub

Public Sub Add_Shift_To_Combo()
On Error GoTo Handle
    Dim rsShift As New ADODB.Recordset
    Set rsShift = OpenCriticalTable("select Shift_ID,Shift_Name from Work_Shift", cnData)
    With cboShift
        .Clear
        If rsShift.State <> 0 And rsShift.RecordCount > 0 Then rsShift.MoveFirst
        Do While Not rsShift.EOF
            .AddItem rsShift.Fields("Shift_Name")
            .ItemData(cboShift.NewIndex) = rsShift.Fields("Shift_ID")
        rsShift.MoveNext
        Loop
        .ListIndex = 0
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " - Add_Shift_To_Combo"
End Sub
