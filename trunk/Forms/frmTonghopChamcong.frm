VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmTonghopChamcong 
   Caption         =   "Tæng hîp ngµy c«ng"
   ClientHeight    =   10335
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
   Icon            =   "frmTonghopChamcong.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10335
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin prjTouchScreen.MyButton cmdClose 
      Height          =   975
      Left            =   13680
      TabIndex        =   5
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTonghopChamcong.frx":000C
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
      Height          =   8775
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15478
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tæng hîp ngµy c«ng"
      TabPicture(0)   =   "frmTonghopChamcong.frx":0028
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "gridChamcong"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSDataGridLib.DataGrid gridChamcong 
         Height          =   8295
         Left            =   0
         TabIndex        =   4
         Top             =   300
         Width           =   14895
         _ExtentX        =   26273
         _ExtentY        =   14631
         _Version        =   393216
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
   Begin MSComCtl2.DTPicker dtpFromdate 
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      Format          =   108265473
      UpDown          =   -1  'True
      CurrentDate     =   41082
   End
   Begin MSComCtl2.DTPicker dtpTodate 
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   960
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      Format          =   108265473
      UpDown          =   -1  'True
      CurrentDate     =   41082
   End
   Begin prjTouchScreen.MyButton cmdTinhcong 
      Height          =   975
      Left            =   10080
      TabIndex        =   8
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTonghopChamcong.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdExport 
      Height          =   975
      Left            =   11880
      TabIndex        =   9
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1720
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTonghopChamcong.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label Label3 
      Caption         =   "§Õn ngµy:"
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Tõ ngµy:"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Tæng hîp chÊm c«ng "
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmTonghopChamcong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rschamcong As New ADODB.Recordset
Dim strFdate, strTdate, strMonth As String
Dim strChamcong As String
Public ExcelApp As New Excel.Application


Private Sub cmdClose_Click()
    Unload Me
    CloseRecordset rschamcong
End Sub

Private Sub cmdExport_Click()
On Error GoTo Handle
    Call XuatExcel(rschamcong, WorkingFolder & "\Chamcong" & Format(Month(dtpFromdate.Value), "00") & "-" & Format(Year(dtpFromdate.Value), "0000") & ".xls", "ChÊm c«ng 09-2012")
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "  cmdExport_Click"
End Sub


Public Sub XuatExcel(ByVal rs As Recordset, ByVal Pathfilename As String, ByVal Title As String)
    Dim i As Long, j As Long, k As Long, iSocot As Long, iSodong As Long, iWidth As Long
    Dim Temp As String
    Dim aSplit() As String
   
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    aSplit = Split(Pathfilename, "\")
    Temp = aSplit(UBound(aSplit))
   
    ExcelApp.DisplayAlerts = False 'khong cho show thong bao Save
   '-----------------------------------------------------------'
   If Dir(Pathfilename) <> "" Then Kill Pathfilename 'neu co thi xoa file
   '-----------------------------------------------------------'
   ExcelApp.Workbooks.Add
    ExcelApp.Workbooks(Workbooks.count).SaveAs Pathfilename
    Call ExcelApp.Workbooks(Temp).Worksheets.Add
   
    iSocot = rs.Fields.count
    iSodong = rs.RecordCount
   
    ExcelApp.ActiveSheet.Name = Temp
    With ExcelApp.Workbooks(Temp).Worksheets(Temp)
        Cells(1, 1) = Title
        Cells(1, 1).Font.Size = 15
        Cells(1, 1).Font.Bold = True
        Cells(1, 1).HorizontalAlignment = xlCenter
        Range(Cells(1, 1), Cells(1, iSocot + 1)).MergeCells = True
        Cells(3, 1) = "No"
        Range(Cells(3, 1), Cells(3, iSocot + 1)).Font.Size = 11
        Range(Cells(3, 1), Cells(3, iSocot + 1)).Font.Bold = True
        Range(Cells(3, 1), Cells(3, iSocot + 1)).HorizontalAlignment = xlCenter
        For j = 1 To iSocot
            Cells(3, j + 1) = rs.Fields(j - 1).Name
            'dong khung Title     '
           Range(Cells(3, j), Cells(3, j + 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
            If j <> iSocot Then
                Range(Cells(3, j), Cells(3, j + 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
            End If
            Range(Cells(3, j), Cells(3, j + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
        Next j
       
        For i = 1 To iSodong
            Cells(3 + i, 1) = i
            For k = 1 To iSocot
                Cells(3 + i, k + 1) = rs.Fields(k - 1)
               
                'dong khung cac cells  '
               Range(Cells(3 + i, k), Cells(3 + i, k + 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
                If k <> iSocot Then
                    Range(Cells(3 + i, k), Cells(3 + i, k + 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
                End If
                Range(Cells(3 + i, k), Cells(3 + i, k + 1)).Borders(xlEdgeTop).LineStyle = xlContinuous
                If i <> (iSodong - 1) Then
                    Range(Cells(3 + i, k), Cells(3 + i, k + 1)).Borders(xlEdgeBottom).LineStyle = xlContinuous
                End If
                Cells(3 + i, k + 1).ColumnWidth = Len(rs.Fields(k - 1)) + 10
            Next k
            rs.MoveNext
        Next i
        Range(Cells(3, iSocot + 1), Cells(iSodong + 3, iSocot + 1)).Borders(xlEdgeRight).LineStyle = xlContinuous
    End With
   
    ExcelApp.ActiveWorkbook.Save
    Screen.MousePointer = vbDefault
    ExcelApp.Visible = True 'cho show Excel
   
End Sub


Private Sub cmdTinhcong_Click()
On Error GoTo Handle
    strFdate = Format(Year(dtpFromdate.Value), "0000") & Format(Month(dtpFromdate.Value), "00") & Format(Day(dtpFromdate.Value), "00")
    strTdate = Format(Year(dtpTodate.Value), "0000") & Format(Month(dtpTodate.Value), "00") & Format(Day(dtpTodate.Value), "00")
    strMonth = Format(Month(dtpFromdate.Value), "00") & Format(Year(dtpFromdate.Value), "0000")
 
    If Left(strFdate, 6) <> Left(strTdate, 6) Then Exit Sub
    strChamcong = "SELECT Distinct Ngaycong" & strMonth & ".Emp_ID,Ngaycong" & strMonth & ".Emp_Name ," & _
                  " Work_Shift.Shift_Name,Work_Shift.InTime , Work_Shift.OutTime," & _
                  "[01In], [01Out] , [02In] ,[04In], [04Out], [05In], [05Out], [06In], [06Out]," & _
                  "[07In], [07Out], [08In],[08Out], [09In], [09Out],  [10In]," & _
                  "[10Out],[11In], [11Out], [12In], [12Out], [13In], [13Out], [14In]," & _
                  "[14Out],[15In], [15Out], [16In], [16Out], [17In], [17Out], [18In], [18Out]," & _
                  "[19In], [19Out],[20In], [20Out], [21In], [21Out], [22In], [22Out], [23In], [23Out]," & _
                  "[24In], [24Out], [25In], [25Out], [26In], [26Out], [27In], [27Out]," & _
                  "[28In], [28Out], [29In], [29Out], [30In], [30Out], [31In], [31Out]" & _
                " FROM (Work_Shift INNER JOIN Employee ON Work_Shift.Shift_ID = Employee.Shift)" & _
                " INNER JOIN Ngaycong" & strMonth & " ON Employee.Cashier_ID = Ngaycong" & strMonth & ".Emp_ID"
           

        Set rschamcong = OpenCriticalTable(strChamcong, cnData)
        Set gridChamcong.DataSource = rschamcong
        gridChamcong.AllowRowSizing = True
        Call init_Datagrid
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  -cmdTinhcong_Click "
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    dtpFromdate.Value = "01/" & Format(Month(Date), "00") & "/" & Format(Year(Date), "0000")
    dtpTodate.Value = Date
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name
End Sub

Public Sub init_Datagrid()
    On Error GoTo Handle
        With gridChamcong
            .Columns(0).Caption = "M· NV"
            .Columns(0).Width = 700
            .Columns(1).Caption = "Tªn NV"
            .Columns(1).Width = 1700
            .Columns(2).Caption = "Ca"
            .Columns(2).Width = 1000
            .Columns(3).Caption = "Giê vµo"
            .Columns(3).Width = 1000
            .Columns(4).Caption = "Giê ra"
            .Columns(4).Width = 1000
            For i = 5 To 35
'                .Columns(i).Caption = Format(i - 4, "00")
                .Columns(i).AllowSizing = True
                .Columns(i).Width = 1000
            Next
             
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " init_Datagrid"
End Sub
