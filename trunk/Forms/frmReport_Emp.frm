VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmReport_Emp 
   Caption         =   "B¸o c¸o chi tiÕt theo nh©n viªn phôc vô"
   ClientHeight    =   9540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14385
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
   ScaleHeight     =   9540
   ScaleWidth      =   14385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton cmdClose 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   4320
      Width           =   4215
      _ExtentX        =   7435
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
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReport_Emp.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdPrint 
      Height          =   735
      Left            =   2280
      TabIndex        =   10
      Top             =   3480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "&In b¸o c¸o"
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
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReport_Emp.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdDone 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "&Läc"
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
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReport_Emp.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Frame Frame3 
      Caption         =   "M· hµng"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   4335
      Begin VB.ComboBox cboMahang 
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
         Left            =   120
         TabIndex        =   13
         Text            =   "M· hµng"
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Nh©n viªn"
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
      Begin VB.ComboBox cboNhanvien 
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
         Left            =   120
         TabIndex        =   12
         Text            =   "Nh©n viªn phôc vô"
         Top             =   360
         Width           =   4095
      End
   End
   Begin VB.Frame fraDate 
      Caption         =   "Ngµy"
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4335
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   495
         Left            =   540
         TabIndex        =   5
         Top             =   240
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
         Format          =   63963137
         UpDown          =   -1  'True
         CurrentDate     =   40330
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   495
         Left            =   2700
         TabIndex        =   6
         Top             =   240
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
         Format          =   63963137
         UpDown          =   -1  'True
         CurrentDate     =   40330
      End
      Begin VB.Label lblFromdate 
         Caption         =   "Tõ :"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   405
      End
      Begin VB.Label lblDenngay 
         Caption         =   "§Õn:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2100
         TabIndex        =   7
         Top             =   330
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chi tiÕt theo nh©n viªn"
      ForeColor       =   &H00FF0000&
      Height          =   9495
      Left            =   4560
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin MSDataGridLib.DataGrid dtgItems 
         Height          =   9015
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   15901
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   21
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
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
End
Attribute VB_Name = "frmReport_Emp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsItems As New ADODB.Recordset
Dim strSql As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub Get_Emp()
On Error GoTo Handle
     Dim rsNhanvien As New ADODB.Recordset
     Set rsNhanvien = OpenCriticalTable("Select Cashier_ID,EmpName from Employee", cnData)
     cboNhanvien.Clear
     Do While Not rsNhanvien.EOF
        With cboNhanvien
            .AddItem rsNhanvien.Fields("EmpName")
            .ItemData(cboNhanvien.NewIndex) = rsNhanvien.Fields("Cashier_ID")
        End With
     rsNhanvien.MoveNext
     Loop
     cboNhanvien.ListIndex = 0
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Lay nhan vien vao Combo"
End Sub


Public Sub Get_Items()
On Error GoTo Handle
     Dim rsPLU As New ADODB.Recordset
     Set rsPLU = Open_Table(cnData, "Inventory")
     cboMahang.Clear
     cboMahang.AddItem "TÊt c¶", 0
     Do While Not rsPLU.EOF
        With cboMahang
            .AddItem rsPLU.Fields("ItemName")
            .ItemData(cboMahang.NewIndex) = Right("000000000000" & rsPLU.Fields("ItemNum"), 12)
        End With
     rsPLU.MoveNext
     Loop
     cboMahang.ListIndex = 0
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Lay Ma hang vao Combo"
End Sub

Private Sub cmdDone_Click()
On Error GoTo Handle
    Dim strSQLSort As String
    If cboMahang.ListIndex = 0 Then
        strSQLSort = "SELECT Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName ,Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS price, Invoice_Totals.Cashier_ID" & _
                " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number" & _
                " WHERE left(Invoice_Totals.DateTime,8)>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "' and  left(Invoice_Totals.DateTime,8)>='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'" & _
                " and Invoice_Totals.OrderMan = '" & Format(cboNhanvien.ItemData(cboNhanvien.ListIndex), "00") & "'" & _
                " GROUP BY Invoice_Itemized.DiffItemName, Invoice_Totals.Cashier_ID, Invoice_Itemized.ItemNum"
    Else
        strSQLSort = "SELECT Invoice_Itemized.ItemNum , Invoice_Itemized.DiffItemName,Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS price, Invoice_Totals.Cashier_ID " & _
                " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number" & _
                " WHERE left(Invoice_Totals.DateTime,8)>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "' and  left(Invoice_Totals.DateTime,8)>='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'" & _
                " and Invoice_Totals.OrderMan = '" & Format(cboNhanvien.ItemData(cboNhanvien.ListIndex), "00") & "' and Invoice_Itemized.ItemNum='" & Right("000000000000" & cboMahang.ItemData(cboMahang.ListIndex), 12) & "'" & _
                " GROUP BY Invoice_Itemized.DiffItemName, Invoice_Totals.Cashier_ID, Invoice_Itemized.ItemNum"
    End If
    
    Set rsItems = OpenCriticalTable(strSQLSort, cnData)
    If Not rsItems.EOF Then
        Set dtgItems.DataSource = rsItems
        Call Set_ColumnHeader_Grid
    Else
        Set dtgItems.DataSource = Nothing
    End If
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdDone_Click"
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim iReport As CRAXDDRT.Report
    Dim SQL, sql1 As String
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Select Case cboMahang.ListIndex
        Case 0
            SQL = "SELECT Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName ,Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS price, Invoice_Totals.Cashier_ID" & _
                " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number" & _
                " WHERE left(Invoice_Totals.DateTime,8)>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "' and  left(Invoice_Totals.DateTime,8)>='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'" & _
                " and Invoice_Totals.OrderMan = '" & Format(cboNhanvien.ItemData(cboNhanvien.ListIndex), "00") & "'" & _
                " GROUP BY Invoice_Itemized.DiffItemName, Invoice_Totals.Cashier_ID, Invoice_Itemized.ItemNum"
        Case Else
        sql1 = "SELECT Invoice_Itemized.ItemNum , Invoice_Itemized.DiffItemName,Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS price, Invoice_Totals.Cashier_ID " & _
                " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number" & _
                " WHERE left(Invoice_Totals.DateTime,8)>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "' and  left(Invoice_Totals.DateTime,8)>='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'" & _
                " and Invoice_Totals.OrderMan = '" & Format(cboNhanvien.ItemData(cboNhanvien.ListIndex), "00") & "' and Invoice_Itemized.ItemNum='" & Right("000000000000" & cboMahang.ItemData(cboMahang.ListIndex), 12) & "'" & _
                " GROUP BY Invoice_Itemized.DiffItemName, Invoice_Totals.Cashier_ID, Invoice_Itemized.ItemNum"
       
    End Select
    
    Set crEmp_Details = Nothing
        cmd.ActiveConnection = cnData
        Select Case cboMahang.ListIndex
        Case 0
            cmd.CommandText = SQL
        Case Else
            cmd.CommandText = sql1
        End Select
        cmd.Execute
    With crEmp_Details
        
        .Database.AddADOCommand cnData, cmd
        .txtPluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .txtPluName.SetUnboundFieldSource "{ado.DiffItemName}"
        .txtCost.SetUnboundFieldSource "{ado.Price}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        
        .EmpID.SetText Format(cboNhanvien.ItemData(cboNhanvien.ListIndex), "00")
        .EmpName.SetText cboNhanvien.Text
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtCost
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With

    End With
    Set iReport = crEmp_Details
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
Exit Sub
Handle:
Exit Sub
MsgBox Err.Number & Err.Description & Me.name & " cmdBaocaochitiet_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    dtpFromDate.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    dtpToDate.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    Call Get_Emp
    Call Get_Items
Exit Sub
Handle:
    MsgBox Err.ne & Err.Description & " Form_Load"
End Sub

Public Sub Set_ColumnHeader_Grid()
On Error GoTo Handle
    With dtgItems
        .Columns(0).Caption = "M· hµng"
        .Columns(0).Width = 2000
        .Columns(1).Caption = "Tªn hµng"
        .Columns(1).Width = 3000
        .Columns(2).Caption = "Sè l­îng"
        .Columns(2).Width = 1200
        .Columns(3).Caption = "§¬n gi¸"
        .Columns(3).Width = 1200
        
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & " - Gan tieu de cho Grid"
End Sub
