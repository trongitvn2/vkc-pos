VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmShowReport_Details 
   Caption         =   "B¸o c¸o chi tiÕt"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShowReport_Details.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picToolsBar 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   15165
      TabIndex        =   0
      Top             =   0
      Width           =   15195
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   120
         TabIndex        =   12
         Top             =   -120
         Width           =   3735
         Begin VB.OptionButton optPayment 
            Caption         =   "Läc theo h×nh thøc thanh to¸n"
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   3495
         End
         Begin VB.OptionButton optLocation 
            Caption         =   "Läc theo khu vùc"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Value           =   -1  'True
            Width           =   3375
         End
      End
      Begin VB.Frame Frame1 
         Height          =   1095
         Left            =   3840
         TabIndex        =   11
         Top             =   -120
         Width           =   4335
         Begin prjTouchScreen.MyButton cmdView 
            Height          =   735
            Left            =   3360
            TabIndex        =   16
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1296
            BTYPE           =   5
            TX              =   "Xem"
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
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   16711680
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmShowReport_Details.frx":000C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin VB.ComboBox cboFilter 
            Height          =   390
            Left            =   120
            TabIndex        =   15
            Text            =   "Combo1"
            Top             =   360
            Width           =   3135
         End
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         Height          =   870
         Left            =   10680
         Picture         =   "frmShowReport_Details.frx":0028
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   885
      End
      Begin VB.TextBox txtPage 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   8940
         TabIndex        =   6
         Top             =   75
         Width           =   855
      End
      Begin VB.CommandButton cmdFirst 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8160
         Picture         =   "frmShowReport_Details.frx":0712
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.CommandButton cmdPrevious 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   8520
         Picture         =   "frmShowReport_Details.frx":0A54
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.CommandButton cmdNext 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   9840
         Picture         =   "frmShowReport_Details.frx":0D96
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.CommandButton cmdLast 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   10260
         Picture         =   "frmShowReport_Details.frx":10D8
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.ComboBox cboZoom 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8280
         TabIndex        =   1
         Text            =   "cboZoom"
         Top             =   600
         Width           =   2415
      End
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   870
         Left            =   13440
         TabIndex        =   8
         Top             =   0
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1535
         BTYPE           =   5
         TX              =   "&§ãng"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmShowReport_Details.frx":141A
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
         Height          =   870
         Left            =   11640
         TabIndex        =   9
         Top             =   0
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   1535
         BTYPE           =   5
         TX              =   "XuÊt sang d¹ng kh¸c"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmShowReport_Details.frx":1436
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
   Begin CRVIEWERLibCtl.CRViewer crvReport 
      CausesValidation=   0   'False
      Height          =   6735
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1140
      Width           =   9975
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   0   'False
      EnableGroupTree =   0   'False
      EnableNavigationControls=   0   'False
      EnableStopButton=   0   'False
      EnablePrintButton=   0   'False
      EnableZoomControl=   0   'False
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   0   'False
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmShowReport_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CRReport As New CRAXDDRT.Report
Dim iReport As New CRAXDDRT.Report
Dim TotalRptPage As Integer
Dim isLoad As Boolean
Dim FromDate, ToDate As String
Dim ReportNum As Integer
Dim DescArrReport() As String

Public Property Let Report(ByVal vNewValue As Variant)
    Set CRReport = vNewValue
End Property

Private Sub cboZoom_Change()
On Error GoTo errHdl

    Select Case cboZoom.ListIndex
        Case 0
            crvReport.Zoom 1
        Case 1
            crvReport.Zoom 2
        Case 2
            crvReport.Zoom 400
        Case 3
            crvReport.Zoom 300
        Case 4
            crvReport.Zoom 200
        Case 5
            crvReport.Zoom 150
        Case 6
            crvReport.Zoom 100
        Case 7
            crvReport.Zoom 75
        Case 8
            crvReport.Zoom 50
        Case 9
            crvReport.Zoom 25
        Case Else
            If IsNumeric(cboZoom.Text) Then
                If Val(cboZoom.Text) < 1000 And Val(cboZoom.Text) > 10 Then crvReport.Zoom CInt(cboZoom.Text)
            End If
    End Select

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cboZoom_Click()
On Error GoTo errHdl
    Call cboZoom_Change
    
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdClose_Click()
On Error GoTo errHdl

    Set iReport = Nothing
    Set CRReport = Nothing
    Unload Me

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdExport_Click()
On Error GoTo Handle
    iReport.Export
'    Select Case ReportNum
'        Case 2
'            Call Xuat_Excel
'        Case Else
'
'    End Select
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdExport_Click"
End Sub

Private Sub cmdFirst_Click()
On Error GoTo errHdl

    crvReport.ShowFirstPage
    While crvReport.IsBusy
        DoEvents
    Wend
    txtPage.Text = crvReport.GetCurrentPageNumber & " / " & TotalRptPage

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdLast_Click()
On Error GoTo errHdl

   crvReport.ShowLastPage

    While crvReport.IsBusy
        DoEvents
    Wend
    txtPage.Text = crvReport.GetCurrentPageNumber & " / " & TotalRptPage

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdNext_Click()
On Error GoTo errHdl

  crvReport.ShowNextPage
    
    While crvReport.IsBusy
        DoEvents
    Wend
    txtPage.Text = crvReport.GetCurrentPageNumber & " / " & TotalRptPage

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo errHdl

    crvReport.ShowPreviousPage
    While crvReport.IsBusy
        DoEvents
    Wend
    txtPage.Text = crvReport.GetCurrentPageNumber & " / " & TotalRptPage

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub


Private Sub cmdPrint_Click()
On Error GoTo errHdl
    myPrint iReport, crvReport.GetCurrentPageNumber, TotalRptPage

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdView_Click()
    Call ViewReport
End Sub

Private Sub Form_Activate()
    If isLoad = True Then Exit Sub
    DescArrReport = LoadLanguage(LngFile, "#05:001:")
    If optLocation.Value = True Then
        Call load_Location
    ElseIf optPayment.Value = True Then
        Call LoadPayment
    End If
    isLoad = True
End Sub

Private Sub Form_Resize()
On Error GoTo errHdl

    picToolsBar.Width = Me.ScaleWidth
    crvReport.Width = Me.ScaleWidth
    crvReport.Left = 0
    crvReport.Height = Me.ScaleHeight - (picToolsBar.Height + 720)

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errHdl
    Set iReport = Nothing
    Set crNewBalance = Nothing
    Set CRReport = Nothing
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub Form_Load()
On Error GoTo Handle
Set iReport = CRReport
isLoad = False
    With cboZoom
        .AddItem "Fix Page", 0
        .AddItem "Full Page", 1
        .AddItem "400 %", 2
        .AddItem "300 %", 3
        .AddItem "200 %", 4
        .AddItem "150 %", 5
        .AddItem "100 %", 6
        .AddItem "75 %", 7
        .AddItem "50 %", 8
        .AddItem "25 %", 9
    End With
    cboZoom.ListIndex = 0
    
    With crvReport
        .DisplayBorder = False
        .ReportSource = iReport
        .EnableSearchControl = False
        .EnableStopButton = False
        .EnableGroupTree = False
        .EnableAnimationCtrl = False
        .EnablePopupMenu = False
        .EnableToolbar = False
        .DisplayToolbar = False
        .DisplayTabs = False
        .ToolTipText = ""
        .ViewReport
        
        While .IsBusy
            DoEvents
        Wend
        .ShowLastPage
        While .IsBusy
            DoEvents
        Wend
       TotalRptPage = .GetCurrentPageNumber
        .ShowFirstPage
        While .IsBusy
            DoEvents
        Wend
    End With
    Me.WindowState = 2
    txtPage.Text = crvReport.GetCurrentPageNumber & " / " & TotalRptPage
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Private Sub Xuat_Excel()
Dim rsCustomer As New ADODB.Recordset
Dim sTapTinExcel As String ' Ta^.p tin Excel ca^`n ke^'t xua^'t
Dim NxtLine, lc As Integer
' Mo+? database ví du. kèm theo VB6 hoa(.c Access
sTapTinExcel = "C:\Customers.XLS"
Dim strSql As String
strSql = "SELECT Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price,Avg(Invoice_Itemized.PricePer)*Sum(Invoice_Itemized.Quantity)as Amount, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime" & _
          " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number " & _
          " WHERE Left([Invoice_Totals].[DateTime],8)>='" & FromDate & "' And Left([Invoice_Totals].[DateTime],8)<='" & ToDate & "'" & _
          " GROUP BY Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer"

Set rsCustomer = OpenCriticalTable(strSql, cnData)
Dim xlApp As Object ' Kho+?i do^.ng Excel, nhung không hie^?n thi.
Set xlApp = CreateObject("Excel.Application")

Dim xlBook As Object
Set xlBook = xlApp.Workbooks.Add ' Thêm workbook mo+'i

Dim xlSheet As Object
Set xlSheet = xlBook.Worksheets(1) ' Làm vie^.c vo+'i Sheet1 trong Excel

xlSheet.Cells.Font.name = ".vnArial" ' Font cho toàn bo^. sheet
xlSheet.Range("A1:E1").Select ' Tiêu ?e^` chie^'m 11 co^.t
With xlApp.Selection
.Font.Size = 14
.Font.Bold = True
.RowHeight = .Font.Size * 1.4
End With

With xlSheet
    .Columns.EntireColumn.AutoFit ' Chi?nh kích thu+o+'c các co^.t cho kho+'p du+~ lie^.u
    
    .Rows("1:1").Insert Shift:=xlDown ' Chèn thêm dòng tu+.a
    .Range("A1").FormulaR1C1 = "B¸o c¸o chi tiÕt"
    .Range("A2").FormulaR1C1 = "Tõ ngµy:" & gfCONVERT_STRING_TO_DATE(FromDate) & "§Õn ngµy:" & gfCONVERT_STRING_TO_DATE(ToDate)
    .Range("A2:E2").Select
End With
' Chi?nh da.ng dòng tu+.a
With xlSheet.Rows("1:1").Font
    .Size = 18
    .Bold = True
End With
' Ðo.c du+~ lie^.u tu+` table Customers, gán vào các co^.t trong excel
Dim i As Integer
For i = 1 To rsCustomer.Fields.count ' Tên các co^.t ?u+o+.c in ?a^.m
    xlApp.ActiveSheet.Cells(3, i).Font.Bold = True
Next

NxtLine = 3 ' Ba('t ?a^`u in các tên cu?a co^.t du+~ lie^.u
For lc = 0 To rsCustomer.Fields.count - 3
    xlApp.ActiveSheet.Cells(NxtLine, lc + 1).Value = rsCustomer.Fields(lc).name
Next

NxtLine = 4 ' Ba('t ?a^`u in du+~ lie^.u
Do Until rsCustomer.EOF
    For lc = 0 To rsCustomer.Fields.count - 3
        xlApp.ActiveSheet.Cells(NxtLine, lc + 1).Value = rsCustomer.Fields(lc)
        xlApp.Columns.EntireColumn.AutoFit
        If rsCustomer.Fields.Item(lc).name <> "DATE" Then
            xlApp.ActiveSheet.Cells(NxtLine, lc + 1).Value = rsCustomer.Fields(lc)
        Else ' Ch?nh d?ng d? li?u ngày tháng
            xlApp.ActiveSheet.Cells(NxtLine, lc + 1).Value = Format(rsCustomer.Fields(lc), "dd/mm/yy")
        End If
    Next
    rsCustomer.MoveNext
    NxtLine = NxtLine + 1
Loop

' Merge 11 ô ?e^? dòng tu+.a chu+'a he^'t trong 11 ô này
xlSheet.Range("A1:K1").Select
xlApp.Selection.Merge
' Ghi lên dia
xlBook.SaveAs FileName:=sTapTinExcel
xlBook.Saved = True ' Ghi ta^.p tin thành công
xlApp.Quit ' Ðóng Excel

rsCustomer.Close ' Ðóng table
End Sub


Public Property Let Get_fDate(ByVal vNewValue As Variant)
    FromDate = vNewValue
End Property

Public Property Let Get_tDate(ByVal vNewValue As Variant)
    ToDate = vNewValue
End Property

Public Property Let Report_Number(ByVal vNewValue As Variant)
    ReportNum = vNewValue
End Property

Public Sub load_Location()
On Error GoTo Handle
    Dim rsLocation As New ADODB.Recordset
    If cnData.State = adStateClosed Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    Set rsLocation = Open_Table(cnData, "Table_Diagram_Sections")
    With cboFilter
    .Clear
    .AddItem "TÊt c¶"
        If rsLocation.RecordCount > 0 Then rsLocation.MoveFirst
        With rsLocation
            Do While Not .EOF
                cboFilter.AddItem .Fields("Section_ID")
                cboFilter.ItemData(cboFilter.NewIndex) = CInt(.Fields("Location_ID"))
            .MoveNext
            Loop
        End With
        .ListIndex = 0
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " load_Location"
End Sub

Public Sub LoadPayment()
On Error GoTo Handle
Dim arrdes() As String
Dim i As Integer
arrdes = LoadLanguage(LngFile, "#02:003:")
    With cboFilter
        .Clear
        .AddItem "TÊt c¶"
        For i = 3 To 8
            .AddItem arrdes(i)
        Next
        .ListIndex = 0
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " LoadPayment"

End Sub
Private Sub ViewReport()
On Error GoTo errHdl
    Dim cmd As New ADODB.Command
    Dim SQL, SQLSort As String
'    SQL = "SELECT Invoice_Totals.Payment_Method,Invoice_Totals.Station_ID,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price" & _
'                            " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number " & _
'                            " WHERE Left([Invoice_Totals].[DateTime],8)>='" & FromDate & "' And Left([Invoice_Totals].[DateTime],8)<='" & ToDate & "'" & _
'                            " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'"
    
    SQL = "SELECT Invoice_Totals.Payment_Method,Invoice_Totals.Station_ID,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Inventory.Dept_ID, Departments.Description" & _
            " FROM Departments INNER JOIN (Inventory INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Departments.Dept_ID = Inventory.Dept_ID" & _
            " WHERE Left([Invoice_Totals].[DateTime],8)>='" & FromDate & "' And Left([Invoice_Totals].[DateTime],8)<='" & ToDate & "'" & _
             " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'"
    If optPayment.Value = True Then
        Select Case cboFilter.ListIndex
            Case 0
                 SQLSort = " GROUP BY Invoice_Totals.Payment_Method,Invoice_Totals.Station_ID, Inventory.Dept_ID, Departments.Description,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer"
            Case 1
                SQLSort = " And [Invoice_Totals].[Payment_Method]='C' GROUP BY Invoice_Totals.Payment_Method,Invoice_Totals.Station_ID, Inventory.Dept_ID, Departments.Description,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer"
            Case 2
                SQLSort = " And [Invoice_Totals].[Payment_Method]='CT' GROUP BY Invoice_Totals.Payment_Method,Invoice_Totals.Station_ID, Inventory.Dept_ID, Departments.Description,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer"
            Case 3
                SQLSort = " And [Invoice_Totals].[Payment_Method]='GC' GROUP BY Invoice_Totals.Payment_Method,Invoice_Totals.Station_ID, Inventory.Dept_ID, Departments.Description,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer"
            Case 4
                SQLSort = " And [Invoice_Totals].[Payment_Method]='OA' GROUP BY Invoice_Totals.Payment_Method,Invoice_Totals.Station_ID, Inventory.Dept_ID, Departments.Description,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer"
            Case 5
                SQLSort = " And [Invoice_Totals].[Payment_Method]='CC' GROUP BY Invoice_Totals.Payment_Method,Invoice_Totals.Station_ID, Inventory.Dept_ID, Departments.Description,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer"
            Case 6
                SQLSort = " And [Invoice_Totals].[Payment_Method]='ROA' GROUP BY Invoice_Totals.Payment_Method,Invoice_Totals.Station_ID, Inventory.Dept_ID, Departments.Description,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer"
        End Select
    ElseIf optLocation.Value = True Then
        If cboFilter.ListIndex = 0 Then
            SQLSort = " GROUP BY Invoice_Totals.Payment_Method,Invoice_Totals.Station_ID, Inventory.Dept_ID, Departments.Description,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer"
        Else
            SQLSort = " And [Invoice_Totals].[Station_ID]='" & Format(cboFilter.ItemData(cboFilter.ListIndex), "00") & "' GROUP BY Invoice_Totals.Payment_Method,Invoice_Totals.Station_ID, Inventory.Dept_ID, Departments.Description,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer"
        End If
    End If
    
    SQL = SQL & SQLSort
    Set crSaleReport = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crSaleReport
        .Database.AddADOCommand cnData, cmd
        .txtPluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .txtPluName.SetUnboundFieldSource "{ado.DiffItemName}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtCost.SetUnboundFieldSource "{ado.Price}"
        
        
        .txtGroup.SetUnboundFieldSource "{ado.Description}"
        .txtFromDate.SetText gfCONVERT_STRING_TO_DATE(FromDate)
        .txtToDate.SetText gfCONVERT_STRING_TO_DATE(ToDate)
        .lblNum.SetText ToDate
        .lblNgay.SetText "Ngµy " & Right(ToDate, 2) & " th¸ng " & Mid(ToDate, 5, 2) & " n¨m " & Left(ToDate, 4)
        'canh le
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign
        .lblStt.SetText DescArrReport(43)
        .txtTitle.SetText DescArrReport(42)
        .lblItemcode.SetText DescArrReport(44)
        .lblItemName.SetText DescArrReport(45)
        .lblUnit.SetText DescArrReport(48)
        .lblQty.SetText DescArrReport(46)
        .lblPrice.SetText DescArrReport(47)
        .lblAmount.SetText DescArrReport(49)
        .lblInword.SetText DescArrReport(56)
        .lblCashier.SetText DescArrReport(57)
        .lblChief.SetText DescArrReport(58)
        .lblDirector.SetText DescArrReport(59)
        .lblSign1.SetText DescArrReport(60)
        .lblSign2.SetText DescArrReport(60)
        .lblSign3.SetText DescArrReport(60)
        .lblFromdate.SetText DescArrReport(40)
        .lblToDate.SetText DescArrReport(41)

         .Section1.Suppress = True
         .Section3.Suppress = True
         .Section12.Suppress = True
         .Section5.Suppress = False
         .Section14.Suppress = True
         
         If optPayment.Value = True Then
            .txtPaymentMethod.SetUnboundFieldSource "{ado.Payment_Method}"
             .Section16.Suppress = False
             .Section3.Suppress = True
         ElseIf optLocation.Value = True Then
            .txtStationID.SetUnboundFieldSource "{ado.Station_ID}"
             .Section16.Suppress = True
             .Section3.Suppress = False
         End If
        
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
        With .txtStore
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtSumStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
    End With
    
    Set iReport = crSaleReport
    With crvReport
        .DisplayBorder = False
        .ReportSource = iReport
        .EnableSearchControl = False
        .EnableStopButton = False
        .EnableGroupTree = False
        .EnableAnimationCtrl = False
        .EnablePopupMenu = False
        .EnableToolbar = False
        .DisplayToolbar = False
        .DisplayTabs = False
        .ToolTipText = ""
        .ViewReport
        While .IsBusy
            DoEvents
        Wend
         .ShowLastPage
         Delay (500)
        TotalRptPage = .GetCurrentPageNumber
        While .IsBusy
            DoEvents
        Wend
       .ShowFirstPage
'
        While .IsBusy
            DoEvents
        Wend
        
    End With
    txtPage.Text = crvReport.GetCurrentPageNumber & " / " & TotalRptPage
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub optLocation_Click()
    Call load_Location
End Sub

Private Sub optPayment_Click()
Call LoadPayment
End Sub
