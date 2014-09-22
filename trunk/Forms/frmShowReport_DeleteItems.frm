VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmShowReport_DeleteItems 
   Caption         =   "B∏o c∏o"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13290
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShowReport_DeleteItems.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   13290
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Height          =   825
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   13245
      TabIndex        =   0
      Top             =   0
      Width           =   13275
      Begin VB.ComboBox CboFilter 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         ItemData        =   "frmShowReport_DeleteItems.frx":000C
         Left            =   5640
         List            =   "frmShowReport_DeleteItems.frx":001C
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   240
         Width           =   4095
      End
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   615
         Left            =   11640
         TabIndex        =   9
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "ß„n&g"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   16711680
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmShowReport_DeleteItems.frx":005C
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
         Height          =   615
         Left            =   9840
         TabIndex        =   8
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BTYPE           =   3
         TX              =   "&Xu t sang dπng kh∏c"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   16711680
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmShowReport_DeleteItems.frx":0078
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
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
         Left            =   60
         TabIndex        =   7
         Text            =   "cboZoom"
         Top             =   180
         Width           =   1575
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
         Left            =   4140
         Picture         =   "frmShowReport_DeleteItems.frx":0094
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
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
         Left            =   3720
         Picture         =   "frmShowReport_DeleteItems.frx":03D6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
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
         Left            =   2400
         Picture         =   "frmShowReport_DeleteItems.frx":0718
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   405
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
         Left            =   1980
         Picture         =   "frmShowReport_DeleteItems.frx":0A5A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   405
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
         Left            =   2820
         TabIndex        =   2
         Top             =   195
         Width           =   855
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   510
         Left            =   4740
         Picture         =   "frmShowReport_DeleteItems.frx":0D9C
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   195
         Width           =   885
      End
   End
   Begin CRVIEWERLibCtl.CRViewer crvReport 
      CausesValidation=   0   'False
      Height          =   7695
      Left            =   240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   960
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
Attribute VB_Name = "frmShowReport_DeleteItems"
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

Public Property Let Report(ByVal vNewValue As Variant)
    Set CRReport = vNewValue
End Property

Private Sub CboFilter_Change()
    Call View_Report
End Sub

Private Sub CboFilter_Click()
    Call View_Report
End Sub

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

Private Sub Form_Activate()
    If isLoad = True Then Exit Sub
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
CboFilter.ListIndex = 0
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
' Mo+? database vÌ du. kËm theo VB6 hoa(.c Access
sTapTinExcel = "C:\Customers.XLS"
Dim strSql As String
strSql = "SELECT Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price,Avg(Invoice_Itemized.PricePer)*Sum(Invoice_Itemized.Quantity)as Amount, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime" & _
          " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number " & _
          " WHERE Left([Invoice_Totals].[DateTime],8)>='" & FromDate & "' And Left([Invoice_Totals].[DateTime],8)<='" & ToDate & "'" & _
          " GROUP BY Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer"

Set rsCustomer = OpenCriticalTable(strSql, cnData)
Dim xlApp As Object ' Kho+?i do^.ng Excel, nhung khÙng hie^?n thi.
Set xlApp = CreateObject("Excel.Application")

Dim xlBook As Object
Set xlBook = xlApp.Workbooks.Add ' ThÍm workbook mo+'i

Dim xlSheet As Object
Set xlSheet = xlBook.Worksheets(1) ' L‡m vie^.c vo+'i Sheet1 trong Excel

xlSheet.Cells.Font.name = ".vnArial" ' Font cho to‡n bo^. sheet
xlSheet.Range("A1:E1").Select ' TiÍu ?e^` chie^'m 11 co^.t
With xlApp.Selection
.Font.Size = 14
.Font.Bold = True
.RowHeight = .Font.Size * 1.4
End With

With xlSheet
    .Columns.EntireColumn.AutoFit ' Chi?nh kÌch thu+o+'c c·c co^.t cho kho+'p du+~ lie^.u
    
    .Rows("1:1").Insert Shift:=xlDown ' ChËn thÍm dÚng tu+.a
    .Range("A1").FormulaR1C1 = "B∏o c∏o chi ti’t"
    .Range("A2").FormulaR1C1 = "Tı ngµy:" & gfCONVERT_STRING_TO_DATE(FromDate) & "ß’n ngµy:" & gfCONVERT_STRING_TO_DATE(ToDate)
    .Range("A2:E2").Select
End With
' Chi?nh da.ng dÚng tu+.a
With xlSheet.Rows("1:1").Font
    .Size = 18
    .Bold = True
End With
' –o.c du+~ lie^.u tu+` table Customers, g·n v‡o c·c co^.t trong excel
Dim i As Integer
For i = 1 To rsCustomer.Fields.count ' TÍn c·c co^.t ?u+o+.c in ?a^.m
    xlApp.ActiveSheet.Cells(3, i).Font.Bold = True
Next

NxtLine = 3 ' Ba('t ?a^`u in c·c tÍn cu?a co^.t du+~ lie^.u
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
        Else ' Ch?nh d?ng d? li?u ng‡y th·ng
            xlApp.ActiveSheet.Cells(NxtLine, lc + 1).Value = Format(rsCustomer.Fields(lc), "dd/mm/yy")
        End If
    Next
    rsCustomer.MoveNext
    NxtLine = NxtLine + 1
Loop

' Merge 11 Ù ?e^? dÚng tu+.a chu+'a he^'t trong 11 Ù n‡y
xlSheet.Range("A1:K1").Select
xlApp.Selection.Merge
' Ghi lÍn dia
xlBook.SaveAs FileName:=sTapTinExcel
xlBook.Saved = True ' Ghi ta^.p tin th‡nh cÙng
xlApp.Quit ' –Ûng Excel

rsCustomer.Close ' –Ûng table
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

Public Sub View_Report()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL As String
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Select Case CboFilter.ListIndex
        Case 0
            SQL = "SELECT DISTINCT Items_Deleted.Sec_ID, Items_Deleted.PrintCount,  items_Deleted.Invoice_Num as  Invoice_No, Items_Deleted.Table_ID, Items_Deleted.Cashier_ID, Items_Deleted.PluNo, Items_Deleted.Quantity, Items_Deleted.Price, Items_Deleted.Quantity*Items_Deleted.Price AS Amount, Left([DateTime],8) AS DateInvoice, Items_Deleted.Ordered, Items_Deleted.Reason, Right([DateTime],8) AS TimeInvoice, Inventory.ItemName, Inventory.Unit" & _
                  " FROM Inventory INNER JOIN Items_Deleted ON Inventory.ItemNum = Items_Deleted.PluNo" & _
                  " Where Left([DateTime],8)>='" & FromDate & "' and Left([DateTime],8)<='" & ToDate & "'"
        Case 1
            SQL = "SELECT DISTINCT Items_Deleted.Sec_ID, Items_Deleted.PrintCount,  items_Deleted.Invoice_Num as Invoice_No, Items_Deleted.Table_ID, Items_Deleted.Cashier_ID, Items_Deleted.PluNo, Items_Deleted.Quantity, Items_Deleted.Price, Items_Deleted.Quantity*Items_Deleted.Price AS Amount, Left([DateTime],8) AS DateInvoice, Items_Deleted.Ordered, Items_Deleted.Reason, Right([DateTime],8) AS TimeInvoice, Inventory.ItemName, Inventory.Unit" & _
                  " FROM Inventory INNER JOIN Items_Deleted ON Inventory.ItemNum = Items_Deleted.PluNo" & _
                  " Where Items_Deleted.Ordered=false and Left([DateTime],8)>='" & FromDate & "' and Left([DateTime],8)<='" & ToDate & "'"
        Case 2
            SQL = "SELECT DISTINCT Items_Deleted.Sec_ID, Items_Deleted.PrintCount,  items_Deleted.Invoice_Num as Invoice_No, Items_Deleted.Table_ID, Items_Deleted.Cashier_ID, Items_Deleted.PluNo, Items_Deleted.Quantity, Items_Deleted.Price, Items_Deleted.Quantity*Items_Deleted.Price AS Amount, Left([DateTime],8) AS DateInvoice, Items_Deleted.Ordered, Items_Deleted.Reason, Right([DateTime],8) AS TimeInvoice, Inventory.ItemName, Inventory.Unit" & _
                  " FROM Inventory INNER JOIN Items_Deleted ON Inventory.ItemNum = Items_Deleted.PluNo" & _
                  " Where Items_Deleted.Ordered=True and Left([DateTime],8)>='" & FromDate & "' and Left([DateTime],8)<='" & ToDate & "'"
        Case 3
            SQL = "SELECT DISTINCT Items_Deleted.Sec_ID, Items_Deleted.PrintCount,  items_Deleted.Invoice_Num as Invoice_No, Items_Deleted.Table_ID, Items_Deleted.Cashier_ID, Items_Deleted.PluNo, Items_Deleted.Quantity, Items_Deleted.Price, Items_Deleted.Quantity*Items_Deleted.Price AS Amount, Left([DateTime],8) AS DateInvoice, Items_Deleted.Ordered, Items_Deleted.Reason, Right([DateTime],8) AS TimeInvoice, Inventory.ItemName, Inventory.Unit" & _
                  " FROM Inventory INNER JOIN Items_Deleted ON Inventory.ItemNum = Items_Deleted.PluNo" & _
                  " Where Items_Deleted.PrintCount>0 and Left([DateTime],8)>='" & FromDate & "' and Left([DateTime],8)<='" & ToDate & "'"
        'Case 4
    End Select
    
    Set crDeleteItems = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crDeleteItems
        .Database.AddADOCommand cnData, cmd
        .txtserver.SetUnboundFieldSource "{ado.Sec_ID}"
        .txtTable.SetUnboundFieldSource "{ado.Table_ID}"
        .txtBill.SetUnboundFieldSource "{ado.Invoice_No}"
        .txtCashier.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtPluCode.SetUnboundFieldSource "{ado.PluNo}"
        .txtQty.SetUnboundFieldSource "{ado.Quantity}"
        .txtItemName.SetUnboundFieldSource "{ado.ItemName}"
        .txtUnit.SetUnboundFieldSource "{ado.Unit}"
        .txtCost.SetUnboundFieldSource "{ado.Price}"
        .txtAmt.SetUnboundFieldSource "{ado.Amount}"
        .txtReason.SetUnboundFieldSource "{ado.Reason}"
        .txtDate.SetUnboundFieldSource "{ado.DateInvoice}"
        .txtTime.SetUnboundFieldSource "{ado.TimeInvoice}"
        .printcount.SetUnboundFieldSource "{ado.PrintCount}"
        .blOrder.SetUnboundFieldSource "{ado.Ordered}"
        .txtFromDate.SetText gfCONVERT_STRING_TO_DATE(FromDate)
        .txtToDate.SetText gfCONVERT_STRING_TO_DATE(ToDate)
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
        With .Field11
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
    Set iReport = crDeleteItems
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
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name
End Sub


Public Property Let Let_Fromdate(ByVal vNewValue As Variant)
    FromDate = vNewValue
End Property

Public Property Let Let_Todate(ByVal vNewValue As Variant)
    ToDate = vNewValue
End Property

