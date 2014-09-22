VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmShowReport 
   Caption         =   "B¸o c¸o"
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
   Icon            =   "frmShowReport.frx":0000
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
      Height          =   705
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   13245
      TabIndex        =   1
      Top             =   0
      Width           =   13275
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
         TabIndex        =   10
         Text            =   "cboZoom"
         Top             =   60
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
         Picture         =   "frmShowReport.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   3720
         Picture         =   "frmShowReport.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   2400
         Picture         =   "frmShowReport.frx":0690
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   60
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
         Picture         =   "frmShowReport.frx":09D2
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
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
         TabIndex        =   5
         Top             =   75
         Width           =   855
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   510
         Left            =   4740
         Picture         =   "frmShowReport.frx":0D14
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   75
         Width           =   885
      End
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   585
         Left            =   7740
         TabIndex        =   2
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1032
         BTYPE           =   14
         TX              =   "&§ãng"
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
         BCOL            =   16578804
         BCOLO           =   16777152
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmShowReport.frx":13FE
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
         Height          =   525
         Left            =   5730
         TabIndex        =   3
         Top             =   60
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   926
         BTYPE           =   14
         TX              =   "XuÊt sang d¹ng kh¸c"
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
         BCOL            =   16578804
         BCOLO           =   16777152
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmShowReport.frx":141A
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
      Height          =   7695
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1020
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
Attribute VB_Name = "frmShowReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CRReport As New CRAXDDRT.Report
Dim iReport As New CRAXDDRT.Report
Dim TotalRptPage As Integer
Dim isLoad As Boolean
Dim Fromdate, Todate As String
Dim ReportNum As Integer

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
MsgBox Err.Number & Err.Description & Me.Name & " cmdExport_Click"
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
    MsgBox Err.Number & Err.Description & Me.Name & " Form_Load"
End Sub

Private Sub Xuat_Excel()
Dim rsCustomer As New ADODB.Recordset
Dim sTapTinExcel As String ' Ta^.p tin Excel ca^`n ke^'t xua^'t
Dim NxtLine, lc As Integer
' Mo+? database ví du. kèm theo VB6 hoa(.c Access
sTapTinExcel = "C:\Customers.XLS"
Dim strSQL As String
strSQL = "SELECT Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price,Avg(Invoice_Itemized.PricePer)*Sum(Invoice_Itemized.Quantity)as Amount, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime" & _
          " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number " & _
          " WHERE Left([Invoice_Totals].[DateTime],8)>='" & Fromdate & "' And Left([Invoice_Totals].[DateTime],8)<='" & Todate & "'" & _
          " GROUP BY Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer"

Set rsCustomer = OpenCriticalTable(strSQL, cnData)
Dim xlApp As Object ' Kho+?i do^.ng Excel, nhung không hie^?n thi.
Set xlApp = CreateObject("Excel.Application")

Dim xlBook As Object
Set xlBook = xlApp.Workbooks.Add ' Thêm workbook mo+'i

Dim xlSheet As Object
Set xlSheet = xlBook.Worksheets(1) ' Làm vie^.c vo+'i Sheet1 trong Excel

xlSheet.Cells.Font.Name = ".vnArial" ' Font cho toàn bo^. sheet
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
    .Range("A2").FormulaR1C1 = "Tõ ngµy:" & gfCONVERT_STRING_TO_DATE(Fromdate) & "§Õn ngµy:" & gfCONVERT_STRING_TO_DATE(Todate)
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
    xlApp.ActiveSheet.Cells(NxtLine, lc + 1).Value = rsCustomer.Fields(lc).Name
Next

NxtLine = 4 ' Ba('t ?a^`u in du+~ lie^.u
Do Until rsCustomer.EOF
    For lc = 0 To rsCustomer.Fields.count - 3
        xlApp.ActiveSheet.Cells(NxtLine, lc + 1).Value = rsCustomer.Fields(lc)
        xlApp.Columns.EntireColumn.AutoFit
        If rsCustomer.Fields.Item(lc).Name <> "DATE" Then
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
    Fromdate = vNewValue
End Property

Public Property Let Get_tDate(ByVal vNewValue As Variant)
    Todate = vNewValue
End Property

Public Property Let Report_Number(ByVal vNewValue As Variant)
    ReportNum = vNewValue
End Property
