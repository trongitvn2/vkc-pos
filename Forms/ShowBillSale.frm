VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmShowBillSale 
   Caption         =   "In Bill"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ShowBillSale.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
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
      ScaleWidth      =   10125
      TabIndex        =   0
      Top             =   0
      Width           =   10155
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   9090
         Top             =   60
      End
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   585
         Left            =   210
         TabIndex        =   4
         Top             =   30
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1032
         BTYPE           =   14
         TX              =   "&§ãng"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777152
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "ShowBillSale.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.CommandButton cmdPrint 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3690
         Picture         =   "ShowBillSale.frx":0028
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   1035
      End
      Begin VB.ComboBox cboZoom 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1800
         TabIndex        =   1
         Text            =   "cboZoom"
         Top             =   120
         Width           =   1575
      End
   End
   Begin CRVIEWERLibCtl.CRViewer crvReport 
      CausesValidation=   0   'False
      Height          =   7695
      Left            =   240
      TabIndex        =   3
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
Attribute VB_Name = "frmShowBillSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iReport As New CRAXDDRT.Report
Dim cmd As New ADODB.Command
Dim DescArr() As String
Dim isLoading As Boolean
Dim Document As CRAXDDRT.Report
Dim BillNO As Double
Dim i As Integer
Dim prt As Printer
Dim PrinterName As String

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
    Set cmd = Nothing
    Set iReport = Nothing
    Unload Me
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdPrint_Click()
On Error GoTo errHdl
    iReport.SelectPrinter GetSettingStr("Receip", "Receipt_DeviceName", True, myIniFile), GetSettingStr("Report", "Report_DeviceName", True, myIniFile), Printer.Port
    iReport.PrintOut True
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Activate()
    If isLoading = True Then Exit Sub
    isLoading = True
End Sub


Private Sub Form_Load()
On Error GoTo Handle
    isLoading = False
    Dim str As String
    DescArr = LoadLanguage(LngFile, "#02:005:")
    ReportDone
    If ArrayFlag(SF(0), 8) = 1 Then
        With frmSelectPrint
            .Show vbModal
            PrinterName = .LetPrinter
        End With
    Else
        If ArrayFlag(SF(6), 5) = 1 Then
            PrinterName = Get_Printer(Sec_ID)
        Else
            PrinterName = GetSettingStr("Receipt", "Receipt_DeviceName", True, myIniFile)
        End If
    End If
    iReport.SelectPrinter True, PrinterName, Printer.Port
    iReport.PrintOut False
    If ArrayFlag(SF(0), 2) = 1 Then
        iReport.PrintOut False
    End If
    If ArrayFlag(SF(6), 6) = 1 Then Call OpenPrinterCashDraw(PrinterName)
    Unload Me
Handle:
Exit Sub
'errHdl:
'    MsgBox Err.Number & " - " & Err.Description
End Sub


Private Sub ReportDone()
On Error Resume Next
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
    Dim ReceiptReport As CRAXDDRT.Report
    Dim DescArrPayment() As String
    DescArrPayment = LoadLanguage(LngFile, "#02:003:")

    If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    If ArrayFlag(SF(0), 5) = 0 Then
        If ArrayFlag(SF(6), 2) = 0 Then
        SQL = " SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.CustNum, Invoice_Totals.Discount," & _
            " Invoice_Totals.Total_Price, Left(Invoice_Totals_Notes.OpenTime,8) AS DateIn, substring(Invoice_Totals_Notes.OpenTime,9,8) AS TimeIn,Left(Invoice_Totals_Notes.ClosingTime,8) AS DateOut, substring(Invoice_Totals_Notes.ClosingTime,9,8) AS TimeOut, " & _
            " Invoice_Totals.Service_Charge, Invoice_Totals.VATFee, Invoice_Totals.Adjustment1," & _
            " Invoice_Totals.Adj1Rate,Invoice_Totals.Personals, Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate," & _
            " Invoice_Totals.Adjustment3,Invoice_Totals.Adj3Rate, Invoice_Totals.Adjustment4,Invoice_Totals.Adj4Rate, Invoice_Totals.AddMoney," & _
            " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount," & _
            " Invoice_Totals.Grand_Total," & _
            " Invoice_Totals.Amt_Tendered, Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, " & _
            " Invoice_Totals.OrderMan, Invoice_Totals.Station_ID, Invoice_Totals.Payment_Method,Invoice_Totals.Tax_Rate_ID, Invoice_Totals.Reserve," & _
            " Invoice_Itemized.ItemNum,Invoice_Itemized.LineNum,Invoice_Itemized.Line_Disc_Desc,Invoice_Itemized.LineDisc, Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer, Sum(Invoice_Itemized.Amt) AS Amt, Invoice_Itemized.DiffItemName, Invoice_Totals.Orig_OnHoldID" & _
            " FROM ((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN (Inventory INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID) ON Invoice_Itemized.ItemNum = Inventory.ItemNum) INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number " & _
            " Where (((Invoice_Totals.Invoice_Number) = " & BillNO & "))" & _
            " GROUP BY Invoice_Totals.Invoice_Number, Invoice_Totals.CustNum, Invoice_Totals.Discount," & _
            " Invoice_Totals.Tax_Rate_ID, Invoice_Totals.Total_Price, " & _
            " Invoice_Totals.Service_Charge, Invoice_Totals.VATFee, Invoice_Totals.Adjustment1,Invoice_Totals.Adj1Rate," & _
            " Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate, Invoice_Totals.Adjustment3,Invoice_Totals.Adj3Rate,Invoice_Totals.Adj4Rate," & _
            " Invoice_Totals.Adjustment4, Invoice_Totals.AddMoney, Invoice_Totals.OrderMan," & _
            " Left(Invoice_Totals_Notes.OpenTime,8),substring(Invoice_Totals_Notes.OpenTime,9,8), Left(Invoice_Totals_Notes.ClosingTime,8),substring(Invoice_Totals_Notes.ClosingTime,9,8),Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered,Invoice_Totals.Reserve," & _
            " Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID," & _
            " Invoice_Totals.Payment_Method,Invoice_Totals.Personals, Invoice_Itemized.ItemNum, Invoice_Itemized.PricePer," & _
            " Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc,Invoice_Itemized.LineNum,Invoice_Itemized.Line_Disc_Desc," & _
            " Invoice_Totals.Orig_OnHoldID, Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate," & _
            " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount" & _
            " ORDER BY Invoice_Itemized.LineNum Desc"
        Else
        SQL = " SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.CustNum, Invoice_Totals.Discount," & _
            " Invoice_Totals.Total_Price, Left(Invoice_Totals_Notes.OpenTime,8) AS DateIn, substring(Invoice_Totals_Notes.OpenTime,9,8) AS TimeIn, Left(Invoice_Totals_Notes.ClosingTime,8) AS DateOut, substring(Invoice_Totals_Notes.ClosingTime,9,8) AS TimeOut," & _
            " Invoice_Totals.Service_Charge, Invoice_Totals.VATFee, Invoice_Totals.Adjustment1,Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate," & _
            " Invoice_Totals.Adj1Rate,Invoice_Totals.Personals, Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate, " & _
            " Invoice_Totals.Adjustment3, Invoice_Totals.Adj3Rate,Invoice_Totals.Adjustment4,Invoice_Totals.Adj4Rate, Invoice_Totals.AddMoney,Invoice_Totals.Reserve," & _
             " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount," & _
            " Invoice_Totals.Grand_Total," & _
            " Invoice_Totals.Amt_Tendered, Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, " & _
            " Invoice_Totals.OrderMan, Invoice_Totals.Station_ID, Invoice_Totals.Payment_Method,Invoice_Totals.Tax_Rate_ID, " & _
            " Invoice_Itemized.ItemNum,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc, Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer, Sum(Invoice_Itemized.Amt) AS Amt, Invoice_Itemized.DiffItemName, Invoice_Totals.Orig_OnHoldID" & _
            " FROM ((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN (Inventory INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID) ON Invoice_Itemized.ItemNum = Inventory.ItemNum) INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number " & _
            " Where (((Invoice_Totals.Invoice_Number) = " & BillNO & "))" & _
            " GROUP BY Invoice_Totals.Invoice_Number, Invoice_Totals.CustNum, Invoice_Totals.Discount," & _
            " Invoice_Totals.Total_Price,  " & _
            " Invoice_Totals.Service_Charge, Invoice_Totals.VATFee, Invoice_Totals.Adjustment1,Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate," & _
            " Invoice_Totals.Adj1Rate, Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate,Invoice_Totals.Reserve," & _
            " Invoice_Totals.Adjustment3,Invoice_Totals.Adj3Rate, Invoice_Totals.Adjustment4,Invoice_Totals.Adj4Rate, Invoice_Totals.AddMoney,Invoice_Totals.Tax_Rate_ID," & _
            " Invoice_Totals.OrderMan,Invoice_Totals.Personals, substring(Invoice_Totals_Notes.ClosingTime,9,8), Invoice_Totals.Grand_Total," & _
            " Invoice_Totals.Amt_Tendered, Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID," & _
            " Invoice_Totals.Station_ID, Invoice_Totals.Payment_Method, Invoice_Itemized.ItemNum," & _
            " Invoice_Itemized.PricePer, Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc," & _
            " Invoice_Totals.Orig_OnHoldID, Left(Invoice_Totals_Notes.OpenTime,8),substring(Invoice_Totals_Notes.OpenTime,9,8), Left(Invoice_Totals_Notes.ClosingTime,8),substring(Invoice_Totals_Notes.ClosingTime,9,8)," & _
             " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount" & _
            " ORDER BY Invoice_Itemized.ItemNum ASC"
        End If
    Else
        If ArrayFlag(SF(6), 1) = 0 Then

            SQL = "SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.CustNum," & _
                " Invoice_Totals.Discount, Invoice_Totals.Total_Price,Left(Invoice_Totals_Notes.OpenTime,8) AS DateIn, substring(Invoice_Totals_Notes.OpenTime,9,8) AS TimeIn,Left(Invoice_Totals_Notes.ClosingTime,8) AS DateOut, substring(Invoice_Totals_Notes.ClosingTime,9,8) AS TimeOut," & _
                " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adjustment1,Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate," & _
                " Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate,Invoice_Totals.Adjustment3,Invoice_Totals.Adj3Rate,Invoice_Totals.Tax_Rate_ID," & _
                " Invoice_Totals.Adjustment4,Invoice_Totals.Adj4Rate,Invoice_Totals.AddMoney," & _
                " Invoice_Totals.Grand_Total,Invoice_Totals.Personals, Invoice_Totals.Amt_Tendered,Invoice_Totals.Reserve," & _
               " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount," & _
                " Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, Invoice_Totals.OrderMan," & _
                " Invoice_Totals.Station_ID,Invoice_Totals.Payment_Method,Invoice_Itemized.ItemNum, " & _
                " Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer,Invoice_Itemized.LineNum,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc," & _
                " sum(Invoice_Itemized.Amt) as Amt, " & _
                " Invoice_Itemized.DiffItemName ,Invoice_Totals.Orig_OnHoldID,MainGroup.GroupNo " & _
                " FROM((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN " & _
                " (Inventory INNER JOIN (Departments INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo) ON Inventory.Dept_ID = Departments.Dept_ID) ON Invoice_Itemized.ItemNum = Inventory.ItemNum)" & _
                " INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number " & _
                " Where Invoice_Totals.Invoice_Number=" & BillNO & _
                " group by Invoice_Itemized.LineNum,Invoice_Itemized.ItemNum,Invoice_Totals.Invoice_Number," & _
                " Invoice_Totals.CustNum,Invoice_Totals.Discount," & _
                " Invoice_Totals.Total_Price,Invoice_Totals.Grand_Total,Invoice_Totals.Amt_Tendered,Invoice_Totals.Amt_Change,substring(Invoice_Totals_Notes.ClosingTime,9,8)," & _
                " Invoice_Totals.Cashier_ID, Invoice_Totals.OrderMan, Invoice_Totals.Station_ID,Invoice_Totals.Reserve, Left(Invoice_Totals_Notes.OpenTime,8),substring(Invoice_Totals_Notes.OpenTime,9,8), Left(Invoice_Totals_Notes.ClosingTime,8),substring(Invoice_Totals_Notes.ClosingTime,9,8)," & _
                " Invoice_Itemized.PricePer,Invoice_Totals.Personals, Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc ,Invoice_Totals.Orig_OnHoldID,Invoice_Totals.Payment_Method, " & _
                " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Tax_Rate_ID,Invoice_Totals.Adjustment1,Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate,Invoice_Totals.Adjustment3," & _
                " Invoice_Totals.Adj3Rate,Invoice_Totals.Adjustment4,Invoice_Totals.Adj4Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate,Invoice_Totals.AddMoney,MainGroup.GroupNo, " & _
                " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount" & _
                " order by Invoice_Itemized.LineNum Desc"
        Else
            SQL = "SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.CustNum," & _
                " Invoice_Totals.Discount, Invoice_Totals.Total_Price,Left(Invoice_Totals_Notes.OpenTime,8) AS DateIn, substring(Invoice_Totals_Notes.OpenTime,9,8) AS TimeIn,Left(Invoice_Totals_Notes.ClosingTime,8) AS DateOut, substring(Invoice_Totals_Notes.ClosingTime,9,8) AS TimeOut," & _
                " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adjustment1,Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate," & _
                " Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate,Invoice_Totals.Adjustment3,Invoice_Totals.Reserve," & _
                " Invoice_Totals.Adj3Rate,Invoice_Totals.Adjustment4,Invoice_Totals.Adj4Rate,Invoice_Totals.AddMoney,Invoice_Totals.Grand_Total,Invoice_Totals.Personals, Invoice_Totals.Amt_Tendered," & _
                " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount," & _
                " Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, Invoice_Totals.OrderMan,Invoice_Totals.Tax_Rate_ID," & _
                " Invoice_Totals.Station_ID,Invoice_Totals.Payment_Method,Invoice_Itemized.ItemNum, " & _
                " Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc," & _
                " sum(Invoice_Itemized.Amt) as Amt, " & _
                " Invoice_Itemized.DiffItemName ,Invoice_Totals.Orig_OnHoldID,MainGroup.GroupNo " & _
                " FROM((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN " & _
                " (Inventory INNER JOIN (Departments INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo) ON Inventory.Dept_ID = Departments.Dept_ID) ON Invoice_Itemized.ItemNum = Inventory.ItemNum)" & _
                " INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number " & _
                " Where Invoice_Totals.Invoice_Number=" & BillNO & _
                " group by Invoice_Itemized.ItemNum,Invoice_Totals.Invoice_Number,Invoice_Totals.CustNum,Invoice_Totals.Discount," & _
                " Invoice_Totals.Total_Price,Invoice_Totals.Grand_Total, Left(Invoice_Totals_Notes.OpenTime,8),substring(Invoice_Totals_Notes.OpenTime,9,8), Left(Invoice_Totals_Notes.ClosingTime,8),substring(Invoice_Totals_Notes.ClosingTime,9,8)," & _
                " Invoice_Totals.Amt_Tendered,Invoice_Totals.Amt_Change,substring(Invoice_Totals_Notes.ClosingTime,9,8)," & _
                " Invoice_Totals.Cashier_ID,Invoice_Totals.Personals, Invoice_Totals.OrderMan, Invoice_Totals.Station_ID,Invoice_Totals.Tax_Rate_ID," & _
                " Invoice_Itemized.PricePer, Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc ," & _
                " Invoice_Itemized.Line_Disc_Desc,Invoice_Totals.Orig_OnHoldID,Invoice_Totals.Payment_Method, Invoice_Totals.Reserve," & _
                " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adjustment1,Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate,Invoice_Totals.Adjustment3," & _
                " Invoice_Totals.Adj3Rate,Invoice_Totals.Adjustment4,Invoice_Totals.Adj4Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate,Invoice_Totals.AddMoney,MainGroup.GroupNo, " & _
                " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount" & _
                " order by Invoice_Itemized.ItemNum ASC"
        End If
   End If
    Dim CRXReportField As CRAXDDRT.DatabaseFieldDefinition
     
    Set crSaleBill = Nothing
    Set crSaleBill58 = Nothing
    Set crSaleBill75 = Nothing
     Set crSaleBillA5 = Nothing
    If ReceiptType = "80" Then
        Set ReceiptReport = crSaleBill
    ElseIf ReceiptType = "58" Then
        Set ReceiptReport = crSaleBill58
    ElseIf ReceiptType = "75" Then
        Set ReceiptReport = crSaleBill75
    Else
        Set ReceiptReport = crSaleBillA5
    End If
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With ReceiptReport
        .Database.AddADOCommand cnData, cmd
        .txtPluCode.SetUnboundFieldSource "{ado.DiffItemName}"
        .txtPluName.SetUnboundFieldSource "{ado.ItemNum}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtCost.SetUnboundFieldSource "{ado.PricePer}"
        .txtAmt.SetUnboundFieldSource "{ado.Amt}"
        .LineDisc.SetUnboundFieldSource "{ado.LineDisc}"
        .txtBillNo.SetUnboundFieldSource "{ado.Invoice_Number}"
        .txtCashier.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtNVPV.SetUnboundFieldSource "{ado.OrderMan}"
        .txtPayment.SetUnboundFieldSource "{ado.Amt_Tendered}"
        .txtChange.SetUnboundFieldSource "{ado.Amt_Change}"
        .txtDiscount.SetUnboundFieldSource "{ado.Discount}"
        .txtTable.SetUnboundFieldSource "{ado.Orig_OnHoldID}"
        .txtserver.SetUnboundFieldSource "{ado.Station_ID}"
        .txtMethod.SetUnboundFieldSource "{ado.Payment_Method}"
        '.TxtTotal.SetUnboundFieldSource "{ado.Grand_Total}"
        .txtServ.SetUnboundFieldSource "{ado.Service_Charge}"
        .txtVAT.SetUnboundFieldSource "{ado.VATFee}"
        .txtAdj1.SetUnboundFieldSource "{ado.Adjustment1}"
        .txtAdj1Rate.SetUnboundFieldSource "{ado.Adj1Rate}"
        .txtAdj2.SetUnboundFieldSource "{ado.Adjustment2}"
        .txtAdj2Rate.SetUnboundFieldSource "{ado.Adj2Rate}"
        .txtAdj3.SetUnboundFieldSource "{ado.Adjustment3}"
        .txtAdj3Rate.SetUnboundFieldSource "{ado.Adj3Rate}"
        .txtAdj4.SetUnboundFieldSource "{ado.Adjustment4}"
        .txtAdj4Rate.SetUnboundFieldSource "{ado.Adj4Rate}"
        
        .txtAdj5.SetUnboundFieldSource "{ado.Adjustment5}"
        .txtAdj5Rate.SetUnboundFieldSource "{ado.Adj5Rate}"

        .txtAdj6.SetUnboundFieldSource "{ado.Adjustment6}"
        .txtAdj6Rate.SetUnboundFieldSource "{ado.Adj6Rate}"

        .txtMoney.SetUnboundFieldSource "{ado.AddMoney}"
        .txtDateIn.SetUnboundFieldSource "{ado.DateIn}"
        .txtTimeIn.SetUnboundFieldSource "{ado.TimeIn}"
        .txtDateOut.SetUnboundFieldSource "{ado.DateOut}"
        .txtTimeOut.SetUnboundFieldSource "{ado.TimeOut}"
        .txtCustomer.SetUnboundFieldSource "{ado.CustNum}"
        .txtMixmatch.SetUnboundFieldSource "{ado.Tax_Rate_ID}"
        .txtsokhach.SetUnboundFieldSource "{ado.Personals}"
        .txtLineDiscDesc.SetUnboundFieldSource "{ado.Line_Disc_Desc}"
        .txtCAPAYMENT.SetUnboundFieldSource "{ado.CA_Amount}"
        .txtOAPAYMENT.SetUnboundFieldSource "{ado.OA_Amount}"
        .txtROAPAYMENT.SetUnboundFieldSource "{ado.ROA_Amount}"
        .txtCCPAYMENT.SetUnboundFieldSource "{ado.CC_Amount}"
        .txtCTPAYMENT.SetUnboundFieldSource "{ado.CT_Amount}"
        .txtGCPAYMENT.SetUnboundFieldSource "{ado.GC_Amount}"
        .txtDatcoc.SetUnboundFieldSource "{ado.Reserve}"
        If ArrayFlag(SF(0), 5) = 1 Then
            .txtMaingroup.SetUnboundFieldSource "{ado.GroupNo}"
        End If
        
        .lblTitle.SetText DescArr(1)
        .lblBillNo.SetText DescArr(2)
        .lblTable.SetText DescArr(3)
        .lblItems.SetText DescArr(4)
        .lblQty.SetText DescArr(5)
        .lblPrice.SetText DescArr(6)
        .lblAmt.SetText DescArr(7)
        .lblTotal.SetText DescArr(8)
'        .lblDiscount.SetText DescArr(9)
        .lblTender.SetText DescArr(10)
        .lblChange.SetText DescArr(11)
        .lblRead.SetText DescArr(12)
        .lblCashier.SetText DescArr(13)
        .lblPhuthu.SetText DescArr(14)
        .lblTotal1.SetText DescArr(15)
        .lblServer.SetText DescArr(16)
        .lblIn.SetText DescArr(17)
        .lblOut.SetText DescArr(18)
        .lblCash.SetText DescArr(19)
        .lblOrder.SetText DescArr(20)
        .lblCustome.SetText DescArr(21)
        .lblSignal.SetText DescArr(22)
'        .lblAdj1.SetText DescArr(25)
'        .lblAdj2.SetText DescArr(26)
        .lblPhuphi.SetText DescArr(27)
        .lblVAT.SetText DescArr(29)
        '.lblPrintCount.SetText DescArr(30)
        .lblTotalItems.SetText DescArr(31)
        
        
        .lblCA.SetText DescArrPayment(3)
        .lblCT.SetText DescArrPayment(4)
        .lblGC.SetText DescArrPayment(5)
        .lblOA.SetText DescArrPayment(6)
        .lblCC.SetText DescArrPayment(7)
        .lblROA.SetText DescArrPayment(8)
        
        'canh le
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign
        
'        With .txtMoney
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtAdj1
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtAdj2
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtAdj3
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtAdj4
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtQty
'            .DecimalPlaces = DecimalQtyNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtTotalAmt
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtPayment
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtChange
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtAmtDist
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtCost
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtServAmt
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'
'        With .SumMaingroup
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtQty
'            .DecimalPlaces = DecimalQtyNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
    End With
    
    Set iReport = ReceiptReport
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
        .ShowFirstPage
        While .IsBusy
            DoEvents
        Wend
    End With
'Exit Sub
'errHdl:
'    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errHdl
    
    i = 0
    Set cmd = Nothing
    Set iReport = Nothing
    Set crSaleBill = Nothing
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub


Public Property Let GetBill(ByVal vNewValue As Variant)
    BillNO = vNewValue
End Property

