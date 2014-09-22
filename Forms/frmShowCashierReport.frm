VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmShowCashierReport 
   Caption         =   "B¸o c¸o ca"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11370
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
   Icon            =   "frmShowCashierReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   14205
      TabIndex        =   0
      Top             =   120
      Width           =   14235
      Begin VB.ComboBox cboCashier 
         Height          =   345
         Left            =   6840
         TabIndex        =   13
         Top             =   240
         Width           =   3975
      End
      Begin VB.ComboBox cboGroup 
         Height          =   345
         Left            =   6840
         TabIndex        =   12
         Top             =   240
         Width           =   3975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nhãm hµng"
         Height          =   255
         Left            =   5280
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nh©n viªn"
         Height          =   255
         Left            =   5280
         TabIndex        =   10
         Top             =   120
         Width           =   1455
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
         Left            =   3360
         TabIndex        =   9
         Top             =   135
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
         Left            =   2520
         Picture         =   "frmShowCashierReport.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   120
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
         Left            =   2940
         Picture         =   "frmShowCashierReport.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
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
         Left            =   4260
         Picture         =   "frmShowCashierReport.frx":0690
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
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
         Left            =   4680
         Picture         =   "frmShowCashierReport.frx":09D2
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   13290
         Top             =   60
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
         Left            =   11400
         Picture         =   "frmShowCashierReport.frx":0D14
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   1035
      End
      Begin VB.ComboBox cboZoom 
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Text            =   "cboZoom"
         Top             =   120
         Width           =   2175
      End
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   585
         Left            =   12600
         TabIndex        =   3
         Top             =   30
         Width           =   1455
         _ExtentX        =   2566
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
         BCOL            =   16777215
         BCOLO           =   16777152
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmShowCashierReport.frx":13FE
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1740
      Width           =   12615
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
Attribute VB_Name = "frmShowCashierReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CRReport As New CRAXDDRT.Report
Dim PrinterName As String
Dim iReport As New CRAXDDRT.Report
Dim TotalPage As Integer
Dim isLoad As Boolean
Dim FromDate, ToDate As String
Dim TotalRptPage As Integer
Dim filterRep As Integer


Public Property Let Report(ByVal vNewValue As Variant)
    Set CRReport = vNewValue
End Property

Private Sub cboCashier_Change()
    ReportDone
End Sub

Private Sub cboCashier_Click()
    Call cboCashier_Change
End Sub

Private Sub cboGroup_Change()
    ReportDone
End Sub

Private Sub cboGroup_Click()
    Call cboGroup_Change
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
    'iReport.PrintOut False
    If cboGroup.ListIndex = 0 Then
        iReport.SelectPrinter GetSettingStr("Receip", "Receipt_DeviceName", True, myIniFile), GetSettingStr("Report", "Report_DeviceName", True, myIniFile), Printer.Port
        iReport.PrintOut False
    Else
        myPrint iReport, crvReport.GetCurrentPageNumber, TotalRptPage
    End If
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
    Set CRReport = Nothing
    
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub
Private Sub Form_Load()
On Error GoTo Handle
'PrinterName = GetSettingStr("Receip", "Receipt_DeviceName", True, myIniFile)
    isLoad = False
    Option2.Value = True
'Set iReport = CRReport
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
    Call AddGroup
    Call AddCashier
    cboZoom.ListIndex = 0
    cboGroup.ListIndex = 0
    ReportDone
    Me.WindowState = 2
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Public Sub AddGroup()
    On Error GoTo Handle
        Dim rsGroup As New ADODB.Recordset
        Set rsGroup = Open_Table(cnData, "Departments")
        With cboGroup
            .Clear
            .AddItem "TÊt c¶"
            Do While Not rsGroup.EOF
                .AddItem rsGroup.Fields("Description")
            rsGroup.MoveNext
            Loop
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Public Sub AddCashier()
    On Error GoTo Handle
        Dim rsCashier As New ADODB.Recordset
        Set rsCashier = LoadPasswordData
        With cboCashier
            .Clear
            .AddItem "TÊt c¶"
            Do While Not rsCashier.EOF
                .AddItem rsCashier.Fields("ID")
            rsCashier.MoveNext
            Loop
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub
Private Sub ReportDone()
On Error GoTo errHdl
    Dim cmd As New ADODB.Command
    Dim SQL, SQLSort As String
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    
Select Case filterRep
    Case 0: SQLSort = " Order by Invoice_Itemized.ItemNum  ASC"
    Case 1: SQLSort = " Order by Invoice_Itemized.DiffItemName  ASC"
    Case 2: SQLSort = " Order by sum(Invoice_Itemized.Quantity)  DESC"
    Case Else: SQLSort = " Order by Invoice_Itemized.ItemNum  ASC"
End Select
    
    'Khong danh cho karaoke
    If Option2.Value = True Then
        Select Case cboGroup.ListIndex
            Case 0:
                SQL = "SELECT Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime, Inventory.Dept_ID" & _
                     " FROM Invoice_Totals INNER JOIN (Inventory INNER JOIN Invoice_Itemized ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number" & _
                      " WHERE Left([Invoice_Totals].[DateTime],8)>='" & FromDate & "' And Left([Invoice_Totals].[DateTime],8)<='" & ToDate & "'" & _
                      " GROUP BY Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer, Inventory.Dept_ID" & SQLSort
            Case Else
                    SQL = "SELECT Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime, Inventory.Dept_ID" & _
                     " FROM Invoice_Totals INNER JOIN (Inventory INNER JOIN Invoice_Itemized ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number" & _
                      " WHERE Left([Invoice_Totals].[DateTime],8)>='" & FromDate & "' And Left([Invoice_Totals].[DateTime],8)<='" & ToDate & "' and Inventory.Dept_ID='" & Format(cboGroup.ListIndex, "000") & "'" & _
                      " GROUP BY Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer, Inventory.Dept_ID" & SQLSort
            
        End Select
    ElseIf Option1.Value = True Then
        If cboCashier.ListIndex = 0 Then
        
        SQL = "SELECT Invoice_Totals.Cashier_ID,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime" & _
             " FROM Invoice_Totals INNER JOIN (Inventory INNER JOIN Invoice_Itemized ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number" & _
              " WHERE Left([Invoice_Totals].[DateTime],8)>='" & FromDate & "' And Left([Invoice_Totals].[DateTime],8)<='" & ToDate & "'" & _
              " GROUP BY Invoice_Totals.Cashier_ID,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer" & SQLSort
        Else
            SQL = "SELECT Invoice_Totals.Cashier_ID,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime" & _
             " FROM Invoice_Totals INNER JOIN (Inventory INNER JOIN Invoice_Itemized ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number" & _
              " WHERE Left([Invoice_Totals].[DateTime],8)>='" & FromDate & "' And Left([Invoice_Totals].[DateTime],8)<='" & ToDate & "' and Invoice_Totals.Cashier_ID='" & cboCashier.Text & "'" & _
              " GROUP BY Invoice_Totals.Cashier_ID,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer" & SQLSort
        End If
    End If
    
    Set crDetail80 = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crDetail80
        .Database.AddADOCommand cnData, cmd
        .txtPluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .txtPluName.SetUnboundFieldSource "{ado.DiffItemName}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtCost.SetUnboundFieldSource "{ado.Price}"

        .txtFromDate.SetText gfCONVERT_STRING_TO_DATE(FromDate)
        .txtToDate.SetText gfCONVERT_STRING_TO_DATE(ToDate)
        .Section1.Suppress = True
        .Section3.Suppress = True
        .Section5.Suppress = True
        .Section4.Suppress = True
        .Section2.Suppress = True
        
        'canh le
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign
        
        If Option2.Value = True Then
            .Section13.Suppress = True
            .Section5.Suppress = False
            .txtGroup.SetUnboundFieldSource "{ado.Dept_ID}"
        Else
            .Section13.Suppress = False
            .Section5.Suppress = True
            .txtCashier.SetUnboundFieldSource "{ado.Cashier_ID}"
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
    
    Set iReport = crDetail80
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

Public Property Let Let_Fromdate(ByVal vNewValue As Variant)
    FromDate = vNewValue
End Property

Public Property Let Let_Todate(ByVal vNewValue As Variant)
    ToDate = vNewValue
End Property

Private Sub Option2_Click()
    cboCashier.Visible = False
    cboGroup.Visible = True
End Sub

Private Sub Option1_Click()
    cboCashier.Visible = True
    cboGroup.Visible = False
    cboCashier.ListIndex = 0
End Sub



Public Property Let filter_report(ByVal vNewValue As Variant)
filterRep = vNewValue
End Property
