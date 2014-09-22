VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmShowStockB 
   Caption         =   "Bao cao kho B"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14385
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShowStockB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9210
   ScaleWidth      =   14385
   StartUpPosition =   3  'Windows Default
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
      ScaleWidth      =   14205
      TabIndex        =   0
      Top             =   0
      Width           =   14235
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
         Left            =   4740
         TabIndex        =   9
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
         Left            =   8820
         Picture         =   "frmShowStockB.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Left            =   8400
         Picture         =   "frmShowStockB.frx":034E
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Left            =   7080
         Picture         =   "frmShowStockB.frx":0690
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Left            =   6660
         Picture         =   "frmShowStockB.frx":09D2
         Style           =   1  'Graphical
         TabIndex        =   5
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
         Left            =   7500
         TabIndex        =   4
         Top             =   75
         Width           =   855
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   510
         Left            =   9420
         Picture         =   "frmShowStockB.frx":0D14
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   75
         Width           =   885
      End
      Begin VB.ComboBox cboStock 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "frmShowStockB.frx":13FE
         Left            =   120
         List            =   "frmShowStockB.frx":140B
         TabIndex        =   2
         Top             =   60
         Width           =   4455
      End
      Begin prjTouchScreen.MyButton cmdExport 
         Height          =   615
         Left            =   10440
         TabIndex        =   1
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "Xu t sang dπng kh∏c"
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
         MICON           =   "frmShowStockB.frx":1439
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   615
         Left            =   12120
         TabIndex        =   10
         Top             =   0
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "ß„ng"
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
         MICON           =   "frmShowStockB.frx":1455
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
      TabIndex        =   11
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
Attribute VB_Name = "frmShowStockB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CRReport As New CRAXDDRT.Report
Dim iReport As New CRAXDDRT.Report
Dim ReceiptReport As New CRAXDDRT.Report
Dim TotalRptPage As Integer
Dim Report_Number As Integer
Dim from_Date, To_Date As String
Dim Table_Month As String
Dim fLoad As Boolean

Public Property Let Report(ByVal vNewValue As Variant)
    Set CRReport = vNewValue
End Property

Private Sub cboStock_Change()
    Set iReport = Nothing
    Set crStock = Nothing
    If fLoad Then
        Select Case Report_Number
            Case 1:
                Call Tonkho_Report
            Case 2:
                Call Instock_Report
            Case 3:
                Call Outstock_Report
            Case 4:
                Call MoveInOutstock_Report
            Case 5:
                Call Tonkho_Report80
            Case 6:
                Call Outstock_Report80
            Case 7:
                MoveInOutstock_Report80
        End Select
    End If
End Sub

Private Sub cboStock_Click()
Call cboStock_Change
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
    If fLoad = True Then Exit Sub
    fLoad = True
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
'Call GetStock_cbo
'Set iReport = CRReport
Table_Month = Mid(To_Date, 5, 2) & Mid(To_Date, 3, 2)
cboStock.ListIndex = 0
fLoad = False
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
    Select Case Report_Number
        Case 1:
            Call Tonkho_Report
        Case 2:
            Call Instock_Report
        Case 3:
            Call Outstock_Report
        Case 4:
            Call MoveInOutstock_Report
        Case 5:
            Call Tonkho_Report80
        Case 6:
            Call Outstock_Report80
        Case 7:
            Call MoveInOutstock_Report80
    End Select
   
    Me.WindowState = 2
    txtPage.Text = crvReport.GetCurrentPageNumber & " / " & TotalRptPage
    fLoad = True
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Public Sub GetStock_cbo()
    On Error GoTo Handle
        Dim rsStock As New ADODB.Recordset
        Set rsStock = Open_Table(cnData, "Stock_List")
        cboStock.Clear
    Do While Not rsStock.EOF
        With cboStock
            .AddItem rsStock.Fields("Stock_Name")
        End With
    rsStock.MoveNext
    Loop
    cboStock.ListIndex = 0
    Exit Sub
Handle:
        MsgBox Err.Number & Err.Description & _
        Me.name & " GetStock_cbo"
End Sub

Public Property Let Report_ID(ByVal vNewValue As Variant)
    Report_Number = vNewValue
End Property

Public Sub Instock_Report()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL, sql1 As String
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Select Case cboStock.ListIndex
        Case 0
            SQL = "SELECT Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description, Sum(Inventory_InB" & Table_Month & ".Quantity) AS Qty, Avg(Inventory_InB" & Table_Month & ".CostPer) AS Cost, Sum(Inventory_InB" & Table_Month & ".Amount) AS Amt" & _
                 " FROM Instock_MasterB INNER JOIN Inventory_InB" & Table_Month & " ON Instock_MasterB.Doc_Number = Inventory_InB" & Table_Month & ".Doc_Number" & _
                 " WHERE (((Instock_MasterB.DateTime)>='" & from_Date & "' And (Instock_MasterB.DateTime)<='" & To_Date & "') AND ((Instock_MasterB.InOutType)<>'O'))" & _
                 " GROUP BY Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description"
                 
        Case 1
            SQL = "SELECT Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description, Sum(Inventory_InB" & Table_Month & ".Quantity) AS Qty, Avg(Inventory_InB" & Table_Month & ".CostPer) AS Cost, Sum(Inventory_InB" & Table_Month & ".Amount) AS Amt" & _
                 " FROM Instock_MasterB INNER JOIN Inventory_InB" & Table_Month & " ON Instock_MasterB.Doc_Number = Inventory_InB" & Table_Month & ".Doc_Number" & _
                 " WHERE (((Instock_MasterB.DateTime)>='" & from_Date & "' And (Instock_MasterB.DateTime)<='" & To_Date & "') AND ((Instock_MasterB.InOutType)<>'O')) and Stock_ID='01'" & _
                 " GROUP BY Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description"
                        
        Case 2
            SQL = "SELECT Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description, Sum(Inventory_InB" & Table_Month & ".Quantity) AS Qty, Avg(Inventory_InB" & Table_Month & ".CostPer) AS Cost, Sum(Inventory_InB" & Table_Month & ".Amount) AS Amt" & _
                 " FROM Instock_MasterB INNER JOIN Inventory_InB" & Table_Month & " ON Instock_MasterB.Doc_Number = Inventory_InB" & Table_Month & ".Doc_Number" & _
                 " WHERE (((Instock_MasterB.DateTime)>='" & from_Date & "' And (Instock_MasterB.DateTime)<='" & To_Date & "') AND ((Instock_MasterB.InOutType)<>'O')) and Stock_ID='02'" & _
                 " GROUP BY Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description"
                   
    End Select
    
    Set crStock = Nothing
    
    
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
        
    With crStock
        Select Case cboStock.ListIndex
                           
        End Select
        .Database.AddADOCommand cnData, cmd
        
        '.GroupA.SetUnboundFieldSource "{ado.GroupA}"
        .PluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .PluName.SetUnboundFieldSource "{ado.Description}"
        '.Unit.SetUnboundFieldSource "{ado.Unit}"
        .Qty.SetUnboundFieldSource "{ado.Qty}"
        .Price.SetUnboundFieldSource "{ado.cost}"
        .Amount.SetUnboundFieldSource "{ado.Amt}"
        .StockID.SetUnboundFieldSource "{ado.Stock_ID}"
        .txtFromDate.SetText gfCONVERT_STRING_TO_DATE(from_Date)
        .txtToDate.SetText gfCONVERT_STRING_TO_DATE(To_Date)
        .txtTitle.SetText "B∏o c∏o nhÀp kho"
        With .Qty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Price
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Amount
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        
        With .SumAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .GrandAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    Set iReport = crStock
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
MsgBox Err.Number & Err.Description & Me.name & " Instock_Report"
End Sub

Public Property Let Get_FromDate(ByVal vNewValue As Variant)
    from_Date = vNewValue
End Property

Public Property Let Get_ToDate(ByVal vNewValue As Variant)
    To_Date = vNewValue
End Property

'Bao cao xuat kho
Public Sub Outstock_Report()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL, sql1 As String
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Select Case cboStock.ListIndex
        Case 0
            SQL = "SELECT Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description, Sum(Inventory_InB" & Table_Month & ".Quantity) AS Qty, Avg(Inventory_InB" & Table_Month & ".CostPer) AS Cost, Sum(Inventory_InB" & Table_Month & ".Amount) AS Amt" & _
                 " FROM Instock_MasterB INNER JOIN Inventory_InB" & Table_Month & " ON Instock_MasterB.Doc_Number = Inventory_InB" & Table_Month & ".Doc_Number" & _
                 " WHERE (((Instock_MasterB.DateTime)>='" & from_Date & "' And (Instock_MasterB.DateTime)<='" & To_Date & "') AND ((Instock_MasterB.InOutType)='O'))" & _
                 " GROUP BY Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description"
                 
        Case 1
            SQL = "SELECT Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description, Sum(Inventory_InB" & Table_Month & ".Quantity) AS Qty, Avg(Inventory_InB" & Table_Month & ".CostPer) AS Cost, Sum(Inventory_InB" & Table_Month & ".Amount) AS Amt" & _
                 " FROM Instock_MasterB INNER JOIN Inventory_InB" & Table_Month & " ON Instock_MasterB.Doc_Number = Inventory_InB" & Table_Month & ".Doc_Number" & _
                 " WHERE (((Instock_MasterB.DateTime)>='" & from_Date & "' And (Instock_MasterB.DateTime)<='" & To_Date & "') AND ((Instock_MasterB.InOutType)='O')) and Stock_ID='01'" & _
                 " GROUP BY Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description"
                        
        Case 2
            SQL = "SELECT Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description, Sum(Inventory_InB" & Table_Month & ".Quantity) AS Qty, Avg(Inventory_InB" & Table_Month & ".CostPer) AS Cost, Sum(Inventory_InB" & Table_Month & ".Amount) AS Amt" & _
                 " FROM Instock_MasterB INNER JOIN Inventory_InB" & Table_Month & " ON Instock_MasterB.Doc_Number = Inventory_InB" & Table_Month & ".Doc_Number" & _
                 " WHERE (((Instock_MasterB.DateTime)>='" & from_Date & "' And (Instock_MasterB.DateTime)<='" & To_Date & "') AND ((Instock_MasterB.InOutType)='O')) and Stock_ID='02'" & _
                 " GROUP BY Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description"
                   
    End Select
    
    Set crStock = Nothing
    
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crStock
        Select Case cboStock.ListIndex
                           
        End Select
        .Database.AddADOCommand cnData, cmd
        
        '.GroupA.SetUnboundFieldSource "{ado.GroupA}"
        .PluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .PluName.SetUnboundFieldSource "{ado.Description}"
        '.Unit.SetUnboundFieldSource "{ado.Unit}"
        .Qty.SetUnboundFieldSource "{ado.Qty}"
        .Price.SetUnboundFieldSource "{ado.cost}"
'        .Amount.SetUnboundFieldSource "{ado.Amt}"
        .StockID.SetUnboundFieldSource "{ado.Stock_ID}"
        .txtFromDate.SetText gfCONVERT_STRING_TO_DATE(from_Date)
        .txtToDate.SetText gfCONVERT_STRING_TO_DATE(To_Date)
        .txtTitle.SetText "B∏o c∏o xu t kho"
        With .Qty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Price
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Amount
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        
        With .SumAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .GrandAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    Set iReport = crStock
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
MsgBox Err.Number & Err.Description & Me.name & " Outstock_Report"
End Sub
'Bao cao xuat kho 80
Public Sub Outstock_Report80()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL, sql1 As String
    Dim CRReport As CRAXDDRT.Report
    
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Select Case cboStock.ListIndex
        Case 0
            SQL = "SELECT Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description, Sum(Inventory_InB" & Table_Month & ".Quantity) AS Qty, Avg(Inventory_InB" & Table_Month & ".CostPer) AS Cost, Sum(Inventory_InB" & Table_Month & ".Amount) AS Amt" & _
                 " FROM Instock_MasterB INNER JOIN Inventory_InB" & Table_Month & " ON Instock_MasterB.Doc_Number = Inventory_InB" & Table_Month & ".Doc_Number" & _
                 " WHERE (((Instock_MasterB.DateTime)>='" & from_Date & "' And (Instock_MasterB.DateTime)<='" & To_Date & "') AND ((Instock_MasterB.InOutType)='O'))" & _
                 " GROUP BY Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description"
                 
        Case 1
            SQL = "SELECT Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description, Sum(Inventory_InB" & Table_Month & ".Quantity) AS Qty, Avg(Inventory_InB" & Table_Month & ".CostPer) AS Cost, Sum(Inventory_InB" & Table_Month & ".Amount) AS Amt" & _
                 " FROM Instock_MasterB INNER JOIN Inventory_InB" & Table_Month & " ON Instock_MasterB.Doc_Number = Inventory_InB" & Table_Month & ".Doc_Number" & _
                 " WHERE (((Instock_MasterB.DateTime)>='" & from_Date & "' And (Instock_MasterB.DateTime)<='" & To_Date & "') AND ((Instock_MasterB.InOutType)='O')) and Stock_ID='01'" & _
                 " GROUP BY Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description"
                        
        Case 2
            SQL = "SELECT Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description, Sum(Inventory_InB" & Table_Month & ".Quantity) AS Qty, Avg(Inventory_InB" & Table_Month & ".CostPer) AS Cost, Sum(Inventory_InB" & Table_Month & ".Amount) AS Amt" & _
                 " FROM Instock_MasterB INNER JOIN Inventory_InB" & Table_Month & " ON Instock_MasterB.Doc_Number = Inventory_InB" & Table_Month & ".Doc_Number" & _
                 " WHERE (((Instock_MasterB.DateTime)>='" & from_Date & "' And (Instock_MasterB.DateTime)<='" & To_Date & "') AND ((Instock_MasterB.InOutType)='O')) and Stock_ID='02'" & _
                 " GROUP BY Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description"
                   
    End Select
    
    Set crStock80 = Nothing
    Set crStock58 = Nothing
    
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    If ReceiptType = 80 Then
        Set CRReport = crStock80
    Else
        Set CRReport = crStock58
    End If
    
    With CRReport
    
        Select Case cboStock.ListIndex
                           
        End Select
        .Database.AddADOCommand cnData, cmd
        
        .PluName.SetUnboundFieldSource "{ado.Description}"
        .Qty.SetUnboundFieldSource "{ado.Qty}"
        .StockID.SetUnboundFieldSource "{ado.Stock_ID}"
        .txtFromDate.SetText gfCONVERT_STRING_TO_DATE(from_Date)
        .txtToDate.SetText gfCONVERT_STRING_TO_DATE(To_Date)
        .txtTitle.SetText "B∏o c∏o xu t kho"
        With .Qty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With

    End With
    Set iReport = CRReport
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
MsgBox Err.Number & Err.Description & Me.name & " Outstock_Report"
End Sub

'Bao cao ton kho
Public Sub Tonkho_Report()
On Error GoTo Handle
    Dim rsTem As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim SQL, sql1 As String
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Select Case cboStock.ListIndex
        Case 0
            SQL = "SELECT TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".ItemNum, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Description, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Unit, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Stock_ID, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Quantity, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".CostPer, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & ".Amount" & _
            " FROM TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00")
        Case 1
           SQL = "SELECT TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".ItemNum, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Description, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Unit, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Stock_ID, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Quantity, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".CostPer, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & ".Amount" & _
            " FROM TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            " Where TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Stock_ID='01'"
        Case 2
            SQL = "SELECT TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".ItemNum, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Description, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Unit, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Stock_ID, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Quantity, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".CostPer, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & ".Amount" & _
            " FROM TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            " Where TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Stock_ID='02'"
    End Select
    Set rsTem = OpenCriticalTable(SQL, cnData)
    If rsTem.RecordCount = 0 Then
        MsgBox "Kh´ng c„ d˜ li÷u "
        Exit Sub
    End If
    Set crStockInOut = Nothing
        
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crStockInOut
        .Database.AddADOCommand cnData, cmd
        '.GroupA.SetUnboundFieldSource "{ado.Dept}"
        .PluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .PluName.SetUnboundFieldSource "{ado.Description}"
        .Unit.SetUnboundFieldSource "{ado.Unit}"
        .Qty.SetUnboundFieldSource "{ado.Quantity}"
        .Price.SetUnboundFieldSource "{ado.CostPer}"
        '.Amount.SetUnboundFieldSource "{ado.Amt}"
        .txtTinhden.Suppress = False
        .txtTinhden.SetText "T›nh Æ’n ngµy:"
        .txtFromDate.SetText gfCONVERT_STRING_TO_DATE(To_Date)
        .txtToDate.Suppress = True
        .Text8.Suppress = True
        .Text3.Suppress = True
        .txtTitle.SetText "B∏o c∏o TÂn kho"
        With .Qty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Price
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Amount
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        
'        With .SumAmt
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'
'        End With
        With .GrandAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    Set iReport = crStockInOut
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
MsgBox Err.Number & Err.Description & Me.name & " Tonkho_Report"
End Sub
'Bao cao ton kho 80mm
Public Sub Tonkho_Report80()
On Error GoTo Handle
    Dim rsTem As New ADODB.Recordset
    Dim cmd As New ADODB.Command
    Dim SQL, sql1 As String
    Dim CRReport As CRAXDDRT.Report
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Select Case cboStock.ListIndex
        Case 0
            SQL = "SELECT TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".ItemNum, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Description, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Unit, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Stock_ID, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Quantity, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".CostPer, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & ".Amount" & _
            " FROM TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00")
        Case 1
           SQL = "SELECT TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".ItemNum, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Description, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Unit, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Stock_ID, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Quantity, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".CostPer, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & ".Amount" & _
            " FROM TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            " Where TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Stock_ID='01'"
        Case 2
            SQL = "SELECT TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".ItemNum, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Description, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Unit, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Stock_ID, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Quantity, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".CostPer, TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & ".Amount" & _
            " FROM TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            " Where TonB" & Mid(To_Date, 5, 2) & Format(Mid(To_Date, 3, 2), "00") & _
            ".Stock_ID='02'"
    End Select
    Set rsTem = OpenCriticalTable(SQL, cnData)
    If rsTem.RecordCount = 0 Then
        MsgBox "Kh´ng c„ d˜ li÷u "
        Exit Sub
    End If
    Set crStockInOut80 = Nothing
    Set crStockInOut58 = Nothing
    
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    If ReceiptType = 80 Then
        Set CRReport = crStockInOut80
    Else
        Set CRReport = crStockInOut58
    End If
    
    With CRReport
    
        .Database.AddADOCommand cnData, cmd
        .PluName.SetUnboundFieldSource "{ado.Description}"
        .Qty.SetUnboundFieldSource "{ado.Quantity}"
        .txtTinhden.Suppress = False
        .txtTinhden.SetText "T›nh Æ’n ngµy:"
        .txtFromDate.SetText gfCONVERT_STRING_TO_DATE(To_Date)
'        .txtTodate.Suppress = True
'        .Text8.Suppress = True
'        .Text3.Suppress = True
        .txtTitle.SetText "B∏o c∏o TÂn kho"
        With .Qty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With

    End With
    Set iReport = CRReport
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
MsgBox Err.Number & Err.Description & Me.name & " Tonkho_Report"
End Sub

'Bao cao xuat nhap ton kho
Public Sub MoveInOutstock_Report()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL, sql1 As String
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Select Case cboStock.ListIndex
        Case 0
            SQL = "SELECT Stock_ReportB.Stock_ID, Stock_ReportB.ItemCode, Stock_ReportB.ItemName," & _
            " Stock_ReportB.Unit, sum(Stock_ReportB.First_Qty) as qty,sum( Stock_ReportB.First_Amt) as First_Amt," & _
            " sum(Stock_ReportB.Instock) as In_Qty, sum(Stock_ReportB.In_Amt) as In_Amount," & _
            " sum(Stock_ReportB.OutStock) As Out_Qty,sum(Stock_ReportB.Out_Amt) as Out_Amt" & _
            " From Stock_ReportB" & _
            " GROUP BY Stock_ReportB.Stock_ID, Stock_ReportB.ItemCode, Stock_ReportB.ItemName, Stock_ReportB.Unit"
        Case 1
            SQL = "SELECT Stock_ReportB.Stock_ID, Stock_ReportB.ItemCode, Stock_ReportB.ItemName," & _
            " Stock_ReportB.Unit, sum(Stock_ReportB.First_Qty) as qty,sum( Stock_ReportB.First_Amt) as First_Amt," & _
            " sum(Stock_ReportB.Instock) as In_Qty, sum(Stock_ReportB.In_Amt) as In_Amount," & _
            " sum(Stock_ReportB.OutStock) As Out_Qty,sum(Stock_ReportB.Out_Amt) as Out_Amt" & _
            " From Stock_ReportB" & _
            " where Stock_ReportB.Stock_ID='01'" & _
            " GROUP BY Stock_ReportB.Stock_ID, Stock_ReportB.ItemCode, Stock_ReportB.ItemName, Stock_ReportB.Unit"
        Case 2
            SQL = "SELECT Stock_ReportB.Stock_ID, Stock_ReportB.ItemCode, Stock_ReportB.ItemName," & _
            " Stock_ReportB.Unit, sum(Stock_ReportB.First_Qty) as qty,sum( Stock_ReportB.First_Amt) as First_Amt," & _
            " sum(Stock_ReportB.Instock) as In_Qty, sum(Stock_ReportB.In_Amt) as In_Amount," & _
            " sum(Stock_ReportB.OutStock) As Out_Qty,sum(Stock_ReportB.Out_Amt) as Out_Amt" & _
            " From Stock_ReportB" & _
            " where Stock_ReportB.Stock_ID='02'" & _
            " GROUP BY Stock_ReportB.Stock_ID, Stock_ReportB.ItemCode, Stock_ReportB.ItemName, Stock_ReportB.Unit"
    End Select
    
    Set crStockMoveInOut = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crStockMoveInOut
        .Database.AddADOCommand cnData, cmd
        
        .PluCode.SetUnboundFieldSource "{ado.ItemCode}"
        .PluName.SetUnboundFieldSource "{ado.ItemName}"
        .FirstQty.SetUnboundFieldSource "{ado.qty}"
        .FirstAmt.SetUnboundFieldSource "{ado.First_Amt}"
        .InQty.SetUnboundFieldSource "{ado.In_Qty}"
        .InAmt.SetUnboundFieldSource "{ado.In_Amount}"
        .OutQty.SetUnboundFieldSource "{ado.Out_Qty}"
        .OutAmt.SetUnboundFieldSource "{ado.Out_Amt}"
        .txtUnit.SetUnboundFieldSource "{ado.Unit}"
        .txtFromDate.SetText gfCONVERT_STRING_TO_DATE(from_Date)
        .txtToDate.SetText gfCONVERT_STRING_TO_DATE(To_Date)
        .txtTitle.SetText "B∏o c∏o Xu t nhÀp tÂn kho"
        With .FirstQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .FirstAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .InQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With

        With .InAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .OutQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With

        With .OutAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .LastQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With

        With .LastAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        
        With .Sumfirst
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .SumIn
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .SumOut
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .SumLast
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    Set iReport = crStockMoveInOut
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
MsgBox Err.Number & Err.Description & Me.name & " MoveInOutstock_Report"
End Sub


'Bao cao xuat nhap ton kho80
Public Sub MoveInOutstock_Report80()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL, sql1 As String
    Dim CRReport As CRAXDDRT.Report
    
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Select Case cboStock.ListIndex
        Case 0
            SQL = "SELECT Stock_ReportB.Stock_ID, Stock_ReportB.ItemCode, Stock_ReportB.ItemName," & _
            " Stock_ReportB.Unit, sum(Stock_ReportB.First_Qty) as qty,sum( Stock_ReportB.First_Amt) as First_Amt," & _
            " sum(Stock_ReportB.Instock) as In_Qty, sum(Stock_ReportB.In_Amt) as In_Amount," & _
            " sum(Stock_ReportB.OutStock) As Out_Qty,sum(Stock_ReportB.Out_Amt) as Out_Amt" & _
            " From Stock_ReportB" & _
            " GROUP BY Stock_ReportB.Stock_ID, Stock_ReportB.ItemCode, Stock_ReportB.ItemName, Stock_ReportB.Unit"
        Case 1
            SQL = "SELECT Stock_ReportB.Stock_ID, Stock_ReportB.ItemCode, Stock_ReportB.ItemName," & _
            " Stock_ReportB.Unit, sum(Stock_ReportB.First_Qty) as qty,sum( Stock_ReportB.First_Amt) as First_Amt," & _
            " sum(Stock_ReportB.Instock) as In_Qty, sum(Stock_ReportB.In_Amt) as In_Amount," & _
            " sum(Stock_ReportB.OutStock) As Out_Qty,sum(Stock_ReportB.Out_Amt) as Out_Amt" & _
            " From Stock_ReportB" & _
            " where Stock_ReportB.Stock_ID='01'" & _
            " GROUP BY Stock_ReportB.Stock_ID, Stock_ReportB.ItemCode, Stock_ReportB.ItemName, Stock_ReportB.Unit"
        Case 2
            SQL = "SELECT Stock_ReportB.Stock_ID, Stock_ReportB.ItemCode, Stock_ReportB.ItemName," & _
            " Stock_ReportB.Unit, sum(Stock_ReportB.First_Qty) as qty,sum( Stock_ReportB.First_Amt) as First_Amt," & _
            " sum(Stock_ReportB.Instock) as In_Qty, sum(Stock_ReportB.In_Amt) as In_Amount," & _
            " sum(Stock_ReportB.OutStock) As Out_Qty,sum(Stock_ReportB.Out_Amt) as Out_Amt" & _
            " From Stock_ReportB" & _
            " where Stock_ReportB.Stock_ID='02'" & _
            " GROUP BY Stock_ReportB.Stock_ID, Stock_ReportB.ItemCode, Stock_ReportB.ItemName, Stock_ReportB.Unit"
    End Select
    
    Set crStockMoveInOut80 = Nothing
    Set crStockMoveInOut58 = Nothing
    
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    If ReceiptType = 80 Then
        Set CRReport = crStockMoveInOut80
    Else
        Set CRReport = crStockMoveInOut58
    End If
    
    With CRReport
        .Database.AddADOCommand cnData, cmd
        .PluCode.SetUnboundFieldSource "{ado.ItemCode}"
        .PluName.SetUnboundFieldSource "{ado.ItemName}"
        .FirstQty.SetUnboundFieldSource "{ado.qty}"
        .InQty.SetUnboundFieldSource "{ado.In_Qty}"
        .OutQty.SetUnboundFieldSource "{ado.Out_Qty}"
        .txtFromDate.SetText gfCONVERT_STRING_TO_DATE(from_Date)
        .txtToDate.SetText gfCONVERT_STRING_TO_DATE(To_Date)
        .txtTitle.SetText "BC Xu t nhÀp tÂn"
        With .FirstQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .InQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With

        With .OutQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With

        With .LastQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With

       

    End With
    Set iReport = CRReport
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
MsgBox Err.Number & Err.Description & Me.name & " MoveInOutstock_Report"
End Sub
