VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmShow_Vendor_Stock_report 
   Caption         =   "B∏o c∏o kho"
   ClientHeight    =   9900
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   15240
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
   ScaleHeight     =   9900
   ScaleWidth      =   15240
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
      Height          =   945
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   14205
      TabIndex        =   1
      Top             =   0
      Width           =   14235
      Begin VB.CommandButton cmdPrint 
         Height          =   510
         Left            =   9420
         Picture         =   "frmShow_Vendor_Stock_report.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   195
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
         Left            =   7620
         TabIndex        =   11
         Top             =   195
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
         Left            =   6780
         Picture         =   "frmShow_Vendor_Stock_report.frx":06EA
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Left            =   7200
         Picture         =   "frmShow_Vendor_Stock_report.frx":0A2C
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Left            =   8520
         Picture         =   "frmShow_Vendor_Stock_report.frx":0D6E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
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
         Left            =   8940
         Picture         =   "frmShow_Vendor_Stock_report.frx":10B0
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
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
         Left            =   5100
         TabIndex        =   6
         Text            =   "cboZoom"
         Top             =   300
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   0
         Width           =   4935
         Begin VB.OptionButton optAll 
            Caption         =   "T t c∂"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton optVen 
            Caption         =   "Nhµ cung c p"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txtVendor_Num 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1680
            TabIndex        =   3
            Top             =   360
            Width           =   3135
         End
      End
      Begin prjTouchScreen.MyButton cmdExport 
         Height          =   615
         Left            =   10440
         TabIndex        =   13
         Top             =   120
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
         MICON           =   "frmShow_Vendor_Stock_report.frx":13F2
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
         TabIndex        =   14
         Top             =   120
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
         MICON           =   "frmShow_Vendor_Stock_report.frx":140E
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
   Begin MSDataGridLib.DataGrid grdVendor 
      Height          =   6015
      Left            =   240
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   10610
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
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
         Size            =   9.75
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
   Begin CRVIEWERLibCtl.CRViewer crvReport 
      CausesValidation=   0   'False
      Height          =   7695
      Left            =   240
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1260
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
Attribute VB_Name = "frmShow_Vendor_Stock_report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CRReport As New CRAXDDRT.Report
Dim iReport As New CRAXDDRT.Report
Dim TotalRptPage As Integer
Dim Report_Number As Integer
Dim from_Date, To_Date As String
Dim Table_Month As String
Dim fLoad As Boolean
Dim rsVendor As New ADODB.Recordset
Dim opt As Integer

Public Property Let Report(ByVal vNewValue As Variant)
    Set CRReport = vNewValue
End Property

Private Sub cboStock_Change()
    Set iReport = Nothing
    Set crStock = Nothing
    If fLoad Then
        Select Case Report_Number
            Case 2:
                Call Instock_Report
        End Select
    End If
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

    Set rsVendor = Open_Table(cnData, "Vendors")
    optAll.Value = True
    opt = 1
    Table_Month = Mid(To_Date, 5, 2) & Mid(To_Date, 3, 2)
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
        Case 2:
            Call Instock_Report
    End Select
   
    Me.WindowState = 2
    txtPage.Text = crvReport.GetCurrentPageNumber & " / " & TotalRptPage
    fLoad = True
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
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
    Select Case opt
        Case 1
            SQL = "SELECT Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description, Sum(Inventory_InB" & Table_Month & ".Quantity) AS Qty, Avg(Inventory_InB" & Table_Month & ".CostPer) AS Cost, Sum(Inventory_InB" & Table_Month & ".Amount) AS Amt" & _
                 " FROM Instock_MasterB INNER JOIN Inventory_InB" & Table_Month & " ON Instock_MasterB.Doc_Number = Inventory_InB" & Table_Month & ".Doc_Number" & _
                 " WHERE (((Instock_MasterB.DateTime)>='" & from_Date & "' And (Instock_MasterB.DateTime)<='" & To_Date & "') AND ((Instock_MasterB.InOutType)<>'O'))" & _
                 " GROUP BY Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description"
                 
        Case 2
            SQL = "SELECT Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description, Sum(Inventory_InB" & Table_Month & ".Quantity) AS Qty, Avg(Inventory_InB" & Table_Month & ".CostPer) AS Cost, Sum(Inventory_InB" & Table_Month & ".Amount) AS Amt" & _
                 " FROM Instock_MasterB INNER JOIN Inventory_InB" & Table_Month & " ON Instock_MasterB.Doc_Number = Inventory_InB" & Table_Month & ".Doc_Number" & _
                 " WHERE (((Instock_MasterB.DateTime)>='" & from_Date & "' And (Instock_MasterB.DateTime)<='" & To_Date & "') AND ((Instock_MasterB.InOutType)<>'O')) and Instock_MasterB.Vendor_Number='" & txtVendor_Num.Text & "'" & _
                 " GROUP BY Instock_MasterB.Stock_ID,Inventory_InB" & Table_Month & ".ItemNum, Inventory_InB" & Table_Month & ".Description"
                        
      
    End Select
    
    Set crStock = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crStock
        .Database.AddADOCommand cnData, cmd
        
        '.GroupA.SetUnboundFieldSource "{ado.GroupA}"
        .PluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .PluName.SetUnboundFieldSource "{ado.Description}"
        '.Unit.SetUnboundFieldSource "{ado.Unit}"
        .Qty.SetUnboundFieldSource "{ado.Qty}"
        .Price.SetUnboundFieldSource "{ado.cost}"
        .Amount.SetUnboundFieldSource "{ado.Amt}"
        .stockID.SetUnboundFieldSource "{ado.Stock_ID}"
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

Private Sub grdVendor_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl
    If KeyAscii = 27 Then
        grdVendor.Visible = False
        txtVendor_Num.SetFocus
    ElseIf KeyAscii = 13 Then
        With rsVendor
            If .RecordCount = 0 Then
                grdVendor.Visible = False
                txtVendor_Num.SetFocus
                MsgBox "M∑ nhµ cung c p nµy kh´ng tÂn tπi", vbExclamation
                Exit Sub
            End If
            txtVendor_Num.Text = grdVendor.Columns(0)
            Call Instock_Report
        End With
        grdVendor.Visible = False
    End If
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - grdVendor_KeyPress "
End Sub

Private Sub optAll_Click()
    optVen.Value = False
    opt = 1
End Sub



Private Sub optVen_Click()
    optAll.Value = False
    opt = 2
    txtVendor_Num.SetFocus
End Sub


Private Sub txtVendor_Num_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Handle
    If KeyCode = vbKeyDown Then
            With rsVendor
                If .State = adStateOpen Then .Close
                If InStr(1, txtVendor_Num.Text, "*", vbTextCompare) > 0 Then
                    .Open "SELECT  Vendor_Number, Vendor_Name, Company FROM Vendors WHERE INSTR(Vendor_Number,""" & Left(txtVendor_Num.Text, Len(Trim(txtVendor_Num.Text)) - 1) & "%"")>0 OR Vendor_Name LIKE '" & _
                    Left(Trim(txtVendor_Num.Text), Len(Trim(txtVendor_Num.Text)) - 1) & "%'  ORDER BY Vendor_Number asc"
                Else
                    .Open "SELECT  Vendor_Number, Vendor_Name, Company FROM Vendors WHERE (INSTR(Vendor_Number,""" & Trim(txtVendor_Num.Text) & """)>0 OR INSTR(Vendor_Name,""" & _
                    Trim(txtVendor_Num.Text) & """)>0) AND TRIM(Vendor_Name)<>"""" ORDER BY Vendor_Number ASC"
                End If
            End With
        
        With grdVendor
            Set .DataSource = rsVendor
            .Columns(0).Caption = "M∑ nhµ cung c p"
            .Columns(0).Width = 2500
            .Columns(1).Caption = "T™n Nhµ cung c p"
            .Columns(1).Width = 4500
            .Columns(1).Alignment = dbgLeft
            .Visible = True
            .SetFocus
            .top = txtVendor_Num.top + txtVendor_Num.Height + 100
            .Left = txtVendor_Num.Left
        End With
    End If
    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub


