VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSochitiet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "B¶n kª chi tiÕt mãn"
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11085
   ScaleWidth      =   15270
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid gridplu 
      Height          =   6015
      Left            =   4440
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10610
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   21
      WrapCellPointer =   -1  'True
      RowDividerStyle =   3
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
   Begin VB.Frame Frame2 
      Caption         =   "B¸o c¸o chi tiÕt mãn"
      Height          =   10935
      Left            =   4680
      TabIndex        =   9
      Top             =   120
      Width           =   10575
      Begin CRVIEWERLibCtl.CRViewer crvReport 
         Height          =   10575
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   10575
         DisplayGroupTree=   -1  'True
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   -1  'True
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   -1  'True
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   -1  'True
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   -1  'True
         DisplayBorder   =   -1  'True
         DisplayTabs     =   -1  'True
         DisplayBackgroundEdge=   -1  'True
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
   End
   Begin VB.Frame fra 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtItemCodes 
         Height          =   495
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tõ ngµy:...............§Õn ngµy :"
         Height          =   855
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   4095
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   270
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
            Format          =   64552961
            UpDown          =   -1  'True
            CurrentDate     =   39448
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   495
            Left            =   1800
            TabIndex        =   7
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
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
            Format          =   64552961
            UpDown          =   -1  'True
            CurrentDate     =   39448
         End
      End
      Begin VB.TextBox txtItemName 
         Height          =   495
         Left            =   1200
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   840
         Width           =   2535
      End
      Begin MSForms.CommandButton cmdclear 
         Height          =   495
         Left            =   3720
         TabIndex        =   15
         Top             =   840
         Width           =   615
         Caption         =   "CLR"
         Size            =   "1085;873"
         FontName        =   ".VnArial"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   495
         Left            =   3720
         TabIndex        =   14
         Top             =   240
         Width           =   615
         Caption         =   "..."
         Size            =   "1085;873"
         FontName        =   ".VnArial"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdPrint 
         Height          =   975
         Left            =   1560
         TabIndex        =   13
         Top             =   2880
         Width           =   1335
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "In b¸o c¸o"
         Size            =   "2355;1720"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdExit 
         Height          =   975
         Left            =   3000
         TabIndex        =   10
         Top             =   2880
         Width           =   1215
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "Tho¸t"
         Size            =   "2143;1720"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdView 
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   1335
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "Xem"
         Size            =   "2355;1720"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tªn hµng"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "M· hµng"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmSochitiet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPLU As New ADODB.Recordset
Dim iReport As New CRAXDDRT.Report
Dim TotalRptPage As Integer
Dim isESC As Boolean

Private Sub cmdclear_Click()
    txtItemCodes.Text = ""
    txtItemName.Text = ""
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Handle
   myPrint iReport, crvReport.GetCurrentPageNumber, TotalRptPage
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdPrint_Click "
End Sub

Private Sub cmdView_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL As String

'     If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    'Khong danh cho karaoke
    SQL = "SELECT Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, right(Invoice_Totals.Invoice_Number,4) as Invoice_No, Invoice_Totals.DateTime" & _
          " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number " & _
          " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "')) and Invoice_Itemized.ItemNum='" & txtItemCodes.Text & "'" & _
          " and Invoice_totals.Status<> 'CO'" & _
          " GROUP BY Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName,Invoice_Totals.DateTime,right(Invoice_Totals.Invoice_Number,4), Invoice_Itemized.PricePer"
    
   
    Set crStockCard = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crStockCard
        .Database.AddADOCommand cnData, cmd
        .DocNo.SetUnboundFieldSource "{ado.Invoice_No}"
        .DocDate.SetUnboundFieldSource "{ado.DateTime}"
        .Qty.SetUnboundFieldSource "{ado.Qty}"
        .Price.SetUnboundFieldSource "{ado.Price}"
        
        .txtDateFrom.SetText dtpFromDate.Value
        .txtDateTo.SetText dtpToDate.Value
        .txtCode.SetText txtItemCodes.Text
        .txtName.SetText txtItemName.Text
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
    End With
    Set iReport = crStockCard
    With crvReport
'        .DisplayBorder = False
        .ReportSource = iReport
        .EnableSearchControl = False
        .EnableStopButton = False
        .EnableGroupTree = False
        .EnableAnimationCtrl = False
        .EnablePopupMenu = False
        .EnableToolbar = True
        '.DisplayToolbar = False
        .DisplayTabs = False
        .ToolTipText = ""
        .ViewReport
        crvReport.Zoom 100
        While .IsBusy
            DoEvents
        Wend
        .ShowLastPage
        TotalRptPage = .GetCurrentPageNumber
        While .IsBusy
            DoEvents
        Wend
        .ShowFirstPage
        While .IsBusy
            DoEvents
        Wend
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdBaocaochitiet_Click"
End Sub

Private Sub CommandButton1_Click()
    Call txtItemCodes_KeyDown(vbKeyDown, 2)
    If isESC = True Then
        gridplu.Visible = False
        isESC = False
        Exit Sub
    End If
    isESC = True

End Sub

Private Sub Form_Load()
On Error GoTo Handle
    dtpFromDate.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    dtpToDate.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    If cnData.State = 0 Then Exit Sub
    Set rsPLU = OpenCriticalTable("select ItemNum,ItemName,Std_price1, std_price2, std_price3,Unit from Inventory", cnData)
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub



Private Sub gridplu_DblClick()
On Error GoTo errHdl
        With rsPLU
            If .RecordCount = 0 Then
                gridplu.Visible = False
                txtItemCodes.SetFocus
                MsgBox "Kh«ng t×m thÊy mÆt hµng ®· chän", vbExclamation
                Exit Sub
            End If
            txtItemCodes.Text = !ItemNum
            txtItemName.Text = !ItemName
        End With
        gridplu.Visible = False
        dtpFromDate.SetFocus
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - gridplu_Click "
End Sub

Private Sub gridplu_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl
    If KeyAscii = 27 Then
        gridplu.Visible = False
        txtItemCodes.SetFocus
    ElseIf KeyAscii = 13 Then
        With rsPLU
            If .RecordCount = 0 Then
                gridplu.Visible = False
                txtItemCodes.SetFocus
                MsgBox "Kh«ng t×m thÊy mÆt hµng ®· chän", vbExclamation
                Exit Sub
            End If
            txtItemCodes.Text = !ItemNum
            txtItemName.Text = !ItemName
        End With
        gridplu.Visible = False
        cmdView.SetFocus
    End If
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - gridplu_KeyPress "

End Sub

Private Sub txtItemCodes_DblClick()
On Error GoTo Handle
    With frmKeyboard
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtItemCodes.Text = .Let_Text_Input
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtItemCodes_DblClick"
End Sub

Private Sub txtItemCodes_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl
    If KeyCode = vbKeyDown Then
            With rsPLU
                If .State = adStateOpen Then .Close
                If InStr(1, txtItemCodes.Text, "*", vbTextCompare) > 0 Then
                    .Open "SELECT  ItemNum, ItemName,Std_price1, std_price2, std_price3, Unit FROM Inventory WHERE INSTR(ItemNum,""" & Left(txtItemCodes.Text, Len(Trim(txtItemCodes.Text)) - 1) & "%"")>0 OR ItemName LIKE '" & _
                    Left(Trim(txtItemCodes.Text), Len(Trim(txtItemCodes.Text)) - 1) & "%'  ORDER BY ItemNum asc"
                Else
                    .Open "SELECT  ItemNum, ItemName,Std_price1, std_price2, std_price3, Unit FROM Inventory WHERE (INSTR(ItemNum,""" & Trim(txtItemCodes.Text) & """)>0 OR INSTR(ItemName,""" & _
                    Trim(txtItemCodes.Text) & """)>0) AND TRIM(ItemName)<>"""" ORDER BY Itemnum ASC"
                End If
            End With
        With gridplu
            Set .DataSource = rsPLU
            .Columns(0).Caption = "M· sè" '"M· hµng"
            .Columns(0).Width = 1600
            .Columns(1).Caption = "DiÔn gi¶i" ' "Tªn hµng"
            .Columns(1).Width = 4000
            .Columns(1).Alignment = dbgLeft
            .Columns(2).Caption = "Gi¸ chu¶n 1"
            .Columns(2).Alignment = dbgCenter
            .Columns(2).Width = 1100
            .Columns(3).Caption = "Gi¸ chu¶n 2"
            .Columns(3).Alignment = dbgCenter
            .Columns(3).Width = 1100
            .Columns(4).Caption = "Gi¸ chu¶n 3"
            .Columns(4).Alignment = dbgCenter
            .Columns(4).Width = 1100
            .Columns(5).Caption = "§VT"
            .Columns(5).Alignment = dbgCenter
            .Columns(5).Width = 1000
            .Visible = True
            .SetFocus
            .top = txtItemCodes.top + txtItemCodes.Height
            .Left = txtItemCodes.Left
        End With
    End If
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - txtPluCode_KeyDown "
End Sub
