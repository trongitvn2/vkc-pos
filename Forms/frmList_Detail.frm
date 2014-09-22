VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmList_Detail 
   Caption         =   "B¶n kª c¸c lo¹i"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6705
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmList_Detail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboType 
      Height          =   390
      ItemData        =   "frmList_Detail.frx":000C
      Left            =   4440
      List            =   "frmList_Detail.frx":0016
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   120
      Width           =   2175
   End
   Begin MSDataGridLib.DataGrid dgrItems 
      Height          =   4335
      Left            =   5160
      TabIndex        =   11
      Top             =   4800
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7646
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   21
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
   Begin VB.CommandButton cmdFind 
      Caption         =   "..."
      Height          =   495
      Left            =   3360
      TabIndex        =   10
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtItemName 
      Height          =   495
      Left            =   3960
      TabIndex        =   9
      Top             =   840
      Width           =   2655
   End
   Begin VB.TextBox txtItemNum 
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   840
      Width           =   2055
   End
   Begin prjTouchScreen.MyButton cmdinstock_List 
      Height          =   1095
      Left            =   360
      TabIndex        =   0
      Top             =   2160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1931
      BTYPE           =   6
      TX              =   "B¶n kª nhËp kho"
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmList_Detail.frx":0031
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdList_Outstock 
      Height          =   1095
      Left            =   3480
      TabIndex        =   2
      Top             =   2160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1931
      BTYPE           =   6
      TX              =   "B¶n kª xuÊt kho"
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmList_Detail.frx":004D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdStockCard 
      Height          =   1095
      Left            =   360
      TabIndex        =   3
      Top             =   3480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1931
      BTYPE           =   6
      TX              =   "ThÎ kho"
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmList_Detail.frx":0069
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdStockDetail 
      Height          =   1095
      Left            =   3480
      TabIndex        =   4
      Top             =   3480
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1931
      BTYPE           =   6
      TX              =   "Sæ chi tiÕt"
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
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmList_Detail.frx":0085
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdexit 
      Height          =   855
      Left            =   1800
      TabIndex        =   5
      Top             =   4920
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1508
      BTYPE           =   6
      TX              =   "Tho¸t"
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
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmList_Detail.frx":00A1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpFromDate 
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   1470
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   63963137
      UpDown          =   -1  'True
      CurrentDate     =   39448
   End
   Begin MSComCtl2.DTPicker dtpToDate 
      Height          =   495
      Left            =   4680
      TabIndex        =   14
      Top             =   1440
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
      Format          =   63963137
      UpDown          =   -1  'True
      CurrentDate     =   39448
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "§Õn ngµy:"
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Tõ ngµy:"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "M· hµng:"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "b¶n kª hµng hãa"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmList_Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsItem As New ADODB.Recordset

Private Sub cboType_Change()
On Error GoTo Handle
    Select Case cboType.ListIndex
        Case 0:
            Set rsItem = OpenCriticalTable("Select ItemNum,ItemName,Unit from Inventory", cnData)
        Case 1:
            Set rsItem = OpenCriticalTable("Select Plucode,PluName,Unit from SetMPLU", cnData)
    End Select

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cboType_Change"
End Sub

Private Sub cboType_Click()
    Call cboType_Change
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
On Error GoTo errHdl
        If cboType.ListIndex = 1 Then
            With rsItem
                If .State = adStateOpen Then .Close
                If InStr(1, txtItemNum.Text, "*", vbTextCompare) > 0 Then
                    .Open "SELECT  PluCode, PluName, Unit FROM SetMPLU WHERE INSTR(PluCode,""" & Left(txtItemNum.Text, Len(Trim(txtItemNum.Text)) - 1) & "%"")>0 OR PluName LIKE '" & _
                    Left(Trim(txtItemNum.Text), Len(Trim(txtItemNum.Text)) - 1) & "%'  ORDER BY PluCode asc"
                Else
                    .Open "SELECT  PluCode, PluName, Unit FROM SetMPLU WHERE (INSTR(PluCode,""" & Trim(txtItemNum.Text) & """)>0 OR INSTR(PluName,""" & _
                    Trim(txtItemNum.Text) & """)>0) AND TRIM(PluName)<>"""" ORDER BY PluCode ASC"
                End If
            End With
        Else
            Set rsItem = OpenCriticalTable("Select ItemNum,ItemName,Unit from Inventory", cnData)
            With rsItem
                If .State = adStateOpen Then .Close
                If InStr(1, txtItemNum.Text, "*", vbTextCompare) > 0 Then
                    .Open "SELECT  ItemNum, ItemName, Unit FROM Inventory WHERE INSTR(ItemNum,""" & Left(txtItemNum.Text, Len(Trim(txtItemNum.Text)) - 1) & "%"")>0 OR ItemName LIKE '" & _
                    Left(Trim(txtItemNum.Text), Len(Trim(txtItemNum.Text)) - 1) & "%'  ORDER BY ItemNum asc"
                Else
                    .Open "SELECT  ItemNum, ItemName, Unit FROM Inventory WHERE (INSTR(ItemNum,""" & Trim(txtItemNum.Text) & """)>0 OR INSTR(ItemName,""" & _
                    Trim(txtItemNum.Text) & """)>0) AND TRIM(ItemName)<>"""" ORDER BY ItemNum ASC"
                End If
            End With
        End If
        With dgrItems
            Set .DataSource = rsItem
            .Columns(0).Caption = "M· hµng"
            .Columns(0).Width = 1800
            .Columns(1).Caption = "Tªn hµng"
            .Columns(1).Width = 2500
            .Columns(1).Alignment = dbgLeft
            .Columns(2).Caption = "§VT"
            .Columns(2).Alignment = dbgCenter
            .Columns(2).Width = 1000
            .Visible = True
            .SetFocus
            .top = txtItemNum.top + 200
            .Left = txtItemNum.Left + 100
        End With
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - cmdFind_Click "
End Sub

Private Sub cmdinstock_List_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim iReport As CRAXDDRT.Report
    Dim SQL, sql1 As String
    Dim DateStock As String
    
    DateStock = Format(Month(dtpToDate.Value), "00") & Right(Format(Year(dtpToDate.Value), "00"), 2)
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
            SQL = "SELECT Instock_MasterB.Doc_Number, Instock_MasterB.DateTime, Instock_MasterB.Vendor_Number, Inventory_InB" & DateStock & ".ItemNum, Inventory_InB" & DateStock & ".Description, Inventory_InB" & DateStock & ".Quantity, Inventory_InB" & DateStock & ".CostPer, Inventory_InB" & DateStock & ".Amount" & _
                 " FROM Instock_MasterB INNER JOIN Inventory_InB" & DateStock & " ON Instock_MasterB.Doc_Number=Inventory_InB" & DateStock & ".Doc_Number " & _
                 " WHERE (((Left(Instock_MasterB.Doc_Number,2))='NK')) and Instock_MasterB.DateTime<='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'" & _
                 " GROUP BY Instock_MasterB.Doc_Number, Instock_MasterB.DateTime, Instock_MasterB.Vendor_Number, Inventory_InB" & DateStock & ".ItemNum, Inventory_InB" & DateStock & ".Description, Inventory_InB" & DateStock & ".Quantity, Inventory_InB" & DateStock & ".CostPer, Inventory_InB" & DateStock & ".Amount" & _
                 " HAVING (((Inventory_InB" & DateStock & ".ItemNum)='" & txtItemNum.Text & "'))"
    
    Set crInstockList = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crInstockList
        .Database.AddADOCommand cnData, cmd
        
        .DocNo.SetUnboundFieldSource "{ado.Doc_Number}"
        .DocDate.SetUnboundFieldSource "{ado.DateTime}"
        .Qty.SetUnboundFieldSource "{ado.Quantity}"
        .Price.SetUnboundFieldSource "{ado.Costper}"
        .Amount.SetUnboundFieldSource "{ado.Amount}"
        .Vendor.SetUnboundFieldSource "{ado.Vendor_Number}"
        .txtCode.SetText txtItemNum.Text
        .txtName.SetText txtItemName.Text
        .txtDateFrom.SetText dtpFromDate.Value
        .txtDateTo.SetText dtpToDate.Value
        .txtTitle.SetText "B¶n kª nhËp kho"
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
    Set iReport = crInstockList
     With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdinstock_List_Click()"
End Sub

Private Sub cmdStockCard_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim iReport As CRAXDDRT.Report
    Dim SQL, sql1 As String
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
            SQL = "SELECT Stock_ReportB.DocNumber, Stock_ReportB.DateTime, Stock_ReportB.ItemCode, Stock_ReportB.ItemName, Stock_ReportB.Unit, Stock_ReportB.First_Qty, Stock_ReportB.Instock, Stock_ReportB.OutStock, Stock_ReportB.Last_Qty" & _
                  " FROM Stock_ReportB" & _
                  " where Stock_ReportB.DateTime >='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "' and Stock_ReportB.DateTime<='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "' and Stock_ReportB.ItemCode='" & txtItemNum.Text & "'" & _
                  " Order by Stock_ReportB.DateTime"
    
    Set crThekho = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crThekho
        .Database.AddADOCommand cnData, cmd
        
        .DocNo.SetUnboundFieldSource "{ado.DocNumber}"
        .DocDate.SetUnboundFieldSource "{ado.DateTime}"
        .FirstQty.SetUnboundFieldSource "{ado.First_Qty}"
        .InQty.SetUnboundFieldSource "{ado.Instock}"
        .OutQty.SetUnboundFieldSource "{ado.Outstock}"
        .txtCode.SetText txtItemNum.Text
        .txtName.SetText txtItemName.Text
        .txtDateFrom.SetText dtpFromDate.Value
        .txtDateTo.SetText dtpToDate.Value
        .txtTitle.SetText "ThÎ kho"
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
    Set iReport = crThekho
     With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdStockCard_Click"
End Sub


Private Sub Command1_Click()

End Sub

Private Sub dgrItems_DblClick()
On Error GoTo Handle

    With rsItem
        If .RecordCount = 0 Then
            dgrItems.Visible = False
            txtItemNum.SetFocus
        
            Exit Sub
        End If
        If cboType.ListIndex = 0 Then
            txtItemNum.Text = !ItemNum
            txtItemName.Text = !ItemName
        Else
            txtItemNum.Text = !PluCode
            txtItemName.Text = !PluName
        End If
        
    
        dgrItems.Visible = False
        
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - dgrItems_DblClick"
End Sub

Private Sub dgrItems_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl
    If KeyAscii = 27 Then
        dgrItems.Visible = False
        txtItemNum.SetFocus
    ElseIf KeyAscii = 13 Then
        With rsItem
            If .RecordCount = 0 Then
                dgrItems.Visible = False
                txtItemNum.SetFocus
            
                Exit Sub
            End If
            If cboType.ListIndex = 0 Then
                txtItemNum.Text = !ItemNum
                txtItemName.Text = !ItemName
            Else
                txtItemNum.Text = !PluCode
                txtItemName.Text = !PluName
            End If
            
        
            dgrItems.Visible = False
            
        End With
    ElseIf KeyAscii = 9 Then
        dgrItems.Visible = False
        txtItemNum.SetFocus
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - dgrItems_KeyPress "
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    cboType.ListIndex = 0
    Set rsItem = OpenCriticalTable("Select ItemNum,ItemName,Unit from Inventory", cnData)
    dtpFromDate.Value = "01/" & Mid(DateDefault, 5, 2) & "/" & Left(DateDefault, 4)
    dtpToDate.Value = Format(Now, "dd/MM/yyyy")
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Private Sub txtItemNum_DblClick()
On Error GoTo Handle
        With frmKeyboard
            .FormCallkeyboard = "Other"
            .txtInput.PasswordChar = ""
            .Show vbModal
            txtItemNum.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtItemNum_DblClick"
End Sub

Private Sub txtItemNum_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl
    If KeyCode = vbKeyDown Then
        If cboType.ListIndex = 1 Then
            With rsItem
                If .State = adStateOpen Then .Close
                If InStr(1, txtItemNum.Text, "*", vbTextCompare) > 0 Then
                    .Open "SELECT  PluCode, PluName, Unit FROM SetMPLU WHERE INSTR(PluCode,""" & Left(txtItemNum.Text, Len(Trim(txtItemNum.Text)) - 1) & "%"")>0 OR PluName LIKE '" & _
                    Left(Trim(txtItemNum.Text), Len(Trim(txtItemNum.Text)) - 1) & "%'  ORDER BY PluCode asc"
                Else
                    .Open "SELECT  PluCode, PluName, Unit FROM SetMPLU WHERE (INSTR(PluCode,""" & Trim(txtItemNum.Text) & """)>0 OR INSTR(PluName,""" & _
                    Trim(txtItemNum.Text) & """)>0) AND TRIM(PluName)<>"""" ORDER BY PluCode ASC"
                End If
            End With
        Else
            Set rsItem = OpenCriticalTable("Select ItemNum,ItemName,Unit from Inventory", cnData)
            With rsItem
                If .State = adStateOpen Then .Close
                If InStr(1, txtItemNum.Text, "*", vbTextCompare) > 0 Then
                    .Open "SELECT  ItemNum, ItemName, Unit FROM Inventory WHERE INSTR(ItemNum,""" & Left(txtItemNum.Text, Len(Trim(txtItemNum.Text)) - 1) & "%"")>0 OR ItemName LIKE '" & _
                    Left(Trim(txtItemNum.Text), Len(Trim(txtItemNum.Text)) - 1) & "%'  ORDER BY ItemNum asc"
                Else
                    .Open "SELECT  ItemNum, ItemName, Unit FROM Inventory WHERE (INSTR(ItemNum,""" & Trim(txtItemNum.Text) & """)>0 OR INSTR(ItemName,""" & _
                    Trim(txtItemNum.Text) & """)>0) AND TRIM(ItemName)<>"""" ORDER BY ItemNum ASC"
                End If
            End With
        End If
        With dgrItems
            Set .DataSource = rsItem
            .Columns(0).Caption = "M· hµng"
            .Columns(0).Width = 1800
            .Columns(1).Caption = "Tªn hµng"
            .Columns(1).Width = 2500
            .Columns(1).Alignment = dbgLeft
            .Columns(2).Caption = "§VT"
            .Columns(2).Alignment = dbgCenter
            .Columns(2).Width = 1000
            .Visible = True
            .SetFocus
            .top = txtItemNum.top + 200
            .Left = txtItemNum.Left + 100
        End With
    End If
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - txtItemNum_KeyDown "
End Sub
