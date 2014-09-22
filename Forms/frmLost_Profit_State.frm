VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLost_Profit_State 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13500
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11040
   ScaleWidth      =   13500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra 
      Height          =   10215
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   13215
      Begin prjTouchScreen.MyProgressBar prbprocess 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         Max             =   300
         Value           =   300
         ProgressLook    =   3
      End
      Begin prjTouchScreen.MyButton cmdCalcu 
         Height          =   615
         Left            =   5640
         TabIndex        =   6
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BTYPE           =   6
         TX              =   "T&›nh l∑i lÁ"
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
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmLost_Profit_State.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.Frame Frame1 
         Height          =   8535
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   12975
         Begin MSFlexGridLib.MSFlexGrid flgProfit 
            Height          =   8175
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   12735
            _ExtentX        =   22463
            _ExtentY        =   14420
            _Version        =   393216
         End
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   495
         Left            =   1140
         TabIndex        =   1
         Top             =   390
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
         Format          =   70516737
         UpDown          =   -1  'True
         CurrentDate     =   40189
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   495
         Left            =   3900
         TabIndex        =   2
         Top             =   360
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
         Format          =   70516737
         UpDown          =   -1  'True
         CurrentDate     =   40189
      End
      Begin prjTouchScreen.MyButton cmdViewReport 
         Height          =   615
         Left            =   5640
         TabIndex        =   7
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BTYPE           =   6
         TX              =   "In b∏o c∏o"
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
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmLost_Profit_State.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdEditSetMenu 
         Height          =   615
         Left            =   7920
         TabIndex        =   8
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BTYPE           =   6
         TX              =   "Hi÷u chÿnh  Æﬁnh m¯c"
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
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmLost_Profit_State.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton MyButton3 
         Height          =   615
         Left            =   7920
         TabIndex        =   9
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         BTYPE           =   6
         TX              =   "..."
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
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmLost_Profit_State.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdExit 
         Height          =   1335
         Left            =   10200
         TabIndex        =   10
         Top             =   240
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2355
         BTYPE           =   6
         TX              =   "&Tho∏t"
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
         BCOLO           =   33023
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmLost_Profit_State.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.Label lblFromdate 
         Caption         =   "Tı ngµy:"
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
         TabIndex        =   4
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label lblDenngay 
         Caption         =   "ß’n ngµy:"
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
         Left            =   2700
         TabIndex        =   3
         Top             =   450
         Width           =   1185
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "L∑i lÁ tr™n m„n"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3600
      TabIndex        =   11
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "frmLost_Profit_State"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim imonth As String
Private Sub cmdCalcu_Click()
On Error GoTo Handle
Dim rsProfit As New ADODB.Recordset
Dim SQL As String
    prbprocess.Value = 0
    SQL = "SELECT Invoice_Itemized.ItemNum,Inventory.Unit, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) as Qty,Sum(Invoice_Itemized.Quantity)* Avg(Invoice_Itemized.PricePer) AS SalingPrice, Sum(Invoice_Itemized.Quantity)*Inventory.Minstock as CostPrice,Sum(Invoice_Itemized.Quantity)* Avg(Invoice_Itemized.PricePer) - Sum(Invoice_Itemized.Quantity)*Inventory.Minstock as Profit, Inventory.Dept_ID, Departments.Description " & _
         " FROM Departments INNER JOIN (Inventory INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Departments.Dept_ID = Inventory.Dept_ID" & _
         " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "' And (Left([Invoice_Totals].[DateTime],8))<='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'))" & _
         " GROUP BY Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Inventory.Minstock, Inventory.Dept_ID, Departments.Description,Invoice_Itemized.PricePer, Inventory.Unit"
    
    imonth = Format(Month(dtpFromDate.Value), "00")
    If Month(dtpFromDate.Value) <> Month(dtpToDate.Value) Then MsgBox "L∑i lÁ chÿ t›nh trong 1 th∏ng"
    
    Set rsProfit = OpenCriticalTable(SQL, cnData)
    Call SetFLGRIDORDER(rsProfit)
        prbprocess.Value = prbprocess.Max

Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & " - cmdCalcu_Click"
End Sub

Private Sub cmdEditSetMenu_Click()
    frmSetMenuLink.Show vbModal
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdViewReport_Click()
On Error GoTo Handle
    Dim SQL As String
    Dim cmd As New ADODB.Command
    Dim iReport As CRAXDDRT.Report
    prbprocess.Value = 0
    If prbprocess.Value < prbprocess.Max Then prbprocess.Value = prbprocess.Max
    
    SQL = "SELECT Invoice_Itemized.ItemNum,Inventory.Unit, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) as Qty,Sum(Invoice_Itemized.Quantity)* Avg(Invoice_Itemized.PricePer) AS SalingPrice, Sum(Invoice_Itemized.Quantity)*Inventory.Minstock as CostPrice, Inventory.Dept_ID, Departments.Description " & _
         " FROM Departments INNER JOIN (Inventory INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Departments.Dept_ID = Inventory.Dept_ID" & _
         " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "' And (Left([Invoice_Totals].[DateTime],8))<='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'))" & _
         " GROUP BY Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Inventory.Minstock, Inventory.Dept_ID, Departments.Description,Invoice_Itemized.PricePer, Inventory.Unit"
'
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set crLost_Profit_Statement = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crLost_Profit_Statement
        .Database.AddADOCommand cnData, cmd
        .txtGroupA.SetUnboundFieldSource "{ado.Description}"
        .txtPluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .txtPluName.SetUnboundFieldSource "{ado.DiffItemName}"
        .txtUnit.SetUnboundFieldSource "{ado.Unit}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtSalingPrice.SetUnboundFieldSource "{ado.SalingPrice}"
        .txtCostPrice.SetUnboundFieldSource "{ado.CostPrice}"
'        .txtCostPrice.SetUnboundFieldSource "{ado.Minstock}"
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate
        With .txtSalingPrice
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtCostPrice
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtGroupSumCost
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtGroupSumSaling
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtGroupSumProfit
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtSumCost
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtSumSaling
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtSumProfit
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtProfit
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    Set iReport = crLost_Profit_Statement
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With


    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdViewReport_Click"
End Sub

Private Sub Form_Load()
 On Error GoTo Handle
   
    dtpFromDate.Value = "01/" & Mid(DateDefault, 5, 2) & "/" & Left(DateDefault, 4)
    dtpToDate.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    Call Set_flgOrder
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & " - Form_Load"
End Sub

Public Sub Set_flgOrder()
    On Error GoTo Handle
    Dim i As Integer
        With flgProfit
            .Cols = 7
            .Rows = 30
            .ColWidth(0) = 1580
            .ColWidth(1) = 3900
            .ColWidth(2) = 1200
            .ColWidth(3) = 1200
            .ColWidth(4) = 1500
            .ColWidth(5) = 1500
            .ColWidth(6) = 1500
            .TextMatrix(0, 0) = "M∑ hµng"
            .TextMatrix(0, 1) = "T™n hµng"
            .TextMatrix(0, 2) = "ß¨n vﬁ t›nh"
            .TextMatrix(0, 3) = "SË l≠Óng"
            .TextMatrix(0, 4) = "Gi∏ b∏n"
            .TextMatrix(0, 5) = "G›a vËn"
            .TextMatrix(0, 6) = "L∑i"
            .ColAlignment(1) = 2
            .ColAlignment(2) = 4
            .ColAlignment(3) = 4
            .ColAlignment(4) = 6
            .ColAlignment(5) = 6
            .ColAlignment(6) = 6
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
            .TextMatrix(1, 4) = ""
            .TextMatrix(1, 5) = ""
            .TextMatrix(1, 6) = ""
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Set_flgOrder"

End Sub


Public Sub SetFLGRIDORDER(rs As ADODB.Recordset)
On Error GoTo Handle
        Dim incount As Integer
        rs.MoveFirst
        prbprocess.Min = 0
        prbprocess.Max = 1000
        With rs
            Do While Not .EOF
                incount = incount + 1
                flgProfit.Rows = rs.RecordCount + 1
                With flgProfit
                    .TextMatrix(incount, 0) = rs!ItemNum
                    .TextMatrix(incount, 1) = rs!DiffItemName
                    .TextMatrix(incount, 2) = rs!Unit
                    .TextMatrix(incount, 3) = rs!Qty
                    .TextMatrix(incount, 4) = Format(rs!SalingPrice, "#,##0")
                    .TextMatrix(incount, 5) = Format(rs!CostPrice, "#,##0")
                    .TextMatrix(incount, 6) = Format(rs!Profit, "#,##0")
                    .CellBackColor = 0
                    
                End With
            rs.MoveNext
            With prbprocess
                If .Value < .Max - (1000 / rs.RecordCount) Then
                    .Value = .Value + 1000 / rs.RecordCount
                Else
                    .Value = .Max
                End If
            End With
            Loop
        End With
        If rs.RecordCount = 0 Then
            With flgProfit
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
                .TextMatrix(1, 4) = ""
                .TextMatrix(1, 5) = ""
                .TextMatrix(1, 6) = ""
            End With
        End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "SetFLGRIDORDER"
End Sub
