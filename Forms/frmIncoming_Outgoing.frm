VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmIncoming_Outgoing 
   Caption         =   "TÊng hÓp thu chi"
   ClientHeight    =   9480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10440
   ClipControls    =   0   'False
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
   ScaleHeight     =   13506.83
   ScaleMode       =   0  'User
   ScaleWidth      =   22422.07
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid dtgNCC 
      Height          =   2415
      Left            =   8040
      TabIndex        =   23
      Top             =   1800
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4260
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
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
   Begin MSDataGridLib.DataGrid dtgrid 
      Height          =   2895
      Left            =   8280
      TabIndex        =   20
      Top             =   1080
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   5106
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   20
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
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
   Begin VB.Frame Frame4 
      Height          =   7695
      Left            =   0
      TabIndex        =   12
      Top             =   1920
      Width           =   15135
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   10680
         TabIndex        =   16
         Text            =   "0"
         Top             =   7080
         Width           =   4215
      End
      Begin MSFlexGridLib.MSFlexGrid flgDetail 
         Height          =   6735
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   11880
         _Version        =   393216
         BackColorFixed  =   -2147483643
         BackColorBkg    =   -2147483643
         GridColorFixed  =   12632256
         GridLinesFixed  =   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "TÊng cÈng:"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   8400
         TabIndex        =   15
         Top             =   7200
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1695
      Left            =   8040
      TabIndex        =   9
      Top             =   240
      Width           =   7095
      Begin VB.TextBox txtMaNCC 
         Height          =   375
         Left            =   3600
         TabIndex        =   22
         Top             =   1080
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cboNCC 
         Height          =   345
         Left            =   120
         TabIndex        =   21
         Text            =   "T t c∂ nhµ cung c p"
         Top             =   1080
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txtma 
         Height          =   375
         Left            =   3600
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cboType 
         Height          =   345
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   3375
      End
      Begin prjTouchScreen.MyButton cmdExit 
         Height          =   615
         Left            =   5760
         TabIndex        =   14
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "ß„n&g"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmIncoming_Outgoing.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdView 
         Height          =   615
         Left            =   4560
         TabIndex        =   11
         Top             =   960
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&In B∏o c∏o"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmIncoming_Outgoing.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdDone 
         Height          =   615
         Left            =   4560
         TabIndex        =   10
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1085
         BTYPE           =   14
         TX              =   "&Th˘c thi"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmIncoming_Outgoing.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.Label lblNCC 
         Caption         =   "Ch‰n nhµ cung c p"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblType 
         Caption         =   "loπi"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Loπi"
      Height          =   1695
      Left            =   4440
      TabIndex        =   5
      Top             =   240
      Width           =   3375
      Begin VB.OptionButton OptTonquy 
         Caption         =   "TÂn qu¸"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.OptionButton OptChi 
         Caption         =   "TÊng hÓp chi"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton OptThu 
         Caption         =   "TÊng hÓp thu"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ThÍi gian"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   495
         Left            =   1500
         TabIndex        =   1
         Top             =   270
         Width           =   2415
         _ExtentX        =   4260
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
         Format          =   63176705
         UpDown          =   -1  'True
         CurrentDate     =   39448
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   495
         Left            =   1500
         TabIndex        =   2
         Top             =   960
         Width           =   2415
         _ExtentX        =   4260
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
         Format          =   63176705
         UpDown          =   -1  'True
         CurrentDate     =   39448
      End
      Begin VB.Label lblFromdate 
         Alignment       =   1  'Right Justify
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
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label lblDenngay 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   3
         Top             =   1050
         Width           =   1245
      End
   End
   Begin VB.Label lblreadnum 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2280
      TabIndex        =   26
      Top             =   9720
      Width           =   12375
   End
   Begin VB.Label Label2 
      Caption         =   "SË ti“n bªng ch˜:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   25
      Top             =   9720
      Width           =   2055
   End
End
Attribute VB_Name = "frmIncoming_Outgoing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPhieuthu As New ADODB.Recordset
Dim rsPhieuChi As New ADODB.Recordset
Dim rsTonghop As New ADODB.Recordset
Dim optSelect As Integer
Dim Total As Double
Dim rsma As New ADODB.Recordset
Dim rsNCC As New ADODB.Recordset

Private Sub cboNCC_Change()
On Error GoTo Handle
Dim str As String

With rsNCC
    If .State = adStateOpen Then .Close
    If InStr(1, cboNCC.Text, "*", vbTextCompare) > 0 Then
        str = "SELECT  Vendor_Number, Vendor_Name FROM Vendors WHERE INSTR(Vendor_Number,""" & Left(cboNCC.Text, Len(Trim(cboNCC.Text)) - 1) & "%"")>0 OR Vendor_Name LIKE '" & _
        Left(Trim(cboNCC.Text), Len(Trim(cboNCC.Text)) - 1) & "%'  ORDER BY Vendor_Number asc"

    Else
        str = "SELECT  Vendor_Number, Vendor_Name FROM Vendors WHERE (INSTR(Vendor_Number,""" & Trim(cboNCC.Text) & """)>0 OR INSTR(Vendor_Name,""" & _
        Trim(cboNCC.Text) & """)>0) ORDER BY Vendor_Number"
    End If
End With
Set rsNCC = OpenCriticalTable(str, cnData)
Set dtgNCC.DataSource = rsNCC
dtgNCC.Visible = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub cboNCC_DropDown()
    Call cboNCC_Change
End Sub

Private Sub cboNCC_KeyDown(KeyCode As Integer, Shift As Integer)
    Call cboNCC_Change
End Sub

Private Sub cboType_Change()
    If OptThu.Value = True Then
        dtgrid.Visible = True
        Call set_Receipt_Cbo
    ElseIf OptChi.Value = True Then
        dtgrid.Visible = True
        Call set_Expense_Cbo
    Else
    
    End If
End Sub

Private Sub cboType_Click()
    Call cboType_Change
End Sub

Private Sub cboType_DropDown()
    Call cboType_Change
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyTab Then dtgrid.SetFocus
End Sub

Private Sub cmdDone_Click()
On Error GoTo Handle
    If dtpFromDate.Value > dtpToDate.Value Then dtpFromDate.Value = dtpToDate.Value
    Select Case optSelect
        Case 0
            If txtma.Text = "" Then
                Call ViewInCom(gfCONVERT_DATE_TO_STRING(dtpFromDate.Value), gfCONVERT_DATE_TO_STRING(dtpToDate.Value))
            Else
                Call ViewInCom_By_Type(gfCONVERT_DATE_TO_STRING(dtpFromDate.Value), gfCONVERT_DATE_TO_STRING(dtpToDate.Value))
            End If
        Case 1
            If txtMaNCC.Text = "" Then
                If txtma.Text = "" Then
                    Call ViewOutCom(gfCONVERT_DATE_TO_STRING(dtpFromDate.Value), gfCONVERT_DATE_TO_STRING(dtpToDate.Value))
                Else
                    Call ViewOutCom_By_Type(gfCONVERT_DATE_TO_STRING(dtpFromDate.Value), gfCONVERT_DATE_TO_STRING(dtpToDate.Value))
                End If
            Else
                If txtma.Text = "" Then
                    Call ViewOutCom_By_NCC(gfCONVERT_DATE_TO_STRING(dtpFromDate.Value), gfCONVERT_DATE_TO_STRING(dtpToDate.Value))
                Else
                    Call ViewOutCom_By_Type_NCC(gfCONVERT_DATE_TO_STRING(dtpFromDate.Value), gfCONVERT_DATE_TO_STRING(dtpToDate.Value))
                End If
            
            End If
        Case 2
        
    End Select
    txtma.Text = ""
    txtMaNCC.Text = ""
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdView_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim iReport As CRAXDDRT.Report
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Select Case optSelect
        Case 0
        
            If txtma.Text = "" Then
                SQL = "SELECT Income.Cashier_ID, Income.DateTime, Income.Customer_ID," & _
                    " Customer.CustName, Income.Receipt_ID, Receipt.DienGiai, Income.Reciever_Name," & _
                    " Income.Amount, Income.Payment_Method" & _
                    " FROM Customer INNER JOIN (Receipt INNER JOIN Income ON Receipt.MaThu" & _
                    " = Income.Receipt_ID) ON Customer.CustNum = Income.Customer_ID" & _
                    " where DateTime>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "'" & _
                    " and DateTime <='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'"
            Else
                SQL = "SELECT Income.Cashier_ID, Income.DateTime, Income.Customer_ID," & _
                    " Customer.CustName, Income.Receipt_ID, Receipt.DienGiai, Income.Reciever_Name," & _
                    " Income.Amount, Income.Payment_Method" & _
                    " FROM Customer INNER JOIN (Receipt INNER JOIN Income ON Receipt.MaThu" & _
                    " = Income.Receipt_ID) ON Customer.CustNum = Income.Customer_ID" & _
                    " where DateTime>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "'" & _
                    " and DateTime <='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'" & _
                    " and Income.Receipt_ID='" & txtma.Text & "'"
                    
            End If
    
            Set crThu = Nothing
                cmd.ActiveConnection = cnData
                cmd.CommandText = SQL
                cmd.Execute
            With crThu
                .Database.AddADOCommand cnData, cmd
                .txtReceiptName.SetUnboundFieldSource "{ado.DienGiai}"
                .txtCustomer.SetUnboundFieldSource "{ado.CustName}"
                .txtPaymentMethod.SetUnboundFieldSource "{ado.Payment_Method}"
                .txtNguoinop.SetUnboundFieldSource "{ado.Reciever_Name}"
                .txtAmount.SetUnboundFieldSource "{ado.Amount}"
        
                .txtFromDate.SetText dtpFromDate.Value
                .txtToDate.SetText dtpToDate.Value
                
                With .txtAmount
                    .DecimalPlaces = DecimalAmtNumber
                    .DecimalSymbol = DecimalMark
                    .ThousandsSeparators = True
                    .ThousandSymbol = DigitGroupMark
        
                End With
                With .txtsumAmt
                    .DecimalPlaces = DecimalAmtNumber
                    .DecimalSymbol = DecimalMark
                    .ThousandsSeparators = True
                    .ThousandSymbol = DigitGroupMark
        
                End With
        
            End With
            
            Set iReport = crThu
            
        Case 1
            SQL = "SELECT Expense.MaChi, Expense.DienGiai, Payouts.DateTime," & _
                  " Payouts.Vendor_Number, Payouts.Amount, Payouts.Description," & _
                  " Payouts.Payment_Method, Payouts.Cashier_ID, Vendors.Vendor_Name" & _
                  " FROM Expense INNER JOIN (Vendors INNER JOIN Payouts ON Vendors.Vendor_Number" & _
                  " = Payouts.Vendor_Number) ON Expense.MaChi = Payouts.Expense_ID" & _
                  " where DateTime>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "'" & _
                  " and DateTime <='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'"
            
            Set crChi = Nothing
                cmd.ActiveConnection = cnData
                cmd.CommandText = SQL
                cmd.Execute
            With crChi
                .Database.AddADOCommand cnData, cmd
                .txtReceiptName.SetUnboundFieldSource "{ado.DienGiai}"
                .txtCustomer.SetUnboundFieldSource "{ado.Vendor_Name}"
                .txtDescription.SetUnboundFieldSource "{ado.Description}"
                .txtNguoinop.SetUnboundFieldSource "{ado.Cashier_ID}"
                .txtAmount.SetUnboundFieldSource "{ado.Amount}"
                .txtDate.SetUnboundFieldSource "{ado.DateTime}"
        
                .txtFromDate.SetText dtpFromDate.Value
                .txtToDate.SetText dtpToDate.Value
                
                With .txtAmount
                    .DecimalPlaces = DecimalAmtNumber
                    .DecimalSymbol = DecimalMark
                    .ThousandsSeparators = True
                    .ThousandSymbol = DigitGroupMark
        
                End With
                With .txtsumAmt
                    .DecimalPlaces = DecimalAmtNumber
                    .DecimalSymbol = DecimalMark
                    .ThousandsSeparators = True
                    .ThousandSymbol = DigitGroupMark
        
                End With
        
            End With
            
            Set iReport = crChi
        
    End Select
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdChi_Click"
End Sub

Private Sub dtgNCC_Click()
    With rsNCC
        txtMaNCC.Text = !Vendor_Number
        cboNCC.Text = !Vendor_Name
        dtgNCC.Visible = False
    End With
End Sub

Private Sub dtgNCC_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl
    If KeyAscii = 13 Then
        With rsNCC
            If .RecordCount = 0 Then
                dtgNCC.Visible = False
                cboNCC.SetFocus
                Exit Sub
            End If
            txtMaNCC.Text = !Vendor_Number
            cboNCC.Text = !Vendor_Name
            dtgNCC.Visible = False
        End With
    ElseIf KeyAscii = 9 Then
        dtgNCC.Visible = False
        cboNCC.SetFocus
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - dtgrid_KeyPress "
End Sub

Private Sub dtgrid_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl
    If KeyAscii = 13 Then
        With rsma
            If .RecordCount = 0 Then
                dtgrid.Visible = False
                cboType.SetFocus
                Exit Sub
            End If
            If OptThu.Value = True Then
                txtma.Text = !MaThu
                cboType.Text = !DienGiai
            ElseIf OptChi.Value = True Then
                txtma.Text = !maChi
                cboType.Text = !DienGiai
            End If
            dtgrid.Visible = False
        End With
    ElseIf KeyAscii = 9 Then
        dtgrid.Visible = False
        cboType.SetFocus
    End If
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - dtgrid_KeyPress "
End Sub

Private Sub Form_Activate()
On Error GoTo errHdl
    Dim DescArr() As String
    Dim ctrl As Control
    DescArr = LoadLanguage(LngFile, "#03:017:")
    If cmdDone.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Activate"


End Sub

Private Sub Form_Load()
On Error GoTo Handle
    dtpFromDate.Value = "01/" & Mid(DateDefault, 5, 2) & "/" & Left(DateDefault, 4)
    dtpToDate.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    OptThu.Value = True
    Call Set_flgInOut
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Form_Load"
End Sub

Private Sub OptChi_Click()
    optSelect = 1
    Set_flgOut
    lblType.Caption = "Ch‰n kho∂n chi"
    cboNCC.Visible = True
    lblNCC.Visible = True
End Sub

Private Sub OptThu_Click()
    optSelect = 0
    Set_flgInOut
    lblType.Caption = "Ch‰n kho∂n thu"
    cboNCC.Visible = False
    dtgNCC.Visible = False
    lblNCC.Visible = False
End Sub

Private Sub OptTonquy_Click()
    optSelect = 2
    Set_flgInOut
    cboType.Enabled = False
End Sub

Public Sub ViewInCom(ByVal Tungay As String, ByVal denngay As String)
On Error GoTo Handle
Dim strSql As String
Dim rsThu As New ADODB.Recordset
strSql = "SELECT Income.ID as ID,Income.Cashier_ID, Income.DateTime, Income.Customer_ID as CustID, Receipt.MaThu," & _
         " Receipt.DienGiai  as Receipt_Des, Income.Reciever_Name, Income.Division, Income.Amount," & _
         " Income.Description, Income.Payment_Method" & _
         " FROM Receipt INNER JOIN Income ON Receipt.MaThu = Income.Receipt_ID" & _
         " where Income.DateTime>='" & Tungay & "' and Income.DateTime<='" & denngay & "'"
    Set rsThu = OpenCriticalTable(strSql, cnData)
    If rsThu.State > 0 And rsThu.RecordCount > 0 Then
        Call SetFLGRIDINOUT(rsThu)
    Else
        Call Set_flgInOut
    End If
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & "- Lay du lieu thu"
End Sub

Public Sub ViewInCom_By_Type(ByVal Tungay As String, ByVal denngay As String)
On Error GoTo Handle
Dim strSql As String
Dim rsThu As New ADODB.Recordset
strSql = "SELECT Income.ID as ID,Income.Cashier_ID, Income.DateTime, Income.Customer_ID as CustID, Receipt.MaThu," & _
         " Receipt.DienGiai  as Receipt_Des, Income.Reciever_Name, Income.Division, Income.Amount," & _
         " Income.Description, Income.Payment_Method" & _
         " FROM Receipt INNER JOIN Income ON Receipt.MaThu = Income.Receipt_ID" & _
         " where Income.DateTime>='" & Tungay & "' and Income.DateTime<='" & denngay & "' and Income.Receipt_ID='" & txtma.Text & "'"
    Set rsThu = OpenCriticalTable(strSql, cnData)
    If rsThu.State > 0 And rsThu.RecordCount > 0 Then
        Call SetFLGRIDINOUT(rsThu)
    Else
        Call Set_flgInOut
    End If
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & "- Lay du lieu thu"
End Sub

Public Sub ViewOutCom(ByVal Tungay As String, ByVal denngay As String)
On Error GoTo Handle
Dim strSql As String
Dim rsChi As New ADODB.Recordset
strSql = "SELECT Payouts.ID as ID, Payouts.DateTime, Payouts.Vendor_Number as CustID, Expense.DienGiai as Receipt_Des," & _
        " Payouts.Amount, Payouts.Description" & _
        " FROM Expense INNER JOIN Payouts ON Expense.MaChi = Payouts.Expense_ID" & _
        " where Payouts.DateTime>='" & Tungay & "' and Payouts.DateTime<='" & denngay & "'"
    Set rsChi = OpenCriticalTable(strSql, cnData)
    If rsChi.State > 0 And rsChi.RecordCount > 0 Then
        Call SetFLGRIDINOUT(rsChi)
    Else
        Call Set_flgOut
    End If
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & "- Lay du lieu thu"
End Sub

Public Sub ViewOutCom_By_Type(ByVal Tungay As String, ByVal denngay As String)
On Error GoTo Handle
Dim strSql As String
Dim rsChi As New ADODB.Recordset
strSql = "SELECT Payouts.ID as ID, Payouts.DateTime, Payouts.Vendor_Number as CustID, Expense.DienGiai as Receipt_Des," & _
        " Payouts.Amount, Payouts.Description" & _
        " FROM Expense INNER JOIN Payouts ON Expense.MaChi = Payouts.Expense_ID" & _
        " where Payouts.DateTime>='" & Tungay & "' and Payouts.DateTime<='" & denngay & "' and Expense_ID='" & txtma.Text & "'"
    Set rsChi = OpenCriticalTable(strSql, cnData)
    If rsChi.State > 0 And rsChi.RecordCount > 0 Then
        Call SetFLGRIDINOUT(rsChi)
    Else
        Call Set_flgOut
    End If
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & "- Lay du lieu thu"
End Sub

Public Sub ViewOutCom_By_NCC(ByVal Tungay As String, ByVal denngay As String)
On Error GoTo Handle
Dim strSql As String
Dim rsChi As New ADODB.Recordset
strSql = "SELECT Payouts.ID as ID, Payouts.DateTime, Payouts.Vendor_Number as CustID, Expense.DienGiai as Receipt_Des," & _
        " Payouts.Amount, Payouts.Description" & _
        " FROM Expense INNER JOIN Payouts ON Expense.MaChi = Payouts.Expense_ID" & _
        " where Payouts.DateTime>='" & Tungay & "' and Payouts.DateTime<='" & denngay & "' and Vendor_Number='" & txtMaNCC.Text & "'"
    Set rsChi = OpenCriticalTable(strSql, cnData)
    If rsChi.State > 0 And rsChi.RecordCount > 0 Then
        Call SetFLGRIDINOUT(rsChi)
    Else
        Call Set_flgOut
    End If
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & "- Lay du lieu thu"
End Sub

Public Sub ViewOutCom_By_Type_NCC(ByVal Tungay As String, ByVal denngay As String)
On Error GoTo Handle
Dim strSql As String
Dim rsChi As New ADODB.Recordset
strSql = "SELECT Payouts.ID as ID, Payouts.DateTime, Payouts.Vendor_Number as CustID, Expense.DienGiai as Receipt_Des," & _
        " Payouts.Amount, Payouts.Description" & _
        " FROM Expense INNER JOIN Payouts ON Expense.MaChi = Payouts.Expense_ID" & _
        " where Payouts.DateTime>='" & Tungay & "' and Payouts.DateTime<='" & denngay & "' and Expense_ID='" & txtma.Text & "' and Vendor_Number='" & txtMaNCC.Text & "'"
    Set rsChi = OpenCriticalTable(strSql, cnData)
    If rsChi.State > 0 And rsChi.RecordCount > 0 Then
        Call SetFLGRIDINOUT(rsChi)
    Else
        Call Set_flgOut
    End If
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & "- Lay du lieu thu"
End Sub

Public Sub Set_flgInOut()
    On Error GoTo Handle
    Dim i As Integer
        With flgDetail
            .Cols = 6
            .Rows = 20
            .ColWidth(0) = 2000
            .ColWidth(1) = 1800
            .ColWidth(2) = 3500
            .ColWidth(3) = 2000
            .ColWidth(4) = 2000
            .ColWidth(5) = 4000
            .TextMatrix(0, 0) = "SË phi’u"
            .TextMatrix(0, 1) = "Ngµy" ' "SÙ' luong"
            .TextMatrix(0, 2) = "Kh∏ch hµng" '" D/Gi·"
            .TextMatrix(0, 3) = "Kho∂n thu" '"T/TiÍn`"
            .TextMatrix(0, 4) = "SË ti“n"
            .TextMatrix(0, 5) = "Di‘n gi∂i"
            .ColAlignment(1) = 2
            .ColAlignment(2) = 4
            .ColAlignment(3) = 6
            .ColAlignment(4) = 6
            i = 1
            Do While i < .Rows - 1
                .TextMatrix(i, 0) = ""
                .TextMatrix(i, 1) = ""
                .TextMatrix(i, 2) = ""
                .TextMatrix(i, 3) = ""
                .TextMatrix(i, 4) = ""
                .TextMatrix(i, 5) = ""
            i = i + 1
            Loop
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"

End Sub

Public Sub Set_flgOut()
    On Error GoTo Handle
    Dim i As Integer
        With flgDetail
            .Cols = 6
            .Rows = 20
            .ColWidth(0) = 2000
            .ColWidth(1) = 1800
            .ColWidth(2) = 3500
            .ColWidth(3) = 2000
            .ColWidth(4) = 2000
            .ColWidth(5) = 4000
            .TextMatrix(0, 0) = "SË phi’u"
            .TextMatrix(0, 1) = "Ngµy" ' "SÙ' luong"
            .TextMatrix(0, 2) = "nhµ cung c p" '" D/Gi·"
            .TextMatrix(0, 3) = "Kho∂n chi" '"T/TiÍn`"
            .TextMatrix(0, 4) = "SË ti“n"
            .TextMatrix(0, 5) = "Di‘n gi∂i"
            .ColAlignment(1) = 2
            .ColAlignment(2) = 4
            .ColAlignment(3) = 6
            .ColAlignment(4) = 6
            i = 1
            Do While i < .Rows - 1
                .TextMatrix(i, 0) = ""
                .TextMatrix(i, 1) = ""
                .TextMatrix(i, 2) = ""
                .TextMatrix(i, 3) = ""
                .TextMatrix(i, 4) = ""
                .TextMatrix(i, 5) = ""
            i = i + 1
            Loop
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"

End Sub

Public Sub SetFLGRIDINOUT(rs As ADODB.Recordset)
On Error GoTo Handle
Dim incount As Integer
        Total = 0
        If rs.EOF Then Exit Sub
        rs.MoveFirst
        With rs
            .Sort = "DateTime ASC"
            Do While Not .EOF
                incount = incount + 1
                flgDetail.Rows = rs.RecordCount + 1
                With flgDetail
                    .TextMatrix(incount, 0) = rs!ID
                    .TextMatrix(incount, 1) = rs!DateTime
                    .TextMatrix(incount, 2) = rs!CustID
                    .TextMatrix(incount, 3) = rs!Receipt_Des
                    .TextMatrix(incount, 4) = Format(rs!Amount, "#,##0")
                    .TextMatrix(incount, 5) = "" & rs!Description
                    .CellBackColor = 0
                    
                End With
                Total = Total + CDbl(rs!Amount)
            rs.MoveNext
            Loop
        End With
        If rs.RecordCount = 0 Then
            With flgDetail
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
                .TextMatrix(1, 4) = ""
                .TextMatrix(1, 5) = ""
            
            End With
        End If
        flgDetail.Row = flgDetail.Rows - 1
        TxtTotal.Text = Format(Total, "#,##0")
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "SetFLGRIDINOUT"
End Sub

Public Sub set_Expense_Cbo()
On Error GoTo Handle
Dim str As String
'Set rsExpense = Open_Table(cnData, "Expense")
'If rsExpense.State <> 0 And rsExpense.RecordCount > 0 Then rsExpense.MoveFirst
With rsma
    If .State = adStateOpen Then .Close
    If InStr(1, cboType.Text, "*", vbTextCompare) > 0 Then
        str = "SELECT  MaChi, Diengiai FROM Expense WHERE INSTR(MaChi,""" & Left(cboType.Text, Len(Trim(cboType.Text)) - 1) & "%"")>0 OR DienGiai LIKE '" & _
        Left(Trim(cboType.Text), Len(Trim(cboType.Text)) - 1) & "%'  ORDER BY MaChi asc"

    Else
        str = "SELECT  MaChi, DienGiai FROM Expense WHERE (INSTR(MaChi,""" & Trim(cboType.Text) & """)>0 OR INSTR(DienGiai,""" & _
        Trim(cboType.Text) & """)>0) ORDER BY MaChi"
    End If
End With
Set rsma = OpenCriticalTable(str, cnData)
Set dtgrid.DataSource = rsma
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub

Public Sub set_Receipt_Cbo()
On Error GoTo Handle
Dim str As String

With rsma
    If .State = adStateOpen Then .Close
    If InStr(1, cboType.Text, "*", vbTextCompare) > 0 Then
        str = "SELECT  MaThu, Diengiai FROM Receipt WHERE INSTR(MaThu,""" & Left(cboType.Text, Len(Trim(cboType.Text)) - 1) & "%"")>0 OR DienGiai LIKE '" & _
        Left(Trim(cboType.Text), Len(Trim(cboType.Text)) - 1) & "%'  ORDER BY MaThu asc"
    Else
        str = "SELECT  MaThu, DienGiai FROM Receipt WHERE (INSTR(MaThu,""" & Trim(cboType.Text) & """)>0 OR INSTR(DienGiai,""" & _
        Trim(cboType.Text) & """)>0) ORDER BY MaThu"
    End If
End With
Set rsma = OpenCriticalTable(str, cnData)
Set dtgrid.DataSource = rsma
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " set_Receipt_Cbo"
End Sub

Private Sub txtTotal_Change()
    lblreadnum.Caption = readnumber(TxtTotal.Text)
End Sub
