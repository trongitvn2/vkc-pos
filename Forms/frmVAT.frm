VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmVAT 
   Caption         =   "KÕt xuÊt H§ VAT"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15240
   ClipControls    =   0   'False
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   11055
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmdlpath 
      Left            =   11400
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjTouchScreen.MyButton cmdFilter 
      Height          =   1095
      Left            =   5520
      TabIndex        =   18
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1931
      BTYPE           =   3
      TX              =   "Läc"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   14.25
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmVAT.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Frame Frame2 
      Caption         =   "Danh s¸ch H§ VAT "
      ForeColor       =   &H00FF0000&
      Height          =   9375
      Left            =   7800
      TabIndex        =   16
      Top             =   1320
      Width           =   7455
      Begin MSFlexGridLib.MSFlexGrid flgVAT 
         Height          =   8775
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   15478
         _Version        =   393216
         ForeColor       =   16711680
         ForeColorFixed  =   -2147483635
         ForeColorSel    =   16711680
      End
   End
   Begin VB.CheckBox chkVAT 
      Caption         =   "VAT"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10800
      TabIndex        =   15
      Top             =   360
      Width           =   2175
   End
   Begin prjTouchScreen.MyButton cmdExport 
      Height          =   615
      Left            =   13200
      TabIndex        =   14
      Top             =   0
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   "KÕt xuÊt d÷ liÖu"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmVAT.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.TextBox txtAmount 
      Height          =   390
      Left            =   7680
      TabIndex        =   13
      Top             =   840
      Width           =   2055
   End
   Begin VB.TextBox txtDate 
      Height          =   390
      Left            =   10800
      TabIndex        =   11
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtInvoice 
      Height          =   390
      Left            =   7680
      TabIndex        =   9
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Danh s¸ch H§ b¸n hµng"
      ForeColor       =   &H00FF0000&
      Height          =   9375
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   7815
      Begin MSFlexGridLib.MSFlexGrid flgSale 
         Height          =   8775
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   15478
         _Version        =   393216
      End
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   13200
      TabIndex        =   0
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   "§ãn&g"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmVAT.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpFDate 
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   72941569
      UpDown          =   -1  'True
      CurrentDate     =   39448
   End
   Begin MSComCtl2.DTPicker dtpTDate 
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   480
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   72941569
      UpDown          =   -1  'True
      CurrentDate     =   39448
   End
   Begin VB.Label Label6 
      Caption         =   "Sè tiÒn:"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6720
      TabIndex        =   12
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Ngµy:"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10080
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "H§ sè:"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   360
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "§Õn"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Tõ:"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Chän d÷ ngµy d÷ liÖu kÕt xuÊt H§"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
End
Attribute VB_Name = "frmVAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsInvoice As New ADODB.Recordset
Dim rsInvoice_temp As New ADODB.Recordset
Dim rsInvoice_Totals As New ADODB.Recordset
Dim strSql As String

Private Sub chkVAT_Click()
On Error GoTo Handle
    With rsInvoice_Totals
    .MoveFirst
    .Find "Invoice_Number=" & txtInvoice.Text, , adSearchForward, adBookmarkFirst
    'MsgBox .EOF
    If Not .EOF Then
        If chkVAT.Value = 1 Then
            .Fields("IsVAT") = True
            .Update
        Else
            .Fields("IsVAT") = False
            .Update
        End If
    End If
End With
Call cmdFilter_Click
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " chkVAT_Click"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
On Error GoTo Handle
Dim fso As New FileSystemObject
Dim path_Save As String
Dim cnSave_VAT As New ADODB.Connection
    With cmdlpath
        .ShowOpen
        path_Save = .FileName
    End With
    If Right(path_Save, 13) = "\Database.mdb" Then path_Save = Left(path_Save, Len(path_Save) - 13)
    If Dir(path_Save & "\Database.mdb", vbDirectory) = "" Then
        MkDir path_Save
        MkDir path_Save & "\Log"
        fso.CopyFile WorkingFolder & "\Database.mdb", path_Save & "\Database.mdb", True
        fso.CopyFile WorkingFolder & "\LoginData.Dat", path_Save & "\LoginData.Dat", True
        Set cnSave_VAT = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    'Delete all data
    cnSave_VAT.Execute "Delete  from Invoice_Totals_Notes"
    cnSave_VAT.Execute "Delete  from Items_Deleted"
    End If
   If cnSave_VAT.State = 0 Then Set cnSave_VAT = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    
    Call Backup_VAT(cnSave_VAT, path_Save)
    MsgBox "§· hoµn thµnh"
Set cnSave_VAT = Nothing
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " -cmdExport_Click "
End Sub

Private Sub cmdFilter_Click()
On Error GoTo Handle
    strSql = "SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.DateTime, IsVAT, Invoice_Totals.Grand_Total, Invoice_Totals.Orig_OnHoldID" & _
                    " From Invoice_Totals" & _
                    " WHERE (((Left([DateTime],8))>='" & gfCONVERT_DATE_TO_STRING(dtpFDate.Value) & "' And (Left([DateTime],8))<='" & gfCONVERT_DATE_TO_STRING(dtpTDate.Value) & "'))" & _
                    " and IsVAT=false and IsgetVAT=false" & _
                    " GROUP BY Invoice_Totals.Invoice_Number,IsVAT, Invoice_Totals.DateTime, Invoice_Totals.Grand_Total, Invoice_Totals.Orig_OnHoldID"
    Set rsInvoice = OpenCriticalTable(strSql, cnData)
    Call setflgInvoice(rsInvoice, flgSale)
    Call Set_FlgVAT
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdFilter_Click"
End Sub

Private Sub flgSale_Click()
On Error GoTo Handle
'    txtInvoice.Text = flgSale.TextMatrix(flgSale.Row, 1)
     txtInvoice.Text = flgSale.TextMatrix(flgSale.Row, 1)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " flgSale_Click"
End Sub

Private Sub flgVAT_Click()
    txtInvoice.Text = flgVAT.TextMatrix(flgVAT.Row, 1)
End Sub

Private Sub Form_Load()
On Error GoTo Handle
Set rsInvoice_Totals = Open_Table(cnData, "Invoice_Totals")
If Not Check_Field_Exist(rsInvoice_Totals, "IsVAT") Then
    cnData.Execute "ALTER TABLE Invoice_Totals ADD COLUMN IsVAT YesNo, IsgetVAT YesNo "
    
End If
    dtpFDate.Value = Format(Date, "dd/MM/yyyy")
    dtpTDate.Value = Format(Date, "dd/MM/yyyy")
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Private Sub setflgInvoice(rs As Recordset, flg As MSFlexGrid)
On Error GoTo errHdl
    Dim intCount    As Integer
    With flg
        .Font = ".vnArial"
        .Rows = rs.RecordCount + 1
        .Cols = 5
        .ColWidth(0) = 1600
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 2000
        .TextMatrix(0, 0) = "Ngµy"
        .TextMatrix(0, 1) = "Sè H§"
        .TextMatrix(0, 2) = "Bµn sè"
        .TextMatrix(0, 3) = "Sè tiÒn"
        .TextMatrix(0, 4) = "VAT"
    End With
    
    If rs Is Nothing Or rs.RecordCount = 0 Then Exit Sub
    If rs.State = 0 Then Exit Sub
    rs.MoveFirst
   flg.Rows = rs.RecordCount + 1
    intCount = 0
    Do While Not rs.EOF
        intCount = intCount + 1
        flg.TextMatrix(intCount, 0) = gfCONVERT_STRING_TO_DATE(Left(rs!DateTime, 8))
        flg.TextMatrix(intCount, 1) = rs!Invoice_Number
        flg.TextMatrix(intCount, 2) = rs!Orig_OnHoldID
        flg.TextMatrix(intCount, 3) = Format(rs!Grand_Total, "#,##0")
        flg.TextMatrix(intCount, 4) = rs!IsVAT
        rs.MoveNext
    Loop
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - setflgInvoice "
End Sub


Public Sub Set_FlgVAT()
On Error GoTo Handle
Dim str As String
str = "SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.DateTime, IsVAT, Invoice_Totals.Grand_Total, Invoice_Totals.Orig_OnHoldID" & _
                    " From Invoice_Totals" & _
                    " WHERE (((Left([DateTime],8))>='" & gfCONVERT_DATE_TO_STRING(dtpFDate.Value) & "' And (Left([DateTime],8))<='" & gfCONVERT_DATE_TO_STRING(dtpTDate.Value) & "'))" & _
                    " and IsVAT=true and IsgetVAT=False" & _
                    " GROUP BY Invoice_Totals.Invoice_Number,IsVAT, Invoice_Totals.DateTime, Invoice_Totals.Grand_Total, Invoice_Totals.Orig_OnHoldID"
    Set rsInvoice_temp = OpenCriticalTable(str, cnData)
    Call setflgInvoice(rsInvoice_temp, flgVAT)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Set_FlgVAT"
End Sub

Private Sub txtInvoice_Change()
On Error GoTo Handle
    With rsInvoice_Totals
            .Find "Invoice_Number=" & txtInvoice.Text, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtInvoice.Text = .Fields("Invoice_number")
                txtDate.Text = gfCONVERT_STRING_TO_DATE(Left(.Fields("DateTime"), 8))
                txtAmount.Text = Format(.Fields("Grand_Total"), "#,##0")
                If .Fields("IsVAT") = True Then
                    chkVAT.Value = 1
                Else
                    chkVAT.Value = 0
                End If
            End If
        End With
    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  txtInvoice_Change"

End Sub

Public Sub Backup_VAT(cnSave_VAT As ADODB.Connection, path_Save As String)
On Error GoTo Handle
Dim j As Integer
Dim Invoice_Num As Double
Dim ans As Integer
    With rsInvoice_temp
        If Not .EOF And .RecordCount > 0 Then .MoveFirst
                j = Get_MaxInvoice_VAT(cnSave_VAT)
                .Sort = "Invoice_Number ASC"
                Do While Not .EOF
                    'Invoice_Num = Right("0000" & j, 4)
                    Invoice_Num = GetMax_Invoice_Backup(path_Save)
                    Call gfBackup_Invoice_Notes(cnSave_VAT, cnData, .Fields("Invoice_Number"), Invoice_Num)
                    Call gfBackup_Invoice_Totals(cnSave_VAT, cnData, .Fields("Invoice_Number"), Invoice_Num)
                    Call gfBackup_Invoice_Itemized(cnSave_VAT, cnData, .Fields("Invoice_Number"), Invoice_Num)
                    Call gfBackup_Deleted_Item(cnSave_VAT, cnData, .Fields("Invoice_Number"), Invoice_Num)
                    Call Mark_IsVAT(.Fields("Invoice_Number"))
                    If ans = 0 Then
                        If MsgBox("B¹n cã muèn xãa nh÷ng bill nµy trong d÷ liÖu chÝnh kh«ng?", vbYesNo) = vbYes Then
                             cnData.Execute "Delete  from invoice_Totals where Invoice_Number=" & .Fields("Invoice_Number")
                             ans = 1
                         Else
                             ans = 2
                         End If
                    ElseIf ans = 1 Then
                        cnData.Execute "Delete  from invoice_Totals where Invoice_Number=" & .Fields("Invoice_Number")
                    End If
                    
                .MoveNext
                j = j + 1
            Loop
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Backup_VAT"
End Sub
Public Sub gfBackup_Invoice_Totals(cnBackup As Connection, cnOrg As Connection, invoice_Num_Org As Double, ByVal Invoice_Num As Double)
On Error GoTo Handle
    Dim rsInvoice_Totals_Org As New ADODB.Recordset
    Dim rsInvoice_Totals_Des As New ADODB.Recordset
    Dim i As Integer
        Set rsInvoice_Totals_Org = OpenCriticalTable("Select * from Invoice_Totals where Invoice_Number=" & invoice_Num_Org, cnOrg)
        Set rsInvoice_Totals_Des = Open_Table(cnBackup, "Invoice_Totals")
        With rsInvoice_Totals_Org
            i = 0
            Do While Not .EOF
                With rsInvoice_Totals_Des
                    .addNew
                    .Fields("Invoice_Number") = Invoice_Num
                    .Fields("Store_ID") = rsInvoice_Totals_Org.Fields("Store_ID")
                    .Fields("CustNum") = rsInvoice_Totals_Org.Fields("CustNum")
                    .Fields("DateTime") = rsInvoice_Totals_Org.Fields("DateTime")
                    .Fields("Total_Cost") = rsInvoice_Totals_Org.Fields("Total_Cost")
                    .Fields("Discount") = rsInvoice_Totals_Org.Fields("Discount")
                    .Fields("KarDiscount") = rsInvoice_Totals_Org.Fields("KarDiscount")
                    .Fields("Total_Price") = rsInvoice_Totals_Org.Fields("Total_Price")
                    .Fields("Total_Tax1") = rsInvoice_Totals_Org.Fields("Total_Tax1")
                    .Fields("Total_Tax2") = rsInvoice_Totals_Org.Fields("Total_Tax2")
                    .Fields("Total_Tax3") = rsInvoice_Totals_Org.Fields("Total_Tax3")
                    .Fields("Grand_Total") = rsInvoice_Totals_Org.Fields("Grand_Total")
                    .Fields("Amt_Tendered") = rsInvoice_Totals_Org.Fields("Amt_Tendered")
                    .Fields("Amt_Change") = rsInvoice_Totals_Org.Fields("Amt_Change")
                    .Fields("InvoiceNotesUsed") = True
                    .Fields("Status") = rsInvoice_Totals_Org.Fields("Status")
                    .Fields("Cashier_ID") = rsInvoice_Totals_Org.Fields("Cashier_ID")
                    .Fields("Station_ID") = rsInvoice_Totals_Org.Fields("Station_ID")
                    .Fields("Payment_Method") = rsInvoice_Totals_Org.Fields("Payment_Method")
                    .Fields("Acct_Balance_Due") = rsInvoice_Totals_Org.Fields("Acct_Balance_Due")
                    
                    .Fields("InvType") = rsInvoice_Totals_Org.Fields("InvType")
                    .Fields("Orig_OnHoldID") = rsInvoice_Totals_Org.Fields("Orig_OnHoldID")
                    .Fields("Tax_Rate_ID") = rsInvoice_Totals_Org.Fields("Tax_Rate_ID")
                    .Fields("OrderMan") = rsInvoice_Totals_Org.Fields("OrderMan")
                    .Fields("Service_Charge") = rsInvoice_Totals_Org.Fields("Service_Charge")
                    .Fields("VATFee") = rsInvoice_Totals_Org.Fields("VATFee")
                    .Fields("Adjustment1") = rsInvoice_Totals_Org.Fields("Adjustment1")
                    .Fields("Adj1Rate") = rsInvoice_Totals_Org.Fields("Adj1Rate")
                    .Fields("Adjustment2") = rsInvoice_Totals_Org.Fields("Adjustment2")
                    .Fields("Adj2Rate") = rsInvoice_Totals_Org.Fields("Adj2Rate")
                    .Fields("Adjustment3") = rsInvoice_Totals_Org.Fields("Adjustment3")
                    .Fields("Adjustment4") = rsInvoice_Totals_Org.Fields("Adjustment4")
                    .Fields("AddMoney") = rsInvoice_Totals_Org.Fields("AddMoney")
                    .Fields("Synchronized") = True
                    .Fields("Personals") = rsInvoice_Totals_Org.Fields("Personals")
                    .Fields("Pro_Desc") = rsInvoice_Totals_Org.Fields("Pro_Desc")
                    .Fields("Reserve") = rsInvoice_Totals_Org.Fields("Reserve")
                    .Fields("OA_Amount") = rsInvoice_Totals_Org.Fields("OA_Amount")
                    .Fields("CA_Amount") = rsInvoice_Totals_Org.Fields("CA_Amount")
                    .Fields("CC_Amount") = rsInvoice_Totals_Org.Fields("CC_Amount")
                    .Fields("ROA_Amount") = rsInvoice_Totals_Org.Fields("ROA_Amount")
                    .Fields("GC_Amount") = rsInvoice_Totals_Org.Fields("GC_Amount")
                    .Fields("CT_Amount") = rsInvoice_Totals_Org.Fields("CT_Amount")
                    .Update
'                    .Requery
                End With
            .MoveNext
            i = i + 1
            Loop
        End With
       ' cnData.Execute "Update Invoice_Totals set Synchronized= Yes where Invoice_Number=" & invoice_Num_Org
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Invoice_Totals"
End Sub

'Backup Invoice Note
Public Sub gfBackup_Invoice_Notes(cnBackup As Connection, cnOrg As Connection, invoice_Num_Org As Double, ByVal Invoice_Num As String)
On Error GoTo Handle
    Dim rsInvoice_Note_Org As New ADODB.Recordset
    Dim rsInvoice_Note_Des As New ADODB.Recordset
        Set rsInvoice_Note_Org = OpenCriticalTable("Select * from Invoice_Totals_Notes where Invoice_Number=" & invoice_Num_Org, cnOrg)
        Set rsInvoice_Note_Des = Open_Table(cnBackup, "Invoice_Totals_notes")
        With rsInvoice_Note_Org
            Do While Not .EOF
                With rsInvoice_Note_Des
                        .addNew
                        .Fields("Invoice_Number") = Invoice_Num
                        .Fields("Store_ID") = rsInvoice_Note_Org.Fields("Store_ID")
                        .Fields("OpenTime") = rsInvoice_Note_Org.Fields("OpenTime")
                        .Fields("ClosingTime") = Right(rsInvoice_Note_Org.Fields("ClosingTime"), 16)
                        .Fields("Total_Minute") = Right(rsInvoice_Note_Org.Fields("Total_Minute"), 16)
                        .Fields("Karaoke_Amount") = rsInvoice_Note_Org.Fields("Karaoke_Amount")
                        .Update
'                        .Requery
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Invoice_Notes"
End Sub

'Backup Invoice Itemized

Public Sub gfBackup_Invoice_Itemized(cnBackup As Connection, cnOrg As Connection, invoice_Num_Org As Double, ByVal Invoice_Num As String)
On Error GoTo Handle
Dim i As Integer
    Dim rsInvoice_Item_Org As New ADODB.Recordset
    Dim rsInvoice_Item_Des As New ADODB.Recordset
        Set rsInvoice_Item_Org = OpenCriticalTable("Select * from Invoice_Itemized where Invoice_Number=" & invoice_Num_Org, cnOrg)
        Set rsInvoice_Item_Des = Open_Table(cnBackup, "Invoice_Itemized")
        i = 0
        With rsInvoice_Item_Org
        
            Do While Not .EOF
                With rsInvoice_Item_Des
'                .Find "Invoice_Number=" & Invoice_Num, , adSearchForward, adBookmarkFirst
'                    If Not .EOF Then Invoice_Num = GetMax_Invoice_Backup(Left(Invoice_Num, 8))
                    .addNew
                    .Fields("Invoice_Number") = Invoice_Num
                    .Fields("LineNum") = rsInvoice_Item_Org.Fields("LineNum")
                    .Fields("ItemNum") = rsInvoice_Item_Org.Fields("ItemNum")
                    .Fields("Quantity") = rsInvoice_Item_Org.Fields("Quantity")
                    .Fields("PricePer") = rsInvoice_Item_Org.Fields("PricePer")
                    .Fields("Tax1Per") = rsInvoice_Item_Org.Fields("Tax1Per")
                    .Fields("Tax2Per") = rsInvoice_Item_Org.Fields("Tax2Per")
                    .Fields("Tax3Per") = rsInvoice_Item_Org.Fields("Tax3Per")
                    .Fields("Serial_Num") = rsInvoice_Item_Org.Fields("Serial_Num")
                    .Fields("Kit_Description") = rsInvoice_Item_Org.Fields("Kit_Description")
                    .Fields("LineDisc") = rsInvoice_Item_Org.Fields("LineDisc")
                    .Fields("DiffItemName") = rsInvoice_Item_Org.Fields("DiffItemName")
                    .Fields("Store_ID") = rsInvoice_Item_Org.Fields("Store_ID")
                    .Fields("Section_ID") = rsInvoice_Item_Org.Fields("Section_ID")
                    .Fields("Person") = rsInvoice_Item_Org.Fields("Person")
                    .Fields("Returned") = rsInvoice_Item_Org.Fields("Returned")
                    .Fields("Line_Disc_Desc") = rsInvoice_Item_Org.Fields("Line_Disc_Desc")
                    .Fields("TimeOrder") = rsInvoice_Item_Org.Fields("TimeOrder")
                    .Update
'                    .Requery
                End With
            .MoveNext
            i = i + 1
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Invoice_Itemized"
End Sub

' Backup Nhun mon bi xoa
Public Sub gfBackup_Deleted_Item(cnBackup As Connection, cnOrg As Connection, invoice_Num_Org As Double, ByVal Invoice_Num As String)
On Error GoTo Handle
    Dim rsInvoice_ItemDelete_Org As New ADODB.Recordset
    Dim rsInvoice_ItemDelete_Des As New ADODB.Recordset
        Set rsInvoice_ItemDelete_Org = OpenCriticalTable("Select * from Items_Deleted where Invoice_Num=" & invoice_Num_Org, cnOrg)
        Set rsInvoice_ItemDelete_Des = Open_Table(cnBackup, "Items_Deleted")
        With rsInvoice_ItemDelete_Org
            Do While Not .EOF
                With rsInvoice_ItemDelete_Des
                    .addNew
                    .Fields("Sec_ID") = rsInvoice_ItemDelete_Org.Fields("Sec_ID")
                    .Fields("Invoice_Num") = Invoice_Num
                    .Fields("Table_ID") = rsInvoice_ItemDelete_Org.Fields("Table_ID")
                    .Fields("Cashier_ID") = rsInvoice_ItemDelete_Org.Fields("Cashier_ID")
                    .Fields("PluNo") = rsInvoice_ItemDelete_Org.Fields("PluNo")
                    .Fields("Quantity") = rsInvoice_ItemDelete_Org.Fields("Quantity")
                    .Fields("Price") = rsInvoice_ItemDelete_Org.Fields("Price")
                    .Fields("Amount") = rsInvoice_ItemDelete_Org.Fields("Amount")
                    .Fields("DateTime") = rsInvoice_ItemDelete_Org.Fields("DateTime")
                    .Fields("Ordered") = rsInvoice_ItemDelete_Org.Fields("Ordered")
                    .Fields("Reason") = rsInvoice_ItemDelete_Org.Fields("Reason")
                    .Fields("PrintCount") = rsInvoice_ItemDelete_Org.Fields("PrintCount")
                    .Fields("Line_Disc") = rsInvoice_ItemDelete_Org.Fields("Line_Disc")
                    .Fields("Line_Disc_Desc") = rsInvoice_ItemDelete_Org.Fields("Line_Disc_Desc")
                    .Update
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Deleted_Item"
End Sub

Public Function GetMax_Invoice_Backup(Path_Backup As String) As Double
On Error GoTo Handle
Dim Max_Invoice As Double
    Dim rsmax As New ADODB.Recordset
    Dim cnmax As New ADODB.Connection
    If Dir(BackupFolder & "\Database.mdb", vbDirectory) <> "" Then
        Set cnmax = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    End If
    Set rsmax = OpenCriticalTable("select Max(Invoice_Number) as maxInvoice from Invoice_Totals", cnmax)
    If rsmax.RecordCount <> 0 Then
        If Not rsmax.EOF Then
            Max_Invoice = rsmax.Fields("maxInvoice") + 1
        End If
    End If
    GetMax_Invoice_Backup = Max_Invoice
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & " GetMax_Invoice_Backup"
End Function

Public Function Get_MaxInvoice_VAT(cn As ADODB.Connection) As String
On Error GoTo Handle
    Dim MaxInvoice As String
    Dim str As String
    Dim rsMaxInvoice As New ADODB.Recordset
    str = "SELECT Max(right(Invoice_Totals.Invoice_Number,4)) AS MaxInvoice_Number" & _
            " From Invoice_Totals "
    
    Set rsMaxInvoice = OpenCriticalTable(str, cn)
    If Not rsMaxInvoice.EOF Then
        MaxInvoice = CDbl("0" & rsMaxInvoice.Fields("MaxInvoice_Number")) + 1
    Else
        MaxInvoice = 1
    End If
    Get_MaxInvoice_VAT = MaxInvoice
Exit Function
Handle:
    MsgBox Err.Description & "  " & Err.Number
End Function


Public Sub Mark_IsVAT(Invoice_Number As Double)
On Error GoTo Handle
    With rsInvoice_Totals
        .Find "Invoice_Number=" & Invoice_Number, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("IsgetVAT") = True
            .Update
        Else
            MsgBox "Ch­a ®¸nh dÊu ®­îc H§ ®· xuÊt VAT", vbInformation
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Mark_IsVAT"
End Sub
