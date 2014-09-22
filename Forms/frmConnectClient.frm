VERSION 5.00
Begin VB.Form frmConnectClient 
   Caption         =   "KÕt nèi"
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11100
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
   ScaleHeight     =   6615
   ScaleWidth      =   11100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSelect 
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   10935
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Lùa chän d÷ liÖu cÇn lÊy"
         Height          =   3375
         Left            =   0
         TabIndex        =   4
         Top             =   720
         Width           =   10935
         Begin VB.CheckBox chkAuto 
            BackColor       =   &H00C0FFFF&
            Caption         =   "LÊy tÊt c¶ d÷ liÖu b¸n hµng d­íi  m¸y con"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   840
            TabIndex        =   19
            Top             =   2880
            Width           =   6615
         End
         Begin VB.CheckBox ChkSale 
            BackColor       =   &H00FFC0C0&
            Caption         =   "D÷ liÖu b¸n hµng"
            Height          =   375
            Left            =   840
            TabIndex        =   12
            Top             =   480
            Width           =   3975
         End
         Begin VB.CheckBox chkTablePlan 
            BackColor       =   &H00FFC0C0&
            Caption         =   "S¬ ®å bµn"
            Height          =   375
            Left            =   840
            TabIndex        =   11
            Top             =   960
            Width           =   3975
         End
         Begin VB.CheckBox chkGroup 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Danh môc Nhãm hµng"
            Height          =   375
            Left            =   6120
            TabIndex        =   10
            Top             =   480
            Width           =   3975
         End
         Begin VB.CheckBox chkItems 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Danh môc hµng"
            Height          =   375
            Left            =   6120
            TabIndex        =   9
            Top             =   960
            Width           =   3975
         End
         Begin VB.CheckBox chkVendor 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Danh môc nhµ cung cÊp"
            Height          =   375
            Left            =   6120
            TabIndex        =   8
            Top             =   1440
            Width           =   3975
         End
         Begin VB.CheckBox chkDMThuchi 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Danh môc kho¶n Thu, Chi"
            Height          =   375
            Left            =   840
            TabIndex        =   7
            Top             =   1440
            Width           =   3975
         End
         Begin VB.CheckBox chkThuchu 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Thu Chi"
            Height          =   375
            Left            =   840
            TabIndex        =   6
            Top             =   1920
            Width           =   3975
         End
         Begin VB.CheckBox chkCustomer 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Danh môc kh¸ch hµng"
            Height          =   375
            Left            =   6120
            TabIndex        =   5
            Top             =   1920
            Width           =   3975
         End
         Begin VB.Label lblCheckAll 
            BackStyle       =   0  'Transparent
            Caption         =   "Chän tÊt c¶"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   2520
            TabIndex        =   16
            Top             =   2520
            Width           =   2055
         End
         Begin VB.Label lblUnCheckAll 
            BackStyle       =   0  'Transparent
            Caption         =   "Bá chän tÊt c¶"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   5160
            TabIndex        =   15
            Top             =   2520
            Width           =   2295
         End
      End
      Begin prjTouchScreen.MyProgressBar prbSync 
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   4080
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   661
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "...."
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   10695
      End
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Cancel          =   -1  'True
      Height          =   855
      Left            =   7680
      TabIndex        =   2
      Top             =   5640
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "§ãng"
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
      MICON           =   "frmConnectClient.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdGetData 
      Height          =   855
      Left            =   4560
      TabIndex        =   1
      Top             =   5640
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "LÊy d÷ liÖu"
      ENAB            =   0   'False
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
      MICON           =   "frmConnectClient.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdConnect 
      Height          =   855
      Left            =   1200
      TabIndex        =   0
      Top             =   5640
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "KiÓm tra kÕt nèi"
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
      MICON           =   "frmConnectClient.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "§­êng dÉn d÷ liÖu kÕt nèi"
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label lblPath 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   3240
      TabIndex        =   13
      Top             =   0
      Width           =   7815
   End
End
Attribute VB_Name = "frmConnectClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isAll As Boolean
Dim AmountBackup As Double

Private Sub chkAuto_Click()
    If chkAuto.Value = 1 Then
        isAll = True
    Else
        isAll = False
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdConnect_Click()
On Error GoTo Handle
    If Check_Connection(BackupFolder & "\Database.mdb", "100881administrator") Then
        cmdConnect.Caption = "KÕt nèi OK!"
        cmdConnect.ForeColor = vbBlue
        cmdConnect.FontBold = True
        fraSelect.Enabled = True
        cmdGetData.Enabled = True
    Else
        cmdConnect.Caption = "Kh«ng thÓ kÕt nèi"
        cmdConnect.ForeColor = vbWhite
        cmdConnect.BackColor = vbRed
    End If
Exit Sub
Handle:
    'Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & userName
End Sub

Private Sub cmdGetData_Click()
On Error GoTo Handle
Dim cnBackup As New ADODB.Connection
Dim cnOrg As New ADODB.Connection
Dim fso As New FileSystemObject

'
If Dir(WorkingFolder & "\Database.mdb", vbDirectory) <> "" Then
    Set cnOrg = Get_Connection()
End If

If Dir(WorkingFolder & "\Database.mdb", vbDirectory) <> "" Then
    fso.CopyFile BackupFolder & "\Database.mdb", WorkingFolder & "\Database" & Format(Now, "dd-MM-yyyy"), True
    Set cnBackup = Get_Connection()
End If
    
    'Dong bo so do ban
    If chkTablePlan.Value = 1 Then Call gfBackup_TablePlan(cnOrg, cnBackup)
    
    'Dong bo nhom hang
    If chkGroup.Value = 1 Then Call gfBackup_Group(cnOrg, cnBackup)
    
    'Dong bo Danh sach hang
    If chkItems.Value = 1 Then Call gfBackup_Items(cnOrg, cnBackup)
    
    'Dong bo du lieu ban hang
    If ChkSale.Value = 1 Then
         Call gfSynchronizeData(isAll)
        'fraSyn.Visible = False
    End If
        
    'Dong bo Khach hang
    If chkCustomer.Value = 1 Then Call gfBackup_Customer(cnOrg, cnBackup)
    
    'Dong bo Nha cung cap
    If chkVendor.Value = 1 Then Call gfBackup_Vendor(cnOrg, cnBackup)
    
    'Dong bo Danh muc thu chi
    If chkDMThuchi.Value = 1 Then Call gfBackup_DMInOut(cnOrg, cnBackup)
'
'    'Dong bo Du lieu thu chi
    If chkThuchu.Value = 1 Then Call gfBackup_InOut(cnOrg, cnBackup)
    
    prbSync.Value = prbSync.Max
    cmdGetData.Enabled = False
    
    lblTitle.Caption = "Hoµn tÊt"
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdOK_Click"
End Sub

Public Sub Update_CO_By_C(ByVal InvoiceDateTime As Double, cn As ADODB.Connection)
On Error GoTo Handle
'    Dim MinInvoice As Double
    Dim rsInvoice_Totals As ADODB.Recordset
'    MinInvoice = Get_MinInvoice(cnData, InvoiceDateTime)
Dim str As String
    str = "Select * from Invoice_totals" & _
            " where Status='CO' and Left(DateTime,8)='" & InvoiceDateTime & "'"
    Set rsInvoice_Totals = OpenCriticalTable(str, cn)
    With rsInvoice_Totals
        Do While Not .EOF
            If .Fields("Grand_total") > 0 Then
                .Fields("Status") = "C"
                .Update
            End If
        rsInvoice_Totals.MoveNext
        Loop
    End With
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Update_CO_By_C"
End Sub

Private Sub gfSynchronizeData(Auto As Boolean)
On Error GoTo Handle
    'Khai bao cac connection
    Dim cnBackup As New ADODB.Connection
    Dim cnOrg As New ADODB.Connection
    Dim rsOrg, rsOn_Hold As New ADODB.Recordset
    Dim rsinvoice_Tottal As New ADODB.Recordset
    'Khai bao bien dem thoi gian
    Dim i As Double
    Dim j, k As Integer
    Dim Invoice_Num As String
    Dim FromDate, ToDate As String
    Label1.Visible = False
    If Dir(BackupFolder & "\Database.mdb", vbDirectory) <> "" Then
        Set cnOrg = Get_Connection()
        Set rsinvoice_Tottal = Open_Table(cnOrg, "Invoice_Totals")
    Else
        Exit Sub
    End If
    
        With rsinvoice_Tottal
            If .RecordCount = 0 Then
                MsgBox "Kh«ng cßn d÷ liÖu ®Ó ®ång bé !", vbInformation
                Exit Sub
            Else
                .Sort = "DateTime  ASC"
                .MoveFirst
                .MoveNext
                FromDate = Left(!DateTime, 8)
                .MoveLast
                ToDate = Left(!DateTime, 8)
            End If
        End With
    ''' Can Modify 25/04/2011
    '''Function: Lay ngay can dong bo trong DB
    If Auto = False Then
        With frmDate_Sync
            .Let_FDate = FromDate
            .Let_TDate = ToDate
            .Show vbModal
            FromDate = .Let_FDate
            ToDate = .Let_TDate
        End With
    End If
    If FromDate = "" Or ToDate = "" Then Exit Sub
    If Dir(WorkingFolder & "\Database.mdb", vbDirectory) <> "" Then
        Set cnBackup = Get_Connection()
    Else
        Exit Sub
    End If
    If cnBackup.State = 0 Then Exit Sub
   
    'Dong bo theo tung ngay
    For i = CDbl(FromDate) To CDbl(ToDate)
        k = 1
        j = Get_MaxInvoice_Backup(cnBackup, i)
        Set rsOrg = OpenCriticalTable("select * from Invoice_Totals where left(Invoice_totals.DateTime,8)='" & i & "' and Invoice_totals.Invoice_Number<>0 and Invoice_totals.Status <>'O' and Synchronized=false", cnOrg)
        prbSync.Max = rsOrg.RecordCount
        With rsOrg
            .Sort = "Invoice_Number ASC"
            Do While Not .EOF
                If prbSync.Value < prbSync.Max Then
                    prbSync.Value = prbSync.Value + 1
                    
                Else
                    prbSync.Value = prbSync.Min
                End If
                Invoice_Num = i & Right("0000" & j, 4)
                Call gfBackup_Invoice_Notes(cnBackup, cnOrg, .Fields("Invoice_Number"), Invoice_Num)
                Call gfBackup_Invoice_Totals(cnBackup, cnOrg, .Fields("Invoice_Number"), Invoice_Num)
                Call gfBackup_Invoice_Itemized(cnBackup, cnOrg, .Fields("Invoice_Number"), Invoice_Num)
                Call gfBackup_Deleted_Item(cnBackup, cnOrg, .Fields("Invoice_Number"), Invoice_Num)
                Call gfBackup_Invoice_Per(cnBackup, cnOrg, .Fields("Invoice_Number"), Invoice_Num)
                Call gfBackup_Invoice_Kitchen_Order_Mast(cnBackup, cnOrg, .Fields("Invoice_Number"), Invoice_Num)
                Call gfBackup_Invoice_Kitchen_Order_Items(cnBackup, cnOrg, .Fields("Invoice_Number"), Invoice_Num)
                Call Delete_Invoice_AmountLarger(AmountBackup, .Fields("Invoice_Number"), cnOrg)
            .MoveNext
            lblTitle.Caption = "Ngµy:" & gfCONVERT_STRING_TO_DATE(i) & " Hãa ®¬n cÇn ®ång bé:" & .RecordCount - k
            Delay (500)
            j = j + 1
            k = k + 1
            Loop
        End With
    Call Update_CO_By_C(i, cnBackup)
    Next i
     '''' Xoa Invoice da dong bo
   
    cnOrg.Execute "Delete  from Tranfer_Joint_table"
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " gfSynchronizeData"
End Sub

Public Sub Delete_Invoice_AmountLarger(ByVal Amount As Double, Invoice As Double, cnOrg As Connection)
    On Error GoTo Handle
    Dim rsOrg As New ADODB.Recordset
    Set rsOrg = OpenCriticalTable("select * from Invoice_Totals where Grand_Total>=" & Amount & " and Invoice_Number <>0  and Invoice_Number=" & Invoice, cnOrg)
    If rsOrg.State <> 0 Then
        If rsOrg.RecordCount > 0 Then
            rsOrg.MoveFirst
        Else
            Exit Sub
        End If
    End If
    With rsOrg
        Do While Not .EOF
            cnOrg.Execute "Delete  from Items_Deleted where Invoice_Num=" & rsOrg.Fields("Invoice_Number")
            cnOrg.Execute "Delete  from Invoice_Totals_Notes where Invoice_Number=" & rsOrg.Fields("Invoice_Number")
           ' cnOrg.Execute "Delete  from Invoice_Totals where Synchronized=true and Invoice_Number<>0"
        .MoveNext
        Loop
    End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Delete_Invoice_AmountLarger"
End Sub

'''''''''''''''''''''''''''''''

' Code Dong bo du lieu
'Dong bo Invoice_Totals

Public Sub gfBackup_Invoice_Totals(cnBackup As Connection, cnOrg As Connection, invoice_Num_Org As Double, Invoice_Num As String)
On Error GoTo Handle
    Dim rsInvoice_Totals_Org As New ADODB.Recordset
    Dim rsInvoice_Totals_Des As New ADODB.Recordset
    Dim i As Integer
'    cnBackup.BeginTrans
'    cnOrg.BeginTrans
        Set rsInvoice_Totals_Org = OpenCriticalTable("Select * from Invoice_Totals where Invoice_Number=" & invoice_Num_Org, cnOrg)
        Set rsInvoice_Totals_Des = Open_Table(cnBackup, "Invoice_Totals")
        With rsInvoice_Totals_Org
            i = 0
            Do While Not .EOF
                With rsInvoice_Totals_Des
'                    .Find "Invoice_Number=" & Invoice_Num, , adSearchForward, adBookmarkFirst
'                    If Not .EOF Then Invoice_Num = GetMax_Invoice_Backup(Left(Invoice_Num, 8))
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
                    .Fields("Adj3Rate") = rsInvoice_Totals_Org.Fields("Adj3Rate")
                    .Fields("Adjustment4") = rsInvoice_Totals_Org.Fields("Adjustment4")
                    .Fields("Adj4Rate") = rsInvoice_Totals_Org.Fields("Adj4Rate")
                    
                    .Fields("Adjustment5") = rsInvoice_Totals_Org.Fields("Adjustment5")
                    .Fields("Adj5Rate") = rsInvoice_Totals_Org.Fields("Adj5Rate")
                    .Fields("Adjustment6") = rsInvoice_Totals_Org.Fields("Adjustment6")
                    .Fields("Adj6Rate") = rsInvoice_Totals_Org.Fields("Adj6Rate")
                    
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
        cnData.Execute "Update Invoice_Totals set Synchronized= Yes where Invoice_Number=" & invoice_Num_Org
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Invoice_Totals"
End Sub

'Backup Invoice Note
Public Sub gfBackup_Invoice_Notes(cnBackup As Connection, cnOrg As Connection, invoice_Num_Org As Double, Invoice_Num As String)
On Error GoTo Handle
    Dim rsInvoice_Note_Org As New ADODB.Recordset
    Dim rsInvoice_Note_Des As New ADODB.Recordset
        Set rsInvoice_Note_Org = OpenCriticalTable("Select * from Invoice_Totals_Notes where Invoice_Number=" & invoice_Num_Org, cnOrg)
        Set rsInvoice_Note_Des = Open_Table(cnBackup, "Invoice_Totals_notes")
        With rsInvoice_Note_Org
            Do While Not .EOF
                
                With rsInvoice_Note_Des
                    .Find "Invoice_Number=" & Invoice_Num, , adSearchForward, adBookmarkFirst
                    If Not .EOF Then Invoice_Num = GetMax_Invoice_Backup(Left(Invoice_Num, 8))
                        .addNew
                        .Fields("Invoice_Number") = Invoice_Num
                        .Fields("Store_ID") = rsInvoice_Note_Org.Fields("Store_ID")
                        .Fields("OpenTime") = rsInvoice_Note_Org.Fields("OpenTime")
                        .Fields("ClosingTime") = rsInvoice_Note_Org.Fields("ClosingTime")
                        .Fields("Total_Minute") = rsInvoice_Note_Org.Fields("Total_Minute")
                        .Fields("Karaoke_Amount") = rsInvoice_Note_Org.Fields("Karaoke_Amount")
                        .Update
                        .Requery
                   
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Invoice_Notes"
End Sub

'Backup Invoice Itemized

Public Sub gfBackup_Invoice_Itemized(cnBackup As Connection, cnOrg As Connection, invoice_Num_Org As Double, Invoice_Num As String)
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
Public Sub gfBackup_Deleted_Item(cnBackup As Connection, cnOrg As Connection, invoice_Num_Org As Double, Invoice_Num As String)
On Error GoTo Handle
    Dim rsInvoice_ItemDelete_Org As New ADODB.Recordset
    Dim rsInvoice_ItemDelete_Des As New ADODB.Recordset
        Set rsInvoice_ItemDelete_Org = OpenCriticalTable("Select * from Items_Deleted where Invoice_Num=" & invoice_Num_Org, cnOrg)
        Set rsInvoice_ItemDelete_Des = Open_Table(cnBackup, "Items_Deleted")
        With rsInvoice_ItemDelete_Org
            Do While Not .EOF
                With rsInvoice_ItemDelete_Des
'                    .Find "Invoice_Number=" & Invoice_Num, , adSearchForward, adBookmarkFirst
'                    If Not .EOF Then Invoice_Num = GetMax_Invoice_Backup(Left(Invoice_Num, 8))
                    
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
                    .Update
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Deleted_Item"
End Sub
'''''''''''''''''''''''''''''''
'Backup Person Mapping
Public Sub gfBackup_Invoice_Per(cnBackup As Connection, cnOrg As Connection, invoice_Num_Org As Double, Invoice_Num As String)
On Error GoTo Handle
    Dim rsInvoice_Per_Org As New ADODB.Recordset
    Dim rsInvoice_Per_Des As New ADODB.Recordset
    Dim i As Integer
    Set rsInvoice_Per_Org = OpenCriticalTable("Select * from Invoice_Totals_Person_Mapping where Invoice_Number=" & invoice_Num_Org, cnOrg)
        Set rsInvoice_Per_Des = Open_Table(cnBackup, "Invoice_Totals_Person_Mapping")
        i = 0
        With rsInvoice_Per_Org
            Do While Not .EOF
                With rsInvoice_Per_Des
'                .Find "Invoice_Number=" & Invoice_Num, , adSearchForward, adBookmarkFirst
'                    If Not .EOF Then Invoice_Num = GetMax_Invoice_Backup(Left(Invoice_Num, 8))
                    .addNew
                    .Fields("Invoice_Number") = Invoice_Num
                    .Fields("Store_ID") = Store_ID
                    .Fields("SeatNum") = rsInvoice_Per_Org.Fields("SeatNum")
                    .Update
                End With
            .MoveNext
            i = i + 1
            Loop
        End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Invoice_Notes"
End Sub
'Backup Nhung chung tu mon goi bep
Public Sub gfBackup_Invoice_Kitchen_Order_Mast(cnBackup As Connection, cnOrg As Connection, invoice_Num_Org As Double, Invoice_Num As String)
On Error GoTo Handle
    Dim rsInvoice_Kitchen_Org As New ADODB.Recordset
    Dim rsInvoice_Kitchen_Des As New ADODB.Recordset
    Dim i As Integer
        Set rsInvoice_Kitchen_Org = OpenCriticalTable("Select * from Kitchen_Order_Master where Invoice_Number=" & invoice_Num_Org, cnOrg)
        Set rsInvoice_Kitchen_Des = Open_Table(cnBackup, "Kitchen_Order_Master")
        i = 0
        With rsInvoice_Kitchen_Org
            Do While Not .EOF
                With rsInvoice_Kitchen_Des
'                .Find "Invoice_Number=" & Invoice_Num, , adSearchForward, adBookmarkFirst
'                    If Not .EOF Then Invoice_Num = GetMax_Invoice_Backup(Left(Invoice_Num, 8))
                    .addNew
                    .Fields("Invoice_Number") = Invoice_Num
                    .Fields("Station_ID") = rsInvoice_Kitchen_Org.Fields("Station_ID")
                    .Fields("Store_ID") = rsInvoice_Kitchen_Org.Fields("Store_ID")
                    .Fields("Cashier_ID") = rsInvoice_Kitchen_Org.Fields("Cashier_ID")
                    .Fields("Table_ID") = rsInvoice_Kitchen_Org.Fields("Table_ID")
                    .Update
                End With
            .MoveNext
            i = i + 1
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Invoice_Kitchen_Order_Mast"
End Sub

Public Sub gfBackup_Invoice_Kitchen_Order_Items(cnBackup As Connection, cnOrg As Connection, invoice_Num_Org As Double, Invoice_Num As String)
On Error GoTo Handle
    Dim rsInvoice_Kitchen_Items_Org As New ADODB.Recordset
    Dim rsInvoice_Kitchen_Items_Des As New ADODB.Recordset
    Dim i As Integer
        Set rsInvoice_Kitchen_Items_Org = OpenCriticalTable("Select * from Kitchen_Order_Items where Invoice_Number=" & Invoice_Num, cnOrg)
        Set rsInvoice_Kitchen_Items_Des = Open_Table(cnBackup, "Kitchen_Order_Items")
        i = 0
        With rsInvoice_Kitchen_Items_Org
            Do While Not .EOF
                With rsInvoice_Kitchen_Items_Des
'                    .Find "Invoice_Number=" & Invoice_Num, , adSearchForward, adBookmarkFirst
'                    If Not .EOF Then Invoice_Num = GetMax_Invoice_Backup(Left(Invoice_Num, 8))
                    .addNew
                    .Fields("Invoice_Number") = Invoice_Num
                    .Fields("ItemName") = rsInvoice_Kitchen_Items_Org.Fields("ItemName")
                    .Fields("ItemNum") = rsInvoice_Kitchen_Items_Org.Fields("ItemNum")
                    .Fields("Quantity") = rsInvoice_Kitchen_Items_Org.Fields("Quantity")
                    .Fields("Price") = rsInvoice_Kitchen_Items_Org.Fields("Price")
                    .Fields("LineNum") = rsInvoice_Kitchen_Items_Org.Fields("LineNum")
                    .Fields("Kit_Desc") = rsInvoice_Kitchen_Items_Org.Fields("Kit_Desc")
                    .Fields("Printer_ID") = rsInvoice_Kitchen_Items_Org.Fields("Printer_ID")
                    .Fields("Send_KP_Date") = rsInvoice_Kitchen_Items_Org.Fields("Send_KP_Date")
                    .Fields("Send_KP_Time") = rsInvoice_Kitchen_Items_Org.Fields("Send_KP_Time")
                    .Update
                End With
            .MoveNext
            i = i + 1
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Invoice_Kitchen_Order_Items"
End Sub


Public Function check_Backup() As Boolean
On Error GoTo Handle
    check_Backup = False
    If ArrayFlag(SF(3), 3) = 1 Then
        check_Backup = True
    End If
    
Exit Function
Handle:
check_Backup = False
MsgBox Err.Number & Err.Description & Me.name & "check_Backup"
End Function

Public Function GetMax_Invoice_Backup(DateMax As String) As Double
On Error GoTo Handle
Dim Max_Invoice As Double
    Dim rsmax As New ADODB.Recordset
    Dim cnmax As New ADODB.Connection
    If Dir(WorkingFolder & "\Database.mdb", vbDirectory) <> "" Then
        Set cnmax = Get_Connection()
    End If
    Set rsmax = OpenCriticalTable("select Max(Invoice_Number) as maxInvoice from Invoice_Totals where left(Invoice_Totals.DateTime,8)='" & DateMax & "'", cnmax)
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


Public Sub gfBackup_TablePlan(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsTable_Org As New ADODB.Recordset
    Dim rsTable_Des As New ADODB.Recordset
    
        lblTitle.Caption = "S¬ då bµn....."
        Delay (500)
        'Cap nhat Khu vuc
        Call gfBackup_Location(cnOrg, cnBackup)
        
        cnBackup.Execute "Delete  from Table_Diagram"
        Set rsTable_Org = Open_Table(cnOrg, "Table_Diagram")
        Set rsTable_Des = Open_Table(cnBackup, "Table_Diagram")
        
        With rsTable_Org
            Do While Not .EOF
                With rsTable_Des
                    .addNew
                    .Fields("Store_ID") = rsTable_Org.Fields("Store_ID")
                    .Fields("Section_ID") = rsTable_Org.Fields("Section_ID")
                    .Fields("Table_Number") = rsTable_Org.Fields("Table_Number")
                    .Fields("ShapeType") = rsTable_Org.Fields("ShapeType")
                    .Fields("XPos") = rsTable_Org.Fields("XPos")
                    .Fields("YPos") = rsTable_Org.Fields("YPos")
                    .Fields("Height") = rsTable_Org.Fields("Height")
                    .Fields("Width") = rsTable_Org.Fields("Width")
                    .Fields("Cost_Center_Index") = rsTable_Org.Fields("Cost_Center_Index")
                    .Fields("NumSeats") = rsTable_Org.Fields("NumSeats")
                    .Update
                    .Requery
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
'    cnBackup.RollbackTrans
'    cnOrg.RollbackTrans
    MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Invoice_Kitchen_Order_Mast"
End Sub

Public Sub gfBackup_Location(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsLocation_Org As New ADODB.Recordset
    Dim rsLocation_Des As New ADODB.Recordset
    'Xoa Du lieu khu vuc
        cnBackup.Execute "Delete  from Table_Diagram_Sections"
        
        Set rsLocation_Org = Open_Table(cnOrg, "Table_Diagram_Sections")
        Set rsLocation_Des = Open_Table(cnBackup, "Table_Diagram_Sections")
        With rsLocation_Org
            Do While Not .EOF
                With rsLocation_Des
                    .addNew
                    .Fields("Store_ID") = rsLocation_Org.Fields("Store_ID")
                    .Fields("Location_ID") = rsLocation_Org.Fields("Location_ID")
                    .Fields("Section_ID") = rsLocation_Org.Fields("Section_ID")
                    .Fields("PriceRate") = rsLocation_Org.Fields("PriceRate")
                    .Fields("VAT") = rsLocation_Org.Fields("VAT")
                    .Fields("Price_Level") = rsLocation_Org.Fields("Price_Level")
                    .Update
                    .Requery
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Location"
End Sub

Public Sub gfBackup_Group(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsGroup_Org As New ADODB.Recordset
    Dim rsGroup_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Nhãm hµng....."
        Delay (500)
        
        Set rsGroup_Org = Open_Table(cnOrg, "Departments")
        Set rsGroup_Des = Open_Table(cnBackup, "Departments")
        
        With rsGroup_Org
            Do While Not .EOF
                With rsGroup_Des
                    .Find "Dept_ID='" & rsGroup_Org.Fields("Dept_ID") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("Store_ID") = rsGroup_Org.Fields("Store_ID")
                            .Fields("Description") = rsGroup_Org.Fields("Description")
                            .Fields("MainGroup") = rsGroup_Org.Fields("MainGroup")
                            .Fields("F") = rsGroup_Org.Fields("F")
                            .Fields("ColorDept") = rsGroup_Org.Fields("ColorDept")
                            .Update
                        Else
                            .addNew
                            .Fields("Dept_ID") = rsGroup_Org.Fields("Dept_ID")
                            .Fields("Store_ID") = rsGroup_Org.Fields("Store_ID")
                            .Fields("Description") = rsGroup_Org.Fields("Description")
                            .Fields("MainGroup") = rsGroup_Org.Fields("MainGroup")
                            .Fields("F") = rsGroup_Org.Fields("F")
                            .Fields("ColorDept") = rsGroup_Org.Fields("ColorDept")
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Group"
End Sub

Public Sub gfBackup_Items(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsItems_Org As New ADODB.Recordset
    Dim rsItems_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Danh môc hµng....."
        Delay (500)
        
        Set rsItems_Org = Open_Table(cnOrg, "Inventory")
        Set rsItems_Des = Open_Table(cnBackup, "Inventory")
        With rsItems_Org
            Do While Not .EOF
                With rsItems_Des
                    .Find "ItemNum='" & rsItems_Org.Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("ItemName") = rsItems_Org.Fields("ItemName")
                            .Fields("Dept_ID") = rsItems_Org.Fields("Dept_ID")
                            .Fields("Std_Price1") = rsItems_Org.Fields("Std_Price1")
                            .Fields("Std_Price2") = rsItems_Org.Fields("Std_Price2")
                            .Fields("Std_Price3") = rsItems_Org.Fields("Std_Price3")
                            .Fields("HH_Price1") = rsItems_Org.Fields("HH_Price1")
                            .Fields("HH_Price2") = rsItems_Org.Fields("HH_Price2")
                            .Fields("HH_Price3") = rsItems_Org.Fields("HH_Price3")
                            .Fields("EV_Price1") = rsItems_Org.Fields("EV_Price1")
                            .Fields("EV_Price2") = rsItems_Org.Fields("EV_Price2")
                            .Fields("EV_Price3") = rsItems_Org.Fields("EV_Price3")
                            '.Fields("LimitPrice") = rsItems_Org.Fields("LimitPrice")
                            .Fields("Unit") = rsItems_Org.Fields("Unit")
                            .Fields("Minstock") = rsItems_Org.Fields("Minstock")
                            .Fields("Modify_Number") = rsItems_Org.Fields("Modify_Number")
                            .Fields("F1") = rsItems_Org.Fields("F1")
                            .Fields("F2") = rsItems_Org.Fields("F2")
                            .Fields("F3") = rsItems_Org.Fields("F3")
                            .Fields("F4") = rsItems_Org.Fields("F4")
                            .Fields("F5") = rsItems_Org.Fields("F5")
                            .Fields("Date_Created") = Date
                            .Fields("Picture") = rsItems_Org.Fields("Picture")
                            .Fields("Print_On_Receipt") = rsItems_Org.Fields("Print_On_Receipt")
                            .Fields("Store_ID") = Store_ID
                            .Update
                        Else
                            .addNew
                            .Fields("ItemNum") = rsItems_Org.Fields("ItemNum")
                            .Fields("ItemName") = rsItems_Org.Fields("ItemName")
                            .Fields("Dept_ID") = rsItems_Org.Fields("Dept_ID")
                            .Fields("Std_Price1") = rsItems_Org.Fields("Std_Price1")
                            .Fields("Std_Price2") = rsItems_Org.Fields("Std_Price2")
                            .Fields("Std_Price3") = rsItems_Org.Fields("Std_Price3")
                            .Fields("HH_Price1") = rsItems_Org.Fields("HH_Price1")
                            .Fields("HH_Price2") = rsItems_Org.Fields("HH_Price2")
                            .Fields("HH_Price3") = rsItems_Org.Fields("HH_Price3")
                            .Fields("EV_Price1") = rsItems_Org.Fields("EV_Price1")
                            .Fields("EV_Price2") = rsItems_Org.Fields("EV_Price2")
                            .Fields("EV_Price3") = rsItems_Org.Fields("EV_Price3")
                            .Fields("LimitPrice") = rsItems_Org.Fields("LimitPrice")
                            .Fields("Unit") = rsItems_Org.Fields("Unit")
                            .Fields("Minstock") = rsItems_Org.Fields("Minstock")
                            .Fields("Modify_Number") = rsItems_Org.Fields("Modify_Number")
                            .Fields("F1") = rsItems_Org.Fields("F1")
                            .Fields("F2") = rsItems_Org.Fields("F2")
                            .Fields("F3") = rsItems_Org.Fields("F3")
                            .Fields("F4") = rsItems_Org.Fields("F4")
                            .Fields("F5") = rsItems_Org.Fields("F5")
                            .Fields("Date_Created") = Date
                            .Fields("Picture") = rsItems_Org.Fields("Picture")
                            .Fields("Print_On_Receipt") = rsItems_Org.Fields("Print_On_Receipt")
                            .Fields("Store_ID") = Store_ID
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Items"
End Sub

Public Sub gfBackup_Customer(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsCust_Org As New ADODB.Recordset
    Dim rsCust_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Kh¸ch hµng....."
        Delay (500)
        Set rsCust_Org = Open_Table(cnOrg, "Customer")
        Set rsCust_Des = Open_Table(cnBackup, "Customer")
        
        With rsCust_Org
            Do While Not .EOF
                With rsCust_Des
                    .Find "CustNum='" & rsCust_Org.Fields("CustNum") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("CustName") = rsCust_Org.Fields("CustName")
                            .Fields("Company") = rsCust_Org.Fields("Company")
                            .Fields("Address") = rsCust_Org.Fields("Address")
                            .Fields("Phone") = rsCust_Org.Fields("Phone")
                            .Fields("Fax") = rsCust_Org.Fields("Fax")
                            .Fields("Cust_Type") = rsCust_Org.Fields("Cust_Type")
                            .Fields("TaxCode") = rsCust_Org.Fields("TaxCode")
                            .Fields("AccountNo") = rsCust_Org.Fields("AccountNo")
                            .Fields("Acct_Open_Date") = rsCust_Org.Fields("Acct_Open_Date")
                            .Fields("Acct_Close_Date") = rsCust_Org.Fields("Acct_Close_Date")
                            .Fields("Acct_Balance") = rsCust_Org.Fields("Acct_Balance")
                            .Fields("Cashier") = rsCust_Org.Fields("Cashier")
                            .Fields("Acct_Max_Balance") = rsCust_Org.Fields("Acct_Max_Balance")
                            .Fields("Birthday") = rsCust_Org.Fields("Birthday")
'                            .Fields("Last_Visit") = rsCust_Org.Fields("Last_Visit")
'                            .Fields("Tax_Rate_ID") = rsCust_Org.Fields("Tax_Rate_ID")
                            .Fields("Point") = rsCust_Org.Fields("Point")
                            .Update
                        Else
                            .addNew
                            .Fields("CustNum") = rsCust_Org.Fields("CustNum")
                            .Fields("CustName") = rsCust_Org.Fields("CustName")
                            .Fields("Company") = rsCust_Org.Fields("Company")
                            .Fields("Address") = rsCust_Org.Fields("Address")
                            .Fields("Phone") = rsCust_Org.Fields("Phone")
                            .Fields("Fax") = rsCust_Org.Fields("Fax")
                            .Fields("Cust_Type") = rsCust_Org.Fields("Cust_Type")
                            .Fields("TaxCode") = rsCust_Org.Fields("TaxCode")
                            .Fields("AccountNo") = rsCust_Org.Fields("AccountNo")
                            .Fields("Acct_Open_Date") = rsCust_Org.Fields("Acct_Open_Date")
                            .Fields("Acct_Close_Date") = rsCust_Org.Fields("Acct_Close_Date")
                            .Fields("Acct_Balance") = rsCust_Org.Fields("Acct_Balance")
                            .Fields("Cashier") = rsCust_Org.Fields("Cashier")
                            .Fields("Acct_Max_Balance") = rsCust_Org.Fields("Acct_Max_Balance")
                            .Fields("Birthday") = rsCust_Org.Fields("Birthday")
'                            .Fields("Last_Visit") = rsCust_Org.Fields("Last_Visit")
'                            .Fields("Tax_Rate_ID") = rsCust_Org.Fields("Tax_Rate_ID")
                            .Fields("Point") = rsCust_Org.Fields("Point")
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Custoer"
End Sub

Public Sub gfBackup_Vendor(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsVendor_Org As New ADODB.Recordset
    Dim rsVendor_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Kh¸ch hµng....."
        Delay (500)
        Set rsVendor_Org = Open_Table(cnOrg, "Vendors")
        Set rsVendor_Des = Open_Table(cnBackup, "Vendors")
        
        With rsVendor_Org
            Do While Not .EOF
                With rsVendor_Des
                    .Find "Vendor_Number='" & rsVendor_Org.Fields("Vendor_Number") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("Vendor_Name") = rsVendor_Org.Fields("Vendor_Name")
                            .Fields("Company") = rsVendor_Org.Fields("Company")
                            .Fields("Address_1") = rsVendor_Org.Fields("Address_1")
                            .Fields("Address_2") = rsVendor_Org.Fields("Address_2")
                            .Fields("Phone") = rsVendor_Org.Fields("Phone")
                            .Fields("Fax") = rsVendor_Org.Fields("Fax")
                            .Fields("Vendor_Tax_ID") = rsVendor_Org.Fields("Vendor_Tax_ID")
                            .Fields("Vendor_AccNo") = rsVendor_Org.Fields("Vendor_AccNo")
                            .Fields("Email") = rsVendor_Org.Fields("Email")
                            .Fields("Website") = rsVendor_Org.Fields("Website")
                            .Update
                        Else
                            .addNew
                            .Fields("Vendor_Number") = rsVendor_Org.Fields("Vendor_Number")
                            .Fields("Vendor_Name") = rsVendor_Org.Fields("Vendor_Name")
                            .Fields("Company") = rsVendor_Org.Fields("Company")
                            .Fields("Address_1") = rsVendor_Org.Fields("Address_1")
                            .Fields("Address_2") = rsVendor_Org.Fields("Address_2")
                            .Fields("Phone") = rsVendor_Org.Fields("Phone")
                            .Fields("Fax") = rsVendor_Org.Fields("Fax")
                            .Fields("Vendor_Tax_ID") = rsVendor_Org.Fields("Vendor_Tax_ID")
                            .Fields("Vendor_AccNo") = rsVendor_Org.Fields("Vendor_AccNo")
                            .Fields("Email") = rsVendor_Org.Fields("Email")
                            .Fields("Website") = rsVendor_Org.Fields("Website")
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "gfBackup_Vendor"
End Sub

Public Sub gfBackup_Thu(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsThu_Org As New ADODB.Recordset
    Dim rsThu_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Danh môc Thu....."
        Delay (500)
        Set rsThu_Org = Open_Table(cnOrg, "Receipt")
        Set rsThu_Des = Open_Table(cnBackup, "Receipt")
        
        With rsThu_Org
            Do While Not .EOF
                With rsThu_Des
                    .Find "MaThu='" & rsThu_Org.Fields("MaThu") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("DienGiai") = rsThu_Org.Fields("DienGiai")
                            .Update
                        Else
                            .addNew
                            .Fields("MaThu") = rsThu_Org.Fields("MaThu")
                            .Fields("DienGiai") = rsThu_Org.Fields("DienGiai")
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Danh muc thu"
End Sub
Public Sub gfBackup_Chi(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsChi_Org As New ADODB.Recordset
    Dim rsChi_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Danh môc Chi....."
        Delay (500)
        Set rsChi_Org = Open_Table(cnOrg, "Expense")
        Set rsChi_Des = Open_Table(cnBackup, "Expense")
        
        With rsChi_Org
            Do While Not .EOF
                With rsChi_Des
                    .Find "MaChi='" & rsChi_Org.Fields("MaChi") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("DienGiai") = rsChi_Org.Fields("DienGiai")
                            .Update
                        Else
                            .addNew
                            .Fields("MaChi") = rsChi_Org.Fields("MaChi")
                            .Fields("DienGiai") = rsChi_Org.Fields("DienGiai")
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Danh muc chi"
End Sub

Public Sub gfBackup_Thutien(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsThu_Org As New ADODB.Recordset
    Dim rsThu_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Thu tiÒn....."
        Delay (500)
        Set rsThu_Org = Open_Table(cnOrg, "Income")
        Set rsThu_Des = Open_Table(cnBackup, "Income")
        
        With rsThu_Org
            Do While Not .EOF
                With rsThu_Des
                    .Find "ID='" & rsThu_Org.Fields("ID") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("Store_ID") = rsThu_Org.Fields("Store_ID")
                            .Fields("Cashier_ID") = rsThu_Org.Fields("Cashier_ID")
                            .Fields("DateTime") = rsThu_Org.Fields("DateTime")
                            .Fields("Customer_ID") = rsThu_Org.Fields("Customer_ID")
                            .Fields("Receipt_ID") = rsThu_Org.Fields("Receipt_ID")
                            .Fields("Reciever_Name") = rsThu_Org.Fields("Reciever_Name")
                            .Fields("Division") = rsThu_Org.Fields("Division")
                            .Fields("Amount") = rsThu_Org.Fields("Amount")
                            .Fields("Description") = rsThu_Org.Fields("Description")
                            .Fields("Payment_Method") = rsThu_Org.Fields("Payment_Method")
                            .Update
                        Else
                            .addNew
                            .Fields("ID") = rsThu_Org.Fields("ID")
                            .Fields("Store_ID") = rsThu_Org.Fields("Store_ID")
                            .Fields("Cashier_ID") = rsThu_Org.Fields("Cashier_ID")
                            .Fields("DateTime") = rsThu_Org.Fields("DateTime")
                            .Fields("Customer_ID") = rsThu_Org.Fields("Customer_ID")
                            .Fields("Receipt_ID") = rsThu_Org.Fields("Receipt_ID")
                            .Fields("Reciever_Name") = rsThu_Org.Fields("Reciever_Name")
                            .Fields("Division") = rsThu_Org.Fields("Division")
                            .Fields("Amount") = rsThu_Org.Fields("Amount")
                            .Fields("Description") = rsThu_Org.Fields("Description")
                            .Fields("Payment_Method") = rsThu_Org.Fields("Payment_Method")
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Thu tien"
End Sub

Public Sub gfBackup_Chitien(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Dim rsChi_Org As New ADODB.Recordset
    Dim rsChi_Des As New ADODB.Recordset
        
        lblTitle.Caption = "Chi tiÒn....."
        Delay (500)
        Set rsChi_Org = Open_Table(cnOrg, "PayOuts")
        Set rsChi_Des = Open_Table(cnBackup, "PayOuts")
        
        With rsChi_Org
            Do While Not .EOF
                With rsChi_Des
                    .Find "ID='" & rsChi_Org.Fields("ID") & "'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            .Fields("Store_ID") = rsChi_Org.Fields("Store_ID")
                            .Fields("Cashier_ID") = rsChi_Org.Fields("Cashier_ID")
                            .Fields("DateTime") = rsChi_Org.Fields("DateTime")
                            .Fields("Vendor_Number") = rsChi_Org.Fields("Vendor_Number")
                            .Fields("Amount") = rsChi_Org.Fields("Amount")
                            .Fields("Description") = rsChi_Org.Fields("Description")
                            .Fields("Payment_Method") = rsChi_Org.Fields("Payment_Method")
                            .Fields("Expense_ID") = rsChi_Org.Fields("Expense_ID")
                            .Fields("Recieve_Name") = rsChi_Org.Fields("Recieve_Name")
                            .Fields("Division") = rsChi_Org.Fields("Division")
                            .Update
                        Else
                            .addNew
                            .Fields("ID") = rsChi_Org.Fields("ID")
                            .Fields("Store_ID") = rsChi_Org.Fields("Store_ID")
                            .Fields("Cashier_ID") = rsChi_Org.Fields("Cashier_ID")
                            .Fields("DateTime") = rsChi_Org.Fields("DateTime")
                            .Fields("Vendor_Number") = rsChi_Org.Fields("Vendor_Number")
                            .Fields("Amount") = rsChi_Org.Fields("Amount")
                            .Fields("Description") = rsChi_Org.Fields("Description")
                            .Fields("Payment_Method") = rsChi_Org.Fields("Payment_Method")
                            .Fields("Expense_ID") = rsChi_Org.Fields("Expense_ID")
                            .Fields("Recieve_Name") = rsChi_Org.Fields("Recieve_Name")
                            .Fields("Division") = rsChi_Org.Fields("Division")
                            .Update
                            .Requery
                        End If
                End With
            .MoveNext
            Loop
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Chi tien"
End Sub


Public Sub gfBackup_DMInOut(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Call gfBackup_Chi(cnOrg, cnBackup)
    Call gfBackup_Thu(cnOrg, cnBackup)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " gfBackup_DMInOut"

End Sub


Public Sub gfBackup_InOut(cnOrg As ADODB.Connection, cnBackup As ADODB.Connection)
On Error GoTo Handle
    Call gfBackup_Chitien(cnOrg, cnBackup)
    Call gfBackup_Thutien(cnOrg, cnBackup)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " gfBackup_InOut"

End Sub

Private Sub Form_Load()
On Error GoTo Handle
    lblPath.Caption = BackupFolder & "\Database.mdb"
    AmountBackup = Get_Amount_Backup
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Form_Load"
End Sub

Private Sub lblCheckAll_Click()
    chkCustomer.Value = 1
    ChkSale.Value = 1
    chkDMThuchi.Value = 1
    chkGroup.Value = 1
    chkItems.Value = 1
    chkTablePlan.Value = 1
    chkThuchu.Value = 1
    chkVendor.Value = 1
End Sub

Private Sub lblUnCheckAll_Click()
    chkCustomer.Value = False
    ChkSale.Value = False
    chkDMThuchi.Value = False
    chkGroup.Value = False
    chkItems.Value = False
    chkTablePlan.Value = False
    chkThuchu.Value = False
    chkVendor.Value = False
End Sub


Public Function Get_Amount_Backup() As Double
    On Error GoTo Handle
    Dim i As Double
    Dim rsInfor As New ADODB.Recordset
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsInfor = Open_Table(cnData, "Setup")
    With rsInfor
        If Not rsInfor.EOF Then
            i = .Fields("AmountLimited")
        End If
    End With
    Get_Amount_Backup = i
    Exit Function
Handle:
    Get_Amount_Backup = 0
    MsgBox Err.Number & Err.Description & Me.name & " Get_Amount_Backup"
End Function

