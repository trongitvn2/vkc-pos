VERSION 5.00
Begin VB.Form frmCalcuStock 
   BackColor       =   &H00800000&
   ClientHeight    =   2280
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7425
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
   ScaleHeight     =   2280
   ScaleWidth      =   7425
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   735
      Left            =   2400
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1296
      BTYPE           =   3
      TX              =   "&ß„ng"
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
      MICON           =   "frmCalcuStock.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ß∑ t›nh tÂn kho xong"
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   ".................."
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ßang t›nh tÂn kho"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "frmCalcuStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsXuatnhap As New ADODB.Recordset
Dim rsInventory As New ADODB.Recordset
Dim rsMPlu  As New ADODB.Recordset
Dim Date_Stock As String
Dim from_Date, To_Date As String
Dim calType As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Public Sub Calculate_Stock()
On Error GoTo Handle
        If Check_Table_exist("Inventory_InB" & Format(Mid(Date_Stock, 5, 2), "00") & Format(Mid(Date_Stock, 3, 2), "00")) = False Then
            Call CreateTable_InStockB(gfCONVERT_STRING_TO_DATE(Date_Stock))
        End If
        'xoa het du lieu bang tam Stock_ReportB
        cnData.Execute "Delete  from Stock_ReportB"
        'Kiem tra bang ton kho thang truoc co chua? Neu co cap nhat vao ton dau
        If Mid(Date_Stock, 5, 2) = "01" Then
            If Check_Table_exist("TonB12" & Format(CInt(Mid(Date_Stock, 3, 2)) - 1, "00")) = False Then
                'Tao bang ton cuoi nam truoc
                Call CreateTable_Ton("B12" & CInt(Mid(Date_Stock, 3, 2)) - 1)
                'Tinh toan ton kho thang 12 nam truoc
                
                'Gan du lieu ton 12 vao thang 1 nam sau
                Call Lay_ton_Cuoi_Dau("12", Mid(To_Date, 3, 2))
                'Lay du lieu nhap xuat trong thang 1
                Call Lay_Nhap_Xuat(Mid(Date_Stock, 5, 2), Mid(Date_Stock, 3, 2))
                'Lay ton cuoi thang 1
                'Call Get_Ton_B(Mid(Date_Stock, 5, 2), Mid(Date_Stock, 3, 2))
                If Check_Table_exist("TonB" & Format(Mid(Date_Stock, 5, 2), "00") & Format(Mid(Date_Stock, 3, 2), "00")) = False Then
                Call CreateTable_Ton("B" & Format(CInt(Mid(Date_Stock, 5, 2)), "00") & Mid(Date_Stock, 3, 2))
            End If
                Call Lay_Xuat_Nhap_Ton(Mid(Date_Stock, 5, 2), Mid(Date_Stock, 3, 2))
            Else
                'Gan du lieu ton 12 vao thang 1 nam sau
                Call Lay_ton_Cuoi_Dau("12", Format(Mid(To_Date, 3, 2) - 1, "00"))
                'Lay du lieu nhap xuat trong thang 1
                Call Lay_Nhap_Xuat(Mid(Date_Stock, 5, 2), Mid(Date_Stock, 3, 2))
                'Lay ton cuoi thang 1
                'Call Get_Ton_B(Mid(Date_Stock, 5, 2), Mid(Date_Stock, 3, 2))
                If Check_Table_exist("TonB" & Format(Mid(Date_Stock, 5, 2), "00") & Format(Mid(Date_Stock, 3, 2), "00")) = False Then
                Call CreateTable_Ton("B" & Format(CInt(Mid(Date_Stock, 5, 2)), "00") & Mid(Date_Stock, 3, 2))
            End If
                Call Lay_Xuat_Nhap_Ton(Mid(Date_Stock, 5, 2), Mid(Date_Stock, 3, 2))
            End If
        
        Else
            If Check_Table_exist("TonB" & Format(CInt(Mid(Date_Stock, 5, 2)) - 1, "00") & Mid(Date_Stock, 3, 2)) = False Then
                'Tao bang ton cuoi thang truoc
                Call CreateTable_Ton("B" & Format(CInt(Mid(Date_Stock, 5, 2)) - 1, "00") & Mid(Date_Stock, 3, 2))
            End If
            'Lay du lieu ton thang truoc gan qua dau thang sau
            Call Lay_ton_Cuoi_Dau(Format(CInt(Mid(Date_Stock, 5, 2)) - 1, "00"), Mid(Date_Stock, 3, 2))
            'Lay du lieu nhap xuat trong thang
            Call Lay_Nhap_Xuat(Mid(Date_Stock, 5, 2), Mid(Date_Stock, 3, 2))
            'Lay ton cuoi thang
            'Call Get_Ton_B(Mid(Date_Stock, 5, 2), Mid(Date_Stock, 3, 2))
            If Check_Table_exist("TonB" & Format(Mid(Date_Stock, 5, 2), "00") & Format(Mid(Date_Stock, 3, 2), "00")) = False Then
                Call CreateTable_Ton("B" & Format(CInt(Mid(Date_Stock, 5, 2)), "00") & Mid(Date_Stock, 3, 2))
            End If
            Call Lay_Xuat_Nhap_Ton(Mid(Date_Stock, 5, 2), Mid(Date_Stock, 3, 2))
        End If

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - Calculate_Stock"
End Sub

Public Property Let Get_Date(ByVal vNewValue As Variant)
    Date_Stock = vNewValue
End Property

Public Sub Add_Data_To_TonA_Month(ByVal Monthstock As String, ByVal YearStock As String, ByVal S As String)
On Error GoTo Handle
    Dim rsTon As New ADODB.Recordset
    Dim rsTondau As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim SQL As String
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    cnData.Execute "Delete  from Stock_Report"
    
    SQL = "SELECT Inventory_In_Master.Doc_Number, Inventory_In_Master.DateTime, Inventory_In_Master.Vendor_Number, Inventory_In_Master.Customer_ID," & _
         " Inventory_In_Master.Stock_ID, Inventory_In_Master.iReason,Inventory_In_Master.InOutType, Inventory_In" & Monthstock & YearStock & ".ItemNum," & _
         " Inventory_In" & Monthstock & YearStock & ".Description, Inventory_In" & Monthstock & YearStock & ".Quantity, Inventory_In" & Monthstock & YearStock & ".CostPer, Inventory_In" & Monthstock & YearStock & ".Amount" & _
         " FROM Inventory_In_Master INNER JOIN Inventory_In" & Monthstock & YearStock & " ON " & _
         " Inventory_In_Master.Doc_Number = Inventory_In" & Monthstock & YearStock & ".Doc_Number" & _
         " GROUP BY Inventory_In_Master.Doc_Number, Inventory_In_Master.DateTime," & _
         " Inventory_In_Master.Vendor_Number, Inventory_In_Master.Customer_ID, Inventory_In_Master.Stock_ID," & _
         " Inventory_In_Master.InOutType, Inventory_In" & Monthstock & YearStock & ".ItemNum, Inventory_In" & Monthstock & YearStock & ".Description," & _
         " Inventory_In" & Monthstock & YearStock & ".Quantity, Inventory_In" & Monthstock & YearStock & ".CostPer, Inventory_In" & Monthstock & YearStock & ".Amount, Inventory_In_Master.iReason"
    If cnData.State <> 0 Then
        Set rsTon = OpenCriticalTable(SQL, cnData)
        If Check_Table_exist("TonA" & Format(CInt(Monthstock) - 1, "00") & YearStock) = False Then
            Call CreateTable_Ton("A" & Format(CInt(Monthstock) - 1, "00") & YearStock)
            Set rsTondau = Open_Table(cnData, "TonA" & Format(CInt(Monthstock) - 1, "00") & YearStock)
        Else
            Set rsTondau = Open_Table(cnData, "TonA" & Format(CInt(Monthstock) - 1, "00") & YearStock)
        End If
        Set rsTemp = Open_Table(cnData, "Stock_Report")
    Else
        Exit Sub
    End If
    If S = "XNT" Then
        With rsTemp
            'Lay du lieu ton dau thang
            With rsTondau
                Do While Not .EOF
                    rsTemp.addNew
                    rsTemp.Fields("ItemCode") = .Fields("ItemNum")
                    rsTemp.Fields("ItemName") = .Fields("Description")
                    rsTemp.Fields("Stock_ID") = .Fields("Stock_ID")
                    If Len(.Fields("ItemNum")) > 6 Then
                        rsInventory.Find "ItemNum='" & .Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                        If Not rsInventory.EOF Then
                            rsTemp.Fields("Unit") = rsInventory.Fields("Unit")
                        End If
                    Else
                        rsMPlu.Find "PluCode='" & .Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                        If Not rsMPlu.EOF Then
                            rsTemp.Fields("Unit") = rsMPlu.Fields("Unit")
                        End If
                    End If
                    rsTemp.Fields("First_Qty") = .Fields("Quantity")
                    rsTemp.Fields("First_Amt") = Abs(.Fields("Quantity") * .Fields("CostPer"))
                    rsTemp.Update
                    rsTemp.Requery
                .MoveNext
                Loop
            End With
            'Lay du lieu xuat nhap ton trong thang
            With rsTon
                Do While Not .EOF
                    rsTemp.addNew
                    rsTemp.Fields("Supplier") = .Fields("Vendor_Number")
                    rsTemp.Fields("Customer") = .Fields("Customer_ID")
                    rsTemp.Fields("Stock_ID") = .Fields("Stock_ID")
                    rsTemp.Fields("Stock_Reason") = .Fields("iReason")
                    rsTemp.Fields("ItemCode") = .Fields("ItemNum")
                    rsTemp.Fields("ItemName") = .Fields("Description")
                    If Len(.Fields("ItemNum")) > 6 Then
                        rsInventory.Find "ItemNum='" & .Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                        If Not rsInventory.EOF Then
                            rsTemp.Fields("Unit") = rsInventory.Fields("Unit")
                        End If
                    Else
                        rsMPlu.Find "PluCode='" & .Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                        If Not rsMPlu.EOF Then
                            rsTemp.Fields("Unit") = rsMPlu.Fields("Unit")
                        End If
                    End If
    '                rsTemp.Fields("First_Qty") = 0
    '                rsTemp.Fields("First_Amt") = 0
                    If .Fields("InOutType") = "I" Then
                        rsTemp.Fields("Instock") = .Fields("Quantity")
                        rsTemp.Fields("In_Amt") = .Fields("Quantity") * .Fields("CostPer")
                    Else
                        rsTemp.Fields("OutStock") = Abs(.Fields("Quantity"))
                        rsTemp.Fields("Out_Amt") = Abs(.Fields("Quantity") * .Fields("CostPer"))
                    End If
    '                rsTemp.Fields("Last_Qty") = .Fields("")
    '                rsTemp.Fields("Last_Amt") = .Fields("")
                    rsTemp.Update
                    rsTemp.Requery
                .MoveNext
                Loop
            End With
        End With
    Else
        With rsTemp
            'Lay du lieu ton dau thang
            With rsTondau
                Do While Not .EOF
                    rsTemp.addNew
                    rsTemp.Fields("ItemCode") = .Fields("ItemNum")
                    rsTemp.Fields("ItemName") = .Fields("Description")
                    rsTemp.Fields("Stock_ID") = .Fields("Stock_ID")
                    If Len(.Fields("ItemNum")) > 6 Then
                        rsInventory.Find "ItemNum='" & .Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                        If Not rsInventory.EOF Then
                            rsTemp.Fields("Unit") = rsInventory.Fields("Unit")
                        End If
                    Else
                        rsMPlu.Find "PluCode='" & .Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                        If Not rsMPlu.EOF Then
                            rsTemp.Fields("Unit") = rsMPlu.Fields("Unit")
                        End If
                    End If
                    rsTemp.Fields("Instock") = .Fields("Quantity")
                    rsTemp.Fields("In_Amt") = Abs(.Fields("Quantity") * .Fields("CostPer"))
                    rsTemp.Update
                    rsTemp.Requery
                .MoveNext
                Loop
            End With
            'Lay du lieu xuat nhap ton trong thang
            With rsTon
                Do While Not .EOF
                    rsTemp.addNew
                    rsTemp.Fields("Supplier") = .Fields("Vendor_Number")
                    rsTemp.Fields("Customer") = .Fields("Customer_ID")
                    rsTemp.Fields("Stock_ID") = .Fields("Stock_ID")
                    rsTemp.Fields("Stock_Reason") = .Fields("iReason")
                    rsTemp.Fields("ItemCode") = .Fields("ItemNum")
                    rsTemp.Fields("ItemName") = .Fields("Description")
                    If Len(.Fields("ItemNum")) > 6 Then
                        rsInventory.Find "ItemNum='" & .Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                        If Not rsInventory.EOF Then
                            rsTemp.Fields("Unit") = rsInventory.Fields("Unit")
                        End If
                    Else
                        rsMPlu.Find "PluCode='" & .Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                        If Not rsMPlu.EOF Then
                            rsTemp.Fields("Unit") = rsMPlu.Fields("Unit")
                        End If
                    End If
                    rsTemp.Fields("Instock") = .Fields("Quantity")
                    rsTemp.Fields("In_Amt") = Abs(.Fields("Quantity") * .Fields("CostPer"))
                    rsTemp.Update
                    rsTemp.Requery
                .MoveNext
                Loop
            End With
        End With
    End If

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - Add_Data_To_TonA_Month"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Set rsInventory = Open_Table(cnData, "Inventory")
    Set rsMPlu = Open_Table(cnData, "SetMPLU")
    To_Date = Date_Stock
    Call Calculate_Stock
    MsgBox "Hoµn t t"
    Label1.Visible = False
    Label2.Visible = False
    Label3.Visible = True
    Delay (50)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - Form_Load"
End Sub

Public Sub Get_Ton(ByVal Monthstock As String, ByVal YearStock As String)
On Error GoTo Handle
    Dim strSql As String
    Dim rsTon As New ADODB.Recordset
    Dim rsTem As New ADODB.Recordset
    If Check_Table_exist("TonA" & Monthstock & YearStock) = False Then
        Call CreateTable_Ton("A" & Monthstock & YearStock)
    Else
        cnData.Execute "Delete  from TonA" & Monthstock & YearStock
    End If
    strSql = "SELECT Stock_Report.ItemCode, Stock_Report.ItemName, Stock_Report.Unit, sum(Stock_Report.Instock) as Qty, sum(Stock_Report.In_Amt) as amt, Stock_Report.Stock_ID" & _
            " From Stock_Report" & _
            " Group by Stock_Report.ItemCode, Stock_Report.ItemName, Stock_Report.Unit, Stock_Report.Stock_ID"
    Set rsTon = Open_Table(cnData, "TonA" & Monthstock & YearStock)
    Set rsTem = OpenCriticalTable(strSql, cnData)
    Do While Not rsTem.EOF
        With rsTon
            .addNew
            .Fields("ItemNum") = rsTem.Fields("ItemCode")
            .Fields("Description") = rsTem.Fields("ItemName")
            .Fields("Unit") = rsTem.Fields("Unit")
            .Fields("Stock_ID") = rsTem.Fields("Stock_ID")
            .Fields("Quantity") = rsTem.Fields("Qty")
            If rsTem.Fields("Qty") <> 0 Then
                .Fields("CostPer") = Abs(rsTem.Fields("amt") / rsTem.Fields("Qty"))
            Else
                .Fields("CostPer") = 0
            End If
            .Fields("Amount") = Abs(rsTem.Fields("amt"))
            .Update
            .Requery
        End With
    rsTem.MoveNext
    Loop
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Public Sub Lay_Nhap_Xuat(ByVal Monthstock As String, ByVal YearStock As String)
On Error GoTo Handle
    Dim rsNX As New ADODB.Recordset
    Dim rsTondau As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim SQL As String
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\database.mdb", "100881administrator")
    'cnData.Execute "Delete  from Stock_ReportB"
    
'    cnData.Execute "Delete  from TonB" & MonthStock & yearStock
    cnData.Execute "Delete  from Inventory_InB" & Monthstock & YearStock & " where left(Doc_Number,9)='" & "XKB20" & YearStock & Monthstock & "'"
'    SQL = "SELECT Instock_MasterB.Doc_Number, Instock_MasterB.DateTime, Instock_MasterB.Vendor_Number, Instock_MasterB.Customer_ID," & _
'         " Instock_MasterB.Stock_ID, Instock_MasterB.iReason,Instock_MasterB.InOutType, Inventory_InB" & Monthstock & YearStock & ".ItemNum," & _
'         " Inventory_InB" & Monthstock & YearStock & ".Description, sum(Inventory_InB" & Monthstock & YearStock & ".Quantity)  as Quantity, avg(Inventory_InB" & Monthstock & YearStock & ".CostPer) as CostPer, sum(Inventory_InB" & Monthstock & YearStock & ".Amount) as Amount" & _
'         " FROM Instock_MasterB INNER JOIN Inventory_InB" & Monthstock & YearStock & " ON " & _
'         " Instock_MasterB.Doc_Number = Inventory_InB" & Monthstock & YearStock & ".Doc_Number" & _
'         " GROUP BY Instock_MasterB.Doc_Number, Instock_MasterB.DateTime," & _
'         " Instock_MasterB.Vendor_Number, Instock_MasterB.Customer_ID, Instock_MasterB.Stock_ID," & _
'         " Instock_MasterB.InOutType, Inventory_InB" & Monthstock & YearStock & ".ItemNum, Inventory_InB" & Monthstock & YearStock & ".Description," & _
'         " Inventory_InB" & Monthstock & YearStock & ".Quantity, Inventory_InB" & Monthstock & YearStock & ".CostPer, Inventory_InB" & Monthstock & YearStock & ".Amount, Instock_MasterB.iReason"
   
    SQL = "SELECT  Instock_MasterB.Vendor_Number, Instock_MasterB.Customer_ID," & _
         " Instock_MasterB.Stock_ID, Instock_MasterB.iReason,Instock_MasterB.InOutType, Inventory_InB" & Monthstock & YearStock & ".ItemNum," & _
         " Inventory_InB" & Monthstock & YearStock & ".Description, sum(Inventory_InB" & Monthstock & YearStock & ".Quantity)  as Quantity, avg(Inventory_InB" & Monthstock & YearStock & ".CostPer) as CostPer, sum(Inventory_InB" & Monthstock & YearStock & ".Amount) as Amount" & _
         " FROM Instock_MasterB INNER JOIN Inventory_InB" & Monthstock & YearStock & " ON " & _
         " Instock_MasterB.Doc_Number = Inventory_InB" & Monthstock & YearStock & ".Doc_Number" & _
         " GROUP BY Instock_MasterB.Vendor_Number, Instock_MasterB.Customer_ID, Instock_MasterB.Stock_ID," & _
         " Instock_MasterB.InOutType, Inventory_InB" & Monthstock & YearStock & ".ItemNum, Inventory_InB" & Monthstock & YearStock & ".Description," & _
         " Inventory_InB" & Monthstock & YearStock & ".Quantity, Inventory_InB" & Monthstock & YearStock & ".CostPer, Inventory_InB" & Monthstock & YearStock & ".Amount, Instock_MasterB.iReason"
        
    'Lay du lieu xuat nhap ton trong thang
            'Lay du lieu xuat ban hang trong thang
                'Lay du lieu xuat khong che bien
                Dim i As String
                i = from_Date
            Do While i <= To_Date
                Call Out_Stock_KCB(i)
                'lay du lieu xuat che bien
                Call Out_Stock_CB(i)
                i = i + 1
            Loop
    'Lay du lieu xuat nhap trong thang
            Set rsNX = OpenCriticalTable(SQL, cnData)
            Set rsTemp = Open_Table(cnData, "Stock_ReportB")
                With rsNX
                    Do While Not .EOF
                        rsTemp.addNew
'                        rsTemp.Fields("DocNumber") = .Fields("Doc_Number")
'                        rsTemp.Fields("DateTime") = .Fields("DateTime")
                        
                        rsTemp.Fields("Supplier") = .Fields("Vendor_Number")
                        rsTemp.Fields("Customer") = .Fields("Customer_ID")
                        rsTemp.Fields("Stock_ID") = .Fields("Stock_ID")
                        rsTemp.Fields("Stock_Reason") = .Fields("iReason")
                        rsTemp.Fields("ItemCode") = .Fields("ItemNum")
                        rsTemp.Fields("ItemName") = .Fields("Description")
                        If Len(.Fields("ItemNum")) > 6 Then
                            rsInventory.Find "ItemNum='" & .Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                            If Not rsInventory.EOF Then
                                rsTemp.Fields("Unit") = rsInventory.Fields("Unit")
                            End If
                        Else
                            rsMPlu.Find "PluCode='" & .Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                            If Not rsMPlu.EOF Then
                                rsTemp.Fields("Unit") = rsMPlu.Fields("Unit")
                            End If
                        End If
        '                rsTemp.Fields("First_Qty") = 0
        '                rsTemp.Fields("First_Amt") = 0
                        If .Fields("InOutType") = "I" Then
                            rsTemp.Fields("Instock") = .Fields("Quantity")
                            rsTemp.Fields("In_Amt") = .Fields("Quantity") * .Fields("CostPer")
                        ElseIf .Fields("InOutType") = "T" Then
                            rsTemp.Fields("Instock") = .Fields("Quantity")
                            rsTemp.Fields("In_Amt") = .Fields("Quantity") * .Fields("CostPer")
                        Else
                            rsTemp.Fields("OutStock") = Abs(.Fields("Quantity"))
                            rsTemp.Fields("Out_Amt") = Abs(.Fields("Quantity") * .Fields("CostPer"))
                        End If
        '                rsTemp.Fields("Last_Qty") = .Fields("")
        '                rsTemp.Fields("Last_Amt") = .Fields("")
                        rsTemp.Update
                        rsTemp.Requery
                    .MoveNext
                    Loop
                End With
            Set rsNX = Nothing
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - Lay_Nhap_Xuat"
End Sub

Public Property Let Let_From_Date(ByVal vNewValue As Variant)
    from_Date = vNewValue
End Property

Public Sub Out_Stock_KCB(ByVal Ngay As String)
    On Error GoTo Handle
        Dim str As String
        Dim rsMasterB As New ADODB.Recordset
        Dim rsInStockB As New ADODB.Recordset
        Dim rsSale As New ADODB.Recordset
        Set rsInventory = Open_Table(cnData, "Inventory")
        Set rsMasterB = Open_Table(cnData, "Instock_MasterB")
        Set rsInStockB = Open_Table(cnData, "Inventory_InB" & Mid(Ngay, 5, 2) & Mid(Ngay, 3, 2))
        str = "select Invoice_Totals.Invoice_Number,Invoice_Totals.DateTime,Invoice_Itemized.ItemNum,Invoice_Itemized.DiffItemName,sum(Invoice_Itemized.Quantity) as Quantity,avg(Invoice_Itemized.PricePer) as Cost FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number" & _
            " where Invoice_Itemized.ItemNum not in (select SetMLink.PLUCode from SetMLink) and left(Invoice_Totals.DateTime,8)='" & Ngay & "'" & _
            " group by Invoice_Totals.Invoice_Number,Invoice_Totals.DateTime,Invoice_Itemized.ItemNum,Invoice_Itemized.DiffItemName"
            
        Set rsSale = OpenCriticalTable(str, cnData)
        With rsSale
            Do While Not .EOF
                With rsMasterB
                    .Find "Doc_Number='XKB" & Ngay & "KCB'", , adSearchForward, adBookmarkFirst
                    If .EOF Then
                        .addNew
                        .Fields("Doc_Number") = "XKB" & Ngay & "KCB"
                        .Fields("Store_ID") = Store_ID
                        .Fields("DateTime") = Ngay
                        .Fields("Customer_ID") = "101"
                        .Fields("iLocked") = True
                        .Fields("iReason") = "XK"
                        .Fields("Stock_ID") = "01"
                        .Fields("InOutType") = "O"
                        .Update
                    End If
                End With
                
                rsInventory.Find "ItemNum='" & rsSale.Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    If rsSale.Fields("ItemNum") <> "KAR" Then
                        If ArrayFlag(rsInventory.Fields("F3"), 8) = 1 Then
                            With rsInStockB
                                .addNew
                                .Fields("Doc_Number") = "XKB" & Ngay & "KCB"
                                .Fields("DateTime") = Ngay
                                .Fields("ItemNum") = rsSale.Fields("ItemNum")
                                .Fields("Description") = rsSale.Fields("DiffItemName")
                                .Fields("Store_ID") = Store_ID
                                .Fields("Quantity") = -rsSale.Fields("Quantity")
                                .Fields("CostPer") = rsSale.Fields("Cost")
                                .Fields("Amount") = rsSale.Fields("Quantity") * rsSale.Fields("Cost")
                                .Update
                            
                            End With
                        End If
                    End If
                End If
                .MoveNext
            Loop
        End With
            
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - Out_Stock_KCB"
End Sub

Public Sub Out_Stock_CB(ByVal Ngay As String)
On Error GoTo Handle
    Dim str, strSetMLink As String
    Dim rsMasterB As New ADODB.Recordset
    Dim rsInStockB As New ADODB.Recordset
    Dim rsSale As New ADODB.Recordset
    Dim rsSetMPLU As New ADODB.Recordset
    
    Set rsMasterB = Open_Table(cnData, "Instock_MasterB")
    Set rsInStockB = Open_Table(cnData, "Inventory_InB" & Mid(Ngay, 5, 2) & Mid(Ngay, 3, 2))
    str = "select Invoice_Totals.Invoice_Number,Invoice_Totals.DateTime,Invoice_Itemized.ItemNum,Invoice_Itemized.DiffItemName,sum(Invoice_Itemized.Quantity) as Quantity,avg(Invoice_Itemized.PricePer) as Cost FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number" & _
        " where Invoice_Itemized.ItemNum in (select SetMLink.PLUCode from SetMLink) and left(Invoice_Totals.DateTime,8)='" & Ngay & "'" & _
        " group by Invoice_Totals.Invoice_Number,Invoice_Totals.DateTime,Invoice_Itemized.ItemNum,Invoice_Itemized.DiffItemName"
     
    strSetMLink = "SELECT SetMLink.PLUCode, SetMPLU.PLUName, SetMLink.SMPLUCode, SetMLink.StockRate," & _
                 " SetMPLU.Cost, SetMPLU.Unit" & _
                " FROM SetMLink INNER JOIN SetMPLU ON SetMLink.SMPLUCode = SetMPLU.PLUCode" & _
                "where SetMLink.PLUCode='"

    Set rsSale = OpenCriticalTable(str, cnData)
    
    With rsSale
        Do While Not .EOF
            With rsMasterB
                .Find "Doc_Number='XKB" & Ngay & "CB'", , adSearchForward, adBookmarkFirst
                If .EOF Then
                    .addNew
                    .Fields("Doc_Number") = "XKB" & Ngay & "CB"
                    .Fields("Store_ID") = Store_ID
                    .Fields("DateTime") = Ngay
                    .Fields("Customer_ID") = "101"
                    .Fields("iLocked") = True
                    .Fields("iReason") = "XK"
                    .Fields("Stock_ID") = "02"
                    .Fields("InOutType") = "O"
                    .Update
                End If
            End With
            
            With rsInStockB
            strSetMLink = "SELECT SetMLink.PLUCode, SetMPLU.PLUName, SetMLink.SMPLUCode, SetMLink.StockRate," & _
                 " SetMPLU.Cost, SetMPLU.Unit" & _
                " FROM SetMLink INNER JOIN SetMPLU ON SetMLink.SMPLUCode = SetMPLU.PLUCode" & _
                " where SetMLink.PLUCode='" & rsSale.Fields("ItemNum") & "'"
                
                Set rsSetMPLU = OpenCriticalTable(strSetMLink, cnData)
                If rsSetMPLU.RecordCount > 0 Then rsSetMPLU.MoveFirst
                    Do While Not rsSetMPLU.EOF
                        .addNew
                        .Fields("Doc_Number") = "XKB" & Ngay & "CB"
                        .Fields("DateTime") = Ngay
                        rsInStockB.Fields("ItemNum") = rsSetMPLU.Fields("SMPLUCode")
                        rsInStockB.Fields("Description") = rsSetMPLU.Fields("PLUName")
                        rsInStockB.Fields("Store_ID") = Store_ID
                        rsInStockB.Fields("Quantity") = -rsSale.Fields("Quantity") * rsSetMPLU.Fields("StockRate") / 1000
                        rsInStockB.Fields("CostPer") = rsSetMPLU.Fields("Cost")
                        rsInStockB.Fields("Amount") = rsInStockB.Fields("Quantity") * rsInStockB.Fields("CostPer")
                        .Update
                   rsSetMPLU.MoveNext
                   Loop
            End With
            .MoveNext
        Loop
    End With
Exit Sub
Handle:
    MsgBox Err.Description & Me.name & " Xuat Kho che bien"

End Sub

Public Sub Get_Ton_B(ByVal Monthstock As String, ByVal YearStock As String)
On Error GoTo Handle
    Dim strSql As String
    Dim rsTon As New ADODB.Recordset
    Dim rsTem As New ADODB.Recordset
    If Check_Table_exist("TonB" & Monthstock & YearStock) = False Then
        Call CreateTable_Ton("B" & Monthstock & YearStock)
    Else
        cnData.Execute "Delete  from TonB" & Monthstock & YearStock
    End If
    strSql = "SELECT Stock_ReportB.ItemCode, Stock_ReportB.ItemName, Stock_ReportB.Unit, sum(Stock_ReportB.Instock) as Qty, sum(Stock_ReportB.In_Amt) as amt, Stock_ReportB.Stock_ID" & _
            " From Stock_ReportB" & _
            " Group by Stock_ReportB.ItemCode, Stock_ReportB.ItemName, Stock_ReportB.Unit, Stock_ReportB.Stock_ID"
    Set rsTon = Open_Table(cnData, "TonB" & Monthstock & YearStock)
    Set rsTem = OpenCriticalTable(strSql, cnData)
    Do While Not rsTem.EOF
        With rsTon
            .addNew
            .Fields("ItemNum") = rsTem.Fields("ItemCode")
            .Fields("Description") = rsTem.Fields("ItemName")
            .Fields("Unit") = rsTem.Fields("Unit")
            .Fields("Stock_ID") = rsTem.Fields("Stock_ID")
            .Fields("Quantity") = rsTem.Fields("Qty")
            If rsTem.Fields("Qty") <> 0 Then
                .Fields("CostPer") = rsTem.Fields("amt") / rsTem.Fields("Qty")
            Else
                .Fields("CostPer") = 0
            End If
            .Fields("Amount") = rsTem.Fields("amt")
            .Update
            .Requery
        End With
    rsTem.MoveNext
    Loop
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub


Public Sub Lay_ton_Cuoi_Dau(ByVal Monthstock As String, ByVal YearStock As String)
On Error GoTo Handle
Dim rsTemp As New ADODB.Recordset
Dim rsTondau As New ADODB.Recordset
     If cnData.State <> 0 Then
        
        If Check_Table_exist("TonB" & Format(CInt(Monthstock), "00") & YearStock) = False Then
            Call CreateTable_Ton("B" & Format(CInt(Monthstock), "00") & YearStock)
            Set rsTondau = Open_Table(cnData, "TonB" & Format(CInt(Monthstock) - 1, "00") & YearStock)
        Else
            Set rsTondau = Open_Table(cnData, "TonB" & Format(Monthstock, "00") & YearStock)
         
        End If
        Set rsTemp = Open_Table(cnData, "Stock_ReportB")
    Else
        Exit Sub
    End If
    
    With rsTemp
            'Lay du lieu ton dau thang
            With rsTondau
                Do While Not .EOF
                    rsTemp.addNew
'                    rsTemp.Fields("DocNumber") = "TD" & Monthstock & YearStock
'                    rsTemp.Fields("DateTime") = Format(Year(Date), "0000") & Format(Month(Date), "00") & "01"
                    rsTemp.Fields("ItemCode") = .Fields("ItemNum")
                    rsTemp.Fields("ItemName") = .Fields("Description")
                    rsTemp.Fields("Stock_ID") = .Fields("Stock_ID")
                    If Len(.Fields("ItemNum")) > 6 Then
                        rsInventory.Find "ItemNum='" & .Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                        If Not rsInventory.EOF Then
                            rsTemp.Fields("Unit") = rsInventory.Fields("Unit")
                        End If
                    Else
                        rsMPlu.Find "PluCode='" & .Fields("ItemNum") & "'", , adSearchForward, adBookmarkFirst
                        If Not rsMPlu.EOF Then
                            rsTemp.Fields("Unit") = rsMPlu.Fields("Unit")
                        End If
                    End If
                    rsTemp.Fields("First_Qty") = .Fields("Quantity")
                    rsTemp.Fields("First_Amt") = .Fields("Quantity") * .Fields("CostPer")
                    rsTemp.Update
                    rsTemp.Requery
                .MoveNext
                Loop
            End With
        End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Lay_ton_Cuoi_Dau"
End Sub

Public Sub Lay_Xuat_Nhap_Ton(ByVal Monthstock As String, ByVal YearStock As String)
On Error GoTo Handle
    Dim rsTam As New ADODB.Recordset
    Dim rsTon As New ADODB.Recordset
    Dim rsStock_ReportB As New ADODB.Recordset
    Dim SQL As String
    cnData.Execute "Delete  from TonB" & Monthstock & YearStock
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    SQL = "SELECT Stock_ReportB.Stock_ID, Stock_ReportB.ItemCode, Stock_ReportB.ItemName, Stock_ReportB.Unit, Sum(Stock_ReportB.First_Qty) AS FirstQty, Sum(Stock_ReportB.First_Amt) AS FirstAmt, Sum(Stock_ReportB.Instock) AS InQty, Sum(Stock_ReportB.In_Amt) AS InAmt, Sum(Stock_ReportB.OutStock) AS OutQty, Sum(Stock_ReportB.First_Qty)+Sum(Stock_ReportB.Instock)-Sum(Stock_ReportB.OutStock) AS LastQty" & _
          " From Stock_ReportB" & _
          " GROUP BY Stock_ReportB.Stock_ID, Stock_ReportB.ItemCode, Stock_ReportB.ItemName, Stock_ReportB.Unit,Stock_ReportB.DocNumber,Stock_ReportB.DateTime"
    'Lay du lieu tong Stock_reportB len recordset rsTam
    Set rsTam = OpenCriticalTable(SQL, cnData)
    'Xoa du lieu trong bang Stock_ReportB
    cnData.Execute "Delete  from Stock_ReportB"
    'Gan du lieu tro xuong Stock_ReportB
    Set rsStock_ReportB = Open_Table(cnData, "Stock_ReportB")
    Set rsTon = Open_Table(cnData, "TonB" & Monthstock & YearStock)
    With rsTam
        'If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            With rsStock_ReportB
                .addNew
'                .Fields("DocNumber") = rsTam.Fields("DocNumber")
'                .Fields("DateTime") = rsTam.Fields("DateTime")
                .Fields("Stock_ID") = rsTam.Fields("Stock_ID")
                .Fields("ItemCode") = rsTam.Fields("ItemCode")
                .Fields("ItemName") = rsTam.Fields("ItemName")
                .Fields("Unit") = rsTam.Fields("Unit")
                .Fields("First_Qty") = rsTam.Fields("FirstQty")
                .Fields("First_Amt") = rsTam.Fields("FirstAmt")
                .Fields("InStock") = rsTam.Fields("InQty")
                .Fields("In_Amt") = rsTam.Fields("InAmt")
                .Fields("OutStock") = rsTam.Fields("OutQty")
                If rsTam.Fields("InQty") > 0 Then
                    .Fields("Out_Amt") = rsTam.Fields("OutQty") * rsTam.Fields("InAmt") / rsTam.Fields("InQty")
                Else
                    .Fields("Out_Amt") = 0
                End If
                .Fields("Last_Qty") = rsTam.Fields("LastQty")
                .Update
'                .Requery
            End With
            With rsTon
                .addNew
                .Fields("Stock_ID") = rsTam.Fields("Stock_ID")
                .Fields("ItemNum") = rsTam.Fields("ItemCode")
                .Fields("Description") = rsTam.Fields("ItemName")
                .Fields("Unit") = rsTam.Fields("Unit")
                .Fields("Quantity") = rsTam.Fields("LastQty")
                If rsTam.Fields("InQty") > 0 Then
                    .Fields("CostPer") = rsTam.Fields("InAmt") / rsTam.Fields("InQty")
                    .Fields("Amount") = rsTam.Fields("LastQty") * rsTam.Fields("InAmt") / rsTam.Fields("InQty")
                Else
                    .Fields("CostPer") = 0
                    .Fields("Amount") = 0
                End If
                
                .Update
'                .Requery
            End With
        .MoveNext
        Loop
    
    End With
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Lay_Xuat_Nhap_Ton"
End Sub
