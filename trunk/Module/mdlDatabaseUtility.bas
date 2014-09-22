Attribute VB_Name = "mdlDatabaseUtility"
Option Explicit
Public myProvider As String
'Public Const myProvider As String = "Provider=SQLNCLI10;Server=Khaccan\SQLEXPRESS;Database=VKC-POS;Uid=sa; Pwd=131112;"
'myProvider = "Provider=SQLNCLI10;Driver={SQL Server Native Client 10.0};Integrated Security=SSPI;Port=1433;Server=" & ServerName & ";Database=" & DataBaseName & ";Uid=SA; Pwd=" & DB_Password & ";"
Public Function Get_Connection(SVName As String, DBName As String, USID As String, USPass As String) As ADODB.Connection
On Error GoTo errHdl
    Dim cnTemp As New ADODB.Connection
    With cnTemp
        myProvider = "Provider=SQLOLEDB;Server=" & SVName & ";Database=" & DBName & ";Uid=" & USID & "; Pwd=" & USPass & ";Trusted_Connection=No;"
        .CursorLocation = adUseClient
        .ConnectionString = myProvider
        .Open
    End With
    Set Get_Connection = cnTemp
    Exit Function

errHdl:
    MsgBox Err.Number & Err.Description & " -------Kh«ng kÕt nèi ®Õn d÷ liÖu !"
    'frmConnect_Data.Show vbModal
End Function

Public Function Check_Connection(ByVal DatabaseFile As String, ByVal Password As String) As Boolean
On Error GoTo errHdl

    Dim cnTemp As New ADODB.Connection
   
    With cnTemp
        myProvider = "Provider=SQLOLEDB;Server=" & ServerName & ";Database=" & DataBaseName & ";Uid=" & UserLog & "; Pwd=" & DB_Password & ";Trusted_Connection =No;"
        .CursorLocation = adUseClient
        .ConnectionString = myProvider
        .Open
    End With
    If cnTemp.State = 1 Then Check_Connection = True
    Exit Function

errHdl:
    Check_Connection = False
    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & "Check_connection" & vbTab & userName
    MsgBox "Kh«ng thÓ kÕt nèi víi m¸y con, vui lßng kiÓm tra l¹i ®­êng dÉn hoÆc ®­êng truyÒn Internet! C¶m ¬n!", vbInformation
End Function

Public Function Open_Table(cn As ADODB.Connection, ByVal TableName As String) As ADODB.Recordset
''    Dim catTable As New ADOX.Catalog
    Dim rsTemp As New ADODB.Recordset
    Dim iInc As Integer
    Dim SQLStr As String
'    Dim flagFound As Boolean
    
    On Error GoTo ErrHandle
        If Check_Table_exist(TableName) Then
        '    flagFound = False
            SQLStr = "Select * From [" & TableName & "]"
            If Not (cn Is Nothing) Then
                If cn.State = adStateOpen Then
                    rsTemp.Open SQLStr, cn, adOpenKeyset, adLockPessimistic
                Else
                    Set cn = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
                    rsTemp.Open SQLStr, cn, adOpenKeyset, adLockPessimistic
                End If
        '        rstemp.Requery
            End If
            Set Open_Table = rsTemp
        Else
            MsgBox "Kh«ng t×m thÊy d÷ liÖu yªu cÇu"
        End If
    Exit Function
ErrHandle:

    MsgBox "Bao loi" & "Open_Table"
End Function

Public Sub CloseRecordset(ByRef rs As ADODB.Recordset)
    On Error GoTo Err_CloseRecordset
    If rs.State = adStateOpen Then Set rs = Nothing
    Exit Sub
    
Err_CloseRecordset:
    MsgBox "Bao loi" & "CloseRecordset"
End Sub

Public Function OpenTable(ByVal TableName As String, ByVal strOrder As String, ByVal cn As ADODB.Connection) As ADODB.Recordset
On Error GoTo errHdl
    'If Check_Table_exist(TableName) Then
        Dim rsTemp As New ADODB.Recordset
        Dim SQLStr As String
        
        'If cn.State = 0 Then Set cn = Get_Connection
        If strOrder = "" Then
            SQLStr = "Select * From [" & TableName & "]"
        Else
            SQLStr = "Select * From [" & TableName & "] order by " & strOrder
        End If
        rsTemp.Open SQLStr, cn, adOpenDynamic, adLockOptimistic
        
        Set OpenTable = rsTemp
'    Else

    Exit Function
errHdl:
    MsgBox Err.Description & " : " & Err.Description & vbCrLf _
    & "mdlDatabaseUtility - Opentable"
End Function

Public Function OpenCriticalTable(ByVal vCritical As String, ByVal cn As ADODB.Connection) As ADODB.Recordset
On Error GoTo errHdl
    Dim rsTemp As New ADODB.Recordset
'    If cn.State = 0 Then Set cn = Get_Connection
    rsTemp.Open vCritical, cn, adOpenKeyset, adLockPessimistic
    
    Set OpenCriticalTable = rsTemp
    'Set cnData = Nothing

    Exit Function
errHdl:
        MsgBox Err.Number & Err.Description & " : " & Err.Description & vbCrLf _
'        & "mdlDatabaseUtility - OpenCriticalTable"
    
End Function
Public Function CreateTable_InStock(DateIn As String) As ADOX.Table
On Error GoTo errHdl
    Dim tblInStock As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
     
    With tblInStock
        .name = "Inventory_In" & Format(Month(DateIn), "00") & Right(Format(Year(DateIn), "00"), 2)
        .Columns.Append "Doc_Number", adVarWChar, 20
        .Columns.Append "DateTime", adVarWChar, 20
        .Columns.Append "ItemNum", adVarWChar, 20
        .Columns.Append "Description", adVarWChar, 50
        .Columns("Description").Attributes = adColNullable
        .Columns.Append "Store_ID", adVarWChar, 10
        .Columns.Append "Quantity", adDouble
        .Columns.Append "CostPer", adDouble
        .Columns.Append "Amount", adDouble
    End With
    Set CreateTable_InStock = tblInStock
    cat.Tables.Append tblInStock
    
    cat.Tables.Refresh
    Delay (500)
    
    Set cat = Nothing
    Set tblInStock = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " CreateTable_InStock"
End Function
Public Function CreateTable_InStockB(DateIn As String) As ADOX.Table
On Error GoTo errHdl
    Dim tblInStock As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblInStock
        .name = "Inventory_InB" & Format(Month(DateIn), "00") & Right(Format(Year(DateIn), "00"), 2)
        .Columns.Append "Doc_Number", adVarWChar, 20
        .Columns.Append "DateTime", adVarWChar, 20
        .Columns.Append "ItemNum", adVarWChar, 20
        .Columns.Append "Description", adVarWChar, 50
        .Columns("Description").Attributes = adColNullable
        .Columns.Append "Store_ID", adVarWChar, 10
        .Columns.Append "Quantity", adDouble
        .Columns.Append "CostPer", adDouble
        .Columns.Append "Amount", adDouble
    End With
    Set CreateTable_InStockB = tblInStock
    cat.Tables.Append tblInStock
    
    cat.Tables.Refresh
    Delay (500)
    
    Set cat = Nothing
    Set tblInStock = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " CreateTable_InStock"
End Function
Public Function CreateTable_In() As ADOX.Table
On Error GoTo errHdl
    Dim tblInStock As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
     
    With tblInStock
        .name = "Inventory_In"
        .Columns.Append "Doc_Number", adVarWChar, 20
        .Columns.Append "DateTime", adVarWChar, 20
        .Columns.Append "Stock_ID", adVarWChar, 2
        .Columns.Append "GroupA", adVarWChar, 3
        .Columns("GroupA").Attributes = adColNullable
        .Columns.Append "ItemNum", adVarWChar, 20
        .Columns.Append "Description", adVarWChar, 50
        .Columns("Description").Attributes = adColNullable
        .Columns.Append "Unit", adVarWChar, 50
        .Columns("Unit").Attributes = adColNullable
        .Columns.Append "Quantity", adDouble
        .Columns.Append "CostPer", adDouble
    End With
    Set CreateTable_In = tblInStock
    cat.Tables.Append tblInStock
    
    cat.Tables.Refresh
    Delay (500)
    
    Set cat = Nothing
    Set tblInStock = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " CreateTable_InStock"
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Tao bang xuat kho
Public Function CreateTable_Out() As ADOX.Table
On Error GoTo errHdl
    Dim tblOutStock As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblOutStock
        .name = "Inventory_Out"
        .Columns.Append "Doc_Number", adVarWChar, 20
        .Columns.Append "DateTime", adVarWChar, 20
        .Columns.Append "Stock_ID", adVarWChar, 2
        .Columns.Append "GroupA", adVarWChar, 3
        .Columns("GroupA").Attributes = adColNullable
        .Columns.Append "ItemNum", adVarWChar, 20
        .Columns.Append "Description", adVarWChar, 50
        .Columns("Description").Attributes = adColNullable
        .Columns.Append "Unit", adVarWChar, 50
        .Columns("Unit").Attributes = adColNullable
        .Columns.Append "Quantity", adDouble
        .Columns.Append "CostPer", adDouble
    End With
    Set CreateTable_Out = tblOutStock
    cat.Tables.Append tblOutStock
    
    cat.Tables.Refresh
    Delay (500)
    
    Set cat = Nothing
    Set tblOutStock = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " CreateTable_Out"
End Function
Public Function CreateTable_InOut() As ADOX.Table
On Error GoTo errHdl
    Dim tblInOutStock As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
     
    With tblInOutStock
        .name = "Inventory_InOut"
        .Columns.Append "Stock_ID", adVarWChar, 2
        .Columns.Append "ItemNum", adVarWChar, 20
        .Columns.Append "Description", adVarWChar, 50
        .Columns("Description").Attributes = adColNullable
        .Columns.Append "Unit", adVarWChar, 50
        .Columns("Unit").Attributes = adColNullable
        .Columns.Append "QuantityFirst", adDouble
        .Columns.Append "CostPerFirst", adDouble
        .Columns.Append "QuantityIn", adDouble
        .Columns.Append "CostPerIn", adDouble
        .Columns.Append "QuantityOut", adDouble
         .Columns.Append "CostPerOut", adDouble
         .Columns.Append "QuantityLast", adDouble
         .Columns.Append "CostPerLast", adDouble
         .Columns.Append "Dept", adVarWChar, 3
         .Columns("Dept").Attributes = adColNullable
    End With
    Set CreateTable_InOut = tblInOutStock
    cat.Tables.Append tblInOutStock
    
    cat.Tables.Refresh
    Delay (500)
    
    Set cat = Nothing
    Set tblInOutStock = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " CreateTable_MoveInOut"
End Function


'Tao bang Ton cuoi thang
Public Function CreateTable_Ton(DateIn As String) As ADOX.Table
On Error GoTo errHdl
    Dim tblTon As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
           
    With tblTon
        .name = "Ton" & Format(DateIn, "00")
        .Columns.Append "ItemNum", adVarWChar, 20
        .Columns.Append "Description", adVarWChar, 50
        .Columns("Description").Attributes = adColNullable
        .Columns.Append "Unit", adVarWChar, 50
        .Columns("Unit").Attributes = adColNullable
        .Columns.Append "Stock_ID", adVarWChar, 10
        .Columns("Stock_ID").Attributes = adColNullable
        .Columns.Append "Quantity", adDouble
        .Columns("Quantity").Attributes = adColNullable
        .Columns.Append "CostPer", adDouble
        .Columns("CostPer").Attributes = adColNullable
        .Columns.Append "Amount", adDouble
        .Columns("Amount").Attributes = adColNullable
    End With
    Set CreateTable_Ton = tblTon
    
     cat.Tables.Append CreateTable_Ton
     cat.Tables.Refresh
     Set cat = Nothing
     
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " CreateTable_Ton"
End Function

Public Function CreateTable_OutStock(DateIn As String) As ADOX.Table
On Error GoTo errHdl
    Dim tblOutStock As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
     
    With tblOutStock
        .name = "Inventory_Out" & Format(Month(DateIn), "00")
        .Columns.Append "Doc_Number", adVarWChar, 20
        .Columns.Append "ItemNum", adVarWChar, 20
        .Columns.Append "Description", adVarWChar, 50
        .Columns("Description").Attributes = adColNullable
        .Columns.Append "Quantity", adDouble
        .Columns.Append "Costper", adDouble
        .Columns.Append "Amount", adDouble
        '.Columns.Append "Cost_Price", adDouble
    End With
    
    Set CreateTable_OutStock = tblOutStock
    cat.Tables.Append tblOutStock
    
    cat.Tables.Refresh
    Delay (500)
    
    Set cat = Nothing
    
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " CreateTable_OutStock"
End Function

Public Function CreateTable_MoveInOut(DateIn As String) As ADOX.Table
On Error GoTo errHdl
    Dim tblMove_In_Out As New ADOX.Table
    With tblMove_In_Out
        .name = "Move_In_Out" & Format(Month(DateIn), "00")
        .Columns.Append "ItemNum", adVarWChar, 20
        .Columns.Append "Description", adVarWChar, 50
        .Columns("Description").Attributes = adColNullable
        .Columns.Append "Store_ID", adVarWChar, 10
        .Columns.Append "In_Quantity", adDouble
        .Columns.Append "Out_Quantity", adDouble
        .Columns.Append "Quantity_Stock", adDouble
        .Columns.Append "Cost_Price", adDouble
    End With
    Set CreateTable_MoveInOut = tblMove_In_Out
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " CreateTable_MoveInOut"
End Function

Public Function CreateTable_Temp() As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    With tblTemp
        .name = "Inventory_Calcu_Temp"
        .Columns.Append "Doc_Number", adVarWChar, 20
        .Columns("Doc_Number").Attributes = adColNullable
        .Columns.Append "DateTime", adVarWChar, 20
        .Columns("DateTime").Attributes = adColNullable
        .Columns.Append "Store_ID", adVarWChar, 10
        .Columns("Store_ID").Attributes = adColNullable
        .Columns.Append "Vendor_Number", adVarWChar, 12
        .Columns("Vendor_Number").Attributes = adColNullable
        .Columns.Append "Org_Doc_Number", adVarWChar, 20
        .Columns("Org_Doc_Number").Attributes = adColNullable
        .Columns.Append "Date_Org", adVarWChar, 20
        .Columns("Date_Org").Attributes = adColNullable
        .Columns.Append "Cashier_ID", adVarWChar, 2
        .Columns("Cashier_ID").Attributes = adColNullable
        .Columns.Append "Delivery_Person", adVarWChar, 50
        .Columns("Delivery_Person").Attributes = adColNullable
        .Columns.Append "Discount", adDouble
        .Columns("Discount").Attributes = adColNullable
        .Columns.Append "iLocked", adBoolean
        .Columns.Append "iReason", adVarWChar, 8
        .Columns("iReason").Attributes = adColNullable
        .Columns.Append "Stock_ID", adVarWChar, 10
        .Columns("Stock_ID").Attributes = adColNullable
        .Columns.Append "ItemNum", adVarWChar, 20
        .Columns.Append "Description", adVarWChar, 50
        .Columns("Description").Attributes = adColNullable
        .Columns.Append "Dept", adVarWChar, 50
        .Columns("Dept").Attributes = adColNullable
        .Columns.Append "Unit", adVarWChar, 20
        .Columns("Unit").Attributes = adColNullable
        .Columns.Append "LastQuantity", adDouble
        .Columns("LastQuantity").Attributes = adColNullable
        .Columns.Append "LastCost", adDouble
        .Columns("LastCost").Attributes = adColNullable
        .Columns.Append "In_Quantity", adDouble
        .Columns("In_Quantity").Attributes = adColNullable
        .Columns.Append "Out_Quantity", adDouble
        .Columns("Out_Quantity").Attributes = adColNullable
        .Columns.Append "Quantity_Stock", adDouble
        .Columns("Quantity_Stock").Attributes = adColNullable
        .Columns.Append "Cost_Price", adDouble
        .Columns("Cost_Price").Attributes = adColNullable
        .Columns.Append "Sale_Price", adDouble
        .Columns("Sale_Price").Attributes = adColNullable
        .Columns.Append "Amount", adDouble
        .Columns("Amount").Attributes = adColNullable
    End With
    Set CreateTable_Temp = tblTemp
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " CreateTable_Temp"
End Function
'Tao bang chuyen gop ban
Public Function Create_Table_Joint_Tranfer() As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "Tranfer_Joint_table"
        .Columns.Append "Org_bill", adDouble
        .Columns.Append "Des_bill", adDouble
        
        .Columns.Append "Org_Location", adVarWChar, 10
        .Columns("Org_Location").Attributes = adColNullable
        .Columns.Append "Des_Location", adVarWChar, 10
        .Columns("Des_Location").Attributes = adColNullable
        
        .Columns.Append "Org_Table", adVarWChar, 50
        .Columns("Org_Table").Attributes = adColNullable
        .Columns.Append "Des_Table", adVarWChar, 5
        .Columns("Des_Table").Attributes = adColNullable
        
        .Columns.Append "DateTime", adVarWChar, 20
        .Columns("DateTime").Attributes = adColNullable
        .Columns.Append "Cashier_ID", adVarWChar, 10
        .Columns("Cashier_ID").Attributes = adColNullable
        .Columns.Append "State", adNumeric
        .Columns("State").Attributes = adColNullable
    End With
    Set Create_Table_Joint_Tranfer = tblTemp
    cat.Tables.Append Create_Table_Joint_Tranfer
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " Create_Table_Joint_Tranfer"
End Function

'Tao bang Khoa so
Public Function Create_Table_Lock() As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "Lock_Month"
        .Columns.Append "Month_Lock", adVarWChar, 6
        .Columns.Append "Date_Lock", adVarWChar, 8
        .Columns.Append "Value", adBoolean
    End With
    Set Create_Table_Lock = tblTemp
    cat.Tables.Append Create_Table_Lock
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " Create_Table_Lock"
End Function

Public Function Create_Pending_Orders() As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "Pending_Orders"
        .Columns.Append "Invoice_Number", adDouble
        .Columns.Append "Station_ID", adVarWChar, 4
        .Columns.Append "Store_ID", adVarWChar, 4
        .Columns.Append "Cashier_ID", adVarWChar, 8
        .Columns.Append "OnHoldID", adVarWChar, 50
        .Columns.Append "Resend", adBoolean
        .Columns.Append "Personal", adDouble
        
    End With
    Set Create_Pending_Orders = tblTemp
    cat.Tables.Append Create_Pending_Orders
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " Create_Pending_Orders"
End Function
Public Function Create_Pending_Orders_item() As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "Pending_Orders_Items"
        .Columns.Append "Invoice_Number", adDouble
        .Columns.Append "ItemNo", adVarWChar, 12
        .Columns.Append "ItemName", adVarWChar, 40
        .Columns.Append "Quan", adDouble
        .Columns.Append "IsModifier", adBoolean
        .Columns.Append "Store_ID", adVarWChar, 10
        .Columns.Append "Price", adDouble
        .Columns.Append "QuanBurned", adDouble
        .Columns.Append "LineNum", adDouble
        .Columns.Append "Kit_Desc", adVarWChar, 50
        .Columns.Append "PrintID", adVarWChar, 3
        
    End With
    Set Create_Pending_Orders_item = tblTemp
    cat.Tables.Append Create_Pending_Orders_item
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " Create_Pending_Orders"
End Function

Public Function Create_MismatchTable() As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "MismatchTable"
        .Columns.Append "ID", adVarWChar, 2
        .Columns.Append "MismatchName", adVarWChar, 50
        .Columns.Append "FromDate", adVarWChar, 12
        .Columns.Append "ToDate", adVarWChar, 12
        
    End With
    Set Create_MismatchTable = tblTemp
    cat.Tables.Append Create_MismatchTable
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " Create_MismatchTable"
End Function
Public Function Create_Attendent() As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "Attendent"
        .Columns.Append "EmpID", adVarWChar, 20
        .Columns.Append "DateTime", adVarWChar, 30
        .Columns.Append "TimeType", adVarWChar, 1
        .Columns.Append "InOutRight", adBoolean
        .Columns.Append "ID", adNumeric
    End With
    Set Create_Attendent = tblTemp
    cat.Tables.Append Create_Attendent
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " Create_Attendent"
End Function

Public Function Create_Table_Reserverd() As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "Table_Reservered"
        .Columns.Append "Reservered_Code", adVarWChar, 20
        .Columns.Append "DateTime", adVarWChar, 20
        .Columns.Append "CustName", adVarWChar, 50
        .Columns.Append "Address", adVarWChar, 100
        .Columns.Append "Phone", adVarWChar, 30
        .Columns.Append "Seat_Num", adDouble
        .Columns.Append "Date_Reservered", adVarWChar, 12
        .Columns.Append "Time_Reservered", adVarWChar, 12
        .Columns.Append "Table_ID", adVarWChar, 30
        .Columns.Append "Amount", adDouble
        .Columns.Append "Description", adVarWChar, 250
        .Columns.Append "Cashier_ID", adVarWChar, 2
        .Columns.Append "IsUsed", adBoolean
    End With
    Set Create_Table_Reserverd = tblTemp
    cat.Tables.Append Create_Table_Reserverd
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " Create_Table_Reserverd"
End Function

Public Function Create_Table_Discount() As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "Promotion"
        .Columns.Append "Pro_ID", adVarWChar, 2
        .Columns.Append "Pro_Name", adVarWChar, 50
        .Columns.Append "Pro_Value", adDouble
                
    End With
    Set Create_Table_Discount = tblTemp
    cat.Tables.Append Create_Table_Discount
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " Create_Table_Discount"
End Function

Public Function Create_Table_Reserved_Details() As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "Table_Reserved_Details"
        .Columns.Append "Reservered_Code", adVarWChar, 20
        .Columns.Append "Table_ID", adVarWChar, 50
        .Columns.Append "ItemNum", adVarWChar, 12
        .Columns.Append "ItemName", adVarWChar, 50
        .Columns.Append "Qty", adDouble
        .Columns.Append "Price", adDouble
        .Columns.Append "Description", adVarWChar, 250
        .Columns("Description").Attributes = adColNullable
        .Columns.Append "KP", adVarWChar, 2
        .Columns.Append "LineDisc", adDouble
        .Columns("LineDisc").Attributes = adColNullable
        .Columns.Append "Line_Disc_Desc", adVarWChar, 100
        .Columns("Line_Disc_Desc").Attributes = adColNullable
        .Columns.Append "Kit_Desc", adVarWChar, 100
        .Columns("Kit_Desc").Attributes = adColNullable
    End With
    Set Create_Table_Reserved_Details = tblTemp
    cat.Tables.Append Create_Table_Reserved_Details
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " Create_Table_Reserved_Details"
End Function

Public Function Create_Promotion_Reason() As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "Promotion_Reason"
        .Columns.Append "ID", adDouble
        .Columns.Append "Pro_Desc", adVarWChar, 200
    End With
    Set Create_Promotion_Reason = tblTemp
    cat.Tables.Append Create_Promotion_Reason
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " Create_Promotion_Reason"

End Function
Public Function Create_Customer_Point_Sale() As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "Customer_Point_Sale"
        .Columns.Append "ID", adVarWChar, 1
        .Columns.Append "TypeMismatch_Name", adVarWChar, 50
        .Columns.Append "Amount_Get_Point", adDouble
        .Columns.Append "Point", adDouble
        .Columns.Append "BirthPoint", adDouble
        .Columns.Append "AmountSale", adDouble
        .Columns.Append "PointSale", adDouble
    End With
    Set Create_Customer_Point_Sale = tblTemp
    cat.Tables.Append Create_Customer_Point_Sale
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " Create_Customer_Point_Sale"
End Function

Public Function Create_Table_Printer_Location() As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "Setup_Printer_Location"
        .Columns.Append "Location_ID", adVarWChar, 2
        .Columns.Append "Receipt_Name", adVarWChar, 100
        .Columns.Append "Printer1", adVarWChar, 100
        .Columns.Append "Printer1_Used", adBoolean
        .Columns.Append "Printer2", adVarWChar, 100
        .Columns.Append "Printer2_Used", adBoolean
        .Columns.Append "Printer3", adVarWChar, 100
        .Columns.Append "Printer3_Used", adBoolean
    End With
    Set Create_Table_Printer_Location = tblTemp
    cat.Tables.Append Create_Table_Printer_Location
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " Create_Table_Printer_Location"
End Function
Public Function Check_Table_exist(TableName As String) As Boolean
On Error GoTo Handle
    Dim cat As New ADOX.Catalog
    Dim bln As Boolean
    Dim i As Integer
    bln = False
    Check_Table_exist = False
    cat.ActiveConnection = myProvider
        For i = 0 To cat.Tables.count - 1
            If UCase(cat.Tables(i).name) = UCase(TableName) Then
                bln = True
                Exit For
            End If
        Next
        Check_Table_exist = bln
    Exit Function
Handle:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "Check_Table_exist"
End Function

Public Function Create_Customer_Type() As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "Customer_Type"
        .Columns.Append "CustType_ID", adVarWChar, 50
        .Columns.Append "CustType_Name", adVarWChar, 100
        .Columns.Append "Promotion", adDouble
        .Columns.Append "Pro_Value", adDouble
        .Columns.Append "Note", adVarWChar, 200
    End With
    Set Create_Customer_Type = tblTemp
    cat.Tables.Append Create_Customer_Type
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " Create_Customer_Type"
End Function

Public Function Check_Field_Exist(rs As ADODB.Recordset, ByVal FieldName As String) As Boolean
    On Error GoTo Handle
        Dim fFound As Boolean
        Dim i As Integer
        fFound = False
        For i = 0 To rs.Fields.count - 1
        DoEvents
            If rs.Fields(i).name = FieldName Then
                fFound = True
                Exit For
            End If
        Next i
        Check_Field_Exist = fFound
    Exit Function
Handle:
    MsgBox Err.Number & Err.Description & " - Check_Field_Exist"
    Check_Field_Exist = False
End Function

Public Function check_Get_Point_Cust() As Boolean
On Error GoTo Handle
    Dim isSale As Boolean
    Dim rsPointsale As New ADODB.Recordset
    Set rsPointsale = Open_Table(cnData, "Customer_Point_Sale")
    With rsPointsale
        If .Fields("ID") = 2 Then
            isSale = True
        End If
    End With
check_Get_Point_Cust = isSale
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & " check_Get_Point_Cust"
    check_Get_Point_Cust = False

End Function

Function Check_Table_On_DB(TableName As String) As Boolean
    Dim rs As New ADODB.Recordset
    Dim kq As Boolean
    kq = False
    Set rs = cnData.OpenSchema(adSchemaTables)
    Do While Not rs.EOF
       If rs("TABLE_TYPE") = "TABLE" Then
            'Kiem tra cac bang trong CSDL co ton tai Table_Name ko?
            If rs("TABLE_NAME") = TableName Then
                kq = True
                Exit Do
            End If
      End If
      rs.MoveNext
    Loop
    Check_Table_On_DB = kq
End Function


Public Function create_tblNgaycong(strDate As String) As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    Dim KEY As New ADOX.KEY
    Dim i As Integer
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "Ngaycong" & strDate
        .Columns.Append "Emp_ID", adVarWChar, 12
        .Columns.Append "Emp_Name", adVarWChar, 50
        For i = 1 To 31
            .Columns.Append Format(i, "00") & "In", adVarWChar, 12
            .Columns(Format(i, "00") & "In").Attributes = adColNullable
            .Columns.Append Format(i, "00") & "Out", adVarWChar, 12
            .Columns(Format(i, "00") & "Out").Attributes = adColNullable
        Next
    End With
    Set create_tblNgaycong = tblTemp
    cat.Tables.Append create_tblNgaycong
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " create_tblNgaycong"
End Function

Public Function Create_Table_Att(strMonth As String) As ADOX.Table
On Error GoTo errHdl
    Dim tblTemp As New ADOX.Table
    Dim cat As New ADOX.Catalog
    Dim KEY As New ADOX.KEY
    cat.ActiveConnection = myProvider
    With tblTemp
        .name = "Att" & strMonth
        .Columns.Append "Emp_ID", adVarWChar, 12
        .Columns.Append "Date_Log", adVarWChar, 12
        .Columns.Append "LogIn_Time", adVarWChar, 12
        .Columns.Append "LogOut_Time", adVarWChar, 12
    End With
    Set Create_Table_Att = tblTemp
    cat.Tables.Append Create_Table_Att
    Set cat = Nothing
    Set tblTemp = Nothing
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & " Create_Table_Att"
End Function
