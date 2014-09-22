Attribute VB_Name = "mdlGeneralStock"
Option Explicit
Public gcnIUSETOWORK As ADODB.Connection
Public gcnIUSETOLOCK  As ADODB.Connection
Public gblnShowQty              As Boolean
Public gblnShowStockInTrans     As Boolean


Public Enum StockManagerSelect
    AveragePrice = 0
    FIFO = 1
    LIFO = 2
End Enum

Public gblnCallPurseFail    As Boolean  'true: that bai,
                                        'false: thanh cong

Public StockManagerType As StockManagerSelect

Private Type InfoTable
    Name As String
    Length As String
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



'*********************************************************
'Chuc nang  :kiem tra trong bang PLU co field MinStock?
'           neu k co thi tao moi
'Tham so vao: strpath: duong dan toi main data
'Tham so ra :khong
'Nguoi tao  :Hai-31/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sCHECK_FIELD_MINSTOCK(ByVal strPath As String, _
            ByVal sPLUCode As String)
On Error GoTo errHdl
    Dim rsPLU As New ADODB.Recordset
    Dim blnFound As Boolean
    Dim i As Integer
    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881")
    
    blnFound = False
    Set rsPLU = cnData.Execute("select * from Inventory")
    'end hai
    
    For i = 0 To rsPLU.Fields.count - 1
        If rsPLU.Fields(i).Name = "MinStock" Then
            blnFound = True
            Exit For
        End If
    Next i
    Set rsPLU = Nothing
    
    If Not blnFound Then
        cnData.Execute "ALTER TABLE PLU " _
                             & "ADD COLUMN MinStock Double;"
    End If
        
    Set rsPLU = cnData.Execute("select MinStock " & _
            "from PLU where PluCode='" & sPLUCode & "'")
    
    If Not (rsPLU.EOF And rsPLU.BOF) Then
        If rsPLU!MinStock & "" = "" Then
            
            cnData.Execute "update Inventory set MinStock=5" & _
            " where PluCode='" & sPLUCode & "'"
            
        End If
    End If
        
    Set rsPLU = Nothing
    Set cnData = Nothing
Exit Sub
errHdl:
    Set rsPLU = Nothing
    Set cnData = Nothing

    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sCHECK_FIELD_MINSTOCK"
End Sub

'*********************************************************
'Chuc nang  : chuyen kieu ngay dd/mm/yyyy (hoac mm/dd/yyyy)
'           thanh kieu chuoi dang dd/mm/yyyy
'Tham so vao: pDatIn: ngay can chuyen
'Tham so ra : ngay kieu chuoi
'Nguoi tao  : Hai-30/10/06
'Nguoi sua  :
'*********************************************************
Public Function gfCONVERT_DATE_TO_STRING_USE_FORMAT( _
        ByVal pDatIn As Date) As String
        
On Error GoTo errHdl
    Dim strRet As String
    
    gfCONVERT_DATE_TO_STRING_USE_FORMAT = ""
    
    strRet = Format(Day(pDatIn), "00") & "/"
    strRet = strRet & Format(Month(pDatIn), "00") & "/"
    strRet = strRet & Year(pDatIn)
    
    gfCONVERT_DATE_TO_STRING_USE_FORMAT = strRet
    
Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gfCONVERT_DATE_TO_STRING_USE_FORMAT"
End Function


'*********************************************************
'Chuc nang  : lay nam tu kieu chuoi dang yyyymmdd
'Tham so vao: pStrDateIn: chuoi ngay
'Tham so ra : nam: kieu integer
'Nguoi tao  : Hai-25/10/06
'Nguoi sua  :
'*********************************************************
Public Function gfGET_YEAR_FROM_STRING_DATE(ByVal pStrDateIn _
                    As String) As Integer
On Error GoTo errHdl
       
    gfGET_YEAR_FROM_STRING_DATE = CInt(Left(pStrDateIn, 4))
    
Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gfGET_YEAR_FROM_STRING_DATE"
End Function
'*********************************************************
'Chuc nang  : lay thang tu kieu chuoi dang yyyymmdd
'Tham so vao: pStrDateIn: chuoi ngay
'Tham so ra : thang: kieu integer
'Nguoi tao  : Hai-25/10/06
'Nguoi sua  :
'*********************************************************
Public Function gfGET_MONTH_FROM_STRING_DATE(ByVal pStrDateIn _
                    As String) As Integer
On Error GoTo errHdl
       
    gfGET_MONTH_FROM_STRING_DATE = CInt(Mid(pStrDateIn, 5, 2))
    
Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gfGET_MONTH_FROM_STRING_DATE"
End Function

'*********************************************************
'Chuc nang  : lay NGAY tu kieu chuoi dang yyyymmdd
'Tham so vao: pStrDateIn: chuoi ngay
'Tham so ra : ngay: kieu integer
'Nguoi tao  : Hai-25/10/06
'Nguoi sua  :
'*********************************************************
Public Function gfGET_DAY_FROM_STRING_DATE(ByVal pStrDateIn _
                    As String) As Integer
On Error GoTo errHdl
       
    gfGET_DAY_FROM_STRING_DATE = CInt(Right(pStrDateIn, 2))
    
Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gfGET_DAY_FROM_STRING_DATE"
End Function

'*********************************************************
'Chuc nang  : tao bang SYYYYMM chua ton kho tinh den 24h
'           cuoi thang MM nam YYYY trong file ALLSTOCK.DAT
'Tham so vao: pStrPath : duong dan RootDir\REPORT
'Tham so ra : khong
'Nguoi tao  : Hai-25/10/06
'Nguoi sua  :
'*********************************************************
Public Sub gsCREATE_PURSE_TABLE(ByVal pStrPath As String, _
                    ByVal pStrYearMonth As String)
On Error GoTo errHdl
       
    Dim cat         As New ADOX.Catalog
    Dim tbl         As New ADOX.Table
    Dim strPath     As String
   
       
    'kiem tra co duong dan nay chua?
    If Dir(pStrPath, vbDirectory) = "" Then
        MkDir pStrPath
    End If
    
    strPath = pStrPath & "\Database.mdb"
    
    'kiem tra co tao file AllStock.Dat nay chua bang nay chua?
    If Dir(strPath) <> "" Then
        cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & _
         ";Jet OLEDB:Database Password=100881;"
    Else
        cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & _
         ";Jet OLEDB:Database Password=100881;"
    End If
    
    'kiem tra bang SYYYYMM co chua?
    For Each tbl In cat.Tables
        If UCase(tbl.Name) = "S" & UCase(pStrYearMonth) Then
            Set tbl = Nothing
            Set cat = Nothing
            Exit Sub
        End If
    Next
        
    'tao bang SYYYYMM
    With tbl
        .Name = "S" & UCase(pStrYearMonth)
        .Columns.Append "Store_ID", adVarWChar, 10
        .Columns.Append "ItemNum", adVarWChar, 50
        .Columns.Append "Qty", adDouble
        .Columns.Append "Purse", adDouble
    End With
    
    cat.Tables.Append tbl
    
    Set tbl = Nothing
    Set cat = Nothing
    
Exit Sub
errHdl:
    Set tbl = Nothing
    Set cat = Nothing

    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gsCREATE_PURSE_TABLE"
End Sub
'*********************************************************
'Chuc nang  : tinh gia von
'Tham so vao:pstrYearFrom,pstrMonthFrom,pstrYearTo,
'               pstrMonthTo
'Tham so ra : khong
'Nguoi tao  : Hai-25/10/06
'Nguoi sua  :
'*********************************************************
Public Sub gsCAL_PURSE(ByVal pStrYearFrom As String, _
        ByVal pStrMonthFrom As String, _
        ByVal pstrYearTo As String, _
        ByVal pstrMonthTo As String)
On Error GoTo errHdl
    Dim strYYYYMM       As String         'nam thang
    
    gblnCallPurseFail = False
    
    'b0 tao 2 bang tam
    Call sCREATE_TABLE_TON_ERROR("Purse.Dat")
    Call sCREATE_TABLE_TON_TEMP("Purse.Dat")
    
    'b1 tinh ton kho dau ky
    Call sGET_STOCK_FIRST(pStrMonthFrom, pStrYearFrom)
    
    'b2 tinh gia von cho cac thang
    strYYYYMM = pStrYearFrom & pStrMonthFrom
    Do While strYYYYMM <= (pstrYearTo & pstrMonthTo)
    
        DoEvents
        'tinh gia von tung thang
        Call sCALL_PURSE_ONE_MONTH(strYYYYMM)
        
        'lay nam thang tiep theo
        strYYYYMM = gfGET_YEAR_MONTH_NEXT( _
                Right(strYYYYMM, 2), Left(strYYYYMM, 4))
    Loop
    
    'b3: in error
    Call sPRINT_TON_ERROR("Purse.Dat", "")
    
    'b4:xoa 2 bang tam
    If Dir(WorkingFolder & "\PURSE.DAT") <> "" Then
        Kill WorkingFolder & "\PURSE.DAT"
    End If
    
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gsCAL_PURSE"
End Sub
'*********************************************************
'Chuc nang  : tao bang tam TON_ERROR trong file PURSE
'             de tinh gia von
'Tham so vao: pStrFileName: ten file can tao
'Tham so ra : khong
'Nguoi tao  : Hai-25/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sCREATE_TABLE_TON_ERROR(ByVal pStrFileName As String)
On Error GoTo errHdl
    Dim cat             As New ADOX.Catalog
    Dim cmdTabl         As New ADODB.Command
    
    Dim strPath         As String
    Dim strSQL          As String
   
 
    strPath = WorkingFolder & pStrFileName
   
    'kiem tra co tao file PURSE.DAT co chua?
    
    cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & _
     ";Jet OLEDB:Database Password=100881;"
    
    Set cat = Nothing
    Sleep 500
       
    'tao bang TON_ERROR
    cmdTabl.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & _
         ";Jet OLEDB:Database Password=100881;"
         
    cmdTabl.CommandText = "CREATE TABLE TON_ERROR ( " & _
            "DateTime NVarChar(8)," & _
            "Doc_Number NVarChar(50)," & _
            "oStore_ID NVarChar(10)," & _
            "ItemNum NVarChar(50)," & _
            "Quantity Float )"
    
    cmdTabl.Execute
    
    Set cmdTabl = Nothing

Exit Sub
errHdl:
    Set cmdTabl = Nothing
    Set cat = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sCREATE_TABLE_TON_ERROR"
End Sub
'*********************************************************
'Chuc nang  : tao bang tam TON_TEMP trong file PURSE
'             de tinh gia von
'Tham so vao: pStrFileName: ten file can tao
'Tham so ra : khong
'Nguoi tao  :
'Nguoi sua  :
'*********************************************************
Private Sub sCREATE_TABLE_TON_TEMP(ByVal pStrFileName As String)
On Error GoTo errHdl
    Dim cat         As New ADOX.Catalog
    Dim tblTemp     As ADOX.Table
    Dim tbl         As New ADOX.Table
    Dim strPath     As String
    
    
    strPath = WorkingFolder & pStrFileName
   
    'kiem tra co tao file PURSE.DAT co chua?
    If Dir(strPath) <> "" Then
        cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & _
         ";Jet OLEDB:Database Password=100881;"
    End If
    
    'kiem tra bang TON_TEMP co chua?
    For Each tblTemp In cat.Tables
        If UCase(tblTemp.Name) = "TON_TEMP" Then
            Set tblTemp = Nothing
            Set tbl = Nothing
            Set cat = Nothing
            Exit Sub
        End If
    Next
        
    'tao bang TON_TEMP
    With tbl
        .Name = "TON_TEMP"
        .Columns.Append "Store_ID", adVarWChar, 10
        .Columns.Append "ItemNum", adVarWChar, 50
        .Columns.Append "Quantity", adDouble
        .Columns.Append "Costper", adDouble
        .Columns("Costper").Attributes = adColNullable
        .Columns.Append "Price", adDouble
        .Columns("Price").Attributes = adColNullable
        
    End With
    
    cat.Tables.Append tbl
    Set tbl = Nothing
    Set tblTemp = Nothing
    Set cat = Nothing

Exit Sub
errHdl:
    Set tbl = Nothing
    Set tblTemp = Nothing
    Set cat = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sCREATE_TABLE_TON_TEMP"
End Sub
'*********************************************************
'Chuc nang  : tinh ton kho dau ky
'Tham so vao: pStrMonthFrom: thang dau
'             pStrYearFrom: nam dau
'Tham so ra : khong
'Nguoi tao  : Hai-25/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sGET_STOCK_FIRST(ByVal pStrMonthFrom As String, _
        ByVal pStrYearFrom As String)
On Error GoTo errHdl
    Dim strSQL          As String
    Dim strPath         As String
    
    Dim strStartDate    As String   'ngay khoi dong
    Dim strYearMonthFirst As String 'nam thang lien truoc
    
    'lay ngay khoi dong
    strStartDate = fGET_START_OR_LOCK_DATE("StartDate")
    
    'copy du lieu ton kho thang truoc, neu co
    If (Mid(strStartDate, 5, 2) <> pStrMonthFrom) Or _
        (Left(strStartDate, 4) <> pStrYearFrom) Then
        
        'lay nam thang lien truoc
        strYearMonthFirst = _
        mdlGeneralStock.gfGET_YEAR_MONTH_PREVIOUS(pStrMonthFrom, _
                            pStrYearFrom)
                            
        'chuyen du lieu tu bang SYYYYMM vao TON_TEMP
        Call sCOPY_STOCK_FIRST(strYearMonthFirst, "Purse.Dat")
        
    End If
    
    
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sGET_STOCK_FIRST"
End Sub

'*********************************************************
'Chuc nang  : tinh thang nam lien truoc
'Tham so vao: pStrMonth: thang
'             pStrYear: nam
'Tham so ra : thang nam lien truoc dang yyyymm
'Nguoi tao  : Hai-25/10/06
'Nguoi sua  :
'*********************************************************
Public Function gfGET_YEAR_MONTH_PREVIOUS(ByVal _
    pStrMonth As String, ByVal pStrYear As String) As String
On Error GoTo errHdl
    Dim intMonth, intYear As Integer
    gfGET_YEAR_MONTH_PREVIOUS = ""
    
    intMonth = CInt(pStrMonth) - 1
    
    'neu la pStrMonth=01
    
    If intMonth = 0 Then
        intYear = CInt(pStrYear) - 1
        
        gfGET_YEAR_MONTH_PREVIOUS = Format(intYear, "0000") & _
                                                    "12"
    Else
        gfGET_YEAR_MONTH_PREVIOUS = pStrYear & _
                Format(intMonth, "00")
    End If
    
    
Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gfGET_YEAR_MONTH_PREVIOUS"
End Function
'*********************************************************
'Chuc nang  : lay ra ngay khoi dong/ngay Lock
'Tham so vao: pStrKindDate: loai ngay StartDate hay LockDate
'Tham so ra : ngay khoi dong kieu chuoi
'Nguoi tao  : Hai-26/10/06
'Nguoi sua  :
'*********************************************************
Private Function fGET_START_OR_LOCK_DATE(ByVal _
        pStrKindDate As String) As String
On Error GoTo errHdl
    Dim rsStartLock     As ADODB.Recordset
    Dim strSQL          As String
    Dim strPath         As String
    
    Dim strdate    As String   'ngay khoi dong
    
    fGET_START_OR_LOCK_DATE = ""
    
    'lay ngay khoi dong
    strSQL = "select " & pStrKindDate & " from StartLock"
    Set rsStartLock = cnData.Execute(strSQL)
    
    strdate = rsStartLock.Fields(0).Value & ""
    
    Set rsStartLock = Nothing
    
    fGET_START_OR_LOCK_DATE = strdate
    
Exit Function
errHdl:
    Set rsStartLock = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - fGET_START_OR_LOCK_DATE"
End Function
'*********************************************************
'Chuc nang  : tao du lieu ton kho dau ky, copy tu
'               bang SYYYYMM vao bang TON_TEMP
'Tham so vao: pStrYYYYMM: thang nam; pStrFileName:ten file data
'Tham so ra : khong
'Nguoi tao  : Hai-26/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sCOPY_STOCK_FIRST(ByVal pStrYYYYMM As String, _
        ByVal pStrFileName As String)
On Error GoTo errHdl

    Dim cnTON_TEMP      As ADODB.Connection
    Dim rsSYYYYMM       As ADODB.Recordset 'SL ton thang nam
    Dim strSQL          As String
    Dim strPath         As String
    
    'mo bang TON_TEMP
    strPath = WorkingFolder & pStrFileName
    Set cnTON_TEMP = Get_Connection(strPath, "100881")
    
    'mo bang SYYYYMM
    strSQL = "select * from S" & pStrYYYYMM
    Set rsSYYYYMM = cnData.Execute(strSQL)
        
    'chuyen du lieu tu bang SYYYYMM vao TON_TEMP
    If Not (rsSYYYYMM.EOF And rsSYYYYMM.BOF) Then
        rsSYYYYMM.MoveFirst
    End If
    
    Do While Not rsSYYYYMM.EOF
        strSQL = "insert into TON_TEMP(Store_ID," & _
                "ItemNum,Quantity,Costper,Price) Values " & _
                "('" & rsSYYYYMM!Store_ID & "','" & _
                rsSYYYYMM!ItemNum & "'," & _
                rsSYYYYMM!Quantity & "," & rsSYYYYMM!Costper & ",0)"
                        
        cnTON_TEMP.Execute strSQL
        rsSYYYYMM.MoveNext
    Loop
    
    Set rsSYYYYMM = Nothing
    Set cnTON_TEMP = Nothing
    
Exit Sub
errHdl:
    Set rsSYYYYMM = Nothing
    Set cnTON_TEMP = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sCOPY_STOCK_FIRST"
End Sub
'*********************************************************
'Chuc nang  : tinh thang nam tiep theo
'Tham so vao: pStrMonth: thang
'             pStrYear: nam
'Tham so ra : thang nam tiep theo dang yyyymm
'Nguoi tao  : Hai-26/10/06
'Nguoi sua  :
'*********************************************************
Public Function gfGET_YEAR_MONTH_NEXT(ByVal _
    pStrMonth As String, ByVal pStrYear As String) As String
    
On Error GoTo errHdl
    Dim intMonth, intYear As Integer
    
    gfGET_YEAR_MONTH_NEXT = ""
    
    intMonth = CInt(pStrMonth) + 1
    
    'neu la pStrMonth=12
    
    If intMonth = 13 Then
        intYear = CInt(pStrYear) + 1
        
        gfGET_YEAR_MONTH_NEXT = Format(intYear, "0000") & _
                                                    "01"
    Else
        gfGET_YEAR_MONTH_NEXT = pStrYear & _
                Format(intMonth, "00")
    End If
    
    
Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gfGET_YEAR_MONTH_NEXT"
End Function

'*********************************************************
'Chuc nang  : tinh gia von hang xuat 1 thang
'Tham so vao: pStrYYYYMM: thang nam
'Tham so ra : khong
'Nguoi tao  : Hai-26/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sCALL_PURSE_ONE_MONTH(ByVal pStrYYYYMM As String)
On Error GoTo errHdl

    'b1.1 lay du lieu tu phieu nhap hang hoa
    Call sGET_PURSE_INSTOCK_PLU(pStrYYYYMM)
     
    'b1.3 tinh don gia von binh quan
    Call sSET_PRICE_IN_TONTEMP
        
   
    'b3.1 tinh gia von cho cac phieu xuat kho HH
    Call sGET_PURSE_OUTSTOCK_PLU(pStrYYYYMM)
    

Exit Sub
errHdl:
    
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sCALL_PURSE_ONE_MONTH"
End Sub

'*********************************************************
'Chuc nang  : tinh gia von hang nhap hang hoa trong 1 thang
'Tham so vao: pStrYYYYMM: thang nam
'Tham so ra : khong
'Nguoi tao  : Hai-26/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sGET_PURSE_INSTOCK_PLU(ByVal pStrYYYYMM As String)
On Error GoTo errHdl
    Dim rsDoc       As ADODB.Recordset
    
    Dim strSQL      As String
    Dim strPath     As String
    
    strSQL = "select distinct Store_ID,Doc_Number " & _
    "from Inventory_In_Master " & _
    "where DateTime='" & pStrYYYYMM & "' and Stock_ID='02'"
    
    Set rsDoc = cnData.Execute(strSQL)
    If Not (rsDoc.EOF And rsDoc.BOF) Then rsDoc.MoveFirst
    
    Do While Not rsDoc.EOF
        DoEvents
        
        'lay cac chi tiet nhap kho cong don vao bang TON_TEMP
        Call sINSERT_INSTOCK_PLU_TO_TONTEMP( _
            rsDoc!Store_ID, rsDoc!Doc_Number, pStrYYYYMM)
            
        rsDoc.MoveNext
    Loop
    
    Set rsDoc = Nothing
Exit Sub
errHdl:
    Set rsDoc = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sGET_PURSE_INSTOCK_PLU"
End Sub

'*********************************************************
'Chuc nang  : them chi tiet nhap kho vao TON_TEMP
'Tham so vao: pStrSiteID: ma site
'           pStrNetID: ma net; pStrYYYYMM: nam thang
'Tham so ra : khong
'Nguoi tao  : Hai-26/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sINSERT_INSTOCK_PLU_TO_TONTEMP( _
    ByVal pStrSiteID As String, _
    ByVal pStrInDocNo As String, ByVal pStrYYYYMM As String)
On Error GoTo errHdl

    Dim rsP_StockIn  As ADODB.Recordset
    Dim iMonth As String
    Dim strSQL      As String
    Dim strPath     As String
    iMonth = Right(pStrYYYYMM, 2)
    strSQL = "select ItemNum,sum(Quantity) as SQty, " & _
    "sum(Amount) as STotal from Inventory_In" & iMonth & _
    "where " & _
    "Doc_Number='" & pStrInDocNo & "' "

    
    Set rsP_StockIn = cnData.Execute(strSQL)
    If Not (rsP_StockIn.EOF And rsP_StockIn.BOF) Then
        rsP_StockIn.MoveFirst
    End If
    
    Do While Not rsP_StockIn.EOF
        DoEvents
        'lay cac chi tiet nhap kho cong vao bang TON_TEMP
        Call sADD_PURSE_TO_TONTEMP(pStrSiteID, rsP_StockIn!ItemNum, CDbl(IIf(rsP_StockIn!Sqty & "" = "", "0", rsP_StockIn!Sqty)), _
            CDbl(IIf(rsP_StockIn!STotal & "" = "", "0", rsP_StockIn!STotal)), "Purse.Dat")
        
        rsP_StockIn.MoveNext
    Loop
    
    Set rsP_StockIn = Nothing
Exit Sub
errHdl:
    Set rsP_StockIn = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sINSERT_INSTOCK_PLU_TO_TONTEMP"
End Sub

'*********************************************************
'Chuc nang  : cong gia von vao TON_TEMP
'Tham so vao: pStrSiteID: ma site
'           pStrNetID: ma net; pStrPluCode: ma hang
'           pDblPurse: gia von;pStrFileName: ten file tam
'Tham so ra : khong
'Nguoi tao  : Hai-26/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sADD_PURSE_TO_TONTEMP( _
    ByVal pStrSiteID As String, _
    ByVal pstrPluCode As String, ByVal pdblQty As Double, _
    ByVal pDblPurse As Double, ByVal pStrFileName As String)
On Error GoTo errHdl

    Dim cnP_TonTemp  As ADODB.Connection
    Dim rsP_TonTemp  As ADODB.Recordset
    
    Dim strSQL      As String
    Dim strPath     As String
    Dim strCON      As String   'dieu kien tim
    
    'lay duong dan file Purse.Dat
    strPath = WorkingFolder & pStrFileName
    
    'neu k co duong dan thi thoat
    If Dir(strPath) = "" Then Exit Sub
    
    'mo bang
    Set cnP_TonTemp = mdlDatabaseUtility.Get_Connection _
                (strPath, "100881")
                
    Set rsP_TonTemp = mdlDatabaseUtility.Open_Table( _
            cnP_TonTemp, "TON_TEMP")
    
    'neu co du lieu
    If Not (rsP_TonTemp.EOF And rsP_TonTemp.BOF) Then
        'tim theo ma hang
        strCON = "ItemNum='" & pstrPluCode & "' " & _
                "and Store_ID='" & pStrSiteID & "' "
        
        rsP_TonTemp.Filter = strCON
            
        'neu tim thay
        If Not rsP_TonTemp.EOF Then
            rsP_TonTemp!Qty = rsP_TonTemp!Qty + pdblQty
            rsP_TonTemp!Purse = rsP_TonTemp!Purse + pDblPurse
            rsP_TonTemp.Update
            
            Set rsP_TonTemp = Nothing
            Set cnP_TonTemp = Nothing
            Exit Sub
        End If
    End If
    
    'neu k co du lieu hay k tim thay thi them moi
    rsP_TonTemp.addNew
    rsP_TonTemp!Store_ID = pStrSiteID
    rsP_TonTemp!ItemNum = pstrPluCode
    rsP_TonTemp!Qty = pdblQty
    rsP_TonTemp!Costper = pDblPurse
    rsP_TonTemp!Price = 0
    rsP_TonTemp.Update
            
    Set rsP_TonTemp = Nothing
    Set cnP_TonTemp = Nothing
Exit Sub
errHdl:
    Set rsP_TonTemp = Nothing
    Set cnP_TonTemp = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sADD_PURSE_TO_TONTEMP"
End Sub

'*********************************************************
'Chuc nang  : tinh gia von hang nhap nguyen lieu trong 1 thang
'Tham so vao: pStrYYYYMM: thang nam
'Tham so ra : khong
'Nguoi tao  : Hai-27/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sGET_PURSE_INSTOCK_SM(ByVal pStrYYYYMM As String)
On Error GoTo errHdl
    Dim rsDoc       As ADODB.Recordset
    
    Dim strSQL      As String
    Dim strPath     As String
    
                
    strSQL = "select distinct Store_ID,Doc_Number " & _
    "from Inventory_In_Master " & _
    "where DateTime='" & pStrYYYYMM & "' and Stock_ID='01'" & _
    "group by Store_ID,Doc_Number"
    
    Set rsDoc = cnData.Execute(strSQL)
    If Not (rsDoc.EOF And rsDoc.BOF) Then rsDoc.MoveFirst
    
    Do While Not rsDoc.EOF
        DoEvents
        
        'lay cac chi tiet nhap kho NL cong don vao bang TON_TEMP
        Call sINSERT_INSTOCK_SM_TO_TONTEMP( _
            rsDoc!Store_ID, _
            rsDoc!Doc_Number, pStrYYYYMM)
            
        rsDoc.MoveNext
    Loop
    
    Set rsDoc = Nothing
Exit Sub
errHdl:
    Set rsDoc = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sGET_PURSE_INSTOCK_SM"
End Sub
'*********************************************************
'Chuc nang  : them chi tiet nhap kho nguyen lieu vao TON_TEMP
'Tham so vao: pStrSiteID: ma site
'           pStrNetID: ma net; pStrYYYYMM: nam thang
'Tham so ra : khong
'Nguoi tao  : Khac Can
'Nguoi sua  :
'*********************************************************
Private Sub sINSERT_INSTOCK_SM_TO_TONTEMP( _
    ByVal pStrSiteID As String, _
    ByVal pStrInDocNo As String, ByVal pStrYYYYMM As String)
On Error GoTo errHdl

    Dim rsP_StockIn  As ADODB.Recordset
    Dim iMonth As String
    Dim strSQL      As String
    Dim strPath     As String
    
    iMonth = Right(pStrYYYYMM, 2)
    
    strSQL = "select ItemNum,sum(Quantity) as SQty, " & _
    "sum(Amount) as STotal from Inventory_In" & iMonth & _
    "where Doc_Number='" & pStrInDocNo & "' " & _
    "group by ItemNum"
    
    Set rsP_StockIn = cnData.Execute(strSQL)
    If Not (rsP_StockIn.EOF And rsP_StockIn.BOF) Then
        rsP_StockIn.MoveFirst
    End If
    
    Do While Not rsP_StockIn.EOF
        DoEvents
        'lay cac chi tiet nhap kho nguyen lieu
        'cong vao bang TON_TEMP
        Call sADD_PURSE_TO_TONTEMP(pStrSiteID, _
            rsP_StockIn!ItemNum, CDbl(IIf(rsP_StockIn!Sqty & "" = "", "0", rsP_StockIn!Sqty)), _
            CDbl(IIf(rsP_StockIn!STotal & "" = "", "0", rsP_StockIn!STotal)), "Purse.Dat")
        
        rsP_StockIn.MoveNext
    Loop
    
    Set rsP_StockIn = Nothing
Exit Sub
errHdl:
    Set rsP_StockIn = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sINSERT_INSTOCK_SM_TO_TONTEMP"
End Sub

'*********************************************************
'Chuc nang  : tinh don gia von trong TON_TEMP
'Tham so vao: khong
'Tham so ra : khong
'Nguoi tao  : Khac Can
'Nguoi sua  :
'*********************************************************
Private Sub sSET_PRICE_IN_TONTEMP()
On Error GoTo errHdl

    Dim cnP_TonTemp  As ADODB.Connection
    Dim rsP_TonTemp  As ADODB.Recordset
    
    Dim strPath     As String
    Dim dblPrice    As Double
    
    'lay duong dan file Purse.Dat
    strPath = WorkingFolder & "\Purse.Dat"
    
    'neu k co duong dan thi thoat
    If Dir(strPath) = "" Then Exit Sub
    
    'mo bang
    Set cnP_TonTemp = mdlDatabaseUtility.Get_Connection _
                (strPath, "100881")
                
    Set rsP_TonTemp = mdlDatabaseUtility.Open_Table( _
            cnP_TonTemp, "TON_TEMP")
    
    'neu co du lieu
    If Not (rsP_TonTemp.EOF And rsP_TonTemp.BOF) Then
        rsP_TonTemp.MoveFirst
    End If
    
    'cap nhat don gia
    Do While Not rsP_TonTemp.EOF
        If rsP_TonTemp!Qty = 0 Then
            dblPrice = 0
            rsP_TonTemp!Purse = 0
        Else
            dblPrice = rsP_TonTemp!Purse / rsP_TonTemp!Qty
        End If
        
        rsP_TonTemp!Price = dblPrice
        rsP_TonTemp.Update
        
        rsP_TonTemp.MoveNext
    Loop
    
    Set rsP_TonTemp = Nothing
    Set cnP_TonTemp = Nothing
Exit Sub
errHdl:
    Set rsP_TonTemp = Nothing
    Set cnP_TonTemp = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sSET_PRICE_IN_TONTEMP"
End Sub


'*********************************************************
'Chuc nang  : tinh gia von cho cac chi tiet xuat chuyen kho
'Tham so vao: pStrSiteID: ma site
'           pStrNetID: ma net; pStrYYYYMM: nam thang
'           pBlnOutStock: true nghia la phieu nay la PXuat
'Tham so ra : khong
'Nguoi tao  : Hai-27/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sSET_PURSE_OUTSTOCK_PLU_DETAIL( _
    ByVal pStrSiteID As String, _
    ByVal pStroDocNo As String, ByVal pStrYYYYMM As String, _
    Optional ByVal pBlnOutStock As Boolean = False)
On Error GoTo errHdl

    Dim rsP_StockOut  As ADODB.Recordset
    Dim iMonth As String
    Dim strSQL      As String
    Dim strPath     As String
    Dim dblQtyIn    As Double
    Dim dblPrice    As Double
    'start hai 17/01/07
    Dim dbloPurse As Double
    
    iMonth = Right(pStrYYYYMM, 2)
    
    strSQL = "select Doc_Number,ItemNum,Quantity, " & _
    "CostPer from Inventory_Out" & iMonth & _
    "where Doc_Number='" & pStroDocNo & "' " & _
    "order by Doc_Number"
    'end hai
    
    Set rsP_StockOut = New ADODB.Recordset
    rsP_StockOut.Open strSQL, cnData, adOpenStatic, adLockOptimistic, adCmdText
    
    If Not (rsP_StockOut.EOF And rsP_StockOut.BOF) Then
        rsP_StockOut.MoveFirst
        
        Do While Not rsP_StockOut.EOF
            DoEvents
            'kiem tra va so sanh so luong xuat va ton kho
            dblQtyIn = fGET_NUM_IN_TONTEMP(pStrSiteID, _
                 rsP_StockOut!ItemNum, "Qty", "Purse.Dat")
            
            
            If Round(rsP_StockOut!Quantity, 3) = Round(dblQtyIn, 3) Then
                'vet kho
                'start hai 17/01/07
                dbloPurse = fGET_NUM_IN_TONTEMP( _
                    pStrSiteID, _
                    rsP_StockOut!ItemNum, "Purse", "Purse.Dat")
                
                
                strSQL = "update Inventory_Out" & iMonth & " Set CostPer=" & _
                    dbloPurse
                    
                cnData.Execute strSQL
                
                'hai sua 07/12/06
                'lay cac chi tiet xuat kho tru don vao bang TON_TEMP
                Call sMINUS_PURSE_TO_TONTEMP(pStrSiteID, _
                rsP_StockOut!ItemNum, rsP_StockOut!Quantity, _
                dbloPurse, "Purse.Dat")
                'end hai
                
                
                If pBlnOutStock = False Then
                    'cap nhat gia von trong nhap kho
                    
                    Call sUPDATE_PURSE_PLU_OUT_TO_IN(pStrSiteID, _
                         rsP_StockOut!Doc_Number, _
                        rsP_StockOut!ItemNum, _
                        dbloPurse, pStrYYYYMM)
                        
                End If
                
            ElseIf Round(rsP_StockOut!Quantity, 3) < Round(dblQtyIn, 3) Then
                'tinh gia von
                dblPrice = fGET_NUM_IN_TONTEMP(pStrSiteID, _
                     rsP_StockOut!ItemNum, "Price", "Purse.Dat")
                
                strSQL = "update Inventory_Out" & iMonth & " Set oPurse=" & _
                     Round(rsP_StockOut!Quantity * dblPrice, 0)

                cnData.Execute strSQL
                
                'lay cac chi tiet xuat kho tru don vao bang TON_TEMP
                Call sMINUS_PURSE_TO_TONTEMP(pStrSiteID, _
                rsP_StockOut!ItemNum, rsP_StockOut!Quantity, _
                Round(rsP_StockOut!Quantity * dblPrice, 0), "Purse.Dat")
                
                
                If pBlnOutStock = False Then
                    'cap nhat gia von trong nhap kho
                        
                    Call sUPDATE_PURSE_PLU_OUT_TO_IN(pStrSiteID, _
                         rsP_StockOut!Doc_Number, _
                        rsP_StockOut!ItemNum, _
                        Round(rsP_StockOut!Quantity * dblPrice, 0), pStrYYYYMM)
                    
                End If
            Else
                'chuyen vao TON_ERROR
                Call fINSERT_DETAIL_IN_TON_ERROR( _
                    rsP_StockOut!oDate, rsP_StockOut!oDoc, _
                    pStrSiteID, _
                    rsP_StockOut!oplucode, _
                    CDbl(IIf(rsP_StockOut!oQty & "" = "", "0", rsP_StockOut!oQty)), "Purse.Dat")
                    
            End If
            
            rsP_StockOut.MoveNext
        Loop
    End If
    
    Set rsP_StockOut = Nothing
Exit Sub
errHdl:
    Set rsP_StockOut = Nothing
    MsgBox Err.Number & " : " & Err.Description _
        & vbCrLf & "mdlGeneralStock : sSET_PURSE_OUTSTOCK_PLU_DETAIL"
End Sub

'*********************************************************
'Chuc nang  : lay so(don gia,so luong,gia von)
'              cua mat hang trong TONTEMP
'Tham so vao: pStrSiteID: ma site;pStrNetID:ma net
'           pStrPluCode:ma hang;pStrFileName: ten file tam
'Tham so ra : don gia
'Nguoi tao  : Hai-27/10/06
'Nguoi sua  :
'*********************************************************
Private Function fGET_NUM_IN_TONTEMP(ByVal pStrSiteID _
        As String, _
        ByVal pstrPluCode As String, _
        ByVal pStrFieldName As String, _
        pStrFileName As String) As Double
        
On Error GoTo errHdl
    Dim cnP_TonTemp  As ADODB.Connection
    Dim rsP_TonTemp  As ADODB.Recordset
    
    Dim strSQL      As String
    Dim strPath     As String
    Dim dblRet      As Double
    
    fGET_NUM_IN_TONTEMP = 0
    
    'lay duong dan file Purse.Dat
    strPath = WorkingFolder & pStrFileName
    
    'neu k co duong dan thi thoat
    If Dir(strPath) = "" Then Exit Function
    
    'mo ket noi
    Set cnP_TonTemp = mdlDatabaseUtility.Get_Connection _
                (strPath, "100881")
    
    strSQL = "select " & pStrFieldName & _
            " from TON_TEMP where " & _
            "Store_ID='" & pStrSiteID & "' and " & _
            "ItemNum='" & pstrPluCode & "'"
                
    Set rsP_TonTemp = cnP_TonTemp.Execute(strSQL)
            
    If rsP_TonTemp.EOF And rsP_TonTemp.BOF Then
        dblRet = 0
    Else
        dblRet = rsP_TonTemp.Fields(pStrFieldName).Value
    End If
        
    Set rsP_TonTemp = Nothing
    Set cnP_TonTemp = Nothing
    
    fGET_NUM_IN_TONTEMP = dblRet
    
Exit Function
errHdl:
    Set rsP_TonTemp = Nothing
    Set cnP_TonTemp = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - fGET_NUM_IN_TONTEMP"
End Function
'*********************************************************
'Chuc nang  : chuyen thong tin vao TON_ERROR
'Tham so vao: pStroDate: ngay xuat;pStroDocNo:so Ctu xuat
'   pStroSiteID :ma site;pStrNetID:ma net;
'   pStrPluCode:ma hang;pDbloQty: so luong xuat;
'   pStrFilename:ten file tam
'Tham so ra : khong co
'Nguoi tao  : Hai-27/10/06
'Nguoi sua  :
'*********************************************************
Private Sub fINSERT_DETAIL_IN_TON_ERROR( _
    ByVal pStroDate As String, ByVal pStroDocNo As String, _
    ByVal pStroSiteID As String, _
    ByVal pStroPluCode As String, ByVal pDbloQty As Double, _
    ByVal pStrFileName As String)
        
On Error GoTo errHdl
    Dim cnP_TonTemp  As ADODB.Connection
    Dim rsP_TonTemp  As ADODB.Recordset
    
    Dim strSQL      As String
    Dim strPath     As String
    
    'lay duong dan file Purse.Dat
    strPath = WorkingFolder & pStrFileName
    
    'neu k co duong dan thi thoat
    If Dir(strPath) = "" Then Exit Sub
    
    'mo ket noi
    Set cnP_TonTemp = mdlDatabaseUtility.Get_Connection _
                (strPath, "100881")
    
    strSQL = "insert into TON_ERROR(DateTime," & _
        "Doc_Number,Store_ID,ItemNum,Quantity) " & _
        "Values ('" & pStroDate & "','" & _
        pStroDocNo & "','" & pStroSiteID & "','" & _
        pStroPluCode & "'," & _
        pDbloQty & ")"
                
     cnP_TonTemp.Execute strSQL
            
    Set rsP_TonTemp = Nothing
    Set cnP_TonTemp = Nothing
        
Exit Sub
errHdl:
    Set rsP_TonTemp = Nothing
    Set cnP_TonTemp = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - fINSERT_DETAIL_IN_TON_ERROR"
End Sub

'*********************************************************
'Chuc nang  : cap nhat gia von cua chi tiet xuat hang hoa
'cua fieu chuyen kho vao phieu nhap kho tuong ung
'Tham so vao: pStrSiteID: ma site xuat,pStrNetID: ma net xuat
'           pStrOutDocNo:so Ctu xuat; pStrPluCode: ma hang
'           pDblPurse:gia von;pStrYYYYMM: nam thang
'Tham so ra : khong
'Nguoi tao  : Hai-26/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sUPDATE_PURSE_PLU_OUT_TO_IN( _
    ByVal pStrSiteID As String, _
    ByVal pStrOutDocNo As String, _
    ByVal pstrPluCode As String, ByVal pDblPurse As Double, _
    ByVal pStrYYYYMM As String)
    
On Error GoTo errHdl

    Dim rsP_StockIn  As ADODB.Recordset
    Dim iMonth As String
    Dim strSiteIDIn As String
    Dim strSQL      As String
    
    strSQL = "update Inventory_In" & iMonth & _
        " set Amount= " & pDblPurse & ",CostPer=" & pDblPurse & "/ Quantity" & _
        " where Doc_Number='" & pStrOutDocNo & "' and" & _
        " ItemNum='" & pstrPluCode & "'"
    
    cnData.Execute strSQL
    
    
    Set rsP_StockIn = Nothing
Exit Sub
errHdl:
    Set rsP_StockIn = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sUPDATE_PURSE_PLU_OUT_TO_IN"
End Sub
'*********************************************************


'*********************************************************
'Chuc nang  : them chi tiet xuat kho vao TON_TEMP
'Tham so vao: pStrSiteID: ma site
'           pStrNetID: ma net; pStrYYYYMM: nam thang
'Tham so ra : khong
'Nguoi tao  : Hai-27/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sINSERT_OUTSTOCK_PLU_TO_TONTEMP( _
    ByVal pStrSiteID As String, _
    ByVal pStrInDocNo As String, ByVal pStrYYYYMM As String)
On Error GoTo errHdl

    Dim cnP_StockIO  As ADODB.Connection
    Dim rsP_StockOut  As ADODB.Recordset
    Dim iMonth As String
    Dim strSQL      As String
    Dim strPath     As String
    iMonth = Right(pStrYYYYMM, 2)
    
    strSQL = "select ItemNum,sum(Quantity) as SQty, " & _
    "sum(Costper) as SPurse from Inventory_Out" & iMonth & _
    "where " & _
    "Doc_Number='" & pStrInDocNo & "' " & _
    "group by ItemNum"
    
    Set rsP_StockOut = cnData.Execute(strSQL)
    If Not (rsP_StockOut.EOF And rsP_StockOut.BOF) Then
        rsP_StockOut.MoveFirst
    End If
    
    Do While Not rsP_StockOut.EOF
        DoEvents
        'lay cac chi tiet xuat kho tru don vao bang TON_TEMP
        Call sMINUS_PURSE_TO_TONTEMP(pStrSiteID, _
                rsP_StockOut!ItemNum, CDbl(IIf(rsP_StockOut!Sqty & "" = "", "0", rsP_StockOut!Sqty)), _
                CDbl(IIf(rsP_StockOut!SPurse & "" = "", "0", rsP_StockOut!SPurse)), "Purse.Dat")
        
        rsP_StockOut.MoveNext
    Loop
    
    Set rsP_StockOut = Nothing
Exit Sub
errHdl:
    Set rsP_StockOut = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sINSERT_OUTSTOCK_PLU_TO_TONTEMP"
End Sub

'*********************************************************
'Chuc nang  : tru don gia von vao TON_TEMP
'Tham so vao: pStrSiteID: ma site
'           pStrNetID: ma net; pStrPluCode: ma hang
'           pDblPurse: gia von;pStrFileName:ten file tam
'Tham so ra : khong
'Nguoi tao  : Hai-26/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sMINUS_PURSE_TO_TONTEMP( _
    ByVal pStrSiteID As String, _
    ByVal pstrPluCode As String, ByVal pdblQty As Double, _
    ByVal pDblPurse As Double, pStrFileName As String)
On Error GoTo errHdl

    Dim cnP_TonTemp  As ADODB.Connection
    Dim rsP_TonTemp  As ADODB.Recordset
    
    Dim strSQL      As String
    Dim strPath     As String
    Dim strCON      As String   'dieu kien tim
    
    'lay duong dan file Purse.Dat
    strPath = WorkingFolder & pStrFileName
    
    'neu k co duong dan thi thoat
    If Dir(strPath) = "" Then Exit Sub
    
    'mo bang
    Set cnP_TonTemp = mdlDatabaseUtility.Get_Connection _
                (strPath, "100881")
                
    Set rsP_TonTemp = mdlDatabaseUtility.Open_Table( _
            cnP_TonTemp, "TON_TEMP")
    
    'neu co du lieu
    If Not (rsP_TonTemp.EOF And rsP_TonTemp.BOF) Then
        'tim theo ma hang
        strCON = "ItemNum='" & pstrPluCode & "' " & _
                "and Store_ID='" & pStrSiteID & "' "
        
        rsP_TonTemp.Filter = strCON
            
        'neu tim thay
        If Not rsP_TonTemp.EOF Then
            rsP_TonTemp!Quantity = Round(CDbl(rsP_TonTemp!Qty), 3) - Round(CDbl(pdblQty), 3)
            rsP_TonTemp!Costper = Round(CDbl(rsP_TonTemp!Purse), 3) - Round(CDbl(pDblPurse), 3)
            rsP_TonTemp.Update
            
            Set rsP_TonTemp = Nothing
            Set cnP_TonTemp = Nothing
            Exit Sub
        Else
            rsP_TonTemp.Filter = adFilterNone
            'start hai 26/12/06
            rsP_TonTemp.addNew
            rsP_TonTemp!Store_ID = pStrSiteID
            rsP_TonTemp!ItemNum = pstrPluCode
            rsP_TonTemp!Quantity = 0 - Round(CDbl(pdblQty), 3)
            rsP_TonTemp!Costper = 0 - Round(CDbl(pDblPurse), 3)
            rsP_TonTemp.Update
            
            'end hai
        End If
     Else
        'start hai 26/12/06
        
        rsP_TonTemp.addNew
        rsP_TonTemp!Store_ID = pStrSiteID
        rsP_TonTemp!ItemNum = pstrPluCode
        rsP_TonTemp!Quantity = 0 - Round(CDbl(pdblQty), 3)
        rsP_TonTemp!Costper = 0 - Round(CDbl(pDblPurse), 3)
        rsP_TonTemp.Update
        
        'end hai
    End If
    
    Set rsP_TonTemp = Nothing
    Set cnP_TonTemp = Nothing
Exit Sub
errHdl:
    Set rsP_TonTemp = Nothing
    Set cnP_TonTemp = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sMINUS_PURSE_TO_TONTEMP"
End Sub




'*********************************************************
'Chuc nang  : tinh gia von xuat kho hang hoa trong 1 thang
'Tham so vao: pStrYYYYMM: thang nam
'Tham so ra : khong
'Nguoi tao  : Hai-28/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sGET_PURSE_OUTSTOCK_PLU(ByVal pStrYYYYMM As String)
On Error GoTo errHdl
    Dim cnStockMng  As ADODB.Connection
    Dim rsDoc       As ADODB.Recordset
    
    Dim strSQL      As String
    Dim strPath     As String
    
                
    strSQL = "select distinct Store_ID,Doc_Number,DateTime " & _
    "from Inventory_Out_Master " & _
    "where DateTime='" & pStrYYYYMM & "'" & _
    "group by Store_ID,Doc_Number,DateTime order by DateTime,Doc_Number"
    
    Set rsDoc = cnData.Execute(strSQL)
    If Not (rsDoc.EOF And rsDoc.BOF) Then rsDoc.MoveFirst
    
    Do While Not rsDoc.EOF
        DoEvents
        
        
        'tinh gia von cho cac mat hang trong phieu
        Call sSET_PURSE_OUTSTOCK_PLU_DETAIL(rsDoc!Store_ID, _
             rsDoc!Doc_Number, pStrYYYYMM, True)
        
        
        rsDoc.MoveNext
    Loop
    
    Set rsDoc = Nothing
Exit Sub
errHdl:
    Set rsDoc = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sGET_PURSE_OUTSTOCK_PLU"
End Sub

'*********************************************************
'Chuc nang  : in danh sach cac mat hang k tinh dc gia von
'Tham so vao: pStrFilename: ten file tam
'Tham so ra : khong
'Nguoi tao  : Hai-28/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sPRINT_TON_ERROR(ByVal pStrFileName As String, _
    pStrToDate As String)
On Error GoTo errHdl
    Dim cnP_TonTemp     As ADODB.Connection
    Dim arrCap()        As String
    Dim strPath         As String
    Dim strdate         As String
    
    'lay duong dan file Purse.Dat
    strPath = WorkingFolder & pStrFileName
    'neu k co duong dan thi thoat
    If Dir(strPath) = "" Then Exit Sub
    
    'tao du lieu in
    Call sCREATE_VIRTUAL_TABLE_TON_ERROR(pStrFileName)
    
    Set cnP_TonTemp = mdlDatabaseUtility.Get_Connection _
                (strPath, "100881")
    
    If fINIT_DATA_TO_PRINT_REPORT_ERR(cnP_TonTemp) = False Then Exit Sub
    
    
    'lay ket noi du lieu
    Dim cmd As New ADODB.Command
    With cmd
        .ActiveConnection = cnP_TonTemp
        .CommandType = adCmdText
        .CommandText = "Select * from REPORT_ERROR " & _
            "order by Store_ID,DateTime,Doc_Number,ItemNum"
        .Execute
    End With
        
    With crPLUPurseError
        If pStrToDate <> "" Then
           
            strdate = gfCONVERT_STRING_TO_DATE(pStrToDate)
                    
            .txtDateFromTo.SetText strdate
            
        Else
            .txtDateFromTo.SetText ""
        End If
        
        .Database.AddADOCommand cnP_TonTemp, cmd
        .PluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .PluName.SetUnboundFieldSource "{ado.Description}"
        .Site.SetUnboundFieldSource "{ado.Store_ID}"
        .DateOpen.SetUnboundFieldSource "{ado.DateTime}"
        .DocNo.SetUnboundFieldSource "{ado.Doc_Number}"
        .Qty.SetUnboundFieldSource "{ado.Quantity}"
        .Unit.SetUnboundFieldSource "{ado.Unit}"
        
        With .Qty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            If DigitsGroup > 0 Then .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        
        .PaperSize = crPaperA4
        .LeftMargin = 1000
    End With
    
    
    With frmShowReport
        .Report = crPLUPurseError
        .WindowState = 2
        .Show vbModal
    End With
    
    Set crPLUPurseError = Nothing
    Set cmd = Nothing
    Set cnP_TonTemp = Nothing
Exit Sub
errHdl:
    Set crPLUPurseError = Nothing
    
    Set cmd = Nothing
    Set cnP_TonTemp = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sPRINT_TON_ERROR"
End Sub

'*********************************************************
'Chuc nang  : tao bang ao de chua noi dung in TON_ERROR
'Tham so vao: pStrFilename: ten file tam
'Tham so ra : khong
'Nguoi tao  : Hai-28/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sCREATE_VIRTUAL_TABLE_TON_ERROR(ByVal pStrFileName As String)
On Error GoTo errHdl
    
    Dim cat_TonErr   As New ADOX.Catalog
    Dim tblReportErr As New ADOX.Table
    
    Dim strSQL      As String
    Dim strPath     As String
    
    strPath = WorkingFolder & "\TEMP\" & pStrFileName
    
    'tao bang tam
    cat_TonErr.ActiveConnection = "Provider=Microsoft.Jet." & _
        "OLEDB.4.0;Data Source=" & strPath & _
        ";Jet OLEDB:Database Password=100881;"
    Sleep 500
    
    With tblReportErr
        .Name = "REPORT_ERROR"
        .Columns.Append "DateTime", adVarWChar, 8
        .Columns.Append "Doc_Number", adVarWChar, 50
        .Columns.Append "Store_ID", adVarWChar, 10
        .Columns.Append "ItemNum", adVarWChar, 50
        .Columns.Append "Description", adVarWChar, 100
        .Columns.Append "Quantity", adDouble
        .Columns.Append "Unit", adVarWChar, 20
    End With
    cat_TonErr.Tables.Append tblReportErr
    
    Sleep 1000
    Set tblReportErr = Nothing
    Set cat_TonErr = Nothing
    
    
Exit Sub
errHdl:
    Set tblReportErr = Nothing
    Set cat_TonErr = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sCREATE_VIRTUAL_TABLE_TON_ERROR"
End Sub

'*********************************************************
'Chuc nang  : tim ten hang hoa/ nguyen lieu
'Tham so vao: pStrSiteID: ma Site;pStrNetID:ma Net
'           pStrCode: ma hang hoa/ nguyen lieu
'           pBlnUnit= true neu tim DVT
'Tham so ra : khong
'Nguoi tao  : Hai-28/10/06
'Nguoi sua  :
'*********************************************************
Public Function gfFIND_NAME_PLU_OR_SM(ByVal pStrSiteID As String, _
         ByVal pstrCode As String, _
        Optional ByVal pBlnUnit As Boolean = False) As String
On Error GoTo errHdl

    Dim rsP_MainDB   As ADODB.Recordset
    
    Dim strSQL      As String
    Dim strPath     As String
    
    gfFIND_NAME_PLU_OR_SM = ""
                
    'lay ten hang hoa
    If pBlnUnit = False Then
        strSQL = "select ItemName from Inventory where " & _
        "ItemNum='" & pstrCode & "'"
    Else
        strSQL = "select Unit from Inventory where " & _
        "ItemNum='" & pstrCode & "'"
    End If
        
    Set rsP_MainDB = cnData.Execute(strSQL)
    
    If rsP_MainDB.EOF And rsP_MainDB.BOF Then
        Set rsP_MainDB = Nothing
    Else
        gfFIND_NAME_PLU_OR_SM = rsP_MainDB.Fields(0).Value
        Exit Function
    End If
    
    'neu k co , lay ten nguyen lieu
    If pBlnUnit = False Then
        strSQL = "select PLUName from SetMPLU where " & _
        "PLUCode='" & pstrCode & "'"
    Else
        strSQL = "select Unit from SetMPLU where " & _
        "PLUCode='" & pstrCode & "'"
    End If
    
    Set rsP_MainDB = cnData.Execute(strSQL)
   
    If Not (rsP_MainDB.EOF And rsP_MainDB.BOF) Then
        gfFIND_NAME_PLU_OR_SM = rsP_MainDB.Fields(0).Value
    End If
    
    Set rsP_MainDB = Nothing
Exit Function
errHdl:
    Set rsP_MainDB = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gfFIND_NAME_PLU_OR_SM"
End Function

'*********************************************************
'Chuc nang  : tao bang ao de chua noi dung in TON_ERROR
'Tham so vao: pCnTon_temp : connection
'Tham so ra : true: co du lieu; false:khong co du lieu
'Nguoi tao  : Hai-28/10/06
'Nguoi sua  :
'*********************************************************
Private Function fINIT_DATA_TO_PRINT_REPORT_ERR( _
        ByVal pCnTon_temp As ADODB.Connection) As Boolean
On Error GoTo errHdl
    Dim rsP_TonTemp     As ADODB.Recordset
    Dim rsP_ReportErr   As New ADODB.Recordset
    
    Dim strSQL      As String
    Dim strPath     As String
    
    fINIT_DATA_TO_PRINT_REPORT_ERR = False
    
    Set rsP_TonTemp = pCnTon_temp.Execute _
                                ("select * from TON_ERROR")
    
    If rsP_TonTemp.EOF And rsP_TonTemp.BOF Then
        Exit Function
    Else
        gblnCallPurseFail = True
        rsP_TonTemp.MoveFirst
    End If
    
    rsP_ReportErr.Open "REPORT_ERROR", pCnTon_temp, adOpenStatic, _
            adLockOptimistic, adCmdTable
    
    
    'them vao bang tam REPORT_ERROR
    Do While Not rsP_TonTemp.EOF
        rsP_ReportErr.addNew
        rsP_ReportErr!DateTime = rsP_TonTemp!DateTime
        rsP_ReportErr!Doc_Number = rsP_TonTemp!Doc_Number
        rsP_ReportErr!Store_ID = rsP_TonTemp!Store_ID
        rsP_ReportErr!ItemNum = rsP_TonTemp!ItemNum
        rsP_ReportErr!Description = gfFIND_NAME_PLU_OR_SM( _
            rsP_TonTemp!Store_ID, _
            rsP_TonTemp!ItemNum)
        rsP_ReportErr!Quantity = rsP_TonTemp!Quantity
        rsP_ReportErr!Unit = gfFIND_NAME_PLU_OR_SM( _
            rsP_TonTemp!Store_ID, _
            rsP_TonTemp!ItemNum, True)
        rsP_ReportErr.Update
        
        rsP_TonTemp.MoveNext
    Loop
    
    Set rsP_ReportErr = Nothing
    Set rsP_TonTemp = Nothing
    fINIT_DATA_TO_PRINT_REPORT_ERR = True
    
Exit Function
errHdl:
    Set rsP_ReportErr = Nothing
    Set rsP_TonTemp = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - fINIT_DATA_TO_PRINT_REPORT_ERR"
End Function

'*********************************************************
'Chuc nang  : tinh ton kho bat ky
'Tham so vao: pStrStockDate : ngay can tinh ton
'Tham so ra : ton kho tinh duoc luu trong file
'           (strRootDir)\Temp\StockAny.Dat
'Nguoi tao  : Hai-30/10/06
'Nguoi sua  :
'*********************************************************
Public Sub gsSTOCK_ANY(ByVal pStrStockDate As String)
On Error GoTo errHdl
    Dim strMocDate          As String   'ngay lam moc
    
    'b0: kiem tra dieu kien
    If fCHECK_DATE_4STOCK_ANY(pStrStockDate) = False Then
        Exit Sub
    End If
    
    'b1: tao table nhap
    Call sCREATE_TABLE_TON_ERROR("StockAny.Dat")
    Call sCREATE_TABLE_TON_TEMP("StockAny.Dat")
    
    'b2: xac dinh so du
    strMocDate = fDEFINE_BALANCE(pStrStockDate)
    If strMocDate = "END" Then Exit Sub
    
    'b3: tinh ton kho
    Call sSTOCK_FROM_DATE_TO_DATE(strMocDate, pStrStockDate)
        
    'b4:in danh sach lam am kho
    Call sPRINT_TON_ERROR("StockAny.Dat", pStrStockDate)
    
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gsSTOCK_ANY"
End Sub

'*********************************************************
'Chuc nang  : kiem tra dieu kien ngay
'Tham so vao: pStrStockDate : ngay can tinh ton
'Tham so ra : ton kho tinh duoc luu trong file
'           (strRootDir)\Temp
'Nguoi tao  : Hai-30/10/06
'Nguoi sua  :
'*********************************************************
Private Function fCHECK_DATE_4STOCK_ANY( _
            ByVal pStrStockDate As String) As Boolean
On Error GoTo errHdl
    Dim rsP_StartLock   As ADODB.Recordset
    
    Dim arrStrMes()     As String
    Dim strPath         As String
    Dim strdate         As String 'kieu chuoi ngay dd/mm/yyyy
    Dim blnRet          As Boolean
    
    blnRet = False
'    arrStrMes = LoadLanguage(LngFile, "#03:049:")
            
    Set rsP_StartLock = cnData.Execute("select * " & _
            "from STARTLOCK")
    
    'neu k co mau tin nao
    If rsP_StartLock.EOF And rsP_StartLock.BOF Then
        MsgBox "Kh«ng cã ngµy khëi ®éng"
    Else
        'neu StartDate la rong
        If rsP_StartLock.Fields("StartDate").Value & "" = "" Then
            MsgBox "Ch­a nhËp kho"
        Else
        'kiem tra ngay can tinh co < StartDate k?
            If pStrStockDate < _
                rsP_StartLock.Fields("StartDate").Value Then
                
                strdate = " " & Right(pStrStockDate, 2) & _
                    "/" & Mid(pStrStockDate, 5, 2) & _
                    "/" & Left(pStrStockDate, 4)
                MsgBox "§· khãa sæ" & _
                    strdate, vbExclamation
            Else
                blnRet = True
            End If
        End If
    End If
    
    Set rsP_StartLock = Nothing
    
    fCHECK_DATE_4STOCK_ANY = blnRet
    
Exit Function
errHdl:
    Set rsP_StartLock = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - fCHECK_DATE_4STOCK_ANY"
End Function

'*********************************************************
'Chuc nang  : tra ve ngay cuoi thang
'Tham so vao: ppstrDate : ngay
'Tham so ra : tra ve ngay cuoi thang
'Nguoi tao  : Tam-30/10/06
'Nguoi sua  :
'*********************************************************
Public Function gfgetLastDateOfMonth(pstrDate As String) As String
On Error GoTo errHdl

    Dim strMonth As String
    Dim strYear As String
    Dim strReturn As String
    
    strMonth = Mid(pstrDate, 5, 2)
    strYear = Left(pstrDate, 4)
    
    Select Case strMonth
        Case "01", "03", "05", "07", "08", "10", "12"
            strReturn = "31"
        Case "04", "06", "09", "11"
            strReturn = "30"
        Case "02"
            If ((CInt(strYear) Mod 4) = 0 And (CInt(strYear) Mod 100) <> 0) _
                Or (CInt(strYear) Mod 400) = 0 Then
                strReturn = "29"
            Else
                strReturn = "28"
            End If
        Case Else
    End Select
    
    gfgetLastDateOfMonth = strReturn
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gfgetLastDateOfMonth"
End Function

'*********************************************************
'Chuc nang  : xac dinh so du lam moc tinh
'Tham so vao: pStrStockDate : ngay
'Tham so ra : tra ve ngay lam moc tinh
'   neu tra ve la "END" thi cham dut thu tuc tinh TON BAT KY
'Nguoi tao  : Hai-30/10/06
'Nguoi sua  :
'*********************************************************
Public Function fDEFINE_BALANCE(pStrStockDate As String) As String
On Error GoTo errHdl
    Dim strMocDate          As String
    Dim strLockDate         As String   'ngay khoa so
    Dim strStartDate        As String   'ngay khoi dong
    Dim strYear, strMonth   As String
    
    'lay ngay
    strMocDate = fGET_START_OR_LOCK_DATE("StartDate")
    strStartDate = fGET_START_OR_LOCK_DATE("StartDate")
    strLockDate = fGET_START_OR_LOCK_DATE("LockDate")
    
    If (strLockDate <> "") And (strStartDate < strLockDate) Then
        If Not (Left(strStartDate, 6) = Left(pStrStockDate, 6)) Then
            If pStrStockDate > strLockDate Then
                strYear = Left(strLockDate, 4)
                strMonth = Mid(strLockDate, 5, 2)
                'chuyen du lieu vao TON_TEMP
                Call sCOPY_STOCK_FIRST(strYear & strMonth, _
                            "StockAny.Dat")
                            
                'tang StrMocDate=LockDate+1
                strMocDate = gfGET_YEAR_MONTH_NEXT( _
                    Mid(strLockDate, 5, 2), Left(strLockDate, 4)) _
                    & "01"
            Else
                'kiem tra co fai la ngay cuoi thang k?
                If Right(pStrStockDate, 2) = _
                    gfgetLastDateOfMonth(pStrStockDate) Then
                    
                    strYear = Left(pStrStockDate, 4)
                    strMonth = Mid(pStrStockDate, 5, 2)
                    'chuyen du lieu vao TON_TEMP
                    Call sCOPY_STOCK_FIRST(strYear & strMonth, _
                            "StockAny.Dat")
                            
                    'xoa table TON_ERROR
                    Call sDELETE_TABLE_IN_TEMP(WorkingFolder & _
                        "\StockAny.Dat", "TON_ERROR")
                        
                    strMocDate = "END"
                Else
                    'lay thang lien truoc
                    strYear = gfGET_YEAR_MONTH_PREVIOUS( _
                        Mid(pStrStockDate, 5, 2), Left(pStrStockDate, 4))
                                        
                    strMonth = Right(strYear, 2)
                    strYear = Left(strYear, 4)
                    'chuyen du lieu vao TON_TEMP
                    Call sCOPY_STOCK_FIRST(strYear & strMonth, _
                            "StockAny.Dat")
                    
                    strMocDate = Left(pStrStockDate, 6) & "01"
                End If
            End If
        End If
    End If
    
    fDEFINE_BALANCE = strMocDate
Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - fDEFINE_BALANCE"
End Function

'*********************************************************
'Chuc nang  : xoa noi dung bang trong file tam
'Tham so vao: pStrFilePath: duong dan ten file,
'               pStrTableName: ten bang
'Tham so ra : khong
'Nguoi tao  : Hai-30/10/06
'Nguoi sua  :
'*********************************************************
Public Sub sDELETE_TABLE_IN_TEMP(ByVal pStrFilePath As String, _
        ByVal pstrTableName As String)
On Error GoTo errHdl
    Dim cnP_Temp    As ADODB.Connection
        
    'kiem tra duong dan
    If Dir(pStrFilePath) = "" Then Exit Sub
    
    Set cnP_Temp = mdlDatabaseUtility.Get_Connection(pStrFilePath, _
            "100881")
    
    cnP_Temp.Execute "delete * from " & pstrTableName
       
    Set cnP_Temp = Nothing
    
Exit Sub
errHdl:
    Set cnP_Temp = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sDELETE_TABLE_IN_TEMP"
End Sub

'*********************************************************
'Chuc nang  : tinh ton kho trong khoang thoi gian
'Tham so vao: pStrDateFrom : ngay bat dau
'           pStrDateTo:ngay ket thuc
'Tham so ra :
'Nguoi tao  : Hai-30/10/06
'Nguoi sua  :
'*********************************************************
Public Sub sSTOCK_FROM_DATE_TO_DATE( _
        ByVal pStrDateFrom As String, ByVal pStrDateTo As String)
On Error GoTo errHdl
    Dim strYYYYMM       As String
    Dim strQuartPath    As String   'duong dan QUY
    Dim strYYYYMMmax    As String   'nam thang cuoi cung
                                    'cua quy pstrDateFrom
    Dim intCoutn        As Integer
    
    'lay nam thang cuoi cua quy
    For intCoutn = 0 To 2
        If (CInt(Mid(pStrDateTo, 5, 2)) + intCoutn) Mod 3 = 0 Then
            strYYYYMMmax = Left(pStrDateTo, 4) + _
                Format(CInt(Mid(pStrDateTo, 5, 2)) + intCoutn, "00")
        End If
    Next
    
    'tinh gia von cho cac thang
    strYYYYMM = Left(pStrDateFrom, 6)
    Do While strYYYYMM <= strYYYYMMmax
    
        DoEvents
       
        'tinh ton cho cac phieu nhap kho HH
        Call sGET_STOCK_IN(Left(strYYYYMM, 4), _
             pStrDateFrom, pStrDateTo)
        
            
        'tinh ton cho cac phieu xuat kho HH
        Call sGET_STOCK_OUT(Left(strYYYYMM, 4), _
             pStrDateFrom, pStrDateTo)
        
        'lay 3 nam thang tiep theo
        strYYYYMM = gfGET_YEAR_MONTH_NEXT( _
                Right(strYYYYMM, 2), Left(strYYYYMM, 4))
        'lay nam thang tiep theo 2
        strYYYYMM = gfGET_YEAR_MONTH_NEXT( _
                Right(strYYYYMM, 2), Left(strYYYYMM, 4))

        'lay nam thang tiep theo 3
        strYYYYMM = gfGET_YEAR_MONTH_NEXT( _
                Right(strYYYYMM, 2), Left(strYYYYMM, 4))
    Loop
    
        
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sSTOCK_FROM_DATE_TO_DATE"
End Sub

'*********************************************************
'Chuc nang  : tinh ton kho hang nhap HH/NL trong 1 thang
'Tham so vao: pStrYear:nam ; pStrQuartFolder: thu muc QUY
'           pstrDateFrom: Tu ngay,pStrDateTo:den ngay
'pblnPLU: =true/False nghia la hang hoa/nguyen lieu
'Tham so ra : khong
'Nguoi tao  : Hai-30/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sGET_STOCK_IN(ByVal pStrYear As String, _
    pStrDateFrom As String, pStrDateTo As String)
    
On Error GoTo errHdl

    Dim rsDoc       As ADODB.Recordset
    
    Dim strSQL      As String
    Dim strPath     As String
    
                
    strSQL = "select distinct Store_ID,Doc_Number " & _
    "from Inventory_In_Master " & _
    " where DateTime>='" & pStrDateFrom & "' and " & _
    "DateTime <='" & pStrDateTo & "'" & _
    "group by Store_ID,Doc_Number"
    
    Set rsDoc = cnData.Execute(strSQL)
    If Not (rsDoc.EOF And rsDoc.BOF) Then rsDoc.MoveFirst
    
    Do While Not rsDoc.EOF
        DoEvents
        
        Call sINSERT_STOCK_IN_TO_TONTEMP(pStrYear, rsDoc!Store_ID, _
                 rsDoc!Doc_Number)
       
        rsDoc.MoveNext
    Loop
    
    Set rsDoc = Nothing
Exit Sub
errHdl:
    Set rsDoc = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sGET_STOCK_IN"
End Sub

'*********************************************************
'Chuc nang  : them chi tiet nhap kho vao TON_TEMP
'Tham so vao: pStrYear: nam; pStrSiteID: ma site;pStrNetID: ma net
'           pStrQuartFolder: thu muc QUY;
'   pblnPLU: = true/False : tinh cho hang hoa/ nguyen lieu
'Tham so ra : khong
'Nguoi tao  : Hai-30/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sINSERT_STOCK_IN_TO_TONTEMP(ByVal pStrYear As String, _
    ByVal pStrSiteID As String, _
    ByVal pStrInDocNo As String)
    
On Error GoTo errHdl
    Dim iMonth As String
    Dim rsP_StockIn  As ADODB.Recordset
    
    Dim strSQL      As String
    Dim strPath     As String
    
    iMonth = mdlGeneralStock.gfGET_MONTH_FROM_STRING_DATE(pStrYear)
       
    strSQL = "select ItemNum,sum(Quantity) as SQty, " & _
    "sum(Amount) as STotal from Inventory_In" & iMonth & _
    "where Doc_Number='" & pStrInDocNo & "' " & _
    "group by ItemNum"
    'end hai
    
    Set rsP_StockIn = cnData.Execute(strSQL)
    If Not (rsP_StockIn.EOF And rsP_StockIn.BOF) Then
        rsP_StockIn.MoveFirst
    End If
    
    Do While Not rsP_StockIn.EOF
        DoEvents
        'lay cac chi tiet nhap kho cong vao bang TON_TEMP
        Call sADD_PURSE_TO_TONTEMP(pStrSiteID, _
            rsP_StockIn!ItemNum, CDbl(IIf(rsP_StockIn!Sqty & "" = "", "0", rsP_StockIn!Sqty)), _
            CDbl(IIf(rsP_StockIn!STotal & "" = "", "0", rsP_StockIn!STotal)), "StockAny.Dat")
        
        rsP_StockIn.MoveNext
    Loop
    
    Set rsP_StockIn = Nothing
Exit Sub
errHdl:
    Set rsP_StockIn = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sINSERT_STOCK_IN_TO_TONTEMP"
End Sub

'*********************************************************
'Chuc nang  : tinh gia von xuat kho HH/NL trong 1 thang
'Tham so vao: pStrYear: nam;pStrQuartFolder: thu muc QUY
'   pstrDateFrom:tu ngay;pStrDateTo:den ngay
'   pblnPLU:= true/false nghia la HH/NL
'Tham so ra : khong
'Nguoi tao  : Hai-30/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sGET_STOCK_OUT(ByVal pStrYear As String, _
    pStrDateFrom As String, pStrDateTo As String)
    
On Error GoTo errHdl
    Dim rsDoc       As ADODB.Recordset
    
    Dim strSQL      As String
    Dim strPath     As String
    
                
    strSQL = "select distinct Store_ID,Doc_Number " & _
    "from Inventory_In_Master" & _
    " where DateTime>='" & pStrDateFrom & "' and " & _
    "DateTime <='" & pStrDateTo & "' " & _
    "group by Store_ID,Doc_Number"
    
    Set rsDoc = cnData.Execute(strSQL)
    If Not (rsDoc.EOF And rsDoc.BOF) Then rsDoc.MoveFirst
    
    Do While Not rsDoc.EOF
        DoEvents
        
            'tru don HH vao bang TON_TEMP
            Call sINSERT_STOCK_OUT_TO_TONTEMP(pStrYear, _
                rsDoc!Store_ID, rsDoc!Doc_Number)
        
        rsDoc.MoveNext
    Loop
    
    Set rsDoc = Nothing
Exit Sub
errHdl:
    Set rsDoc = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sGET_STOCK_OUT"
End Sub

'*********************************************************
'Chuc nang  : tru don cac chi tiet xuat kho vao TON_TEMP
'Tham so vao: pStrYear: nam; pStrSiteID: ma site
'           pStrNetID: ma net; pStrQuartFolder: thu muc QUY
'           pblnPLU: true/false nghia la HH/NL
'Tham so ra : khong
'Nguoi tao  : Hai-30/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sINSERT_STOCK_OUT_TO_TONTEMP(ByVal pStrYear As String, _
    ByVal pStrSiteID As String, _
    ByVal pStroDocNo As String)
On Error GoTo errHdl

    Dim rsP_StockOut  As ADODB.Recordset
    
    Dim strSQL      As String
    Dim strPath     As String
    Dim dblQtyIn    As Double
    Dim dblPrice    As Double
    Dim iMonth As String
    iMonth = gfGET_MONTH_FROM_STRING_DATE(pStrYear)
    
    strSQL = "select Doc_Number,ItemNum,Quantity, " & _
    "CostPer from Inventory_Out" & iMonth & _
    " where Doc_Number='" & pStroDocNo & "' " & _
    "order by Doc_Number"
    'end hai
    
    'mo recordset
    Set rsP_StockOut = cnData.Execute(strSQL)
    
    If Not (rsP_StockOut.EOF And rsP_StockOut.BOF) Then
        rsP_StockOut.MoveFirst
    End If
        
    Do While Not rsP_StockOut.EOF
        DoEvents
        'kiem tra va so sanh so luong xuat va ton kho
        dblQtyIn = fGET_NUM_IN_TONTEMP(pStrSiteID, _
             rsP_StockOut!ItemNum, "Qty", "StockAny.Dat")
        
        'tru don trong TON_TEMP
        Call sMINUS_PURSE_TO_TONTEMP(pStrSiteID, _
            rsP_StockOut!ItemNum, CDbl(IIf(rsP_StockOut!oQty & "" = "", "0", rsP_StockOut!oQty)), _
            CDbl(IIf(rsP_StockOut!oPurse & "" = "", "0", rsP_StockOut!oPurse)), "StockAny.Dat")
                
        If Round(CDbl(IIf(rsP_StockOut!oQty & "" = "", "0", rsP_StockOut!oQty)), 3) > Round(dblQtyIn, 3) Then
           
            'chuyen vao TON_ERROR
            Call fINSERT_DETAIL_IN_TON_ERROR( _
                rsP_StockOut!oDate, rsP_StockOut!oDoc, _
                pStrSiteID, _
                rsP_StockOut!oplucode, _
                dblQtyIn - CDbl(IIf(rsP_StockOut!oQty & "" = "", "0", rsP_StockOut!oQty)), "StockAny.Dat")
                
        End If
        
        rsP_StockOut.MoveNext
    Loop
    
       
    Set rsP_StockOut = Nothing
Exit Sub
errHdl:
    Set rsP_StockOut = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sINSERT_STOCK_OUT_TO_TONTEMP"
End Sub


'*********************************************************
'Chuc nang  : them chi tiet xuat hang hoa
'Tham so vao:cac field trong bang OutStock
'Tham so ra :khong
'Nguoi tao  :Hai-31/10/06
'Nguoi sua  :
'*********************************************************
Private Sub sADD_OUTSTOCK(ByRef rsOStock As ADODB.Recordset, _
    ByVal mDate As String, ByVal mDoc As String, _
    ByVal mPluCode As String, ByVal mQty As Double, _
    ByVal mCost As Double, ByVal MAmount As Double, _
    Optional isSold As Boolean, Optional pCnDB As ADODB.Connection)
    
On Error GoTo errHdl
    Dim strSQL As String
    
    With rsOStock
        If rsOStock Is Nothing Then Exit Sub
        If rsOStock.State = adStateClosed Then Exit Sub
        strSQL = "insert into OutStock(oDate,oDoc,oplucode,oQty," & _
        "oCost,oTotal,oSold,oPurse) values(" & _
        "'" & mDate & "','" & mDoc & "','" & mPluCode & _
        "'," & mQty & "," & mCost & "," & MAmount & _
        "," & isSold & ",0)"
        
        pCnDB.Execute strSQL
                
    End With

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - sADD_OUTSTOCK"
End Sub



'*********************************************************
'Chuc nang  : keim tra chung tu co bi khoa k?
'Tham so vao:pStrSql:cau sql;pcnDoc:ket noi
'Tham so ra :true:co/ false:khong
'Nguoi tao  :Hai-02/11/06
'Nguoi sua  :
'*********************************************************
Public Function gfCHECK_LOCK_DOC(ByVal pStrSQL As String, _
    ByVal pcnDoc As ADODB.Connection) As Boolean
    
On Error GoTo errHdl
    Dim rsTempDoc As ADODB.Recordset
    Dim arrMes()      As String
    
    gfCHECK_LOCK_DOC = False
    
    Set rsTempDoc = pcnDoc.Execute(pStrSQL)
        
    gfCHECK_LOCK_DOC = rsTempDoc.Fields(0).Value
    
    
    If gfCHECK_LOCK_DOC = True Then
        MsgBox "®· khãa chøng tõ "
    End If
    
    Set rsTempDoc = Nothing
    
Exit Function
errHdl:
    Set rsTempDoc = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gfCHECK_LOCK_DOC"
End Function

'*********************************************************
'Chuc nang  : tra ve gia tri kieu chuoi trong record
'Tham so vao:pStrSql:cau sql;pcnDoc:ket noi
'Tham so ra :choui gia tri
'Nguoi tao  :Hai-02/11/06
'Nguoi sua  :
'*********************************************************
Public Function gfRETURN_STRING_VALUE(ByVal pStrSQL As String, _
    ByVal pcnDoc As ADODB.Connection) As String
    
On Error GoTo errHdl
    Dim rsTemp As ADODB.Recordset
    
    gfRETURN_STRING_VALUE = ""
    
    Set rsTemp = pcnDoc.Execute(pStrSQL)
        
    gfRETURN_STRING_VALUE = "                                      " & _
                            rsTemp.Fields(0).Value
    
    
    Set rsTemp = Nothing
    
Exit Function
errHdl:
    Set rsTemp = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gfRETURN_STRING_VALUE"
End Function

'*********************************************************
'Chuc nang  : tra ve gia tri chuoi tien dung de update vao DB
'Tham so vao:pStrValueMoney: chuoi tien;pIntSiteID;pintNetId
'Tham so ra :so tien kieu so
'Nguoi tao  :Hai-02/11/06
'Nguoi sua  :
'*********************************************************


'*********************************************************
'Chuc nang  :xoa ky tu trong chuoi
'Tham so vao:pStrValue: chuoi chua dau ky tu, pstrChar:ky tu
'Tham so ra : chuoi khong con ky tu nua
'Nguoi tao  :Hai-03/11/06
'Nguoi sua  :
'*********************************************************
Public Function gfDELETE_CHAR_IN_STRING(ByVal pStrValue As String, _
                ByVal pstrChar As String) As String
    
On Error GoTo errHdl
    Dim strTemp         As String
    Dim intCout         As Integer
    gfDELETE_CHAR_IN_STRING = ""
    
    intCout = 1
    strTemp = ""
    Do While intCout <= Len(pStrValue)
        If Mid(pStrValue, intCout, 1) <> pstrChar Then
            strTemp = strTemp & Mid(pStrValue, intCout, 1)
        End If
        
        intCout = intCout + 1
    Loop
    
    gfDELETE_CHAR_IN_STRING = strTemp
    
Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gfDELETE_CHAR_IN_STRING"
End Function

'*********************************************************
'Chuc nang  :kiem tra co ai dang khoa so hay k?
'Tham so vao:pStrWorkingPath: duong dan thu muc du lieu
'Tham so ra :true:co nguoi mo/false: k co nguoi nao mo
'Nguoi tao  :Hai-04/11/06
'Nguoi sua  :
'*********************************************************
Public Function gfCHECK_IUSETOLOCK(ByVal pStrWorkingPath _
        As String) As Boolean
    
On Error GoTo errHdl
'start hai 06/01/07
'''    Dim cnTemp          As New ADODB.Connection
'''    Dim rsTemp          As New ADODB.Recordset
    Dim arrMes()        As String
    Dim strPath         As String
    Dim blnRet          As Boolean
    Dim fso         As New FileSystemObject
    Dim fileTemp    As File
    Dim cat As New ADOX.Catalog
    Dim tbl As New ADOX.Table
    Dim gconStrPassDB As String
    gconStrPassDB = "100881"
    blnRet = False
    arrMes = LoadLanguage(LngFile, "#03:050:")
    
    If Dir(pStrWorkingPath & "\Report", vbDirectory) = "" Then Exit Function
    
    strPath = pStrWorkingPath & "\Report\IUSETOLOCK.Dat"
    
    'kiem tra co tao file IUSETOLOCK.Dat nay co chua?
On Error GoTo errHdlUp
    If Dir(strPath) <> "" Then
        Set fileTemp = fso.GetFile(strPath)
        fileTemp.Name = "LOCK.dat"
        fileTemp.Name = "IUSETOLOCK.dat"
        
    Else
        cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & _
         ";Jet OLEDB:Database Password=" & gconStrPassDB & ";"
    
        With tbl
            .Name = "IUSETOLOCK"
            .Columns.Append "MYCORP", adVarWChar, 50
        End With
        cat.Tables.Append tbl
    
        Set tbl = Nothing
        Set cat = Nothing
    
    End If
    
    gfCHECK_IUSETOLOCK = blnRet
    
Exit Function
errHdlUp:
    blnRet = True
'''    cnTemp.RollbackTrans
    Set tbl = Nothing
    Set cat = Nothing
    Set fso = Nothing
    gfCHECK_IUSETOLOCK = blnRet
    
'''    Set rsTemp = Nothing
'''    Set cnTemp = Nothing
    
Exit Function
errHdl:
    Set tbl = Nothing
    Set cat = Nothing
    Set fso = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & " mdlGeneralStock - gfCHECK_IUSETOLOCK "
End Function


'*********************************************************
'Chuc nang  : tao gia tri ban dau cho table IUSETOLOCK,
'   IUSETOWORK
'Tham so vao: pStrPath : duong dan RootDir\REPORT
'Tham so ra : pstrPath duong dan file
'Nguoi tao  : Hai-04/11/06
'Nguoi sua  :
'*********************************************************
Public Sub sINIT_DATA_4TABLE_IUSETO(ByVal pStrPath As String)
On Error GoTo errHdl
    Dim cnTemp      As ADODB.Connection
    Dim rstTemp     As ADODB.Recordset
    Dim strPath     As String
    
    
    'kiem tra co duong dan nay chua?
    If Dir(pStrPath, vbDirectory) = "" Then
        Exit Sub
    End If
    
    Set cnTemp = mdlDatabaseUtility.Get_Connection(pStrPath, _
            "100881")
    
    Set rstTemp = mdlDatabaseUtility.Open_Table(cnTemp, "IUSETOWORK")
    If (rstTemp.EOF And rstTemp.BOF) Then
        rstTemp.addNew
        rstTemp.Fields(0) = "PTV2008"
        rstTemp.Update
    End If
    Set rstTemp = Nothing
    
    Set rstTemp = mdlDatabaseUtility.Open_Table(cnTemp, "IUSETOLOCK")
    If (rstTemp.EOF And rstTemp.BOF) Then
        rstTemp.addNew
        rstTemp.Fields(0) = "PTV2008"
        rstTemp.Update
    End If
    Set rstTemp = Nothing
    Set cnTemp = Nothing
    
Exit Sub
errHdl:
    Set rstTemp = Nothing
    Set cnTemp = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & " mdlGeneralStock - sINIT_DATA_4TABLE_IUSETO "
End Sub


'*********************************************************
'Chuc nang  :dong mo table IUSETOWORK
'Tham so vao:pStrWorkingPath:duong dan thu muc goc ;
'            true:mo;false:dong
'Tham so ra :khong
'Nguoi tao  :Hai-04/11/06
'Nguoi sua  :
'*********************************************************
Public Sub gsOPEN_CLOSE_IUSETOWORK(ByVal pStrWorkingPath As _
        String, Optional ByVal pblnOpen As Boolean = True)
    
On Error GoTo errHdl
    Dim strPath         As String
    
    If pblnOpen = False Then
        If Not (gcnIUSETOWORK Is Nothing) Then
            gcnIUSETOWORK.Close
        End If
        Set gcnIUSETOWORK = Nothing
    End If
    
    If pblnOpen = True Then
        strPath = pStrWorkingPath & "\Report\IUSETOWORK.Dat"
        'kiem tra co tao file IUSETOWORK.Dat nay co chua?
        If Dir(strPath) <> "" Then
            Set gcnIUSETOWORK = mdlDatabaseUtility.Get_Connection( _
                strPath, "100881")
            
        End If
    End If
    
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & " mdlGeneralStock - gsOPEN_CLOSE_IUSETOWORK"
End Sub

'*********************************************************
'Chuc nang  :dong mo table IUSETOLOCK
'Tham so vao:pStrWorkingPath:duong dan thu muc goc ;
'            true:mo;false:dong
'Tham so ra :khong
'Nguoi tao  :Hai-06/11/06
'Nguoi sua  :
'*********************************************************
Public Sub gsOPEN_CLOSE_IUSETOLOCK(ByVal pStrWorkingPath As _
        String, Optional ByVal pblnOpen As Boolean = True)
    
On Error GoTo errHdl
    Dim strPath         As String
    
    'dong
    If pblnOpen = False Then
        If Not (gcnIUSETOLOCK Is Nothing) Then
            gcnIUSETOLOCK.Close
        End If
        Set gcnIUSETOLOCK = Nothing
    End If
    
    'mo
    If pblnOpen = True Then
        strPath = pStrWorkingPath & "\Report\IUSETOLOCK.Dat"
        'kiem tra co tao file IUSETOLOCK.Dat nay co chua?
        If Dir(strPath) <> "" Then
            Set gcnIUSETOLOCK = mdlDatabaseUtility.Get_Connection( _
                strPath, "100881")
            

            
        End If
    End If
    

Exit Sub
errHdl:
'''    Set rsTemp = Nothing
    'end hai
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & " mdlGeneralStock - gsOPEN_CLOSE_IUSETOLOCK "
End Sub

'*********************************************************
'Chuc nang  :kiem tra co ai dang lam viec hay k?
'Tham so vao:pStrWorkingPath: duong dan thu muc du lieu
'Tham so ra :true:co nguoi lam viec/false: k co nguoi lam viec
'Nguoi tao  :Hai-06/11/06
'Nguoi sua  :
'*********************************************************
Public Function gfCHECK_IUSETOWORK(ByVal pStrWorkingPath _
        As String) As Boolean
    
On Error GoTo errHdl
    Dim arrMes()        As String
    Dim strPath         As String
    Dim blnRet          As Boolean
    Dim fso         As New FileSystemObject
    Dim fileTemp As File
    Dim cat As New ADOX.Catalog
    Dim tbl As New ADOX.Table
    
    blnRet = False
    arrMes = LoadLanguage(LngFile, "#03:050:")
    
    If Dir(pStrWorkingPath & "\Report") = "" Then Exit Function
    
    strPath = pStrWorkingPath & "\Report\IUSETOWORK.Dat"
    
    'kiem tra co tao file IUSETOWORK.Dat nay co chua?

On Error GoTo errHdlUp
    If Dir(strPath) <> "" Then
        Set fileTemp = fso.GetFile(strPath)
        fileTemp.Name = "WORK.Dat"
        fileTemp.Name = "IUSETOWORK.Dat"
        
    Else
        'tao file IUSETOWORK.dat
        cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & _
         ";Jet OLEDB:Database Password=" & "100881" & ";"
        
        'tao bang IUSETOWORK
        With tbl
            .Name = "IUSETOWORK"
            .Columns.Append "MYCORP", adVarWChar, 50
        End With
        cat.Tables.Append tbl
        
        Set tbl = Nothing
        Set cat = Nothing
    End If
    
    
    gfCHECK_IUSETOWORK = blnRet
    
Exit Function
errHdlUp:
    blnRet = True
    Set tbl = Nothing
    Set cat = Nothing
    Set fso = Nothing
    gfCHECK_IUSETOWORK = blnRet
    

Exit Function
errHdl:
    Set tbl = Nothing
    Set cat = Nothing
    Set fso = Nothing


    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & " mdlGeneralStock - gfCHECK_IUSETOWORK "
End Function


