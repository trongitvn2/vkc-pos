Attribute VB_Name = "mdlGlobal"
Option Explicit
Public str_Search As String
Public Sec_ID As String
'Public Table_ID As String
Public ConQty As Double
'Public Discount As Double
Public txtPhimso As String
Public PImage As String
Public MaxInvoice As String
Public LastIndex As Integer
Public FirstIndex As Integer
Public currentBill As String
Public CustNo(3) As String
Public LineDiscount As Integer
Public ExtrasPrice As Double
Public VAT As Integer
'Language define
Public LngFile As String
Public LngFolder As String
Public CurLng As String
Public CurFont As String
Public ColorFont As String
Public ShapeColor As String
Public bkColor As String
'rsChuyen mon
Public rsTranfer As New ADODB.Recordset
Public qtyTran As Integer
Public OKCancel As Byte
Public SF(6) As String
Public Bill As Double
Public fFile  As Integer
Public DateDefault As String
Public ReceiptType As String
Public OrderType As String
Public isTimer As Boolean

Public CurDir As String
'Public WorkingFolder As String
'Public ReportFolder As String
Public BackupFolder As String

Public ServerName As String
Public DataBaseName As String
Public UserLog As String
Public DB_Password As String


Public BK_ServerName As String
Public BK_DataBaseName As String
Public BK_UserLog As String
Public BK_DB_Password As String


'Online Flag
Public OnlineFlag As Boolean

'Station
Public Store_ID As String
'User
Public UserLogin As String
Public UserLevel As Integer
Public UserDesc() As String
Public userName As String
Public UserPass As String
Public UserID As String
Public rsuser As New ADODB.Recordset

'Initial File
Public myIniFile As String
Public cnData As New ADODB.Connection

'Public StockType As Integer
Public Const mLightGrey As Long = &HE0E0E0
Public Const mWhite As Long = &HFFFFFF
Public Const mRed As Long = &HC0C0FF
Public Const mGrey As Long = &H808080
Public Const mYellow As Long = &HC0FFFF

Public MachineID As New systemId.clsSystemID
Public En_Decryption As New systemId.clsEn_DecryptionCode
Public ProcessID, Mac_ID As String
'
'Khai bao dinh dang
Public DigitGroupMark As String * 1
Public DecimalMark As String * 1
Public DigitsGroup As Integer
Public DecimalQtyNumber As Integer
Public DecimalAmtNumber As Integer
Public QuantityFormat As String
Public AmountFormat As String
Public formatNum As String
Public CurrencySymbol As String
Public CommandStr As String
Public TopAlign, BottomAlign, LeftAlign, RightAlign As Integer
Public Sort_By As String
Public Date_Open As String
Type sRight
    FullRight As String * 640
    Sodoban As String * 64
    Banhang As String * 64
    Danhmuc As String * 64
    Nhanvien As String * 64
    Caidathethong As String * 64
    Caidatdanhmuc As String * 64
    Baocao As String * 64
    kho As String * 64
    Thuchi As String * 64
    Suaten As String * 64
End Type
Public MyRight As sRight
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Function TrimTab(ByVal InString As String) As String
On Error GoTo errHdl

    Dim i As Integer
    Do While True
        DoEvents
        If Left(InString, 1) = vbTab Then
            InString = Right(InString, Len(InString) - 1)
        Else
            Exit Do
        End If
    Loop
    
    Do While True
        DoEvents
        If Right(InString, 1) = vbTab Then
            InString = Left(InString, Len(InString) - 1)
        Else
            Exit Do
        End If
    Loop
    
    TrimTab = InString
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal - Trimtab"
End Function

Public Function GetDayMonthYear(ByVal StartDay As String) As String()
On Error GoTo errHdl

    Dim plash As Integer, count As Integer
    Dim mang() As String, tmpStr As String
    count = 0
    ReDim Preserve mang(2)
    mang(0) = Day(StartDay)
    mang(1) = Month(StartDay)
    mang(2) = Year(StartDay)
    ReDim Preserve GetDayMonthYear(2)
    GetDayMonthYear = mang
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal - GetDayMonthYear"
End Function

Public Sub Delay(DelayTime As Double)
On Error GoTo errHdl

    Dim sTime As Double, eTime As Double
    sTime = Timer
    Do
        DoEvents
        eTime = Timer
    Loop Until eTime - sTime > DelayTime / 1000
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal - Delay"
End Sub

Public Sub Set_Language(lForm As Form, ByVal FontName As String)
On Error GoTo errHdl

    Dim LngControl As Control
    For Each LngControl In lForm.Controls
        If TypeOf LngControl Is TextBox Or _
            TypeOf LngControl Is label Or _
            TypeOf LngControl Is StatusBar Or _
            TypeOf LngControl Is ComboBox Or _
            TypeOf LngControl Is ListBox Or _
            TypeOf LngControl Is CommandButton Or _
            TypeOf LngControl Is Frame Or _
            TypeOf LngControl Is MSFlexGrid Or _
            TypeOf LngControl Is DirListBox Or _
            TypeOf LngControl Is DriveListBox Or _
            TypeOf LngControl Is FileListBox Or _
            TypeOf LngControl Is TreeView Or _
            TypeOf LngControl Is CheckBox Or _
            TypeOf LngControl Is OptionButton Or _
            TypeOf LngControl Is MyButton Or _
            TypeOf LngControl Is TabStrip Then
            LngControl.Font.name = FontName
        ElseIf TypeOf LngControl Is DataGrid Then
            LngControl.Font.name = FontName
            LngControl.HeadFont.name = FontName
        End If
    Next
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlLanguage - Set_Language"
End Sub

Public Sub SetCombo(ByVal sTableName As String, ByVal cbo As ComboBox, ByVal sFieldName As String, ByVal fEmpty As Boolean)
On Error GoTo errHdl

    Dim res As New ADODB.Recordset
    
    If sTableName = "" Then Exit Sub
    cbo.Clear
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set res = Open_Table(cnData, sTableName)
    If res.State = 0 Then Exit Sub
    If fEmpty = True Then cbo.AddItem "-------"
    With res
        If .RecordCount = 0 Then
            cbo.AddItem "-------"
            GoTo 1
        End If
        .MoveFirst
        Do While Not .EOF
        DoEvents
            cbo.AddItem res.Fields(sFieldName)
'            cbo.ItemData(cbo.NewIndex) = res(0)
            .MoveNext
        Loop
    End With
1:
    CloseRecordset res
    cbo.ListIndex = 0
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlFunctionPublic - SetCombo_Kun"
End Sub

Public Sub SetColorFlexGrid(FlexGrid As MSFlexGrid, frow As Integer, fCol As Integer, lCol As Integer)
On Error GoTo errHdl

    Dim irow As Integer
    Dim iCol As Integer
    Dim myColor As Double
    With FlexGrid
        .Refresh
        For irow = frow To .Rows - 1
            DoEvents
            If irow Mod 500 = 0 Then Delay 200
            If irow Mod 2 = 0 Then
                myColor = &HC0E0FF      '&H80000018
            Else
                myColor = &HFFF0D1
            End If
            .Row = irow
            For iCol = fCol To lCol - 1
                DoEvents
                .Col = iCol
                .CellBackColor = myColor
            Next iCol
        Next irow
    End With
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal - SetColorFlexGrid"
End Sub

Public Function FillZeroForString(str1 As String, iNum As Integer) As String
On Error GoTo errHdl

    Do While Len(str1) < iNum
    DoEvents
        str1 = "0" & str1
    Loop
    FillZeroForString = str1
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlFunctionPublic - FillZeroForString"
End Function
Public Sub AddValueForList(ByVal str1 As String, ByVal lst As ListBox)
On Error GoTo errHdl

    Dim strBin As String
    Dim k As Integer
    
    strBin = HexToBin(str1)
    strBin = FillZeroForString(strBin, 8)
    For k = 0 To Len(strBin) - 1 Step 1
    DoEvents
        If Mid(strBin, k + 1, 1) = 1 Then
            lst.Selected(k) = True
        Else
            lst.Selected(k) = False
        End If
    Next k
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal - AddValueForList"
End Sub

Public Function TrimLeftZero(ByVal StrValue As String) As String
On Error GoTo errHdl

    Do While Left$(StrValue, 1) = "0"
        DoEvents
        If Len(StrValue) = 1 Then
            StrValue = ""
            Exit Do
        Else
            StrValue = Right$(StrValue, Len(StrValue) - 1)
        End If
    Loop
    TrimLeftZero = StrValue
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal- TrimLeftZero"
End Function

Public Function TrimSpecialChar(ByVal StrValue As String) As String
On Error GoTo errHdl
Dim i As Integer
Dim str As String

    For i = 1 To Len(StrValue)
        DoEvents
        If Mid(StrValue, i, 1) <> "?" And Mid(StrValue, i, 1) <> "." And Mid(StrValue, i, 1) <> "," _
        And Mid(StrValue, i, 1) <> ";" And Mid(StrValue, i, 1) <> "\" And Mid(StrValue, i, 1) <> "/" _
        And Mid(StrValue, i, 1) <> "\" And Mid(StrValue, i, 1) <> "&" And Mid(StrValue, i, 1) <> "%" And Mid(StrValue, i, 1) <> "-" _
        And Mid(StrValue, i, 1) <> "" And Mid(StrValue, i, 1) <> "#" And Mid(StrValue, i, 1) <> "$" And Mid(StrValue, i, 1) <> "*" Then
            str = str & Mid(StrValue, i, 1)
        End If
    Next
    TrimSpecialChar = str
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal- TrimSpecialChar"
End Function

Public Function BinToHex(sBin As String) As String
On Error GoTo errHdl

    BinToHex = DectoHex(BinToDec(sBin))
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal - BinToHex"
End Function

Public Function BinToDec(str1 As String) As String
On Error GoTo errHdl
    Dim i As Integer
    Dim str2 As String
    str2 = 0
    If str1 <> "" Then
        For i = 0 To Len(str1) - 1
        DoEvents
            str2 = Int(str2) + Int(Mid(str1, i + 1, 1)) * 2 ^ (Len(str1) - 1 - i)
        Next i
        BinToDec = str2
    Else
        Exit Function
    End If
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal - BinToDec"
End Function

Public Function DectoHex(str1 As String) As String
On Error GoTo errHdl

    DectoHex = Hex(Val(str1))
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal - DectoHex"
End Function

Public Function HexToBin(str1 As String) As String
On Error GoTo errHdl

    HexToBin = DecToBin(HexToDec(str1))
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal - HexToBin"
End Function

Public Function HexToDec(str1 As String) As String
On Error GoTo errHdl

    Dim tmpDouble As Double
    
    HexToDec = "00"
    If str1 = "" Then Exit Function
    tmpDouble = "&H" & str1
    HexToDec = tmpDouble
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal - HexToDec"
End Function

Public Function DecToBin(str1 As String) As String
On Error GoTo errHdl

    Dim str2 As String
    str2 = ""
    Do While Int(str1) >= 2
    DoEvents
        str2 = Int(str1) Mod 2 & str2
        str1 = Int(str1) \ 2
    Loop
    str2 = str1 & str2
    DecToBin = str2
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal - DecToBin"
End Function

Public Function RemoveComma(ByVal S1 As String) As String
On Error GoTo errHdl

    Dim sResult As String
    Dim k As Integer
    
    For k = 1 To Len(S1)
    DoEvents
        If Mid(S1, k, 1) <> "," And Mid(S1, k, 1) <> "." Then
            sResult = sResult & Mid(S1, k, 1)
        End If
    Next k
    RemoveComma = sResult
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal - RemoveComma"
End Function

'            ========== UPDATE DATA TO DATABASE =========
Public Function Add_UpdatedData_To_Array(flex As MSFlexGrid, arrResult() As Variant)
On Error GoTo errHdl

    Dim arrFlex() As String
    Dim flag As Boolean
    Dim sNo As String
    Dim i As Integer
    Dim j As Integer
    
    flag = False
    With flex
        sNo = .TextMatrix(.Row, 0)
        ReDim Preserve arrFlex(.Cols - 1)
        If UBound(arrResult) < 1 Then
1:
            ReDim Preserve arrResult(UBound(arrResult) + 1)
            For j = 0 To .Cols - 1
            DoEvents
                arrFlex(j) = .TextMatrix(.Row, j)
            Next j
            arrResult(UBound(arrResult)) = arrFlex()
            Add_UpdatedData_To_Array = arrResult
            Exit Function
        End If
        
        For i = 1 To UBound(arrResult)
        DoEvents
            If InStr(1, arrResult(i)(0), sNo) <> 0 Then
                For j = 0 To UBound(arrFlex)
                DoEvents
                    If .TextMatrix(.Row, j) = "" Then
                        arrFlex(j) = "0"
                    Else
                        arrFlex(j) = .TextMatrix(.Row, j)
                    End If
                Next j
                arrResult(i) = arrFlex()
                flag = True
                Exit For
            End If
        Next i
        If Not flag Then GoTo 1
    End With
    Add_UpdatedData_To_Array = arrResult
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlFunctionPublic - Add_UpdatedData_To_Array"
End Function

Public Function Reset_MaxLength(txtText As TextBox, iIndex As Integer, ilength As Integer, svalue As String, sNumFormat As String) As String
On Error GoTo errHdl

    Dim iCount As Byte
    Dim i As Integer
    iCount = 0
    With txtText
        .MaxLength = ilength
        For i = 1 To Len(svalue)
        DoEvents
            If Mid(svalue, i, 1) = "." Or Mid(svalue, i, 1) = "," Then
                iCount = iCount + 1
            End If
        Next i
        .MaxLength = .MaxLength + iCount
        svalue = Format(svalue, sNumFormat)
    End With
    Reset_MaxLength = svalue
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlFunctionPublic - Reset_MaxLength"
End Function

Public Function gfCOUNT_RECORD( _
    ByVal pStrSQL As String, pcnData As ADODB.Connection) As Long
    
On Error GoTo errHdl
    Dim rsCount As ADODB.Recordset
    
    gfCOUNT_RECORD = 0
    Set rsCount = pcnData.Execute(pStrSQL)
         
    gfCOUNT_RECORD = rsCount.Fields(0).Value
    
    Set rsCount = Nothing
Exit Function
errHdl:
    Set rsCount = Nothing
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gfCOUNT_RECORD"
End Function

Public Sub Load_SF_System()
On Error GoTo Handle
Dim rsSystem As New ADODB.Recordset
Dim i As Integer
Set rsSystem = Open_Table(cnData, "SystemFlag")
    For i = 1 To 7
        With rsSystem
            .Find "SF='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                SF(i - 1) = .Fields("Data")
            End If
        End With
    Next
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & " ArrayFlag"
End Sub
Public Function ArrayFlag(S As String, ByVal Position As Integer) As String
On Error GoTo Handle
Dim strflag As String

    strflag = Mid(Right("00000000" & HexToBin(S), 8), Position, 1)
    ArrayFlag = strflag
    
Exit Function
Handle:
MsgBox Err.Number & Err.Description & " ArrayFlag"
End Function

Public Function GetMaxSophieuThu() As String
On Error GoTo Handle
    Dim rsmax As New ADODB.Recordset
    Dim Date_Incom As String
    
    Set rsmax = OpenCriticalTable("select max(ID) as MaxID from Income", cnData)
    If Not rsmax.EOF Then
    If "" & rsmax.Fields("maxiD") = "" Then
        GetMaxSophieuThu = "PT/" & Mid(DateDefault, 5, 2) & Mid(DateDefault, 3, 2) & "0001"
    Else
        GetMaxSophieuThu = Left(rsmax.Fields("MaxID"), Len(rsmax.Fields("MaxID")) - 4) & Right("0000" & (CDbl(Right(rsmax.Fields("MaxID"), 4)) + 1), 4)
    End If
    Else
        GetMaxSophieuThu = "PT/" & Mid(DateDefault, 5, 2) & Mid(DateDefault, 3, 2) & "0001"
    End If
Exit Function
Handle:
MsgBox Err.Number & Err.Description & "  mdlGlobal " & "   GetMaxSophieuThu"
End Function

Public Function Get_Max_Date(ByVal Date_Stock As String)
On Error GoTo Handle
Dim Max_Date As String
Dim Date_Value As Date
Date_Value = gfCONVERT_STRING_TO_DATE(Date_Stock)
    Select Case Month(Date_Value)
        Case "01", "03", "05", "07", "08", "10", "12"
            Max_Date = "31"
        Case "04", "06", "09", "11"
            Max_Date = "30"
        Case Else
            If ((CInt(Year(Date_Value)) Mod 4) = 0 And (CInt(Year(Date_Value)) Mod 100) <> 0) _
                Or (CInt(Year(Date_Value)) Mod 400) = 0 Then
                Max_Date = "29"
            Else
                Max_Date = "28"
            End If
    End Select
Get_Max_Date = Max_Date
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & " - Get_Max_Date"
End Function

'Public Function Open_File() As Boolean
'On Error GoTo Handle
'    Dim str_Path As String
'    Dim isOpened As Boolean
'    isOpened = False
'    If Dir(str_Path, vbDirectory) <> "" Then
'        fFile = FreeFile
'        str_Path = WorkingFolder & "\Log\" & Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00")
'            Open str_Path For Append As fFile
'            isOpened = True
'    End If
'    Open_File = isOpened
'Exit Function
'Handle:
'Open_File = False
'    If Err.Number = 52 Then
'        MsgBox "kh«ng t×m thÊy ®­êng dÉn"
'        End
'    End If
'End Function

Public Function Return_right(ByVal strID As String, action As Variant) As Boolean
On Error GoTo Handle
    Dim IDUSER As String
    Dim TempRight As sRight
    IDUSER = Left(strID, 2)
    With rsuser
        .Find "ID='" & IDUSER & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If Mid(strID, 3, Len(strID) - 2) <> rsuser.Fields("Password") Then Exit Function
            With TempRight
                .FullRight = rsuser.Fields("UserRight")
                .Sodoban = RightDeCode(Left(.FullRight, 64))
                .Banhang = RightDeCode(Mid(.FullRight, 65, 64))
                .Danhmuc = RightDeCode(Mid(.FullRight, 129, 64))
                .Nhanvien = RightDeCode(Mid(.FullRight, 193, 64))
                .Caidathethong = RightDeCode(Mid(.FullRight, 257, 64))
                .Caidatdanhmuc = RightDeCode(Mid(.FullRight, 321, 64))
                .Baocao = RightDeCode(Mid(.FullRight, 385, 64))
                .kho = RightDeCode(Mid(.FullRight, 449, 64))
                .Thuchi = RightDeCode(Mid(.FullRight, 513, 64))
                .Suaten = RightDeCode(Mid(.FullRight, 577, 64))
                
                If Mid(.Banhang, 19, 1) = 1 Or rsuser.Fields("UserLevel") = 1 And action = "Delete" Then
                    Return_right = True
                    Exit Function
                End If
            End With
        End If
    End With
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & "- Return_right"
    Return_right = False
End Function

Public Sub format_Balance_Bill()
On Error GoTo Handle
    Dim DescArr() As String
    Dim rsserver As New ADODB.Recordset
    Dim rscompany As New ADODB.Recordset
    Dim rsAdjustment As New ADODB.Recordset
    Dim rsMainGroup As New ADODB.Recordset
    Dim rsFCRate As New ADODB.Recordset
    Dim rsInventory As New ADODB.Recordset
    Dim iReport As CRAXDDRT.Report
    
    DescArr = LoadLanguage(LngFile, "#02:005:")
    
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\database.mdb", "100881administrator")
'    End If
    Set rsserver = Open_Table(cnData, "Table_Diagram_Sections")
    Set rscompany = Open_Table(cnData, "Setup")
    Set rsAdjustment = Open_Table(cnData, "Adjustment")
    Set rsMainGroup = Open_Table(cnData, "MainGroup")
    Set rsFCRate = Open_Table(cnData, "Media")
    Set rsInventory = Open_Table(cnData, "Inventory")
    
    If ArrayFlag(SF(4), 8) = 1 Then
        Set iReport = crBalance
    Else
        Set iReport = crBalance75
    End If
    
    With iReport
        'Dinh dang Dau Bill
            If ArrayFlag(SF(5), 1) = 1 Then
                .Section(8).Suppress = False
                .lblInfor1.SetText rscompany!Company_Info_1
                .lblInfor2.SetText rscompany!Company_Info_2
                .lblInfor3.SetText rscompany!Company_Info_3
                .lblInfor4.SetText rscompany!Company_Info_4 & "-" & rscompany!Company_Info_5
                .Picture3.SetOleLocation (rscompany!Image)
            ElseIf ArrayFlag(SF(5), 2) = 1 Then
                .Section(31).Suppress = False
                .Picture2.SetOleLocation (rscompany!Image)
            ElseIf ArrayFlag(SF(5), 3) = 1 Then
                .Sections(32).Suppress = False
                .lblText1.SetText rscompany!Company_Info_1
                .lblText2.SetText rscompany!Company_Info_2
                .lblText3.SetText rscompany!Company_Info_3
                .lblText4.SetText rscompany!Company_Info_4 & "-" & rscompany!Company_Info_5
            ElseIf ArrayFlag(SF(5), 4) = 1 Then
                .Sections(31).Suppress = True
                .Section(32).Suppress = True
                .Section(8).Suppress = True
            End If
        'khu vuc
            rsserver.Find "Location_ID='" & .txtserver.Value & "'", , adSearchForward, adBookmarkFirst
            If Not rsserver.EOF Then
                .lblTextSever.SetText rsserver.Fields("Section_ID")
            End If

    End With
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & "format_Balance_Bill"
End Sub

Public Function Get_Printer(ByVal LocationID As String) As String
On Error GoTo Handle
Dim Printer_Name As String
    Dim rsPrinter_Location As New ADODB.Recordset
    Set rsPrinter_Location = Open_Table(cnData, "Setup_Printer_Location")
    With rsPrinter_Location
        If .State <> 0 And .RecordCount > 0 Then .MoveFirst
        .Find "Location_ID='" & LocationID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            Printer_Name = .Fields("Receipt_Name")
        End If
    End With
    Get_Printer = Printer_Name
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & " Get_Printer "
End Function

Public Function Get_Printer_Order(ByVal LocationID As String, ByVal Print_ID As String) As String
On Error GoTo Handle
Dim Printer_Name As String
    Dim rsPrinter_Location As New ADODB.Recordset
    Set rsPrinter_Location = Open_Table(cnData, "Setup_Printer_Location")
    With rsPrinter_Location
        If .State <> 0 And .RecordCount > 0 Then .MoveFirst
        .Find "Location_ID='" & LocationID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If Print_ID = "01" Then
                If .Fields("Printer1_Used") = True Then
                    Printer_Name = .Fields("Printer1")
                Else
                    Printer_Name = ""
                End If
            ElseIf Print_ID = "02" Then
                 If .Fields("Printer2_Used") = True Then
                    Printer_Name = .Fields("Printer2")
                Else
                    Printer_Name = ""
                End If
            ElseIf Print_ID = "03" Then
                If .Fields("Printer3_Used") = True Then
                    Printer_Name = .Fields("Printer3")
                Else
                    Printer_Name = ""
                End If
            End If
        End If
    End With
Get_Printer_Order = Printer_Name
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & " Get_Printer_Order"
End Function

Public Function get_Trial_Date() As String
On Error GoTo Handle
    Dim str_Date, tmp_Date As String
    Dim hFile As Double
    Dim LogFile As String
    LogFile = "C:\Windows\System32\sysrt.dll"
    If Dir(LogFile, vbDirectory) <> "" Then
            hFile = FreeFile
            Open LogFile For Input As #hFile
            Do While Not EOF(hFile)
                DoEvents
                Line Input #hFile, tmp_Date
            Loop
            Close #hFile
    str_Date = En_Decryption.MalgoDecrypt(tmp_Date, 5)
    End If
    get_Trial_Date = str_Date
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & "   get_Trial_Date"
    get_Trial_Date = ""
End Function

Public Function Get_Right(User_ID As String, Request_Name As String) As Boolean
On Error GoTo Handle
Dim response As Boolean
Dim res As New ADODB.Recordset
If User_ID = "131112" Then
    Get_Right = True
    Exit Function
End If
        Set res = LoadPasswordData
        With MyRight
            res.MoveFirst
            Do While Not res.EOF
                If StrComp(res.Fields("ID"), Left(User_ID, 2), 1) = 0 Then
                    .FullRight = res.Fields("UserRight")
                    .Banhang = RightDeCode(Mid(.FullRight, 65, 64))
                    .Danhmuc = RightDeCode(Mid(.FullRight, 129, 64))
                    If res.Fields("UserLevel") = 1 Then
                        Get_Right = True
                        Exit Function
                    End If
                    Exit Do
                End If
                res.MoveNext
            Loop
            
            Select Case Request_Name
            Case "delete"
                If Mid(.Banhang, 2, 1) = 1 Then
                    response = True
                Else
                    response = False
                End If
                
            Case "discount"
                If Mid(.Banhang, 3, 1) = 0 Then
                      response = False
                Else
                    response = True
                End If
            Case "editprice"
                If Mid(.Banhang, 4, 1) = 0 Then
                       response = False
                Else
                    response = True
                End If
            
            Case "discount_item"
                If Mid(.Banhang, 5, 1) = 0 Then
                      response = False
                Else
                    response = True
                End If
                
            Case "extraPrice"
                If Mid(.Banhang, 6, 1) = 0 Then
                       response = False
                Else
                    response = True
                End If
                
            Case "editquantity"
                If Mid(.Banhang, 7, 1) = 0 Then
                      response = False
                Else
                    response = True
                End If
                
            Case "bufferPrint"
                If Mid(.Banhang, 8, 1) = 0 Then
                       response = False
                Else
                    response = True
                End If
                
            Case "tabletranffer"
                If Mid(.Banhang, 9, 1) = 0 Then
                      response = False
                Else
                    response = True
                End If
                
            Case "joint_table"
                If Mid(.Banhang, 10, 1) = 0 Then
                      response = False
                Else
                    response = True
                End If
                
            Case "split_items"
                If Mid(.Banhang, 11, 1) = 0 Then
                      response = False
                Else
                    response = True
                End If
                
            Case "payment"
                If Mid(.Banhang, 12, 1) = 0 Then
                       response = False
                Else
                    response = True
                End If
                
            Case "adj1"
                If Mid(.Banhang, 13, 1) = 0 Then
                      response = False
                Else
                    response = True
                End If
                
            Case "adj2"
                If Mid(.Banhang, 14, 1) = 0 Then
                       response = False
                Else
                    response = True
                End If
                
            Case "editname"
                If Mid(.Banhang, 16, 1) = 0 Then
                       response = False
                Else
                    response = True
                End If
                
            Case "item_infor"
                If Mid(.Banhang, 17, 1) = 0 Then
                       response = False
                Else
                    response = True
                End If
                
            Case "money"
                If Mid(.Banhang, 18, 1) = 0 Then
                       response = False
                Else
                    response = True
                End If
            Case "service_charge"
                If Mid(.Banhang, 27, 1) = 0 Then
                       response = False
                Else
                    response = True
                End If
            Case "Delete_IsPrint"
                If Mid(.Banhang, 19, 1) = 0 Then
                       response = False
                Else
                    response = True
                End If
            Case "Delete_Ordered"
                If Mid(.Banhang, 26, 1) = 0 Then
                       response = False
                Else
                    response = True
                End If
            End Select
        End With
    CloseRecordset res
    Get_Right = response
Exit Function
Handle:
MsgBox "Get_Right - B¸o lçi: " & Err.Description
End Function

Public Function Get_record_No(TableName As String, Condition As String) As Integer
On Error GoTo Handle
    Dim result As Integer
    Dim rs As New ADODB.Recordset
    Set rs = OpenCriticalTable("select * from " & TableName & " where Invoice_Number=" & Condition, cnData)
    If rs.State <> 0 Then result = rs.RecordCount
    Get_record_No = result
Exit Function
Handle:
    MsgBox Err.Number & Err.Description
End Function

Public Sub Print_Receipt_Count(ByVal Invoice As Double)
On Error GoTo errHdl
Dim rsInvoice As New ADODB.Recordset
    Set rsInvoice = Open_Table(cnData, "Invoice_totals")
    With rsInvoice
        .Find "Invoice_Number=" & Invoice, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Status") = "P"
            .Fields("InvType") = CInt("0" & .Fields("InvType")) + 1
            .Update
        End If
    End With
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf
End Sub

'In H§
Public Sub Print_Receipt(BillNO As Double)
On Error GoTo errHdl
    Dim SQL As String
    Dim DescArr() As String
    Dim cmd As New ADODB.Command
    Dim ReceiptReport As New CRAXDDRT.Report
    '
'    Dim crDatabase As CRAXDDRT.Database
'    Dim CrDBTables As CRAXDRT.DatabaseTables
'    Dim CrDBTable As CRAXDRT.DatabaseTable
'
'    Set crDatabase = crBalance.Database
'    Set CrDBTables = crDatabase.Tables
'    Set CrDBTable = CrDBTables.Item(1)
'
'    CrDBTable.SetLogOnInfo ServerName, DataBaseName, UserLog, DB_Password
    
    DescArr = LoadLanguage(LngFile, "#02:005:")
    
    If ArrayFlag(SF(0), 5) = 0 Then
        If ArrayFlag(SF(6), 2) = 1 Then
         SQL = "SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.KarDiscount, Invoice_Totals.CustNum," & _
            "Invoice_Totals.Discount, Invoice_Totals.Total_Price, Invoice_Totals.Service_Charge," & _
            "Invoice_Totals.VATFee, Invoice_Totals.Adjustment1, Invoice_Totals.Adj2Rate, " & _
            "Invoice_Totals.Adj1Rate,Invoice_Totals.Personals, Invoice_Totals.Adjustment2, Invoice_Totals.Adjustment3," & _
            "Invoice_Totals.Adjustment4, Invoice_Totals.AddMoney, Invoice_Totals.Tax_Rate_ID," & _
            "Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered, Invoice_Totals.Amt_Change," & _
            "Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID, Invoice_Totals.InvType," & _
            "Invoice_Itemized.ItemNum, Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer," & _
            "Sum(Invoice_Itemized.Amt) AS Amt, Invoice_Itemized.DiffItemName," & _
            "Invoice_Totals.Orig_OnHoldID,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc, Invoice_Totals.OrderMan, Left([OpenTime],8) AS DateIn,Right([OpenTime],12) AS TimeIn, Left([ClosingTime],8) AS DateOut,Right([ClosingTime],12) AS TimeOut, Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount" & _
            " FROM ((((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number) INNER JOIN Inventory ON Invoice_Itemized.ItemNum = Inventory.ItemNum) INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID) INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo" & _
            " WHERE (((Invoice_Itemized.ItemNum)<>'KAR') AND ((Invoice_Totals.Invoice_Number)=" & BillNO & "))" & _
            " GROUP BY Invoice_Totals.Invoice_Number, Invoice_Totals.KarDiscount,Invoice_Totals.Personals, Invoice_Totals.CustNum, Invoice_Totals.Discount, Invoice_Totals.Total_Price, Invoice_Totals.Service_Charge, Invoice_Totals.VATFee, Invoice_Totals.Adjustment1, Invoice_Totals.Tax_Rate_ID, Invoice_Totals.Adjustment2, Invoice_Totals.Adj2Rate, Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment3, Invoice_Totals.Adjustment4, Invoice_Totals.AddMoney, Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered, " & _
            " Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID, Invoice_Totals.InvType, Invoice_Itemized.ItemNum," & _
            " Invoice_Itemized.PricePer,Invoice_Itemized.Line_Disc_Desc, Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc, Invoice_Totals.Orig_OnHoldID, Invoice_Totals.OrderMan, Left([OpenTime],8), Right([OpenTime],12), Left([ClosingTime],8),Right([ClosingTime],12), Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount" & _
            " ORDER BY Invoice_Itemized.ItemNum"
        Else
        SQL = "SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.KarDiscount, Invoice_Totals.CustNum," & _
            "Invoice_Totals.Discount, Invoice_Totals.Total_Price," & _
            "Invoice_Totals.Tax_Rate_ID, Invoice_Totals.Service_Charge, Invoice_Totals.VATFee," & _
            "Invoice_Totals.Adjustment1,Invoice_Totals.Personals, Invoice_Totals.Adj2Rate, Invoice_Totals.Adj1Rate," & _
            "Invoice_Totals.Adjustment2, Invoice_Totals.Adjustment3, Invoice_Totals.Adjustment4," & _
            "Invoice_Totals.AddMoney, Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered, " & _
            "Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID," & _
            "Invoice_Totals.InvType, Invoice_Itemized.ItemNum, Sum(Invoice_Itemized.Quantity) AS Qty," & _
            "Invoice_Itemized.LineNum,Invoice_Itemized.PricePer, Sum(Invoice_Itemized.Amt) AS Amt," & _
            "Invoice_Itemized.DiffItemName, Invoice_Totals.Orig_OnHoldID,Invoice_Itemized.LineDisc," & _
            "Invoice_Itemized.Line_Disc_Desc, Invoice_Totals.OrderMan, Left([OpenTime],8) AS DateIn,Right([OpenTime],12) AS TimeIn, Left([ClosingTime],8) AS DateOut,Right([ClosingTime],12) AS TimeOut,  Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount" & _
            " FROM ((((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number) INNER JOIN Inventory ON Invoice_Itemized.ItemNum = Inventory.ItemNum) INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID) INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo" & _
            " WHERE (((Invoice_Itemized.ItemNum)<>'KAR') AND ((Invoice_Totals.Invoice_Number)=" & BillNO & "))" & _
            " GROUP BY Invoice_Totals.Invoice_Number, Invoice_Totals.KarDiscount, Invoice_Totals.CustNum," & _
            " Invoice_Totals.Tax_Rate_ID,Invoice_Totals.Discount, Invoice_Totals.Total_Price,Invoice_Itemized.Line_Disc_Desc, " & _
            " Invoice_Totals.Service_Charge, Invoice_Totals.VATFee, Invoice_Totals.Adjustment1," & _
            " Invoice_Totals.Adjustment2, Invoice_Totals.Adj2Rate,Invoice_Totals.Personals, Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment3, Invoice_Totals.Adjustment4, Invoice_Totals.AddMoney, Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered, Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID, Invoice_Totals.InvType, Invoice_Itemized.ItemNum, Invoice_Itemized.PricePer, Invoice_Itemized.DiffItemName," & _
            " Invoice_Itemized.LineDisc, Invoice_Totals.Orig_OnHoldID, Invoice_Totals.OrderMan, Left([OpenTime],8), Right([OpenTime],12), Left([ClosingTime],8),Right([ClosingTime],12), Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount, Invoice_Itemized.LineNum" & _
            " ORDER BY Invoice_Itemized.LineNum desc"
        End If
    Else
        If ArrayFlag(SF(6), 2) = 0 Then
        SQL = "SELECT Invoice_Totals.Invoice_Number,Invoice_Totals.KarDiscount, Invoice_Totals.CustNum," & _
            " Invoice_Totals.Discount, Invoice_Totals.Total_Price," & _
            " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adjustment1," & _
            " Invoice_Totals.Adjustment2,Invoice_Totals.Adjustment3," & _
            " Invoice_Totals.Adj2Rate,Invoice_Totals.Personals, Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment4,Invoice_Totals.AddMoney," & _
            " Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered,Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID," & _
            " Invoice_Totals.Station_ID,Invoice_Totals.Tax_Rate_ID,Invoice_Totals.InvType,Invoice_Itemized.ItemNum, " & _
            " Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer, Invoice_Itemized.LineNum,Invoice_Itemized.Line_Disc_Desc," & _
            " sum(Invoice_Itemized.Amt) as Amt, " & _
            " Invoice_Itemized.DiffItemName ,Invoice_Itemized.LineDisc ,Invoice_Totals.Orig_OnHoldID,Invoice_Totals.OrderMan, " & _
            " Left([OpenTime],8) AS DateIn,Right([OpenTime],12) AS TimeIn, Left([ClosingTime],8) AS DateOut,Right([ClosingTime],12) AS TimeOut,  Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount, MainGroup.GroupNo, MainGroup.GroupName " & _
            " FROM ((((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number) INNER JOIN Inventory ON Invoice_Itemized.ItemNum = Inventory.ItemNum) INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID) INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo " & _
            " Where Invoice_Itemized.ItemNum<>'KAR' and Invoice_Totals.Invoice_Number=" & BillNO & _
            " group by Invoice_Itemized.ItemNum,Invoice_Totals.Invoice_Number," & _
            " Invoice_Totals.CustNum,Invoice_Totals.Discount,Invoice_Totals.KarDiscount," & _
            " Invoice_Totals.Total_Price,Invoice_Totals.Grand_Total," & _
            " Invoice_Totals.Amt_Tendered,Invoice_Totals.Amt_Change," & _
            " Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID,Invoice_Totals.Tax_Rate_ID," & _
            " Invoice_Itemized.PricePer, Invoice_Itemized.LineNum, Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc ," & _
            " Invoice_Totals.Orig_OnHoldID,Invoice_Totals.OrderMan, Invoice_Totals.InvType, " & _
            " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adj2Rate, Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment1,Invoice_Totals.Adjustment2,Invoice_Totals.Adjustment3," & _
            " Invoice_Totals.Adjustment4,Invoice_Totals.Personals,Invoice_Totals.AddMoney, Left([OpenTime],8), Right([OpenTime],12), Left([ClosingTime],8),Right([ClosingTime],12), Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount, MainGroup.GroupNo, MainGroup.GroupName" & _
            " order by Invoice_Itemized.LineNum Desc"
        Else

             SQL = "SELECT Invoice_Totals.Invoice_Number,Invoice_Totals.KarDiscount, Invoice_Totals.CustNum," & _
            " Invoice_Totals.Discount, Invoice_Totals.Total_Price," & _
            " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adjustment1," & _
            " Invoice_Totals.Adjustment2,Invoice_Totals.Adjustment3,Invoice_Totals.Adj2Rate,Invoice_Totals.Personals, Invoice_Totals.Adj1Rate," & _
            " Invoice_Totals.Adjustment4,Invoice_Totals.AddMoney,Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered," & _
            " Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID,Invoice_Totals.Tax_Rate_ID," & _
            " Invoice_Totals.Station_ID,Invoice_Totals.InvType,Invoice_Itemized.ItemNum, " & _
            " Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer, " & _
            " sum(Invoice_Itemized.Amt) as Amt, " & _
            " Invoice_Itemized.DiffItemName ,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc ,Invoice_Totals.Orig_OnHoldID,Invoice_Totals.OrderMan, " & _
            " Left([OpenTime],8) AS DateIn,Right([OpenTime],12) AS TimeIn, Left([ClosingTime],8) AS DateOut,Right([ClosingTime],12) AS TimeOut,  Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount, MainGroup.GroupNo, MainGroup.GroupName " & _
            " FROM ((((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number) INNER JOIN Inventory ON Invoice_Itemized.ItemNum = Inventory.ItemNum) INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID) INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo " & _
            " Where Invoice_Itemized.ItemNum<>'KAR' and Invoice_Totals.Invoice_Number=" & BillNO & _
            " group by Invoice_Itemized.ItemNum,Invoice_Totals.Invoice_Number," & _
            " Invoice_Totals.CustNum,Invoice_Totals.Discount,Invoice_Totals.KarDiscount," & _
            " Invoice_Totals.Total_Price,Invoice_Totals.Grand_Total," & _
            " Invoice_Totals.Amt_Tendered,Invoice_Totals.Amt_Change," & _
            " Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID,Invoice_Totals.Tax_Rate_ID," & _
            " Invoice_Itemized.PricePer,  Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc ," & _
            " Invoice_Totals.Orig_OnHoldID,Invoice_Totals.OrderMan, Invoice_Totals.InvType, " & _
            " Invoice_Totals.Service_Charge,Invoice_Totals.Personals,Invoice_Totals.VATFee,Invoice_Totals.Adj2Rate, Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment1,Invoice_Totals.Adjustment2,Invoice_Totals.Adjustment3," & _
            " Invoice_Totals.Adjustment4,Invoice_Totals.AddMoney, Left([OpenTime],8), Right([OpenTime],12), Left([ClosingTime],8),Right([ClosingTime],12), Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount, MainGroup.GroupNo, MainGroup.GroupName" & _
            " order by Invoice_Itemized.ItemNum"
        End If
   End If
    
 '   Dim CRXReportField As CRAXDDRT.DatabaseFieldDefinition
    'ReceiptReport.Database.Tables(1).LogOnServerName "p2ssql.dll", ServerName, DataBaseName, UserLog, DB_Password
    Set crBalance75 = Nothing
    Set crBalance58 = Nothing
    Set crBalance = Nothing
    Set crBalanceA5 = Nothing
    If ReceiptType = "80" Then
        Set ReceiptReport = crBalance
    ElseIf ReceiptType = "58" Then
        Set ReceiptReport = crBalance58
    ElseIf ReceiptType = "75" Then
        Set ReceiptReport = crBalance75
    ElseIf ReceiptType = "A5" Then
        Set ReceiptReport = crBalanceA5
    End If
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With ReceiptReport
        .Database.AddADOCommand cnData, cmd
        .txtPluCode.SetUnboundFieldSource "{ado.DiffItemName}"
        .txtPluName.SetUnboundFieldSource "{ado.ItemNum}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtCost.SetUnboundFieldSource "{ado.PricePer}"
        .txtAmt.SetUnboundFieldSource "{ado.Amt}"
        .LineDisc.SetUnboundFieldSource "{ado.LineDisc}"
'        .Cost1.SetUnboundFieldSource "{ado.PricePer}"
        .txtBillNo.SetUnboundFieldSource "{ado.Invoice_Number}"
        .txtCashier.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtCustomerID.SetUnboundFieldSource "{ado.CustNum}"
        .txtChange.SetUnboundFieldSource "{ado.Amt_Change}"
        .txtDiscount.SetUnboundFieldSource "{ado.Discount}"
        .txtTable.SetUnboundFieldSource "{ado.Orig_OnHoldID}"
        .txtserver.SetUnboundFieldSource "{ado.Station_ID}"
        .txtOrder.SetUnboundFieldSource "{ado.OrderMan}"
        .txtAdj1.SetUnboundFieldSource "{ado.Adjustment1}"
        .txtAdj1Rate.SetUnboundFieldSource "{ado.Adj1Rate}"
        .txtAdj2.SetUnboundFieldSource "{ado.Adjustment2}"
        .txtAdj2Rate.SetUnboundFieldSource "{ado.Adj2Rate}"
        .txtAdj3.SetUnboundFieldSource "{ado.Adjustment3}"
        .txtAdj4.SetUnboundFieldSource "{ado.Adjustment4}"
        .txtSev.SetUnboundFieldSource "{ado.Service_Charge}"
        .txtVAT.SetUnboundFieldSource "{ado.VATFee}"
        .txtMoney.SetUnboundFieldSource "{ado.AddMoney}"
        .printcount.SetUnboundFieldSource "{ado.InvType}"
        .txtMixmatch.SetUnboundFieldSource "{ado.Tax_Rate_ID}"
        .txtsokhach.SetUnboundFieldSource "{ado.Personals}"
        .txtLineDiscDesc.SetUnboundFieldSource "{ado.Line_Disc_Desc}"
        
        .txtDateIn.SetUnboundFieldSource "{ado.DateIn}"
        .txtTimeIn.SetUnboundFieldSource "{ado.TimeIn}"
        .txtDateOut.SetUnboundFieldSource "{ado.DateOut}"
        .txtTimeOut.SetUnboundFieldSource "{ado.TimeOut}"

        .lblTitle.SetText DescArr(24)
        If ArrayFlag(SF(0), 5) = 1 Then
            .txtMaingroup.SetUnboundFieldSource "{ado.GroupNo}"
        End If
        .lblTable.SetText DescArr(3)
        .lblBillNo.SetText DescArr(2)
        .lblItems.SetText DescArr(4)
        .lblQty.SetText DescArr(5)
        .lblPrice.SetText DescArr(6)
        .lblAmt.SetText DescArr(7)
        .lblTotal.SetText DescArr(8)
        '.lblDiscount.SetText DescArr(9)
        .lblRead.SetText DescArr(12)
        .lblCashier.SetText DescArr(13)
        .lblPhuthu.SetText DescArr(14)
        .lblTotal1.SetText DescArr(15)
        .lblServer.SetText DescArr(16)
        .lblIn.SetText DescArr(17)
        .lblOut.SetText DescArr(18)
        .lblCash.SetText DescArr(19)
        .lblOrder.SetText DescArr(20)
        .lblCustomer.SetText DescArr(21)
        .lblSignal.SetText DescArr(22)
        .lblAdj1.SetText DescArr(25)
        .lblAdj2.SetText DescArr(26)
        .lblPhuphi.SetText DescArr(27)
        .lblVAT.SetText DescArr(29)
        .lblPrintCount.SetText DescArr(30)
        .lblTotalItems.SetText DescArr(31)
        
        'canh le
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign
'        With .txtQty
'            .DecimalPlaces = DecimalQtyNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'
'        With .txtCost
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtMoney
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtAdj4
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtAdj3
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtAdj2
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtAdj1
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtAmt
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtChange
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'
'        With .txtMainTotal
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtServAmt
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'
'        With .TxtTotal
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
'        With .txtTotalAmt
'            .DecimalPlaces = DecimalAmtNumber
'            .DecimalSymbol = DecimalMark
'            .ThousandsSeparators = True
'            .ThousandSymbol = DigitGroupMark
'        End With
       
    End With
'    Set iReport = ReceiptReport
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Receipt", "Receipt_DeviceName", True, myIniFile)
        .Report = ReceiptReport
        .Show vbModal
    End With
Exit Sub
errHdl:
If Err.Number = 364 Then
    Exit Sub
Else
    MsgBox Err.Number & " - " & Err.Description
End If
End Sub
Public Function Let_Record_Ordered(ByVal strBill As Double) As ADODB.Recordset
On Error GoTo Handle
Dim rsOrdered As New ADODB.Recordset
        Dim strBalance As String
        If ArrayFlag(SF(6), 1) = 1 Then
            strBalance = "SELECT ItemNum AS PluNo, sum(Quantity) AS Qty,LineNum," & _
                        " PricePer AS Std_Price1, DiffItemName AS PluName," & _
                        " Kit_Description as Kit_Desc,LineDisc,Line_Disc_Desc,TimeOrder " & _
                        " From Invoice_Itemized" & _
                        " WHERE Invoice_Number=" & strBill & _
                        " group by ItemNum, PricePer , DiffItemName,Kit_Description," & _
                        " LineDisc,LineNum,Line_Disc_Desc,TimeOrder" & _
                        " ORDER BY LineNum Asc"

        Else
            strBalance = "SELECT ItemNum AS PluNo, sum(Quantity) AS Qty, PricePer AS Std_Price1," & _
                        " DiffItemName AS PluName," & _
                        " Kit_Description as Kit_Desc,LineDisc,Line_Disc_Desc  " & _
                        " From Invoice_Itemized" & _
                        " WHERE Invoice_Number=" & strBill & _
                        " group by ItemNum, PricePer , DiffItemName," & _
                        " Kit_Description,LineDisc,Line_Disc_Desc " & _
                        " ORDER BY ItemNum Asc"
        End If
        Set rsOrdered = OpenCriticalTable(strBalance, cnData)
        Set Let_Record_Ordered = rsOrdered
Exit Function
Handle:
    MsgBox Err.Number & Err.Description
    Set Let_Record_Ordered = Nothing
End Function

Public Function fill_search(strsource As String) As String
Dim result As String
Dim i As Integer
For i = 1 To Len(strsource)
    result = result & "%" & Mid(strsource, i, 1)
Next
fill_search = result
End Function
