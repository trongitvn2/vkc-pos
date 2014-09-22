Attribute VB_Name = "PublicFunction"
Public DefaultSite As String
Public SiteFile As String

Function readnumber(ByVal S As String) As String
    Dim so
    Dim hang
    Dim tuChuoi As Double
    so = Array("kh«ng", "mét", "hai", "ba", "bèn", "n¨m", "s¸u", "b¶y", "t¸m", "chÝn")
    hang = Array("", "ngh×n", "triÖu", "tû")
    Dim i, j, donvi, chuc, tram As Integer
    Dim str As String
    If S = "" Then Exit Function
    tuChuoi = Round(CDbl(S), 0)
    S = tuChuoi
    str = ""
    i = Len(S)
    For j = 0 To i - 1
        If Left(S, 1) = "0" Then
            S = Right(S, i - j)
        End If
    Next j
    If S = "0" Then
        readnumber = so(0)
        Exit Function
    End If
    i = Len(S)
    If i = 0 Then
        str = ""
    Else
        j = 0
        Do While i > 0
            donvi = Int(Mid(S, i, 1))
            i = i - 1
            If i > 0 Then
                chuc = Int(Mid(S, i, 1))
            Else
                chuc = -1
            End If
            i = i - 1
            If i > 0 Then
                tram = Int(Mid(S, i, 1))
            Else
                tram = -1
            End If
            i = i - 1
            If donvi > 0 Or chuc > 0 Or tram > 0 Or j = 3 Then
                str = hang(j) & " " & str
            End If
            j = j + 1
            If j > 3 Then
                j = 1
            End If
            If donvi = 1 And chuc > 1 Then
                str = "mèt" & " " & str
            Else
                If donvi = 5 And chuc > 0 Then
                    str = "l¨m" & " " & str
                ElseIf donvi > 0 Then
                    str = so(donvi) & " " & str
                End If
            End If
            If chuc < 0 Then
                Exit Do
            Else
                If chuc = 0 And donvi > 0 Then
                    str = "lÎ" & " " & str
                ElseIf chuc = 1 Then
                    str = "m­êi" & " " & str
                ElseIf chuc > 1 Then
                    str = so(chuc) & " " & "m­¬i" & " " & str
                End If
            End If
            If tram < 0 Then
                Exit Do
            Else
                If tram > 0 Or chuc > 0 Or donvi > 0 Then
                    str = so(tram) & " " & "tr¨m" & " " & str
                End If
            End If
        Loop
    End If
    str = Trim(str)
    If str <> "" Then
     str = UCase(Left(str, 1)) & Right(str, Len(str) - 1)
    
     readnumber = str
     End If
End Function


Public Function LoadLanguage(ByVal LngFile As String, ByVal DescPos As String)
On Error GoTo errHdl

    Dim LngArray() As String
    Dim hFile As Integer
    Dim tmpStr As String
    Dim lFound As Boolean
    
    lFound = False
    hFile = FreeFile
    Open LngFile For Input As #hFile
    Do While Not EOF(hFile)
        DoEvents
        Line Input #hFile, tmpStr
        If Left(tmpStr, Len(DescPos)) = DescPos Then
            lFound = True
            tmpStr = Right(tmpStr, Len(tmpStr) - Len(DescPos))
                ReDim Preserve LngArray(CLng(Left(tmpStr, InStr(tmpStr, ":") - 1)))
            LngArray(CLng(Left(tmpStr, InStr(tmpStr, ":") - 1))) = Right(tmpStr, Len(tmpStr) - InStr(tmpStr, ":"))
        Else
            If lFound And Left(tmpStr, 4) = "#000" Then Exit Do
        End If
    Loop
    Close #hFile
    LoadLanguage = LngArray
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlLanguage - LoadLanguage"
End Function

Public Sub CheckSitePath(ByVal sgpFile As String)
On Error GoTo errHdl

    If Dir(RemoveExtFile(SiteFile), vbDirectory) = "" Then MkDir RemoveExtFile(SiteFile)
    WorkingFolder = RemoveExtFile(SiteFile)
   
Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlFile - CheckSitePath"
End Sub

Public Function RemoveExtFile(ByVal FileName As String) As String
On Error GoTo errHdl

    Dim fso As New FileSystemObject
    Dim tmpStr As String
    
    If FileName = "" Then
        RemoveExtFile = ""
        Exit Function
    End If
    tmpStr = fso.GetExtensionName(FileName)
    RemoveExtFile = Replace(FileName, "." & tmpStr, "")

Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "PublicFunction - RemoveExtFile"
End Function
Public Sub LoadFont(ByRef strFont As String, ByVal fLang As String)
On Error GoTo errHdl

    Dim fL As Integer
    Dim tmpStr As String
    
    fL = FreeFile
    Open fLang For Input As #fL
    Do While Not EOF(fL)
        DoEvents
        Line Input #fL, tmpStr
        If Left(tmpStr, 11) = "#99:001:02:" Then
            strFont = Right(tmpStr, Len(tmpStr) - 11)
            Exit Do
        End If
    Loop
    Close #fL
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlLanguage - LoadFont"
End Sub

'*********************************************************
'Chuc nang  : xoa cac file tmp
'Tham so vao: khong
'Tham so ra : khong
'Nguoi tao  : Hai-09/10/06
'Nguoi sua  :
'*********************************************************
Public Sub gsDELETE_TMP_FILE()
On Error GoTo errHdl

    Dim strPath   As String
    'xoa o dia goc
   
    strPath = Left(App.Path, 3) & "*.tmp"
    
    If Dir(strPath) <> "" Then
        Kill strPath
    End If
    
    'xoa thu muc chua file chay exe
    strPath = App.Path & "\*.tmp"
    If Dir(strPath) <> "" Then
        Kill strPath
    End If
    ' Xóa thu muc chua data
    strPath = Left(WorkingFolder, 2) & "\*.tmp"
    If Dir(strPath) <> "" Then
        Kill strPath
    End If
    ' Xóa thu muc chua data
    strPath = "C:\*.tmp"
    If Dir(strPath) <> "" Then
        Kill strPath
    End If
Exit Sub
errHdl:
    MsgBox Err.Number & Err.Description & "Delete temp file not completed!"
End Sub

Public Function gfCONVERT_DATE_TO_STRING(ByVal pDatIn As Date) _
                    As String
On Error GoTo errHdl
    Dim strRet As String
    
    gfCONVERT_DATE_TO_STRING = ""
    strRet = Year(pDatIn)
    strRet = strRet & Format(Month(pDatIn), "00")
    strRet = strRet & Format(Day(pDatIn), "00")
    gfCONVERT_DATE_TO_STRING = strRet
    
Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gfCONVERT_DATE_TO_STRING"
End Function
'*********************************************************
'Chuc nang  : chuyen kieu chuoi dang yyyymmdd thanh
'           kieu chuoi dang dd/mm/yyyy
'Tham so vao: pStrDateIn: chuoi ngay can chuyen
'Tham so ra : chuoi ngay dd/mm/yyyy
'Nguoi tao  : Can-25/10/06
'Nguoi sua  :
'*********************************************************
Public Function gfCONVERT_STRING_TO_DATE(ByVal pStrDateIn _
                    As String) As String
On Error GoTo errHdl
    Dim strDate As String
    
    strDate = Right(pStrDateIn, 2) & "/" & _
                 Mid(pStrDateIn, 5, 2) & _
                "/" & Left(pStrDateIn, 4)
    
    gfCONVERT_STRING_TO_DATE = strDate
    
Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGenaralStock - gfCONVERT_STRING_TO_DATE"
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Ham tra ve gia tri lon nhat cua cot fieldName co trong ban TableName
'Tham so vao: Ten ban, Ten cot
'Tham so ra : Gia tri lon nhat cua cot trong bang

Public Function GetMax_ID(TableName As String, FieldName As String) As String
On Error GoTo Handle
Dim result As String
Dim str As String
Dim rsmax As New ADODB.Recordset
    str = "Select max(" & FieldName & ") from " & TableName & """"
    Set rsmax = OpenCriticalTable(str, cnData)
    If rsmax.RecordCount > 0 And Not rsmax.EOF Then
        result = Right("00" & rsmax.Fields(0) + 1, 2)
    Else
        result = Right("00" & "1", 2)
    End If
GetMax_ID = result
Exit Function
Handle:
MsgBox Err.Number & Err.Description & " GetMax_ID"

End Function


'Public Function Get_Price_Kar(Dateprice As Date) As Double
'On Error GoTo Handle
'Dim Kar_Price As Double
'Dim i As Integer
'Dim iweek1, iweek2 As Integer
'Dim Arr() As String
'Dim rsKar As New ADODB.Recordset
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    Set rsKar = Open_Table(cnData, "Setup_Karaoke")
'    For i = 1 To 3
'        With rsKar
'            .Find "ID='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
'            If Not .EOF Then
'                Select Case i
'                    Case 1
'                        Select Case .Fields("Weekday_from")
'                            Case "Monday": iweek1 = 1
'                            Case "Tuesday": iweek1 = 2
'                            Case "Wednesday": iweek1 = 3
'                            Case "Thursday": iweek1 = 4
'                            Case "Friday": iweek1 = 5
'                            Case "Saturday": iweek1 = 6
'                            Case "Sunday": iweek1 = 7
'                        End Select
'                        Select Case .Fields("Weekday_To")
'                            Case "Monday": iweek2 = 1
'                            Case "Tuesday": iweek2 = 2
'                            Case "Wednesday": iweek2 = 3
'                            Case "Thursday": iweek2 = 4
'                            Case "Friday": iweek2 = 5
'                            Case "Saturday": iweek2 = 6
'                            Case "Sunday": iweek2 = 7
'                        End Select
'                        If Weekday(Dateprice, vbMonday) >= iweek1 And Weekday(Dateprice, vbMonday) <= iweek2 Then
'                            If Format(Now, "HH:mm:ss") >= Format(.Fields("From_Time1"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(.Fields("To_Time1"), "HH:mm:ss") Then
'                                Kar_Price = .Fields("Price1")
'                            ElseIf Format(Now, "HH:mm:ss") >= Format(.Fields("From_Time2"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(.Fields("To_Time2"), "HH:mm:ss") Then
'                                Kar_Price = .Fields("Price2")
'                            ElseIf Format(Now, "HH:mm:ss") >= Format(.Fields("From_Time3"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(.Fields("To_Time3"), "HH:mm:ss") Then
'                                Kar_Price = .Fields("Price3")
'                            End If
'                        End If
'                    Case 2
'                        Select Case .Fields("Weekday_from")
'                            Case "Monday": iweek1 = 1
'                            Case "Tuesday": iweek1 = 2
'                            Case "Wednesday": iweek1 = 3
'                            Case "Thursday": iweek1 = 4
'                            Case "Friday": iweek1 = 5
'                            Case "Saturday": iweek1 = 6
'                            Case "Sunday": iweek1 = 7
'                        End Select
'                        Select Case .Fields("Weekday_To")
'                            Case "Monday": iweek2 = 1
'                            Case "Tuesday": iweek2 = 2
'                            Case "Wednesday": iweek2 = 3
'                            Case "Thursday": iweek2 = 4
'                            Case "Friday": iweek2 = 5
'                            Case "Saturday": iweek2 = 6
'                            Case "Sunday": iweek2 = 7
'                        End Select
'                        If Weekday(Dateprice, vbMonday) >= iweek1 And Weekday(Dateprice, vbMonday) <= iweek2 Then
'                            If Format(Now, "HH:mm:ss") >= Format(.Fields("From_Time1"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(.Fields("To_Time1"), "HH:mm:ss") Then
'                                Kar_Price = .Fields("Price1")
'                            ElseIf Format(Now, "HH:mm:ss") >= Format(.Fields("From_Time2"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(.Fields("To_Time2"), "HH:mm:ss") Then
'                                Kar_Price = .Fields("Price2")
'                            ElseIf Format(Now, "HH:mm:ss") >= Format(.Fields("From_Time3"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(.Fields("To_Time3"), "HH:mm:ss") Then
'                                Kar_Price = .Fields("Price3")
'                            End If
'                        End If
'                    Case 3
'                        Dim j As Integer
'                        Call AddArr(.Fields("Weekday_From"), Arr())
'                        For j = 0 To UBound(Arr()) - 1
'                            If Format(Day(Dateprice), "00") & "/" & Format(Month(Dateprice), "00") = Arr(j) Then
'                                If Format(Now, "HH:mm:ss") >= Format(.Fields("From_Time1"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(.Fields("To_Time1"), "HH:mm:ss") Then
'                                    Kar_Price = .Fields("Price1")
'                                ElseIf Format(Now, "HH:mm:ss") >= Format(.Fields("From_Time2"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(.Fields("To_Time2"), "HH:mm:ss") Then
'                                    Kar_Price = .Fields("Price2")
'                                ElseIf Format(Now, "HH:mm:ss") >= Format(.Fields("From_Time3"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(.Fields("To_Time3"), "HH:mm:ss") Then
'                                    Kar_Price = .Fields("Price3")
'                                End If
'                            End If
'                        Next
'                End Select
'            End If
'        End With
'    Next
'    Get_Price_Kar = Kar_Price
'Exit Function
'Handle:
'    Get_Price_Kar = 0
'    MsgBox Err.Number & Err.Description & "Get_Price_Kar"
'End Function

Public Function Get_Price_Kar(ByVal Dateprice As Date, sTime As String, Price_Level As Integer) As Double
On Error GoTo Handle
Dim Kar_Price As Double
Dim i As Integer
Dim iweek1, iweek2 As Integer
Dim Arr() As String
Dim rsKar As New ADODB.Recordset
    If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    Set rsKar = Open_Table(cnData, "Setup_Karaoke")
    For i = 1 To 3
        With rsKar
            .Find "ID='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                Select Case i
                    Case 1
                        Select Case .Fields("Weekday_from")
                            Case "Monday": iweek1 = 1
                            Case "Tuesday": iweek1 = 2
                            Case "Wednesday": iweek1 = 3
                            Case "Thursday": iweek1 = 4
                            Case "Friday": iweek1 = 5
                            Case "Saturday": iweek1 = 6
                            Case "Sunday": iweek1 = 7
                        End Select
                        Select Case .Fields("Weekday_To")
                            Case "Monday": iweek2 = 1
                            Case "Tuesday": iweek2 = 2
                            Case "Wednesday": iweek2 = 3
                            Case "Thursday": iweek2 = 4
                            Case "Friday": iweek2 = 5
                            Case "Saturday": iweek2 = 6
                            Case "Sunday": iweek2 = 7
                        End Select
                        If Weekday(Dateprice, vbMonday) >= iweek1 And Weekday(Dateprice, vbMonday) <= iweek2 Then
                            If Format(sTime, "HH:mm:ss") >= Format(.Fields("From_Time1"), "HH:mm:ss") And Format(sTime, "HH:mm:ss") <= Format(.Fields("To_Time1"), "HH:mm:ss") Then
                                Select Case Price_Level
                                    Case 1: Kar_Price = .Fields("Price1")
                                    Case 2: Kar_Price = .Fields("Price11")
                                    Case 3: Kar_Price = .Fields("Price12")
                                    Case 4: Kar_Price = .Fields("Price13")
                                End Select
                            ElseIf Format(sTime, "HH:mm:ss") >= Format(.Fields("From_Time2"), "HH:mm:ss") And Format(sTime, "HH:mm:ss") <= Format(.Fields("To_Time2"), "HH:mm:ss") Then
                                
                                Select Case Price_Level
                                    Case 1: Kar_Price = .Fields("Price2")
                                    Case 2: Kar_Price = .Fields("Price21")
                                    Case 3: Kar_Price = .Fields("Price22")
                                    Case 4: Kar_Price = .Fields("Price23")
                                End Select
                                
                            ElseIf Format(sTime, "HH:mm:ss") >= Format(.Fields("From_Time3"), "HH:mm:ss") And Format(sTime, "HH:mm:ss") <= Format(.Fields("To_Time3"), "HH:mm:ss") Then
                                
                                Select Case Price_Level
                                    Case 1: Kar_Price = .Fields("Price3")
                                    Case 2: Kar_Price = .Fields("Price31")
                                    Case 3: Kar_Price = .Fields("Price32")
                                    Case 4: Kar_Price = .Fields("Price33")
                                End Select
                                
                            End If
                        End If
                    Case 2
                        Select Case .Fields("Weekday_from")
                            Case "Monday": iweek1 = 1
                            Case "Tuesday": iweek1 = 2
                            Case "Wednesday": iweek1 = 3
                            Case "Thursday": iweek1 = 4
                            Case "Friday": iweek1 = 5
                            Case "Saturday": iweek1 = 6
                            Case "Sunday": iweek1 = 7
                        End Select
                        Select Case .Fields("Weekday_To")
                            Case "Monday": iweek2 = 1
                            Case "Tuesday": iweek2 = 2
                            Case "Wednesday": iweek2 = 3
                            Case "Thursday": iweek2 = 4
                            Case "Friday": iweek2 = 5
                            Case "Saturday": iweek2 = 6
                            Case "Sunday": iweek2 = 7
                        End Select
                        If Weekday(Dateprice, vbMonday) >= iweek1 And Weekday(Dateprice, vbMonday) <= iweek2 Then
                            If Format(sTime, "HH:mm:ss") >= Format(.Fields("From_Time1"), "HH:mm:ss") And Format(sTime, "HH:mm:ss") <= Format(.Fields("To_Time1"), "HH:mm:ss") Then
                                Select Case Price_Level
                                    Case 1: Kar_Price = .Fields("Price1")
                                    Case 2: Kar_Price = .Fields("Price11")
                                    Case 3: Kar_Price = .Fields("Price12")
                                    Case 4: Kar_Price = .Fields("Price13")
                                End Select
                            ElseIf Format(sTime, "HH:mm:ss") >= Format(.Fields("From_Time2"), "HH:mm:ss") And Format(sTime, "HH:mm:ss") <= Format(.Fields("To_Time2"), "HH:mm:ss") Then
                               Select Case Price_Level
                                    Case 1: Kar_Price = .Fields("Price2")
                                    Case 2: Kar_Price = .Fields("Price21")
                                    Case 3: Kar_Price = .Fields("Price22")
                                    Case 4: Kar_Price = .Fields("Price23")
                                End Select
                            ElseIf Format(sTime, "HH:mm:ss") >= Format(.Fields("From_Time3"), "HH:mm:ss") And Format(sTime, "HH:mm:ss") <= Format(.Fields("To_Time3"), "HH:mm:ss") Then
                               Select Case Price_Level
                                    Case 1: Kar_Price = .Fields("Price3")
                                    Case 2: Kar_Price = .Fields("Price31")
                                    Case 3: Kar_Price = .Fields("Price32")
                                    Case 4: Kar_Price = .Fields("Price33")
                                End Select
                            End If
                        End If
                    Case 3
                        Dim j As Integer
                        Call AddArr(.Fields("Weekday_From"), Arr())
                        For j = 0 To UBound(Arr()) '- 1
                            If Format(Day(Dateprice), "00") & "/" & Format(Month(Dateprice), "00") = Arr(j) Then
                                If Format(sTime, "HH:mm:ss") >= Format(.Fields("From_Time1"), "HH:mm:ss") And Format(sTime, "HH:mm:ss") <= Format(.Fields("To_Time1"), "HH:mm:ss") Then
                                    Select Case Price_Level
                                        Case 1: Kar_Price = .Fields("Price1")
                                        Case 2: Kar_Price = .Fields("Price11")
                                        Case 3: Kar_Price = .Fields("Price12")
                                        Case 4: Kar_Price = .Fields("Price13")
                                    End Select
                                ElseIf Format(sTime, "HH:mm:ss") >= Format(.Fields("From_Time2"), "HH:mm:ss") And Format(sTime, "HH:mm:ss") <= Format(.Fields("To_Time2"), "HH:mm:ss") Then
                                    Select Case Price_Level
                                        Case 1: Kar_Price = .Fields("Price2")
                                        Case 2: Kar_Price = .Fields("Price21")
                                        Case 3: Kar_Price = .Fields("Price22")
                                        Case 4: Kar_Price = .Fields("Price23")
                                    End Select
                                ElseIf Format(sTime, "HH:mm:ss") >= Format(.Fields("From_Time3"), "HH:mm:ss") And Format(sTime, "HH:mm:ss") <= Format(.Fields("To_Time3"), "HH:mm:ss") Then
                                   Select Case Price_Level
                                        Case 1: Kar_Price = .Fields("Price3")
                                        Case 2: Kar_Price = .Fields("Price31")
                                        Case 3: Kar_Price = .Fields("Price32")
                                        Case 4: Kar_Price = .Fields("Price33")
                                    End Select
                                End If
                            End If
                        Next
                End Select
            End If
        End With
    Next
    Get_Price_Kar = Kar_Price
Exit Function
Handle:
    Get_Price_Kar = 0
    MsgBox Err.Number & Err.Description & "Get_Price_Kar"
End Function
Public Sub AddArr(str As String, Arr() As String)
Dim plash As Integer
Dim count As Integer
Dim chuoi, tmpStr As String
    chuoi = str
    count = 0
    Do While Len(chuoi) > 0
        plash = InStr(1, chuoi, ";", 0)
        ReDim Preserve Arr(count)
        If plash <> 0 Then
            tmpStr = Mid(chuoi, 1, plash - 1)
            Arr(count) = tmpStr
            If Len(chuoi) - Len(tmpStr) - 1 > 0 Then
                chuoi = Mid(chuoi, plash + 1, Len(chuoi) - Len(tmpStr) - 1)
            Else
                chuoi = ""
            End If
        Else
            Arr(count) = tmpStr
            Exit Do
        End If
        count = count + 1
    Loop
        
End Sub

Public Function RightDeCode(S1 As String) As String
    Dim sResult As String
    Dim i As Integer
    
    sResult = ""
    If S1 = "" Then GoTo 1
    For i = 1 To Len(S1) Step 2
    DoEvents
        If Mid(S1, i, 2) <> "-1" And Mid(S1, 1, 2) <> "  " Then
            sResult = sResult & FillZeroForString(HexToBin(Mid(S1, i, 2)), 8)
        End If
    Next i
1:  RightDeCode = sResult
End Function

