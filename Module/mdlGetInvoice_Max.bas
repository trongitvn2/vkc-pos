Attribute VB_Name = "mdlGetInvoice_Max"
Public Function GetMaxInvoice_Number() As String
On Error GoTo Handle
    Dim MaxInvoice As String
    Dim str As String
    Dim rsMaxInvoice As New ADODB.Recordset
    str = "SELECT Max(Invoice_Totals.Invoice_Number) AS MaxInvoice_Number" & _
            " From Invoice_Totals "
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsMaxInvoice = OpenCriticalTable(str, cnData)
    If Not rsMaxInvoice.EOF Then
        MaxInvoice = CDbl("0" & rsMaxInvoice.Fields("MaxInvoice_Number")) + 1
    Else
        MaxInvoice = 1
    End If
    GetMaxInvoice_Number = MaxInvoice
Exit Function
Handle:
    MsgBox Err.ne & Err.Description & "  GetMaxInvoice_Number"
End Function


Public Function Get_MaxInvoice_Backup(cnBackup As ADODB.Connection, ByVal Date_Invoice As String) As String
On Error GoTo Handle
    Dim MaxInvoice As String
    Dim str As String
    Dim rsMaxInvoice As New ADODB.Recordset
    str = "SELECT Max(right(Invoice_Totals.Invoice_Number,4)) AS MaxInvoice_Number" & _
            " From Invoice_Totals " & _
            " where left(DateTime,8)='" & Date_Invoice & "'"
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsMaxInvoice = OpenCriticalTable(str, cnBackup)
    If Not rsMaxInvoice.EOF Then
        MaxInvoice = CDbl("0" & rsMaxInvoice.Fields("MaxInvoice_Number")) + 1
    Else
        MaxInvoice = 1
    End If
    Get_MaxInvoice_Backup = MaxInvoice
Exit Function
Handle:
    MsgBox Err.ne & Err.Description & "  Get_MaxInvoice_Backup"
End Function

Public Function Get_MinInvoice(cn As ADODB.Connection, ByVal Date_Invoice As String) As String
On Error GoTo Handle
    Dim MinInvoice As String
    Dim str As String
    Dim rsMinInvoice As New ADODB.Recordset
    str = "SELECT Min(right(Invoice_Totals.Invoice_Number,4)) AS MinInvoice_Number" & _
            " From Invoice_Totals " & _
            " where left(DateTime,8)='" & Date_Invoice & "'"
    Set rsMinInvoice = OpenCriticalTable(str, cn)
    If Not rsMinInvoice.EOF Then
        MinInvoice = CDbl("0" & rsMinInvoice.Fields("MinInvoice_Number")) + 1
    Else
        MinInvoice = 1
    End If
    Get_MinInvoice = MinInvoice
Exit Function
Handle:
    MsgBox Err.ne & Err.Description & "  - Get_MinInvoice"
End Function

