Attribute VB_Name = "mdlPassword"
Option Explicit


Private Function GetBCC(str$) As Integer
On Error GoTo errHdl

    Dim i As Integer, m As Integer
        m = Asc(Mid$(str$, 1, 1))
    For i = 2 To Len(str$) Step 1
        m = m Xor Asc(Mid$(str$, i, 1))
    Next
    GetBCC = m
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlPassword - GetBCC"
End Function



Public Function LoadPasswordData() As ADODB.Recordset
On Error GoTo errHdl

    Dim rs As New ADODB.Recordset
    Dim rsUser_Data As New ADODB.Recordset
    If cnData.State = adStateClosed Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    With rs
        .Fields.Append "ID", adVarWChar, 10
        .Fields.Append "UserName", adVarWChar, 50
        .Fields.Append "UserLevel", adVarWChar, 1
        .Fields.Append "Password", adVarWChar, 32
        .Fields.Append "UserRight", adVarWChar, 640
        .Fields("UserRight").Attributes = adColNullable
        .Open
    End With
   Set rsUser_Data = Open_Table(cnData, "User_Login")
   With rsUser_Data
        Do While Not .EOF
            With rs
                If Not .EOF And .RecordCount > 0 Then .MoveLast
                .addNew
                !ID = Trim(rsUser_Data.Fields("ID"))
                !userName = rsUser_Data.Fields("UserName")
                !UserLevel = rsUser_Data.Fields("UserLevel")
                !Password = En_Decryption.MalgoDecrypt(Trim(rsUser_Data.Fields("Password")), 10)
                !UserRight = rsUser_Data.Fields("UserRight")
                .Update
            End With
        .MoveNext
        Loop
        If rs.RecordCount > 0 Then rs.MoveFirst
    End With
    
    Set LoadPasswordData = rs

    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlPassword - LoadPasswordData"
End Function
Public Function SavePasswordData(rsTemp As ADODB.Recordset)
On Error GoTo Handle
Dim rsRight As New ADODB.Recordset
cnData.Execute "delete from user_Login"
Set rsRight = Open_Table(cnData, "User_Login")
    With rsRight
            With rsTemp
            If .RecordCount > 0 Then .MoveFirst
                Do While Not .EOF
                        rsRight.addNew
                        rsRight.Fields("ID") = .Fields("ID")
                        rsRight.Fields("UserName") = .Fields("UserName")
                        rsRight.Fields("UserLevel") = .Fields("UserLevel")
                        rsRight.Fields("Password") = En_Decryption.MalgoEncrypt(.Fields("Password"), 10)
                        rsRight.Fields("UserRight") = .Fields("UserRight")
                        rsRight.Update
                        rsRight.Requery
                .MoveNext
                Loop
            End With
    End With
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & "  SavePasswordData"

End Function


Private Function FillInteger(InputInt As Integer) As String
On Error GoTo errHdl

    Dim tmpInt As String
    tmpInt = CStr(InputInt)
    If Len(tmpInt) < 3 Then
        Do While Len(tmpInt) < 3
            tmpInt = "0" & tmpInt
        Loop
    Else
        tmpInt = Right(tmpInt, 3)
    End If
    FillInteger = tmpInt
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlPassword - FillInteger"
End Function

