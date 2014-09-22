Attribute VB_Name = "mdlPrivateProfile"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal SectionName As String, ByVal KeyName As String, ByVal Default As String, ByVal ReturnedString As String, ByVal StringSize As Long, ByVal FileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal SectionName As String, ByVal KeyName As String, ByVal KeyValue As String, ByVal FileName As String) As Long

Public Sub SaveSettingStr(ByVal Section As String, ByVal KeyName As String, ByVal Setting As String, IniFile As String)
On Error GoTo errHdl

    Dim lRet As Long
    lRet = WritePrivateProfileString(Section, KeyName, Setting, IniFile)
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlPrivateProfile - SaveSettingStr"
End Sub

Public Function GetSettingStr(ByVal Section As String, _
    ByVal KeyName As String, ByVal DefaultValue As String, _
    IniFile As String) As String
On Error GoTo errHdl

    Dim lRet As Long
    Dim sBuf As String * 512
    lRet = GetPrivateProfileString(Section, KeyName, _
        DefaultValue, sBuf, Len(sBuf), IniFile)
    GetSettingStr = TrimNull(sBuf)
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlPrivateProfile - GetSettingStr"
End Function

Private Function TrimNull(ByVal InString As String) As String
On Error GoTo errHdl

    Dim lPos As Long
    TrimNull = Trim$(InString)
    lPos = InStr(TrimNull, vbNullChar)
    If lPos > 0 Then TrimNull = Left(TrimNull, lPos - 1)
    
    Exit Function
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlPrivateProfile - TrimNull"
End Function


