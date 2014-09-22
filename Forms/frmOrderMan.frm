VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmOrderMan 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Nh©n viªn phôc vô"
   ClientHeight    =   10995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10995
   ScaleWidth      =   10725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdExit 
      Height          =   1260
      Left            =   9105
      TabIndex        =   2
      Top             =   4335
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   2223
      BTYPE           =   6
      TX              =   "&Tho¸t"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   16711680
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOrderMan.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdOK 
      Height          =   1335
      Left            =   9105
      TabIndex        =   1
      Top             =   2895
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   2355
      BTYPE           =   6
      TX              =   "&OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   16711680
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmOrderMan.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid flgEmployee 
      Height          =   11025
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   19447
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      GridLinesFixed  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOrderMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmployee As New ADODB.Recordset
Dim rsWork_Shift As New ADODB.Recordset
Dim DescArr() As String
Dim rsTem As New ADODB.Recordset
Dim strEmploy_ID As String

Private Sub setflgEmployee(rs As ADODB.Recordset)
On Error GoTo errHdl
    Dim intCount    As Integer
    With flgEmployee
    .Cols = 3
    .Rows = 2
        .Font = ".vnArial"
        .ColWidth(0) = 1500
        .ColWidth(1) = 9000
        .ColWidth(2) = 4000

        .TextMatrix(0, 0) = DescArr(4)
        .TextMatrix(0, 1) = DescArr(6)
        .TextMatrix(0, 2) = DescArr(7)
        
    End With
    
    If rs Is Nothing Then Exit Sub
    If rs.State = 0 Then Exit Sub
    If rs.RecordCount > 0 Then rs.MoveFirst
    If rs.EOF And rs.BOF Then
        With flgEmployee
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
        End With
        Exit Sub
    End If
   flgEmployee.Rows = rs.RecordCount + 1
    intCount = 0
    Do While Not rs.EOF
        intCount = intCount + 1
        flgEmployee.TextMatrix(intCount, 0) = rs!Emp_ID
        flgEmployee.TextMatrix(intCount, 1) = rs!Emp_Name
        flgEmployee.TextMatrix(intCount, 2) = rs!Dept

        rs.MoveNext
    Loop
'    SetColorFlexGrid flgEmployee, 1, 1, flgEmployee.Cols
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - setflgEmployee "
End Sub

Private Sub cmdExit_Click()
    Set rsTem = Nothing
    Set rsWork_Shift = Nothing
    Set rsEmployee = Nothing
    Unload Me

End Sub

Private Sub cmdOK_Click()
    Call flgEmployee_DblClick
End Sub

Private Sub flgEmployee_DblClick()
On Error GoTo Handle
    strEmploy_ID = flgEmployee.TextMatrix(flgEmployee.Row, 0)
    
    Unload Me
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " flgEmployee_DblClick"
End Sub

Private Sub Form_Activate()
    Dim ctrl As Control
    If cmdOK.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    DescArr = LoadLanguage(LngFile, "#03:005:")
    Me.Caption = DescArr(1)
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl

End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    Dim strPath As String
    Dim str, str1 As String
    DescArr = LoadLanguage(LngFile, "#03:005:")
    str = "Select * from Employee"
    str1 = "select * from work_Shift" ' where Intime<='" & Time & "' and OutTime<='" & Time & "'"
    Set rsEmployee = OpenCriticalTable(str, cnData)
    Set rsWork_Shift = OpenCriticalTable(str1, cnData)
    With rsTem
    If .State = 0 Then
                .Fields.Append "Emp_ID", adVarWChar, 20
                .Fields.Append "Emp_Name", adVarWChar, 36
                .Fields.Append "Dept", adVarWChar, 20
                .Open
            End If
            Do While Not rsWork_Shift.EOF
                Set rsEmployee = OpenCriticalTable("select * from Employee where shift='" & rsWork_Shift.Fields("Shift_ID") & "'", cnData)
                Do While Not rsEmployee.EOF
                    .addNew
                    .Fields("Emp_ID") = rsEmployee.Fields("Cashier_ID")
                    .Fields("Emp_Name") = rsEmployee.Fields("EmpName")
                    .Fields("Dept") = rsEmployee.Fields("Dept_ID")
                    .Update
                rsEmployee.MoveNext
                Loop
            rsWork_Shift.MoveNext
            Loop
    End With
    Call setflgEmployee(rsTem)
    Exit Sub
Handle:
    MsgBox Err.Number & " " & _
    Err.Description & Me.name & "Form_Load"

End Sub



Private Sub Form_Unload(Cancel As Integer)
Set rsTem = Nothing
Set rsEmployee = Nothing
Set rsWork_Shift = Nothing
End Sub

Public Property Get Let_Emp() As Variant
    Let_Emp = strEmploy_ID
End Property

