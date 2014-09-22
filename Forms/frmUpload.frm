VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmUpload 
   Caption         =   "Import Danh môc "
   ClientHeight    =   2295
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8475
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
   Icon            =   "frmUpload.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2295
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyProgressBar MyProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   661
   End
   Begin prjTouchScreen.MyButton cmdUpload 
      Height          =   735
      Left            =   5760
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "IMPORT"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpload.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdBrowse 
      Height          =   495
      Left            =   6600
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      BTYPE           =   5
      TX              =   "Browse..."
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpload.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox tb_upload 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   6375
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Height          =   735
      Left            =   7080
      TabIndex        =   4
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "§ãng"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   14737632
      BCOLO           =   14737632
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpload.frx":0044
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
      Alignment       =   2  'Center
      Caption         =   "Import dANH MôC MENU"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   5055
   End
End
Attribute VB_Name = "frmUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim Call_Value As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdUpload_Click()
On Error Resume Next
Dim ex As New Excel.Application
Dim wb As Excel.Workbook
Dim ws As Excel.Worksheet
Dim i As Long: i = 1
Dim sqlUpload As String

' If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
Set wb = ex.Workbooks.Open(tb_upload.Text, , True)
Set ws = wb.Worksheets(1)
Do: i = i + 1
If MyProgressBar1.Value < MyProgressBar1.Max Then
    MyProgressBar1.Value = MyProgressBar1.Value + 1
Else
    MyProgressBar1.Value = 0
End If
If Call_Value = "Items" Then
    cnData.Execute "INSERT INTO Inventory (ItemNum,ItemName,Dept_ID,Std_Price1,Std_Price2,Std_Price3, " & _
                        " HH_Price1,HH_Price2,HH_Price3,EV_Price1,EV_Price2,EV_Price3,LimitPrice,Unit,Minstock," & _
                        " Modify_Number,F1,F2,F3,F4,F5,Date_Created,Picture,Print_On_Receipt,Store_ID)" & _
                        " VALUES ('" & ws.Range("A" & i).Value & "', '" & ws.Range("B" & i).Value & "','" & ws.Range("C" & i).Value & "','" & _
                        ws.Range("D" & i).Value & "','" & ws.Range("E" & i).Value & "','" & ws.Range("F" & i).Value & "','" & _
                        ws.Range("G" & i).Value & "','" & ws.Range("H" & i).Value & "','" & ws.Range("I" & i).Value & "','" & _
                        ws.Range("J" & i).Value & "','" & ws.Range("K" & i).Value & "','" & ws.Range("L" & i).Value & "','" & _
                        ws.Range("M" & i).Value & "','" & ws.Range("N" & i).Value & "','" & ws.Range("O" & i).Value & "','" & _
                        ws.Range("P" & i).Value & "','" & ws.Range("Q" & i).Value & "','" & ws.Range("R" & i).Value & "','" & _
                        ws.Range("S" & i).Value & "','" & ws.Range("T" & i).Value & "','" & ws.Range("U" & i).Value & "','" & _
                        ws.Range("V" & i).Value & "','" & ws.Range("W" & i).Value & "','" & ws.Range("X" & i).Value & "','" & _
                        ws.Range("Y" & i).Value & "')"
ElseIf Call_Value = "Department" Then

End If
Loop Until ws.Range("A" & (i + 1)) = Empty
ex.Quit
MsgBox "Upload hoµn tÊt"
MyProgressBar1.Value = MyProgressBar1.Max
Set ex = Nothing
'Handle:
'MsgBox Err.Description
End Sub

Private Sub cmdBrowse_Click()
CommonDialog1.Filter = "(Microsoft Excel Workbook)|*.xls*"
CommonDialog1.ShowOpen
tb_upload.Text = CommonDialog1.FileName
End Sub


Public Property Let FormCall(ByVal vNewValue As Variant)
    Call_Value = vNewValue
End Property
