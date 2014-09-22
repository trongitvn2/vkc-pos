VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmDetailsOrder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Chi tiÕt order"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12855
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
   Icon            =   "frmDetailsOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8640
   ScaleWidth      =   12855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSFlexGridLib.MSFlexGrid flgOrder 
      Height          =   7815
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   13785
      _Version        =   393216
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Cancel          =   -1  'True
      Height          =   735
      Left            =   11280
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDetailsOrder.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblBill 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Hãa ®¬n:"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lbltable 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Bµn:"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmDetailsOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim Bill, Table As String
Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub SetFLGRIDORDER(rs As ADODB.Recordset)
On Error GoTo handle
        Dim incount As Integer
        
        With flgOrder
            .Cols = 9
            .Rows = 20
            .ColWidth(0) = 0
            .ColWidth(1) = 2800
            .ColWidth(2) = 800
            .ColWidth(3) = 1050
            .ColWidth(4) = 1250
            .ColWidth(5) = 2400
            .ColWidth(6) = 1900
            .ColWidth(7) = 1200
            .TextMatrix(0, 1) = "Tªn mãn"
            .TextMatrix(0, 2) = "S.L"
            .TextMatrix(0, 3) = " §.Gi¸"
            .TextMatrix(0, 4) = "T.TiÒn"
            .TextMatrix(0, 5) = "Ghi chó order"
            .TextMatrix(0, 6) = "M¸y in Order"
            .TextMatrix(0, 7) = "Thêi gian"
            .TextMatrix(0, 8) = "Gi¶m % mãn"
            .ColAlignment(1) = 2
            .ColAlignment(2) = 4
            .ColAlignment(3) = 6
            .ColAlignment(4) = 6
            .ColAlignment(5) = 2
        End With
        rs.MoveFirst
        With rs
            .Sort = "Line_Number DeSC"
            Do While Not .EOF
                incount = incount + 1
                flgOrder.Rows = rs.RecordCount + 1
                With flgOrder
                    .TextMatrix(incount, 0) = rs!Line_Number
                    .TextMatrix(incount, 1) = rs!PluName
                    .TextMatrix(incount, 2) = rs!Qty
                    .TextMatrix(incount, 3) = Format(rs!Std_Price1, formatNum)
                    .TextMatrix(incount, 4) = Format(rs!Amt, formatNum)
                    .TextMatrix(incount, 5) = rs.Fields("Kit_Desc")
                    If ArrayFlag(rs.Fields("F2"), 1) = 1 Then
                        .TextMatrix(incount, 6) = "M¸y in BÕp"
                    ElseIf ArrayFlag(rs.Fields("F2"), 2) = 1 Then
                        .TextMatrix(incount, 6) = "M¸y in Pha chÕ"
                    Else
                        .TextMatrix(incount, 6) = " "
                    End If
                    .TextMatrix(incount, 7) = rs.Fields("TimeOrder")
                    .TextMatrix(incount, 8) = rs.Fields("Line_Disc")
                End With
            rs.MoveNext
            Loop
        End With
        If rs.RecordCount = 0 Then
            With flgOrder
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
                .TextMatrix(1, 4) = ""
                .TextMatrix(1, 5) = ""
            End With
        End If
Exit Sub
handle:
    MsgBox Err.Number & Err.Description & Me.Name & "SetFLGRIDORDER"
End Sub

Public Property Let Get_Recordset(ByVal vNewValue As Variant)
    Set rs = vNewValue
End Property

Private Sub Form_Load()
On Error GoTo handle
    Call SetFLGRIDORDER(rs)
    lblTable.Caption = Table
    lblBill.Caption = Bill
Exit Sub
handle:
    MsgBox Err.Number & Err.Description & Me.Name & " Form_Load"
End Sub


Public Property Let LetBill(ByVal vNewValue As Variant)
    Bill = vNewValue
End Property


Public Property Let LetTable(ByVal vNewValue As Variant)
    Table = vNewValue
End Property
