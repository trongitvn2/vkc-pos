VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSplitBill 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Split Bill"
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11085
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid flgSubBill 
      Height          =   3885
      Index           =   0
      Left            =   3720
      TabIndex        =   6
      Top             =   960
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   6853
      _Version        =   393216
      Cols            =   5
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      GridLinesFixed  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjTouchScreen.MyButton cmdSelectAll 
      Height          =   435
      Left            =   30
      TabIndex        =   3
      Top             =   8760
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   767
      BTYPE           =   6
      TX              =   "Select All"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16578804
      BCOLO           =   16578804
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSplitBill.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Frame frmButton 
      Height          =   1545
      Left            =   30
      TabIndex        =   2
      Top             =   9600
      Width           =   15225
      Begin prjTouchScreen.MyButton cmdCombine 
         Height          =   915
         Left            =   2460
         TabIndex        =   7
         Top             =   420
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   1614
         BTYPE           =   14
         TX              =   "&Gép chung"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16578804
         BCOLO           =   16777152
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSplitBill.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdNumBill 
         Height          =   945
         Left            =   90
         TabIndex        =   5
         Top             =   420
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   1667
         BTYPE           =   14
         TX              =   "&Chia ®Òu ra"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16578804
         BCOLO           =   16777152
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSplitBill.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdDone 
         Height          =   915
         Left            =   13020
         TabIndex        =   9
         Top             =   390
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   1614
         BTYPE           =   14
         TX              =   "&Hoµn tÊt"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16578804
         BCOLO           =   16777152
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSplitBill.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flgBill 
      Height          =   7065
      Left            =   30
      TabIndex        =   0
      Top             =   960
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   12462
      _Version        =   393216
      Cols            =   5
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      GridColor       =   8421504
      GridColorFixed  =   8421504
      GridLinesFixed  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjTouchScreen.MyButton cmdDeselect 
      Height          =   435
      Left            =   1920
      TabIndex        =   4
      Top             =   8760
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   767
      BTYPE           =   6
      TX              =   "Deselect"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16578804
      BCOLO           =   16578804
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSplitBill.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   615
      Left            =   30
      TabIndex        =   12
      Top             =   0
      Width           =   15255
      ForeColor       =   16711680
      BackColor       =   14737632
      Caption         =   "SPLIT BILL"
      Size            =   "26908;1085"
      FontName        =   ".VnArialH"
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   34
      ParagraphAlign  =   3
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tæng céng:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   60
      TabIndex        =   11
      Top             =   8100
      Width           =   1725
   End
   Begin VB.Label lblTotalAmt 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   1800
      TabIndex        =   10
      Top             =   8100
      Width           =   1875
   End
   Begin VB.Label lblSubBill 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bill 1:"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   8
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label lblOrgBill 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Bill gèc:"
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   30
      TabIndex        =   1
      Top             =   600
      Width           =   3645
   End
End
Attribute VB_Name = "frmSplitBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsOrgBill As New ADODB.Recordset
Dim TotalAmt As Double
Dim Desarr() As String
Dim rsSubCheck As New ADODB.Recordset
Dim iClickSub As Boolean
Dim cnSplit_Bill As New ADODB.Connection
Dim Bill_Org As Integer
Dim rsInvoice_subCheck As New ADODB.Recordset
Dim rsInvoice_subCheck_Items As New ADODB.Recordset
''

Private Sub cmdDone_Click()
    Unload Me
    With frmSplitBillDone
        .Get_Master_Bill = Get_Bill_Number
        .Show vbModal
    End With
End Sub

Private Sub flgBill_Click()
On Error Resume Next
        With rsSubCheck
            If .State = 0 Then
                .Fields.Append "Line_Number", adDouble
                .Fields.Append "PLUNo", adVarWChar, 20
                .Fields.Append "PLUName", adVarWChar, 50
                .Fields.Append "Qty", adDouble
                .Fields.Append "Std_Price1", adDouble
                .Fields.Append "Amt", adDouble
                .Open
           End If
           If iClickSub = False Then
                .addNew
                rsOrgBill.Find "Line_Number=" & flgBill.TextMatrix(flgBill.Row, 5), , adSearchForward, adBookmarkFirst
                
                If Not rsOrgBill.EOF Then
                    .Fields("PluNo") = rsOrgBill!PluNo
                    .Fields("PluName") = rsOrgBill!PluName
                    If rsOrgBill!Qty > 1 Then
                        frmQtyTranfer.Show vbModal
                        .Fields("Qty") = frmQtyTranfer.Let_Result
                        rsOrgBill!Qty = rsOrgBill!Qty - .Fields("Qty")
                        rsOrgBill.Update
                        If rsOrgBill!Qty = 0 Then rsOrgBill.Delete adAffectCurrent
                    Else
                    
                        .Fields("Qty") = rsOrgBill!Qty
                    End If
                    .Fields("Std_Price1") = rsOrgBill!Std_Price1
                    .Fields("Amt") = rsOrgBill!Amt
                    .Fields("Line_number") = rsOrgBill!Line_Number
                    .Update
                End If
            Else
                rsOrgBill.addNew
                rsOrgBill.Fields("PluNo") = .Fields("PluNo")
                rsOrgBill.Fields("PluName") = .Fields("PluName")
                rsOrgBill.Fields("Qty") = .Fields("Qty")
                rsOrgBill.Fields("Std_Price1") = .Fields("Std_Price1")
                rsOrgBill.Fields("Amt") = .Fields("Amt")
                rsOrgBill.Fields("Line_number") = .Fields("Line_number")
                .Update
                
            End If
            
        End With
       ' Call SetFLGRIDORDER(rsOrgBill)
End Sub

Private Sub flgSubBill_Click(Index As Integer)
On Error Resume Next
    iClickSub = True
    If rsSubCheck.RecordCount > 0 Then
        Call SetFLGRIDORDER(rsSubCheck, flgSubBill(Index))
'        With rsSubCheck
'            .MoveFirst
'            Do While Not .EOF
'                CallDeleteRecords (.Fields("Line_Number"))
'            .MoveNext
'            Loop
'        End With
        Call SetFLGRIDORDER(rsOrgBill, flgBill)
        
        Set rsSubCheck = Nothing
        Load flgSubBill(Index + 1)
        With flgSubBill(Index + 1)
        If (Index + 1) Mod 3 = 0 Then
            .top = flgSubBill(Index).top + flgSubBill(Index).Height + lblSubBill(Index).Height
            .Left = flgSubBill(0).Left
        Else
            .top = flgSubBill(Index).top
            .Left = flgSubBill(Index).Left + 50 + flgSubBill(Index).Width
        End If
            .Visible = True
        End With
        Load lblSubBill(Index + 1)
        With lblSubBill(Index + 1)
            .top = flgSubBill(Index + 1).top - 300
            .Left = flgSubBill(Index + 1).Left
            .Caption = "Bill" & " " & Index + 2 & ":"
            .Visible = True
        End With
        Call Set_flgOrder(flgSubBill(Index + 1))
    Else
    
    End If
    iClickSub = False
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    Desarr = LoadLanguage(LngFile, "#01:007:")
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Desarr = LoadLanguage(LngFile, "#01:007:")
    
    If Dir(WorkingFolder & "\Database.mdb", vbDirectory) <> "" Then Set cnSplit_Bill = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    If cnData.State <> 0 Then
        Set rsInvoice_subCheck = Open_Table(cnData, "Invoice_SubCheck")
        Set rsInvoice_subCheck_Items = Open_Table(cnData, "Invoice_SubCheck_Items")
    End If
    
    Call Set_flgOrder(flgBill)
    Call Set_flgOrder(flgSubBill(0))
    Call SetFLGRIDORDER(rsOrgBill, flgBill)
    If rsOrgBill.State = 1 Then
        rsOrgBill.MoveFirst
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub
Public Sub Set_flgOrder(flg As MSFlexGrid)
    On Error GoTo Handle
    Dim i As Integer
        With flg
            .Cols = 6
            .Rows = 30
            .ColWidth(0) = 0
            .ColWidth(1) = 1600
            .ColWidth(2) = 350
            .ColWidth(3) = 700
            .ColWidth(4) = 700
            .ColWidth(5) = 0
            .TextMatrix(0, 1) = Desarr(19) '"Tên món"
            .TextMatrix(0, 2) = Desarr(20) ' "Sô' luong"
            .TextMatrix(0, 3) = Desarr(21) '" D/Giá"
            .TextMatrix(0, 4) = Desarr(22) '"T/Tiên`"
            .ColAlignment(1) = 2
            .ColAlignment(2) = 4
            .ColAlignment(3) = 6
            .ColAlignment(4) = 6
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
            .TextMatrix(1, 4) = ""
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"

End Sub

Public Sub SetFLGRIDORDER(rs As ADODB.Recordset, flg As MSFlexGrid)
On Error GoTo Handle
        Dim incount As Integer
        rs.MoveFirst
        Do While Not rs.EOF
            incount = incount + 1
            flg.Rows = rs.RecordCount + 1
            With flg
                .TextMatrix(incount, 1) = rs!PluName
                .TextMatrix(incount, 2) = rs!Qty
                .TextMatrix(incount, 3) = Format(rs!Std_Price1, formatNum)
                .TextMatrix(incount, 4) = Format(rs!Amt, formatNum)
                .TextMatrix(incount, 5) = rs!Line_Number
            End With
            TotalAmt = TotalAmt + rs!Amt
        rs.MoveNext
        Loop
    If discount > 0 Then
        lblTotalAmt.Caption = Format(TotalAmt - TotalAmt * discount / 100, formatNum)
    Else
    
        lblTotalAmt.Caption = Format(TotalAmt, formatNum)
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "SetFLGRIDORDER"
End Sub

Public Property Let GetRecordset(ByVal vNewValue As Variant)
    Set rsOrgBill = vNewValue
End Property


Private Sub Form_Unload(Cancel As Integer)
    TotalAmt = 0
    Set rsSubCheck = Nothing
    Set rsOrgBill = Nothing
End Sub

Public Sub CallDeleteRecords(Line_Delete As Integer)
    On Error GoTo Handle
        With rsOrgBill
            .Find "Line_Number=" & Line_Delete, , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    .Delete adAffectCurrent
                End If
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " CallDeleteRecords"
End Sub

Public Property Let Get_Bill_Number(ByVal vNewValue As Variant)
    Bill_Org = vNewValue
End Property
Public Property Get Get_Bill_Number() As Variant
   Get_Bill_Number = Bill_Org
End Property

Public Function Update_Invoice_Subcheck() As Boolean
On Error GoTo Handle
    Update_Invoice_Subcheck = False
    With rsInvoice_subCheck
        .addNew
        .Fields("") = ""
        
    End With
    
Exit Function

Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Function
