VERSION 5.00
Begin VB.Form frmCashMedia 
   Caption         =   "Thanh to¸n ngo¹i tÖ"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10350
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdexit 
      Height          =   735
      Left            =   5280
      TabIndex        =   2
      Top             =   6120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "Th&o¸t"
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
      BCOL            =   14215660
      BCOLO           =   12648447
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCashMedia.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Frame fraMedia 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      Begin prjTouchScreen.MyButton cmdMedia 
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1508
         BTYPE           =   14
         TX              =   "Media"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   12648447
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCashMedia.frx":001C
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
   Begin prjTouchScreen.MyButton cmdChangeRate 
      Height          =   735
      Left            =   2760
      TabIndex        =   3
      Top             =   6120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "&Thay ®æi tû gi¸"
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
      BCOL            =   14215660
      BCOLO           =   12648447
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmCashMedia.frx":0038
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
Attribute VB_Name = "frmCashMedia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsMedia As New ADODB.Recordset
Dim DescArr() As String
Dim Total As Double
Dim Customer As String
Dim BillNO As String
Dim rsItemPayment As New ADODB.Recordset
Dim service_Charge, Discount As Integer
Dim Adjtotal1, Adjtotal2, Adjtotal3, Adjtotal4 As Double

Private Sub cmdChangeRate_Click()
    frmMedia.Show vbModal
End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set rsMedia = Nothing
End Sub

Private Sub cmdMedia_Click(Index As Integer)
    rsMedia.Find "MediaID='" & Right("00" & Index, 2) & "'", , adSearchForward, adBookmarkFirst
    If Not rsMedia.EOF Then
        With frmCash
            If CDbl("0" & rsMedia.Fields("FCRATE")) = 0 Then
                MsgBox "Tû gi¸ quy ®æi kh«ng thÓ b»ng 0"
                Exit Sub
            End If
            If Discount > 0 Then
                .GetTotals = Round((Total - Total * Discount / 100) / rsMedia.Fields("FCRATE"), 3)
            Else
                .GetTotals = Round(Total / rsMedia.Fields("FCRATE"), 3)
            End If
            .GetCustomer = Get_Cust
            .GetBillNo = GetBill_Number
'            .GetRecord = Get_Record_Payment
            .Get_Payment_Method = rsMedia.Fields("MEDIAID")
            .Show vbModal
        End With
    End If
    'Unload Me
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
If cmdExit.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
DescArr = LoadLanguage(LngFile, "#02:019:")
'Load caption Command
Me.Caption = DescArr(1)
cmdChangeRate.Caption = DescArr(2)
cmdExit.Caption = DescArr(3)

    Set rsMedia = Open_Table(cnData, "Media")
    Call LoadCommand(rsMedia, "MediaName")
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    If cnData.State = 0 Then
        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    End If
    Set rsMedia = Open_Table(cnData, "Media")
    Call LoadCommand(rsMedia, "MediaName")
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Public Sub LoadCommand(rs As ADODB.Recordset, strTenfield1 As String)
On Error Resume Next 'GoTo Handle
Dim Index As Integer
Dim sodong As Integer
Dim i, j As Integer
Index = 1
rs.MoveFirst
If rs.RecordCount Mod 5 > 0 Then
    sodong = rs.RecordCount / 5 + 1
Else
    sodong = rs.RecordCount / 5
End If
If rs.RecordCount > 0 Then
    For i = 1 To sodong
        For j = 1 To 5
                Load cmdMedia(Index)
                With cmdMedia(Index)
                    If i = 1 Then
                        If Index Mod 6 = 0 Then
                            .Left = fraMedia.Left + 300
                            .top = cmdMedia(Index - 1).top + cmdMedia(Index - 1).Height + 200
                        Else
                            .top = cmdMedia(Index - 1).top
                            If j = 1 Then
                                 .Left = fraMedia.Left + 300
                            Else
                                .Left = cmdMedia(Index - 1).Left + 300 + cmdMedia(Index - 1).Width
                            End If
                        End If
                    Else
                        If (Index - 1) Mod 5 = 0 Then
                            .Left = fraMedia.Left + 300
                            .top = cmdMedia(Index - 1).top + cmdMedia(Index - 1).Height + 200
                        Else
                            .top = cmdMedia(Index - 1).top
                            If j = 1 Then
                               .Left = fraMedia.Left + 300
                            Else
                                .Left = cmdMedia(Index - 1).Left + 300 + cmdMedia(Index - 1).Width
                            End If
                        End If
                    End If
                        If Not rs.EOF Then
                            .Caption = rs.Fields("" & strTenfield1 & "") '& Chr(13) & rs.Fields("" & strTenfield2 & "")
                            .ToolTipText = "Ty gia: " & rs.Fields("FCRATE")
                        Else
                            Exit Sub
                        End If
                        .Visible = True
                        .Height = 900
                        .Width = 1600
            
                End With
            rs.MoveNext
            Index = Index + 1
        Next j
    Next i

End If
Exit Sub
'Handle:
'    MsgBox Err.Number & Err.Description & Me.Name & "  LoadCommandSub"
End Sub

Public Property Get Get_TotalAmt() As Variant
    Get_TotalAmt = Total
End Property

Public Property Let Get_TotalAmt(ByVal vNewValue As Variant)
    Total = vNewValue
End Property

Public Property Get GetBill_Number() As Variant
    GetBill_Number = BillNO
End Property

Public Property Let GetBill_Number(ByVal vNewValue As Variant)
    BillNO = vNewValue
End Property

Public Property Get Get_Cust() As Variant
    Get_Cust = Customer
End Property

Public Property Let Get_Cust(ByVal vNewValue As Variant)
    Customer = vNewValue
End Property

Public Property Get Get_Record_Payment() As Variant
    Set Get_Record_Payment = rsItemPayment
End Property

Public Property Let Get_Record_Payment(ByVal vNewValue As Variant)
   Set rsItemPayment = vNewValue
End Property

Public Property Let Get_Service_Charge(ByVal vNewValue As Variant)
    service_Charge = vNewValue
End Property

Public Property Let Get_Adj1(ByVal vNewValue As Variant)
    Adjtotal1 = vNewValue
End Property
Public Property Let Get_Adj2(ByVal vNewValue As Variant)
    Adjtotal2 = vNewValue
End Property
Public Property Let Get_Adj3(ByVal vNewValue As Variant)
    Adjtotal3 = vNewValue
End Property
Public Property Let Get_Adj4(ByVal vNewValue As Variant)
    Adjtotal4 = vNewValue
End Property


Public Property Get Get_Discount() As Variant
    Get_Discount = Discount
End Property

Public Property Let Get_Discount(ByVal vNewValue As Variant)
    Discount = vNewValue
End Property
