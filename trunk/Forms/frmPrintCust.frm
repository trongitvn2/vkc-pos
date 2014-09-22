VERSION 5.00
Begin VB.Form frmPrintCust 
   Caption         =   "Th«ng tin kh¸ch hµng"
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12585
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
   ScaleHeight     =   7260
   ScaleWidth      =   12585
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5040
      Top             =   6240
   End
   Begin VB.Frame Frame1 
      Height          =   6135
      Left            =   6120
      TabIndex        =   6
      Top             =   0
      Width           =   6255
      Begin VB.PictureBox Pic1 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   2760
         Picture         =   "frmPrintCust.frx":0000
         ScaleHeight     =   975
         ScaleWidth      =   975
         TabIndex        =   21
         Top             =   4080
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblMess 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   975
         Left            =   120
         TabIndex        =   20
         Top             =   5040
         Width           =   6015
      End
      Begin VB.Label lblDiscount 
         Caption         =   "Gi¶m:"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   5535
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2400
         TabIndex        =   8
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Tæng céng:"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1935
      End
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   855
      Left            =   9360
      TabIndex        =   1
      Top             =   6240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "§ãng"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
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
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPrintCust.frx":0509
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdPrint 
      Height          =   855
      Left            =   6360
      TabIndex        =   0
      Top             =   6240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "L­u Th«ng Tin"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
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
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPrintCust.frx":0525
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdAddNew 
      Height          =   735
      Left            =   1920
      TabIndex        =   19
      Top             =   4920
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "T¹o míi kh¸ch hµng"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
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
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPrintCust.frx":0541
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblCustPhone 
      Alignment       =   2  'Center
      Caption         =   " "
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   18
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Label lblCustAdd 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   17
      Top             =   1800
      Width           =   4455
   End
   Begin VB.Label lblCustName 
      Caption         =   " "
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   16
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label lblPoint 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   15
      Top             =   4200
      Width           =   3615
   End
   Begin VB.Label lblDateBirth 
      Alignment       =   1  'Right Justify
      Caption         =   " "
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   14
      Top             =   3000
      Width           =   4215
   End
   Begin VB.Label lblAccountBalance 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   13
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label b 
      Caption         =   "C«ng nî hiÖn t¹i:"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label lbl 
      Caption         =   "§iÓm tÝch lòy:"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label v 
      Caption         =   "Ngµy sinh:"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label lblPhone 
      Caption         =   "®iÖn tho¹i:"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblAdd 
      Caption         =   "§i¹ chØ:"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblCustomerName 
      Caption         =   "Tªn kh¸ch hµng:"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Th«ng tin kh¸ch hµng"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   5775
   End
End
Attribute VB_Name = "frmPrintCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cust_ID As String
Dim Cust_Dis, adj1, adj2 As Integer
Dim Total, Amount As Double
Dim rscust As New ADODB.Recordset
Dim Accepted As Boolean
Dim i
Dim Amount_Get_Point, Pnt As Integer

Private Sub cmdAddNew_Click()
    frmAddCustomer.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Handle
    If Accepted = True Then
        With frmOrder
            .Get_Discount = Cust_Dis
            CustNo(0) = Trim(Cust_ID)
            .Get_Adj1 = adj1
            .Get_Adj2 = adj2
        End With
        Set rscust = Open_Table(cnData, "Customer")
        With rscust
        .Find "CustNum='" & Cust_ID & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Fields("Totals") = Amount
                Amount = .Fields("Totals")
                .Fields("Point") = Get_Point
                .Update
            End If
        End With
    End If
    Unload Me
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdPrint_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
Dim strSql As String

'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    
    strSql = "SELECT Customer.CustNum, Customer.CustName, Customer.Company, Customer.Address, Customer.Phone, Customer.Fax, Customer.TaxCode, Customer.AccountNo, Customer.Acct_Open_Date, Customer.Acct_Close_Date, Customer.Acct_Balance, Customer.Cashier, Customer.Acct_Max_Balance, Customer.Birthday, Customer.Point, Customer.Totals, Customer_Type.Promotion, Customer_Type.Pro_Value" & _
                   " FROM Customer_Type INNER JOIN Customer ON Customer_Type.CustType_ID = Customer.Cust_Type "
    Set rscust = OpenCriticalTable(strSql, cnData)
    If Not Check_Field_Exist(rscust, "Totals") Then
        cnData.Execute "ALTER TABLE Customer " _
                             & "ADD COLUMN Totals Double;"
    End If
    Call Load_CustInfor
    lblTotal.Caption = Format(Total, "#,##0")
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Public Property Let Get_CustID(ByVal vNewValue As Variant)
    Cust_ID = vNewValue
End Property

Public Property Get Return_Dist() As Variant
    Return_Dist = Cust_Dis
End Property

Public Sub Load_CustInfor()
On Error GoTo Handle
    With rscust
        .Find "CustNum='" & Cust_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            lblCustName.Caption = "" & .Fields("CustName")
            lblCustAdd.Caption = "" & .Fields("Address")
            lblCustPhone.Caption = "" & .Fields("Phone")
            lblDateBirth.Caption = .Fields("Birthday")
            lblAccountBalance.Caption = CDbl("0" & .Fields("Acct_Balance"))
           Select Case .Fields("Promotion")
                Case 0
                    Cust_Dis = 0
                    adj1 = 0
                    adj2 = 0
                    lblDiscount.Caption = "Kh«ng cã h×nh thøc khuyÕn m·i cho kh¸ch hµng hµy"
                Case 1
                    Cust_Dis = .Fields("Pro_Value")
                    adj1 = 0
                    adj2 = 0
                    lblDiscount.Caption = "Gi¶m tæng H§ " & .Fields("Pro_Value") & "%"
                Case 2
                    Cust_Dis = 0
                    adj1 = .Fields("Pro_Value")
                    adj2 = 0
                    lblDiscount.Caption = "Gi¶m tæng Thøc ¨n " & .Fields("Pro_Value") & "%"
                Case 3
                    Cust_Dis = 0
                    adj1 = 0
                    adj2 = .Fields("Pro_Value")
                    lblDiscount.Caption = "Gi¶m tæng Thøc uèng " & .Fields("Pro_Value") & "%"
            End Select
            Amount = CDbl("0" & .Fields("Totals"))
            CustNo(1) = "" & .Fields("CustName")
            lblPoint.Caption = Get_Point '''' Ham lay diem tich luy
            cmdAddNew.Enabled = False
            If Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") > Format(Year(.Fields("Acct_Close_Date")), "0000") & Format(Month(.Fields("Acct_Close_Date")), "00") & Format(Day(.Fields("Acct_Close_Date")), "00") Then
                Accepted = False
                lblMess.Caption = "Tµi kho¶n thÎ nµy ®· hÕt h¹n, vui lßng liªn hÖ bé phËn qu¶n lý thÎ!"
            Else
                Accepted = True
                lblMess.Visible = False
                Pic1.Visible = False
            End If
        Else
            lblCustName.Font.Size = 10
            lblCustName.ForeColor = vbRed
            lblCustName.Caption = "Kh«ng t×m thÊy th«ng tin kh¸ch hµng !"
            
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Load_CustInfor"
End Sub

Public Property Let Get_Total(ByVal vNewValue As Variant)
    Total = vNewValue
End Property

Private Sub Form_Unload(Cancel As Integer)
    Cust_Dis = 0
    Cust_ID = ""
End Sub

Private Sub Timer1_Timer()
On Error GoTo Handle
    If Accepted Then Exit Sub
    i = i + 1
        If i Mod 2 = 0 Then
            Pic1.Visible = True
        Else
            Pic1.Visible = False
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Timer1_Timer"
End Sub

Public Function Get_Point() As Integer
On Error GoTo Handle
    Dim rsPoint As New ADODB.Recordset
    Dim kq, i As Integer
    Set rsPoint = Open_Table(cnData, "Customer_Point_Sale")
    With rsPoint
        If Not .EOF Then
            Amount_Get_Point = .Fields("Amount_Get_Point")
            Pnt = .Fields("Point")
        End If
    End With
    
    i = Int(Amount / Amount_Get_Point)
    Do Until i = 0
        kq = kq + Pnt
    i = i - 1
    Loop
    Get_Point = CDbl("0" & kq)
Exit Function
Handle:
    Get_Point = 0
    MsgBox Err.Number & Err.Description & Me.name & " Get_Point"

End Function
