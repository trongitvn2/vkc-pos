VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmTotalDetails 
   BackColor       =   &H000000FF&
   Caption         =   " "
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10425
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FFFF&
      Height          =   9495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin prjTouchScreen.MyButton cmdClose 
         Cancel          =   -1  'True
         Height          =   975
         Left            =   1680
         TabIndex        =   1
         Top             =   7680
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1720
         BTYPE           =   3
         TX              =   "&Close"
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
         FCOL            =   255
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTotalDetails.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "chi tiÕt hãa ®¬n"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   30
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bµn:"
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
         Left            =   120
         TabIndex        =   28
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lblTable 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   960
         TabIndex        =   27
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "H.§¬n:"
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
         Left            =   120
         TabIndex        =   26
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblBill 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   1440
         TabIndex        =   25
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Gi¶m % tæng H§:"
         BeginProperty Font 
            Name            =   ".VnArial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Gi¶m % Thøc ¨n:"
         BeginProperty Font 
            Name            =   ".VnArial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Gi¶m % thøc uèng:"
         BeginProperty Font 
            Name            =   ".VnArial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "PhÝ phôc vô:"
         BeginProperty Font 
            Name            =   ".VnArial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   4560
         Width           =   2655
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Phô thu tiÒn mÆt:"
         BeginProperty Font 
            Name            =   ".VnArial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   5040
         Width           =   2655
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "ThuÕ VAT:"
         BeginProperty Font 
            Name            =   ".VnArial Narrow"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   6000
         Width           =   2655
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Thanh to¸n:"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   6480
         Width           =   2655
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
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
         Left            =   2640
         TabIndex        =   16
         Top             =   2640
         Width           =   2775
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   15
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label lblAdj1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   14
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label lblAdj2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   13
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label lblSerCharge 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   12
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label lblMoney 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   11
         Top             =   5040
         Width           =   1575
      End
      Begin VB.Label lblVAT 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         TabIndex        =   10
         Top             =   6000
         Width           =   1575
      End
      Begin VB.Label lblCash 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   9
         Top             =   6480
         Width           =   2775
      End
      Begin VB.Label lblDiscountPer 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   8
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblVATPer 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   7
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Label lblSerchargePer 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   6
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label lblAdj2Per 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   5
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lblAdj1Per 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   4
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   3240
         X2              =   5520
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "TiÒn tr­íc thuÕ:"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   5520
         Width           =   2655
      End
      Begin VB.Label lblNoneAddOn 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   5400
         Width           =   2775
      End
   End
   Begin CRVIEWERLibCtl.CRViewer crvReport 
      CausesValidation=   0   'False
      Height          =   10575
      Left            =   5640
      TabIndex        =   29
      Top             =   0
      Width           =   4815
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmTotalDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Total, Cash, BillNO, discount, adj1, adj2, Adj3, Adj4, Adj5, Adj6, Money, VAT, Sercharge As Double
Dim DiscountPer, Adj1Per, Adj2Per, Adj3Per, Adj4Per, Adj5Per, Adj6Per, VATPer, SerchargePer As Integer
Dim TableNo As String

Private Sub cmdClose_Click()
    Unload Me
End Sub



Private Sub Form_Load()
On Error GoTo Handle
    discount = -DiscountPer * Total / 100
    Sercharge = SerchargePer * Total / 100
    lblTable.Caption = TableNo
    lblBill.Caption = BillNO
    lblTotal.Caption = Format(Total, "#,##0")
    lblDiscount.Caption = Format(discount, "#,##0")
    lblDiscountPer.Caption = DiscountPer & "%"
    lblAdj1.Caption = Format(adj1, "#,##0")
    lblAdj1Per.Caption = Adj1Per & "%"
    lblAdj2.Caption = Format(adj2, "#,##0")
    lblAdj2Per.Caption = Adj2Per & "%"
    lblSerCharge.Caption = Format(Sercharge, "#,##0")
    lblSerchargePer.Caption = SerchargePer & "%"
    lblMoney.Caption = Format(Money, "#,##0")
    
    lblNoneAddOn.Caption = Format(Total + discount + adj1 + adj2 + Adj3 + Adj4 + Adj5 + Adj6 + Sercharge + Money, "#,##0")
    VAT = (Total + discount + adj1 + adj2 + adj2 + Adj3 + Adj4 + Adj5 + Adj6 + Sercharge + Money) * VATPer / 100
    lblVAT.Caption = Format(VAT, "#,##0")
    lblVATPer.Caption = VATPer & "%"
    lblCash.Caption = Format(Total + discount + adj1 + adj2 + Adj3 + Adj4 + Adj5 + Adj6 + Sercharge + Money + VAT, "#,##0")
    Call Load_Bill(BillNO)
Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & "Form_load"
    MsgBox Err.Number & Err.Description & Me.name
End Sub

Public Property Let Get_Total(ByVal vNewValue As Variant)
    Total = vNewValue
End Property
Public Property Let Get_Table(ByVal vNewValue As Variant)
    TableNo = vNewValue
End Property
Public Property Let Get_Bill(ByVal vNewValue As Variant)
    BillNO = vNewValue
End Property

Public Property Let Get_Cash(ByVal vNewValue As Variant)
    Cash = vNewValue
End Property
Public Property Let Get_DiscountPer(ByVal vNewValue As Variant)
    DiscountPer = vNewValue
End Property


Public Property Let Get_Adj1Per(ByVal vNewValue As Variant)
    Adj1Per = vNewValue
End Property
Public Property Let Get_Adj1(ByVal vNewValue As Variant)
    adj1 = vNewValue
End Property
Public Property Let Get_Adj2(ByVal vNewValue As Variant)
    adj2 = vNewValue
End Property
Public Property Let Get_Adj2Per(ByVal vNewValue As Variant)
    Adj2Per = vNewValue
End Property

Public Property Let Get_Adj3Per(ByVal vNewValue As Variant)
    Adj3Per = vNewValue
End Property
Public Property Let Get_Adj3(ByVal vNewValue As Variant)
    Adj3 = vNewValue
End Property

Public Property Let Get_Adj4Per(ByVal vNewValue As Variant)
    Adj4Per = vNewValue
End Property
Public Property Let Get_Adj4(ByVal vNewValue As Variant)
    Adj4 = vNewValue
End Property

Public Property Let Get_Adj5Per(ByVal vNewValue As Variant)
    Adj5Per = vNewValue
End Property
Public Property Let Get_Adj5(ByVal vNewValue As Variant)
    Adj5 = vNewValue
End Property


Public Property Let Get_Adj6Per(ByVal vNewValue As Variant)
    Adj6Per = vNewValue
End Property
Public Property Let Get_Adj6(ByVal vNewValue As Variant)
    Adj6 = vNewValue
End Property



Public Property Let Get_Sercharge(ByVal vNewValue As Variant)
    SerchargePer = vNewValue
End Property
Public Property Let Get_Money(ByVal vNewValue As Variant)
    Money = vNewValue
End Property
Public Property Let Get_VAT(ByVal vNewValue As Variant)
    VATPer = vNewValue
End Property

Public Sub Load_Bill(ByVal Bill_No As Double)
    On Error Resume Next
    Dim cmd As New ADODB.Command
    Dim SQL As String
    Dim RptID As Integer
    Dim ReceiptReport As CRAXDDRT.Report
    Dim iReport As CRAXDDRT.Report
    Dim DescArr() As String
    
    '
    
    Dim crDatabase As CRAXDRT.Database
    Dim CrDBTables As CRAXDRT.DatabaseTables
    Dim CrDBTable As CRAXDRT.DatabaseTable
    
    Set crDatabase = ReceiptReport.Database
    Set CrDBTables = crDatabase.Tables
    Set CrDBTable = CrDBTables.Item(1)
    
    CrDBTable.SetLogOnInfo ServerName, DataBaseName, UserLog, DB_Password



    DescArr = LoadLanguage(LngFile, "#02:005:")
    
    If ArrayFlag(SF(0), 5) = 0 Then
        If ArrayFlag(SF(6), 2) = 1 Then
         SQL = "SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.KarDiscount, Invoice_Totals.CustNum," & _
            "Invoice_Totals.Discount, Invoice_Totals.Total_Price, Invoice_Totals.Service_Charge," & _
            "Invoice_Totals.VATFee, Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment1, Invoice_Totals.Adj2Rate, " & _
            "Invoice_Totals.Personals, Invoice_Totals.Adjustment2, Invoice_Totals.Adj3Rate,Invoice_Totals.Adjustment3,Invoice_Totals.Adj5Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj6Rate,Invoice_Totals.Adjustment6," & _
            "Invoice_Totals.Adj4Rate,Invoice_Totals.Adjustment4, Invoice_Totals.AddMoney, Invoice_Totals.Tax_Rate_ID," & _
            "Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered, Invoice_Totals.Amt_Change," & _
            "Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID, Invoice_Totals.InvType,Invoice_Totals.Reserve, " & _
            "Invoice_Itemized.ItemNum, Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer," & _
            "Sum(Invoice_Itemized.Amt) AS Amt, Invoice_Itemized.DiffItemName," & _
            "Invoice_Totals.Orig_OnHoldID,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc, Invoice_Totals.OrderMan, Right([OpenTime],12) AS TimeIn, Right([ClosingTime],12) AS TimeOut, Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount" & _
            " FROM ((((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number) INNER JOIN Inventory ON Invoice_Itemized.ItemNum = Inventory.ItemNum) INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID) INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo" & _
            " WHERE (((Invoice_Itemized.ItemNum)<>'KAR') AND ((Invoice_Totals.Invoice_Number)=" & BillNO & "))" & _
            " GROUP BY Invoice_Totals.Invoice_Number, Invoice_Totals.KarDiscount,Invoice_Totals.Personals," & _
            " Invoice_Totals.CustNum, Invoice_Totals.Discount, Invoice_Totals.Total_Price, Invoice_Totals.Service_Charge, " & _
            " Invoice_Totals.VATFee, Invoice_Totals.Adjustment1, Invoice_Totals.Tax_Rate_ID, Invoice_Totals.Adjustment2, Invoice_Totals.Adj2Rate, Invoice_Totals.Adj5Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj6Rate,Invoice_Totals.Adjustment6," & _
            " Invoice_Totals.Adj3Rate,Invoice_Totals.Adj4Rate,Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment3, Invoice_Totals.Adjustment4, Invoice_Totals.AddMoney,Invoice_Totals.Reserve, Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered, Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID, Invoice_Totals.InvType, Invoice_Itemized.ItemNum, Invoice_Itemized.PricePer,Invoice_Itemized.Line_Disc_Desc, Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc, Invoice_Totals.Orig_OnHoldID, Invoice_Totals.OrderMan, Right([OpenTime],12), Right([ClosingTime],12), Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount" & _
            " ORDER BY Invoice_Itemized.ItemNum"
        Else
        SQL = "SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.KarDiscount, Invoice_Totals.CustNum," & _
            "Invoice_Totals.Discount, Invoice_Totals.Total_Price," & _
            "Invoice_Totals.Tax_Rate_ID, Invoice_Totals.Service_Charge, Invoice_Totals.VATFee," & _
            "Invoice_Totals.Adjustment1,Invoice_Totals.Personals, Invoice_Totals.Adj2Rate, Invoice_Totals.Adj1Rate," & _
            "Invoice_Totals.Adjustment2, Invoice_Totals.Adj3Rate,Invoice_Totals.Adjustment3, Invoice_Totals.Adj4Rate,Invoice_Totals.Adjustment4," & _
            "Invoice_Totals.AddMoney, Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered, Invoice_Totals.Adj5Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj6Rate,Invoice_Totals.Adjustment6," & _
            "Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID,Invoice_Totals.Reserve, " & _
            "Invoice_Totals.InvType, Invoice_Itemized.ItemNum, Sum(Invoice_Itemized.Quantity) AS Qty," & _
            "Invoice_Itemized.LineNum,Invoice_Itemized.PricePer, Sum(Invoice_Itemized.Amt) AS Amt," & _
            "Invoice_Itemized.DiffItemName, Invoice_Totals.Orig_OnHoldID,Invoice_Itemized.LineDisc," & _
            "Invoice_Itemized.Line_Disc_Desc, Invoice_Totals.OrderMan, Right([OpenTime],12) AS TimeIn, Right([ClosingTime],12) AS TimeOut, Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount" & _
            " FROM ((((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number) INNER JOIN Inventory ON Invoice_Itemized.ItemNum = Inventory.ItemNum) INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID) INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo" & _
            " WHERE (((Invoice_Itemized.ItemNum)<>'KAR') AND ((Invoice_Totals.Invoice_Number)=" & BillNO & "))" & _
            " GROUP BY Invoice_Totals.Invoice_Number, Invoice_Totals.KarDiscount, Invoice_Totals.CustNum," & _
            " Invoice_Totals.Tax_Rate_ID,Invoice_Totals.Discount, Invoice_Totals.Total_Price,Invoice_Itemized.Line_Disc_Desc, Invoice_Totals.Reserve, " & _
            " Invoice_Totals.Service_Charge, Invoice_Totals.VATFee, Invoice_Totals.Adjustment1,Invoice_Totals.Adj5Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj6Rate,Invoice_Totals.Adjustment6," & _
            " Invoice_Totals.Adjustment2, Invoice_Totals.Adj2Rate,Invoice_Totals.Adj3Rate,Invoice_Totals.Adj4Rate,Invoice_Totals.Personals, Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment3, Invoice_Totals.Adjustment4, Invoice_Totals.AddMoney, Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered, Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID, Invoice_Totals.InvType, Invoice_Itemized.ItemNum, Invoice_Itemized.PricePer, Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc, Invoice_Totals.Orig_OnHoldID, Invoice_Totals.OrderMan, Right([OpenTime],12), Right([ClosingTime],12), Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount, Invoice_Itemized.LineNum" & _
            " ORDER BY Invoice_Itemized.LineNum desc"
        End If
    Else
        If ArrayFlag(SF(6), 2) = 0 Then
        SQL = "SELECT Invoice_Totals.Invoice_Number,Invoice_Totals.KarDiscount, Invoice_Totals.CustNum," & _
            " Invoice_Totals.Discount, Invoice_Totals.Total_Price," & _
            " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adjustment1," & _
            " Invoice_Totals.Adjustment2,Invoice_Totals.Adj3Rate,Invoice_Totals.Adjustment3,Invoice_Totals.Adj5Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj6Rate,Invoice_Totals.Adjustment6," & _
            " Invoice_Totals.Adj2Rate,Invoice_Totals.Personals, Invoice_Totals.Adj1Rate,Invoice_Totals.Adj4Rate,Invoice_Totals.Adjustment4,Invoice_Totals.AddMoney," & _
            " Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered,Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID,Invoice_Totals.Reserve, " & _
            " Invoice_Totals.Station_ID,Invoice_Totals.Tax_Rate_ID,Invoice_Totals.InvType,Invoice_Itemized.ItemNum, " & _
            " Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer, Invoice_Itemized.LineNum,Invoice_Itemized.Line_Disc_Desc," & _
            " sum(Invoice_Itemized.Amt) as Amt, " & _
            " Invoice_Itemized.DiffItemName ,Invoice_Itemized.LineDisc ,Invoice_Totals.Orig_OnHoldID,Invoice_Totals.OrderMan, " & _
            " Right([OpenTime],12) AS TimeIn, Right([ClosingTime],12) AS TimeOut, Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount, MainGroup.GroupNo, MainGroup.GroupName " & _
            " FROM ((((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number) INNER JOIN Inventory ON Invoice_Itemized.ItemNum = Inventory.ItemNum) INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID) INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo " & _
            " Where Invoice_Itemized.ItemNum<>'KAR' and Invoice_Totals.Invoice_Number=" & BillNO & _
            " group by Invoice_Itemized.ItemNum,Invoice_Totals.Invoice_Number," & _
            " Invoice_Totals.CustNum,Invoice_Totals.Discount,Invoice_Totals.KarDiscount," & _
            " Invoice_Totals.Total_Price,Invoice_Totals.Grand_Total," & _
            " Invoice_Totals.Amt_Tendered,Invoice_Totals.Amt_Change," & _
            " Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID,Invoice_Totals.Tax_Rate_ID," & _
            " Invoice_Itemized.PricePer, Invoice_Itemized.LineNum, Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc ," & _
            " Invoice_Totals.Orig_OnHoldID,Invoice_Totals.OrderMan, Invoice_Totals.InvType, Invoice_Totals.Reserve,Invoice_Totals.Adj4Rate, Invoice_Totals.Adj5Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj6Rate,Invoice_Totals.Adjustment6," & _
            " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adj2Rate, Invoice_Totals.Adj3Rate,Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment1,Invoice_Totals.Adjustment2,Invoice_Totals.Adjustment3," & _
            " Invoice_Totals.Adjustment4,Invoice_Totals.Personals,Invoice_Totals.AddMoney, Right([OpenTime],12), Right([ClosingTime],12), Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount, MainGroup.GroupNo, MainGroup.GroupName" & _
            " order by Invoice_Itemized.LineNum Desc"
        Else

             SQL = "SELECT Invoice_Totals.Invoice_Number,Invoice_Totals.KarDiscount, Invoice_Totals.CustNum," & _
            " Invoice_Totals.Discount, Invoice_Totals.Total_Price," & _
            " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adjustment1,Invoice_Totals.Adj5Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj6Rate,Invoice_Totals.Adjustment6," & _
            " Invoice_Totals.Adjustment2,Invoice_Totals.Adj3Rate,Invoice_Totals.Adjustment3,Invoice_Totals.Adj2Rate,Invoice_Totals.Personals, Invoice_Totals.Adj1Rate," & _
            " Invoice_Totals.Adj4Rate,Invoice_Totals.Adjustment4,Invoice_Totals.AddMoney,Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered," & _
            " Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID,Invoice_Totals.Tax_Rate_ID,Invoice_Totals.Reserve, " & _
            " Invoice_Totals.Station_ID,Invoice_Totals.InvType,Invoice_Itemized.ItemNum, " & _
            " Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer, " & _
            " sum(Invoice_Itemized.Amt) as Amt, " & _
            " Invoice_Itemized.DiffItemName ,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc ,Invoice_Totals.Orig_OnHoldID,Invoice_Totals.OrderMan, " & _
            " Right([OpenTime],12) AS TimeIn, Right([ClosingTime],12) AS TimeOut, Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount, MainGroup.GroupNo, MainGroup.GroupName " & _
            " FROM ((((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number) INNER JOIN Inventory ON Invoice_Itemized.ItemNum = Inventory.ItemNum) INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID) INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo " & _
            " Where Invoice_Itemized.ItemNum<>'KAR' and Invoice_Totals.Invoice_Number=" & BillNO & _
            " group by Invoice_Itemized.ItemNum,Invoice_Totals.Invoice_Number," & _
            " Invoice_Totals.CustNum,Invoice_Totals.Discount,Invoice_Totals.KarDiscount," & _
            " Invoice_Totals.Total_Price,Invoice_Totals.Grand_Total," & _
            " Invoice_Totals.Amt_Tendered,Invoice_Totals.Amt_Change," & _
            " Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID,Invoice_Totals.Tax_Rate_ID,Invoice_Totals.Reserve, " & _
            " Invoice_Itemized.PricePer,  Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc ," & _
            " Invoice_Totals.Orig_OnHoldID,Invoice_Totals.OrderMan, Invoice_Totals.InvType,Invoice_Totals.Adj5Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj6Rate,Invoice_Totals.Adjustment6, " & _
            " Invoice_Totals.Service_Charge,Invoice_Totals.Personals,Invoice_Totals.VATFee,Invoice_Totals.Adj2Rate, Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment1,Invoice_Totals.Adjustment2,Invoice_Totals.Adj3Rate,Invoice_Totals.Adjustment3," & _
            " Invoice_Totals.Adj4Rate,Invoice_Totals.Adjustment4,Invoice_Totals.AddMoney, Right([OpenTime],12), Right([ClosingTime],12), Invoice_Totals_Notes.Total_Minute, Invoice_Totals_Notes.Karaoke_Amount, MainGroup.GroupNo, MainGroup.GroupName" & _
            " order by Invoice_Itemized.ItemNum"
        End If
   End If
    
    Dim CRXReportField As CRAXDDRT.DatabaseFieldDefinition
    Set crBalance75 = Nothing
    Set crBalance58 = Nothing
    Set crBalance = Nothing
    If ReceiptType = "80" Then
        Set ReceiptReport = crBalance
    ElseIf ReceiptType = "58" Then
        Set ReceiptReport = crBalance58
    ElseIf ReceiptType = "75" Then
        Set ReceiptReport = crBalance75
    ElseIf ReceiptType = "A5" Then
        Set ReceiptReport = crBalanceA5
    End If
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
'        Set rs = OpenCriticalTable(SQL, cnData)
'    Print #fFile, "In H§" & vbTab & ":" & userName
'        With rs
'            Print #fFile, "Bµn:" & .Fields("Orig_OnHoldID") & vbTab & "H§ sè:" & .Fields("Invoice_Number")
'            Do While Not rs.EOF
'                Print #fFile, vbTab & .Fields("ItemNum") & vbTab & .Fields("DiffItemName") & vbTab & .Fields("Qty") & vbTab & .Fields("PricePer") & vbTab & .Fields("amt")
'                .MoveNext
'            Loop
'        End With
    'Print #fFile, "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
    
    With ReceiptReport
        .Database.AddADOCommand cnData, cmd
        .txtPluCode.SetUnboundFieldSource "{ado.DiffItemName}"
        .txtPluName.SetUnboundFieldSource "{ado.ItemNum}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtCost.SetUnboundFieldSource "{ado.PricePer}"
        .txtAmt.SetUnboundFieldSource "{ado.Amt}"
        .LineDisc.SetUnboundFieldSource "{ado.LineDisc}"
'        .Cost1.SetUnboundFieldSource "{ado.PricePer}"
        .txtBillNo.SetUnboundFieldSource "{ado.Invoice_Number}"
        .txtCashier.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtCustomerID.SetUnboundFieldSource "{ado.CustNum}"
        .txtChange.SetUnboundFieldSource "{ado.Amt_Change}"
        .txtDiscount.SetUnboundFieldSource "{ado.Discount}"
        .txtTable.SetUnboundFieldSource "{ado.Orig_OnHoldID}"
        .txtserver.SetUnboundFieldSource "{ado.Station_ID}"
        .txtOrder.SetUnboundFieldSource "{ado.OrderMan}"
       
       .txtAdj1.SetUnboundFieldSource "{ado.Adjustment1}"
        .txtAdj1Rate.SetUnboundFieldSource "{ado.Adj1Rate}"
        
        .txtAdj2.SetUnboundFieldSource "{ado.Adjustment2}"
        .txtAdj2Rate.SetUnboundFieldSource "{ado.Adj2Rate}"
        
        .txtAdj3.SetUnboundFieldSource "{ado.Adjustment3}"
        .txtAdj3Rate.SetUnboundFieldSource "{ado.Adj3Rate}"
        
        .txtAdj4.SetUnboundFieldSource "{ado.Adjustment4}"
        .txtAdj4Rate.SetUnboundFieldSource "{ado.Adj4Rate}"
        
        .txtAdj5.SetUnboundFieldSource "{ado.Adjustment5}"
        .txtAdj5Rate.SetUnboundFieldSource "{ado.Adj5Rate}"
        
        .txtAdj6.SetUnboundFieldSource "{ado.Adjustment6}"
        .txtAdj6Rate.SetUnboundFieldSource "{ado.Adj6Rate}"
        
        .txtSev.SetUnboundFieldSource "{ado.Service_Charge}"
        .txtVAT.SetUnboundFieldSource "{ado.VATFee}"
        .txtMoney.SetUnboundFieldSource "{ado.AddMoney}"
        .printcount.SetUnboundFieldSource "{ado.InvType}"
        .txtMixmatch.SetUnboundFieldSource "{ado.Tax_Rate_ID}"
        .txtsokhach.SetUnboundFieldSource "{ado.Personals}"
        .txtLineDiscDesc.SetUnboundFieldSource "{ado.Line_Disc_Desc}"
        .txtReserved.SetUnboundFieldSource "{ado.Reserve}"

        .lblTitle.SetText DescArr(24)
        If ArrayFlag(SF(0), 5) = 1 Then
            .txtMaingroup.SetUnboundFieldSource "{ado.GroupNo}"
        End If
        .lblTable.SetText DescArr(3)
        .lblBillNo.SetText DescArr(2)
        .lblItems.SetText DescArr(4)
        .lblQty.SetText DescArr(5)
        .lblPrice.SetText DescArr(6)
        .lblAmt.SetText DescArr(7)
        .lblTotal.SetText DescArr(8)
        '.lblDiscount.SetText DescArr(9)
        .lblRead.SetText DescArr(12)
        .lblCashier.SetText DescArr(13)
        .lblPhuthu.SetText DescArr(14)
        .lblTotal1.SetText DescArr(15)
        .lblServer.SetText DescArr(16)
        .lbldate.SetText DescArr(17)
        .lblTime.SetText DescArr(18)
        .lblCash.SetText DescArr(19)
        .lblOrder.SetText DescArr(20)
        .lblCustomer.SetText DescArr(21)
        .lblSignal.SetText DescArr(22)
        .lblAdj1.SetText DescArr(25)
        .lblAdj2.SetText DescArr(26)
        .lblPhuphi.SetText DescArr(27)
        .lblVAT.SetText DescArr(29)
        .lblPrintCount.SetText DescArr(30)
        
        
        'canh le
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        
        With .txtCost
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtMoney
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAdj4
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAdj3
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAdj2
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAdj1
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtChange
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        
        With .txtMainTotal
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtServAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        
        With .TxtTotal
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtTotalAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
       
    End With
'    Call format_Balance_Bill
    Set iReport = ReceiptReport
    With crvReport
        .DisplayBorder = False
        .ReportSource = iReport
        .EnableSearchControl = False
        .EnableStopButton = False
        .EnableGroupTree = False
        .EnableAnimationCtrl = False
        .EnablePopupMenu = False
        .EnableToolbar = False
        .DisplayToolbar = False
        .DisplayTabs = False
        .ToolTipText = ""
        .ViewReport
        .Zoom 100
        While .IsBusy
            DoEvents
        Wend
        .ShowLastPage
        While .IsBusy
            DoEvents
        Wend
        .ShowFirstPage
        While .IsBusy
            DoEvents
        Wend
    End With
    
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

