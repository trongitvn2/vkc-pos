VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCash 
   Caption         =   "Cash"
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10365
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdmoney 
      Height          =   1480
      Index           =   7
      Left            =   120
      Picture         =   "Form1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      Tag             =   "2000"
      Top             =   4560
      Width           =   3615
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1480
      Index           =   6
      Left            =   120
      Picture         =   "Form1.frx":8635
      Style           =   1  'Graphical
      TabIndex        =   27
      Tag             =   "5000"
      Top             =   6000
      Width           =   3615
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1455
      Index           =   5
      Left            =   120
      Picture         =   "Form1.frx":1221A
      Style           =   1  'Graphical
      TabIndex        =   26
      Tag             =   "10000"
      Top             =   7485
      Width           =   3615
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1450
      Index           =   4
      Left            =   120
      Picture         =   "Form1.frx":1B202
      Style           =   1  'Graphical
      TabIndex        =   25
      Tag             =   "20000"
      Top             =   8925
      Width           =   3615
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1450
      Index           =   3
      Left            =   3720
      Picture         =   "Form1.frx":24798
      Style           =   1  'Graphical
      TabIndex        =   24
      Tag             =   "500000"
      Top             =   8925
      Width           =   3615
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1455
      Index           =   2
      Left            =   3720
      Picture         =   "Form1.frx":2A60F
      Style           =   1  'Graphical
      TabIndex        =   23
      Tag             =   "200000"
      Top             =   7485
      Width           =   3615
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1480
      Index           =   0
      Left            =   3720
      Picture         =   "Form1.frx":2FD9D
      Style           =   1  'Graphical
      TabIndex        =   22
      Tag             =   "100000"
      Top             =   6000
      Width           =   3615
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1480
      Index           =   1
      Left            =   3720
      Picture         =   "Form1.frx":35030
      Style           =   1  'Graphical
      TabIndex        =   21
      Tag             =   "50000"
      Top             =   4560
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   7440
      TabIndex        =   7
      Top             =   1200
      Width           =   4575
      Begin MSForms.CommandButton cmdOrther 
         Height          =   1095
         Left            =   120
         TabIndex        =   35
         Top             =   8040
         Width           =   2175
         ForeColor       =   16711680
         BackColor       =   -2147483638
         VariousPropertyBits=   8388635
         Caption         =   "..."
         Size            =   "3836;1931"
         FontName        =   ".VnArialH"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdGiftCard 
         Height          =   1095
         Left            =   2295
         TabIndex        =   34
         Tag             =   "L5"
         Top             =   8040
         Width           =   2175
         ForeColor       =   16711680
         BackColor       =   -2147483638
         VariousPropertyBits=   8388635
         Caption         =   "PhiÕu quµ tÆng"
         Size            =   "3836;1931"
         FontName        =   ".VnArialH"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdBalance 
         Height          =   1095
         Left            =   2295
         TabIndex        =   33
         Tag             =   "L6"
         Top             =   6960
         Width           =   2175
         ForeColor       =   16711680
         BackColor       =   -2147483638
         VariousPropertyBits=   8388635
         Caption         =   "C«ng nî"
         Size            =   "3836;1931"
         FontName        =   ".VnArialH"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdCheck 
         Height          =   1095
         Left            =   120
         TabIndex        =   32
         Tag             =   "L8"
         Top             =   6960
         Width           =   2175
         ForeColor       =   16711680
         BackColor       =   -2147483638
         VariousPropertyBits=   8388635
         Caption         =   "Bill ký"
         Size            =   "3836;1931"
         FontName        =   ".VnArialH"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdCashTranfer 
         Height          =   1095
         Left            =   2295
         TabIndex        =   31
         Tag             =   "L4"
         Top             =   5880
         Width           =   2175
         ForeColor       =   16711680
         BackColor       =   -2147483638
         VariousPropertyBits=   8388635
         Caption         =   "chuyÓn kho¶n"
         Size            =   "3836;1931"
         FontName        =   ".VnArialH"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdCredit 
         Height          =   1095
         Left            =   120
         TabIndex        =   30
         Tag             =   "L7"
         Top             =   5880
         Width           =   2175
         ForeColor       =   16711680
         BackColor       =   -2147483638
         VariousPropertyBits=   8388635
         Caption         =   "ThÎ tÝn dông (Visa, Master..)"
         Size            =   "3836;1931"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAlpha 
         Cancel          =   -1  'True
         Height          =   1335
         Index           =   13
         Left            =   120
         TabIndex        =   29
         Tag             =   "L12"
         Top             =   4560
         Width           =   2175
         ForeColor       =   16711680
         BackColor       =   255
         Caption         =   "Tho¸t"
         Size            =   "3836;2355"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   315
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCash 
         Height          =   1335
         Left            =   2280
         TabIndex        =   0
         Tag             =   "L3"
         Top             =   4560
         Width           =   2175
         ForeColor       =   16777215
         BackColor       =   16711680
         VariousPropertyBits=   8388635
         Caption         =   "TiÒn mÆt"
         Size            =   "3836;2355"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   435
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAlpha 
         Height          =   1095
         Index           =   11
         Left            =   3000
         TabIndex        =   19
         Top             =   3480
         Width           =   1455
         ForeColor       =   16711680
         BackColor       =   -2147483638
         Caption         =   "000"
         Size            =   "2566;1931"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   525
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAlpha 
         Height          =   1095
         Index           =   10
         Left            =   1560
         TabIndex        =   18
         Top             =   3480
         Width           =   1455
         ForeColor       =   16711680
         BackColor       =   -2147483638
         Caption         =   "00"
         Size            =   "2566;1931"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   525
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAlpha 
         Height          =   1095
         Index           =   9
         Left            =   120
         TabIndex        =   17
         Top             =   3480
         Width           =   1455
         ForeColor       =   16711680
         BackColor       =   -2147483638
         Caption         =   "0"
         Size            =   "2566;1931"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   525
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAlpha 
         Height          =   1095
         Index           =   8
         Left            =   3000
         TabIndex        =   16
         Top             =   2400
         Width           =   1455
         ForeColor       =   16711680
         BackColor       =   -2147483638
         Caption         =   "9"
         Size            =   "2566;1931"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   525
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAlpha 
         Height          =   1095
         Index           =   7
         Left            =   1560
         TabIndex        =   15
         Top             =   2400
         Width           =   1455
         ForeColor       =   16711680
         BackColor       =   -2147483638
         Caption         =   "8"
         Size            =   "2566;1931"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   525
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAlpha 
         Height          =   1095
         Index           =   6
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   1455
         ForeColor       =   16711680
         BackColor       =   -2147483638
         Caption         =   "7"
         Size            =   "2566;1931"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   525
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAlpha 
         Height          =   1095
         Index           =   5
         Left            =   3000
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
         ForeColor       =   16711680
         BackColor       =   -2147483638
         Caption         =   "6"
         Size            =   "2566;1931"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   525
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAlpha 
         Height          =   1095
         Index           =   4
         Left            =   1560
         TabIndex        =   12
         Top             =   1320
         Width           =   1455
         ForeColor       =   16711680
         BackColor       =   -2147483638
         Caption         =   "5"
         Size            =   "2566;1931"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   525
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAlpha 
         Height          =   1095
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1455
         ForeColor       =   16711680
         BackColor       =   -2147483638
         Caption         =   "4"
         Size            =   "2566;1931"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   525
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAlpha 
         Height          =   1095
         Index           =   2
         Left            =   3000
         TabIndex        =   10
         Top             =   240
         Width           =   1455
         ForeColor       =   16711680
         BackColor       =   -2147483638
         Caption         =   "3"
         Size            =   "2566;1931"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   525
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAlpha 
         Height          =   1095
         Index           =   1
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   1455
         ForeColor       =   16711680
         BackColor       =   -2147483638
         Caption         =   "2"
         Size            =   "2566;1931"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   525
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdAlpha 
         Height          =   1095
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1455
         ForeColor       =   16711680
         BackColor       =   -2147483638
         Caption         =   "1"
         Size            =   "2566;1931"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   525
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   7440
      TabIndex        =   5
      Top             =   510
      Width           =   2925
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   7215
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   705
         Left            =   90
         TabIndex        =   3
         Top             =   630
         Width           =   7035
      End
      Begin MSFlexGridLib.MSFlexGrid flgPayment 
         Height          =   3105
         Left            =   60
         TabIndex        =   2
         Top             =   1470
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   5477
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblAmountList 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000B&
         Caption         =   "H×nh thøc thanh to¸n"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   30
         TabIndex        =   4
         Tag             =   "L2"
         Top             =   150
         Width           =   4395
      End
   End
   Begin MSForms.CommandButton cmdAlpha 
      Height          =   735
      Index           =   12
      Left            =   10440
      TabIndex        =   20
      Top             =   480
      Width           =   1455
      ForeColor       =   255
      BackColor       =   -2147483638
      Caption         =   "Xãa"
      Size            =   "2566;1296"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label lblTenderAmount 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "NhËp sè tiÒn thanh to¸n"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   7440
      TabIndex        =   6
      Tag             =   "L1"
      Top             =   120
      Width           =   3525
   End
End
Attribute VB_Name = "frmCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iCash, isclick As Boolean
Dim Total, totals, i As Double
Dim Payment_Method As String
Dim Customer As String
Dim DescArr() As String
Dim BillNO As Double
Dim rsInvoice_Items As New ADODB.Recordset
Dim rsInvoice_Onhold As New ADODB.Recordset
Dim rsInvoice_Total As New ADODB.Recordset
Dim rsInvoice_Notes As New ADODB.Recordset
Dim isActived As Boolean
Dim diemtichluy As Double
Dim rsPayment As New ADODB.Recordset
Dim CardID As String
Dim OA_Amount, CA_Amount, GC_Amount, CC_Amount, CT_Amount, ROA_Amount, Payment_Totals As Double
Dim State_Payment As Integer

Private Sub cmdAlpha_Click(Index As Integer)
    Select Case Index
        Case 0 To 11:
            If isclick = True Then
                 txtQty.Text = Format(txtQty.Text & cmdAlpha(Index).Caption, "#,##0")
            Else
                txtQty.Text = cmdAlpha(Index).Caption
                isclick = True
            End If
        Case 12:
            txtQty.Text = ""
        Case 13:
            Unload Me
            iCash = False
        Unload Me
    End Select
End Sub

Public Property Let GetTotals(ByVal vNewValue As Variant)
    totals = vNewValue
End Property
Public Property Let GetTotal(ByVal vNewValue As Variant)
    Total = vNewValue
End Property

Private Sub cmdBalance_Click()
On Error GoTo Handle
    Payment_Method = "OA"
    i = i + CDbl("0" & txtQty.Text)
    If Trim(Customer) = "101" Then
        MsgBox "Kh¸ch v·ng lai kh«ng ®­îc l­u vµo c«ng nî ", vbInformation
    With frmFindCustomer
        .FormCall = "CustomerSelect"
        .Show vbModal
    End With
'        Exit Sub
'    Else
    
        isclick = False
        OA_Amount = CDbl("0" & txtQty.Text)
        Call Add_rsPayment("OA", cmdBalance.Caption, OA_Amount)
        txtQty.Text = Format(CDbl("0" & txtAmount.Text) - Payment_Totals, "#,##0")
        Payment_Method = "OA"
        Call Cash
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdBalance_Click"
End Sub
Public Sub CheckRight()
On Error GoTo Handle
    Dim res As New ADODB.Recordset
        Set res = LoadPasswordData
        With MyRight
            res.MoveFirst
            Do While Not res.EOF
                If StrComp(res.Fields("ID"), UserID, 1) = 0 Then
                    .FullRight = res.Fields("UserRight")
                    .Banhang = RightDeCode(Mid(.FullRight, 65, 64))
                    Exit Do
                End If
                res.MoveNext
            Loop
            
            If Mid(.Banhang, 20, 1) = 0 Then
                  cmdCash.Enabled = False
            Else: cmdCash.Enabled = True
            End If
            
            If Mid(.Banhang, 21, 1) = 0 Then
                  cmdCashTranfer.Enabled = False
            Else: cmdCashTranfer.Enabled = True
            End If
            
            If Mid(.Banhang, 22, 1) = 0 Then
                  cmdCredit.Enabled = False
            Else: cmdCredit.Enabled = True
            End If
            
            If Mid(.Banhang, 23, 1) = 0 Then
                  cmdBalance.Enabled = False
            Else: cmdBalance.Enabled = True
            End If
            
            If Mid(.Banhang, 24, 1) = 0 Then
                  cmdGiftCard.Enabled = False
            Else: cmdGiftCard.Enabled = True
            End If
            
             If Mid(.Banhang, 25, 1) = 0 Then
                  cmdCheck.Enabled = False
            Else: cmdCheck.Enabled = True
            End If
        End With
    CloseRecordset res
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " CheckRight"
End Sub


Private Sub cmdCash_Click()
On Error GoTo Handle
        isclick = False
        iCash = True
        i = i + CDbl("0" & txtQty.Text)
        If Val("0" & totals) - Val("0" & Payment_Totals) <= CDbl("0" & txtQty.Text) Then
            CA_Amount = (totals + totals * VAT / 100) - Payment_Totals
        Else
            CA_Amount = CDbl("0" & txtQty.Text)
        End If
        Call Add_rsPayment("CA", cmdCash.Caption, CA_Amount)
        txtQty.Text = Format(CDbl("0" & txtAmount.Text) - Payment_Totals, "#,##0")
        Payment_Method = "C"
        'goi ham thanh toan
        Call Cash
Exit Sub
Handle:
    Exit Sub
    With frmChangeBill
        .Let_Bill = Bill
        .Show vbModal
    End With
    Exit Sub
MsgBox Err.Number & Err.Description & Me.name & "  cmdCash_Click"
End Sub

Private Sub cmdCheck_Click()
   On Error GoTo Handle
        isclick = False
        i = i + CDbl("0" & txtQty.Text)
        ROA_Amount = CDbl("0" & txtQty.Text)
        Call Add_rsPayment("ROA", cmdCheck.Caption, ROA_Amount)
        txtQty.Text = Format(CDbl("0" & txtAmount.Text) - Payment_Totals, "#,##0")
        Payment_Method = "ROA"
        Call Cash
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdCheck_Click"
End Sub

Private Sub cmdCredit_Click()
    On Error GoTo Handle
    Dim iscredit As Boolean
    Dim rsCredit_Payment As New ADODB.Recordset
        isclick = False
        With frmCredit_Card_infor
            .Show vbModal
            iscredit = .Let_OK
          Set rsCredit_Payment = .Let_Records
        End With
        If iscredit = True Then
            i = i + CDbl("0" & txtQty.Text)
            CC_Amount = CDbl("0" & txtQty.Text)
            Call Add_rsPayment("CC", cmdCredit.Caption, CC_Amount)
            txtQty.Text = Format(CDbl("0" & txtAmount.Text) - Payment_Totals, "#,##0")
            Call Save_Credit(rsCredit_Payment)
            Call Print_Credit(BillNO)
            Payment_Method = "CC"
            Call Cash
        End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdCredit_Click"
End Sub

Private Sub cmdCashTranfer_Click()
On Error GoTo Handle
        isclick = False
        i = i + CDbl("0" & txtQty.Text)
        CT_Amount = CDbl("0" & txtQty.Text)
        Call Add_rsPayment("CT", cmdCashTranfer.Caption, CT_Amount)
        txtQty.Text = Format(CDbl("0" & txtAmount.Text) - Payment_Totals, "#,##0")
        Payment_Method = "CT"
        Call Cash
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdCashTranfer_Click"
End Sub

Private Sub cmdGiftCard_Click()
On Error GoTo Handle
    With frmGiftCard_Pay
        .Let_Payment = totals - Payment_Totals
        .Show vbModal
        GC_Amount = .Let_Amount
        CardID = .Let_CardID
    End With
    If GC_Amount <> 0 Then
         isclick = False
        Call Add_rsPayment("GC", cmdGiftCard.Caption, GC_Amount)
        txtQty.Text = Format(CDbl("0" & txtAmount.Text) - Payment_Totals, "#,##0")
        Payment_Method = "GC"
        Call Cash
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub cmdmoney_Click(Index As Integer)
On Error GoTo Handle
If isclick = True Then
    txtQty.Text = CDbl("0" & txtQty.Text) + CDbl(cmdmoney(Index).Tag)
Else
    txtQty.Text = CDbl(cmdmoney(Index).Tag)
    isclick = True
End If
    If CDbl("0" & txtQty.Text) >= CDbl("0" & txtAmount.Text) Then
        Call cmdCash_Click
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
Dim ctrl As Control
If isActived = True Then Exit Sub
isActived = True
If cmdCash.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    DescArr = LoadLanguage(LngFile, "#02:003:")
    For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
        Next ctrl
    If UserLevel <> 1 Then CheckRight
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Activate"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        Call cmdCash_Click
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - Form_KeyPress"
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    iCash = False
    isActived = False
    isclick = False
    Payment_Method = "C"
    If rsPayment.State = 0 Then Call Create_Payment
    Call Set_flgPayment
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    txtAmount.Locked = True
    Set rsInvoice_Onhold = OpenCriticalTable("select * from Invoice_OnHold", cnData)
    Set rsInvoice_Total = OpenCriticalTable("Select * from Invoice_Totals", cnData)
    Set rsInvoice_Notes = OpenCriticalTable("select * from Invoice_Totals_Notes", cnData)
    If Not Check_Field_Exist(rsInvoice_Total, "CA_Amount") Then
        cnData.Execute "ALTER TABLE Invoice_Totals ADD COLUMN OA_Amount Double,CA_Amount double, CC_Amount double, ROA_Amount double, GC_Amount double, CT_Amount double "
    End If
    With rsInvoice_Notes
        .Find "Invoice_Number=" & BillNO, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            'Totals = Totals - Get_Reserve_Amount(BillNO)
            
            txtQty.Text = Format(totals + totals * VAT / 100 - Get_Reserve_Amount(BillNO), formatNum)
            txtAmount.Text = Format(txtQty.Text, formatNum)
        End If
    End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   Form_Load"
End Sub

Public Property Get GetCustomer() As Variant
    GetCustomer = Customer
End Property

Public Property Let GetCustomer(ByVal vNewValue As Variant)
    Customer = vNewValue
End Property

Public Property Get GetBillNo() As Variant
    GetBillNo = BillNO
End Property

Public Property Let GetBillNo(ByVal vNewValue As Variant)
   BillNO = vNewValue
End Property

Public Property Let Get_Payment_Method(ByVal vNewValue As Variant)
    Payment_Method = vNewValue
End Property


Public Sub Update_Invoice_Notes()
 On Error GoTo Handle
Dim rsLocation As New ADODB.Recordset
  
    With rsInvoice_Notes
    If .State = 0 Then Set rsInvoice_Notes = Open_Table(cnData, "Invoice_Totals_Notes")
    If .RecordCount > 0 Then .MoveFirst
      .Find "Invoice_Number=" & BillNO, , adSearchForward, adBookmarkFirst
          If Not .EOF Then
            If UCase(Trim(.Fields("ClosingTime"))) = "C" Then
                    .Fields("ClosingTime") = DateDefault & Format(Now, "HH:mm:ss")
                    .Update
                End If
      End If
    End With
 
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Update_Invoice_Notes"
End Sub

Public Function gfUpdate_Invoice_Totals() As Boolean
On Error GoTo Handle
Dim reserve_Value As Double
    gfUpdate_Invoice_Totals = False
    Set rsInvoice_Total = OpenCriticalTable("select * from Invoice_Totals", cnData)
    reserve_Value = Get_Reserve_Amount(BillNO)
        With rsInvoice_Total
            .Find "Invoice_Number=" & BillNO, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                !Store_ID = Store_ID
                !CustNum = CustNo(0)
                !Total_Price = Total
                !Total_Tax1 = totals
                !Grand_Total = totals + totals * VAT / 100
                If OA_Amount = 0 And ROA_Amount = 0 And CT_Amount = 0 And CC_Amount = 0 And GC_Amount = 0 Then
                    !CA_Amount = !Grand_Total - reserve_Value
                Else
                    !CA_Amount = CA_Amount
                End If
                !OA_Amount = OA_Amount
                !ROA_Amount = ROA_Amount
                !CT_Amount = CT_Amount
                !CC_Amount = CC_Amount
                !GC_Amount = GC_Amount
                !Status = Payment_Method
                '!cashier_ID = UserID
                !Amt_Tendered = i 'Payment_Totals
                !Amt_Change = !Amt_Tendered + reserve_Value - !Grand_Total
                !Payment_Method = Payment_Method
                If Get_Cash_by_Time = True Then
                    !DateTime = DateDefault & Format(Now, "HH:mm:ss")
                End If
                !Synchronized = "False"
                rsInvoice_Total.Update
                .Requery
            End If
        End With
gfUpdate_Invoice_Totals = True
Exit Function

Handle:
    MsgBox Err.Number & Err.Description & Me.name & "gfupdate_Invoice_Totals"
    gfUpdate_Invoice_Totals = False
End Function

Public Function gfDelete_Invoice_Onhold() As Boolean
On Error GoTo Handle
    gfDelete_Invoice_Onhold = False
     With rsInvoice_Onhold
     If .State = 1 And .RecordCount > 0 Then
        .MoveFirst
     Else
        Exit Function
     End If
      .Find "Invoice_Number=" & BillNO, , adSearchForward, adBookmarkFirst
          If Not .EOF Then
              .Delete adAffectCurrent
              .Requery
          End If
    End With
    gfDelete_Invoice_Onhold = True
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " gfDelete_Invoice_Onhold"
    gfDelete_Invoice_Onhold = False
End Function


Private Sub Form_Unload(Cancel As Integer)
    diemtichluy = 0
    Customer = ""
   Total = 0
   totals = 0
   OA_Amount = 0
   ROA_Amount = 0
   CA_Amount = 0
   GC_Amount = 0
   CT_Amount = 0
   GC_Amount = 0
   CC_Amount = 0
   i = 0
   Payment_Totals = 0
   CloseRecordset rsPayment
End Sub

Private Sub txtQty_Change()
On Error GoTo Handle
    txtQty.Text = Format(txtQty.Text, "#,##0")
    txtQty.SelStart = Len(txtQty.Text)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtQty_Change"
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    Select Case KeyAscii
        Case 13
            Call cmdCash_Click
        Case 8
        Case 48 To 57
        Case Else:   KeyAscii = 0
    End Select
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtQty_KeyPress "
End Sub
'Luu no vao cong no khach hang
'Tham so truyen vao la ma khach hang

Public Function update_Balance(S As String) As Boolean
On Error GoTo Handle
Dim isUpdate As Boolean
    Dim rsCustomer As New ADODB.Recordset
    Dim strCus As String
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    
    Set rsCustomer = Open_Table(cnData, "Customer")
    With rsCustomer
        If Not .EOF And .RecordCount > 0 Then .MoveFirst
        .Find "CustNum='" & S & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If Val(.Fields("Acct_Balance")) + totals >= CDbl("0" & .Fields("Acct_Max_Balance")) Then
                MsgBox " C«ng nî cña b¹n ®¹t ®Õn møc tèi ®a, vui lßng thanh tãan bít tr­íc khi ghi nî"
                isUpdate = False
            Else
                .Fields("Acct_Balance") = Val(.Fields("Acct_Balance")) + totals
                .Update
                isUpdate = True
            End If
        End If
    End With
    update_Balance = isUpdate
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name
End Function

Public Function Get_Cash_by_Time() As Boolean
On Error GoTo Handle
Dim iOpen As Boolean
    If ArrayFlag(SF(0), 7) = 1 Then iOpen = True
    Get_Cash_by_Time = iOpen
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Get_Cash_by_Time"

End Function

Public Property Get Return_Amt() As Variant
    Return_Amt = CDbl(txtQty.Text)
End Property

Public Property Let Get_Diem(ByVal vNewValue As Variant)
    diemtichluy = vNewValue
End Property

Public Sub update_Diem(strID As String)
    On Error GoTo Handle
    Dim rsCustomer As New ADODB.Recordset
    Set rsCustomer = Open_Table(cnData, "Customer")
    If rsCustomer.RecordCount = 0 Then
        Exit Sub
    Else
        rsCustomer.MoveFirst
    End If
    With rsCustomer
        .Find "Custnum='" & strID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Acct_Balance") = .Fields("Acct_Balance") + OA_Amount
            .Fields("Point") = CDbl("0" & .Fields("Point")) + diemtichluy
            .Update
'            .Requery
        End If
    End With
    
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  update_Diem"
End Sub

Public Sub Set_flgPayment()
    On Error GoTo Handle
        With flgPayment
            .Cols = 2
            .Rows = 7
            .ColWidth(0) = 4400
            .ColWidth(1) = 2600
            .TextMatrix(0, 0) = "H×nh thøc TT"
            .TextMatrix(0, 1) = "Sè tiÒn"
            .ColAlignment(0) = 2
            .ColAlignment(1) = 6
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Set_flgOrder"
End Sub

Public Sub Create_Payment()
 On Error GoTo Handle
        With rsPayment
            If .State = 0 Then
                .Fields.Append "Payment_ID", adVarWChar, 10
                .Fields.Append "Payment_Name", adVarWChar, 100
                .Fields.Append "Amount", adDouble
                .Open
            End If
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Create_Payment"
End Sub

Public Sub Add_rsPayment(ByVal PaymentID As String, Payment_Name As String, ByVal Amount As Double)
On Error GoTo Handle
        With rsPayment
            .Find "Payment_ID='" & PaymentID & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("Payment_ID") = PaymentID
                .Fields("Payment_Name") = Payment_Name
                .Fields("Amount") = Amount
                .Update
            Else
                .Fields("Amount") = .Fields("Amount") + Amount
                .Update
            End If
        End With
        Call SetFLGRID_Payment(rsPayment)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Add_rsPayment"
End Sub
Public Sub SetFLGRID_Payment(rs As ADODB.Recordset)
On Error GoTo Handle
        Dim incount As Integer
        Payment_Totals = 0
        If rs.RecordCount = 0 Then GoTo 1
        rs.MoveFirst
        With rs
            Do While Not .EOF
                incount = incount + 1
                flgPayment.Rows = rs.RecordCount + 1
                With flgPayment
                    .TextMatrix(incount, 0) = rs.Fields(1)
                    .TextMatrix(incount, 1) = Format(rs.Fields(2), formatNum)
                    Payment_Totals = Payment_Totals + CDbl("0" & rs.Fields(2))
                End With
            rs.MoveNext
            Loop
        End With
1:
        If rs.RecordCount = 0 Then
            For incount = 1 To flgPayment.Rows - 1
                With flgPayment
                    .TextMatrix(incount, 0) = ""
                    .TextMatrix(incount, 1) = ""
                End With
            Next
        End If
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "SetFLGRID_Payment"
End Sub

Public Sub Cash()
On Error GoTo Handle
  Dim TC As Double
  Dim Tender As Double
        TC = totals + totals * VAT / 100 - Get_Reserve_Amount(BillNO)
        Tender = i
        Bill = BillNO
        If CDbl("0" & Payment_Totals) >= CDbl("0" & TC) Then 'Val(Totals)
            If gfUpdate_Invoice_Totals = True Then
                    Call Update_Invoice_Notes
                    If gfDelete_Invoice_Onhold = False Then
                        If State_Payment = 1 Then
                            Unload Me
                        End If
                        Exit Sub
                    End If
            Else
                    Exit Sub
            End If
            Call update_Diem(Customer)
            If CardID <> "" Then Call Update_GC_Balance(CardID)
    Unload Me
'         Thoat khoi giao dien ban hang
            If ArrayFlag(SF(4), 6) = 1 Then
                With frmShowBillSale
                    .GetBill = Bill
                    .Show vbModal
                End With
            Else
                With frmChange
                    .GetTotal = TC
                    .GetTender_Amt = Tender
                    .Show vbModal
                End With
                If ArrayFlag(SF(6), 6) = 1 Then Call OpenPrinterCashDraw(GetSettingStr("Receipt", "Receipt_DeviceName", True, myIniFile))
            End If
        End If
Exit Sub
Handle:
        With frmChange
            .GetTotal = TC
            .GetTender_Amt = Tender
            .Show vbModal
        End With
    Exit Sub
    MsgBox Err.Number & Err.Description & Me.name & " - Cash"
End Sub

Public Sub Update_GC_Balance(ByVal Card_ID As String)
On Error GoTo Handle
    Dim rsGiftCard As New ADODB.Recordset
    Set rsGiftCard = Open_Table(cnData, "Gift_Cards")
    With rsGiftCard
        .Find "Card_ID='" & Card_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Balance_Amount") = .Fields("Balance_Amount") - GC_Amount
            If .Fields("Balance_Amount") = 0 Then
                .Fields("Valid") = False
            End If
            .Update
        End If
        
    End With
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " -Update_GC_Balance"
End Sub

Public Sub Save_Credit(rs As ADODB.Recordset)
On Error GoTo Handle
    Dim rsInvoice_Sub_Payment As New ADODB.Recordset
    Set rsInvoice_Sub_Payment = Open_Table(cnData, "Invoice_Sub_Payment")
    With rsInvoice_Sub_Payment
        .Find "Invoice_Number='" & BillNO & "'", , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields("Invoice_Number") = BillNO
            .Fields("Store_ID") = Sec_ID
            .Fields("InvoiceRefNum") = rs.Fields("Transaction_Code")
            .Fields("Banking_Name") = rs.Fields("Banking_Name")
            .Fields("Account_Name") = rs.Fields("Account_Name")
            .Fields("Account_Type") = rs.Fields("Card_Type")
            .Fields("Account_ID") = rs.Fields("Card_Code")
            .Fields("Account_Expire") = rs.Fields("Card_Expired")
            .Fields("Account_Add") = rs.Fields("Account_Add")
            .Fields("Amount") = CC_Amount
            .Update
        Else
            MsgBox "§· tån t¹i Invoice trong D÷ liÖu"
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Save_Credit "
End Sub

Public Sub Print_Credit(BillID As Double)
On Error GoTo Handle
Dim cmd As New ADODB.Command
    Dim SQL As String
            SQL = "SELECT Invoice_Sub_Payment.Invoice_Number, Invoice_Sub_Payment.InvoiceRefNum," & _
            " Invoice_Sub_Payment.Banking_Name, Invoice_Sub_Payment.Account_Name, " & _
            " Invoice_Sub_Payment.Account_Type, Invoice_Sub_Payment.Account_ID, " & _
            " Invoice_Sub_Payment.Account_Expire, Invoice_Sub_Payment.Amount" & _
            " From Invoice_Sub_Payment" & _
            " Where Invoice_Number=" & BillID & _
            " GROUP BY Invoice_Sub_Payment.Invoice_Number, Invoice_Sub_Payment.InvoiceRefNum, " & _
            " Invoice_Sub_Payment.Banking_Name, Invoice_Sub_Payment.Account_Name, " & _
            " Invoice_Sub_Payment.Account_Type, Invoice_Sub_Payment.Account_ID, " & _
            " Invoice_Sub_Payment.Account_Expire, Invoice_Sub_Payment.Amount"
    Set crCredit_Confirm = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crCredit_Confirm
        .Database.AddADOCommand cnData, cmd
        .Transaction.SetUnboundFieldSource "{ado.InvoiceRefNum}"
        .bankingName.SetUnboundFieldSource "{ado.Banking_Name}"
        .CardName.SetUnboundFieldSource "{ado.Account_Name}"
        .CardNo.SetUnboundFieldSource "{ado.Account_ID}"
        .CardType.SetUnboundFieldSource "{ado.Account_Type}"
        .CardExpire.SetUnboundFieldSource "{ado.Account_Expire}"
        .Amount.SetUnboundFieldSource "{ado.Amount}"
    End With
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = crCredit_Confirm
        .Show vbModal
    End With
Exit Sub
Handle:
Exit Sub
    MsgBox Err.Number & Err.Description & Me.name & " Print_Credit"
End Sub


'Public Sub Print_GiftCard(CardID As Double)
'On Error GoTo errHdl
'    Dim SQL As String
'    Dim iReport As CRAXDDRT.Report
'    Dim cmd As New ADODB.Command
'    SQL = "SELECT Gift_Cards.Card_ID, Gift_Cards.Balance, Gift_Cards.Balance_Due, Gift_Cards.Balance_Amount, Gift_Cards.Open_Date, Gift_Cards.Exp_Date" & _
'                " FROM Gift_Cards where Card_ID='" & txtCard_ID.Text & "'"
'
'    Dim CRXReportField As CRAXDDRT.DatabaseFieldDefinition
'        Set crGiftCard = Nothing
'        cmd.ActiveConnection = cnData
'        cmd.CommandText = SQL
'        cmd.Execute
'    With crGiftCard
'        .Database.AddADOCommand cnData, cmd
'        .txtCardID.SetUnboundFieldSource "{ado.Card_ID}"
'        .txtAmount.SetUnboundFieldSource "{ado.Balance_Amount}"
'        .txtAmount2.SetUnboundFieldSource "{ado.Balance_Amount}"
'        .txtDateOpen.SetUnboundFieldSource "{ado.Open_Date}"
'        .txtDateExpired.SetUnboundFieldSource "{ado.Exp_Date}"
'    End With
'    Set iReport = crGiftCard
'    With crvGiftCard
'        .DisplayBorder = False
'        .ReportSource = iReport
'        .EnableSearchControl = False
'        .EnableStopButton = False
'        .EnableGroupTree = False
'        .EnableAnimationCtrl = False
'        .EnablePopupMenu = False
'        .EnableToolbar = False
'        .DisplayToolbar = False
'        .DisplayTabs = False
'        .ToolTipText = ""
'        .ViewReport
'        crvGiftCard.Zoom 100
'        While .IsBusy
'            DoEvents
'        Wend
'        .ShowLastPage
'        While .IsBusy
'            DoEvents
'        Wend
'        .ShowFirstPage
'        While .IsBusy
'            DoEvents
'        Wend
'    End With
'Exit Sub
'errHdl:
'    MsgBox Err.Number & " - " & Err.Description & "Load_GiftCard"
'End Sub


Public Function Get_Reserve_Amount(BillNum As Double) As Double
On Error GoTo Handle
Dim result As Double
    With rsInvoice_Total
        .Find "Invoice_Number=" & BillNum, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            result = CDbl("0" & .Fields("Reserve"))
        Else
            result = 0
        End If
    End With
    Get_Reserve_Amount = result
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Get_Reserve_Amount"

End Function

Public Property Let form_call(ByVal vNewValue As Variant)
    State_Payment = vNewValue
End Property
