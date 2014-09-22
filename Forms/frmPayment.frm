VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPayment 
   Caption         =   "Payment"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPayment.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10455
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   0
      TabIndex        =   30
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
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   630
         Width           =   7035
      End
      Begin MSFlexGridLib.MSFlexGrid flgPayment 
         Height          =   3105
         Left            =   60
         TabIndex        =   32
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
         TabIndex        =   33
         Tag             =   "L2"
         Top             =   150
         Width           =   4395
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
      Left            =   7320
      TabIndex        =   29
      Top             =   510
      Width           =   2925
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
      Left            =   7320
      TabIndex        =   8
      Top             =   1200
      Width           =   4575
      Begin MSForms.CommandButton cmdAlpha 
         Height          =   1095
         Index           =   0
         Left            =   120
         TabIndex        =   28
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
      Begin MSForms.CommandButton cmdAlpha 
         Height          =   1095
         Index           =   1
         Left            =   1560
         TabIndex        =   27
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
         Index           =   2
         Left            =   3000
         TabIndex        =   26
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
         Index           =   3
         Left            =   120
         TabIndex        =   25
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
         Index           =   4
         Left            =   1560
         TabIndex        =   24
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
         Index           =   5
         Left            =   3000
         TabIndex        =   23
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
         Index           =   6
         Left            =   120
         TabIndex        =   22
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
         Index           =   7
         Left            =   1560
         TabIndex        =   21
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
         Index           =   8
         Left            =   3000
         TabIndex        =   20
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
         Index           =   9
         Left            =   120
         TabIndex        =   19
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
         Index           =   11
         Left            =   3000
         TabIndex        =   17
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
      Begin MSForms.CommandButton cmdCash 
         Height          =   1335
         Index           =   0
         Left            =   2280
         TabIndex        =   16
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
      Begin MSForms.CommandButton cmdclose 
         Cancel          =   -1  'True
         Height          =   1335
         Left            =   120
         TabIndex        =   15
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
         Height          =   1095
         Index           =   1
         Left            =   120
         TabIndex        =   14
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
      Begin MSForms.CommandButton cmdCash 
         Height          =   1095
         Index           =   2
         Left            =   2295
         TabIndex        =   13
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
      Begin MSForms.CommandButton cmdCash 
         Height          =   1095
         Index           =   3
         Left            =   120
         TabIndex        =   12
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
      Begin MSForms.CommandButton cmdCash 
         Height          =   1095
         Index           =   4
         Left            =   2295
         TabIndex        =   11
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
      Begin MSForms.CommandButton cmdCash 
         Height          =   1095
         Index           =   5
         Left            =   2295
         TabIndex        =   10
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
      Begin MSForms.CommandButton cmdCash 
         Height          =   1095
         Index           =   6
         Left            =   120
         TabIndex        =   9
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
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1480
      Index           =   1
      Left            =   3600
      Picture         =   "frmPayment.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "50000"
      Top             =   4560
      Width           =   3615
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1480
      Index           =   0
      Left            =   3600
      Picture         =   "frmPayment.frx":5F35
      Style           =   1  'Graphical
      TabIndex        =   6
      Tag             =   "100000"
      Top             =   6000
      Width           =   3615
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1455
      Index           =   2
      Left            =   3600
      Picture         =   "frmPayment.frx":B1C8
      Style           =   1  'Graphical
      TabIndex        =   5
      Tag             =   "200000"
      Top             =   7485
      Width           =   3615
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1450
      Index           =   3
      Left            =   3600
      Picture         =   "frmPayment.frx":10956
      Style           =   1  'Graphical
      TabIndex        =   4
      Tag             =   "500000"
      Top             =   8925
      Width           =   3615
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1450
      Index           =   4
      Left            =   0
      Picture         =   "frmPayment.frx":167CD
      Style           =   1  'Graphical
      TabIndex        =   3
      Tag             =   "20000"
      Top             =   8925
      Width           =   3615
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1455
      Index           =   5
      Left            =   0
      Picture         =   "frmPayment.frx":1FD63
      Style           =   1  'Graphical
      TabIndex        =   2
      Tag             =   "10000"
      Top             =   7485
      Width           =   3615
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1480
      Index           =   6
      Left            =   0
      Picture         =   "frmPayment.frx":28D4B
      Style           =   1  'Graphical
      TabIndex        =   1
      Tag             =   "5000"
      Top             =   6000
      Width           =   3615
   End
   Begin VB.CommandButton cmdmoney 
      Height          =   1480
      Index           =   7
      Left            =   0
      Picture         =   "frmPayment.frx":32930
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "2000"
      Top             =   4560
      Width           =   3615
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
      Left            =   7320
      TabIndex        =   35
      Tag             =   "L1"
      Top             =   120
      Width           =   3525
   End
   Begin MSForms.CommandButton cmdAlpha 
      Height          =   735
      Index           =   12
      Left            =   10320
      TabIndex        =   34
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
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ispment, iclose, isclick As Boolean
Dim Grand_Total As Double
Dim Payment_Method As String
Dim rsInvoice_Totals As New ADODB.Recordset
Dim Invoice_Num As Double
Dim rsInvoice_Notes As New ADODB.Recordset
Dim rsinvoice_hold As New ADODB.Recordset

Public Property Get ispayment() As Variant
    ispayment = ispment
End Property

Public Property Get Is_close() As Variant
    Is_close = iclose
End Property

Public Property Let Get_Grand_Total(ByVal vNewValue As Variant)
    Grand_Total = vNewValue
End Property

Private Sub cmdAlpha_Click(Index As Integer)
Select Case Index
        Case 0 To 11:
            If isclick = True Then
                 txtQty.Text = Format(txtQty.Text & cmdAlpha(Index).Caption, "#,##0")
            Else
                txtQty.Text = Format(txtQty.Text & cmdAlpha(Index).Caption, "#,##0")
                isclick = True
            End If
        Case 12:
            txtQty.Text = ""
       
    End Select
End Sub


Public Function Payment(Method As String, Tender As Double) As Boolean
    On Error GoTo Handle
    Dim paymented As Boolean
    Set rsInvoice_Totals = Open_Table(cnData, "Invoice_Totals")
    With rsInvoice_Totals
    .Find "Invoice_Number=" & Invoice_Num, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Status") = "C"
            .Fields("Amt_Tendered") = Tender
            .Fields("Amt_Change") = Tender - .Fields("Grand_Total")
            .Fields("Payment_Method") = Method
            .Fields("Synchronized") = "False"
            Select Case Method
                Case "C", "CA"
                    .Fields("CA_Amount") = txtAmount.Text
                Case "CT"
                    .Fields("CT_Amount") = txtAmount.Text
                Case "OA"
                    .Fields("OA_Amount") = txtAmount.Text
                Case "CC"
                    .Fields("CC_Amount") = txtAmount.Text
                Case "ROA"
                    .Fields("ROA_Amount") = txtAmount.Text
                Case "OA"
                    .Fields("OA_Amount") = txtAmount.Text
                Case "GC"
                    .Fields("GC_Amount") = txtAmount.Text
            End Select
            .Update
            paymented = True
        End If
    End With
    Payment = paymented
    Exit Function
Handle:
    MsgBox Err.Number & Err.Description & "Payment "
End Function

Private Sub cmdCash_Click(Index As Integer)
On Error GoTo Handle
Dim menthod As String
Dim TC As Double
    Select Case Index
        Case 0
            menthod = "C"
        Case 1
            menthod = "CC"
        Case 2
            menthod = "CT"
        Case 3
            menthod = "ROA"
        Case 4
            menthod = "OA"
        Case 5
            menthod = "GC"
    End Select
    If CDbl("0" & txtQty.Text) = 0 Then txtQty.Text = txtAmount.Text
    Cash (menthod)
     Unload Me
    If ArrayFlag(SF(4), 6) = 1 Then
        With frmShowBillSale
            .GetBill = Invoice_Num
            .Show vbModal
        End With
    Else
        With frmChange
            .GetTotal = TC
            .GetTender_Amt = CDbl("0" & txtQty.Text)
            .Show vbModal
        End With
        If ArrayFlag(SF(6), 6) = 1 Then Call OpenPrinterCashDraw(GetSettingStr("Receipt", "Receipt_DeviceName", True, myIniFile))
        Unload Me
    End If
Exit Sub
Handle:
Exit Sub
MsgBox Err.Number & Err.Description
End Sub

Private Sub cmdClose_Click()
iclose = True
    Unload Me
End Sub

Public Sub Cash(paymenthod As String)
On Error GoTo Handle
If Payment(paymenthod, txtQty.Text) Then
    fUpdate_OnHold (Invoice_Num)
    fUpdate_Invoice_Notes (Invoice_Num)
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description
End Sub

Public Property Let Get_Invoice_Number(ByVal vNewValue As Variant)
Invoice_Num = vNewValue
End Property

Public Sub fUpdate_OnHold(Bill As Double)
On Error GoTo Handle
 Set rsinvoice_hold = Open_Table(cnData, "Invoice_OnHold")
 
    With rsinvoice_hold
    If .RecordCount > 0 Then .MoveFirst
        .Find "Invoice_Number=" & Bill, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Delete adAffectCurrent
            .Requery
        End If
    End With
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description
End Sub

Public Sub fUpdate_Invoice_Notes(Bill As Double)
On Error GoTo Handle
    Set rsInvoice_Notes = Open_Table(cnData, "Invoice_Totals_Notes")
    With rsInvoice_Notes
    If .RecordCount > 0 Then .MoveFirst
        .Find "Invoice_Number=" & Bill, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("ClosingTime") = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") & Format(Now, "HH:mm:ss")
            .Update
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description
End Sub

Private Sub cmdmoney_Click(Index As Integer)
On Error GoTo Handle
If isclick = True Then
    txtQty.Text = CDbl("0" & txtQty.Text) + CDbl(cmdmoney(Index).Tag)
Else
    txtQty.Text = CDbl(cmdmoney(Index).Tag)
    isclick = True
End If
    If CDbl("0" & txtQty.Text) = 0 Then txtQty.Text = txtAmount.Text
    If CDbl(txtQty.Text) >= CDbl(txtAmount.Text) Then
    Cash ("C")
     Unload Me
    If ArrayFlag(SF(4), 6) = 1 Then
        With frmShowBillSale
            .GetBill = Invoice_Num
            .Show vbModal
        End With
    Else
        With frmChange
            .GetTotal = TC
            .GetTender_Amt = CDbl("0" & txtQty.Text)
            .Show vbModal
        End With
        If ArrayFlag(SF(6), 6) = 1 Then Call OpenPrinterCashDraw(GetSettingStr("Receipt", "Receipt_DeviceName", True, myIniFile))
        Unload Me
        End If
    End If
Exit Sub
Handle:
Exit Sub
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
Dim ctrl As Control
If isActived = True Then Exit Sub
isActived = True
If cmdCash(0).Font.name <> CurFont Then Call Set_Language(Me, CurFont)
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

Private Sub Form_Load()
On Error GoTo Handle
    txtAmount.Text = Format(Grand_Total, "#,##0")
    txtQty.Text = Format(Grand_Total, "#,##0")
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    isclick = False
    iclose = False
    ispment = False
    CloseRecordset rsInvoice_Totals
    CloseRecordset rsInvoice_Notes
    CloseRecordset rsinvoice_hold
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
                  cmdCash(0).Enabled = False
            Else: cmdCash(0).Enabled = True
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
