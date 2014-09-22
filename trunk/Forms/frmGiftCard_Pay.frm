VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmGiftCard_Pay 
   Caption         =   "ThÎ quµ tÆng"
   ClientHeight    =   6735
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAmount 
      Height          =   495
      Left            =   7080
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   8175
      Begin CRVIEWERLibCtl.CRViewer crvGiftCard 
         CausesValidation=   0   'False
         Height          =   4095
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   7935
         DisplayGroupTree=   0   'False
         DisplayToolbar  =   -1  'True
         EnableGroupTree =   -1  'True
         EnableNavigationControls=   -1  'True
         EnableStopButton=   -1  'True
         EnablePrintButton=   -1  'True
         EnableZoomControl=   0   'False
         EnableCloseButton=   -1  'True
         EnableProgressControl=   -1  'True
         EnableSearchControl=   -1  'True
         EnableRefreshButton=   -1  'True
         EnableDrillDown =   -1  'True
         EnableAnimationControl=   -1  'True
         EnableSelectExpertButton=   0   'False
         EnableToolbar   =   0   'False
         DisplayBorder   =   0   'False
         DisplayTabs     =   0   'False
         DisplayBackgroundEdge=   0   'False
         SelectionFormula=   ""
         EnablePopupMenu =   -1  'True
         EnableExportButton=   0   'False
         EnableSearchExpertButton=   0   'False
         EnableHelpButton=   0   'False
      End
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   1095
      Left            =   4800
      TabIndex        =   5
      Top             =   5520
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1931
      BTYPE           =   3
      TX              =   "§ãng"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmGiftCard_Pay.frx":0000
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
      Height          =   1095
      Left            =   1320
      TabIndex        =   4
      Top             =   5520
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1931
      BTYPE           =   3
      TX              =   "§ång ý"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16761024
      BCOLO           =   16761024
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmGiftCard_Pay.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2880
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   1800
      Width           =   3855
   End
   Begin VB.TextBox txtCard_ID 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Sè tiÒn:"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "M· thÎ:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmGiftCard_Pay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isOKCancel As Boolean
Dim GiftCard_Amount As Double
Dim Pay_Amount As Double
Dim rsGiftCard As New ADODB.Recordset
Dim CardID As String
Private Sub cmdCancel_Click()
    sOKCancel = False
    If isOKCancel Then
        CardID = ""
        If txtAmount.Text > Pay_Amount Then
            GiftCard_Amount = Pay_Amount
        Else
            GiftCard_Amount = txtAmount.Text
        End If
    Else
        GiftCard_Amount = 0
    End If
    Unload Me
End Sub

Private Sub cmdOK_Click()
    isOKCancel = True
    If isOKCancel Then
        CardID = txtCard_ID.Text
        If txtAmount.Text > Pay_Amount Then
            GiftCard_Amount = Pay_Amount
        Else
            GiftCard_Amount = txtAmount.Text
        End If
    Else
        GiftCard_Amount = 0
    End If
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Handle
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    Set rsGiftCard = Open_Table(cnData, "Gift_Cards")
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - Form_Load"
End Sub

Public Property Get Let_Amount() As Variant
    Let_Amount = GiftCard_Amount
End Property

Public Sub Load_GiftCard(CardID As String)
On Error GoTo errHdl
    Dim SQL As String
    Dim iReport As CRAXDDRT.Report
    Dim cmd As New ADODB.Command
    SQL = "SELECT Gift_Cards.Card_ID, Gift_Cards.Balance, Gift_Cards.Balance_Due, Gift_Cards.Balance_Amount, Gift_Cards.Open_Date, Gift_Cards.Exp_Date" & _
                " FROM Gift_Cards where Card_ID='" & txtCard_ID.Text & "'"

    Dim CRXReportField As CRAXDDRT.DatabaseFieldDefinition
     
        Set crGiftCard = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crGiftCard
        .Database.AddADOCommand cnData, cmd
        .txtCardID.SetUnboundFieldSource "{ado.Card_ID}"
        .txtAmount.SetUnboundFieldSource "{ado.Balance_Amount}"
        .txtAmount2.SetUnboundFieldSource "{ado.Balance_Amount}"
        .txtDateOpen.SetUnboundFieldSource "{ado.Open_Date}"
        .txtDateExpired.SetUnboundFieldSource "{ado.Exp_Date}"
    End With
    Set iReport = crGiftCard
    With crvGiftCard
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
        crvGiftCard.Zoom 100
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
    MsgBox Err.Number & " - " & Err.Description & "Load_GiftCard"
End Sub

Public Property Get Let_CardID() As Variant
    Let_CardID = CardID
End Property


Private Sub txtCard_ID_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        With rsGiftCard
            .Find "Card_ID='" & txtCard_ID.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If gfCONVERT_DATE_TO_STRING(Format(Now, "dd/MM/yyyy")) >= .Fields("Open_Date") And gfCONVERT_DATE_TO_STRING(Format(Now, "dd/MM/yyyy")) <= .Fields("Exp_Date") And .Fields("Valid") = True Then
                    txtAmount.Text = CDbl("0" & .Fields("Balance_Amount"))
                    Call Load_GiftCard(txtCard_ID.Text)
                Else
                    txtAmount.Text = 0
                    MsgBox "PhiÕu nµy kh«ng cã gi¸ trÞ trong kho¶ng thêi gian nµy hoÆc hÕt h¹n sö dông", vbInformation
                End If
           End If
        End With
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtCard_ID_KeyPress "
End Sub


Public Property Let Let_Payment(ByVal vNewValue As Variant)
    Pay_Amount = vNewValue
End Property
