VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmGiftCard 
   Caption         =   "PhiÕu quµ tÆng"
   ClientHeight    =   10110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
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
   ScaleHeight     =   10110
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin CRVIEWERLibCtl.CRViewer crvGiftCard 
      CausesValidation=   0   'False
      Height          =   4095
      Left            =   120
      TabIndex        =   24
      Top             =   4080
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
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   8040
      TabIndex        =   20
      Top             =   -120
      Width           =   7215
      Begin VB.OptionButton optValid 
         Caption         =   "§· sö dông/HÕt h¹n"
         Height          =   255
         Index           =   2
         Left            =   4680
         TabIndex        =   23
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optValid 
         Caption         =   "Cßn h¹n sö  dông"
         Height          =   255
         Index           =   1
         Left            =   2160
         TabIndex        =   22
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton optValid 
         Caption         =   "TÊt c¶"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   240
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Danh s¸ch phiÕu  quµ tÆng"
      Height          =   9495
      Left            =   8040
      TabIndex        =   15
      Top             =   480
      Width           =   7215
      Begin MSFlexGridLib.MSFlexGrid flgGiftCard 
         Height          =   9135
         Left            =   45
         TabIndex        =   19
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   16113
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame frmCmd 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   360
      TabIndex        =   14
      Top             =   8160
      Width           =   7380
      Begin prjTouchScreen.MyButton cmdAdd 
         Height          =   945
         Left            =   60
         TabIndex        =   5
         Tag             =   "L5"
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1667
         BTYPE           =   14
         TX              =   "&Thªm"
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
         MICON           =   "frmGiftCard.frx":0000
         PICN            =   "frmGiftCard.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdUpdate 
         Height          =   945
         Left            =   1500
         TabIndex        =   6
         Tag             =   "L6"
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1667
         BTYPE           =   14
         TX              =   "&CËp nhËt"
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
         MICON           =   "frmGiftCard.frx":046E
         PICN            =   "frmGiftCard.frx":048A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdDelete 
         Height          =   945
         Left            =   2850
         TabIndex        =   7
         Tag             =   "L7"
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1667
         BTYPE           =   14
         TX              =   "&Xãa"
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
         MICON           =   "frmGiftCard.frx":09CE
         PICN            =   "frmGiftCard.frx":09EA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdHelp 
         Height          =   945
         Left            =   4200
         TabIndex        =   9
         Tag             =   "L8"
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1667
         BTYPE           =   14
         TX              =   "In PhiÕu"
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
         MICON           =   "frmGiftCard.frx":1024
         PICN            =   "frmGiftCard.frx":1040
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   945
         Left            =   5640
         TabIndex        =   8
         Tag             =   "L9"
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1667
         BTYPE           =   14
         TX              =   "&§ãng"
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
         MICON           =   "frmGiftCard.frx":37F2
         PICN            =   "frmGiftCard.frx":380E
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
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7935
      Begin VB.TextBox txtBalance_Due 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         TabIndex        =   17
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtCard_ID 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1080
         MaxLength       =   12
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox txtBalance 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1080
         TabIndex        =   2
         Text            =   " "
         Top             =   1560
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker dtpOpenDate 
         Height          =   450
         Left            =   1080
         TabIndex        =   3
         Top             =   2580
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   794
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   62914561
         UpDown          =   -1  'True
         CurrentDate     =   39254
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker dtpExpire_Date 
         Height          =   450
         Left            =   4800
         TabIndex        =   4
         Top             =   2535
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   794
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "MM/dd/yyyy"
         Format          =   62914561
         UpDown          =   -1  'True
         CurrentDate     =   39254
      End
      Begin VB.Label Label2 
         Caption         =   "Tµi kho¶n kh¶ dông:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   18
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lblCard_ID 
         Alignment       =   1  'Right Justify
         Caption         =   "M· thÎ:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   45
         TabIndex        =   13
         Tag             =   "L1"
         Top             =   330
         Width           =   1200
      End
      Begin VB.Label lblBalance 
         Alignment       =   1  'Right Justify
         Caption         =   "Tµi kho¶n:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Tag             =   "L2"
         Top             =   1200
         Width           =   1230
      End
      Begin VB.Label lblOpenDate 
         Alignment       =   1  'Right Justify
         Caption         =   "Ngµy ph¸t hµnh"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Tag             =   "L3"
         Top             =   2160
         Width           =   1965
      End
      Begin VB.Label lblExpire_Date 
         Alignment       =   1  'Right Justify
         Caption         =   "Ngµy hÕt h¹n:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4050
         TabIndex        =   10
         Tag             =   "L4"
         Top             =   2160
         Width           =   1995
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Danh s¸ch phiÕu quµ tÆng"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   720
      TabIndex        =   16
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmGiftCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsGiftCard As New ADODB.Recordset
    Dim iReport As CRAXDDRT.Report


Private Sub cmdAdd_Click()
    Call Lock4Items(False)
    Call Init_4_Items
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = True
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Set rsGiftCard = Nothing
End Sub

Private Sub cmdDelete_Click()
On Error GoTo Handle
    If MsgBox("B¹n cã ch¾c ch¾n muèn xãa thÎ quµ tÆng nµy ?", vbYesNo) = vbYes Then
        With rsGiftCard
            .Find "Card_ID='" & Trim(txtCard_ID.Text) & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Delete adAffectCurrent
            End If
        End With
    End If
    Call Delete_Receipt
    Call optValid_Click(0)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdDelete_Click"
End Sub

Private Sub cmdPrint()
    
End Sub

Private Sub cmdHelp_Click()
On Error GoTo errHdl
    myPrint iReport, crvGiftCard.GetCurrentPageNumber, 1

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo Handle
    With rsGiftCard
        .Find "Card_ID='" & Trim(txtCard_ID.Text) & "'", , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields("Card_ID") = Trim(txtCard_ID.Text)
            .Fields("Balance") = CDbl("0" & txtBalance.Text)
            .Fields("Balance_Due") = CDbl("0" & txtBalance_Due.Text)
            .Fields("Balance_Amount") = CDbl("0" & txtBalance_Due.Text)
            .Fields("CashID") = UserID
            .Fields("Open_Date") = gfCONVERT_DATE_TO_STRING(dtpOpenDate.Value)
            .Fields("Exp_Date") = gfCONVERT_DATE_TO_STRING(dtpExpire_Date.Value)
            .Fields("Valid") = 1
            .Update
            .Requery
        Else
            .Fields("Card_ID") = Trim(txtCard_ID.Text)
            .Fields("Balance") = CDbl("0" & txtBalance.Text)
            .Fields("Balance_Due") = CDbl("0" & txtBalance_Due.Text)
            .Fields("Balance_Amount") = CDbl("0" & txtBalance_Due.Text)
            .Fields("Open_Date") = gfCONVERT_DATE_TO_STRING(dtpOpenDate.Value)
            .Fields("Exp_Date") = gfCONVERT_DATE_TO_STRING(dtpOpenDate.Value)
            .Update
            .Requery
        End If
    End With
    Call optValid_Click(0)
    Call Save_Receipt
    cmdUpdate.Enabled = False
    cmdAdd.Enabled = True
    cmdAdd.SetFocus
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdUpdate_Click "
End Sub


Private Sub dtpOpenDate_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then dtpExpire_Date.SetFocus
End Sub

Private Sub flgGiftCard_Click()
On Error GoTo errHdl
    txtCard_ID.Text = flgGiftCard.TextMatrix(flgGiftCard.Row, 0)
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- flgGiftCard_Click"
End Sub

Private Sub Form_Activate()
On Error GoTo errHdl
    Dim DescArr() As String
    Dim ctrl As Control
    DescArr = LoadLanguage(LngFile, "#03:016:")
'    If cmdAdd.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(10)
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Activate"

End Sub

Private Sub Form_Load()
On Error GoTo Handle
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsGiftCard = Open_Table(cnData, "Gift_Cards")
    Call Lock4Items(True)
    Call Lock_Command(True)
    Call Set_Data_Text
    Call Set_flg
    Call SetFLGRID_Gift(rsGiftCard)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub

Public Sub Lock4Items(b As Boolean)
On Error GoTo Handle
    txtCard_ID.Locked = b
    txtBalance.Locked = b
    dtpOpenDate.Enabled = Not b
    dtpExpire_Date.Enabled = Not b
    txtBalance_Due.Locked = b
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Lock4Items "
End Sub

Public Sub Lock_Command(b As Boolean)
On Error GoTo Handle
    cmdAdd.Enabled = b
    cmdUpdate.Enabled = Not b
    cmdDelete.Enabled = b
    cmdHelp.Enabled = b
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Lock_Command "
End Sub

Public Sub Init_4_Items()
    On Error GoTo Handle
    txtCard_ID.Text = ""
    txtBalance.Text = ""
    txtBalance_Due.Text = ""
    dtpOpenDate.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    dtpExpire_Date.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    txtCard_ID.SetFocus
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Lock4Items "
End Sub


Private Sub optValid_Click(Index As Integer)
    On Error GoTo Handle
        If optValid(0).Value = True Then
            Set rsGiftCard = OpenCriticalTable("select * from Gift_Cards ", cnData)
        ElseIf optValid(1).Value = True Then
            Set rsGiftCard = OpenCriticalTable("select * from Gift_Cards where Valid=1", cnData)
        Else
            Set rsGiftCard = OpenCriticalTable("select * from Gift_Cards where Valid=0", cnData)
        End If
        Call SetFLGRID_Gift(rsGiftCard)
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " optValid_Click"
End Sub

Private Sub txtBalance_Due_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtBalance_Due.Text = Format(txtBalance_Due.Text, formatNum)
    End If
End Sub

Private Sub txtBalance_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtBalance_Due.SetFocus
        txtBalance.Text = Format(txtBalance.Text, formatNum)
    End If
End Sub

Private Sub txtCard_ID_Change()
    On Error GoTo Handle
    With rsGiftCard
        .Find "Card_ID='" & Trim(txtCard_ID.Text) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtBalance.Text = Format(.Fields("BALANCE"), formatNum)
            txtBalance_Due.Text = Format(.Fields("BALANCE_Due"), formatNum)
            dtpOpenDate.Value = gfCONVERT_STRING_TO_DATE(.Fields("Open_Date"))
            dtpExpire_Date.Value = gfCONVERT_STRING_TO_DATE(.Fields("Exp_Date"))
            Call Load_GiftCard(txtCard_ID.Text)

        End If
    End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtCard_ID_Change"
End Sub

Private Sub txtCard_ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtBalance.SetFocus
End Sub

Public Sub Set_Data_Text()
If rsGiftCard.State = 1 And rsGiftCard.RecordCount > 0 Then
    rsGiftCard.MoveFirst
Else
    Exit Sub
End If
With rsGiftCard
    txtCard_ID.Text = .Fields("Card_ID")
    txtBalance.Text = Format(.Fields("Balance"), formatNum)
    txtBalance_Due.Text = Format(.Fields("Balance_Due"), formatNum)
    dtpOpenDate.Value = gfCONVERT_STRING_TO_DATE(.Fields("Open_Date"))
    dtpExpire_Date.Value = gfCONVERT_STRING_TO_DATE(.Fields("EXP_DATE"))
End With
End Sub

Public Sub Set_flg()
    On Error GoTo Handle
        With flgGiftCard
            .Cols = 5
            .Rows = 20
            .ColWidth(0) = 1200
            .ColWidth(1) = 1400
            .ColWidth(2) = 1400
            .ColWidth(3) = 1550
            .ColWidth(4) = 1550
            .TextMatrix(0, 0) = "M· sè"
            .TextMatrix(0, 1) = "Tµi kho¶n"
            .TextMatrix(0, 2) = "TK kh¶ dông"
            .TextMatrix(0, 3) = "Ngµy ph¸t hµnh"
            .TextMatrix(0, 4) = "H¹n sö dông"
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
    MsgBox Err.Number & Err.Description & Me.name & "  flgGiftCard"
End Sub

Public Sub SetFLGRID_Gift(rs As ADODB.Recordset)
On Error GoTo Handle
        Dim incount As Integer
        If rs.RecordCount = 0 Then GoTo 1
        rs.MoveFirst
        With rs
            .Sort = "Card_ID ASC"
            Do While Not .EOF
                incount = incount + 1
                flgGiftCard.Rows = rs.RecordCount + 1
                With flgGiftCard
                    .TextMatrix(incount, 0) = rs.Fields(0)
                    .TextMatrix(incount, 1) = Format(rs.Fields(1), formatNum)
                    .TextMatrix(incount, 2) = Format(rs.Fields(2), formatNum)
                    .TextMatrix(incount, 3) = gfCONVERT_STRING_TO_DATE(rs.Fields(5))
                    .TextMatrix(incount, 4) = gfCONVERT_STRING_TO_DATE(rs.Fields(6))
                End With
            rs.MoveNext
            Loop
        End With
1:
        If rs.RecordCount = 0 Then
            For incount = 1 To flgGiftCard.Rows - 1
                With flgGiftCard
                    .TextMatrix(incount, 0) = ""
                    .TextMatrix(incount, 1) = ""
                    .TextMatrix(incount, 2) = ""
                    .TextMatrix(incount, 3) = ""
                    .TextMatrix(incount, 4) = ""
                End With
            Next
        End If
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "SetFLGRIDORDER"
End Sub


Public Sub Load_GiftCard(CardID As String)
On Error GoTo errHdl
    Dim SQL As String
    Dim cmd As New ADODB.Command
    SQL = "SELECT Gift_Cards.Card_ID, Gift_Cards.Balance, Gift_Cards.Balance_Due,Gift_Cards.Balance_Amount,  Gift_Cards.Open_Date, Gift_Cards.Exp_Date" & _
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


Public Sub Save_Receipt()
On Error GoTo Handle
Dim rsPhieuthu As New ADODB.Recordset
Set rsPhieuthu = Open_Table(cnData, "Income")
With rsPhieuthu
        .Find "ID='" & "PQT\" & txtCard_ID.Text & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
              .addNew
              .Fields("ID") = "PQT\" & Trim(txtCard_ID.Text)
              .Fields("Store_ID") = Store_ID
              .Fields("Cashier_ID") = UserID
              .Fields("DateTime") = gfCONVERT_DATE_TO_STRING(dtpOpenDate.Value)
              .Fields("Receipt_ID") = "BH"
              .Fields("Customer_ID") = "101"
              .Fields("Reciever_Name") = ""
              .Fields("Division") = ""
              .Fields("Payment_Method") = ""
              .Fields("Amount") = CDbl("0" & txtBalance.Text)
              .Fields("Description") = "Thu tiÒn phiÕu quµ tÆng " & txtCard_ID.Text
              .Update
            End If
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name
End Sub


Public Sub Delete_Receipt()
On Error GoTo Handle
Dim rsPhieuthu As New ADODB.Recordset
Set rsPhieuthu = Open_Table(cnData, "Income")
With rsPhieuthu
        .Find "ID='" & "PQT\" & txtCard_ID.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
             .Delete adAffectCurrent
             .Requery
            End If
        End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Delete_Receipt"
End Sub
