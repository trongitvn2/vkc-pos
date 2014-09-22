VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPreviewBill 
   Caption         =   "Danh s¸ch hãa ®¬n ®· thanh to¸n"
   ClientHeight    =   11055
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
   ScaleHeight     =   11055
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdSearch 
      Height          =   855
      Left            =   8040
      TabIndex        =   22
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "T×m &kiÕm"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   14.25
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
      MICON           =   "frmPreviewBill.frx":0000
      PICN            =   "frmPreviewBill.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Frame Frame5 
      Caption         =   "NhËp d÷ liÖu cÇn t×m"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   7935
      Begin VB.TextBox txtTableID 
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   5520
         TabIndex        =   23
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox txtInvoice 
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   1800
         TabIndex        =   0
         Top             =   390
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Sè bµn:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4560
         TabIndex        =   24
         Top             =   450
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Sè H§:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   960
         TabIndex        =   21
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame frapreview 
      Caption         =   "Hßa ®¬n b¸n hµng"
      Height          =   11415
      Left            =   10680
      TabIndex        =   13
      Top             =   120
      Width           =   5055
      Begin CRVIEWERLibCtl.CRViewer crvBill 
         CausesValidation=   0   'False
         Height          =   10575
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   4335
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
   Begin VB.Frame Frame1 
      Caption         =   "T×m Hãa ®¬n theo kho¶ng thêi gian"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   6255
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   6015
         Begin VB.OptionButton Opt75 
            Caption         =   "Khæ in 75mm"
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
            Left            =   3360
            TabIndex        =   18
            Top             =   180
            Width           =   1935
         End
         Begin VB.OptionButton Opt80 
            Caption         =   "Khæ in 80mm"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   480
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   495
         Left            =   1260
         TabIndex        =   7
         Top             =   420
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
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
         CalendarForeColor=   255
         CalendarTitleForeColor=   255
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   64290817
         UpDown          =   -1  'True
         CurrentDate     =   39448
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   495
         Left            =   4560
         TabIndex        =   8
         Top             =   420
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
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
         CalendarForeColor=   255
         CalendarTitleForeColor=   255
         Format          =   64290817
         UpDown          =   -1  'True
         CurrentDate     =   39448
      End
      Begin VB.Label lblDenngay 
         Alignment       =   1  'Right Justify
         Caption         =   "§Õn ngµy:"
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
         Left            =   3240
         TabIndex        =   10
         Tag             =   "L3"
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label lblFromdate 
         Alignment       =   1  'Right Justify
         Caption         =   "Tõ ngµy:"
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
         TabIndex        =   9
         Tag             =   "L2"
         Top             =   480
         Width           =   1005
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1850
      Left            =   6360
      TabIndex        =   3
      Top             =   960
      Width           =   4215
      Begin prjTouchScreen.MyButton cmddelete 
         Height          =   855
         Left            =   2760
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1508
         BTYPE           =   6
         TX              =   "&Xãa"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPreviewBill.frx":0656
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdExit 
         Cancel          =   -1  'True
         Height          =   855
         Left            =   2760
         TabIndex        =   4
         Tag             =   "L7"
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1508
         BTYPE           =   6
         TX              =   "Th&o¸t"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPreviewBill.frx":0672
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdEdit 
         Height          =   855
         Left            =   1320
         TabIndex        =   5
         Tag             =   "L5"
         Top             =   120
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1508
         BTYPE           =   6
         TX              =   "&Söa ch÷a"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPreviewBill.frx":068E
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
         Height          =   855
         Left            =   1320
         TabIndex        =   11
         Top             =   960
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1508
         BTYPE           =   6
         TX              =   "XuÊt sang d¹ng kh¸c"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPreviewBill.frx":06AA
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
         Left            =   0
         TabIndex        =   15
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1508
         BTYPE           =   6
         TX              =   "&In"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPreviewBill.frx":06C6
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
         Height          =   855
         Left            =   0
         TabIndex        =   19
         Tag             =   "L4"
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1508
         BTYPE           =   6
         TX              =   "&Thùc thi"
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
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPreviewBill.frx":06E2
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
   Begin VB.Frame Frame4 
      Height          =   8415
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   10575
      Begin MSFlexGridLib.MSFlexGrid flgPreviewBill 
         Height          =   7695
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   13573
         _Version        =   393216
         BackColorFixed  =   -2147483643
         BackColorBkg    =   -2147483643
         GridColorFixed  =   12632256
         WordWrap        =   -1  'True
         ScrollTrack     =   -1  'True
         GridLinesFixed  =   1
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
End
Attribute VB_Name = "frmPreviewBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBill As New ADODB.Recordset
Dim DescArr() As String
Dim DescArr1() As String
Dim Location, Table As String
Dim Bill_Number As String
Dim discount As Integer
Dim iReport As New CRAXDDRT.Report
Dim isLoaded As Boolean


 
Public Sub SelectRight()
On Error GoTo Handle
    If UserLevel <> 1 Then
        cmdEdit.Visible = False
'        cmdDelete.Visible = False
        With cmdDone
            .Left = 0
            .Width = 4200
            .Height = 855
        End With
    End If
    'If UserID = "131112" Then cmdDelete.Visible = True
    If UserLevel = 1 Then cmdDelete.Visible = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " SelectRight"
End Sub

Private Sub cmddelete_Click()
On Error GoTo Handle
    
'   Call Open_File
'    Print #fFile, "Hñy Bill " & vbTab & Now & vbTab & flgPreviewBill.TextMatrix(flgPreviewBill.Row, 1) & vbTab & flgPreviewBill.TextMatrix(flgPreviewBill.Row, 2) & vbTab & flgPreviewBill.TextMatrix(flgPreviewBill.Row, 3) & vbTab & userName
'    Close #fFile
    If MsgBox("B¹n cã ch¾c ch¾n muèn xãa hãa ®¬n nµy, xãa bá kh«ng phôc håi l¹i ®­îc !", vbYesNo) = vbYes Then
        Call Delete_Invoice(CDbl("0" & Bill_Number), cnData)
        Call cmdDone_Click
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmddelete_Click"
End Sub

Private Sub cmdDone_Click()
On Error GoTo Handle
Dim strSql As String
        If UserLevel = 1 Then
            strSql = "select Invoice_Totals.DateTime,Invoice_Totals.Invoice_Number,Invoice_Totals.Orig_OnHoldID," & _
                    " Invoice_Totals.Total_Price,Invoice_Totals.Discount,Invoice_Totals.Grand_Total," & _
                    " Invoice_Totals.CA_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.Station_ID " & _
                    " from Invoice_Totals " & _
                    " where left(DateTime,8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "'" & _
                    " and left(DateTime,8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'" & _
                    " and Status<>'O' and Status<>'P' and Status<>'CO' order by Invoice_Totals.Invoice_number"
        Else
            strSql = "select Invoice_Totals.DateTime,Invoice_Totals.Invoice_Number,Invoice_Totals.Orig_OnHoldID," & _
                    " Invoice_Totals.Total_Price,Invoice_Totals.Discount,Invoice_Totals.Grand_Total," & _
                    " Invoice_Totals.CA_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.Station_ID " & _
                    " from Invoice_Totals " & _
                    " where left(DateTime,8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "'" & _
                    " and left(DateTime,8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'" & _
                    " and Status<>'O' and Status<>'P' and Status<>'CO' and Cashier_ID='" & UserID & "' order by Invoice_Totals.Invoice_number"
        End If
    Set rsBill = OpenCriticalTable(strSql, cnData)
    Call SetFlexPreviewBill(rsBill)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdDone_Click"
End Sub

Private Sub cmdEdit_Click()
Dim rsOrdered As New ADODB.Recordset
On Error GoTo Handle
If Bill_Number = "" Then Exit Sub
    Set rsOrdered = Let_Record_Ordered(Bill_Number)
    With frmOrder
        .FormCall = 1
        .Get_Secion = Trim(Location)
        .Get_Record_Ordered = rsOrdered
        .GetBill_Number = Bill_Number
        .Get_Table_ID = Trim(Table)
        .cmdNewBalance.Enabled = False
        .Get_Discount = discount
        .Show vbModal
    End With
    'Print #fFile, "Söa Bill " & vbTab & Now & vbTab & flgPreviewBill.TextMatrix(flgPreviewBill.Row, 1) & vbTab & flgPreviewBill.TextMatrix(flgPreviewBill.Row, 2) & vbTab & flgPreviewBill.TextMatrix(flgPreviewBill.Row, 3) & vbTab & UserID & ":" & userName
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdEdit_Click"
End Sub

Private Sub cmdExit_Click()
    Set rsBill = Nothing
    Unload Me
End Sub

Private Sub cmdHelp_Click()
On Error GoTo Handle
    iReport.Export
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " --- cmd_Export_to_other format"

End Sub

Private Sub cmdPrint_Click()
On Error GoTo Handle
Dim PrinterName As String
        With frmSelectPrint
            .Show vbModal
            PrinterName = .LetPrinter
        End With
        If PrinterName = "" Then Exit Sub
        iReport.SelectPrinter True, PrinterName, Printer.Port
        iReport.PrintOut False
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub cmdSearch_Click()
On Error GoTo Handle
Bill_Number = 0
    If txtInvoice.Text = "" Then
        If txtTableID.Text = "" Then
            MsgBox "NhËp d÷ liÖu cÇn t×m", vbInformation
        Else
            Call Load_Table
        End If
        Exit Sub
    Else
        Call Load_Bill(CDbl(txtInvoice.Text))
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdSearch_Click"
End Sub

Private Sub flgPreviewBill_Click()
On Error GoTo Handle
    Location = flgPreviewBill.TextMatrix(flgPreviewBill.Row, 8)
    Bill_Number = flgPreviewBill.TextMatrix(flgPreviewBill.Row, 9)
    Table = flgPreviewBill.TextMatrix(flgPreviewBill.Row, 2)
    discount = CDbl("0" & flgPreviewBill.TextMatrix(flgPreviewBill.Row, 4))
    If Bill_Number <> "" Then
        Call Load_Bill(CDbl(Bill_Number))
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " flgPreviewBill_Click"
End Sub

Private Sub flgPreviewBill_DblClick()
    If UserLevel = 1 Then
        Call cmdEdit_Click
    Else
        With frmViewBill
            .GetBill = Bill_Number
            .Show vbModal
        End With
    End If
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    Dim ctrl As Control
    Me.Caption = DescArr(8)
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
 Next ctrl
 If UserLevel <> 1 Then Check_right
 If isLoaded Then cmdDone_Click
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    DescArr = LoadLanguage(LngFile, "#03:007:")
    DescArr1 = LoadLanguage(LngFile, "#02:005:")
    Call SelectRight
    dtpFromDate.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    dtpToDate.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
    Call cmdDone_Click
    If UserLevel <> 1 Then
        dtpFromDate.Enabled = False
        dtpToDate.Enabled = False
    End If
    isLoaded = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub
Public Sub SetFlexPreviewBill(rs As ADODB.Recordset)
On Error GoTo errHdl
    Dim intCount    As Integer
    With flgPreviewBill
        .Cols = 10
        .Rows = 2
        .Font = ".vnArial"
        .ColWidth(0) = 1300
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 1400
        .ColWidth(4) = 1200
        .ColWidth(5) = 1300
        .ColWidth(6) = 1300
        .ColWidth(7) = 1700
        .ColWidth(8) = 1400
        .ColWidth(9) = 0
        .TextMatrix(0, 0) = DescArr(10)
        .TextMatrix(0, 1) = DescArr(9)
        .TextMatrix(0, 2) = DescArr(11)
        .TextMatrix(0, 3) = DescArr(12)
        .TextMatrix(0, 4) = DescArr(13)
        .TextMatrix(0, 5) = DescArr(14)
        .TextMatrix(0, 6) = DescArr(15)
        .TextMatrix(0, 7) = DescArr(16)
        .TextMatrix(0, 8) = DescArr(17)
    End With
    
    If rs Is Nothing Then Exit Sub
    If rs.State = 0 Then Exit Sub
    
    If rs.EOF And rs.BOF Then
        With flgPreviewBill
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
            .TextMatrix(1, 4) = ""
            .TextMatrix(1, 5) = ""
            .TextMatrix(1, 6) = ""
            .TextMatrix(1, 7) = ""
            .TextMatrix(1, 8) = ""
             .TextMatrix(1, 9) = ""
        End With
        Exit Sub
    End If
   flgPreviewBill.Rows = rs.RecordCount + 1
    intCount = 0
    Do While Not rs.EOF
        intCount = intCount + 1
        flgPreviewBill.TextMatrix(intCount, 0) = gfCONVERT_STRING_TO_DATE(Left(rs!DateTime, 8))
        flgPreviewBill.TextMatrix(intCount, 1) = Right("0000" & rs!Invoice_Number, 4)
        flgPreviewBill.TextMatrix(intCount, 2) = rs!Orig_OnHoldID
        flgPreviewBill.TextMatrix(intCount, 3) = Format(rs!Total_Price, formatNum)
        flgPreviewBill.TextMatrix(intCount, 4) = rs!discount
        flgPreviewBill.TextMatrix(intCount, 5) = Format(rs!Grand_Total, formatNum)
        flgPreviewBill.TextMatrix(intCount, 6) = Format(rs!CA_Amount, formatNum)
        flgPreviewBill.TextMatrix(intCount, 7) = Format(rs!CT_Amount, formatNum)
        flgPreviewBill.TextMatrix(intCount, 8) = rs!Station_ID
        flgPreviewBill.TextMatrix(intCount, 9) = rs!Invoice_Number
        rs.MoveNext
    Loop
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - flgPreviewBill "
End Sub

Private Sub Form_Unload(Cancel As Integer)
Bill_Number = ""
Table = ""
currentBill = ""
discount = 0
isLoaded = False
End Sub

Public Sub Delete_Invoice(S As Double, cn As ADODB.Connection)
On Error GoTo Handle
Dim rsOnHold As New ADODB.Recordset
Dim rsInvoice_Notes As New ADODB.Recordset
Dim rsInvoice_Total As New ADODB.Recordset
Dim rsInvoice_Personal As New ADODB.Recordset

'Dim rsInvoice_Notes As New ADODB.Recordset
Call Add_to_Cancel_Invoice(S)
Set rsOnHold = Open_Table(cnData, "Invoice_OnHold")
Set rsInvoice_Notes = Open_Table(cnData, "Invoice_Totals_Notes")
Set rsInvoice_Total = Open_Table(cnData, "Invoice_Totals")

cn.Execute "delete from Invoice_Itemized where Invoice_Number=" & S

cn.Execute "delete  from Invoice_Totals_Person_Mapping where Invoice_Number=" & S
'Xoa invoice tam trong Invoice_Total
    With rsInvoice_Total
        .Find "Invoice_Number=" & S, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Delete adAffectCurrent
            .Requery
        End If
    End With
    
'Xoa invoice tam trong Invoice_Notes
    With rsInvoice_Notes
        .Find "Invoice_Number=" & S, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Delete adAffectCurrent
            .Requery
        End If
    End With
'Xoa Ban tam trong Table_OnHold
    With rsOnHold
        .Find "Invoice_Number=" & S, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Delete adAffectCurrent
            .Requery
        End If
    End With
    cnData.Execute "Delete  from Items_Deleted where Invoice_Num=" & S
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Public Sub Load_Bill(Bill_No As Double)
    On Error Resume Next
    Dim cmd As New ADODB.Command
    Dim SQL As String
    Dim RptID As Integer
    Dim ReceiptReport As CRAXDDRT.Report
    Dim rs As New ADODB.Recordset
    If ArrayFlag(SF(0), 5) = 0 Then
        If ArrayFlag(SF(6), 2) = 0 Then
        SQL = " SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.CustNum, Invoice_Totals.Discount," & _
            " Invoice_Totals.Total_Price, Left(Invoice_Totals_Notes.OpenTime,8) AS DateIn, substring(Invoice_Totals_Notes.OpenTime,9,8) AS TimeIn,Left(Invoice_Totals_Notes.ClosingTime,8) AS DateOut, substring(Invoice_Totals_Notes.ClosingTime,9,8) AS TimeOut, " & _
            " Invoice_Totals.Service_Charge, Invoice_Totals.VATFee, Invoice_Totals.Adjustment1," & _
            " Invoice_Totals.Adj1Rate,Invoice_Totals.Personals, Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate," & _
            " Invoice_Totals.Adjustment3,Invoice_Totals.Adj3Rate, Invoice_Totals.Adjustment4,Invoice_Totals.Adj4Rate, Invoice_Totals.AddMoney," & _
            " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount," & _
            " Invoice_Totals.Grand_Total," & _
            " Invoice_Totals.Amt_Tendered, Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, " & _
            " Invoice_Totals.OrderMan, Invoice_Totals.Station_ID, Invoice_Totals.Payment_Method,Invoice_Totals.Tax_Rate_ID, Invoice_Totals.Reserve," & _
            " Invoice_Itemized.ItemNum,Invoice_Itemized.LineNum,Invoice_Itemized.Line_Disc_Desc,Invoice_Itemized.LineDisc, Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer, Sum(Invoice_Itemized.Amt) AS Amt, Invoice_Itemized.DiffItemName, Invoice_Totals.Orig_OnHoldID" & _
            " FROM ((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN (Inventory INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID) ON Invoice_Itemized.ItemNum = Inventory.ItemNum) INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number " & _
            " Where (((Invoice_Totals.Invoice_Number) = " & Bill_No & "))" & _
            " GROUP BY Invoice_Totals.Invoice_Number, Invoice_Totals.CustNum, Invoice_Totals.Discount," & _
            " Invoice_Totals.Tax_Rate_ID, Invoice_Totals.Total_Price, " & _
            " Invoice_Totals.Service_Charge, Invoice_Totals.VATFee, Invoice_Totals.Adjustment1,Invoice_Totals.Adj1Rate," & _
            " Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate, Invoice_Totals.Adjustment3,Invoice_Totals.Adj3Rate,Invoice_Totals.Adj4Rate," & _
            " Invoice_Totals.Adjustment4, Invoice_Totals.AddMoney, Invoice_Totals.OrderMan," & _
            " Left(Invoice_Totals_Notes.OpenTime,8),substring(Invoice_Totals_Notes.OpenTime,9,8), Left(Invoice_Totals_Notes.ClosingTime,8),substring(Invoice_Totals_Notes.ClosingTime,9,8),Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered,Invoice_Totals.Reserve," & _
            " Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID," & _
            " Invoice_Totals.Payment_Method,Invoice_Totals.Personals, Invoice_Itemized.ItemNum, Invoice_Itemized.PricePer," & _
            " Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc,Invoice_Itemized.LineNum,Invoice_Itemized.Line_Disc_Desc," & _
            " Invoice_Totals.Orig_OnHoldID, Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate," & _
            " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount" & _
            " ORDER BY Invoice_Itemized.LineNum Desc"
        Else
        SQL = " SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.CustNum, Invoice_Totals.Discount," & _
            " Invoice_Totals.Total_Price, Left(Invoice_Totals_Notes.OpenTime,8) AS DateIn, substring(Invoice_Totals_Notes.OpenTime,9,8) AS TimeIn, Left(Invoice_Totals_Notes.ClosingTime,8) AS DateOut, substring(Invoice_Totals_Notes.ClosingTime,9,8) AS TimeOut," & _
            " Invoice_Totals.Service_Charge, Invoice_Totals.VATFee, Invoice_Totals.Adjustment1,Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate," & _
            " Invoice_Totals.Adj1Rate,Invoice_Totals.Personals, Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate, " & _
            " Invoice_Totals.Adjustment3, Invoice_Totals.Adj3Rate,Invoice_Totals.Adjustment4,Invoice_Totals.Adj4Rate, Invoice_Totals.AddMoney,Invoice_Totals.Reserve," & _
             " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount," & _
            " Invoice_Totals.Grand_Total," & _
            " Invoice_Totals.Amt_Tendered, Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, " & _
            " Invoice_Totals.OrderMan, Invoice_Totals.Station_ID, Invoice_Totals.Payment_Method,Invoice_Totals.Tax_Rate_ID, " & _
            " Invoice_Itemized.ItemNum,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc, Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer, Sum(Invoice_Itemized.Amt) AS Amt, Invoice_Itemized.DiffItemName, Invoice_Totals.Orig_OnHoldID" & _
            " FROM ((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN (Inventory INNER JOIN Departments ON Inventory.Dept_ID = Departments.Dept_ID) ON Invoice_Itemized.ItemNum = Inventory.ItemNum) INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number " & _
            " Where (((Invoice_Totals.Invoice_Number) = " & Bill_No & "))" & _
            " GROUP BY Invoice_Totals.Invoice_Number, Invoice_Totals.CustNum, Invoice_Totals.Discount," & _
            " Invoice_Totals.Total_Price,  " & _
            " Invoice_Totals.Service_Charge, Invoice_Totals.VATFee, Invoice_Totals.Adjustment1,Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate," & _
            " Invoice_Totals.Adj1Rate, Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate,Invoice_Totals.Reserve," & _
            " Invoice_Totals.Adjustment3,Invoice_Totals.Adj3Rate, Invoice_Totals.Adjustment4,Invoice_Totals.Adj4Rate, Invoice_Totals.AddMoney,Invoice_Totals.Tax_Rate_ID," & _
            " Invoice_Totals.OrderMan,Invoice_Totals.Personals, substring(Invoice_Totals_Notes.ClosingTime,9,8), Invoice_Totals.Grand_Total," & _
            " Invoice_Totals.Amt_Tendered, Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID," & _
            " Invoice_Totals.Station_ID, Invoice_Totals.Payment_Method, Invoice_Itemized.ItemNum," & _
            " Invoice_Itemized.PricePer, Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc," & _
            " Invoice_Totals.Orig_OnHoldID, Left(Invoice_Totals_Notes.OpenTime,8),substring(Invoice_Totals_Notes.OpenTime,9,8), Left(Invoice_Totals_Notes.ClosingTime,8),substring(Invoice_Totals_Notes.ClosingTime,9,8)," & _
             " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount" & _
            " ORDER BY Invoice_Itemized.ItemNum ASC"
        End If
    Else
        If ArrayFlag(SF(6), 1) = 0 Then

            SQL = "SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.CustNum," & _
                " Invoice_Totals.Discount, Invoice_Totals.Total_Price,Left(Invoice_Totals_Notes.OpenTime,8) AS DateIn, substring(Invoice_Totals_Notes.OpenTime,9,8) AS TimeIn,Left(Invoice_Totals_Notes.ClosingTime,8) AS DateOut, substring(Invoice_Totals_Notes.ClosingTime,9,8) AS TimeOut," & _
                " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adjustment1,Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate," & _
                " Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate,Invoice_Totals.Adjustment3,Invoice_Totals.Adj3Rate,Invoice_Totals.Tax_Rate_ID," & _
                " Invoice_Totals.Adjustment4,Invoice_Totals.Adj4Rate,Invoice_Totals.AddMoney," & _
                " Invoice_Totals.Grand_Total,Invoice_Totals.Personals, Invoice_Totals.Amt_Tendered,Invoice_Totals.Reserve," & _
               " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount," & _
                " Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, Invoice_Totals.OrderMan," & _
                " Invoice_Totals.Station_ID,Invoice_Totals.Payment_Method,Invoice_Itemized.ItemNum, " & _
                " Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer,Invoice_Itemized.LineNum,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc," & _
                " sum(Invoice_Itemized.Amt) as Amt, " & _
                " Invoice_Itemized.DiffItemName ,Invoice_Totals.Orig_OnHoldID,MainGroup.GroupNo " & _
                " FROM((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN " & _
                " (Inventory INNER JOIN (Departments INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo) ON Inventory.Dept_ID = Departments.Dept_ID) ON Invoice_Itemized.ItemNum = Inventory.ItemNum)" & _
                " INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number " & _
                " Where Invoice_Totals.Invoice_Number=" & Bill_No & _
                " group by Invoice_Itemized.LineNum,Invoice_Itemized.ItemNum,Invoice_Totals.Invoice_Number," & _
                " Invoice_Totals.CustNum,Invoice_Totals.Discount," & _
                " Invoice_Totals.Total_Price,Invoice_Totals.Grand_Total,Invoice_Totals.Amt_Tendered,Invoice_Totals.Amt_Change,substring(Invoice_Totals_Notes.ClosingTime,9,8)," & _
                " Invoice_Totals.Cashier_ID, Invoice_Totals.OrderMan, Invoice_Totals.Station_ID,Invoice_Totals.Reserve, Left(Invoice_Totals_Notes.OpenTime,8),substring(Invoice_Totals_Notes.OpenTime,9,8), Left(Invoice_Totals_Notes.ClosingTime,8),substring(Invoice_Totals_Notes.ClosingTime,9,8)," & _
                " Invoice_Itemized.PricePer,Invoice_Totals.Personals, Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc ,Invoice_Totals.Orig_OnHoldID,Invoice_Totals.Payment_Method, " & _
                " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Tax_Rate_ID,Invoice_Totals.Adjustment1,Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate,Invoice_Totals.Adjustment3," & _
                " Invoice_Totals.Adj3Rate,Invoice_Totals.Adjustment4,Invoice_Totals.Adj4Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate,Invoice_Totals.AddMoney,MainGroup.GroupNo, " & _
                " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount" & _
                " order by Invoice_Itemized.LineNum Desc"
        Else
            SQL = "SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.CustNum," & _
                " Invoice_Totals.Discount, Invoice_Totals.Total_Price,Left(Invoice_Totals_Notes.OpenTime,8) AS DateIn, substring(Invoice_Totals_Notes.OpenTime,9,8) AS TimeIn,Left(Invoice_Totals_Notes.ClosingTime,8) AS DateOut, substring(Invoice_Totals_Notes.ClosingTime,9,8) AS TimeOut," & _
                " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adjustment1,Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate," & _
                " Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate,Invoice_Totals.Adjustment3,Invoice_Totals.Reserve," & _
                " Invoice_Totals.Adj3Rate,Invoice_Totals.Adjustment4,Invoice_Totals.Adj4Rate,Invoice_Totals.AddMoney,Invoice_Totals.Grand_Total,Invoice_Totals.Personals, Invoice_Totals.Amt_Tendered," & _
                " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount," & _
                " Invoice_Totals.Amt_Change, Invoice_Totals.Cashier_ID, Invoice_Totals.OrderMan,Invoice_Totals.Tax_Rate_ID," & _
                " Invoice_Totals.Station_ID,Invoice_Totals.Payment_Method,Invoice_Itemized.ItemNum, " & _
                " Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc," & _
                " sum(Invoice_Itemized.Amt) as Amt, " & _
                " Invoice_Itemized.DiffItemName ,Invoice_Totals.Orig_OnHoldID,MainGroup.GroupNo " & _
                " FROM((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN " & _
                " (Inventory INNER JOIN (Departments INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo) ON Inventory.Dept_ID = Departments.Dept_ID) ON Invoice_Itemized.ItemNum = Inventory.ItemNum)" & _
                " INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number " & _
                " Where Invoice_Totals.Invoice_Number=" & Bill_No & _
                " group by Invoice_Itemized.ItemNum,Invoice_Totals.Invoice_Number,Invoice_Totals.CustNum,Invoice_Totals.Discount," & _
                " Invoice_Totals.Total_Price,Invoice_Totals.Grand_Total, Left(Invoice_Totals_Notes.OpenTime,8),substring(Invoice_Totals_Notes.OpenTime,9,8), Left(Invoice_Totals_Notes.ClosingTime,8),substring(Invoice_Totals_Notes.ClosingTime,9,8)," & _
                " Invoice_Totals.Amt_Tendered,Invoice_Totals.Amt_Change,substring(Invoice_Totals_Notes.ClosingTime,9,8)," & _
                " Invoice_Totals.Cashier_ID,Invoice_Totals.Personals, Invoice_Totals.OrderMan, Invoice_Totals.Station_ID,Invoice_Totals.Tax_Rate_ID," & _
                " Invoice_Itemized.PricePer, Invoice_Itemized.DiffItemName,Invoice_Itemized.LineDisc ," & _
                " Invoice_Itemized.Line_Disc_Desc,Invoice_Totals.Orig_OnHoldID,Invoice_Totals.Payment_Method, Invoice_Totals.Reserve," & _
                " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adjustment1,Invoice_Totals.Adj1Rate,Invoice_Totals.Adjustment2,Invoice_Totals.Adj2Rate,Invoice_Totals.Adjustment3," & _
                " Invoice_Totals.Adj3Rate,Invoice_Totals.Adjustment4,Invoice_Totals.Adj4Rate,Invoice_Totals.Adjustment5,Invoice_Totals.Adj5Rate, Invoice_Totals.Adjustment6,Invoice_Totals.Adj6Rate,Invoice_Totals.AddMoney,MainGroup.GroupNo, " & _
                " Invoice_Totals.CA_Amount,Invoice_Totals.OA_Amount,Invoice_Totals.ROA_Amount,Invoice_Totals.CC_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.GC_Amount" & _
                " order by Invoice_Itemized.ItemNum ASC"
        End If
   End If
    Set rs = OpenCriticalTable(SQL, cnData)
    If rs.State <> 0 Then
        If rs.RecordCount = 0 Then
            MsgBox "Kh«ng t×m thÊy sè Bill theo yªu cÇu"
            Exit Sub
        Else
            If Val(Bill_Number & "0") = 0 Then Bill_Number = Trim(txtInvoice.Text)
        End If
    Else
        Exit Sub
    End If
    
    
    Dim CRXReportField As CRAXDDRT.DatabaseFieldDefinition
     
    Set crSaleBill = Nothing
    Set crSaleBill58 = Nothing
    Set crSaleBill75 = Nothing
    Set crSaleBillA5 = Nothing
    If ReceiptType = "80" Then
        Set ReceiptReport = crSaleBill
    ElseIf ReceiptType = "58" Then
        Set ReceiptReport = crSaleBill58
    ElseIf ReceiptType = "75" Then
        Set ReceiptReport = crSaleBill75
    Else
        Set ReceiptReport = crSaleBillA5
    End If
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
        
    With ReceiptReport
        .Database.AddADOCommand cnData, cmd
        .txtPluCode.SetUnboundFieldSource "{ado.DiffItemName}"
        .txtPluName.SetUnboundFieldSource "{ado.ItemNum}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtCost.SetUnboundFieldSource "{ado.PricePer}"
        .txtAmt.SetUnboundFieldSource "{ado.Amt}"
        .LineDisc.SetUnboundFieldSource "{ado.LineDisc}"
        .txtBillNo.SetUnboundFieldSource "{ado.Invoice_Number}"
        .txtCashier.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtNVPV.SetUnboundFieldSource "{ado.OrderMan}"
        .txtPayment.SetUnboundFieldSource "{ado.Amt_Tendered}"
        .txtChange.SetUnboundFieldSource "{ado.Amt_Change}"
        .txtDiscount.SetUnboundFieldSource "{ado.Discount}"
        .txtTable.SetUnboundFieldSource "{ado.Orig_OnHoldID}"
        .txtserver.SetUnboundFieldSource "{ado.Station_ID}"
        .txtMethod.SetUnboundFieldSource "{ado.Payment_Method}"
        '.TxtTotal.SetUnboundFieldSource "{ado.Grand_Total}"
        .txtServ.SetUnboundFieldSource "{ado.Service_Charge}"
        .txtVAT.SetUnboundFieldSource "{ado.VATFee}"
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
        
        .txtMoney.SetUnboundFieldSource "{ado.AddMoney}"
        .txtDateIn.SetUnboundFieldSource "{ado.DateIn}"
        .txtTimeIn.SetUnboundFieldSource "{ado.TimeIn}"
        .txtDateOut.SetUnboundFieldSource "{ado.DateOut}"
        .txtTimeOut.SetUnboundFieldSource "{ado.TimeOut}"
        .txtCustomer.SetUnboundFieldSource "{ado.CustNum}"
        .txtMixmatch.SetUnboundFieldSource "{ado.Tax_Rate_ID}"
        .txtsokhach.SetUnboundFieldSource "{ado.Personals}"
        .txtLineDiscDesc.SetUnboundFieldSource "{ado.Line_Disc_Desc}"
        .txtCAPAYMENT.SetUnboundFieldSource "{ado.CA_Amount}"
        .txtOAPAYMENT.SetUnboundFieldSource "{ado.OA_Amount}"
        .txtROAPAYMENT.SetUnboundFieldSource "{ado.ROA_Amount}"
        .txtCCPAYMENT.SetUnboundFieldSource "{ado.CC_Amount}"
        .txtCTPAYMENT.SetUnboundFieldSource "{ado.CT_Amount}"
        .txtGCPAYMENT.SetUnboundFieldSource "{ado.GC_Amount}"
        .txtDatcoc.SetUnboundFieldSource "{ado.Reserve}"
        If ArrayFlag(SF(0), 5) = 1 Then
            .txtMaingroup.SetUnboundFieldSource "{ado.GroupNo}"
        End If
        
        .lblTitle.SetText DescArr1(1)
        .lblBillNo.SetText DescArr1(2)
        .lblTable.SetText DescArr1(3)
        .lblItems.SetText DescArr1(4)
        .lblQty.SetText DescArr1(5)
        .lblPrice.SetText DescArr1(6)
        .lblAmt.SetText DescArr1(7)
        .lblTotal.SetText DescArr1(8)
'        .lblDiscount.SetText DescArr1(9)
        .lblTender.SetText DescArr1(10)
        .lblChange.SetText DescArr1(11)
        .lblRead.SetText DescArr1(12)
        .lblCashier.SetText DescArr1(13)
        .lblPhuthu.SetText DescArr1(14)
        .lblTotal1.SetText DescArr1(15)
        .lblServer.SetText DescArr1(16)
        .lblIn.SetText DescArr1(17)
        .lblOut.SetText DescArr1(18)
        .lblCash.SetText DescArr1(19)
        .lblOrder.SetText DescArr1(20)
        .lblCustome.SetText DescArr1(21)
        .lblSignal.SetText DescArr1(22)
'        .lblAdj1.SetText DescArr1(25)
'        .lblAdj2.SetText DescArr1(26)
        .lblPhuphi.SetText DescArr1(27)
        .lblVAT.SetText DescArr1(29)
        '.lblPrintCount.SetText DescArr1(30)
        .lblTotalItems.SetText DescArr1(31)
        
        'canh le
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign
        
        With .txtMoney
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
        With .txtAdj2
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
        With .txtAdj4
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
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
        With .txtPayment
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
        With .txtAmtDist
            .DecimalPlaces = DecimalAmtNumber
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
        With .txtServAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        
        With .SumMaingroup
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    
    Set iReport = ReceiptReport
    With crvBill
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
'Exit Sub
'errHdl:
'    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub MyButton1_Click()

End Sub



Private Sub Opt75_Click()
On Error GoTo Handle
    Opt75.Value = True
    Opt80.Value = False
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Opt75_Click"
End Sub

Private Sub Opt80_Click()
On Error GoTo Handle
    Opt75.Value = False
    Opt80.Value = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Opt80_Click"
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdSearch_Click
End Sub

Public Sub Load_Table()
On Error GoTo Handle
    Dim rsTable_Find As New ADODB.Recordset
    Dim strSql As String
    strSql = "select Invoice_Totals.DateTime,Invoice_Totals.Invoice_Number,Invoice_Totals.Orig_OnHoldID," & _
                    " Invoice_Totals.Total_Price,Invoice_Totals.Discount,Invoice_Totals.Grand_Total," & _
                    " Invoice_Totals.CA_Amount,Invoice_Totals.CT_Amount,Invoice_Totals.Station_ID " & _
                    " from Invoice_Totals  " & _
                    " WHERE left(DateTime,8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "'" & _
                    " and left(DateTime,8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'" & _
                    " and Status<>'O' and Status<>'P' and Status<>'CO' and Invoice_Totals.Orig_OnHoldID='" & txtTableID.Text & Chr(13) & "'" & _
                    " ORDER by Invoice_Totals.Invoice_number"
    Set rsTable_Find = OpenCriticalTable(strSql, cnData)
    Call SetFlexPreviewBill(rsTable_Find)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Load_Table"
End Sub

Public Sub Check_right()
 Dim res As New ADODB.Recordset
        
        Set res = LoadPasswordData
        With MyRight
            res.MoveFirst
            Do While Not res.EOF
                If StrComp(res.Fields("UserName"), userName, 1) = 0 Then
                    .FullRight = res.Fields("UserRight")
                    .Nhanvien = RightDeCode(Mid(.FullRight, 193, 64))
                    Exit Do
                End If
                res.MoveNext
            Loop
            If Mid(.Nhanvien, 12, 1) = 0 Then
                  dtpFromDate.Enabled = False
                  dtpToDate.Enabled = False
            Else
                dtpFromDate.Enabled = True
                dtpToDate.Enabled = True
            End If
        End With
End Sub

Public Sub Add_to_Cancel_Invoice(Invoice As Double)
On Error GoTo Handle
    Dim rsInvoice_Totals As New ADODB.Recordset
    Dim strOrg As String
    Dim DateDelete As String
    Dim rsinvoice_Cancel As New ADODB.Recordset
    Dim rsInvoice_Cancel_Items As New ADODB.Recordset
    If cnData.State = adStateClosed Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    strOrg = "SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.Station_ID, Invoice_Totals.Store_ID, Invoice_Totals.Cashier_ID, Invoice_Totals.Orig_OnHoldID, " & _
                    " Invoice_Totals.DateTime, Invoice_Totals.Status, Invoice_Itemized.LineNum, Invoice_Itemized.ItemNum, Invoice_Itemized.Quantity," & _
                    " Invoice_Itemized.PricePer , Invoice_Itemized.Line_Disc_Desc, Invoice_Itemized.DiffItemName,Invoice_Itemized.Kit_Description" & _
                    " FROM Invoice_Totals Left JOIN Invoice_Itemized ON Invoice_Totals.[Invoice_Number] = Invoice_Itemized.[Invoice_Number]" & _
                    " where Invoice_Totals.Invoice_Number=" & Invoice
    Set rsInvoice_Totals = OpenCriticalTable(strOrg, cnData)
    Set rsinvoice_Cancel = Open_Table(cnData, "Invoice_Cancel")
    Set rsInvoice_Cancel_Items = Open_Table(cnData, "Invoice_Cancel_Items")
    
    With rsInvoice_Totals
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            'cap nhat vao Invoice_Cancel
            With rsinvoice_Cancel
                DateDelete = Left(rsInvoice_Totals.Fields("DateTime"), 8)
                .Find "Invoice_number=" & DateDelete & rsInvoice_Totals.Fields("Invoice_Number"), , adSearchForward, adBookmarkFirst
                If .EOF Then
                    .addNew
                    .Fields("Invoice_Number") = DateDelete & rsInvoice_Totals.Fields("Invoice_number")
                    .Fields("DateTime") = Trim(rsInvoice_Totals.Fields("DateTime"))
                    .Fields("Staion_ID") = Trim(rsInvoice_Totals.Fields("Station_ID"))
                    .Fields("Table_ID") = Trim(rsInvoice_Totals.Fields("Orig_OnHoldID"))
                    .Fields("Cashier_ID") = Trim(rsInvoice_Totals.Fields("Cashier_ID"))
                    .Fields("Cashier_Cancel") = UserID
                    .Fields("CO_DateTime") = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") & Format(Now, "HH:mm:ss")
                    .Fields("Invoice_Status") = Trim(rsInvoice_Totals.Fields("Status"))
                    .Update
                End If
            End With
            'cap nhat vao Invoice_Cancel_Items
            With rsInvoice_Cancel_Items
                    .addNew
                    .Fields("Invoice_Number") = DateDelete & rsInvoice_Totals.Fields("Invoice_number")
                    .Fields("LineNum") = rsInvoice_Totals.Fields("LineNum")
                    .Fields("ItemNum") = Trim(rsInvoice_Totals.Fields("ItemNum"))
                    .Fields("ItemName") = Trim(rsInvoice_Totals.Fields("DiffItemName"))
                    .Fields("Quantity") = rsInvoice_Totals.Fields("Quantity")
                    .Fields("Price") = rsInvoice_Totals.Fields("PricePer")
                    .Fields("Item_CO_DateTime") = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") & Format(Now, "HH:mm:ss")
                    .Fields("Kit_Desc") = Trim(rsInvoice_Totals.Fields("Kit_Description"))
                    .Update
            End With
        .MoveNext
        Loop
    End With
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name
End Sub
