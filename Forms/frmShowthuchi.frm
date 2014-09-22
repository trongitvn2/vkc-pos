VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmShowthuchi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShowthuchi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picToolsBar 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   705
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   11805
      TabIndex        =   0
      Top             =   0
      Width           =   11835
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   585
         Left            =   7740
         TabIndex        =   10
         Top             =   60
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1032
         BTYPE           =   14
         TX              =   "&ß„ng"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
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
         MICON           =   "frmShowthuchi.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdExport 
         Height          =   525
         Left            =   5730
         TabIndex        =   9
         Top             =   60
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   926
         BTYPE           =   14
         TX              =   "Xu t sang dπng kh∏c"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   9.75
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
         MICON           =   "frmShowthuchi.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   510
         Left            =   4740
         Picture         =   "frmShowthuchi.frx":0044
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   75
         Width           =   885
      End
      Begin VB.TextBox txtPage 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2820
         TabIndex        =   6
         Top             =   75
         Width           =   855
      End
      Begin VB.CommandButton cmdFirst 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1980
         Picture         =   "frmShowthuchi.frx":072E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.CommandButton cmdPrevious 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2400
         Picture         =   "frmShowthuchi.frx":0A70
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.CommandButton cmdNext 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3720
         Picture         =   "frmShowthuchi.frx":0DB2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.CommandButton cmdLast 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4140
         Picture         =   "frmShowthuchi.frx":10F4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   405
      End
      Begin VB.ComboBox cboZoom 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Text            =   "cboZoom"
         Top             =   60
         Width           =   1575
      End
   End
   Begin CRVIEWERLibCtl.CRViewer crvReport 
      CausesValidation=   0   'False
      Height          =   6105
      Left            =   375
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   840
      Width           =   9105
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   0   'False
      EnableGroupTree =   0   'False
      EnableNavigationControls=   0   'False
      EnableStopButton=   0   'False
      EnablePrintButton=   0   'False
      EnableZoomControl=   0   'False
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   0   'False
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   0   'False
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmShowthuchi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim iReport As CRAXDDRT.Report
Dim TotalRptPage As Integer
Dim mRsExcel As ADODB.Recordset
Dim mStrNameExcel   As String
Dim crView As CRAXDDRT.CRPaperOrientation

Private Sub cboZoom_Change()
On Error GoTo errHdl

    Select Case cboZoom.ListIndex
    Case 0
        crvReport.Zoom 1
    Case 1
        crvReport.Zoom 2
    Case 2
        crvReport.Zoom 400
    Case 3
        crvReport.Zoom 300
    Case 4
        crvReport.Zoom 200
    Case 5
        crvReport.Zoom 150
    Case 6
        crvReport.Zoom 100
    Case 7
        crvReport.Zoom 75
    Case 8
        crvReport.Zoom 50
    Case 9
        crvReport.Zoom 25
    Case Else
        If IsNumeric(cboZoom.Text) Then
            If Val(cboZoom.Text) < 1000 And Val(cboZoom.Text) > 10 Then crvReport.Zoom CInt(cboZoom.Text)
        End If
    End Select

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cboZoom_Click()
On Error GoTo errHdl
    Call cboZoom_Change

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cboZoom_KeyPress(KeyAscii As Integer)
On Error GoTo errHdl

    If KeyAscii < &H30 Or KeyAscii > &H39 Then
        If KeyAscii <> &H8 Then KeyAscii = 0
    End If
    
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdClose_Click()
On Error GoTo errHdl
    Unload Me
    
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdExport_Click()
On Error GoTo errHdl

    'ExportReport iReport
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdFirst_Click()
On Error GoTo errHdl

    crvReport.ShowFirstPage
    While crvReport.IsBusy
        DoEvents
    Wend
    txtPage.Text = crvReport.GetCurrentPageNumber & " / " & TotalRptPage

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdLast_Click()
On Error GoTo errHdl

   crvReport.ShowLastPage

    While crvReport.IsBusy
        DoEvents
    Wend
    txtPage.Text = crvReport.GetCurrentPageNumber & " / " & TotalRptPage

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdNext_Click()
On Error GoTo errHdl

  crvReport.ShowNextPage
    
    While crvReport.IsBusy
        DoEvents
    Wend
    txtPage.Text = crvReport.GetCurrentPageNumber & " / " & TotalRptPage

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdPrevious_Click()
On Error GoTo errHdl

    crvReport.ShowPreviousPage
    While crvReport.IsBusy
        DoEvents
    Wend
    txtPage.Text = crvReport.GetCurrentPageNumber & " / " & TotalRptPage

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdPrint_Click()
On Error GoTo errHdl

    myPrint iReport, crvReport.GetCurrentPageNumber, TotalRptPage

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo errHdl

    If KeyCode = vbKeyP And Shift = vbCtrlMask Then
        Call cmdPrint_Click
    End If

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo errHdl

    Dim DescArr() As String
    DescArr = LoadLanguage(LngFile, "#02:004:")
    'Me.Caption = frmViewExternalReport.lstReport.list(frmViewExternalReport.lstReport.ListIndex)
    cmdExport.Caption = DescArr(1)
    cmdClose.Caption = DescArr(2)
    With cboZoom
        .AddItem DescArr(3), 0
        .AddItem DescArr(4), 1
        .AddItem "400 %", 2
        .AddItem "300 %", 3
        .AddItem "200 %", 4
        .AddItem "150 %", 5
        .AddItem "100 %", 6
        .AddItem "75 %", 7
        .AddItem "50 %", 8
        .AddItem "25 %", 9
    End With
    With iReport
        .PaperSize = crPaperA4
        .PaperOrientation = crPortrait
    End With
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
        While .IsBusy
            DoEvents
        Wend
        .ShowLastPage
        While .IsBusy
            DoEvents
        Wend
        TotalRptPage = .GetCurrentPageNumber
        .ShowFirstPage
        While .IsBusy
            DoEvents
        Wend
        txtPage.Text = .GetCurrentPageNumber & " / " & TotalRptPage
    End With
    Me.WindowState = 2
    
    cboZoom.ListIndex = 0
    Screen.MousePointer = vbDefault

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub Form_Resize()
On Error GoTo errHdl

    With crvReport
        .Left = 0
        .top = picToolsBar.top + picToolsBar.Height + 120
        .Height = Me.ScaleHeight - (picToolsBar.top + picToolsBar.Height + 120)
        .Width = Me.ScaleWidth
    End With
    picToolsBar.Left = 0
    picToolsBar.Width = Me.ScaleWidth

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub

Public Property Let Report(ByVal vReport As CRAXDDRT.Report)
On Error GoTo errHdl

    Set iReport = vReport

Exit Property
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Property


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo errHdl

    Set iReport = Nothing
    Set crPhieuchi = Nothing
    Set crPhieuthu = Nothing
Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description
End Sub


