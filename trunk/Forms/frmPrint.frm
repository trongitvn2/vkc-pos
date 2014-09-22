VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "M∏y in"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   645
      Left            =   4500
      TabIndex        =   18
      Tag             =   "L13"
      Top             =   4200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1138
      BTYPE           =   14
      TX              =   "&Tho∏t"
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
      MICON           =   "frmPrint.frx":0000
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
      Height          =   615
      Left            =   2580
      TabIndex        =   17
      Tag             =   "L12"
      Top             =   4200
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1085
      BTYPE           =   14
      TX              =   "&Print"
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
      MICON           =   "frmPrint.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdPrintSetup 
      Height          =   645
      Left            =   240
      TabIndex        =   16
      Tag             =   "L11"
      Top             =   4170
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1138
      BTYPE           =   14
      TX              =   "&Cµi Æ∆t m∏y in"
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
      MICON           =   "frmPrint.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComctlLib.ImageList PaperImageList 
      Left            =   2520
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   26
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":0054
            Key             =   "pPortrait"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":0AA6
            Key             =   "pLandscape"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":13F8
            Key             =   "CollateF"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":6C4C
            Key             =   "CollateT"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraOrientation 
      Caption         =   "Chi“u gi y"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   3480
      TabIndex        =   14
      Tag             =   "L9"
      Top             =   2520
      Width           =   2685
      Begin VB.CheckBox chkCollate 
         Caption         =   "X’p trang"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   150
         TabIndex        =   15
         Tag             =   "L10"
         Top             =   225
         Width           =   1515
      End
      Begin VB.Image imgCollate 
         Height          =   960
         Left            =   525
         Picture         =   "frmPrint.frx":C4A0
         Top             =   525
         Width           =   1755
      End
   End
   Begin VB.ComboBox cboPrinter 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   4965
   End
   Begin VB.Frame fraCopies 
      Caption         =   "SË l≠Óng b∂n in"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   120
      TabIndex        =   10
      Tag             =   "L7"
      Top             =   2520
      Width           =   2415
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   375
         Left            =   1875
         TabIndex        =   12
         Top             =   675
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCopies 
         Height          =   375
         Left            =   195
         TabIndex        =   13
         Top             =   675
         Width           =   1575
      End
      Begin VB.Label lblCopies 
         Caption         =   "SË b∂n in:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   11
         Tag             =   "L8"
         Top             =   435
         Width           =   1455
      End
   End
   Begin VB.Frame fraRange 
      Caption         =   "Ch‰n in"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Tag             =   "L2"
      Top             =   1320
      Width           =   6045
      Begin VB.TextBox txtToPage 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3600
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtFromPage 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2010
         TabIndex        =   7
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton optFromPage 
         Caption         =   "Tı t&rang:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   570
         TabIndex        =   6
         Tag             =   "L5"
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optSelecPage 
         Caption         =   "Trang hi÷&n thÍi"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   5
         Tag             =   "L4"
         Top             =   240
         Width           =   1815
      End
      Begin VB.OptionButton optAll 
         Caption         =   "&T t c∂"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   570
         TabIndex        =   4
         Tag             =   "L3"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblToPage 
         Caption         =   "ß’n trang:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   8
         Tag             =   "L6"
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Label lblStatus 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   4980
   End
   Begin VB.Label lblPrinter 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4965
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'Dim Printer_Name As String
'Dim Copies As Integer
Dim StartPageNumber As Integer
Dim StopPageNumber As Integer
Dim CurrentPage As Integer
'Dim minPage As Long
Dim MaxPage As Long
Dim Document As CRAXDDRT.Report
Dim Collate As Boolean
Dim i As Integer
Private Const PAPER_PORTRAIT As Integer = 1
Private Const PAPER_LANDSCAPE As Integer = 2

Private Sub cboPrinter_Change()
    Dim pHandle As Long
    Screen.MousePointer = vbHourglass
    Set Printer = Printers(cboPrinter.ListIndex)
    lblPrinter.Caption = "M∏y in: " & Printer.DeviceName
    DoEvents
    On Error Resume Next
    pHandle = Printer.hdc
    If Err Then
        lblStatus = "Trπng th∏i m∏y in: Tæt"
        cmdOk.Enabled = False
    Else
        lblStatus = "Trπng th∏i m∏y in: MÎ"
        cmdOk.Enabled = True
    End If
    Screen.MousePointer = vbNormal
'    cboQuanlity.ListIndex = Printer.PrintQuality
End Sub

Private Sub cboPrinter_Click()
    Call cboPrinter_Change
End Sub

Private Sub chkCollate_Click()
    If chkCollate.Value Then
        Collate = True
        imgCollate.Picture = PaperImageList.ListImages("CollateT").Picture
    Else
        Collate = False
        imgCollate.Picture = PaperImageList.ListImages("CollateF").Picture
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim Orent As Integer
    Dim tmpPage As Integer
    If optAll.Value Then
        StartPageNumber = 1
        StopPageNumber = MaxPage
    ElseIf optSelecPage.Value Then
        StartPageNumber = CurrentPage
        StopPageNumber = CurrentPage
    ElseIf optFromPage.Value Then
        If CInt(txtToPage.Text) < CInt(txtFromPage.Text) Then
            tmpPage = CInt(txtToPage.Text)
            txtToPage.Text = txtFromPage.Text
            txtFromPage.Text = tmpPage
        End If
        StartPageNumber = CInt(txtFromPage.Text)
        StopPageNumber = CInt(txtToPage.Text)
    End If
'    Document.SelectPrinter GetSettingStr("Receip", "Receipt_DeviceName", True, myIniFile), GetSettingStr("Report", "Report_DeviceName", True, myIniFile), Printer.Port
'    Document.PrintOut False, CInt(txtCopies.Text), True, StartPageNumber, StopPageNumber
    If Printer.DeviceName <> Document.PrinterName Then
        Orent = Document.PaperOrientation
        Document.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
        Document.PaperSize = crPaperA4
        Document.PaperOrientation = Orent
    End If
    Document.PrintOut False, CInt(txtCopies.Text), True, StartPageNumber, StopPageNumber
    
    Unload Me
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    Dim ctrl As Control
    Dim DescArr() As String
    If cmdPrintSetup.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
        DescArr = LoadLanguage(LngFile, "#03:024:")
        Me.Caption = DescArr(1)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
        Next ctrl
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Form_Load"

End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Dim prn As Printer
    'Temporary define
    cboPrinter.Clear
    For Each prn In Printers
        cboPrinter.AddItem prn.DeviceName
    Next
    For i = 0 To cboPrinter.ListCount - 1
        If cboPrinter.List(i) = Printer.DeviceName Then
            cboPrinter.ListIndex = CDbl(GetSettingStr("Report", "Report", True, myIniFile))
            Exit For
        End If
    Next
    lblPrinter.Caption = "M∏y in: " & Printer.DeviceName
    optAll.Value = True
    txtCopies.Text = 1
    chkCollate.Value = 1
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " Form_Load"
    'Call cmdOK_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Document = Nothing
End Sub

Private Sub optFromPage_Click()
    If optFromPage.Value Then
        txtFromPage.SetFocus
        If Trim(txtFromPage.Text) = "" Then
            txtFromPage.Text = "1"
            txtToPage.Text = MaxPage
        End If
        txtFromPage.SelStart = 0
        txtFromPage.SelLength = 9999
    End If
End Sub

Private Sub txtCopies_Change()
    If CInt(txtCopies.Text) <= 0 Then
        txtCopies.Text = 1
    ElseIf CInt(txtCopies.Text) > 256 Then
        txtCopies.Text = 256
    End If
End Sub

Private Sub txtCopies_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        Call UpDown1_DownClick
    ElseIf KeyCode = vbKeyUp Then
        Call UpDown1_UpClick
    End If
End Sub

Private Sub txtCopies_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc("0") And KeyAscii > Asc("9") Then KeyAscii = 0
End Sub

Private Sub txtFromPage_Change()
    If txtFromPage.Text <> "" Then
        If Not optFromPage.Value Then optFromPage.Value = True
    End If
    If CInt(txtFromPage.Text) > MaxPage Then
        txtFromPage.Text = MaxPage
    ElseIf CInt(txtFromPage.Text) = 0 Then
        txtFromPage.Text = 1
    End If
End Sub

Private Sub txtFromPage_KeyPress(KeyAscii As Integer)
    If KeyAscii < &H30 Or KeyAscii > &H39 Then
        If KeyAscii <> &H8 Then KeyAscii = 0
    End If
End Sub

Private Sub txtToPage_Change()
    If Not optFromPage.Value Then optFromPage.Value = True
    If CInt(txtToPage.Text) > MaxPage Then
        txtToPage.Text = MaxPage
    ElseIf CInt(txtToPage.Text) = 0 Then
        txtToPage.Text = 1
    End If
End Sub

Private Sub txtToPage_GotFocus()
    txtToPage.SelStart = 0
    txtToPage.SelLength = 9999
End Sub

Private Sub txtToPage_KeyPress(KeyAscii As Integer)
    If KeyAscii < &H30 Or KeyAscii > &H39 Then
        If KeyAscii <> &H8 Then KeyAscii = 0
    End If
End Sub

Private Sub UpDown1_DownClick()
    txtCopies.Text = CInt(txtCopies.Text) - 1
End Sub

Private Sub UpDown1_UpClick()
    txtCopies.Text = CInt(txtCopies.Text) + 1
End Sub

Public Property Let MaxPageNumber(ByVal vMaxPage As Integer)
    MaxPage = vMaxPage
End Property

Public Property Let CurrentPageNumber(ByVal vCurrentPage As Integer)
    CurrentPage = vCurrentPage
End Property

Public Property Let DocumentPrint(ByVal vDocument As CRAXDDRT.Report)
    Set Document = vDocument
End Property
