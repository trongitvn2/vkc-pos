VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmPrintDefault 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Default Printer"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
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
   Icon            =   "frmPrintDefault.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdHelp 
      Height          =   795
      Left            =   1980
      TabIndex        =   9
      Tag             =   "L6"
      Top             =   7470
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1402
      BTYPE           =   6
      TX              =   "&Trî gióp"
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
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPrintDefault.frx":000C
      PICN            =   "frmPrintDefault.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdSave 
      Cancel          =   -1  'True
      Height          =   765
      Left            =   150
      TabIndex        =   8
      Tag             =   "L5"
      Top             =   7470
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1349
      BTYPE           =   6
      TX              =   "&L­u"
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
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPrintDefault.frx":0662
      PICN            =   "frmPrintDefault.frx":067E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Frame fraPrintertype 
      Caption         =   "Printer type"
      Height          =   1785
      Left            =   3780
      TabIndex        =   3
      Top             =   6480
      Width           =   6405
      Begin VB.OptionButton Option3 
         Caption         =   "Full Size Printer"
         Height          =   375
         Left            =   180
         TabIndex        =   6
         Top             =   1140
         Width           =   6135
      End
      Begin VB.OptionButton Option2 
         Caption         =   "A4 Page Size of Printer"
         Height          =   435
         Left            =   180
         TabIndex        =   5
         Top             =   660
         Width           =   6135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Thermal Receipt Printer"
         Height          =   435
         Left            =   180
         TabIndex        =   4
         Top             =   270
         Value           =   -1  'True
         Width           =   6165
      End
   End
   Begin VB.Frame fraPrint 
      Caption         =   "Select local Window Printer"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   3780
      TabIndex        =   1
      Tag             =   "L8"
      Top             =   90
      Width           =   6405
      Begin MSForms.ListBox lstPrinter 
         Height          =   5835
         Left            =   30
         TabIndex        =   2
         Top             =   420
         Width           =   6285
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "11086;10292"
         MatchEntry      =   1
         FontName        =   ".VnArial"
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame fraPrintType 
      Caption         =   "Select Friendly Printer"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   60
      TabIndex        =   0
      Tag             =   "L7"
      Top             =   90
      Width           =   3705
      Begin VB.ListBox lstReport 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5520
         Left            =   30
         TabIndex        =   7
         Top             =   420
         Width           =   3615
      End
   End
   Begin prjTouchScreen.MyButton cmdDelete 
      Height          =   795
      Left            =   1980
      TabIndex        =   10
      Top             =   6600
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   1402
      BTYPE           =   6
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
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPrintDefault.frx":0BC2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdAddnew 
      Height          =   765
      Left            =   150
      TabIndex        =   11
      Top             =   6630
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1349
      BTYPE           =   6
      TX              =   "&Thªm míi"
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
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmPrintDefault.frx":0BDE
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
Attribute VB_Name = "frmPrintDefault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arrdesc() As String
Dim settingReport As String
Dim settingReceipt As String
Dim settingBarcode As String
Dim Arr(3) As String
Dim rsFriendPrint As New ADODB.Recordset
Dim rsMapingPrint As New ADODB.Recordset
Dim rsVirtual As New ADODB.Recordset
Dim FlagChange As Boolean

Private Sub cmdAddNew_Click()
    With frmKeyboard
        .FormCallkeyboard = "AddPrint"
        .txtInput.PasswordChar = ""
        .Show vbModal
    End With
End Sub

Private Sub cmddelete_Click()
On Error GoTo Handle
If MsgBox("B¹n cã muèn xãa m¸y in nµy kh«ng?", vbYesNo) = vbYes Then
    With rsFriendPrint
        .Find "PrtID='" & Format(lstReport.ListCount - 3, "00") & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Delete adAffectCurrent
        End If
    End With
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdDelete_Click"
End Sub

Private Sub cmdSave_Click()
 Dim str As String
 Dim i As Integer
 Dim rsSystemPrint As New ADODB.Recordset
    If Option1.Value = True Then
        str = "Receipt"
    ElseIf Option2.Value = True Then
        str = "A4"
    Else
        str = "FullSize"
    End If
    SaveSettingStr "LoaiReceipt", "ReceiptType", str, myIniFile
    If FlagChange = True Then
    If cnData.State = adStateClosed Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
        Set rsSystemPrint = OpenCriticalTable("Select * from SystemFlag where SF='02'", cnData)
        cnData.Execute "Delete  from Printer_Mapping"
        If rsVirtual.State <> 0 Then rsVirtual.MoveFirst
        With rsMapingPrint
            Do While Not rsVirtual.EOF
'            .Find "PrinterName='" & rsVirtual.Fields("PrintID") & "'", , adSearchForward, adBookmarkFirst
'                If .EOF Then
                    .addNew
                    .Fields("Station_ID") = Sec_ID
                    .Fields("Station_ID") = Store_ID
                    .Fields("PrinterName") = rsVirtual.Fields("PrintID")
                    .Fields("Details") = rsVirtual.Fields("Details")
                    .Fields("PrtIndex") = rsVirtual.Fields("PrtIndex")
                    '.Fields("Disable") = rsVirtual.Fields("Disable")
                    .Update
'                End If
            rsVirtual.MoveNext
            Loop
        End With
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error GoTo Handle
Dim ctrl As Control
 If cmdAddNew.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
 Me.Caption = Arrdesc(1)
 For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Arrdesc(Mid(ctrl.Tag, 2))
 Next ctrl
' If cnData.State = 0 Then
'    Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
' End If
 If cnData.State <> 0 Then
    Set rsFriendPrint = OpenCriticalTable("select * from Friendly_Printers order by PrtID", cnData)
 End If
 With rsVirtual
    If .State = 0 Then
        .Fields.Append "Station_ID", adVarWChar, 4
        .Fields.Append "PrintID", adVarWChar, 2
        .Fields.Append "Details", adVarWChar, 255
        .Fields.Append "Disable", adBoolean
        .Fields.Append "PrtIndex", adDouble
        .Open
    End If
        Do While Not rsFriendPrint.EOF
            .Find "PrintID='" & rsFriendPrint.Fields("PrtID") & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("Station_ID") = Store_ID
                .Fields("PrintID") = rsFriendPrint.Fields("PrtID")
                .Update
            End If
        rsFriendPrint.MoveNext
        Loop
    
 End With
 Call LoadList

Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo Handle

 Arrdesc = LoadLanguage(LngFile, "#01:006:")
 Arr(0) = GetSettingStr("Report", "Report", True, myIniFile)
 Arr(1) = GetSettingStr("Receipt", "Receipt", True, myIniFile)
 Arr(2) = GetSettingStr("Lable", "Lable", True, myIniFile)
 Dim ctrl As Control
 Option1.Value = True
 Me.Caption = Arrdesc(1)
 For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Arrdesc(Mid(ctrl.Tag, 2))
 Next ctrl
' If cnData.State = 0 Then
'    Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
' End If
 If cnData.State <> 0 Then
    Set rsFriendPrint = OpenCriticalTable("select * from Friendly_Printers order by PrtID", cnData)
    Set rsMapingPrint = OpenCriticalTable("select * from Printer_Mapping order by PrinterName", cnData)
 End If
 Call LoadPrinter
 Call LoadList
 Option1.Value = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "Form_Load"
End Sub

Public Sub LoadPrinter()
On Error GoTo Handle
Dim prt As Printer
    With lstPrinter
        .Clear
        For Each prt In Printers
            .AddItem prt.DeviceName
        Next
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "LoadPrinter"
End Sub
Public Sub LoadList()
On Error GoTo Handle
Dim i As Integer
    With lstReport
        .Clear
        For i = 2 To 4 Step 1
            .AddItem Arrdesc(i)
            .ItemData(lstReport.NewIndex) = 1000
        Next
    End With
    If rsFriendPrint.State <> 0 Then rsFriendPrint.MoveFirst
    With rsFriendPrint
            Do While Not .EOF
                lstReport.AddItem .Fields("PrinterName")
                lstReport.ItemData(lstReport.NewIndex) = 1000
            .MoveNext
            Loop
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "LoadList"
End Sub

Private Sub lstPrinter_Click()
    On Error GoTo Handle
    Select Case lstReport.ListIndex
        Case 0: SaveSettingStr "Receipt", "Receipt", lstPrinter.ListIndex, myIniFile
                SaveSettingStr "Receipt", "Receipt_DeviceName", lstPrinter.Text, myIniFile
        Case 1: SaveSettingStr "Report", "Report", lstPrinter.ListIndex, myIniFile
                SaveSettingStr "Report", "Report_DeviceName", lstPrinter.Text, myIniFile
        Case 2: SaveSettingStr "Lable", "Lable", lstPrinter.ListIndex, myIniFile
                SaveSettingStr "Lable", "Lable_DeviceName", lstPrinter.Text, myIniFile
        Case Else
            FlagChange = True
            With rsVirtual
                .Find "PrintID='" & Format(lstReport.ListIndex - 2, "00") & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    .Fields("Details") = lstPrinter.Text
                    .Fields("PrtIndex") = lstPrinter.ListIndex
                    '.Fields("Disabled") = 0
                    .Update
                End If
            End With
        
    End Select
    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "lstReport_Click"
End Sub

Private Sub lstReport_Click()
On Error GoTo Handle
    Select Case lstReport.ListIndex
        Case 0: If IsNumeric(Arr(0)) Then lstPrinter.Selected(Arr(0)) = True
        Case 1: If IsNumeric(Arr(1)) Then lstPrinter.Selected(Arr(1)) = True
        Case 2: If IsNumeric(Arr(2)) Then lstPrinter.Selected(Arr(2)) = True
        Case Else
            With rsMapingPrint
                .Find "PrinterName='" & Format(lstReport.ListIndex - 2, "00") & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    lstPrinter.Selected(.Fields("PrtIndex")) = True
                End If
            End With
    End Select
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "lstReport_Click"
End Sub
