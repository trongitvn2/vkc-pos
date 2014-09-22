VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTableSelect 
   Caption         =   "Chän bµn ®Æt"
   ClientHeight    =   10260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   ClipControls    =   0   'False
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
   ScaleHeight     =   10260
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNum 
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   13800
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Height          =   975
      Left            =   13800
      TabIndex        =   6
      Top             =   6000
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Close"
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
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTableSelect.frx":0000
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
      Height          =   975
      Left            =   13800
      TabIndex        =   5
      Top             =   4920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "OK"
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
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTableSelect.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   9225
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13695
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFC0&
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         Height          =   1035
         Index           =   0
         Left            =   1560
         Top             =   960
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblTable 
         Alignment       =   2  'Center
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         Caption         =   "#1"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Index           =   0
         Left            =   1590
         TabIndex        =   2
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
         WordWrap        =   -1  'True
      End
      Begin VB.Shape Shape2 
         Height          =   1785
         Index           =   0
         Left            =   1470
         Top             =   3660
         Visible         =   0   'False
         Width           =   2805
      End
      Begin VB.Label lblSection 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Left            =   11760
         TabIndex        =   1
         Top             =   240
         Width           =   1545
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   1095
      Left            =   0
      TabIndex        =   3
      Top             =   9120
      Width           =   15525
      _ExtentX        =   27384
      _ExtentY        =   1931
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "   "
      TabPicture(0)   =   "frmTableSelect.frx":0038
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdSection(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin prjTouchScreen.MyButton cmdSection 
         Height          =   1005
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   40
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1773
         BTYPE           =   6
         TX              =   "Section"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421631
         BCOLO           =   8421631
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmTableSelect.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
   End
End
Attribute VB_Name = "frmTableSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Drag As Boolean
Dim rsSection As New ADODB.Recordset
Dim CountTable As Integer
Dim CountSection As Integer
Dim rsTable As New ADODB.Recordset
Dim iLoad As Boolean
Dim iLoadSection As Boolean
Dim rsInvoice_On_Holds As New ADODB.Recordset
Dim indexTable As Integer
Dim rsAlign As New ADODB.Recordset
Dim Table_Number As String

Private Sub cmdClose_Click()
    Table_Number = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Table_Number = txtNum.Text
    Unload Me
End Sub

Private Sub cmdSection_Click(Index As Integer)
    On Error GoTo Handle
    Dim ctrl As Control
        Sec_ID = Format(cmdSection(Index).Tag, "00")
        cmdSection(Index).BackColor = vbGreen
        Call LoadTable(CStr(Sec_ID))
        lblSection.Caption = cmdSection(Index).Caption
        iLoad = True
        For Each ctrl In Me
        If ctrl.name = "cmdSection" Then
            ctrl.ForeColor = vbBlue
        End If
    Next ctrl
    cmdSection(Index).ForeColor = vbRed
    Set rsInvoice_On_Holds = OpenCriticalTable("select * from Invoice_OnHold where Section_ID='" & Sec_ID & "'", cnData)
    Exit Sub
    
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdSection_Click "
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    Dim DescArr() As String
    Dim ctrl As Control
    If iLoad = True Then Exit Sub
    iLoad = True
    DescArr = LoadLanguage(LngFile, "#03:014:")
    Me.Caption = DescArr(1)
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    
    Call Load_Section
    If Sec_ID <> "" Then
        Call LoadTable(Sec_ID)
    Else
        Call LoadTable("01")
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   Form_Activate"
End Sub


Private Sub Form_Load()
On Error GoTo Handle
    iLoad = False
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsInvoice_On_Holds = Open_Table(cnData, "Invoice_OnHold")
    
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Handle
'    Set cnData = Nothing
    Set rsSection = Nothing
    Set rsTable = Nothing
    CountTable = 0
    CountSection = 0
    iLoadSection = False

Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   Form_Unload"
End Sub


Private Sub lblTable_Click(Index As Integer)
    On Error GoTo Handle
    Dim i As Integer
        Dim tableCaption  As String
        tableCaption = Left(lblTable(Index).Caption, InStr(Replace(lblTable(Index).Caption, Chr(13) & Chr(13), Chr(13)), Chr(13)))
        'tableCaption = Replace(tableCaption, Chr(13), "")
        txtNum.Text = tableCaption
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " lblTable_Click  "
End Sub

Public Sub Load_Section()
    On Error GoTo Handle
    Dim ctrl As Control
        Dim i, a, b As Integer
        i = 1
        a = 0
        If cnData.State > 0 Then
             Set rsSection = OpenCriticalTable("select * from Table_Diagram_Sections order by Location_ID ASC", cnData)
        Else
            Exit Sub
        End If
        If rsSection.EOF Then Exit Sub
        If iLoadSection = True Then
            For Each ctrl In Me
                If TypeOf ctrl Is MyButton And ctrl.name = "cmdSection" Then
                    a = a + 1
                End If
            Next
            For b = 1 To a - 1
                Unload cmdSection(b)
            Next
            
        End If
            Do While Not rsSection.EOF
                Load cmdSection(i)
                With cmdSection(i)
                    If i = 1 Then
                        .Left = cmdSection(i - 1).Left + 80
                    Else
                        .Left = cmdSection(i - 1).Left + cmdSection(i - 1).Width + 80
                    End If
                    .top = cmdSection(i - 1).top
                    .Visible = True
                    .Caption = rsSection.Fields("Section_ID")
                    .Tag = rsSection.Fields("Location_ID")
                End With
                i = i + 1
            rsSection.MoveNext
            Loop
        CountSection = rsSection.RecordCount
        iLoadSection = True
        a = 0
        
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   Load_Section"
End Sub

Public Sub LoadTable(Section_ID As String)
On Error GoTo Handle
Dim rscolor As New ADODB.Recordset
Dim rsSeatedColor As New ADODB.Recordset
Dim rsInvoice_Total As New ADODB.Recordset
Dim rsVacantColor As New ADODB.Recordset
Dim i, j As Integer
i = 1: j = 1
    Dim str As String
    Dim ctrl As Control
    If CountTable > 0 Then
        For j = 1 To CountTable
            Unload lblTable(j)
            Unload Shape1(j)
        Next
    End If
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    'Lay Bang mau
    Dim TypeColor, SeatedColor, BlankTable As String
    TypeColor = "RESERVED"
    SeatedColor = "SEATED"
    BlankTable = "VACANT"
    Set rscolor = OpenCriticalTable("Select ReserveValue from ColorTablePlan where ReserveType='" & TypeColor & "'", cnData)
    Set rsSeatedColor = OpenCriticalTable("Select ReserveValue from ColorTablePlan where ReserveType='" & SeatedColor & "'", cnData)
    Set rsVacantColor = OpenCriticalTable("Select ReserveValue from ColorTablePlan where ReserveType='" & BlankTable & "'", cnData)

    str = "select * from Table_Diagram where Section_ID='" & Section_ID & "'"
    Set rsTable = OpenCriticalTable(str, cnData)
    CountTable = rsTable.RecordCount
    Dim strTableTotal As String
    Do While Not rsTable.EOF
        Load lblTable(i)
        With lblTable(i)
            .Left = rsTable.Fields("XPOS")
            .top = rsTable.Fields("YPOS")
            .Height = rsTable.Fields("Height")
            .Width = rsTable.Fields("width")
            strTableTotal = "SELECT Invoice_OnHold.Invoice_Number, Invoice_Totals.Store_ID," & _
            "Invoice_OnHold.OnHoldID, Invoice_Totals.Grand_Total, Invoice_Totals.Total_Price, " & _
            "Invoice_Totals.Orig_OnHoldID, Invoice_OnHold.Section_ID FROM Invoice_OnHold" & _
            " INNER JOIN Invoice_Totals ON Invoice_OnHold.Invoice_Number = Invoice_Totals.Invoice_Number " & _
            " where Invoice_OnHold.OnHoldID = '" & rsTable.Fields("Table_number") & Chr(13) & "' and Invoice_OnHold.Section_ID='" & Section_ID & "'"
            Set rsInvoice_Total = OpenCriticalTable(strTableTotal, cnData)
            If rsInvoice_Total.RecordCount > 0 Then
                If CDbl("0" & rsInvoice_Total.Fields("Grand_Total")) > 0 Then
                    .Caption = rsTable.Fields("Table_Number") & Chr(13) & Format(rsInvoice_Total.Fields("Grand_Total"), formatNum)
                    .BackStyle = 1
                    .BackColor = rscolor.Fields("ReserveValue")
                    .FontSize = rsTable.Fields("Cost_Center_Index")
                Else
                    .Caption = rsTable.Fields("Table_Number") & Chr(13) & Format(rsInvoice_Total.Fields("Grand_Total"), formatNum)
                    .BackStyle = 1
                    .BackColor = rsSeatedColor.Fields("ReserveValue")
                    .FontSize = rsTable.Fields("Cost_Center_Index")
                End If
            Else
                .Caption = rsTable.Fields("Table_Number") & Chr(13)
                .FontSize = rsTable.Fields("Cost_Center_Index")
                .BackStyle = 1
                .BackColor = rsVacantColor.Fields("ReserveValue")
            End If

            .Visible = True
        End With
        Load Shape1(i)
        With Shape1(i)
            .Left = lblTable(i).Left - 40
            .top = lblTable(i).top - 45
            .Height = lblTable(i).Height + 100
            .Width = lblTable(i).Width + 100
            .Shape = rsTable.Fields("ShapeType")
            .Visible = True
        End With
    rsTable.MoveNext
    i = i + 1
    Loop

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  LoadTable"
End Sub


Public Property Get Let_Table_Num() As Variant
    Let_Table_Num = Table_Number
End Property

Public Property Get Let_SectionID() As Variant
    Let_SectionID = Sec_ID
End Property
