VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmemployee 
   Caption         =   "Nh©n viªn"
   ClientHeight    =   10995
   ClientLeft      =   60
   ClientTop       =   240
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
   ScaleHeight     =   10995
   ScaleWidth      =   15240
   StartUpPosition =   2  'CenterScreen
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
      Height          =   1365
      Left            =   8400
      TabIndex        =   23
      Top             =   9650
      Width           =   6660
      Begin prjTouchScreen.MyButton cmdThem 
         Height          =   945
         Left            =   60
         TabIndex        =   24
         Tag             =   "16"
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1667
         BTYPE           =   14
         TX              =   "Thªm"
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
         MICON           =   "frmemployee.frx":0000
         PICN            =   "frmemployee.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdCapnhat 
         Height          =   945
         Left            =   1350
         TabIndex        =   25
         Tag             =   "L17"
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1667
         BTYPE           =   14
         TX              =   "&CËp nhËt"
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
         MICON           =   "frmemployee.frx":046E
         PICN            =   "frmemployee.frx":048A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdXoa 
         Height          =   945
         Left            =   2655
         TabIndex        =   26
         Tag             =   "L19"
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1667
         BTYPE           =   14
         TX              =   "&Xãa"
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
         MICON           =   "frmemployee.frx":09CE
         PICN            =   "frmemployee.frx":09EA
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
         Left            =   3960
         TabIndex        =   27
         Tag             =   "L20"
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1667
         BTYPE           =   14
         TX              =   "&Gióp ®ì"
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
         MICON           =   "frmemployee.frx":1024
         PICN            =   "frmemployee.frx":1040
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
         Left            =   5280
         TabIndex        =   28
         Tag             =   "L21"
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1667
         BTYPE           =   14
         TX              =   "&§ãng"
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
         MICON           =   "frmemployee.frx":167A
         PICN            =   "frmemployee.frx":1696
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
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   13680
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraSup 
      Caption         =   "Danh môc nh©n viªn"
      Height          =   11025
      Left            =   0
      TabIndex        =   11
      Tag             =   "L1"
      Top             =   0
      Width           =   8265
      Begin MSFlexGridLib.MSFlexGrid flgEmployee 
         Height          =   10665
         Left            =   60
         TabIndex        =   12
         Top             =   300
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   18812
         _Version        =   393216
         BackColor       =   16777215
         BackColorFixed  =   16777215
         BackColorBkg    =   16777215
         GridColorFixed  =   12632256
         GridLinesFixed  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   8430
      ScaleHeight     =   945
      ScaleWidth      =   6465
      TabIndex        =   0
      Top             =   120
      Width           =   6525
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Employee_ID"
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   90
         TabIndex        =   2
         Top             =   45
         Width           =   6105
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Employee_Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   90
         TabIndex        =   1
         Top             =   480
         Width           =   6135
      End
   End
   Begin TabDlg.SSTab TabSupplier 
      Height          =   8025
      Left            =   8370
      TabIndex        =   3
      Top             =   1590
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   14155
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   882
      TabCaption(0)   =   "Th«ng tin nh©n viªn"
      TabPicture(0)   =   "frmemployee.frx":7930
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraIn"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraIn 
         Height          =   7335
         Left            =   90
         TabIndex        =   4
         Top             =   570
         Width           =   6525
         Begin VB.Frame frmImage 
            Caption         =   "H×nh 3x4"
            Height          =   2295
            Left            =   4560
            TabIndex        =   35
            Tag             =   "L15"
            Top             =   4920
            Width           =   1815
            Begin VB.Image Employee_Pict 
               Height          =   1935
               Left            =   120
               Stretch         =   -1  'True
               Top             =   240
               Width           =   1575
            End
         End
         Begin VB.TextBox txtMail 
            Height          =   555
            Left            =   510
            TabIndex        =   33
            Top             =   6120
            Width           =   4035
         End
         Begin VB.TextBox txtPhone 
            Height          =   495
            Left            =   480
            TabIndex        =   29
            Top             =   5160
            Width           =   3045
         End
         Begin VB.ComboBox cboWork_Shift 
            Height          =   345
            Left            =   2760
            TabIndex        =   22
            Text            =   "Ca"
            Top             =   3480
            Width           =   1695
         End
         Begin VB.CheckBox chkActive 
            Caption         =   "Active"
            Height          =   495
            Left            =   4680
            TabIndex        =   18
            Tag             =   "L10"
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtDept_Name 
            Height          =   495
            Left            =   2430
            TabIndex        =   17
            Top             =   2520
            Width           =   3975
         End
         Begin VB.ComboBox cboDept 
            Height          =   345
            Left            =   480
            TabIndex        =   16
            Text            =   "Phßng ban"
            Top             =   2640
            Width           =   1935
         End
         Begin VB.TextBox txtEmployee_Swip 
            Height          =   525
            Left            =   3270
            TabIndex        =   13
            Top             =   840
            Width           =   3135
         End
         Begin VB.TextBox txtEmployee_Name 
            Height          =   495
            Left            =   510
            TabIndex        =   7
            Top             =   1650
            Width           =   5895
         End
         Begin VB.TextBox txtEmployee_ID 
            Height          =   525
            Left            =   510
            TabIndex        =   6
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtAdd1 
            Height          =   495
            Left            =   510
            TabIndex        =   5
            Top             =   4200
            Width           =   5895
         End
         Begin MSComCtl2.DTPicker dtpEmployee_Birthday 
            Height          =   375
            Left            =   480
            TabIndex        =   20
            Top             =   3480
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
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
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   63242241
            UpDown          =   -1  'True
            CurrentDate     =   39448
         End
         Begin VB.Label lblMail 
            Caption         =   "Email:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Tag             =   "L14"
            Top             =   5760
            Width           =   1125
         End
         Begin VB.Label lblPhone 
            Caption         =   "§iÖn tho¹i:"
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Tag             =   "L13"
            Top             =   4800
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "B¶ng c«ng viÖc"
            Height          =   315
            Left            =   3960
            TabIndex        =   31
            Tag             =   "L12"
            Top             =   3120
            Width           =   1575
         End
         Begin MSForms.ComboBox cboJobCode 
            Height          =   375
            Left            =   4560
            TabIndex        =   30
            Top             =   3480
            Width           =   1815
            VariousPropertyBits=   746604571
            DisplayStyle    =   3
            Size            =   "3201;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   ".VnArial"
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            Caption         =   "Ca"
            Height          =   285
            Left            =   2640
            TabIndex        =   21
            Tag             =   "L9"
            Top             =   3120
            Width           =   825
         End
         Begin VB.Label lblFax 
            Caption         =   "Ngµy sinh"
            Height          =   285
            Left            =   120
            TabIndex        =   19
            Tag             =   "L8"
            Top             =   3120
            Width           =   1665
         End
         Begin VB.Label Label3 
            Caption         =   "Phßng ban"
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Tag             =   "L7"
            Top             =   2280
            Width           =   1665
         End
         Begin VB.Label Label2 
            Caption         =   "M· Swip"
            Height          =   285
            Left            =   3120
            TabIndex        =   14
            Tag             =   "L5"
            Top             =   480
            Width           =   2265
         End
         Begin VB.Label lblSupName 
            Caption         =   "Tªn nh©n viªn"
            Height          =   285
            Left            =   120
            TabIndex        =   10
            Tag             =   "L6"
            Top             =   1380
            Width           =   1665
         End
         Begin VB.Label lblSupCode 
            Caption         =   "M· Nh©n viªn"
            Height          =   285
            Left            =   120
            TabIndex        =   9
            Tag             =   "L4"
            Top             =   450
            Width           =   1665
         End
         Begin VB.Label lblAdd 
            Caption         =   "§Þa chØ:"
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Tag             =   "L11"
            Top             =   3960
            Width           =   1665
         End
      End
   End
End
Attribute VB_Name = "frmemployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsEmployee As New ADODB.Recordset
Dim strPath As String
Dim DescArr() As String
Dim wYear As Double
Dim wMonth As Integer
Dim rsCom_Dept As New ADODB.Recordset


Private Sub cboDept_Click()
    Call cboDept_Change
End Sub

Private Sub cmdCapnhat_Click()
    Call UpdateDatabase
    Call LoadControl
    If cmdThem.Enabled = True Then
        cmdThem.SetFocus
    Else
        cmdThem.Enabled = True
        cmdThem.SetFocus
    End If
End Sub

Private Sub cboDept_Change()
With rsCom_Dept
    .Find "Dept_ID='" & Trim(cboDept.Text) & "'", , adSearchForward, adBookmarkFirst
    If Not .EOF Then
        txtDept_Name.Text = .Fields("Dept_Name")
    End If
End With
End Sub

Private Sub cmdClose_Click()
    Set rsEmployee = Nothing
    Unload Me
End Sub

Private Sub cmdThem_Click()
On Error GoTo Handle
    If cmdThem.Caption = Trim(DescArr(16)) Then
        Call UnlockText
        Call DeleteTextbox
    ElseIf cmdThem.Caption = DescArr(18) Then
        Call UnlockText
    End If
Exit Sub
Handle:
MsgBox Err.Number & " " & Err.Description & " " & _
            Me.name & " " & "cmdThem _click"

End Sub
Private Sub DeleteTextbox()
    On Error GoTo Handle
        cmdThem.Caption = DescArr(18)
        Call Add_Dept_t_Employ
        Call Add_Work_Shift
        Call Add_JobCode
        txtEmployee_ID.Text = ""
        txtEmployee_Name.Text = ""
        txtEmployee_Swip.Text = ""
        dtpEmployee_Birthday.Value = Date
        txtAdd1.Text = ""
        txtPhone.Text = ""
        txtMail.Text = ""
        txtEmployee_ID.SetFocus
        
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "DeleteTextbox"
End Sub
Private Sub UpdateDatabase()
    On Error GoTo Handle
        With rsEmployee
            .Find "Cashier_ID='" & txtEmployee_ID.Text & _
                "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
            End If
                .Fields("Cashier_ID") = txtEmployee_ID.Text
                .Fields("JobCode_ID") = Format(cboJobCode.ListIndex + 1, "00")
                .Fields("Dept_ID") = Trim(cboDept.Text)
                .Fields("Swipe_ID") = txtEmployee_Swip.Text
                .Fields("EmpName") = txtEmployee_Name.Text
                .Fields("Address") = txtAdd1.Text
                .Fields("Phone") = txtPhone.Text
                .Fields("EMail") = txtMail.Text
                .Fields("Birthday") = dtpEmployee_Birthday.Value
                .Fields("Picture") = ""
                .Fields("Shift") = Format(cboWork_Shift.ListIndex + 1, "00")
                If chkActive.Value = 1 Then .Fields("Disabled") = True
                .Update
                .Requery
'            Else
'                MsgBox DescArr(8), vbOKOnly
'                Call DeleteTextbox
'            End If
        End With
        Call setflgEmployee
        cmdThem.Caption = Trim(DescArr(16))
        
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "UpdateDatabase"
End Sub

Private Sub cmdXoa_Click()
    On Error GoTo Handle
    Dim ans As Integer
    ans = MsgBox("B¹n cã ch¾c ch¨n muèn xãa danh môc nµy kh«ng?", vbYesNo)
    If ans = vbYes Then
        With rsEmployee
            .Find "Cashier_ID='" & flgEmployee.TextMatrix(flgEmployee.Row, 0) & _
                    "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Or .BOF Then
                .Delete adAffectCurrent
                .MoveNext
                .Requery
            End If
            Call Form_Load
        End With
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "cmdXoa_Click"
End Sub

Private Sub Employee_Pict_Click()
Dim fso As New FileSystemObject
Dim P As String
    With comdlg
         .FileName = ""
        .Filter = "Image(*.jpg)|*.bmp|*.jpg"
        .DefaultExt = "*.jpg"
        .InitDir = App.Path
        .ShowOpen
        If .FileName <> "" Then
            Employee_Pict.Picture = LoadPicture(.FileName)
            P = .FileName
        End If
        If rsEmployee.State = 1 And rsEmployee.RecordCount > 0 Then
            rsEmployee.MoveFirst
        Else
            Exit Sub
        End If
        With rsEmployee
            .Find "Cashier_ID='" & flgEmployee.TextMatrix(flgEmployee.Row, 0) & _
                        "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Or .BOF Then
                .Fields("Picture") = P
                .Update
            End If
            
        End With
    End With
End Sub

Private Sub flgEmployee_EnterCell()
    On Error GoTo Handle
    With rsEmployee
        .Find "Cashier_ID='" & flgEmployee.TextMatrix(flgEmployee.Row, 0) & _
                "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtEmployee_ID.Text = .Fields("Cashier_ID")
            cboJobCode.ListIndex = CDbl(.Fields("JobCode_ID")) - 1
            cboDept.Text = .Fields("Dept_ID")
            txtEmployee_Swip.Text = .Fields("Swipe_ID")
            txtEmployee_Name.Text = .Fields("EmpName")
            txtAdd1.Text = .Fields("Address")
            txtPhone.Text = .Fields("Phone")
            txtMail.Text = .Fields("EMail")
            dtpEmployee_Birthday.Value = .Fields("Birthday")
            Employee_Pict.Picture = LoadPicture(.Fields("Picture"))
            cboWork_Shift.ListIndex = CDbl(.Fields("Shift")) - 1
            If !Disabled = True Then
                chkActive.Value = 1
            Else
                chkActive.Value = 0
            End If
            lblNo.Caption = !cashier_ID
            lblName.Caption = !EmpName
        End If
    End With
    Exit Sub
    
Handle:
    MsgBox Err.Number & " " & Err.Description & vbCrLf _
    & Me.name & " flgEmployee_EnterCell"
End Sub

Private Sub Form_Activate()
    Dim ctrl As Control
    If cmdThem.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    DescArr = LoadLanguage(LngFile, "#03:005:")
    Me.Caption = DescArr(1)
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    Dim strPath As String
    Dim str As String
    DescArr = LoadLanguage(LngFile, "#03:005:")
    str = "Select * from Employee"
    Set rsEmployee = OpenCriticalTable(str, cnData)
    Set rsCom_Dept = Open_Table(cnData, "Company_Dept")
    
    Call Add_Dept_t_Employ
    Call Add_Work_Shift
    Call Add_JobCode
    Call setflgEmployee
    Call LockText
    Exit Sub
Handle:
    MsgBox Err.Number & " " & _
    Err.Description & Me.name & "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Handle
    Set rsEmployee = Nothing
Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & "" & _
                        Me.name & " " & "form_Unload"
End Sub
Private Sub setflgEmployee()
On Error GoTo errHdl
    Dim intCount    As Integer
    With flgEmployee
    .Cols = 10
    .Rows = 2
        .Font = ".vnArial"
        .ColWidth(0) = 1500
        .ColWidth(1) = 1800
        .ColWidth(2) = 3000
        .ColWidth(3) = 1300
        .ColWidth(4) = 1500
        .ColWidth(5) = 2000
        .ColWidth(6) = 2500
        .ColWidth(7) = 2000
        .ColWidth(8) = 1500
        .ColWidth(9) = 2500

        .TextMatrix(0, 0) = DescArr(4)
        .TextMatrix(0, 1) = DescArr(5)
        .TextMatrix(0, 2) = DescArr(6)
        .TextMatrix(0, 3) = DescArr(7)
        .TextMatrix(0, 4) = DescArr(8)
        .TextMatrix(0, 5) = DescArr(9)
        .TextMatrix(0, 6) = DescArr(11)
        .TextMatrix(0, 7) = DescArr(12)
        .TextMatrix(0, 8) = DescArr(13)
        .TextMatrix(0, 9) = DescArr(14)
        
    End With
    
    If rsEmployee Is Nothing Then Exit Sub
    If rsEmployee.State = 0 Then Exit Sub
    
    If rsEmployee.EOF And rsEmployee.BOF Then
        With flgEmployee
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
   flgEmployee.Rows = rsEmployee.RecordCount + 1
    intCount = 0
    Do While Not rsEmployee.EOF
        intCount = intCount + 1
        flgEmployee.TextMatrix(intCount, 0) = rsEmployee!cashier_ID
        flgEmployee.TextMatrix(intCount, 1) = rsEmployee!Swipe_ID
        flgEmployee.TextMatrix(intCount, 2) = rsEmployee!EmpName
        flgEmployee.TextMatrix(intCount, 3) = rsEmployee!Dept_ID
        flgEmployee.TextMatrix(intCount, 4) = rsEmployee!Birthday
        flgEmployee.TextMatrix(intCount, 5) = rsEmployee!Dept_ID
        flgEmployee.TextMatrix(intCount, 6) = rsEmployee!Address
        flgEmployee.TextMatrix(intCount, 7) = rsEmployee!JobCode_ID
        flgEmployee.TextMatrix(intCount, 8) = rsEmployee!Phone
        flgEmployee.TextMatrix(intCount, 9) = rsEmployee!Email

        rsEmployee.MoveNext
    Loop
'    SetColorFlexGrid flgEmployee, 1, 1, flgEmployee.Cols

    Call flgEmployee_EnterCell
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - setflgEmployee "
End Sub
Private Sub LoadControl()
    On Error GoTo Handle
    
    With rsEmployee
        .Find "Cashier_ID='" & !cashier_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtEmployee_ID.Text = .Fields("Cashier_ID")
            cboJobCode.ListIndex = CDbl("0" & .Fields("JobCode_ID")) - 1
            cboDept.Text = .Fields("Dept_ID")
            txtEmployee_Swip.Text = .Fields("Swipe_ID")
            txtEmployee_Name.Text = .Fields("EmpName")
            txtAdd1.Text = .Fields("Address")
            txtPhone.Text = .Fields("Phone")
            txtMail.Text = .Fields("EMail")
            dtpEmployee_Birthday.Value = .Fields("Birthday")
            Employee_Pict.Picture = LoadPicture(.Fields("Picture"))
            cboWork_Shift.ListIndex = CDbl("0" & .Fields("Shift")) - 1
            If !Disabled = True Then
                chkActive.Value = 1
            Else
                chkActive.Value = 0
            End If
            lblNo.Caption = !cashier_ID
            lblName.Caption = !EmpName
            .Requery
        End If
    End With
    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " LoadControl"
End Sub

Public Sub UnlockText()
    On Error GoTo Handle
         txtEmployee_ID.Locked = False
        cboJobCode.Enabled = True
        cboDept.Enabled = True
        txtEmployee_Swip.Locked = False
        txtEmployee_Name.Locked = False
        txtAdd1.Locked = False
        txtPhone.Locked = False
        txtMail.Locked = False
        dtpEmployee_Birthday.Enabled = True
        cboWork_Shift.Enabled = True
        txtEmployee_ID.SetFocus
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub
Public Sub LockText()
    On Error GoTo Handle
        txtEmployee_ID.Locked = True
        cboJobCode.Enabled = True
        cboDept.Enabled = True
        txtEmployee_Swip.Locked = True
        txtEmployee_Name.Locked = True
        txtAdd1.Locked = True
        txtPhone.Locked = True
        txtMail.Locked = True
        dtpEmployee_Birthday.Enabled = True
        cboWork_Shift.Enabled = True
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub

Private Sub txtEmployee_Name_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtEmployee_Name.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtEmployee_ID_DblClick "

End Sub

Private Sub txtEmployee_Name_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        cboDept.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtEmployee_Name_KeyPress"

End Sub

Private Sub txtEmployee_ID_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtEmployee_ID.Text = .Let_Text_Input
       End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtEmployee_ID_DblClick "

End Sub

Private Sub txtEmployee_ID_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        txtEmployee_Name.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtEmployee_ID_KeyPress"
End Sub

Public Sub Add_Dept_t_Employ()
On Error GoTo Handle
    With cboDept
        .Clear
        If rsCom_Dept.State = 1 And rsCom_Dept.RecordCount > 0 Then
            rsCom_Dept.MoveFirst
        Else
            Exit Sub
        End If
        Do While Not rsCom_Dept.EOF
            .AddItem rsCom_Dept.Fields("Dept_ID")
        rsCom_Dept.MoveNext
        Loop
    End With
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name
End Sub

Public Sub Add_Work_Shift()
On Error GoTo Handle
Dim rsWork_Shift As New ADODB.Recordset
Set rsWork_Shift = Open_Table(cnData, "Work_Shift")
If rsWork_Shift.State = 1 And rsWork_Shift.RecordCount > 0 Then
    rsWork_Shift.MoveFirst
Else
    Exit Sub
End If
    With cboWork_Shift
        .Clear
        Do While Not rsWork_Shift.EOF
            .AddItem rsWork_Shift.Fields("Shift_Name")
        rsWork_Shift.MoveNext
        Loop
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Add_Work_Shift"
End Sub
Public Sub Add_JobCode()
On Error GoTo Handle
Dim rsJobcode As New ADODB.Recordset
Set rsJobcode = Open_Table(cnData, "JobCode")
If rsJobcode.State = 1 And rsJobcode.RecordCount > 0 Then
    rsJobcode.MoveFirst
Else
    Exit Sub
End If
    With cboJobCode
        .Clear
        Do While Not rsJobcode.EOF
            .AddItem rsJobcode.Fields("JobCode_Name")
        rsJobcode.MoveNext
        Loop
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Add_Work_Shift"
End Sub

