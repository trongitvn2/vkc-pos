VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGeneralBill 
   Caption         =   "Th«ng tin Bill"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8085
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
   MaxButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog comdlg 
      Left            =   360
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   8281
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Th«ng tin ®Çu bill"
      TabPicture(0)   =   "frmGerneralBill.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Th«ng tin cuèi bill"
      TabPicture(1)   =   "frmGerneralBill.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Th«ng tin ®ång bé d÷ liÖu"
      TabPicture(2)   =   "frmGerneralBill.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   26
         Top             =   480
         Width           =   7575
         Begin VB.TextBox txtAmount 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3840
            TabIndex        =   35
            Top             =   3360
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Frame Frame3 
            Caption         =   "W:186 x H:78 Pixel"
            Height          =   1695
            Left            =   2640
            TabIndex        =   32
            Top             =   120
            Width           =   3495
            Begin VB.Image Image1 
               Height          =   1335
               Left            =   120
               OLEDropMode     =   1  'Manual
               Stretch         =   -1  'True
               ToolTipText     =   "Click here to select Logo"
               Top             =   240
               Width           =   3255
            End
         End
         Begin VB.TextBox txtMailStore 
            Height          =   615
            Left            =   2280
            TabIndex        =   28
            Top             =   2640
            Width           =   5175
         End
         Begin VB.TextBox txtMailHost 
            Height          =   615
            Left            =   2280
            TabIndex        =   27
            Top             =   1920
            Width           =   5175
         End
         Begin MSComCtl2.DTPicker TimeSync 
            Height          =   495
            Left            =   2280
            TabIndex        =   34
            Top             =   3360
            Width           =   1455
            _ExtentX        =   2566
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
            CustomFormat    =   "HH:mm"
            Format          =   64290818
            UpDown          =   -1  'True
            CurrentDate     =   36494
         End
         Begin VB.Label lblTimeSync 
            Alignment       =   1  'Right Justify
            Caption         =   "Thêi gian §ång bé"
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
            Height          =   615
            Left            =   120
            TabIndex        =   33
            Top             =   3360
            Width           =   2055
         End
         Begin VB.Label lblImage 
            Alignment       =   1  'Right Justify
            Caption         =   "Logo h×nh:"
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
            Height          =   495
            Left            =   120
            TabIndex        =   31
            Tag             =   "L14"
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Mail m¸y tr¹m:"
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
            Height          =   615
            Left            =   120
            TabIndex        =   30
            Tag             =   "L16"
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Mail server:"
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
            Height          =   495
            Left            =   120
            TabIndex        =   29
            Tag             =   "L15"
            Top             =   1920
            Width           =   2055
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000A&
         Height          =   3795
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   7515
         Begin VB.TextBox L1 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "VNI-Ariston"
               Size            =   12
               Charset         =   0
               Weight          =   800
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1770
            MaxLength       =   45
            TabIndex        =   17
            Top             =   450
            Width           =   5535
         End
         Begin VB.TextBox L2 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "VNI-Bodon"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   1770
            MaxLength       =   45
            TabIndex        =   16
            Top             =   1080
            Width           =   5535
         End
         Begin VB.TextBox L3 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "VNI-Helve"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1770
            MaxLength       =   45
            TabIndex        =   15
            Top             =   1740
            Width           =   5535
         End
         Begin VB.TextBox L4 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "VNI-Helve"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1770
            MaxLength       =   45
            TabIndex        =   14
            Top             =   2430
            Width           =   5535
         End
         Begin VB.TextBox L5 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "VNI-Helve"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1770
            MaxLength       =   45
            TabIndex        =   13
            Top             =   3090
            Width           =   5535
         End
         Begin VB.Label lblL5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dßng 5:"
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   90
            TabIndex        =   22
            Tag             =   "L19"
            Top             =   3090
            Width           =   1665
         End
         Begin VB.Label lblL4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dßng 4:"
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   90
            TabIndex        =   21
            Tag             =   "L8"
            Top             =   2490
            Width           =   1665
         End
         Begin VB.Label lblL2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dßng 2:"
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   90
            TabIndex        =   20
            Tag             =   "L6"
            Top             =   1050
            Width           =   1665
         End
         Begin VB.Label lblL3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dßng 3:"
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   90
            TabIndex        =   19
            Tag             =   "L7"
            Top             =   1800
            Width           =   1665
         End
         Begin VB.Label lblL1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dßng 1:"
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   90
            TabIndex        =   18
            Tag             =   "L5"
            Top             =   480
            Width           =   1665
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000A&
         Height          =   3795
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   7515
         Begin VB.TextBox L10 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "VNI-Ariston"
               Size            =   12
               Charset         =   0
               Weight          =   800
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1800
            MaxLength       =   45
            TabIndex        =   6
            Top             =   3090
            Width           =   5535
         End
         Begin VB.TextBox L9 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "VNI-Helve"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1770
            MaxLength       =   45
            TabIndex        =   5
            Top             =   2430
            Width           =   5535
         End
         Begin VB.TextBox L8 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "VNI-Helve"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1770
            MaxLength       =   45
            TabIndex        =   4
            Top             =   1740
            Width           =   5535
         End
         Begin VB.TextBox L7 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "VNI-Helve"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   1770
            MaxLength       =   45
            TabIndex        =   3
            Top             =   1080
            Width           =   5535
         End
         Begin VB.TextBox L6 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "VNI-Helve"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1770
            MaxLength       =   45
            TabIndex        =   2
            Top             =   450
            Width           =   5535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dßng 1:"
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   90
            TabIndex        =   11
            Tag             =   "L9"
            Top             =   480
            Width           =   1665
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dßng 3:"
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   90
            TabIndex        =   10
            Tag             =   "L11"
            Top             =   1800
            Width           =   1665
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dßng 2:"
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   90
            TabIndex        =   9
            Tag             =   "L10"
            Top             =   1050
            Width           =   1665
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dßng 4:"
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   90
            TabIndex        =   8
            Tag             =   "L12"
            Top             =   2490
            Width           =   1665
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dßng 5:"
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   90
            TabIndex        =   7
            Tag             =   "L13"
            Top             =   3090
            Width           =   1665
         End
      End
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Cancel          =   -1  'True
      Height          =   855
      Left            =   4080
      TabIndex        =   24
      Tag             =   "L18"
      Top             =   5520
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "§ãng"
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
      BCOL            =   14215660
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmGerneralBill.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdSaveChange 
      Height          =   855
      Left            =   1560
      TabIndex        =   25
      Tag             =   "L17"
      Top             =   5520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "L­u thay ®æi"
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
      BCOL            =   14215660
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmGerneralBill.frx":0070
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "th«ng tin doanh nghiÖp"
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
      Height          =   615
      Left            =   1080
      TabIndex        =   23
      Tag             =   "L1"
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frmGeneralBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSetup As New ADODB.Recordset
Dim str As String
Dim ischange As Boolean

Private Sub cmdClose_Click()
    If ischange = True Then
        If MsgBox("B¹n cã muèn l­u thay ®æi kh«ng?", vbYesNo) = vbYes Then
            cmdSaveChange_Click
        End If
    End If
    Unload Me
End Sub

Private Sub cmdSaveChange_Click()
    On Error GoTo Handle
    'If MsgBox("B¹n cã muèn l­u thay ®æi kh«ng?", vbYesNo) = vbYes Then
        With rsSetup
            .Fields("Company_Info_1") = L1.Text
            .Fields("Company_Info_2") = L2.Text
            .Fields("Company_Info_3") = L3.Text
            .Fields("Company_Info_4") = L4.Text
            .Fields("Company_Info_5") = L5.Text
            .Fields("Invoice_Notes_1") = L6.Text
            .Fields("Invoice_Notes_2") = L7.Text
            .Fields("Invoice_Notes_3") = L8.Text
            .Fields("Invoice_Notes_4") = L9.Text
            .Fields("Invoice_Notes_5") = L10.Text
            .Fields("AmountLimited") = txtAmount.Text
            .Fields("InetEMail") = txtMailHost.Text
            .Fields("Store_Email") = txtMailStore.Text
            .Fields("TimeSync") = Format(TimeSync.Value, "HH:mm:ss")
            .Update
        End With
    'End If
    ischange = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdSaveChange_Click"
End Sub

Private Sub Form_Activate()
On Error GoTo errHdl
    Dim DescArr() As String
    Dim ctrl As Control
    DescArr = LoadLanguage(LngFile, "#03:015:")
'    If cmdSaveChange.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    SSTab1.TabCaption(0) = DescArr(2)
    SSTab1.TabCaption(1) = DescArr(3)
    SSTab1.TabCaption(2) = DescArr(4)
    L2.Font.name = "VNI-Bodon" '"VNI-Thufap2"
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    If UCase(UserID) = "131112" Or UserID = "881507" Then txtAmount.Visible = True
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Activate"

End Sub

Private Sub Form_Load()
On Error GoTo Handle
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\database.mdb", "100881administrator")
'    End If
    Set rsSetup = Open_Table(cnData, "Setup")
    If Not rsSetup.EOF Then
        With rsSetup
            L1.Text = .Fields("Company_Info_1")
            L2.Text = .Fields("Company_Info_2")
            L3.Text = .Fields("Company_Info_3")
            L4.Text = .Fields("Company_Info_4")
            L5.Text = .Fields("Company_Info_5")
            L6.Text = .Fields("Invoice_Notes_1")
            L7.Text = .Fields("Invoice_Notes_2")
            L8.Text = .Fields("Invoice_Notes_3")
            L9.Text = .Fields("Invoice_Notes_4")
            L10.Text = .Fields("Invoice_Notes_5")
            txtAmount.Text = Format(.Fields("AmountLimited"), "#,##0")
            txtMailHost.Text = .Fields("InetEMail")
            txtMailStore.Text = .Fields("Store_Email")
            TimeSync.Value = Format(.Fields("TimeSync"), "HH:mm:ss")
            If Dir(.Fields("Image"), vbDirectory) <> "" Then
                Image1.Picture = LoadPicture(.Fields("Image"))
            End If
        End With
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
ischange = False
CloseRecordset rsSetup
End Sub

Private Sub Image1_Click()
Dim fso As New FileSystemObject
Dim P As String
    With comdlg
         .FileName = ""
         .Filter = "Image(*.jpg)|*.jpg|*.bmp|*.*"
        .DefaultExt = "*.bmp"
        .InitDir = App.Path
        .ShowOpen
        If .FileName <> "" Then
            Image1.Picture = LoadPicture(.FileName)
            P = .FileName
        End If
        With rsSetup
            .Fields("Image") = P
        
            .Update
        End With
    End With
End Sub


Private Sub L1_KeyPress(KeyAscii As Integer)
    ischange = True
End Sub


Private Sub L2_KeyPress(KeyAscii As Integer)
    ischange = True
End Sub
Private Sub L3_KeyPress(KeyAscii As Integer)
    ischange = True
End Sub
Private Sub L4_KeyPress(KeyAscii As Integer)
    ischange = True
End Sub
Private Sub L5_KeyPress(KeyAscii As Integer)
    ischange = True
End Sub
Private Sub L6_KeyPress(KeyAscii As Integer)
    ischange = True
End Sub
Private Sub L7_KeyPress(KeyAscii As Integer)
    ischange = True
End Sub
Private Sub L8_KeyPress(KeyAscii As Integer)
    ischange = True
End Sub
Private Sub L9_KeyPress(KeyAscii As Integer)
    ischange = True
End Sub
Private Sub L10_KeyPress(KeyAscii As Integer)
    ischange = True
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        txtAmount.Text = Format(CDbl("0" & txtAmount.Text), "#,##0")
        txtAmount.SelStart = 9999
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub
