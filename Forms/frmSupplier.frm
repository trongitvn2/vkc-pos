VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSupplier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Danh môc nhµ cung cÊp"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
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
   Icon            =   "frmSupplier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraButton2 
      Height          =   1395
      Left            =   8880
      TabIndex        =   29
      Top             =   9360
      Width           =   6495
      Begin prjTouchScreen.MyButton cmdHelp 
         Height          =   1155
         Left            =   3900
         TabIndex        =   30
         Tag             =   "L16"
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   2037
         BTYPE           =   14
         TX              =   "&Gióp ®ì"
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
         MICON           =   "frmSupplier.frx":000C
         PICN            =   "frmSupplier.frx":0028
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
         Cancel          =   -1  'True
         Height          =   1155
         Left            =   5220
         TabIndex        =   31
         Tag             =   "L17"
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   2037
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
         MICON           =   "frmSupplier.frx":0662
         PICN            =   "frmSupplier.frx":067E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdCancel 
         Height          =   1155
         Left            =   2700
         TabIndex        =   32
         Tag             =   "L18"
         Top             =   180
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   2037
         BTYPE           =   14
         TX              =   "&Hñy bá"
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
         MICON           =   "frmSupplier.frx":6918
         PICN            =   "frmSupplier.frx":6934
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
         Height          =   1155
         Left            =   1380
         TabIndex        =   33
         Tag             =   "L14"
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   2037
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
         MICON           =   "frmSupplier.frx":6B0E
         PICN            =   "frmSupplier.frx":6B2A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdThem 
         Height          =   1155
         Left            =   60
         TabIndex        =   34
         Tag             =   "L13"
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   2037
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
         MICON           =   "frmSupplier.frx":7164
         PICN            =   "frmSupplier.frx":7180
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
      Left            =   8940
      ScaleHeight     =   945
      ScaleWidth      =   6225
      TabIndex        =   20
      Top             =   60
      Width           =   6285
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Group Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   90
         TabIndex        =   22
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label lblNo 
         Alignment       =   2  'Center
         BackColor       =   &H80000008&
         Caption         =   "Group No"
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   90
         TabIndex        =   21
         Top             =   45
         Width           =   5985
      End
   End
   Begin TabDlg.SSTab TabSupplier 
      Height          =   7905
      Left            =   8880
      TabIndex        =   13
      Top             =   1290
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   13944
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   3528
      TabCaption(0)   =   "Th«ng tin nhµ Cung cÊp"
      TabPicture(0)   =   "frmSupplier.frx":75D2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraIn"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame fraIn 
         Height          =   7215
         Left            =   90
         TabIndex        =   14
         Top             =   435
         Width           =   6165
         Begin VB.TextBox txtSTK 
            Height          =   495
            Left            =   3810
            TabIndex        =   8
            Top             =   4800
            Width           =   2235
         End
         Begin VB.TextBox txtMST 
            Height          =   495
            Left            =   1110
            TabIndex        =   7
            Top             =   4830
            Width           =   2355
         End
         Begin VB.TextBox txtFax 
            Height          =   495
            Left            =   3840
            TabIndex        =   6
            Top             =   3990
            Width           =   2205
         End
         Begin VB.TextBox txtPhone 
            Height          =   495
            Left            =   1110
            TabIndex        =   5
            Top             =   3990
            Width           =   2355
         End
         Begin VB.TextBox txtMail 
            Height          =   495
            Left            =   1110
            TabIndex        =   9
            Top             =   5580
            Width           =   4935
         End
         Begin VB.TextBox txtWebsite 
            Height          =   495
            Left            =   1110
            TabIndex        =   10
            Top             =   6300
            Width           =   4935
         End
         Begin VB.TextBox txtCompany 
            Height          =   495
            Left            =   1110
            TabIndex        =   2
            Top             =   1950
            Width           =   4935
         End
         Begin VB.TextBox txtAdd2 
            Height          =   495
            Left            =   1110
            TabIndex        =   4
            Top             =   3150
            Width           =   4935
         End
         Begin VB.TextBox txtAdd1 
            Height          =   495
            Left            =   1110
            TabIndex        =   3
            Top             =   2520
            Width           =   4935
         End
         Begin VB.TextBox txtSupCode 
            Height          =   525
            Left            =   1110
            TabIndex        =   0
            Top             =   570
            Width           =   2775
         End
         Begin VB.TextBox txtSupName 
            Height          =   495
            Left            =   1110
            TabIndex        =   1
            Top             =   1410
            Width           =   4935
         End
         Begin VB.Label lblSTK 
            Alignment       =   1  'Right Justify
            Caption         =   "Sè tµi kho¶n"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   28
            Tag             =   "L10"
            Top             =   4500
            Width           =   1215
         End
         Begin VB.Label lblMST 
            Alignment       =   1  'Right Justify
            Caption         =   "M· sè thuÕ :"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   27
            Tag             =   "L9"
            Top             =   4530
            Width           =   1095
         End
         Begin VB.Label lblFax 
            Caption         =   "Fax:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2970
            TabIndex        =   26
            Tag             =   "L8"
            Top             =   3720
            Width           =   855
         End
         Begin VB.Label lblPhone 
            Caption         =   "§iÖn tho¹i:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Tag             =   "L7"
            Top             =   3750
            Width           =   1215
         End
         Begin VB.Label lblMail 
            Caption         =   "Email:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   24
            Tag             =   "L11"
            Top             =   5340
            Width           =   1125
         End
         Begin VB.Label lblWebsite 
            Caption         =   "Website:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Tag             =   "L12"
            Top             =   6180
            Width           =   1335
         End
         Begin VB.Label lblCompany 
            Caption         =   "C«ng ty:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   120
            TabIndex        =   19
            Tag             =   "L4"
            Top             =   1800
            Width           =   915
         End
         Begin VB.Label Label1 
            Caption         =   "§Þa chØ 2:"
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
            TabIndex        =   18
            Tag             =   "L6"
            Top             =   2940
            Width           =   1005
         End
         Begin VB.Label lblAdd 
            Caption         =   "§Þa chØ 1:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Tag             =   "L5"
            Top             =   2280
            Width           =   1065
         End
         Begin VB.Label lblSupCode 
            Caption         =   "M· nhµ cung cÊp"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Tag             =   "L2"
            Top             =   210
            Width           =   2745
         End
         Begin VB.Label lblSupName 
            Caption         =   "Tªn nhµ cung cÊp"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   120
            TabIndex        =   15
            Tag             =   "L3"
            Top             =   1140
            Width           =   2505
         End
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Th«ng tin nhµ cung cÊp"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1680
         TabIndex        =   35
         Top             =   120
         Width           =   2895
      End
   End
   Begin VB.Frame fraSup 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10185
      Left            =   30
      TabIndex        =   11
      Tag             =   "L1"
      Top             =   900
      Width           =   8745
      Begin MSFlexGridLib.MSFlexGrid flgSupplier 
         Height          =   9825
         Left            =   60
         TabIndex        =   12
         Top             =   300
         Width           =   8625
         _ExtentX        =   15214
         _ExtentY        =   17330
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "danh môc nhµ cung cÊp"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Left            =   480
      TabIndex        =   36
      Top             =   120
      Width           =   8295
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsSupplier As New ADODB.Recordset
Dim strPath As String
Dim DescArr() As String

Private Sub cmdCancel_Click()
    On Error GoTo Handle
        Call DeleteTextbox
        cmdThem.Caption = DescArr(13)
        cmdCancel.Enabled = False
        Call flgSupplier_EnterCell
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "cmdCancel_Click"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdThem_Click()
On Error GoTo Handle
    
    If cmdThem.Caption = DescArr(13) Then
        Call DeleteTextbox
        cmdCancel.Enabled = True
    ElseIf cmdThem.Caption = DescArr(15) Then
            Call UpdateDatabase
    End If
Exit Sub
Handle:
MsgBox Err.Number & " " & Err.Description & " " & _
            Me.name & " " & "cmdThem _click"

End Sub

Private Sub cmdXoa_Click()
    On Error GoTo Handle
    Dim rsInventoryB As New ADODB.Recordset
    Dim strSql As String
    If txtSupCode.Text <> "0000" Then
        strSql = "select * from Instock_MasterB where Vendor_Number ='" & flgSupplier.TextMatrix(flgSupplier.Row, 0) & "'"
        Set rsInventoryB = OpenCriticalTable(strSql, cnData)
        If rsInventoryB.RecordCount = 0 Then
            With rsSupplier
                .Find "Vendor_Number='" & flgSupplier.TextMatrix(flgSupplier.Row, 0) & _
                        "'", , adSearchForward, adBookmarkFirst
                
                If Not .EOF Or .BOF Then
                    .Delete adAffectCurrent
                    .MoveNext
                    .Requery
                End If
                Call Form_Load
            End With
        Else
            MsgBox "Nhµ cung cÊp nµy ®ang sö dông, kh«ng thÓ xãa"
        End If
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "cmdXoa_Click"
End Sub

Private Sub flgSupplier_EnterCell()
    On Error GoTo Handle
    
    With rsSupplier
        .Find "Vendor_Number='" & flgSupplier.TextMatrix(flgSupplier.Row, 0) & _
                "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtSupCode.Text = !Vendor_Number
            txtSupName.Text = !Vendor_Name
            txtCompany.Text = !Company & ""
            txtAdd1.Text = !Address_1 & ""
            txtAdd2.Text = !Address_2 & ""
            txtPhone.Text = !Phone & ""
            txtFax.Text = !Fax & ""
            txtMST.Text = !Vendor_Tax_ID & ""
            txtMail.Text = !Email & ""
            txtWebsite.Text = !website & ""
            lblNo.Caption = !Vendor_Number
            lblName.Caption = !Vendor_Name
        End If
    End With
    Exit Sub
    
Handle:
    MsgBox Err.Number & " " & Err.Description & vbCrLf _
    & Me.name & " flgSupplier_EnterCell"
End Sub

Private Sub Form_Activate()
    On Error GoTo Handle
    
    Dim ctrl As Control
'    If cmdCancel.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    DescArr = LoadLanguage(LngFile, "#01:014:")
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & Me.name & "Form_Active"
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    Dim str As String
    DescArr = LoadLanguage(LngFile, "#01:014:")
    str = "Select * from Vendors"
    Set rsSupplier = OpenCriticalTable(str, cnData)
    Call set_flgSupplier
    cmdCancel.Enabled = False
    Exit Sub
Handle:
    MsgBox Err.Number & " " & _
    Err.Description & Me.name & "Form_Load"
End Sub


Private Sub set_flgSupplier()
On Error GoTo errHdl
    Dim intCount    As Integer
    With flgSupplier
        .Cols = rsSupplier.Fields.count
        .Font = ".vnArial"
        .ColWidth(0) = 1400
        .ColWidth(1) = 3000
        .ColWidth(2) = 4000
        .ColWidth(3) = 1400
        .ColWidth(4) = 1400
        .ColWidth(5) = 1400
        .ColWidth(6) = 1400
        .ColWidth(7) = 1400
        .ColWidth(8) = 1400
        .ColWidth(9) = 1400
        .ColWidth(10) = 1400
        .TextMatrix(0, 0) = DescArr(2)
        .TextMatrix(0, 1) = DescArr(3)
        .TextMatrix(0, 2) = DescArr(4)
        .TextMatrix(0, 3) = DescArr(5)
        .TextMatrix(0, 5) = DescArr(6)
        .TextMatrix(0, 4) = DescArr(7)
        .TextMatrix(0, 6) = DescArr(8)
        .TextMatrix(0, 7) = DescArr(9)
        .TextMatrix(0, 8) = DescArr(10)
        .TextMatrix(0, 9) = DescArr(11)
        .TextMatrix(0, 10) = DescArr(12)
    End With
    
    If rsSupplier Is Nothing Then Exit Sub
    If rsSupplier.State = 0 Then Exit Sub
    
    If rsSupplier.EOF And rsSupplier.BOF Then
        With flgSupplier
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
        .TextMatrix(1, 10) = ""
        End With
        Exit Sub
    End If
   flgSupplier.Rows = rsSupplier.RecordCount + 1
    intCount = 0
    Do While Not rsSupplier.EOF
        intCount = intCount + 1
        flgSupplier.TextMatrix(intCount, 0) = rsSupplier!Vendor_Number
        flgSupplier.TextMatrix(intCount, 1) = rsSupplier!Vendor_Name
        flgSupplier.TextMatrix(intCount, 2) = rsSupplier!Company
        flgSupplier.TextMatrix(intCount, 3) = rsSupplier!Address_1
        flgSupplier.TextMatrix(intCount, 4) = rsSupplier!Address_2
        flgSupplier.TextMatrix(intCount, 5) = rsSupplier!Phone & ""
        flgSupplier.TextMatrix(intCount, 6) = rsSupplier!Fax & ""
        flgSupplier.TextMatrix(intCount, 7) = rsSupplier!Vendor_Tax_ID & ""
        flgSupplier.TextMatrix(intCount, 8) = rsSupplier!Vendor_AccNo & ""
        flgSupplier.TextMatrix(intCount, 9) = rsSupplier!Email & ""
        flgSupplier.TextMatrix(intCount, 10) = rsSupplier!website & ""
        rsSupplier.MoveNext
        
    Loop
'    SetColorFlexGrid flgSupplier, 1, 1, flgSupplier.Cols
    Call flgSupplier_EnterCell
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - set_flgSupplier "
End Sub

Private Sub DeleteTextbox()
    On Error GoTo Handle
        cmdThem.Caption = DescArr(15)
        txtSupCode.Text = ""
        txtSupName.Text = ""
        txtSTK.Text = ""
        txtAdd1.Text = ""
        txtAdd2.Text = ""
        txtCompany.Text = ""
        txtFax.Text = ""
        txtMST.Text = ""
        txtPhone.Text = ""
        txtMail.Text = ""
        txtWebsite.Text = ""
        txtSupCode.SetFocus
        
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "DeleteTextbox"
End Sub

Private Sub UpdateDatabase()
    On Error GoTo Handle
        With rsSupplier
            .Find "Vendor_Number='" & txtSupCode.Text & _
                "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("Vendor_Number") = txtSupCode.Text
                .Fields("Vendor_Name") = txtSupName.Text
                .Fields("Company") = txtCompany.Text & ""
                .Fields("Address_1") = txtAdd1.Text
                .Fields("Address_2") = txtAdd2.Text
                .Fields("Phone") = txtPhone.Text
                .Fields("Fax") = txtFax.Text
                .Fields("Vendor_Tax_ID") = txtMST.Text
                .Fields("Vendor_AccNo") = txtSTK.Text
                .Fields("EMail") = txtMail.Text
                .Fields("Website") = txtWebsite.Text
                .Update
                .Requery
            Else
                .Fields("Vendor_Number") = txtSupCode.Text
                .Fields("Vendor_Name") = txtSupName.Text
                .Fields("Company") = txtCompany.Text & ""
                .Fields("Address_1") = txtAdd1.Text
                .Fields("Address_2") = txtAdd2.Text
                .Fields("Phone") = txtPhone.Text
                .Fields("Fax") = txtFax.Text
                .Fields("Vendor_Tax_ID") = txtMST.Text
                .Fields("Vendor_AccNo") = txtSTK.Text
                .Fields("EMail") = txtMail.Text
                .Fields("Website") = txtWebsite.Text
                .Update
                .Requery
            End If
        End With
        Call set_flgSupplier
        cmdThem.Caption = DescArr(13)
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "UpdateDatabase"
End Sub

Private Sub CancelItem()
    On Error GoTo Handle
        cmdThem.Caption = DescArr(13)
        If txtSupCode.Text = "" Then
            txtSupCode.SetFocus
        ElseIf txtSupName.Text = "" Then
            txtSupName.SetFocus
        End If
        
            
    Exit Sub
Handle:
    MsgBox Err.Number & "" & Err.Description & Me.name & " CancelItem"
End Sub

Private Sub LoadControl()
    On Error GoTo Handle
    
    With rsSupplier
        .Find "Vendor_Number='" & !Vendor_Number & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtSupCode.Text = !Vendor_Number
            txtSupName.Text = !Vendor_Name
            txtSTK.Text = !Vendor_AccNo
            txtAdd1.Text = !Address_1
            txtAdd2.Text = !Address_2
            txtPhone.Text = !Phone
            txtFax.Text = !Fax
            txtMST.Text = !Vendor_Tax_ID
            txtCompany.Text = !Company
            txtMail.Text = !Mail
            txtWebsite.Text = !website
            .Requery
        End If
    End With
    Exit Sub
    
Handle:
    MsgBox Err.Number & " " & Err.Description & vbCrLf _
    & Me.name & " LoadControl"
End Sub



Private Sub txtAdd1_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        .PasswordChar = ""
        .Text = txtAdd1.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .Show vbModal
            txtAdd1.Text = .Let_Text_Input
       End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub txtAdd1_KeyPress(KeyAscii As Integer)
 On Error GoTo Handle
    If KeyAscii = 13 Then
        txtAdd2.SetFocus
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtAdd1_KeyPress"

End Sub


Private Sub txtAdd2_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        .PasswordChar = ""
        .Text = txtAdd2.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .Show vbModal
            txtAdd2.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub txtAdd2_KeyPress(KeyAscii As Integer)
 On Error GoTo Handle
    If KeyAscii = 13 Then
        txtPhone.SetFocus
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtAdd2_KeyPress"

End Sub

Private Sub txtCompany_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        .PasswordChar = ""
        .Text = txtCompany.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .Show vbModal
            txtCompany.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub txtCompany_KeyPress(KeyAscii As Integer)
 On Error GoTo Handle
    If KeyAscii = 13 Then
        txtAdd1.SetFocus
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtCompany_KeyPress"
End Sub

Private Sub txtFax_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        .PasswordChar = ""
        .Text = txtFax.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .Show vbModal
            txtFax.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    Select Case KeyAscii
        Case 13
            txtMST.SetFocus
        Case Is < 32, 48 To 57, 44, 46
        Case Else: KeyAscii = 0
    End Select
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtFax_KeyPress"
End Sub

Private Sub txtMail_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        .PasswordChar = ""
        .Text = txtMail.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .Show vbModal
            txtMail.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub txtMail_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    Select Case KeyAscii
        Case 13
            txtWebsite.SetFocus
    End Select
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtMail_KeyPress"

End Sub

Private Sub txtMST_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        .PasswordChar = ""
        .Text = txtMST.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .Show vbModal
            txtMST.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub txtMST_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    Select Case KeyAscii
        Case 13
            txtSTK.SetFocus
        Case Is < 32, 48 To 57, 44, 46
        Case Else: KeyAscii = 0
    End Select
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtMST_KeyPress"
End Sub

Private Sub txtPhone_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        .PasswordChar = ""
        .Text = txtPhone.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .Show vbModal
            txtPhone.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub txtPhone_KeyPress(KeyAscii As Integer)
     On Error GoTo Handle
    Select Case KeyAscii
        Case 13
            txtPhone.SetFocus
        Case Is < 32, 48 To 57, 44, 46
        Case Else: KeyAscii = 0
    End Select
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtPhone_KeyPress"
End Sub

Private Sub txtSTK_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        .PasswordChar = ""
        .Text = txtSTK.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .Show vbModal
            txtSTK.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub txtSTK_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    Select Case KeyAscii
        Case 13
            txtMail.SetFocus
    End Select
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtPhone_KeyPress"
End Sub


Private Sub txtSupCode_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        .PasswordChar = ""
        .Text = txtSupCode.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .Show vbModal
            txtSupCode.Text = .Let_Text_Input
       End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub txtSupCode_KeyPress(KeyAscii As Integer)
    On Error GoTo Handle
    If KeyAscii = 13 Then
        txtSupName.SetFocus
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtSupCode_KeyPress"
End Sub


Private Sub txtSupName_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        .PasswordChar = ""
        .Text = txtSupName.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .Show vbModal
            txtSupName.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub txtSupName_KeyPress(KeyAscii As Integer)
 On Error GoTo Handle
    If KeyAscii = 13 Then
        txtCompany.SetFocus
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtSupName_KeyPress"
End Sub


Private Sub txtWebsite_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        .PasswordChar = ""
        .Text = txtWebsite.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .Show vbModal
            txtWebsite.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub txtWebsite_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    Select Case KeyAscii
        Case 13
            cmdThem.SetFocus
    End Select
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtWebsite_KeyPress"

End Sub
