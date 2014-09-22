VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPhieuchi 
   Caption         =   "Phi’u chi"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9405
   ClipControls    =   0   'False
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
   Moveable        =   0   'False
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Phi’u chi trong th∏ng"
      Height          =   10335
      Left            =   8760
      TabIndex        =   37
      Top             =   120
      Width           =   6375
      Begin MSFlexGridLib.MSFlexGrid flgPhieuchi 
         Height          =   9975
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   17595
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraNoidung 
      Caption         =   "NÈi dung phi’u chi"
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
      Height          =   8775
      Left            =   30
      TabIndex        =   5
      Tag             =   "L4"
      Top             =   1680
      Width           =   8625
      Begin VB.TextBox txtDiengiai 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1980
         TabIndex        =   32
         Top             =   4200
         Width           =   6345
      End
      Begin VB.ComboBox cboKhoanchi 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1980
         TabIndex        =   17
         Text            =   "kho∂n chi"
         Top             =   630
         Width           =   2445
      End
      Begin VB.TextBox txtDiengiaichi 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4410
         TabIndex        =   16
         Top             =   600
         Width           =   3915
      End
      Begin VB.ComboBox cboVendor 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1980
         TabIndex        =   15
         Top             =   1170
         Width           =   2445
      End
      Begin VB.TextBox txtTenKH 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4410
         TabIndex        =   14
         Top             =   1140
         Width           =   3915
      End
      Begin VB.TextBox txtDiachi 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1980
         TabIndex        =   13
         Top             =   1680
         Width           =   6345
      End
      Begin VB.TextBox txtSoDT 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1980
         TabIndex        =   12
         Top             =   2280
         Width           =   2265
      End
      Begin VB.TextBox txtMST 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5610
         TabIndex        =   11
         Top             =   2280
         Width           =   2715
      End
      Begin VB.TextBox txtNguoinoptien 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1980
         TabIndex        =   10
         Top             =   2910
         Width           =   3615
      End
      Begin VB.TextBox txtBophan 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6660
         TabIndex        =   9
         Top             =   2910
         Width           =   1665
      End
      Begin VB.TextBox txtPhuongthuc 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1980
         TabIndex        =   8
         Text            =   "Ti“n m∆t"
         Top             =   3570
         Width           =   2475
      End
      Begin VB.TextBox txtSotien 
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5970
         TabIndex        =   7
         Top             =   3540
         Width           =   2355
      End
      Begin prjTouchScreen.MyButton cmdLuu 
         Height          =   1035
         Left            =   3360
         TabIndex        =   28
         Tag             =   "L16"
         Top             =   5520
         Width           =   2235
         _extentx        =   3942
         _extenty        =   1826
         btype           =   14
         tx              =   "&L≠u"
         enab            =   0
         font            =   "frmPhieuchi.frx":0000
         coltype         =   2
         focusr          =   -1
         bcol            =   16578804
         bcolo           =   16777152
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmPhieuchi.frx":0028
         picn            =   "frmPhieuchi.frx":0046
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdXem 
         Height          =   1035
         Left            =   5760
         TabIndex        =   29
         Tag             =   "L17"
         Top             =   5520
         Width           =   2235
         _extentx        =   3942
         _extenty        =   1826
         btype           =   14
         tx              =   "&Xem"
         enab            =   0
         font            =   "frmPhieuchi.frx":058A
         coltype         =   2
         focusr          =   -1
         bcol            =   16578804
         bcolo           =   16578804
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmPhieuchi.frx":05B2
         picn            =   "frmPhieuchi.frx":05D0
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdThoat 
         Height          =   1035
         Left            =   5760
         TabIndex        =   30
         Tag             =   "L18"
         Top             =   6720
         Width           =   2205
         _extentx        =   3889
         _extenty        =   1826
         btype           =   14
         tx              =   "Th&o∏t"
         enab            =   -1
         font            =   "frmPhieuchi.frx":2D84
         coltype         =   2
         focusr          =   -1
         bcol            =   16578804
         bcolo           =   16777152
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmPhieuchi.frx":2DAC
         picn            =   "frmPhieuchi.frx":2DCA
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmddelete 
         Height          =   1035
         Left            =   3360
         TabIndex        =   33
         Top             =   6720
         Width           =   2235
         _extentx        =   3942
         _extenty        =   1826
         btype           =   14
         tx              =   "&X„a"
         enab            =   -1
         font            =   "frmPhieuchi.frx":9066
         coltype         =   2
         focusr          =   -1
         bcol            =   16578804
         bcolo           =   16578804
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmPhieuchi.frx":908E
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdPrint 
         Height          =   1035
         Left            =   960
         TabIndex        =   34
         Top             =   6720
         Width           =   2235
         _extentx        =   3942
         _extenty        =   1826
         btype           =   14
         tx              =   "In khÊ 80mm"
         enab            =   -1
         font            =   "frmPhieuchi.frx":90AC
         coltype         =   2
         focusr          =   -1
         bcol            =   16578804
         bcolo           =   16578804
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmPhieuchi.frx":90D4
         picn            =   "frmPhieuchi.frx":90F2
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin prjTouchScreen.MyButton cmdCreat 
         Height          =   1035
         Left            =   960
         TabIndex        =   36
         Top             =   5520
         Width           =   2235
         _extentx        =   3942
         _extenty        =   1826
         btype           =   14
         tx              =   "&Tπo mÌi"
         enab            =   -1
         font            =   "frmPhieuchi.frx":B8A6
         coltype         =   2
         focusr          =   -1
         bcol            =   16578804
         bcolo           =   16777152
         fcol            =   16711680
         fcolo           =   255
         mcol            =   12632256
         mptr            =   1
         micon           =   "frmPhieuchi.frx":B8CE
         umcol           =   -1
         soft            =   0
         picpos          =   0
         ngrey           =   0
         fx              =   0
         hand            =   0
         check           =   0
         value           =   0
      End
      Begin VB.Label lblDiengiai 
         Alignment       =   1  'Right Justify
         Caption         =   "Di‘n gi∂i:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   31
         Tag             =   "L14"
         Top             =   4200
         Width           =   1785
      End
      Begin VB.Label lblKhoanthu 
         Alignment       =   1  'Right Justify
         Caption         =   "Kho∂n chi:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   360
         TabIndex        =   6
         Tag             =   "L5"
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Nhµ cung c p:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   300
         TabIndex        =   18
         Tag             =   "L6"
         Top             =   1170
         Width           =   1605
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "ßﬁa chÿ:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   450
         TabIndex        =   19
         Tag             =   "L7"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "SË ßT:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   180
         TabIndex        =   27
         Tag             =   "L8"
         Top             =   2310
         Width           =   1725
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "M∑ sË thu’:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4320
         TabIndex        =   26
         Tag             =   "L9"
         Top             =   2310
         Width           =   1275
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Ng≠Íi nhÀn ti“n:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   25
         Tag             =   "L10"
         Top             =   2940
         Width           =   1785
      End
      Begin VB.Label Label8 
         Caption         =   "BÈ phÀn:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5670
         TabIndex        =   24
         Tag             =   "L11"
         Top             =   2940
         Width           =   1005
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Ph≠¨ng th¯c chi:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   23
         Tag             =   "L12"
         Top             =   3600
         Width           =   1785
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "SË ti“n:"
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
         Left            =   4560
         TabIndex        =   22
         Tag             =   "L13"
         Top             =   3660
         Width           =   1125
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Bªng ch˜:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   510
         TabIndex        =   21
         Tag             =   "L15"
         Top             =   4830
         Width           =   1275
      End
      Begin VB.Label lblChu 
         Caption         =   "Thµnh ti“n bªng ch˜"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1920
         TabIndex        =   20
         Top             =   4920
         Width           =   6645
      End
   End
   Begin VB.TextBox txtNgay 
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
      Height          =   465
      Left            =   480
      TabIndex        =   3
      Top             =   960
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox txtSophieu 
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
      Height          =   435
      Left            =   2880
      TabIndex        =   2
      Top             =   1110
      Width           =   2565
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   495
      Left            =   6360
      TabIndex        =   35
      Top             =   1080
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   61276161
      UpDown          =   -1  'True
      CurrentDate     =   40553
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ngµy:"
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
      Height          =   315
      Left            =   5640
      TabIndex        =   4
      Tag             =   "L3"
      Top             =   1200
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "SË phi’u:"
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
      Left            =   1770
      TabIndex        =   1
      Tag             =   "L2"
      Top             =   1170
      Width           =   1065
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "phi’u chi"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   3240
      TabIndex        =   0
      Tag             =   "L1"
      Top             =   60
      Width           =   3135
   End
End
Attribute VB_Name = "frmPhieuchi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsPhieuChi As New ADODB.Recordset
Dim rsKhoanchi As New ADODB.Recordset
Dim rsVendor As New ADODB.Recordset

Private Sub cboVendor_Change()
On Error GoTo Handle

    With rsVendor
        .Find "Vendor_Number='" & cboVendor.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtTenKH.Text = .Fields("Vendor_Name") & ""
                txtDiachi.Text = .Fields("Address_1") & ""
                txtSoDT.Text = .Fields("Phone") & ""
                txtMST.Text = .Fields("Vendor_Tax_ID") & ""
            End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "cboVendor_Change"
End Sub

Private Sub cboVendor_Click()
Call cboVendor_Change
End Sub

Private Sub cboVendor_GotFocus()
    SendKeys "%{DOWN}"
End Sub

Private Sub cboVendor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtNguoinoptien.SetFocus
    End If
End Sub

Private Sub cboKhoanchi_Change()
On Error GoTo Handle

    With rsKhoanchi
        .Find "Machi='" & cboKhoanchi.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtDiengiaichi.Text = .Fields("DienGiai")
            End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "cboKhoanchi_Change"
End Sub

Private Sub cboKhoanchi_Click()
    Call cboKhoanchi_Change
End Sub

Private Sub cboKhoanchi_GotFocus()
    SendKeys "%{DOWN}"
End Sub

Private Sub cboKhoanchi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboVendor.SetFocus
    End If
End Sub

Private Sub cmdCreat_Click()
On Error GoTo Handle
    txtSophieu.Enabled = True
    txtNgay.Text = gfCONVERT_STRING_TO_DATE(DateDefault)
    txtSophieu.Text = GetMaxSophieu
    Call Clear_Text
    Call Add_Combo_NCC
    Call Add_Combo_Khoanchi
    cmdLuu.Enabled = True
    cmdCreat.Enabled = False
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdCreat_Click"
End Sub

Private Sub cmddelete_Click()
On Error GoTo Handle
If UserLevel <> 1 Then
    MsgBox "Bπn kh´ng c„ quy“n x„a phi’u"
    Exit Sub
End If
    With rsPhieuChi
        If MsgBox("Bπn c„ muËn x„a phi’u nµy !?", vbYesNo) = vbYes Then
            .Find "ID='" & Trim(txtSophieu.Text) & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Delete adAffectCurrent
                .Requery
            End If
        End If
    End With
    Call Set_FlgReceipt
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "cmddelete_Click"
End Sub

Private Sub cmdLuu_Click()
    On Error GoTo Handle
    Dim ans As Integer
    ans = MsgBox("Bπn c„ muËn l≠u phi’u nµy kh´ng?", vbYesNo)
    If cboKhoanchi.Text = "" Then
        MsgBox "bπn ph∂i ch‰n kho∂n thu"
        Exit Sub
    End If
    If cboVendor.Text = "" Then
        MsgBox "Bπn ph∂i ch‰n kh∏ch hµng nhÀn ti“n"
        Exit Sub
    End If
    If ans = vbYes Then
    
        With rsPhieuChi
            .Find "ID='" & txtSophieu.Text & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("ID") = Trim(txtSophieu.Text)
                .Fields("Store_ID") = Store_ID
                .Fields("Cashier_ID") = UserID
                .Fields("DateTime") = gfCONVERT_DATE_TO_STRING(dtpDate.Value)
                .Fields("Expense_ID") = Trim(cboKhoanchi.Text)
                .Fields("Vendor_Number") = Trim(cboVendor.Text)
                .Fields("Recieve_Name") = txtNguoinoptien.Text
                .Fields("Division") = txtBophan.Text
                .Fields("Payment_Method") = txtPhuongthuc.Text
                .Fields("Amount") = CDbl("0" & txtSotien.Text)
                .Fields("Description") = txtDiengiai.Text
                .Update
            End If
        End With
    End If
    Call Set_FlgReceipt
    cmdXem.Enabled = True
    cmdCreat.Enabled = True
    cmdLuu.Enabled = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   cmdLuu_Click"
End Sub

Private Sub cmdPrint_Click()
On Error GoTo Handle
    Dim iReport As CRAXDDRT.Report
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = "SELECT Payouts.ID, Payouts.Store_ID, Left([DateTime],8) AS DateRe," & _
    " Payouts.Vendor_Number, Payouts.Amount, Payouts.Description," & _
    " Payouts.Payment_Method, Payouts.Expense_ID, Payouts.Recieve_Name," & _
    " Payouts.Division " & _
    " FROM Payouts" & _
    " WHERE ID='" & txtSophieu.Text & "'"
    Set crPhieuchi80 = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crPhieuchi80
        .Database.AddADOCommand cnData, cmd
        .txtSophieu.SetUnboundFieldSource "{ado.ID}"
        .txtNgaythu.SetUnboundFieldSource "{ado.DateRe}"
        .txtMathu.SetUnboundFieldSource "{ado.Expense_ID}"
        .txtNguoinop.SetUnboundFieldSource "{ado.Recieve_Name}"
        .txtDescription.SetUnboundFieldSource "{ado.Description}"
        .txtSotien.SetUnboundFieldSource "{ado.Amount}"
        
        With .txtSotien
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        'canh le
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign
    End With
    
    Set iReport = crPhieuchi80
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = iReport
        .Show vbModal
    End With
Exit Sub
Handle:
    Exit Sub
    MsgBox Err.Number & Err.Description & Me.name & " cmdPrint_Click"
    
End Sub

Private Sub cmdthoat_Click()
    Unload Me
End Sub

Private Sub cmdXem_Click()
       On Error GoTo Handle
    Dim iReport As CRAXDDRT.Report
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = "SELECT Payouts.ID, Payouts.Store_ID, Left([DateTime],8) AS DateRe," & _
    " Payouts.Vendor_Number, Payouts.Amount, Payouts.Description," & _
    " Payouts.Payment_Method, Payouts.Expense_ID, Payouts.Recieve_Name," & _
    " Payouts.Division " & _
    " FROM Payouts" & _
    " WHERE ID='" & txtSophieu.Text & "'"
    Set crPhieuchi = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crPhieuchi
        .Database.AddADOCommand cnData, cmd
        .txtSophieu.SetUnboundFieldSource "{ado.ID}"
        .txtNgaythu.SetUnboundFieldSource "{ado.DateRe}"
        .txtMathu.SetUnboundFieldSource "{ado.Expense_ID}"
        .txtKH.SetUnboundFieldSource "{ado.Vendor_Number}"
        .txtNguoinop.SetUnboundFieldSource "{ado.Recieve_Name}"
        .txtBophan.SetUnboundFieldSource "{ado.Division}"
        .txtHTTT.SetUnboundFieldSource "{ado.Payment_Method}"
        .txtDescription.SetUnboundFieldSource "{ado.Description}"
        .txtSotien.SetUnboundFieldSource "{ado.Amount}"
    End With
    Set iReport = crPhieuchi
    With frmShowthuchi
        .Report = iReport
        .Show vbModal
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdXem_Click"
End Sub

Private Sub dtpDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    txtNgay.Text = dtpDate.Value
End Sub

Private Sub dtpDate_Change()
    txtNgay.Text = dtpDate.Value
    txtSophieu.Text = GetMaxSophieu
End Sub

Private Sub dtpDate_Click()
 txtNgay.Text = dtpDate.Value
End Sub

Private Sub flgPhieuchi_Click()
    txtSophieu.Text = flgPhieuchi.TextMatrix(flgPhieuchi.Row, 0)
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    Dim ctrl As Control
    Dim DescArr() As String
    If cmdLuu.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
        DescArr = LoadLanguage(LngFile, "#03:021:")
        Me.Caption = DescArr(1)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
        Next ctrl
        If UserLevel <> 1 Then CheckRight
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"

End Sub

Private Sub Form_Load()
On Error GoTo Handle
'If cnData.State = 0 Then
'    Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'End If
    Set rsKhoanchi = OpenCriticalTable("Select * from Expense", cnData)
    Set rsVendor = OpenCriticalTable("select * from Vendors", cnData)
    Set rsPhieuChi = OpenCriticalTable("select * from Payouts", cnData)
    dtpDate.Value = Date
Call Locktextbox
Call Set_flgHeader
Call Set_FlgReceipt
Exit Sub
Handle:
MsgBox Err.ne & Err.Description & Me.name & "Form_Load"
End Sub

Public Sub CheckRight()
On Error GoTo Handle
    Dim res As New ADODB.Recordset
        
        Set res = LoadPasswordData
        With MyRight
            res.MoveFirst
            Do While Not res.EOF
                If StrComp(res.Fields("UserName"), userName, 1) = 0 Then
                    .FullRight = res.Fields("UserRight")
                    .Thuchi = RightDeCode(Mid(.FullRight, 513, 64))
                    Exit Do
                End If
                res.MoveNext
            Loop
            If Mid(.Thuchi, 6, 1) = 0 Then
                  cmdCreat.Enabled = False
            Else: cmdCreat.Enabled = True
            End If
            If Mid(.Thuchi, 7, 1) = 0 Then
                  cmdDelete.Enabled = False
            Else: cmdDelete.Enabled = True
            End If
            

        End With

    CloseRecordset res
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " CheckRight"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsKhoanchi = Nothing
    Set rsVendor = Nothing
    Set rsPhieuChi = Nothing
End Sub



Private Sub txtBophan_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        
        .PasswordChar = ""
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtBophan.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtBophan_DblClick"
End Sub

Private Sub txtBophan_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtSotien.SetFocus
    End If
End Sub

Private Sub txtDienGiai_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        
        .PasswordChar = ""
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtDiengiai.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtDiengiai_DblClick"
End Sub

Private Sub txtDiengiai_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdLuu.SetFocus
End If
End Sub

Private Sub txtDiengiaichi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        cboVendor.SetFocus
    End If
End Sub



Private Sub txtNgay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboKhoanchi.SetFocus
 End If
End Sub

Private Sub txtNguoinoptien_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        
        .PasswordChar = ""
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtNguoinoptien.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtNguoinoptien"
End Sub

Private Sub txtNguoinoptien_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtBophan.SetFocus
    End If
End Sub

Private Sub txtSophieu_Change()
On Error GoTo Handle
    With rsPhieuChi
        .Find "ID='" & txtSophieu & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtNgay.Text = gfCONVERT_STRING_TO_DATE(Left(.Fields("DateTime"), 10))
            cboKhoanchi.Text = .Fields("Expense_ID")
            cboVendor.Text = .Fields("Vendor_Number")
            txtNguoinoptien.Text = .Fields("Recieve_Name")
            txtBophan.Text = .Fields("Division")
            txtPhuongthuc.Text = .Fields("Payment_Method")
            txtSotien.Text = Format(.Fields("Amount"), formatNum)
            txtDiengiai.Text = .Fields("Description")
            cmdXem.Enabled = True
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtSophieu_Change "
End Sub

Private Sub txtSotien_Change()
    lblChu.Caption = readnumber(CDbl("0" & txtSotien.Text)) & " ÆÂng./."
End Sub


Private Sub txtSotien_DblClick()
On Error GoTo Handle
    With frmKeyboard.txtInput
        
        .PasswordChar = ""
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        With frmKeyboard
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtSotien.Text = .Let_Text_Input
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  txtSotien_DblClick"
End Sub

Private Sub txtSotien_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtSotien.Text = Format(txtSotien.Text, formatNum)
lblChu.Caption = readnumber(CDbl("0" & txtSotien.Text)) & " ÆÂng./."

    txtDiengiai.SetFocus
 End If
End Sub

Private Sub txtTenKH_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNguoinoptien.SetFocus
    End If
End Sub

Public Sub Locktextbox()
    On Error GoTo Handle
        txtSoDT.Locked = True
        txtNgay.Locked = True
        txtDiengiaichi.Locked = True
        txtTenKH.Locked = True
        txtDiachi.Locked = True
        txtSoDT.Locked = True
        txtMST.Locked = True
        
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   Locktextbox"
End Sub

Public Function GetMaxSophieu() As String
On Error GoTo Handle
    Dim rsmax As New ADODB.Recordset
    Dim date_Payout As String
    date_Payout = gfCONVERT_DATE_TO_STRING(dtpDate.Value)
    
    Set rsmax = OpenCriticalTable("select max(ID) as MaxID from Payouts where Substring(DateTime,5,2)='" & Format(Month(dtpDate.Value), "00") & "'", cnData)
    If Not rsmax.EOF Then
    If "" & rsmax.Fields("maxiD") = "" Then
        GetMaxSophieu = "PC/" & Mid(date_Payout, 5, 2) & Mid(date_Payout, 3, 2) & "0001"
    Else
        GetMaxSophieu = Left(rsmax.Fields("MaxID"), Len(rsmax.Fields("MaxID")) - 4) & Right("0000" & (CDbl(Right(rsmax.Fields("MaxID"), 4)) + 1), 4)
    End If
    Else
        GetMaxSophieu = "PC/" & Mid(date_Payout, 5, 2) & Mid(date_Payout, 3, 2) & "0001"
    End If
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   GetMaxSophieu"
End Function

Public Sub Set_FlgPhieuChi(rs As ADODB.Recordset)
On Error GoTo Handle
        Dim incount As Integer
        If rs.State <> 0 Then
            If rs.RecordCount > 0 Then
                rs.MoveFirst
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        With rs
            .Sort = "DateTime Desc"
            Do While Not .EOF
                incount = incount + 1
                flgPhieuchi.Rows = rs.RecordCount + 1
                With flgPhieuchi
                    .TextMatrix(incount, 0) = rs!ID
                    .TextMatrix(incount, 1) = rs!DienGiai
                    .TextMatrix(incount, 2) = gfCONVERT_STRING_TO_DATE(rs!DateTime)
                    .TextMatrix(incount, 3) = rs!Vendor_Name
                    .TextMatrix(incount, 4) = rs!Recieve_Name & ""
                    .TextMatrix(incount, 5) = Format(rs!Amount, formatNum)
                    .TextMatrix(incount, 6) = rs!Description & " "
                End With
            rs.MoveNext
            Loop
        End With
        If rs.RecordCount = 0 Then
            With flgPhieuchi
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
                .TextMatrix(1, 4) = ""
                .TextMatrix(1, 5) = ""
            End With
        End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "Set_FlgPhieuChi"
End Sub
Public Sub Set_flgHeader()
    On Error GoTo Handle
        With flgPhieuchi
            .Cols = 7
            .Rows = 20
            .ColWidth(0) = 1600
            .ColWidth(1) = 1300
            .ColWidth(2) = 1400
            .ColWidth(3) = 1500
            .ColWidth(4) = 1500
            .ColWidth(5) = 1200
            .ColWidth(6) = 2000
            .TextMatrix(0, 0) = "SË phi’u"
            .TextMatrix(0, 1) = "Kho∂n Chi"
            .TextMatrix(0, 2) = "Ngµy chi"
            .TextMatrix(0, 3) = "Nhµ CC"
            .TextMatrix(0, 4) = "Ng≠Íi nhÀn"
            .TextMatrix(0, 5) = "SË ti“n"
            .TextMatrix(0, 6) = "Di‘n gi∂i thu"
            .ColAlignment(1) = 2
            .ColAlignment(2) = 4
            .ColAlignment(3) = 6
            .ColAlignment(4) = 6
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
            .TextMatrix(1, 4) = ""
            .TextMatrix(1, 5) = ""
            .TextMatrix(1, 6) = ""
            
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Set_flgHeader"
End Sub

Public Sub Set_FlgReceipt()
On Error GoTo Handle
    Dim strFilter As String
    Dim rsReceipt_Inmonth As New ADODB.Recordset
    strFilter = "SELECT Payouts.ID, Expense.DienGiai, Payouts.DateTime," & _
                " Vendors.Vendor_Name, Payouts.Recieve_Name, Payouts.Amount," & _
                " Payouts.Description" & _
                " FROM (Expense INNER JOIN Payouts ON Expense.MaChi = Payouts.Expense_ID)" & _
                " INNER JOIN Vendors ON Payouts.Vendor_Number = Vendors.Vendor_Number" & _
               " where left(Payouts.DateTime,6)='" & Format(Year(Date), "0000") & Format(Month(Date), "00") & "'" & _
               " ORDER BY Payouts.DateTime"
    Set rsReceipt_Inmonth = OpenCriticalTable(strFilter, cnData)
    Call Set_FlgPhieuChi(rsReceipt_Inmonth)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Set_FlgPhieuThu"
End Sub

Public Sub Add_Combo_NCC()
On Error GoTo Handle
If rsVendor.State <> 0 And rsVendor.RecordCount > 0 Then rsVendor.MoveFirst
'Gan list Khach hang
With cboVendor
    .Clear
    Do While Not rsVendor.EOF
        .AddItem rsVendor.Fields("Vendor_Number")
        rsVendor.MoveNext
    Loop
End With

Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Add_Combo_NCC"

End Sub

Public Sub Add_Combo_Khoanchi()
On Error GoTo Handle
If rsKhoanchi.State <> 0 And rsKhoanchi.RecordCount > 0 Then rsKhoanchi.MoveFirst
'Gan list Khoan chi
With cboKhoanchi
    .Clear
    Do While Not rsKhoanchi.EOF
        .AddItem rsKhoanchi.Fields("Machi")
        rsKhoanchi.MoveNext
    Loop
End With

Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Add_Combo_Khoanchi"
End Sub

Public Sub Clear_Text()
On Error GoTo Handle
    txtDiengiaichi.Text = ""
    txtTenKH.Text = ""
    txtDiachi.Text = ""
    txtSoDT.Text = ""
    txtMST.Text = ""
    txtNguoinoptien.Text = ""
    txtBophan.Text = ""
    txtSotien.Text = ""
    txtDiengiai.Text = ""
    lblChu.Caption = ""
    cboVendor.Text = ""
    cboKhoanchi.Text = ""
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Clear_Text"
End Sub


