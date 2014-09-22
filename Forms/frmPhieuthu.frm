VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPhieuthu 
   Caption         =   "Phi’u Thu"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
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
   Icon            =   "frmPhieuthu.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   94905.08
   ScaleMode       =   0  'User
   ScaleWidth      =   3.57753e5
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Danh s∏ch phi’u thu trong th∏ng"
      Height          =   10335
      Left            =   9000
      TabIndex        =   35
      Top             =   120
      Width           =   6255
      Begin MSFlexGridLib.MSFlexGrid flgPhieuthu 
         Height          =   9855
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   17383
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
   Begin VB.Frame fraNoidung 
      Caption         =   "NÈi dung phi’u thu"
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
      Height          =   8895
      Left            =   120
      TabIndex        =   5
      Tag             =   "L4"
      Top             =   1560
      Width           =   8865
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
         Top             =   4320
         Width           =   6705
      End
      Begin prjTouchScreen.MyButton cmdThoat 
         Height          =   915
         Left            =   6900
         TabIndex        =   30
         Tag             =   "L18"
         Top             =   6600
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   1614
         BTYPE           =   14
         TX              =   "Th&o∏t"
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
         MICON           =   "frmPhieuthu.frx":000C
         PICN            =   "frmPhieuthu.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdXem 
         Height          =   915
         Left            =   3570
         TabIndex        =   29
         Tag             =   "L17"
         Top             =   6600
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1614
         BTYPE           =   14
         TX              =   "&Xem"
         ENAB            =   0   'False
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
         BCOLO           =   16578804
         FCOL            =   16711680
         FCOLO           =   255
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPhieuthu.frx":62C2
         PICN            =   "frmPhieuthu.frx":62DE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdLuu 
         Height          =   915
         Left            =   1920
         TabIndex        =   28
         Tag             =   "L16"
         Top             =   6600
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1614
         BTYPE           =   14
         TX              =   "&L≠u"
         ENAB            =   0   'False
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
         MICON           =   "frmPhieuthu.frx":8A90
         PICN            =   "frmPhieuthu.frx":8AAC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
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
         TabIndex        =   25
         Top             =   3540
         Width           =   2715
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
         TabIndex        =   23
         Text            =   "Ti“n m∆t"
         Top             =   3570
         Width           =   2475
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
         TabIndex        =   21
         Top             =   2910
         Width           =   2025
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
         TabIndex        =   19
         Top             =   2910
         Width           =   3615
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
         TabIndex        =   17
         Top             =   2280
         Width           =   3075
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
         TabIndex        =   15
         Top             =   2310
         Width           =   2265
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
         Width           =   6705
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
         TabIndex        =   11
         Top             =   1140
         Width           =   4275
      End
      Begin VB.ComboBox cboKhachhang 
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
         TabIndex        =   10
         Text            =   "Combo2"
         Top             =   1170
         Width           =   2445
      End
      Begin VB.TextBox txtDiengiaithu 
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
         TabIndex        =   7
         Top             =   600
         Width           =   4275
      End
      Begin VB.ComboBox cboKhoanthu 
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
         Left            =   2010
         TabIndex        =   6
         Text            =   "kho∂n thu"
         Top             =   630
         Width           =   2445
      End
      Begin prjTouchScreen.MyButton cmddelete 
         Height          =   915
         Left            =   5235
         TabIndex        =   33
         Top             =   6600
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1614
         BTYPE           =   14
         TX              =   "&X„a"
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
         MICON           =   "frmPhieuthu.frx":8FF0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdCreat 
         Height          =   915
         Left            =   240
         TabIndex        =   37
         Top             =   6600
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1614
         BTYPE           =   14
         TX              =   "&Tπo mÌi"
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
         MICON           =   "frmPhieuthu.frx":900C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
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
         Top             =   4320
         Width           =   1785
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
         Height          =   825
         Left            =   1980
         TabIndex        =   27
         Top             =   4950
         Width           =   6765
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
         Left            =   630
         TabIndex        =   26
         Tag             =   "L15"
         Top             =   4950
         Width           =   1275
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
         TabIndex        =   24
         Tag             =   "L13"
         Top             =   3660
         Width           =   1125
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Ph≠¨ng th¯c TT:"
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
         TabIndex        =   22
         Tag             =   "L12"
         Top             =   3600
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
         TabIndex        =   20
         Tag             =   "L11"
         Top             =   2940
         Width           =   1005
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Ng≠Íi nÈp ti“n:"
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
         TabIndex        =   18
         Tag             =   "L10"
         Top             =   2940
         Width           =   1785
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
         TabIndex        =   16
         Tag             =   "L9"
         Top             =   2310
         Width           =   1275
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
         TabIndex        =   14
         Tag             =   "L8"
         Top             =   2310
         Width           =   1725
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
         TabIndex        =   12
         Tag             =   "L7"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Kh∏ch hµng:"
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
         TabIndex        =   9
         Tag             =   "L6"
         Top             =   1170
         Width           =   1605
      End
      Begin VB.Label lblKhoanthu 
         Alignment       =   1  'Right Justify
         Caption         =   "Kho∂n thu:"
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
         TabIndex        =   8
         Tag             =   ":5"
         Top             =   630
         Width           =   1545
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
      Left            =   7440
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1365
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
      Left            =   2940
      TabIndex        =   2
      Top             =   900
      Width           =   1965
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   495
      Left            =   5760
      TabIndex        =   34
      Top             =   840
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
      Format          =   65208321
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
      Left            =   4980
      TabIndex        =   4
      Tag             =   "L3"
      Top             =   990
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
      Left            =   1830
      TabIndex        =   1
      Tag             =   "L2"
      Top             =   960
      Width           =   1065
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "phi’u thu"
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
      Left            =   2760
      TabIndex        =   0
      Tag             =   "L1"
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmPhieuthu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsPhieuthu As New ADODB.Recordset
Dim rsKhoanThu As New ADODB.Recordset
Dim rsKhachHang As New ADODB.Recordset

Private Sub cboKhachhang_Change()
On Error GoTo Handle

    With rsKhachHang
        .Find "CustNum='" & cboKhachhang.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtTenKH.Text = .Fields("CustName")
                txtDiachi.Text = .Fields("Address")
                txtSoDT.Text = .Fields("Phone")
                txtMST.Text = .Fields("TaxCode")
            End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "cboKhoanthu_Change"
End Sub

Private Sub cboKhachhang_Click()
Call cboKhachhang_Change
End Sub

Private Sub cboKhachhang_GotFocus()
    SendKeys "%{DOWN}"
End Sub

Private Sub cboKhachhang_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtNguoinoptien.SetFocus
    End If
End Sub

Private Sub cboKhoanthu_Change()
On Error GoTo Handle

    With rsKhoanThu
        .Find "MaThu='" & cboKhoanthu.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                txtDiengiaithu.Text = .Fields("DienGiai")
            End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "cboKhoanthu_Change"
End Sub

Private Sub cboKhoanthu_Click()
    Call cboKhoanthu_Change
End Sub

Private Sub cboKhoanthu_GotFocus()
    SendKeys "%{DOWN}"
End Sub

Private Sub cboKhoanthu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboKhachhang.SetFocus
    End If
End Sub

Private Sub cmdClose_Click()

End Sub

Private Sub cmdCreat_Click()
On Error GoTo Handle
    txtNgay.Text = gfCONVERT_STRING_TO_DATE(DateDefault)
    txtSophieu.Text = GetMaxSophieuThu
    Call Clear_Text
    Call Add_Combo_Cust
    Call Add_Combo_Khoanthu
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
    If MsgBox("Bπn muËn x„a phi’u thu nµy ?!", vbYesNo) = vbYes Then
        With rsPhieuthu
            .Find "ID='" & txtSophieu.Text & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Delete adAffectCurrent
                .Requery
            End If
        End With
    End If
    Call Set_FlgReceipt
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  cmddelete_Click"
End Sub

Private Sub cmdLuu_Click()
    On Error GoTo Handle
    Dim ans As Integer
    ans = MsgBox("Bπn c„ muËn l≠u phi’u nµy kh´ng?", vbYesNo)
    If cboKhoanthu.Text = "" Then
        MsgBox "bπn ph∂i ch‰n kho∂n thu"
        Exit Sub
    End If
    If cboKhachhang.Text = "" Then
        MsgBox "Bπn ph∂i ch‰n kh∏ch hµng nhÀn ti“n"
        Exit Sub
    End If
    If ans = vbYes Then
        With rsPhieuthu
        .Find "ID='" & txtSophieu.Text & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
              .addNew
              .Fields("ID") = Trim(txtSophieu.Text)
              .Fields("Store_ID") = Store_ID
              .Fields("Cashier_ID") = UserID
              .Fields("DateTime") = gfCONVERT_DATE_TO_STRING(txtNgay.Text)
              .Fields("Receipt_ID") = Trim(cboKhoanthu.Text)
              .Fields("Customer_ID") = Trim(cboKhachhang.Text)
              .Fields("Reciever_Name") = txtNguoinoptien.Text
              .Fields("Division") = txtBophan.Text
              .Fields("Payment_Method") = txtPhuongthuc.Text
              .Fields("Amount") = CDbl("0" & txtSotien.Text)
              .Fields("Description") = txtDiengiai.Text
              .Update
            End If
        End With
        
        'cap nhat cong no khach hang
        With rsKhachHang
            If Not .EOF And .RecordCount > 0 Then .MoveFirst
            .Find "CustNum='" & Trim(cboKhachhang.Text) & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Fields("Acct_Balance") = CDbl(.Fields("Acct_Balance")) - CDbl(txtSotien.Text)
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
    SQL = "SELECT Income.ID, left(Income.DateTime,10) as DatePay, Income.Customer_ID, Income.Receipt_ID," & _
        " Income.Division, Income.Reciever_Name, Income.Amount, Income.Description," & _
        " Income.Payment_Method " & _
        " FROM Income" & _
        " where ID='" & txtSophieu.Text & "'"
    Set crPhieuthu = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crPhieuthu
        .Database.AddADOCommand cnData, cmd
        .txtSophieu.SetUnboundFieldSource "{ado.ID}"
        .txtNgaythu.SetUnboundFieldSource "{ado.DatePay}"
        .txtMathu.SetUnboundFieldSource "{ado.Receipt_ID}"
        .txtKH.SetUnboundFieldSource "{ado.Customer_ID}"
        .txtNguoinop.SetUnboundFieldSource "{ado.Reciever_Name}"
        .txtBophan.SetUnboundFieldSource "{ado.Division}"
        .txtHTTT.SetUnboundFieldSource "{ado.Payment_Method}"
        .txtDescription.SetUnboundFieldSource "{ado.Description}"
        .txtSotien.SetUnboundFieldSource "{ado.Amount}"
    End With
    Set iReport = crPhieuthu
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
End Sub

Private Sub dtpDate_Click()
     txtNgay.Text = dtpDate.Value
End Sub

Private Sub flgPhieuthu_Click()
    txtSophieu.Text = flgPhieuthu.TextMatrix(flgPhieuthu.Row, 0)
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    Dim ctrl As Control
    Dim DescArr() As String
    If cmdLuu.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
        DescArr = LoadLanguage(LngFile, "#03:022:")
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
'    Set cnData = Get_Connection(strPath, "100881administrator")
'End If
    Set rsKhoanThu = OpenCriticalTable("Select * from Receipt", cnData)
    Set rsKhachHang = OpenCriticalTable("select * from Customer", cnData)
    Set rsPhieuthu = OpenCriticalTable("select * from Income", cnData)
    dtpDate.Value = Date

Call Locktextbox
Call Set_flgHeader
Call Set_FlgReceipt
'txtSophieu.Enabled = False
Exit Sub
Handle:
MsgBox Err.ne & Err.Description & Me.name & "Form_Load"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsKhachHang = Nothing
    Set rsKhoanThu = Nothing
    Set rsPhieuthu = Nothing
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

Private Sub txtDiengiaithu_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        cboKhachhang.SetFocus
    End If
End Sub

Private Sub txtNgay_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboKhoanthu.SetFocus
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
    MsgBox Err.Number & Err.Description & Me.name & "  txtNguoinoptien_DblClick"
End Sub

Private Sub txtNguoinoptien_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtBophan.SetFocus
    End If
End Sub

Private Sub txtSophieu_Change()
On Error GoTo Handle
    With rsPhieuthu
        .Find "ID='" & txtSophieu & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtNgay.Text = gfCONVERT_STRING_TO_DATE(Left(.Fields("DateTime"), 10))
            cboKhoanthu.Text = .Fields("Receipt_ID")
            cboKhachhang.Text = .Fields("Customer_ID")
            txtNguoinoptien.Text = .Fields("Reciever_Name")
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

Private Sub txtSophieu_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    dtpDate.SetFocus
 End If
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
        txtDiengiaithu.Locked = True
        txtTenKH.Locked = True
        txtDiachi.Locked = True
        txtSoDT.Locked = True
        txtMST.Locked = True
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   Locktextbox"
End Sub

Public Sub Set_FlgPhieuThu(rs As ADODB.Recordset)
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
                flgPhieuthu.Rows = rs.RecordCount + 1
                With flgPhieuthu
                    .TextMatrix(incount, 0) = rs!ID
                    .TextMatrix(incount, 1) = rs!DienGiai
                    .TextMatrix(incount, 2) = gfCONVERT_STRING_TO_DATE(rs!DateTime)
                    .TextMatrix(incount, 3) = rs!Reciever_Name
                    .TextMatrix(incount, 4) = Format(rs!Amount, formatNum)
                    .TextMatrix(incount, 5) = rs!Description
                End With
            rs.MoveNext
            Loop
        End With
        If rs.RecordCount = 0 Then
            With flgPhieuthu
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
    MsgBox Err.Number & Err.Description & Me.name & "Set_FlgPhieuThu"
End Sub
Public Sub Set_flgHeader()
    On Error GoTo Handle
        With flgPhieuthu
            .Cols = 6
            .Rows = 20
            .ColWidth(0) = 1500
            .ColWidth(1) = 1200
            .ColWidth(2) = 1200
            .ColWidth(3) = 1200
            .ColWidth(4) = 1200
            .ColWidth(5) = 2000
            .TextMatrix(0, 0) = "SË phi’u"
            .TextMatrix(0, 1) = "Kho∂n thu"
            .TextMatrix(0, 2) = "Ngµy thu"
            .TextMatrix(0, 3) = "Ng≠Íi nhÀn"
            .TextMatrix(0, 4) = "SË ti“n"
            .TextMatrix(0, 5) = "Di‘n gi∂i thu"
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
            
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Set_flgHeader"
End Sub

Public Sub Set_FlgReceipt()
On Error GoTo Handle
    Dim strFilter As String
    Dim rsReceipt_Inmonth As New ADODB.Recordset
    strFilter = "SELECT Income.ID, Receipt.DienGiai, Income.DateTime," & _
               " Income.Reciever_Name, Income.Cashier_ID, Income.Division," & _
               " Income.Amount, Income.Description" & _
               " FROM Receipt INNER JOIN Income ON Receipt.MaThu = Income.Receipt_ID" & _
               " where Left(Income.DateTime,6)='" & Format(Year(Date), "0000") & Format(Month(Date), "00") & "'" & _
               " ORDER BY Income.DateTime"
    Set rsReceipt_Inmonth = OpenCriticalTable(strFilter, cnData)
    Call Set_FlgPhieuThu(rsReceipt_Inmonth)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Set_FlgPhieuThu"
End Sub

Public Sub Add_Combo_Cust()
On Error GoTo Handle
If rsKhachHang.State <> 0 And rsKhachHang.RecordCount > 0 Then rsKhachHang.MoveFirst
'Gan list Khach hang
With cboKhachhang
    .Clear
    Do While Not rsKhachHang.EOF
        .AddItem rsKhachHang.Fields("CustNum")
        rsKhachHang.MoveNext
    Loop
End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Add_Combo_Cust"

End Sub

Public Sub Add_Combo_Khoanthu()
On Error GoTo Handle
If rsKhoanThu.State <> 0 And rsKhoanThu.RecordCount > 0 Then rsKhoanThu.MoveFirst
'Gan list Khoan thu
With cboKhoanthu
    .Clear
    Do While Not rsKhoanThu.EOF
        .AddItem rsKhoanThu.Fields("MaThu")
        rsKhoanThu.MoveNext
    Loop
End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Add_Combo_Khoanthu"
End Sub

Public Sub Clear_Text()
On Error GoTo Handle
    txtDiengiaithu.Text = ""
    txtTenKH.Text = ""
    txtDiachi.Text = ""
    txtSoDT.Text = ""
    txtMST.Text = ""
    txtNguoinoptien.Text = ""
    txtBophan.Text = ""
    txtSotien.Text = ""
    txtDiengiai.Text = ""
    lblChu.Caption = ""
    cboKhachhang.Text = ""
    cboKhoanthu.Text = ""
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Clear_Text"
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
            If Mid(.Thuchi, 3, 1) = 0 Then
                  cmdCreat.Enabled = False
            Else: cmdCreat.Enabled = True
            End If
            If Mid(.Thuchi, 4, 1) = 0 Then
                  cmdDelete.Enabled = False
            Else: cmdDelete.Enabled = True
            End If
            

        End With

    CloseRecordset res
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " CheckRight"
End Sub

