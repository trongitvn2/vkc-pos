VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmWork_Shift 
   Caption         =   "Ca lµm viÖc"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11850
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
   ScaleHeight     =   8685
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   705
      Left            =   5850
      ScaleHeight     =   645
      ScaleWidth      =   5835
      TabIndex        =   6
      Top             =   60
      Width           =   5895
      Begin VB.Label lblNo 
         BackColor       =   &H80000008&
         Caption         =   "Shift_ID"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   45
         Width           =   1695
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000008&
         Caption         =   "Shift_ Name"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   330
         Width           =   2295
      End
   End
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
      Height          =   1125
      Left            =   5760
      TabIndex        =   0
      Top             =   7560
      Width           =   6165
      Begin prjTouchScreen.MyButton cmdThem 
         Height          =   705
         Left            =   60
         TabIndex        =   1
         Tag             =   "L14"
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Thªm"
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
         MICON           =   "frmWork_Shift.frx":0000
         PICN            =   "frmWork_Shift.frx":001C
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
         Height          =   705
         Left            =   1260
         TabIndex        =   2
         Tag             =   "L15"
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1244
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
         MICON           =   "frmWork_Shift.frx":046E
         PICN            =   "frmWork_Shift.frx":048A
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
         Height          =   705
         Left            =   2460
         TabIndex        =   3
         Tag             =   "L17"
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1244
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
         MICON           =   "frmWork_Shift.frx":09CE
         PICN            =   "frmWork_Shift.frx":09EA
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
         Height          =   705
         Left            =   3660
         TabIndex        =   4
         Tag             =   "L18"
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1244
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
         MICON           =   "frmWork_Shift.frx":1024
         PICN            =   "frmWork_Shift.frx":1040
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
         Height          =   705
         Left            =   4850
         TabIndex        =   5
         Tag             =   "L19"
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1244
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
         MICON           =   "frmWork_Shift.frx":167A
         PICN            =   "frmWork_Shift.frx":1696
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
   Begin TabDlg.SSTab tabGroup 
      Height          =   6765
      Left            =   5820
      TabIndex        =   9
      Top             =   780
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   11933
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "B¶ng thêi gian lµm viÖc"
      TabPicture(0)   =   "frmWork_Shift.frx":7930
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   6015
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   5775
         Begin VB.CheckBox chkout 
            Caption         =   "B¾t buéc ra ca"
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
            Left            =   3240
            TabIndex        =   28
            Tag             =   "L11"
            Top             =   4800
            Width           =   2415
         End
         Begin VB.CheckBox chkIn 
            Caption         =   "B¾t buéc vµo ca"
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
            Left            =   720
            TabIndex        =   27
            Tag             =   "L10"
            Top             =   4800
            Width           =   2415
         End
         Begin VB.TextBox CountMinute 
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
            Left            =   2520
            TabIndex        =   24
            Tag             =   "1"
            Top             =   3840
            Width           =   1575
         End
         Begin VB.TextBox txtLeaveSon 
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
            Left            =   2520
            TabIndex        =   21
            Tag             =   "1"
            Top             =   3240
            Width           =   1575
         End
         Begin VB.TextBox txtLateAllow 
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
            Left            =   2520
            TabIndex        =   18
            Tag             =   "1"
            Top             =   2640
            Width           =   1575
         End
         Begin VB.TextBox txtDescription 
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
            Left            =   2520
            TabIndex        =   13
            Tag             =   "1"
            Top             =   840
            Width           =   3135
         End
         Begin VB.TextBox txtShiftID 
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
            Left            =   2520
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   12
            Tag             =   "1"
            Top             =   240
            Width           =   1185
         End
         Begin MSComCtl2.DTPicker dtpStartTime 
            Height          =   495
            Left            =   2520
            TabIndex        =   29
            Top             =   1440
            Width           =   2805
            _ExtentX        =   4948
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
            CustomFormat    =   "hh:mm:ss"
            Format          =   16580610
            UpDown          =   -1  'True
            CurrentDate     =   0.25
         End
         Begin MSComCtl2.DTPicker dtpEndTime 
            Height          =   495
            Left            =   2520
            TabIndex        =   30
            Top             =   2040
            Width           =   2805
            _ExtentX        =   4948
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
            CustomFormat    =   "hh:mm:ss"
            Format          =   16580610
            UpDown          =   -1  'True
            CurrentDate     =   0.25
         End
         Begin VB.Label Label8 
            Caption         =   "Phót"
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
            Left            =   4200
            TabIndex        =   26
            Tag             =   "L12"
            Top             =   3960
            Width           =   735
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Tæng T/Gian lµm viÖc"
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
            TabIndex        =   25
            Tag             =   "L9"
            Top             =   3840
            Width           =   2355
         End
         Begin VB.Label Label6 
            Caption         =   "Phót"
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
            Left            =   4200
            TabIndex        =   23
            Tag             =   "L12"
            Top             =   3360
            Width           =   735
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Cho phÐp vÒ sím"
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
            TabIndex        =   22
            Tag             =   "L7"
            Top             =   3240
            Width           =   2355
         End
         Begin VB.Label Label4 
            Caption         =   "Phót"
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
            Left            =   4200
            TabIndex        =   20
            Tag             =   "L12"
            Top             =   2760
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Cho phÐp ®i trÓ"
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
            TabIndex        =   19
            Tag             =   "L6"
            Top             =   2640
            Width           =   2355
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Thêi gian kÕt thóc"
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
            TabIndex        =   17
            Tag             =   "L5"
            Top             =   2070
            Width           =   2355
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Thêi gian b¾t ®Çu"
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
            TabIndex        =   16
            Tag             =   "L4"
            Top             =   1470
            Width           =   2355
         End
         Begin VB.Label lblExpensesName 
            Alignment       =   1  'Right Justify
            Caption         =   "DiÔn gi¶i"
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
            TabIndex        =   15
            Tag             =   "L3"
            Top             =   870
            Width           =   2355
         End
         Begin VB.Label lblExpensesNo 
            Alignment       =   1  'Right Justify
            Caption         =   "M· sè ca"
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
            TabIndex        =   14
            Tag             =   "L2"
            Top             =   270
            Width           =   2355
         End
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flgWork_Shift 
      Height          =   8625
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   15214
      _Version        =   393216
      Cols            =   7
      BackColor       =   16777215
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      GridColorFixed  =   12632256
      TextStyleFixed  =   3
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
Attribute VB_Name = "frmWork_Shift"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim rsWork_Shift As New ADODB.Recordset
Dim strPath As String
Dim DescArr() As String
Dim wYear As Double
Dim wMonth As Integer

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

Private Sub cmdClose_Click()
    Set rsWork_Shift = Nothing
    Unload Me
End Sub

Private Sub cmdThem_Click()
On Error GoTo Handle
    If cmdThem.Caption = DescArr(14) Then
        Call UnlockText
        Call DeleteTextbox
    ElseIf cmdThem.Caption = DescArr(16) Then
        Call UnlockText
    End If
Exit Sub
Handle:
MsgBox Err.Number & " " & Err.Description & " " & _
            Me.name & " " & "cmdThem _click"

End Sub
Private Sub DeleteTextbox()
    On Error GoTo Handle
        cmdThem.Caption = DescArr(16)
        txtShiftID.Text = GetMax_ID("Work_Shift", "Shift_ID")
       txtDescription.Text = ""
       dtpStartTime.Value = Format(Now, "HH:mm:ss")
       dtpEndTime.Value = Format(Now, "HH:mm:ss")
       txtLateAllow.Text = ""
       txtLeaveSon.Text = ""
       CountMinute.Text = ""
        txtDescription.SetFocus
        
    Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & " " & _
                    Me.name & " " & "DeleteTextbox"
End Sub
Private Sub UpdateDatabase()
    On Error GoTo Handle
        With rsWork_Shift
            .Find "Shift_ID='" & txtShiftID.Text & _
                "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("Shift_ID") = txtShiftID.Text
                .Fields("Shift_Name") = txtDescription.Text
                .Fields("InTime") = dtpStartTime.Value
                .Fields("OutTime") = dtpEndTime.Value
                .Fields("LateTime") = txtLateAllow.Text
                .Fields("LeaveTime") = txtLeaveSon.Text
                .Fields("LongTime") = CountMinute.Text
                If chkIn.Value = 1 Then .Fields("MustIn") = True
                If chkout.Value = 1 Then .Fields("MustOut") = True
                .Update
                .Requery
            Else
                MsgBox DescArr(8), vbOKOnly
                Call DeleteTextbox
                
            End If
        End With
        Call setflgWork_Shift
        cmdThem.Caption = DescArr(14)
        
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
        With rsWork_Shift
            .Find "Shift_ID='" & flgWork_Shift.TextMatrix(flgWork_Shift.Row, 0) & _
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

Private Sub dtpEndTime_Change()
    CountMinute.Text = Hour(dtpEndTime.Value) * 60 + Minute(dtpEndTime.Value) - Hour(dtpStartTime.Value) * 60 - Minute(dtpStartTime.Value)
    CountMinute.Locked = True
End Sub

Private Sub dtpEndTime_LostFocus()
    CountMinute.Text = Hour(dtpEndTime.Value) * 60 + Minute(dtpEndTime.Value) - Hour(dtpStartTime.Value) * 60 - Minute(dtpStartTime.Value)
    CountMinute.Locked = True
End Sub


Private Sub dtpStartTime_Change()
    CountMinute.Text = Hour(dtpEndTime.Value) * 60 + Minute(dtpEndTime.Value) - Hour(dtpStartTime.Value) * 60 - Minute(dtpStartTime.Value)
    CountMinute.Locked = True

End Sub

Private Sub flgWork_Shift_EnterCell()
    On Error GoTo Handle
    With rsWork_Shift
        .Find "Shift_ID='" & flgWork_Shift.TextMatrix(flgWork_Shift.Row, 0) & _
                "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtShiftID.Text = !Shift_ID
            txtDescription.Text = !Shift_Name
            dtpStartTime.Value = !Intime
            dtpEndTime.Value = !OutTime
            txtLateAllow.Text = !LateTime
            txtLeaveSon.Text = !LeaveTime
            CountMinute.Text = !LongTime
            If !MustIn = True Then
                chkIn.Value = 1
            Else
                chkIn.Value = 0
            End If
            If !MustOut = True Then
                chkout.Value = 1
            Else
                chkout.Value = 0
            End If
            lblNo.Caption = !Shift_ID
            lblName.Caption = !Shift_Name
        End If
    End With
    Exit Sub
    
Handle:
    MsgBox Err.Number & " " & Err.Description & vbCrLf _
    & Me.name & " flgWork_Shift_EnterCell"
End Sub

Private Sub Form_Activate()
    Dim ctrl As Control
    If cmdCapnhat.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    DescArr = LoadLanguage(LngFile, "#03:004:")
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
    DescArr = LoadLanguage(LngFile, "#03:004:")
    str = "Select * from Work_Shift"
    Set rsWork_Shift = OpenCriticalTable(str, cnData)
    Call setflgWork_Shift
    Call LockText
    Exit Sub
Handle:
    MsgBox Err.Number & " " & _
    Err.Description & Me.name & "Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Handle
    Set rsWork_Shift = Nothing
Exit Sub
Handle:
    MsgBox Err.Number & " " & Err.Description & "" & _
                        Me.name & " " & "form_Unload"
End Sub
Private Sub setflgWork_Shift()
On Error GoTo errHdl
    Dim intCount    As Integer
    With flgWork_Shift
    .Cols = 7
    .Rows = 2
        .Font = ".vnArial"
        .ColWidth(0) = 1200
        .ColWidth(1) = 3500
        .ColWidth(2) = 3000
        .ColWidth(3) = 3000
        .ColWidth(4) = 2000
        .ColWidth(5) = 2000
        .ColWidth(6) = 1500
        .TextMatrix(0, 0) = DescArr(2)
        .TextMatrix(0, 1) = DescArr(3)
        .TextMatrix(0, 2) = DescArr(4)
        .TextMatrix(0, 3) = DescArr(5)
        .TextMatrix(0, 4) = DescArr(6)
        .TextMatrix(0, 5) = DescArr(7)
        .TextMatrix(0, 6) = DescArr(9)
        
    End With
    
    If rsWork_Shift Is Nothing Then Exit Sub
    If rsWork_Shift.State = 0 Then Exit Sub
    
    If rsWork_Shift.EOF And rsWork_Shift.BOF Then
        With flgWork_Shift
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
            .TextMatrix(1, 4) = ""
            .TextMatrix(1, 5) = ""
            .TextMatrix(1, 6) = ""
            
        End With
        Exit Sub
    End If
   flgWork_Shift.Rows = rsWork_Shift.RecordCount + 1
    intCount = 0
    Do While Not rsWork_Shift.EOF
        intCount = intCount + 1
        flgWork_Shift.TextMatrix(intCount, 0) = rsWork_Shift!Shift_ID
        flgWork_Shift.TextMatrix(intCount, 1) = rsWork_Shift!Shift_Name
        flgWork_Shift.TextMatrix(intCount, 2) = rsWork_Shift!Intime
        flgWork_Shift.TextMatrix(intCount, 3) = rsWork_Shift!OutTime
        flgWork_Shift.TextMatrix(intCount, 4) = rsWork_Shift!LateTime
        flgWork_Shift.TextMatrix(intCount, 5) = rsWork_Shift!LeaveTime
        flgWork_Shift.TextMatrix(intCount, 6) = rsWork_Shift!LongTime
        rsWork_Shift.MoveNext
        
    Loop
'    SetColorFlexGrid flgWork_Shift, 1, 1, flgWork_Shift.Cols

    Call flgWork_Shift_EnterCell
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - setflgWork_Shift "
End Sub
Private Sub LoadControl()
    On Error GoTo Handle
    
    With rsWork_Shift
        .Find "Shift_ID='" & !Shift_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtShiftID.Text = !Shift_ID
           txtDescription.Text = !Shift_Name
           dtpStartTime.Value = !Intime
           dtpEndTime.Value = !OutTime
           txtLateAllow.Text = !LateTime
           txtLeaveSon.Text = !LeaveTime
           CountMinute.Text = !LongTime
            .Requery
        End If
    End With
    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " LoadControl"
End Sub

Public Sub UnlockText()
    On Error GoTo Handle
        txtShiftID.Locked = False
        txtDescription.Locked = False
        txtLateAllow.Locked = False
        txtLeaveSon.Locked = False
        CountMinute.Locked = False
        cmdCapnhat.Enabled = True
        txtShiftID.SetFocus
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub
Public Sub LockText()
    On Error GoTo Handle
        txtShiftID.Locked = True
        txtDescription.Locked = True
        txtLateAllow.Locked = True
        txtLeaveSon.Locked = True
        CountMinute.Locked = True
        cmdCapnhat.Enabled = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " UnlockText"
End Sub

Private Sub txtDescription_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtDescription.Text = .Let_Text_Input
        End With

    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtShiftID_DblClick "

End Sub

Private Sub txtDescription_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        dtpStartTime.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtDescription_KeyPress"

End Sub

Private Sub txtShiftID_DblClick()
    On Error GoTo Handle:
        With frmKeyboard
            .txtInput.PasswordChar = ""
            .FormCallkeyboard = "Other"
            .Show vbModal
            txtShiftID.Text = .Let_Text_Input
       End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtShiftID_DblClick "

End Sub

Private Sub txtShiftID_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then
        txtDescription.SetFocus
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   txtShiftID_KeyPress"
End Sub
