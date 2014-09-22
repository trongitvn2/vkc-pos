VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTimeReport 
   Caption         =   "Giê b¸o c¸o"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7635
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
   ScaleHeight     =   2115
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   735
      Left            =   3360
      TabIndex        =   4
      Top             =   1320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1296
      BTYPE           =   14
      TX              =   "&OK"
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
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTimeReport.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpFrom 
      Height          =   495
      Left            =   1110
      TabIndex        =   0
      Top             =   510
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
      CustomFormat    =   "HH:mm:ss"
      Format          =   48365570
      UpDown          =   -1  'True
      CurrentDate     =   38462.25
   End
   Begin MSComCtl2.DTPicker dtpTo 
      Height          =   495
      Left            =   4710
      TabIndex        =   1
      Top             =   480
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
      CustomFormat    =   "HH:mm:ss"
      Format          =   48365570
      UpDown          =   -1  'True
      CurrentDate     =   38462.5826388889
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Tõ :"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   0
      TabIndex        =   3
      Tag             =   "L6"
      Top             =   570
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "§Õn:"
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   3600
      TabIndex        =   2
      Tag             =   "L7"
      Top             =   540
      Width           =   1005
   End
End
Attribute VB_Name = "frmTimeReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fTime, tTime As String

Private Sub Form_Load()
On Error GoTo Handle
    dtpFrom.Value = Format("00:00:00", "HH:mm:ss")
    dtpTo.Value = Format("23:59:59", "HH:mm:ss")
    fTime = dtpFrom.Value
    tTime = dtpTo.Value
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "Form_Load"
End Sub

Public Property Get GetFTime() As Variant
    GetFTime = fTime
End Property

Public Property Let GetFTime(ByVal vNewValue As Variant)
    fTime = vNewValue
End Property

Public Property Get GetTTime() As Variant
    GetTTime = tTime
End Property

Public Property Let GetTTime(ByVal vNewValue As Variant)
    tTime = vNewValue
End Property

Private Sub MyButton1_Click()
    fTime = dtpFrom.Value
    tTime = dtpTo.Value
    Unload Me
End Sub
