VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDate_Sync 
   Caption         =   "Ngµy"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6255
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
   ScaleHeight     =   3255
   ScaleWidth      =   6255
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton frmOk 
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "&OK"
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
      BCOL            =   14737632
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDate_Sync.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpfdate 
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArialH"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      Format          =   20709377
      UpDown          =   -1  'True
      CurrentDate     =   40157
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   855
      Left            =   3240
      TabIndex        =   2
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "Cancel"
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
      BCOL            =   14737632
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDate_Sync.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtptdate 
      Height          =   615
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArialH"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   16711680
      CalendarTitleForeColor=   16711680
      Format          =   20709377
      UpDown          =   -1  'True
      CurrentDate     =   40157
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "§Õn ngµy:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tõ ngµy:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "§ång bé d÷ liÖu b¸n hµng"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "frmDate_Sync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isOK As Boolean
Dim fdate, tdate As String

Private Sub cmdCancel_Click()
    Unload Me
    fdate = ""
    tdate = ""
End Sub

Private Sub Form_Load()
On Error GoTo handle
    isOK = False
    dtpfdate.Value = gfCONVERT_STRING_TO_DATE(fdate)
    dtptdate.Value = gfCONVERT_STRING_TO_DATE(tdate)
    
Exit Sub
handle:
MsgBox Err.Number & Err.Description & Me.Name & "  Form_Load"
End Sub

Private Sub frmOk_Click()
On Error GoTo handle
    isOK = True
    fdate = gfCONVERT_DATE_TO_STRING(dtpfdate.Value)
    tdate = gfCONVERT_DATE_TO_STRING(dtptdate.Value)
    Unload Me
Exit Sub
handle:
MsgBox Err.Number & Err.Description & Me.Name & "  frmOk_Click"

End Sub

Public Property Let Let_FDate(ByVal vNewValue As Variant)
    fdate = vNewValue
End Property

Public Property Get Let_FDate() As Variant
    Let_FDate = fdate
End Property

Public Property Get Let_TDate() As Variant
    Let_TDate = tdate
End Property

Public Property Let Let_TDate(ByVal vNewValue As Variant)
    tdate = vNewValue
End Property
