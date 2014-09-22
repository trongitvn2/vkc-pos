VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7785
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
   Icon            =   "frmDate.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   7785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton frmOk 
      Height          =   855
      Left            =   1440
      TabIndex        =   2
      Top             =   1680
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
      MICON           =   "frmDate.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComCtl2.DTPicker dtpdate 
      Height          =   735
      Left            =   1680
      TabIndex        =   0
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1296
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
      Format          =   61341697
      CurrentDate     =   40157
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   855
      Left            =   3720
      TabIndex        =   3
      Top             =   1680
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
      MICON           =   "frmDate.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "D÷ liÖu b¸n hµng sÏ l­u vµo ngµy"
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
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frmDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    End
End Sub
Private Sub dtpdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call frmOk_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call frmOk_Click
End Sub

Private Sub Form_Load()
    dtpDate.Value = Date
End Sub

Private Sub frmOk_Click()
On Error GoTo Handle
DateDefault = gfCONVERT_DATE_TO_STRING(dtpDate.Value)
If DateDefault > Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") Or DateDefault < Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date) - 1, "00") Then Exit Sub
'SaveSettingStr "SYSTEM", "isOpen", "1", myIniFile
'SaveSettingStr "SYSTEM", "DateOpen", DateDefault, myIniFile
Unload Me
'With frmLogin
'        .Me_State = 0
'        .Show vbModal
'    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdOk_Click"
End Sub

Private Sub MyButton1_Click()

End Sub
