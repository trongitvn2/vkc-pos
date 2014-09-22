VERSION 5.00
Begin VB.Form frmSelectPrint 
   Caption         =   "Chän m¸y in"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8715
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   8715
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton cmPrint 
      Height          =   975
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1720
      BTYPE           =   1
      TX              =   "m¸y in"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "frmSelectPrint.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton MyButton2 
      Height          =   975
      Left            =   6720
      TabIndex        =   2
      Top             =   6720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1720
      BTYPE           =   1
      TX              =   "&Tho¸t"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   255
      BCOLO           =   16711680
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSelectPrint.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Chän m¸y in"
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
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmSelectPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PrinterName As String

Private Sub cmPrint_Click(Index As Integer)
    PrinterName = cmPrint(Index).Caption
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Dim prn As Printer
    Dim Index As Integer
    Index = 1
    'Temporary define
    For Each prn In Printers
        Load cmPrint(Index)
            With cmPrint(Index)
            If Index = 1 Then
                .top = cmPrint(Index - 1).top
            Else
                .top = cmPrint(Index - 1).top + 100 + cmPrint(Index).Height
            End If
                .Left = 120
                .Caption = prn.DeviceName
                .Visible = True
                .Height = 975
                .Width = 8535
            End With
        Index = Index + 1
    Next
'    For i = 0 To cboPrinter.ListCount - 1
'        If cboPrinter.List(i) = Printer.DeviceName Then
'            cboPrinter.ListIndex = CDbl(GetSettingStr("Report", "Report", True, myIniFile))
'            Exit For
'        End If
'    Next
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " - Form_Load"
End Sub

Private Sub MyButton2_Click()
    PrinterName = ""
    Unload Me
End Sub



Public Property Get LetPrinter() As Variant
    LetPrinter = PrinterName
End Property

