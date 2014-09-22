VERSION 5.00
Begin VB.Form frmQtyTranfer 
   BackColor       =   &H00000000&
   Caption         =   " "
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
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
   ScaleHeight     =   8310
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   7035
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   4215
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1155
         Index           =   0
         Left            =   90
         TabIndex        =   3
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   2037
         BTYPE           =   5
         TX              =   "1"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1155
         Index           =   1
         Left            =   1440
         TabIndex        =   4
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   2037
         BTYPE           =   5
         TX              =   "2"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1155
         Index           =   2
         Left            =   2790
         TabIndex        =   5
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   2037
         BTYPE           =   5
         TX              =   "3"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1155
         Index           =   3
         Left            =   90
         TabIndex        =   6
         Top             =   1425
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   2037
         BTYPE           =   5
         TX              =   "4"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1155
         Index           =   4
         Left            =   1440
         TabIndex        =   7
         Top             =   1425
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   2037
         BTYPE           =   5
         TX              =   "5"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1155
         Index           =   5
         Left            =   2790
         TabIndex        =   8
         Top             =   1425
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   2037
         BTYPE           =   5
         TX              =   "6"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":008C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1155
         Index           =   6
         Left            =   90
         TabIndex        =   9
         Top             =   2640
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   2037
         BTYPE           =   5
         TX              =   "7"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":00A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1155
         Index           =   7
         Left            =   1440
         TabIndex        =   10
         Top             =   2640
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   2037
         BTYPE           =   5
         TX              =   "8"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":00C4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1155
         Index           =   8
         Left            =   2790
         TabIndex        =   11
         Top             =   2640
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   2037
         BTYPE           =   5
         TX              =   "9"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":00E0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1155
         Index           =   10
         Left            =   1440
         TabIndex        =   12
         Top             =   3870
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   2037
         BTYPE           =   5
         TX              =   "00"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":00FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1155
         Index           =   11
         Left            =   2790
         TabIndex        =   13
         Top             =   3870
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   2037
         BTYPE           =   5
         TX              =   "."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":0118
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   855
         Index           =   13
         Left            =   2095
         TabIndex        =   14
         Tag             =   "L4"
         Top             =   5085
         Width           =   1960
         _ExtentX        =   3466
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "&Tho¸t"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":0134
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   945
         Index           =   14
         Left            =   90
         TabIndex        =   15
         Tag             =   "L5"
         Top             =   5985
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   1667
         BTYPE           =   5
         TX              =   "&§ång ý"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":0150
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   1155
         Index           =   9
         Left            =   90
         TabIndex        =   16
         Top             =   3870
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   2037
         BTYPE           =   5
         TX              =   "0"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":016C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdAlpha 
         Height          =   855
         Index           =   12
         Left            =   90
         TabIndex        =   17
         Tag             =   "L3"
         Top             =   5085
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "&Xãa"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   16711680
         FCOLO           =   16777215
         MCOL            =   -2147483638
         MPTR            =   1
         MICON           =   "frmQtyTranfer.frx":0188
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   4
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   40
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "NhËp sè l­îng chuyÓn"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Tag             =   "L2"
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmQtyTranfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim formCallme As Integer
Dim OK, isOK As Boolean
Dim Result As Double
Private Sub cmdAlpha_Click(Index As Integer)
    Select Case Index
        Case 0 To 11:
            Text1.Text = Text1.Text & cmdAlpha(Index).Caption
        Case 12:
            Text1.Text = ""
        Case 13:
            isOK = False
            Unload Me
        Case 14:
        If Text1.Text = "" Then
            MsgBox "NhËp sè l­îng chuyÓn", vbInformation
            Exit Sub
        End If
            isOK = True
           qtyTran = CInt(Text1.Text)
           Result = CInt(Text1.Text)
           OK = True
        
        Unload Me
    End Select
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    Dim ctrl As Control
    Dim DescArr() As String
    If cmdAlpha(1).Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
        DescArr = LoadLanguage(LngFile, "#03:025:")
        Me.Caption = DescArr(1)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
        Next ctrl
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Form_Load"

End Sub

Private Sub Form_Load()
    On Error GoTo Handle
        With Text1
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Form_Load"
End Sub


Public Property Let FormCall(ByVal vNewValue As Variant)
    formCallme = vNewValue
End Property


Public Property Get GetOK() As Variant
    GetOK = OK
End Property

Public Property Get Let_Result() As Variant
    Let_Result = Result
End Property

Private Sub Form_Unload(Cancel As Integer)
    If Not isOK Then
        qtyTran = 0
        Result = 0
    End If
End Sub
