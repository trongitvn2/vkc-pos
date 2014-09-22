VERSION 5.00
Begin VB.Form frmBillDone 
   Caption         =   "Select Bill Done"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
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
   ScaleHeight     =   7965
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdSelect 
      Height          =   855
      Left            =   3000
      TabIndex        =   3
      Top             =   6960
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   1508
      BTYPE           =   6
      TX              =   "&Lùa chän"
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
      BCOL            =   12648447
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBillDone.frx":0000
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
      Height          =   855
      Left            =   1440
      TabIndex        =   2
      Top             =   6960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      BTYPE           =   6
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
      BCOL            =   12648447
      BCOLO           =   12648384
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBillDone.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdBillSub 
      Height          =   1095
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1931
      BTYPE           =   5
      TX              =   "MyButton1"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777088
      BCOLO           =   16761087
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBillDone.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Line Line2 
      DrawMode        =   12  'Nop
      X1              =   120
      X2              =   7560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      DrawMode        =   12  'Nop
      X1              =   120
      X2              =   7560
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Chän hãa ®¬n cÇn thùc hiÖn"
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frmBillDone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bill_Master As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " Form_Load"
End Sub

Public Property Let GetBill_Master(ByVal vNewValue As Variant)
    Bill_Master = vNewValue
End Property


Public Sub LoadCommand(rs As ADODB.Recordset, strTenfield1 As String)
On Error Resume Next 'GoTo Handle
Dim Index As Integer
Dim sodong As Integer
Dim i, j As Integer
Index = 1
rs.MoveFirst
If rs.RecordCount Mod 3 > 0 Then
    sodong = rs.RecordCount / 3 + 1
Else
    sodong = rs.RecordCount / 3
End If
If rs.RecordCount > 0 Then
For i = 1 To sodong
    For j = 1 To 3
            Load cmdBillSub(Index)
            With cmdBillSub(Index)
            If i = 1 Then
                If Index Mod 4 = 0 Then
                    .Left = 500
                    .top = cmdBillSub(Index - 1).top + cmdBillSub(Index - 1).Height + 200
                Else
                    .top = cmdBillSub(Index - 1).top
                    If j = 1 Then
                         .Left = 500
                    Else
                        .Left = cmdBillSub(Index - 1).Left + 500 + cmdBillSub(Index - 1).Width
                    End If
                End If
            Else
                If (Index - 1) Mod 3 = 0 Then
                    .Left = 500
                    .top = cmdBillSub(Index - 1).top + cmdBillSub(Index - 1).Height + 200
                Else
                    .top = cmdBillSub(Index - 1).top
                    If j = 1 Then
                       .Left = 300
                    Else
                        .Left = cmdBillSub(Index - 1).Left + 500 + cmdBillSub(Index - 1).Width
                    End If
                End If
            End If
                If Not rs.EOF Then
                    .Caption = rs.Fields("" & strTenfield1 & "") '& Chr(13) & rs.Fields("" & strTenfield2 & "")
                Else
                    Exit Sub
                End If
                .Visible = True
                .Height = 900
                .Width = 1600
        
            End With
                rs.MoveNext
        Index = Index + 1
    Next j
Next i

End If
Exit Sub
'Handle:
'    MsgBox Err.Number & Err.Description & Me.Name & "  LoadCommandSub"
End Sub
