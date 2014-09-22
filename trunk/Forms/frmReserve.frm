VERSION 5.00
Begin VB.Form frmReserve 
   Caption         =   "Danh s¸ch ®Æt cäc"
   ClientHeight    =   10950
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15240
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
   ScaleHeight     =   10950
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Fra 
      Height          =   10095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      Begin prjTouchScreen.MyButton cmdPhieu 
         Height          =   1455
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Visible         =   0   'False
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   2566
         BTYPE           =   3
         TX              =   "MyButton1"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   16711680
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmReserve.frx":0000
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
   Begin prjTouchScreen.MyButton cmdClose 
      Cancel          =   -1  'True
      Height          =   855
      Left            =   13200
      TabIndex        =   2
      Top             =   10200
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "§ãng"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmReserve.frx":001C
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
Attribute VB_Name = "frmReserve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsReserve As New ADODB.Recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Dim strSQL As String
    strSQL = "SELECT Table_Reservered.Reservered_Code," & _
    " Table_Reservered.DateTime, Table_Reservered.CustName," & _
    " Table_Reservered.Phone, Table_Reservered.Date_Reservered," & _
    " Table_Reservered.Time_Reservered, Table_Reservered.Table_ID," & _
    " Table_Reservered.Amount, Table_Reservered.IsUsed" & _
    " From Table_Reservered" & _
    " where Table_Reservered.Date_Reservered='" & Format(Date, "dd/MM/yyyy") & "' " & _
    " AND Table_Reservered.IsUsed=0"
    If cnData.State <> 0 Then
        Set rsReserve = OpenCriticalTable(strSQL, cnData)
    End If
    If rsReserve.State <> 0 And rsReserve.RecordCount > 0 Then
        rsReserve.MoveFirst
        Call LoadCommand(rsReserve)
    End If
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & ""
End Sub


Public Sub LoadCommand(rs As ADODB.Recordset)
On Error Resume Next 'GoTo Handle
Dim Index As Integer
Dim sodong As Integer
Dim i, J As Integer
Index = 1
rs.MoveFirst
If rs.RecordCount Mod 2 > 0 Then
    sodong = rs.RecordCount / 2 + 1
Else
    sodong = rs.RecordCount / 2
End If
If rs.RecordCount > 0 Then
For i = 1 To sodong
    For J = 1 To 5
            Load cmdPhieu(Index)
            With cmdPhieu(Index)
            If i = 1 Then
                If Index Mod 6 = 0 Then
                    .Left = Fra.Left
                    .top = cmdPhieu(Index - 1).top + cmdPhieu(Index - 1).Height + 200
                Else
                    .top = cmdPhieu(Index - 1).top
                    If J = 1 Then
                        .Left = 200
                    Else
                        .Left = cmdPhieu(Index - 1).Left + cmdPhieu(Index - 1).Width + 100
                    End If
                End If
            Else
                If (Index - 1) Mod 5 = 0 Then
                    .Left = 200
                    .top = cmdPhieu(Index - 1).top + cmdPhieu(Index - 1).Height + 200
                Else
                    .top = cmdPhieu(Index - 1).top
                    If J = 1 Then
                       .Left = 200
                    Else
                        .Left = cmdPhieu(Index - 1).Left + cmdPhieu(Index - 1).Width + 100
                    End If
                End If
            End If
                If Not rs.EOF Then
                    .Caption = rs.Fields("Table_ID") & Chr(13) & rs.Fields("CustName") & Chr(13) & rs.Fields("Phone")
                    .Tag = rs.Fields("Reservered_Code")
                Else
                    Exit Sub
                End If
                .Visible = True
                .Height = 1455
                .Width = 3255
            End With
        rs.MoveNext
        Index = Index + 1
    Next J
Next i

End If
Exit Sub
'Handle:
'    MsgBox Err.Number & Err.Description & Me.Name & "  LoadCommandSub"
End Sub
