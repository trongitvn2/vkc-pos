VERSION 5.00
Begin VB.Form frmSelect_Station 
   Caption         =   "Lùa chän quÇy"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
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
   ScaleHeight     =   4935
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin prjTouchScreen.MyButton cmdSection 
         Height          =   825
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1455
         BTYPE           =   14
         TX              =   "QuÇy"
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
         BCOL            =   12632319
         BCOLO           =   16711680
         FCOL            =   16711680
         FCOLO           =   8421631
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSelect_Station.frx":0000
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
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   825
      Left            =   960
      TabIndex        =   2
      Top             =   3960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1455
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
      BCOL            =   14737632
      BCOLO           =   12648447
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSelect_Station.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdNext 
      Height          =   825
      Left            =   3240
      TabIndex        =   3
      Top             =   3960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1455
      BTYPE           =   14
      TX              =   "TiÕp tôc >>"
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
      BCOL            =   14737632
      BCOLO           =   12648447
      FCOL            =   16711680
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSelect_Station.frx":0038
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
Attribute VB_Name = "frmSelect_Station"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsStation As New ADODB.Recordset
Dim strStationID As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If cmdNext.Caption = "TiÕp tôc >>" Then
        If strStationID = "" Then
            MsgBox "Vui lßng lùa chän quÇy"
        Else
            cmdNext.Caption = "&§ång ý"
        End If
    Else: cmdNext.Caption = "&§ång ý"
        SaveSettingStr "SYSTEM", "Station", strStationID, myIniFile
        Unload Me
    End If
End Sub

Private Sub cmdSection_Click(Index As Integer)
    strStationID = cmdSection(Index).Tag
End Sub

Private Sub Form_Load()
On Error GoTo Handle

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    
    Set rsStation = Open_Table(cnData, "Stations_Location")
    Call LoadCommand(rsStation, "Station_Name")
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Public Sub LoadCommand(rs As ADODB.Recordset, strTenfield1 As String)
On Error Resume Next 'GoTo Handle
Dim Index As Integer
Dim sodong As Integer
Dim i, j As Integer
Index = 1
rs.MoveFirst
If rs.RecordCount Mod 2 > 0 Then
    sodong = rs.RecordCount / 2 + 1
Else
    sodong = rs.RecordCount / 2
End If
If rs.RecordCount > 0 Then
For i = 1 To sodong
    For j = 1 To 2
            Load cmdSection(Index)
            With cmdSection(Index)
            If i = 1 Then
                If Index Mod 3 = 0 Then
                    .Left = fraSection.Left + 500
                    .top = cmdSection(Index - 1).top + cmdSection(Index - 1).Height + 200
                Else
                    .top = cmdSection(Index - 1).top
                    If j = 1 Then
                         .Left = fraSection.Left + 500
                    Else
                        .Left = cmdSection(Index - 1).Left + 500 + cmdSection(Index - 1).Width
                    End If
                End If
            Else
                If (Index - 1) Mod 2 = 0 Then
                    .Left = fraSection.Left + 500
                    .top = cmdSection(Index - 1).top + cmdSection(Index - 1).Height + 200
                Else
                    .top = cmdSection(Index - 1).top
                    If j = 1 Then
                       .Left = fraSection.Left + 300
                    Else
                        .Left = cmdSection(Index - 1).Left + 500 + cmdSection(Index - 1).Width
                    End If
                End If
            End If
                If Not rs.EOF Then
                    .Caption = rs.Fields("" & strTenfield1 & "") '& Chr(13) & rs.Fields("" & strTenfield2 & "")
                    .Tag = rs.Fields("Station_Number")
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
