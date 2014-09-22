VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSetupPrice 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ThiÕt lËp gi¸ theo giê - theo khu vùc"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11505
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdMissmash 
      Height          =   855
      Left            =   5280
      TabIndex        =   27
      Top             =   6720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1508
      BTYPE           =   5
      TX              =   "KhuyÕn m·i"
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
      BCOL            =   14737632
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSetupPrice.frx":0000
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
      Cancel          =   -1  'True
      Height          =   825
      Left            =   8160
      TabIndex        =   2
      Tag             =   "L9"
      Top             =   6720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1455
      BTYPE           =   5
      TX              =   "&§ãng"
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
      MICON           =   "frmSetupPrice.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdSave 
      Height          =   825
      Left            =   2400
      TabIndex        =   1
      Tag             =   "L8"
      Top             =   6735
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1455
      BTYPE           =   5
      TX              =   "&Save change"
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
      MICON           =   "frmSetupPrice.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6465
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   11404
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "ThiÕt lËp gi¸ theo giê"
      TabPicture(0)   =   "frmSetupPrice.frx":0054
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "ThiÕt lËp gi¸ theo khu vùc"
      TabPicture(1)   =   "frmSetupPrice.frx":0070
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "fraSection"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   5775
         Left            =   -66600
         TabIndex        =   23
         Top             =   300
         Width           =   2775
         Begin VB.Label Label9 
            Caption         =   " *-*Chó ý: C¸c kho¶ng thêi gian thiÕt lËp kh«ng ®­îc chång chÐo nhau. "
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
            Height          =   975
            Left            =   120
            TabIndex        =   26
            Top             =   2040
            Width           =   2535
         End
         Begin VB.Label Label8 
            Caption         =   "VÝ dô: Giê cuèi cña kho¶ng thêi gian tr­íc kh«ng ®­îc lín h¬n giê ®Çu cña kho¶ng thêi gian sau"
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
            Height          =   1455
            Left            =   120
            TabIndex        =   25
            Top             =   3120
            Width           =   2535
         End
         Begin VB.Label Label7 
            Caption         =   $"frmSetupPrice.frx":008C
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
            Height          =   1695
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5775
         Left            =   -74940
         TabIndex        =   7
         Top             =   300
         Width           =   8295
         Begin VB.Frame fraStandar 
            Caption         =   "Standar Price ( Gi¸ chuÈn )"
            Height          =   1695
            Left            =   90
            TabIndex        =   18
            Tag             =   "L3"
            Top             =   300
            Width           =   8055
            Begin MSComCtl2.DTPicker dtpFrom 
               Height          =   495
               Index           =   0
               Left            =   1380
               TabIndex        =   19
               Top             =   750
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
               CustomFormat    =   "hh:MM:ss"
               Format          =   64684034
               UpDown          =   -1  'True
               CurrentDate     =   38462.25
            End
            Begin MSComCtl2.DTPicker dtpTo 
               Height          =   495
               Index           =   0
               Left            =   4980
               TabIndex        =   20
               Top             =   720
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
               CustomFormat    =   "hh:MM:ss"
               Format          =   64684034
               UpDown          =   -1  'True
               CurrentDate     =   38462.5826388889
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Tõ :"
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
               Height          =   285
               Left            =   270
               TabIndex        =   22
               Tag             =   "L6"
               Top             =   810
               Width           =   1005
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "§Õn:"
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
               Height          =   285
               Left            =   3870
               TabIndex        =   21
               Tag             =   "L7"
               Top             =   780
               Width           =   1005
            End
         End
         Begin VB.Frame fraHappy 
            Caption         =   "Happy hour Price ( Gi¸ giê vµng )"
            Height          =   1665
            Left            =   90
            TabIndex        =   13
            Tag             =   "L4"
            Top             =   2130
            Width           =   8055
            Begin MSComCtl2.DTPicker dtpFrom 
               Height          =   495
               Index           =   1
               Left            =   1350
               TabIndex        =   14
               Top             =   630
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
               CustomFormat    =   "hh:MM:ss"
               Format          =   64684034
               UpDown          =   -1  'True
               CurrentDate     =   38462.5833333333
            End
            Begin MSComCtl2.DTPicker dtpTo 
               Height          =   495
               Index           =   1
               Left            =   4920
               TabIndex        =   15
               Top             =   630
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
               CustomFormat    =   "hh:MM:ss"
               Format          =   64684034
               UpDown          =   -1  'True
               CurrentDate     =   38462.7493055556
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Tõ :"
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
               Height          =   285
               Left            =   270
               TabIndex        =   17
               Top             =   630
               Width           =   1005
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "§Õn:"
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
               Height          =   285
               Left            =   3870
               TabIndex        =   16
               Top             =   600
               Width           =   1005
            End
         End
         Begin VB.Frame fraEverning 
            Caption         =   "Everning  Price ( Gi¸ buæi tèi )"
            Height          =   1695
            Left            =   90
            TabIndex        =   8
            Tag             =   "L5"
            Top             =   3930
            Width           =   8055
            Begin MSComCtl2.DTPicker dtpFrom 
               Height          =   495
               Index           =   2
               Left            =   1350
               TabIndex        =   9
               Top             =   690
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
               CustomFormat    =   "hh:MM:ss"
               Format          =   64684034
               UpDown          =   -1  'True
               CurrentDate     =   38462.25
            End
            Begin MSComCtl2.DTPicker dtpTo 
               Height          =   495
               Index           =   2
               Left            =   4920
               TabIndex        =   10
               Top             =   690
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
               CustomFormat    =   "hh:MM:ss"
               Format          =   64684034
               UpDown          =   -1  'True
               CurrentDate     =   38462.25
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Tõ :"
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
               Height          =   285
               Left            =   270
               TabIndex        =   12
               Top             =   780
               Width           =   1005
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "§Õn:"
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
               Height          =   285
               Left            =   3870
               TabIndex        =   11
               Top             =   750
               Width           =   1005
            End
         End
      End
      Begin VB.Frame fraSection 
         Caption         =   "PhÇn tr¨m gi¸ ®­îc céng thªm vµo gi¸ chuÈn"
         ForeColor       =   &H00FF0000&
         Height          =   5745
         Left            =   60
         TabIndex        =   3
         Top             =   330
         Width           =   11175
         Begin VB.TextBox txtPercent 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   0
            Left            =   2700
            TabIndex        =   5
            Top             =   510
            Visible         =   0   'False
            Width           =   1365
         End
         Begin prjTouchScreen.MyButton cmdSection 
            Height          =   825
            Index           =   0
            Left            =   390
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   1455
            BTYPE           =   5
            TX              =   "Khu vùc"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   ".VnArialH"
               Size            =   12
               Charset         =   0
               Weight          =   700
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
            MICON           =   "frmSetupPrice.frx":011F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            Value           =   0   'False
         End
         Begin VB.Label lblPer 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   4050
            TabIndex        =   6
            Top             =   600
            Visible         =   0   'False
            Width           =   315
         End
      End
   End
End
Attribute VB_Name = "frmSetupPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLocation As New ADODB.Recordset
Dim rsSetupPrice As New ADODB.Recordset
Dim Desarr() As String
Dim ischange As Boolean


Private Sub cmdClose_Click()
If ischange = True Then
    Call cmdSave_Click
End If
    Unload Me
End Sub

Private Sub cmdMissmash_Click()
    frmMismas.Show vbModal
End Sub

Private Sub cmdSave_Click()
On Error GoTo Handle
Dim i As Integer
If ischange = True Then
    If MsgBox("B¹n cã muèn l­u th«ng tin thay ®æi kh«ng?", vbYesNo) = vbYes Then
        If rsSetupPrice.RecordCount > 0 Then
            With rsSetupPrice
                For i = 1 To 3
                    .Find "ID=" & i, , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                         .Fields("StartTime") = dtpFrom(i - 1).Value
                         .Fields("EndTime") = dtpTo(i - 1).Value
                         .Update
                    End If
                .MoveFirst
                Next i
            End With
        
        End If
        With rsLocation
            For i = 0 To txtPercent.count - 1
                .Find "Location_ID='" & txtPercent(i).Tag & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        .Fields("PriceRate") = txtPercent(i).Text
                        .Update
                    End If
            .MoveFirst
            Next
        End With
    End If
End If
ischange = False
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "SetPriceToform"
End Sub



Private Sub dtpFrom_Change(Index As Integer)
    On Error GoTo Handle
        ischange = True
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  dtpFrom_Change"
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    Dim ctrl As Control
    If cmdClose.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
        Desarr = LoadLanguage(LngFile, "#02:006:")
        Me.Caption = Desarr(1)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Desarr(Mid(ctrl.Tag, 2))
        Next ctrl
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    ischange = False
    Desarr = LoadLanguage(LngFile, "#02:006:")
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsLocation = Open_Table(cnData, "Table_Diagram_Sections")
    Set rsSetupPrice = Open_Table(cnData, "PeriodPrice")
    Call SetPriceToform
    Call LoadCommand(rsLocation, "Section_ID")
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub

Public Sub SetPriceToform()
On Error GoTo Handle
Dim i As Integer
If rsSetupPrice.RecordCount > 0 Then
    With rsSetupPrice
        For i = 1 To 3
            .Find "ID=" & i, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                dtpFrom(i - 1).Value = .Fields("StartTime")
                dtpTo(i - 1).Value = .Fields("EndTime")
            End If
        .MoveFirst
        Next i
    End With
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "SetPriceToform"
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
                        .Left = cmdSection(Index - 1).Left + 500 + cmdSection(Index - 1).Width + txtPercent(Index - 1).Width + 300
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
                        .Left = cmdSection(Index - 1).Left + 500 + cmdSection(Index - 1).Width + txtPercent(Index - 1).Width + 300
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
            Load txtPercent(Index)
            With txtPercent(Index)
                .top = cmdSection(Index).top + 80
                .Left = cmdSection(Index).Left + cmdSection(Index).Width + 100
                .Text = rs.Fields("PriceRate")
                .Tag = rs.Fields("Location_ID")
                .Visible = True
            End With
            Load lblPer(Index)
            With lblPer(Index)
                .top = txtPercent(Index).top
                .Left = txtPercent(Index).Left + txtPercent(Index).Width
                .Visible = True
            End With
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        rs.MoveNext
        Index = Index + 1
    Next j
Next i

End If
Exit Sub
'Handle:
'    MsgBox Err.Number & Err.Description & Me.Name & "  LoadCommandSub"
End Sub

Private Sub Option2_Click()

End Sub

Private Sub txtPercent_Change(Index As Integer)
    On Error GoTo Handle
        ischange = True
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtPercent_Change"
End Sub
