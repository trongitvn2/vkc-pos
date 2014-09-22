VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmMismas 
   Caption         =   "KhuyÕn m·i"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12135
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   12135
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   12135
      Begin VB.Frame Frame2 
         Caption         =   "Lo¹i khuyÕn m·i - thêi gian khuyÕn m·i "
         Height          =   7215
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   4575
         Begin VB.ComboBox cboMiss 
            BeginProperty Font 
               Name            =   ".VnArialH"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   7
            Text            =   "Chän lo¹i khuyÕn m·i"
            Top             =   960
            Width           =   4095
         End
         Begin VB.Frame fraDate 
            ForeColor       =   &H00FF0000&
            Height          =   1695
            Left            =   120
            TabIndex        =   2
            Top             =   1800
            Width           =   4215
            Begin MSComCtl2.DTPicker dtpFromDate 
               Height          =   495
               Left            =   1440
               TabIndex        =   3
               Top             =   360
               Width           =   2535
               _ExtentX        =   4471
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
               CustomFormat    =   "dd/MM/yyyy"
               Format          =   64487425
               UpDown          =   -1  'True
               CurrentDate     =   40610
            End
            Begin MSComCtl2.DTPicker dtpToDate 
               Height          =   495
               Left            =   1440
               TabIndex        =   4
               Top             =   960
               Width           =   2535
               _ExtentX        =   4471
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
               Format          =   64487425
               UpDown          =   -1  'True
               CurrentDate     =   40610
            End
            Begin VB.Label lblFromdate 
               Caption         =   "Tõ ngµy:"
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
               TabIndex        =   6
               Top             =   360
               Width           =   1425
            End
            Begin VB.Label lblDenngay 
               Caption         =   "§Õn ngµy:"
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
               TabIndex        =   5
               Top             =   960
               Width           =   1425
            End
         End
         Begin prjTouchScreen.MyButton cmdClose 
            Cancel          =   -1  'True
            Height          =   855
            Left            =   1320
            TabIndex        =   8
            Top             =   3840
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   1508
            BTYPE           =   5
            TX              =   "&§ãng"
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
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   16711680
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmMismask.frx":0000
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
      Begin prjTouchScreen.MyButton cmdAdd 
         Height          =   855
         Left            =   4800
         TabIndex        =   9
         Top             =   3000
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "Thªm"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   16711680
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMismask.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdRemove 
         Height          =   855
         Left            =   4800
         TabIndex        =   10
         Top             =   3960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "Xãa"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   16711680
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMismask.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid flgMiss 
         Height          =   7815
         Left            =   5760
         TabIndex        =   12
         Top             =   360
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   13785
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "chÝnh s¸ch khuyÕn m¹i"
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
         Height          =   495
         Left            =   720
         TabIndex        =   11
         Top             =   600
         Width           =   4815
      End
   End
End
Attribute VB_Name = "frmMismas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As New ADODB.Recordset
Dim rsMiss As New ADODB.Recordset

Private Sub cmdAdd_Click()
On Error GoTo Handle
If cboMiss.ListIndex = 0 Then
    MsgBox "Vui lßng chän lo¹i khuyÕn m·i !!!"
Else
    With rsMiss
        .Find "ID=" & cboMiss.ListIndex, , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields(0) = cboMiss.ListIndex
            .Fields(1) = cboMiss.Text
            .Fields(2) = dtpFromDate.Value
            .Fields(3) = dtpToDate.Value
            .Update
        Else
            MsgBox "H×nh thøc nµy ®· tån t¹i, nÕu tiÕp tôc vui lßng bá môc hiÖn t¹i"
        End If
    End With
   Call InitFlex(rsMiss)
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - cmdSave_Click"
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdRemove_Click()
On Error GoTo Handle
    With rsMiss
        .Find "ID='" & flgMiss.TextMatrix(flgMiss.Row, 0) & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Delete adAffectCurrent
            .Requery
        End If
    End With
   Call InitFlex(rsMiss)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - cmdRemove_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
If Check_Table_exist("MismatchTable") = False Then
    Call Create_MismatchTable
End If

Call AddMissType
    With flgMiss
        .Cols = 4
        .Rows = 3
        .ColWidth(0) = 500
        .ColWidth(1) = 2200
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .TextMatrix(0, 0) = "ID"
        .TextMatrix(0, 1) = "Tªn khuyÕn m·i"
        .TextMatrix(0, 2) = "Ngµy b¾t ®Çu"
        .TextMatrix(0, 3) = "Ngµy kÕt thóc"
        .ColAlignment(0) = 2
        .ColAlignment(1) = 4
        .ColAlignment(2) = 6
        .ColAlignment(3) = 6
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, 1) = ""
        .TextMatrix(1, 2) = ""
        .TextMatrix(1, 3) = ""
       
    End With
    dtpFromDate.Value = Date
    dtpToDate.Value = Date
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    Set rsMiss = Open_Table(cnData, "MismatchTable")
    Call InitFlex(rsMiss)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & "- Form_Load"
End Sub
Public Sub InitFlex(rs As ADODB.Recordset)
On Error GoTo Handle
    Dim incount As Integer
       If rs.State = 0 Then Exit Sub
       If rs.RecordCount = 0 Then Exit Sub
        rs.MoveFirst
        With rs
            .Sort = "ID ASC"
            Do While Not .EOF
                incount = incount + 1
                flgMiss.Rows = rs.RecordCount + 1
                With flgMiss
                    .TextMatrix(incount, 0) = rs.Fields(0)
                    .TextMatrix(incount, 1) = rs.Fields(1)
                    .TextMatrix(incount, 2) = Format(Day(rs.Fields(2)), "00") & "/" & Format(Month(rs.Fields(2)), "00") & "/" & Format(Year(rs.Fields(2)), "0000")
                    .TextMatrix(incount, 3) = Format(Day(rs.Fields(3)), "00") & "/" & Format(Month(rs.Fields(3)), "00") & "/" & Format(Year(rs.Fields(3)), "0000")
                End With
            rs.MoveNext
            Loop
        End With
        If rs.RecordCount = 0 Then
            With flgMiss
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
            
            End With
        End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - InitFlex"

End Sub


Public Sub AddMissType()
On Error GoTo Handle
    With cboMiss
        .Clear
        .AddItem "Chän h×nh thøc khuyÕn m·i"
        .AddItem "Gi¶m % Tæng H§"
        .AddItem "Gi¶m % Thøc ¨n"
        .AddItem "Gi¶m % thøc uèng"
    End With
    cboMiss.ListIndex = 0
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " - AddMissType"
End Sub
