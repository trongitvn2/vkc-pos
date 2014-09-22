VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAdd_Location 
   Caption         =   "Khu vùc"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10815
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
   ScaleHeight     =   4875
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   0
      TabIndex        =   12
      Top             =   -120
      Width           =   8775
      Begin MSDataGridLib.DataGrid dtgLocation 
         Height          =   4275
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   7541
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   30
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial NarrowH"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   2535
         Left            =   4320
         TabIndex        =   6
         Top             =   2160
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   4471
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Møc gi¸ menu"
         TabPicture(0)   =   "frmAdd_Location.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Lo¹i h×nh"
         TabPicture(1)   =   "frmAdd_Location.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "optType(0)"
         Tab(1).Control(1)=   "optType(1)"
         Tab(1).Control(2)=   "optType(2)"
         Tab(1).ControlCount=   3
         Begin VB.OptionButton optType 
            Caption         =   "Kh¸ch s¹n - Nhµ nghØ..."
            Height          =   375
            Index           =   2
            Left            =   -74880
            TabIndex        =   9
            Top             =   1920
            Width           =   2895
         End
         Begin VB.OptionButton optType 
            Caption         =   "Karaoke- Billiard..."
            Height          =   375
            Index           =   1
            Left            =   -74880
            TabIndex        =   8
            Top             =   1200
            Width           =   2295
         End
         Begin VB.OptionButton optType 
            Caption         =   "Nhµ hµng/ Cafe.."
            Height          =   375
            Index           =   0
            Left            =   -74880
            TabIndex        =   7
            Top             =   480
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.Frame Frame2 
            Height          =   2055
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   4095
            Begin VB.OptionButton OptPrice 
               Caption         =   "Gi¸ 1"
               Height          =   375
               Index           =   0
               Left            =   360
               TabIndex        =   3
               Top             =   480
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton OptPrice 
               Caption         =   "Gi¸ 2"
               Height          =   375
               Index           =   1
               Left            =   360
               TabIndex        =   4
               Top             =   960
               Width           =   1815
            End
            Begin VB.OptionButton OptPrice 
               Caption         =   "Gi¸ 3"
               Height          =   375
               Index           =   2
               Left            =   360
               TabIndex        =   5
               Top             =   1440
               Width           =   1815
            End
         End
      End
      Begin VB.TextBox txtVat 
         Height          =   495
         Left            =   6960
         TabIndex        =   2
         Text            =   "0"
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtServ 
         Height          =   495
         Left            =   4560
         TabIndex        =   1
         Text            =   "0"
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtName 
         Height          =   495
         Left            =   4560
         TabIndex        =   0
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label3 
         Caption         =   "VAT:"
         Height          =   375
         Left            =   6360
         TabIndex        =   15
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "PhÝ phôc vô:"
         Height          =   375
         Left            =   4440
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tªn khu vùc"
         Height          =   495
         Left            =   4440
         TabIndex        =   13
         Top             =   240
         Width           =   1575
      End
   End
   Begin prjTouchScreen.MyButton cmdClose 
      Height          =   1215
      Left            =   8880
      TabIndex        =   11
      Top             =   2280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2143
      BTYPE           =   5
      TX              =   "§ãn&g"
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
      MICON           =   "frmAdd_Location.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdOk 
      Height          =   1095
      Left            =   8880
      TabIndex        =   10
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1931
      BTYPE           =   5
      TX              =   "§å&ng ý"
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
      MICON           =   "frmAdd_Location.frx":0054
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
Attribute VB_Name = "frmAdd_Location"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLocation As New ADODB.Recordset
Dim Price_Level As Integer
Dim Location_type As Integer


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo Handle
 'Dim rsSection As New ADODB.Recordset
        Dim rsmax As New ADODB.Recordset
            'If Dir(WorkingFolder & "\Database.mdb", vbDirectory) <> "" Then
                'Set rsSection = OpenCriticalTable("select Store_ID,Location_ID,Section_ID,PriceRate,VAT,Price_Level ,Service_Charge,TimeLevel,isTimer from Table_Diagram_Sections ", cnData)
                If txtName.Text = "" Then
                    MsgBox "Tªn khu vùc kh«ng ®­îc rçng !"
                    Exit Sub
                    txtName.SetFocus
                End If
                Set rsmax = OpenCriticalTable("select Max(Location_ID) as MaxID from Table_Diagram_Sections", cnData)
                With rsLocation
                    .addNew
                    .Fields("Store_ID") = Store_ID
                    .Fields("Location_ID") = Format(CDbl("0" & rsmax.Fields("maxID")) + 1, "00")
                    .Fields("Section_ID") = txtName.Text
                    .Fields("PriceRate") = 0
                    .Fields("VAT") = txtVat.Text
                    .Fields("Price_Level") = Price_Level
                    .Fields("Service_Charge") = txtServ.Text
                    .Fields("TimeLevel") = 0
                    .Fields("isTimer") = False
                    .Update
                End With
            'End If
        Exit Sub
Handle:
        MsgBox Err.Number & Err.Description & Me.name & " cmdOk_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Price_Level = 0
    Location_type = 0
    If cnData.State = adStateClosed Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    Set rsLocation = Open_Table(cnData, "Table_Diagram_Sections")
    Set dtgLocation.DataSource = rsLocation
    Init_Column_Name
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  -Form_Load "
End Sub

Public Sub Init_Column_Name()
On Error GoTo Handle
    With dtgLocation
        .Columns(0).Caption = ""
        .Columns(0).Width = 0
        .Columns(1).Caption = ""
        .Columns(1).Width = 0
        .Columns(2).Caption = "Tªn KV"
        .Columns(2).Width = 1600
        .Columns(3).Caption = ""
        .Columns(3).Width = 1600
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  -Init_Column_Name "
End Sub

Private Sub OptPrice_Click(Index As Integer)
On Error GoTo Handle
    Select Case Index
        Case 0
            If OptPrice(Index).Value = True Then Price_Level = 1
        Case 1
           If OptPrice(Index).Value = True Then Price_Level = 2
        Case 2
            If OptPrice(Index).Value = True Then Price_Level = 3
    End Select
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  -OptPrice_Click "
End Sub
