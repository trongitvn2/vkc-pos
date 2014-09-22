VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmKitchenView 
   Caption         =   "HiÓn thÞ th«ng tin gëi  bÕp"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   ClipControls    =   0   'False
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
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimeReFresh 
      Interval        =   1000
      Left            =   14160
      Top             =   4440
   End
   Begin prjTouchScreen.MyButton MyButton1 
      Height          =   855
      Left            =   13920
      TabIndex        =   1
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BTYPE           =   6
      TX              =   "Xong"
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
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmKitchenView.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSDataGridLib.DataGrid gridView 
      Height          =   10980
      Left            =   0
      Negotiate       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   19368
      _Version        =   393216
      BorderStyle     =   0
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   21
      TabAction       =   2
      WrapCellPointer =   -1  'True
      RowDividerStyle =   3
      AllowAddNew     =   -1  'True
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
         Name            =   ".VnArial"
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
   Begin prjTouchScreen.MyButton cmdClose 
      Height          =   855
      Left            =   13920
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      BTYPE           =   6
      TX              =   "&§ãng"
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
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmKitchenView.frx":001C
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
Attribute VB_Name = "frmKitchenView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsKitchen As New ADODB.Recordset
Dim rsPending_Item As New ADODB.Recordset
Dim strIndex As String
Dim countTime As Integer

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Set rsPending_Item = Open_Table(cnData, "Pending_Orders_Items")
    Call Refresh_List
    Call Set_Grid
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name
End Sub

Public Sub Set_Grid()
On Error GoTo Handle
    With gridView
        .Columns(0).Caption = ""
        .Columns(0).Width = 0
        .Columns(1).Caption = "Tªn mãn"
        .Columns(1).Width = 3600
        .Columns(2).Caption = "Sè l­îng"
        .Columns(2).Width = 1200
        .Columns(3).Caption = "§¬n gi¸"
        .Columns(3).Width = 1200
        .Columns(4).Caption = "Chó thÝch"
        .Columns(4).Width = 2500
        .Columns(5).Caption = "Bµn"
        .Columns(5).Width = 1500
        .Columns(6).Caption = "Nh©n viªn"
        .Columns(6).Width = 1700
        .Columns(7).Caption = "Thêi gian Order"
        .Columns(7).Width = 1700
    End With

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name
End Sub

Private Sub gridView_Click()
    strIndex = gridView.Columns(0)
End Sub

Private Sub MyButton1_Click()
On Error GoTo Handle
    With rsPending_Item
        .Find "Index='" & strIndex & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Delete adAffectCurrent
            .Requery
        End If
        
    End With
    Call Refresh_List
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name
End Sub

Private Sub TimeReFresh_Timer()
countTime = countTime + 1
If countTime = 10 Then
    Call Refresh_List
    countTime = 0
End If
End Sub

Public Sub Refresh_List()
On Error GoTo Handle
Dim strSQL As String
    strSQL = "SELECT Pending_Orders_Items.Index,Pending_Orders_Items.ItemName, [Quan]-[QuanBurned] AS Quantity, Pending_Orders_Items.Price, Pending_Orders_Items.Kit_Desc, Pending_Orders.OnHoldID, Pending_Orders.Cashier_ID,Pending_Orders_Items.TimeOrder" & _
             " FROM Pending_Orders INNER JOIN Pending_Orders_Items ON Pending_Orders.Invoice_Number = Pending_Orders_Items.Invoice_Number order by Pending_Orders_Items.TimeOrder"
    Set rsKitchen = OpenCriticalTable(strSQL, cnData)
    Set gridView.DataSource = rsKitchen
    Call Set_Grid
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " Refresh_List"
End Sub
