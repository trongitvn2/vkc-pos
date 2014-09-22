VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form frmLocation_Cashier 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ph©n nh©n viªn theo khu vùc"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11925
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
   Icon            =   "frmLocation_Cashier.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   11925
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Nh©n viªn"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8715
      Left            =   75
      TabIndex        =   1
      Top             =   675
      Width           =   11790
      Begin prjTouchScreen.MyButton cmdClose 
         Height          =   1215
         Left            =   7200
         TabIndex        =   6
         Top             =   7275
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   2143
         BTYPE           =   6
         TX              =   "&L­u / §ãng"
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
         BCOL            =   16711680
         BCOLO           =   33023
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmLocation_Cashier.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.Frame Frame2 
         Caption         =   "Khu vùc"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6840
         Left            =   6525
         TabIndex        =   3
         Top             =   225
         Width           =   5115
         Begin VB.TextBox txtFlag 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1725
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   150
            Width           =   1740
         End
         Begin VB.ListBox lstFlag 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   6060
            Left            =   75
            Style           =   1  'Checkbox
            TabIndex        =   4
            Top             =   675
            Width           =   4995
         End
      End
      Begin MSDataGridLib.DataGrid dtgNhanvien 
         Height          =   8190
         Left            =   75
         TabIndex        =   2
         Top             =   300
         Width           =   6390
         _ExtentX        =   11271
         _ExtentY        =   14446
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   0   'False
         HeadLines       =   1
         RowHeight       =   21
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Ph©n chia nh©n viªn theo khu vùc"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   465
      Left            =   1875
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmLocation_Cashier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCashier As New ADODB.Recordset
Dim rsCashier_Location As New ADODB.Recordset
Dim rsLocation As New ADODB.Recordset
Dim cashier_ID As String

Private Sub cmdClose_Click()
    Set rsCashier = Nothing
    Unload Me
End Sub

Private Sub dtgNhanvien_Click()
On Error GoTo Handle
    With rsCashier_Location
        If .State <> 0 Then
            If .RecordCount > 0 Then .MoveFirst
        Else
            Exit Sub
        End If
        cashier_ID = dtgNhanvien.Columns(0).Value
        .Find "Cashier_ID='" & cashier_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            txtFlag.Text = .Fields("Location")
            AddValueForListFlag txtFlag.Text, lstFlag
        Else
            .addNew
            .Fields("Cashier_ID") = cashier_ID
            .Fields("Location") = txtFlag.Text
            .Update
            .Requery
        End If
        
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & " dtgNhanvien_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    Set rsCashier = LoadPasswordData
    Set rsLocation = Open_Table(cnData, "Table_Diagram_Sections")
    Set rsCashier_Location = Open_Table(cnData, "Stations")
    Call Set_DataGrid
    Call Add_Location_to_List
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " Form_Load"
End Sub

Public Sub Set_DataGrid()
On Error GoTo Handle
    With dtgNhanvien
     Set .DataSource = rsCashier
        .Font.Name = ".vnArial"
        .Columns(0).Caption = "Ma NV"
        .Columns(1).Caption = "Ten Nhan vien"
        .Columns(2).Caption = "Cap do"
        .Columns(0).Width = 1000
        .Columns(1).Width = 4000
        .Columns(2).Width = 1200
        .Columns(3).Width = 0
        .Columns(4).Width = 0
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.Name & "  Set_DataGrid"
End Sub

Public Sub Add_Location_to_List()
    On Error GoTo Handle
    Dim i As Integer
        lstFlag.Clear
        With rsLocation
            Do While Not rsLocation.EOF
                lstFlag.AddItem .Fields("Section_ID")
            .MoveNext
            i = i + 1
            Loop
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & " Add_Location_to_List"
End Sub

Private Sub lstFlag_Click()
On Error GoTo Handle
    Dim strflag As String
    Dim i As Integer
        strflag = ""
        For i = 0 To lstFlag.ListCount - 1
            If lstFlag.Selected(i) Then
                strflag = strflag & "1"
            Else: strflag = strflag & "0"
            End If
        Next
        txtFlag.Text = FillZeroForString(BinToHex(strflag), 2)
        With rsCashier_Location
            If .RecordCount > 0 Then .MoveFirst
            .Find "Cashier_ID='" & cashier_ID & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Fields("Location") = txtFlag.Text
                .Update
                .Requery
            End If
        End With
Exit Sub

Handle:
MsgBox Err.Number & Err.Description & Me.Name & ""
End Sub

Private Sub txtFlag_Change()
    AddValueForListFlag txtFlag.Text, lstFlag
End Sub
Public Sub AddValueForListFlag(ByVal str1 As String, ByVal lst As ListBox)
On Error GoTo errHdl

    Dim strBin As String
    Dim k As Integer
    
    strBin = HexToBin(str1)
    strBin = FillZeroForString(strBin, rsLocation.RecordCount)
    For k = 0 To Len(strBin) - 1 Step 1
    DoEvents
        If Mid(strBin, k + 1, 1) = 1 Then
            lst.Selected(k) = True
        Else
            lst.Selected(k) = False
        End If
    Next k
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "mdlGlobal - AddValueForList"
End Sub

