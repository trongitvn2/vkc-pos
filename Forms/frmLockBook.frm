VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmLockBook 
   Caption         =   "Khãa sæ"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
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
   Icon            =   "frmLockBook.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton cmdLock 
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "&Khãa sæ"
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
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLockBook.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      CustomFormat    =   "MM/yyyy"
      Format          =   63569923
      UpDown          =   -1  'True
      CurrentDate     =   40283
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   3840
      TabIndex        =   3
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
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
      BCOL            =   12632256
      BCOLO           =   33023
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmLockBook.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Khãa sæ th¸ng"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frmLockBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLock_Month As New ADODB.Recordset
Dim month_lock As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLock_Click()
    On Error GoTo Handle
    If MsgBox("B¹n ®· tÝnh tån kho vµ hoµn tÊt d÷ liÖu th¸ng " & month_lock & " ch­a?", vbYesNo) = vbYes Then
        If MsgBox("B¹n cã ch¾c ch¾n muèn khãa sæ th¸ng " & month_lock & " kh«ng?", vbYesNo) = vbYes Then
            With rsLock_Month
                .Find "Month_Lock='" & month_lock & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    If .Fields("Value") = True Then
                        MsgBox "Th¸ng nµy ®· khãa sæ råi !"
                    Else
                        .Fields("Value") = True
                        .Update
                    End If
                Else
                    .addNew
                    .Fields("Month_Lock") = month_lock
                    .Fields("Date_Lock") = gfCONVERT_DATE_TO_STRING(Date)
                    .Fields("Value") = True
                    .Update
                    Call Lock_Book(month_lock)
                    MsgBox "§· khãa sæ th¸ng " & month_lock
                End If
            End With
        End If
    Else
        frmCal_TonTemp.Show vbModal
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdLock_Click()"
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
    month_lock = Format(Month(DTPicker1.Value), "00") & Format(Year(DTPicker1.Value), "0000")
End Sub

Private Sub DTPicker1_Change()
    month_lock = Format(Month(DTPicker1.Value), "00") & Format(Year(DTPicker1.Value), "0000")
End Sub

Private Sub DTPicker1_Click()
    month_lock = Format(Month(DTPicker1.Value), "00") & Format(Year(DTPicker1.Value), "0000")
End Sub

Private Sub Form_Load()
On Error GoTo Handle
    DTPicker1.Value = Format(Month(Date), "00") & "/" & Format(Year(Date), "0000")
    month_lock = Format(Month(DTPicker1.Value), "00") & Format(Year(DTPicker1.Value), "0000")
        If Not Check_Table_exist("Lock_Month") Then
            Call Create_Table_Lock
        End If
    If cnData.State <> 0 Then Set rsLock_Month = Open_Table(cnData, "Lock_Month")
    
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Form_Load()"
End Sub

Public Sub Lock_Book(ByVal mLock As String)
On Error GoTo Handle
Dim cnLock_data As New ADODB.Connection
Dim strLock_Path As String
Dim fso As New FileSystemObject

    strLock_Path = WorkingFolder & "\Data " & month_lock
    If Dir(strLock_Path, vbDirectory) = "" Then MkDir (strLock_Path)
     fso.CopyFile WorkingFolder & "\Database.mdb", strLock_Path & "\Database.mdb", True
     fso.CopyFile WorkingFolder & "\LoginData.dat", strLock_Path & "\LoginData.dat", True
     If Dir(strLock_Path & "\Log", vbDirectory) = "" Then MkDir (strLock_Path & "\Log")
     If cnLock_data.State = 0 Then Set cnLock_data = Get_Connection()
     Call Delete_Sale_Data_Lock(cnLock_data)
     'xãa d÷ liÖu trong data hiÖn t¹i
     If MsgBox("B¹n cã muèn xãa d÷ liÖu th¸ng " & month_lock & " ®· khãa sæ kh«ng", vbYesNo) = vbYes Then
        Call Delete_Sale_Data(cnData)
     End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Lock_Book"
End Sub


Public Sub Delete_Sale_Data_Lock(cnLock_data As Connection)
On Error GoTo Handle
    With cnLock_data
    If Left(month_lock, 2) = "01" Then
        'xãa d÷ liÖu th¸ng tr­íc khãa sæ
        .Execute "DELETE Invoice_Totals.*, Invoice_Totals_Notes.*" & _
                        " FROM Invoice_Totals_Notes INNER JOIN Invoice_Totals ON Invoice_Totals_Notes.Invoice_Number = Invoice_Totals.Invoice_Number " & _
                        " where left(Invoice_Totals.DateTime,6)='" & Format(Mid(month_lock, 3, 4) - 1, "0000") & "12" & "'"
        .Execute "DELETE Items_Deleted.*, Items_Deleted.DateTime FROM Items_Deleted WHERE left(Items_Deleted.DateTime,6)='" & Format(Mid(month_lock, 3, 4) - 1, "0000") & "12" & "'"
        
        'xoa du lieu sau th¸ng khãa sæ
        .Execute "DELETE Invoice_Totals.*, Invoice_Totals_Notes.*" & _
                        " FROM Invoice_Totals_Notes INNER JOIN Invoice_Totals ON Invoice_Totals_Notes.Invoice_Number = Invoice_Totals.Invoice_Number " & _
                        " where left(Invoice_Totals.DateTime,6)='" & Mid(month_lock, 3, 4) & Format(Left(month_lock, 2) + 1, "00") & "'"
        .Execute "DELETE Items_Deleted.*, Items_Deleted.DateTime FROM Items_Deleted WHERE left(Items_Deleted.DateTime,6)='" & Format(Mid(month_lock, 3, 4) - 1, "0000") & Format(Left(month_lock, 2) + 1, "00") & "'"
        
    Else
        .Execute "DELETE Invoice_Totals.*, Invoice_Totals_Notes.*" & _
                        " FROM Invoice_Totals_Notes INNER JOIN Invoice_Totals ON Invoice_Totals_Notes.Invoice_Number = Invoice_Totals.Invoice_Number " & _
                        " where left(Invoice_Totals.DateTime,6)<='" & Mid(month_lock, 3, 4) & Format(Left(month_lock, 2) - 1, "00") & "'"
        .Execute "DELETE Items_Deleted.*, Items_Deleted.DateTime FROM Items_Deleted WHERE left(Items_Deleted.DateTime,6)='" & Mid(month_lock, 3, 4) & Format(Left(month_lock, 2) - 1, "00") & "'"
    'Xoa d÷ liÖu sau th¸ng khãa sæ
        .Execute "DELETE Invoice_Totals.*, Invoice_Totals_Notes.*" & _
                        " FROM Invoice_Totals_Notes INNER JOIN Invoice_Totals ON Invoice_Totals_Notes.Invoice_Number = Invoice_Totals.Invoice_Number " & _
                        " where left(Invoice_Totals.DateTime,6)='" & Mid(month_lock, 3, 4) & Format(Left(month_lock, 2) + 1, "00") & "'"
        .Execute "DELETE Items_Deleted.*, Items_Deleted.DateTime FROM Items_Deleted WHERE left(Items_Deleted.DateTime,6)='" & Format(Mid(month_lock, 3, 4) - 1, "0000") & Format(Left(month_lock, 2) + 1, "00") & "'"
        If Check_Table_exist("Inventory_InB" & Format(Left(month_lock, 2) - 2, "00") & Mid(month_lock, 5, 2)) Then
            .Execute "Drop table Inventory_InB" & Format(Left(month_lock, 2) - 2, "00") & Mid(month_lock, 5, 2)
        End If
        If Check_Table_exist("TonB" & Format(Left(month_lock, 2) - 2, "00") & Mid(month_lock, 5, 2)) Then
            .Execute "Drop table TonB" & Format(Left(month_lock, 2) - 2, "00") & Mid(month_lock, 5, 2)
        End If
    End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Delete_Sale_Data"
End Sub

Public Sub Delete_Sale_Data(cnLock_data As Connection)
On Error GoTo Handle
    With cnLock_data
        .Execute "DELETE Invoice_Totals.*, Invoice_Totals_Notes.*" & _
                        " FROM Invoice_Totals_Notes INNER JOIN Invoice_Totals ON Invoice_Totals_Notes.Invoice_Number = Invoice_Totals.Invoice_Number " & _
                        " where left(Invoice_Totals.DateTime,6)='" & Mid(month_lock, 3, 4) & Left(month_lock, 2) & "'"
        .Execute "DELETE Items_Deleted.*, Items_Deleted.DateTime FROM MainGroup, Items_Deleted WHERE left(Items_Deleted.DateTime,6)='" & Mid(month_lock, 3, 4) & Left(month_lock, 2) & "'"
        If Check_Table_exist("Inventory_InB" & Format(Left(month_lock, 2) - 1, "00") & Mid(month_lock, 5, 2)) Then
        .Execute "Drop table Inventory_InB" & Format(Left(month_lock, 2) - 1, "00") & Mid(month_lock, 5, 2)
        End If
        If Check_Table_exist("TonB" & Format(Left(month_lock, 2) - 1, "00") & Mid(month_lock, 5, 2)) Then
            .Execute "Drop table TonB" & Format(Left(month_lock, 2) - 1, "00") & Mid(month_lock, 5, 2)
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Delete_Sale_Data"
End Sub

