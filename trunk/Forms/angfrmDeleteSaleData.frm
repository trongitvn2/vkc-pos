VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmDeleteSaleData 
   Caption         =   "Xãa d÷ liÖu b¸n hµng"
   ClientHeight    =   4215
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7140
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
   ScaleHeight     =   4215
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   855
      Left            =   2520
      TabIndex        =   7
      Tag             =   "L8"
      Top             =   3120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "&Tho¸t"
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
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "angfrmDeleteSaleData.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdOK 
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Tag             =   "L5"
      Top             =   2160
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "&D÷ liÖu b¸n hµng"
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
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "angfrmDeleteSaleData.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.Frame Frame1 
      Caption         =   "Chän kho¶ng thêi gian cÇn xãa"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Tag             =   "L2"
      Top             =   840
      Width           =   6855
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   495
         Left            =   1320
         TabIndex        =   2
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
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
         Format          =   70451201
         UpDown          =   -1  'True
         CurrentDate     =   39448
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   495
         Left            =   4680
         TabIndex        =   3
         Top             =   330
         Width           =   1935
         _ExtentX        =   3413
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
         Format          =   70451201
         UpDown          =   -1  'True
         CurrentDate     =   39448
      End
      Begin VB.Label lblFromdate 
         Alignment       =   1  'Right Justify
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
         Left            =   60
         TabIndex        =   5
         Tag             =   "L3"
         Top             =   450
         Width           =   1125
      End
      Begin VB.Label lblDenngay 
         Alignment       =   1  'Right Justify
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
         Left            =   3420
         TabIndex        =   4
         Tag             =   "L4"
         Top             =   420
         Width           =   1125
      End
   End
   Begin prjTouchScreen.MyButton cmdDeleteStock 
      Height          =   855
      Left            =   2520
      TabIndex        =   8
      Tag             =   "L6"
      Top             =   2160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "&D÷ liÖu kho"
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
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "angfrmDeleteSaleData.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdDeleteInOut 
      Height          =   855
      Left            =   4800
      TabIndex        =   9
      Tag             =   "L7"
      Top             =   2160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   "&D÷ liÖu Thu chi"
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
      FCOLO           =   255
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "angfrmDeleteSaleData.frx":0054
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
      Alignment       =   2  'Center
      Caption         =   "xãa d÷ liÖu"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2160
      TabIndex        =   0
      Tag             =   "L1"
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmDeleteSaleData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsInvoice_Onhold As New ADODB.Recordset
Dim rsInvoice_Items As New ADODB.Recordset
Dim rsInvoice_Total As New ADODB.Recordset
Dim rsInvoice_Totals_Notes As New ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeleteInOut_Click()
On Error GoTo Handle
    Dim fdate, tdate As String
    If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    fdate = gfCONVERT_DATE_TO_STRING(dtpFromDate.Value)
    tdate = gfCONVERT_DATE_TO_STRING(dtpToDate.Value)
    cnData.Execute "Delete  from Payouts where left([DateTime],8)>='" & fdate & "' and left([DateTime],8)<='" & tdate & "'"
    cnData.Execute "Delete  from Income where left([DateTime],8)>='" & fdate & "' and left([DateTime],8)<='" & tdate & "'"
    MsgBox "§· xãa xong", vbInformation
Exit Sub
Handle:
    MsgBox Err.Description & " cmdDeleteInOut_Click"
End Sub

Private Sub cmdDeleteStock_Click()
On Error GoTo Handle
    Dim fdate, tdate As String
    If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    fdate = gfCONVERT_DATE_TO_STRING(dtpFromDate.Value)
    tdate = gfCONVERT_DATE_TO_STRING(dtpToDate.Value)
    cnData.Execute "Delete  from Instock_MasterB where left([DateTime],8)>='" & fdate & "' and left([DateTime],8)<='" & tdate & "'"
    Call Delete_InB(fdate, tdate)
    MsgBox "§· xãa xong", vbInformation
Exit Sub
Handle:
    MsgBox Err.Description & " cmdDeleteStock_Click"
End Sub

Private Sub cmdOK_Click()
On Error GoTo Handle
If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    If dtpFromDate.Value > dtpToDate.Value Then
        MsgBox "B¹n chän sai gi¸ trÞ ngµy !"
        dtpFromDate.Value = dtpToDate.Value
    Else
        If MsgBox("B¹n cã ch¾c ch¾n muèn xãa d÷ liÖu b¸n hµng tõ ngµy " & dtpFromDate.Value & " ®Õn ngµy " & dtpToDate.Value & " kh«ng?", vbYesNo) = vbYes Then
            Dim str As String
                str = "select * from Invoice_Totals where left(Invoice_Totals.[DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' and left(Invoice_Totals.[DateTime],8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'"
            Set rsInvoice_Total = OpenCriticalTable(str, cnData)
            If rsInvoice_Total.RecordCount > 0 Then
                With rsInvoice_Total
                    Do While Not .EOF
                    If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
                     cnData.Execute "Delete  from Invoice_Itemized  where Invoice_Number=" & rsInvoice_Total.Fields("Invoice_Number")
                     cnData.Execute "Delete  from Items_Deleted where Invoice_Num=" & rsInvoice_Total.Fields("Invoice_Number") & " and left(Items_Deleted.[DateTime],8)>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "' and left(Items_Deleted.[DateTime],8)<='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'"
                    cnData.Execute "Delete  from Kitchen_Order_Items where Invoice_Number=" & rsInvoice_Total.Fields("Invoice_Number")
                     cnData.Execute "Delete  from Kitchen_Order_Master where Invoice_Number=" & rsInvoice_Total.Fields("Invoice_Number")
                     cnData.Execute "Delete  from Invoice_Totals_Person_Mapping where Invoice_Number=" & rsInvoice_Total.Fields("Invoice_Number")
                    cnData.Execute "Delete  from Invoice_Totals where Invoice_Number=" & rsInvoice_Total.Fields("Invoice_Number")
                    cnData.Execute "Delete  from Invoice_Totals_Time where Invoice_Number=" & rsInvoice_Total.Fields("Invoice_Number")
                        With rsInvoice_Onhold
                            .Find "Invoice_Number=" & rsInvoice_Total.Fields("Invoice_Number"), , adSearchForward, adBookmarkFirst
                            If Not .EOF Then
                                .Delete adAffectCurrent
                            End If
                        End With
                        
                         With rsInvoice_Totals_Notes
                            .Find "Invoice_Number=" & rsInvoice_Total.Fields("Invoice_Number"), , adSearchForward, adBookmarkFirst
                            If Not .EOF Then
                                .Delete adAffectCurrent
                            End If
                        End With
                        '.Delete adAffectCurrent
                    .MoveNext
                    Loop
                End With
            End If
            If cnData.State = 0 Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
         cnData.Execute "Delete  from Tranfer_Joint_table where left(Tranfer_Joint_table.[DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' and left(Tranfer_Joint_table.[DateTime],8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'"
         
        End If
    End If

MsgBox "Hoµn tÊt !"
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdOK_Click "
End Sub

Private Sub Form_Activate()
On Error GoTo errHdl
    Dim DescArr() As String
    Dim ctrl As Control
    DescArr = LoadLanguage(LngFile, "#03:013:")
    If cmdOk.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & "- Form_Activate"

End Sub

Private Sub Form_Load()
On Error GoTo Handle
'If cnData.State <> 0 Then
'    Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'End If

Set rsInvoice_Totals_Notes = Open_Table(cnData, "Invoice_Totals_Notes")
Set rsInvoice_Onhold = Open_Table(cnData, "Invoice_OnHold")
    dtpFromDate.Value = Date
    dtpToDate.Value = Date
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Public Sub Delete_InB(ByVal fdate As String, ByVal tdate As String)
On Error GoTo Handle
    Dim i As Integer
    i = Mid(fdate, 5, 2)
    If Mid(fdate, 3, 2) = Mid(tdate, 3, 2) Then
        If Mid(fdate, 5, 2) <> Mid(tdate, 5, 2) Then
            Do While i <= Mid(tdate, 5, 2)
                If Check_Table_exist("Inventory_InB" & Format(i, "00") & Mid(fdate, 3, 2)) Then
                    cnData.Execute "Delete  from Inventory_InB" & Format(i, "00") & Mid(fdate, 3, 2) & " where DateTime >='" & fdate & "' and DateTime<='" & tdate & "'"
                End If
                If Check_Table_exist("TonB" & Format(i, "00") & Mid(fdate, 3, 2)) Then
                    cnData.Execute "Delete  from TonB" & Format(i, "00") & Mid(fdate, 3, 2)
                    cnData.Execute "DROP TABLE TonB" & Format(i, "00") & Mid(fdate, 3, 2)
                End If
                i = i + 1
            Loop
        Else
            cnData.Execute "Delete  from Inventory_InB" & Format(i, "00") & Mid(fdate, 3, 2) & " where DateTime >='" & fdate & "' and DateTime<='" & tdate & "'"
        End If
    Else
        MsgBox "ChØ cho phÐp xãa d÷ liÖu kho tõng n¨m", vbInformation
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Delete_InB"

End Sub
