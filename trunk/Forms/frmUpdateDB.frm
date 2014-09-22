VERSION 5.00
Begin VB.Form frmUpdateDB 
   Caption         =   "CËp nhËt d÷ liÖu"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8835
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
   Icon            =   "frmUpdateDB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   8835
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton cmdUpdateCancelOrder 
      Height          =   975
      Left            =   3120
      TabIndex        =   9
      Top             =   3240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Add Table Hñy order"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdateDB.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdclose 
      Height          =   975
      Left            =   5880
      TabIndex        =   4
      Top             =   5520
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "§ãng"
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
      BCOL            =   255
      BCOLO           =   255
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdateDB.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdReserve 
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "CËp nhËt d÷ liÖu ®Æt bµn"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdateDB.frx":0044
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdSup 
      Height          =   975
      Left            =   3120
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Danh môc nhµ cung cÊp mÆc ®Þnh"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdateDB.frx":0060
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdIncom 
      Height          =   975
      Left            =   5880
      TabIndex        =   2
      Top             =   1080
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "CËp nhËt Kho¶n thu mÆc ®Þnh"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdateDB.frx":007C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdPayout 
      Height          =   975
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "CËp nhËt Kho¶n chi mÆc ®Þnh"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdateDB.frx":0098
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdKitchen 
      Height          =   975
      Left            =   3120
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "CËp nhËt Ticket Order"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdateDB.frx":00B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdUpdateBCTH 
      Height          =   975
      Left            =   5880
      TabIndex        =   7
      Top             =   2160
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "CËp nhËt b¸o c¸o tæng hîp"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdateDB.frx":00D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdUpdate_Invoice_Totals 
      Height          =   975
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "CËp nhËt d÷ liÖu gi¶m gi¸"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdateDB.frx":00EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdSync 
      Height          =   975
      Left            =   5880
      TabIndex        =   10
      Top             =   3240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "Add Column Synchronized on Invoice_Totals"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdateDB.frx":0108
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdUpdateKar 
      Height          =   975
      Left            =   3120
      TabIndex        =   11
      Top             =   4320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "CËp nhËt tÝnh giê"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdateDB.frx":0124
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdUpdateKar_menu 
      Height          =   975
      Left            =   360
      TabIndex        =   12
      Top             =   4320
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "CËp nhËt danh môc giê"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdateDB.frx":0140
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdKho 
      Height          =   975
      Left            =   5880
      TabIndex        =   13
      Top             =   4320
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1720
      BTYPE           =   3
      TX              =   "CËp nhËt c«ng nî nhËp kho"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   8421631
      BCOLO           =   8421631
      FCOL            =   16711680
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmUpdateDB.frx":015C
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
      BackStyle       =   0  'Transparent
      Caption         =   "cËp nhËt cÊu tróc d÷ liÖu "
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   360
      Width           =   6615
   End
End
Attribute VB_Name = "frmUpdateDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdKho_Click()
     On Error GoTo Handle
    Dim rsInOut As New ADODB.Recordset
    Set rsInOut = Open_Table(cnData, "Instock_MasterB")
    If Not Check_Field_Exist(rsInOut, "Payment_Method") Then
        cnData.Execute "ALTER TABLE Instock_MasterB ADD [Payment_Method] nvarchar(2),Totals float"
    End If
    MsgBox "Update OK"
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdKitchen_Click"
    
End Sub

Private Sub cmdKitchen_Click()
    On Error GoTo Handle
    Dim rspending As New ADODB.Recordset
    Set rspending = Open_Table(cnData, "Pending_Orders_Items")
    If Not Check_Field_Exist(rspending, "Count") Then
        cnData.Execute "ALTER TABLE Pending_Orders_Items ADD COLUMN [Count] double"
    End If
    MsgBox "Update OK"
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdKitchen_Click"
End Sub

Private Sub cmdPayout_Click()
On Error GoTo Handle
Dim rsChi As New ADODB.Recordset
Set rsChi = Open_Table(cnData, "Expense")
If rsChi.State <> 0 Then
    With rsChi
        .Find "MaChi='" & "T§C" & "'", , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields("MaChi") = "T§C"
            .Fields("DienGiai") = "Trõ §Æt cäc"
            .Update
        End If
        MsgBox "Update OK"
    End With
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdPayout_Click"
End Sub

Private Sub cmdReserve_Click()
    On Error GoTo Handle
    Dim rs As New ADODB.Recordset
    Dim rsReserve As New ADODB.Recordset
    If Not Check_Table_exist("Table_Reserved_Details") Then
        Call Create_Table_Reserved_Details
        Call Create_Table_Reserverd
    End If
    Set rs = Open_Table(cnData, "Table_Reserved_Details")
    Set rsReserve = Open_Table(cnData, "Table_Reservered")
    
        If Not Check_Field_Exist(rs, "LineDisc") Then
            cnData.Execute "ALTER TABLE Table_Reserved_Details ADD COLUMN LineDisc Double ,Line_Disc_Desc char,Kit_Desc char"
        End If
        If Not Check_Field_Exist(rsReserve, "Section_ID") Then
            cnData.Execute "ALTER TABLE Table_Reservered ADD COLUMN Section_ID char"
        End If
          MsgBox "CËp nhËt hoµn tÊt"
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub cmdSup_Click()
On Error GoTo Handle
Dim rssup As New ADODB.Recordset
If Not Check_Table_exist("Vendors") Then Exit Sub
Set rssup = Open_Table(cnData, "Vendors")
If rssup.State <> 0 Then
    With rssup
        .Find "Vendor_Number='" & "0000" & "'", , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields("Vendor_Number") = "0000"
            .Fields("Vendor_Name") = "MÆc ®Þnh"
            .Fields("Company") = "-"
            .Fields("Address_1") = "-"
            .Fields("Address_2") = "-"
            .Update
        End If
        MsgBox "Update OK"
    End With
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdSup_Click"
End Sub

Private Sub cmdIncom_Click()
On Error GoTo Handle
Dim rsReceipt As New ADODB.Recordset
Set rsReceipt = Open_Table(cnData, "Receipt")
If rsReceipt.State <> 0 Then
    With rsReceipt
        .Find "MaThu='" & "§C" & "'", , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields("MaThu") = "§C"
            .Fields("DienGiai") = "§Æt cäc"
            .Update
        End If
        MsgBox "Update OK"
    End With
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdSup_Click"
End Sub

Private Sub cmdUpdate_Invoice_Totals_Click()
On Error GoTo Handle
Dim rsInvoice_Totals As New ADODB.Recordset
Dim rsdiscount As New ADODB.Recordset
Dim i As Integer
 Set rsdiscount = Open_Table(cnData, "Adjustment")
  With rsdiscount
      For i = 5 To 7
        .Find "AdjNo='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields("AdjNo") = Format(i, "00")
            If i <> 7 Then
                .Fields("AdjName") = "Adjustment " & i
            Else
                .Fields("AdjName") = "Gi¶m tæng Hãa ®¬n"
            End If
            .Fields("AdjRate") = 0
            .Update
        Else
            Exit For
            .Fields("AdjName") = "Adjustment 5"
            .Update
        End If
    Next
End With
                
 
Set rsInvoice_Totals = Open_Table(cnData, "Invoice_Totals")
    If Not Check_Field_Exist(rsInvoice_Totals, "Adj3Rate") Then
        cnData.Execute "ALTER TABLE Invoice_Totals ADD COLUMN Adj3Rate Double,Adj4Rate double, Adj5Rate double, Adj6Rate double, Adjustment5 double, Adjustment6 double "
    End If
    MsgBox "Update OK", vbInformation
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   cmdUpdateBCTH_Click"

    
End Sub

Private Sub cmdUpdateBCTH_Click()
On Error GoTo Handle
Dim rsGeneral As New ADODB.Recordset
Set rsGeneral = OpenCriticalTable("select * from RP_General", cnData)
    If Not Check_Field_Exist(rsGeneral, "CountGC") Then
        cnData.Execute "ALTER TABLE RP_General " _
                                 & "ADD COLUMN CountGC Double,AmountGC Double,CountROA Double,AmountROA Double;"
    End If
    If Not Check_Field_Exist(rsGeneral, "CountReserve") Then
        cnData.Execute "ALTER TABLE RP_General " _
                                 & "ADD COLUMN CountReserve Double, AmountReserve Double"
    End If
    If Not Check_Field_Exist(rsGeneral, "CountAdj3") Then
        cnData.Execute "ALTER TABLE RP_General " _
                                 & "ADD COLUMN CountAdj3 Double,CountAdj4 Double,CountAdj5 Double,CountAdj6 Double,Adjustment3 Double,Adjustment4 Double,Adjustment5 Double,Adjustment6 Double"
    End If
    
    MsgBox "Update OK", vbInformation
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   cmdUpdateBCTH_Click"
End Sub

Private Sub cmdUpdateCancelOrder_Click()
On Error GoTo Handle
    Dim str, str1 As String
    If cnData.State = adStateClosed Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    If Not Check_Table_exist("Invoice_Cancel") Then
        str = "CREATE TABLE [dbo].[Invoice_Cancel]([Invoice_Number] [float] NOT NULL, [DateTime] [nVarchar] (20) NOT NULL,[Staion_ID] [nvarchar](12) NOT NULL," & _
                " [Table_ID] [nvarchar](20) NOT NULL, [Cashier_ID] [nvarchar](20) NOT NULL, [Cashier_Cancel] [nvarchar](20) NOT NULL,[CO_DateTime] [nvarchar](20) NOT NULL," & _
                " [Invoice_Status][nvarchar](2)Null," & _
                " CONSTRAINT [PK_Invoice_Cancel] PRIMARY KEY CLUSTERED ( [Invoice_Number] Asc)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]) ON [PRIMARY] "
        str1 = "CREATE TABLE [dbo].[Invoice_Cancel_Items]([Invoice_Number] [float] NOT NULL,[LineNum] [int] NOT NULL," & _
                " [ItemNum] [nvarchar](13) NOT NULL, [ItemName] [nvarchar](100) NOT NULL,[Quantity] [float] NOT NULL," & _
                " [Price] [float] NOT NULL,[Item_CO_DateTime] [nvarchar](20) NOT NULL,[Kit_Desc] [nvarchar](100) NULL" & _
                " ) ON [PRIMARY]"
        cnData.Execute str
        cnData.Execute str1
        MsgBox "Update is OK!", vbInformation
    Else
        MsgBox "Don't need update DB !", vbInformation
    End If
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdUpdateCancelOrder_Click"
End Sub

Private Sub cmdSync_Click()
On Error GoTo Handle
Dim rsTotals As New ADODB.Recordset
Set rsTotals = OpenCriticalTable("select * from Invoice_Totals", cnData)
    If Not Check_Field_Exist(rsTotals, "Synchronized") Then
        cnData.Execute "ALTER TABLE Invoice_Totals " _
                                 & "ADD COLUMN Synchronized bit;"
    End If
    
    MsgBox "Update is OK", vbInformation
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   cmdSync_Click"
End Sub
Private Sub cmdUpdateKar_Click()
On Error GoTo Handle
Dim rsKarSetup As New ADODB.Recordset
Set rsKarSetup = OpenCriticalTable("select * from Table_Diagram_Sections", cnData)
    If Not Check_Field_Exist(rsKarSetup, "isTimer") Then
        cnData.Execute "ALTER TABLE Table_Diagram_Sections " _
                                 & "ADD  TimeLevel int null, isTimer bit;"
        cnData.Execute "update  Table_Diagram_Sections " _
                                 & "set  TimeLevel =0, isTimer 0;"
                                                          
    End If
    cnData.Execute "update  Table_Diagram_Sections " _
                                 & "set  TimeLevel =0, isTimer =0"
    MsgBox "Update is OK", vbInformation
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   cmdUpdateKar_Click"
End Sub

Private Sub cmdUpdateKar_menu_Click()
    Call Update_MainGroup
    Call Update_Group
    Call Update_Items
    MsgBox "Update OK"
End Sub
Public Sub Update_MainGroup()
On Error GoTo Handle
Dim rsMain As New ADODB.Recordset
Set rsMain = Open_Table(cnData, "MainGroup")
If rsMain.State <> 0 Then
    With rsMain
        .Find "GroupNo='" & "99" & "'", , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields("GroupNo") = "99"
            .Fields("GroupName") = "TiÒn giê"
            .Update
        End If
    End With
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Update_Maingroup"
End Sub

Public Sub Update_Group()
On Error GoTo Handle
Dim rsDepartments As New ADODB.Recordset
Set rsDepartments = Open_Table(cnData, "Departments")
If rsDepartments.State <> 0 Then
    With rsDepartments
        .Find "Dept_ID='" & "999" & "'", , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields("GIndex") = "999"
            .Fields("Dept_ID") = "999"
            .Fields("Store_ID") = "01"
            .Fields("Description") = "TiÒn giê"
            .Fields("MainGroup") = "99"
            .Fields("F") = "00"
            .Fields("ColorDept") = "ffff"
            .Update
        End If
    End With
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Update_Group"

End Sub

Public Sub Update_Items()
On Error GoTo Handle
Dim rsInventory As New ADODB.Recordset
Set rsInventory = Open_Table(cnData, "Inventory")
If rsInventory.State <> 0 Then
    With rsInventory
        .Find "Dept_ID='" & "KAR11" & "'", , adSearchForward, adBookmarkFirst
       
        If .EOF Then
             For i = 1 To 3
                .addNew
                .Fields("ItemNum") = "KAR" & i
                .Fields("ItemName") = "TiÒn giê"
                .Fields("Dept_ID") = "999"
                .Fields("Std_Price1") = "0"
                .Fields("Std_Price2") = "0"
                .Fields("Std_Price3") = "0"
                .Fields("HH_Price1") = "0"
                .Fields("HH_Price2") = "0"
                .Fields("HH_Price3") = "0"
                .Fields("EV_Price1") = "0"
                .Fields("EV_Price2") = "0"
                .Fields("EV_Price3") = "0"
                .Fields("LimitPrice") = "80FFFF"
                .Fields("Unit") = "'"
                .Fields("Minstock") = "0"
                .Fields("Modify_Number") = "0"
                .Fields("F1") = "00"
                .Fields("F2") = "00"
                .Fields("F3") = "00"
                .Fields("F4") = "00"
                .Fields("F5") = "00"
                .Fields("Date_Created") = "2014/01/03"
                .Fields("Picture") = "-"
                .Fields("Print_On_Receipt") = "1"
                .Fields("Store_ID") = "01"
                .Update
            Next
            For i = 1 To 3
                .addNew
                .Fields("ItemNum") = "KAR2" & i
                .Fields("ItemName") = "TiÒn giê"
                .Fields("Dept_ID") = "999"
                .Fields("Std_Price1") = "0"
                .Fields("Std_Price2") = "0"
                .Fields("Std_Price3") = "0"
                .Fields("HH_Price1") = "0"
                .Fields("HH_Price2") = "0"
                .Fields("HH_Price3") = "0"
                .Fields("EV_Price1") = "0"
                .Fields("EV_Price2") = "0"
                .Fields("EV_Price3") = "0"
                .Fields("LimitPrice") = "80FFFF"
                .Fields("Unit") = "'"
                .Fields("Minstock") = "0"
                .Fields("Modify_Number") = "0"
                .Fields("F1") = "00"
                .Fields("F2") = "00"
                .Fields("F3") = "00"
                .Fields("F4") = "00"
                .Fields("F5") = "00"
                .Fields("Date_Created") = "2014/01/03"
                .Fields("Picture") = "-"
                .Fields("Print_On_Receipt") = "1"
                .Fields("Store_ID") = "01"
                .Update
            Next
            For i = 1 To 3
                .addNew
                .Fields("ItemNum") = "KAR3" & i
                .Fields("ItemName") = "TiÒn giê"
                .Fields("Dept_ID") = "999"
                .Fields("Std_Price1") = "0"
                .Fields("Std_Price2") = "0"
                .Fields("Std_Price3") = "0"
                .Fields("HH_Price1") = "0"
                .Fields("HH_Price2") = "0"
                .Fields("HH_Price3") = "0"
                .Fields("EV_Price1") = "0"
                .Fields("EV_Price2") = "0"
                .Fields("EV_Price3") = "0"
                .Fields("LimitPrice") = "80FFFF"
                .Fields("Unit") = "'"
                .Fields("Minstock") = "0"
                .Fields("Modify_Number") = "0"
                .Fields("F1") = "00"
                .Fields("F2") = "00"
                .Fields("F3") = "00"
                .Fields("F4") = "00"
                .Fields("F5") = "00"
                .Fields("Date_Created") = "2014/01/03"
                .Fields("Picture") = "-"
                .Fields("Print_On_Receipt") = "1"
                .Fields("Store_ID") = "01"
                .Update
            Next
            For i = 1 To 3
                .addNew
                .Fields("ItemNum") = "KAR1" & i
                .Fields("ItemName") = "TiÒn giê"
                .Fields("Dept_ID") = "999"
                .Fields("Std_Price1") = "0"
                .Fields("Std_Price2") = "0"
                .Fields("Std_Price3") = "0"
                .Fields("HH_Price1") = "0"
                .Fields("HH_Price2") = "0"
                .Fields("HH_Price3") = "0"
                .Fields("EV_Price1") = "0"
                .Fields("EV_Price2") = "0"
                .Fields("EV_Price3") = "0"
                .Fields("LimitPrice") = "80FFFF"
                .Fields("Unit") = "'"
                .Fields("Minstock") = "0"
                .Fields("Modify_Number") = "0"
                .Fields("F1") = "00"
                .Fields("F2") = "00"
                .Fields("F3") = "00"
                .Fields("F4") = "00"
                .Fields("F5") = "00"
                .Fields("Date_Created") = "2014/01/03"
                .Fields("Picture") = "-"
                .Fields("Print_On_Receipt") = "1"
                .Fields("Store_ID") = "01"
                .Update
            Next
        End If
    End With
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Update_Group"
End Sub

