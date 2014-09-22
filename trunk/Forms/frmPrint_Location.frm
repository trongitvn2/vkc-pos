VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmPrint_Location 
   Caption         =   "Lùa chän m¸y in theo khu vùc"
   ClientHeight    =   8760
   ClientLeft      =   120
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
   ScaleHeight     =   8760
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   -240
      Width           =   10935
      Begin VB.Frame fraLocation 
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Left            =   5280
         TabIndex        =   6
         Top             =   1080
         Width           =   5415
         Begin VB.Frame fraPrinter1 
            Caption         =   "KP 1"
            ForeColor       =   &H00FF0000&
            Height          =   1575
            Left            =   120
            TabIndex        =   15
            Top             =   1680
            Width           =   5175
            Begin VB.ComboBox cboPrinter1 
               Height          =   390
               Left            =   120
               TabIndex        =   17
               Text            =   "M¸y in order"
               Top             =   480
               Width           =   4935
            End
            Begin VB.CheckBox chkPrinter1 
               Caption         =   "§ang sö dông"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   495
               Left            =   120
               TabIndex        =   16
               Top             =   960
               Width           =   4455
            End
         End
         Begin VB.Frame fraPrinter2 
            Caption         =   "KP 2"
            ForeColor       =   &H00FF0000&
            Height          =   1575
            Left            =   120
            TabIndex        =   12
            Top             =   3360
            Width           =   5175
            Begin VB.CheckBox chkPrinter2 
               Caption         =   "§ang sö dông"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   495
               Left            =   120
               TabIndex        =   14
               Top             =   960
               Width           =   4455
            End
            Begin VB.ComboBox cboPrinter2 
               Height          =   390
               Left            =   120
               TabIndex        =   13
               Text            =   "M¸y in order"
               Top             =   480
               Width           =   4935
            End
         End
         Begin VB.Frame fraPrinter3 
            Caption         =   "KP 3"
            ForeColor       =   &H00FF0000&
            Height          =   1575
            Left            =   120
            TabIndex        =   9
            Top             =   5040
            Width           =   5175
            Begin VB.CheckBox chkPrinter3 
               Caption         =   "§ang sö dông"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   495
               Left            =   120
               TabIndex        =   11
               Top             =   960
               Width           =   4455
            End
            Begin VB.ComboBox cboPrinter3 
               Height          =   390
               Left            =   120
               TabIndex        =   10
               Text            =   "M¸y in order"
               Top             =   480
               Width           =   4935
            End
         End
         Begin VB.Frame fraReceipt 
            Caption         =   "M¸y in Bill - B¸o c¸o..."
            ForeColor       =   &H00FF0000&
            Height          =   1095
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   5175
            Begin VB.ComboBox cboReceipt 
               Height          =   390
               Left            =   120
               TabIndex        =   8
               Text            =   "M¸y in Bill - B¸o c¸o"
               Top             =   480
               Width           =   4935
            End
         End
      End
      Begin prjTouchScreen.MyButton cmdClose 
         Cancel          =   -1  'True
         Height          =   975
         Left            =   8040
         TabIndex        =   4
         Top             =   7920
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1720
         BTYPE           =   3
         TX              =   "§ãn&g"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPrint_Location.frx":0000
         PICN            =   "frmPrint_Location.frx":001C
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
         Height          =   975
         Left            =   5280
         TabIndex        =   3
         Top             =   7920
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1720
         BTYPE           =   3
         TX              =   "&L­u"
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
         COLTYPE         =   4
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmPrint_Location.frx":0656
         PICN            =   "frmPrint_Location.frx":0672
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin MSFlexGridLib.MSFlexGrid flgLocation 
         Height          =   7215
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   12726
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Caption         =   "Cµi ®Æt m¸y in theo khu vùc"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   1200
         TabIndex        =   5
         Top             =   480
         Width           =   8055
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   4695
      End
   End
End
Attribute VB_Name = "frmPrint_Location"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPrinter_Location As New ADODB.Recordset
Dim rsLocation As New ADODB.Recordset
Dim Location_ID As String

Private Sub cboPrinter1_Change()
On Error GoTo Handle
    With rsPrinter_Location
        .Find "Location_ID='" & Location_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Printer1_Name") = cboPrinter1.Text
            .Update
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cboPrinter1_Change"
End Sub

Private Sub cboPrinter1_Click()
Call cboPrinter1_Change
End Sub

Private Sub cboPrinter2_Change()
On Error GoTo Handle
    With rsPrinter_Location
        .Find "Location_ID='" & Location_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Printer2_Name") = cboPrinter2.Text
            .Update
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cboPrinter2_Change"
End Sub

Private Sub cboPrinter2_Click()
Call cboPrinter2_Change
End Sub

Private Sub cboPrinter3_Change()
On Error GoTo Handle
    With rsPrinter_Location
        .Find "Location_ID='" & Location_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Printer3_Name") = cboPrinter3.Text
            .Update
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cboPrinter3_Change"
End Sub

Private Sub cboPrinter3_Click()
    Call cboPrinter3_Change
End Sub

Private Sub cboReceipt_Change()
On Error GoTo Handle
    With rsPrinter_Location
        .Find "Location_ID='" & Location_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Receipt_Name") = cboReceipt.Text
            .Update
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Private Sub cboReceipt_Click()
Call cboReceipt_Change
End Sub

Private Sub chkPrinter1_Click()
On Error GoTo Handle
Dim chkValue As Boolean
    If chkPrinter1.Value = True Then
        'chkPrinter1.Value = False
        chkValue = Not chkPrinter1.Value
    Else
         'chkPrinter1.Value = True
        chkValue = chkPrinter1.Value
    End If
    With rsPrinter_Location
        .Find "Location_ID='" & Location_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Printer1_Used") = chkValue
            .Update
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " chkPrinter1_Click"
End Sub

Private Sub chkPrinter2_Click()
Dim chkValue As Boolean
    If chkPrinter2.Value = True Then
        'chkPrinter2.Value = False
        chkValue = Not chkPrinter2.Value
    Else
         'chkPrinter2.Value = True
        chkValue = chkPrinter2.Value
    End If
    With rsPrinter_Location
        .Find "Location_ID='" & Location_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Printer2_Used") = chkValue
            .Update
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " chkPrinter2_Click"
End Sub

Private Sub chkPrinter3_Click()
Dim chkValue As Boolean
    If chkPrinter3.Value = True Then
       ' chkPrinter3.Value = False
        chkValue = Not chkPrinter3.Value
    Else
        ' chkPrinter3.Value = True
        chkValue = chkPrinter3.Value
    End If
    With rsPrinter_Location
        .Find "Location_ID='" & Location_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Printer3_Used") = chkValue
            .Update
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " chkPrinter3_Click"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo Handle
    If MsgBox("B¹n cã muèn l­u cÊu h×nh hiÖn t¹i kh«ng?", vbYesNo) = vbYes Then
        Call Save_Data
    Else
        Exit Sub
    End If
    Set rsPrinter_Location = Nothing
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdSave_Click"
End Sub

Private Sub flgLocation_Click()
On Error GoTo Handle
    Location_ID = flgLocation.TextMatrix(flgLocation.Row, 0)
    fraLocation.Caption = flgLocation.TextMatrix(flgLocation.Row, 1)
    With rsPrinter_Location
        .Find "Location_ID='" & Location_ID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            cboReceipt.Text = .Fields("Receipt_Name")
            cboPrinter1.Text = .Fields("Printer1_Name")
            If .Fields("Printer1_Used").Value = True Then
                chkPrinter1.Value = 1
            Else
                chkPrinter1.Value = 0
            End If
            cboPrinter2.Text = .Fields("Printer2_Name")
            If .Fields("Printer2_Used").Value = True Then
                chkPrinter2.Value = 1
            Else
                chkPrinter2.Value = 0
            End If
            cboPrinter3.Text = .Fields("Printer3_Name")
            If .Fields("Printer3_Used").Value = True Then
                chkPrinter3.Value = 1
            Else
                chkPrinter3.Value = 0
            End If
        Else
            .addNew
            .Fields("Location_ID") = Location_ID
            .Fields("Receipt_Name") = ""
            .Fields("Printer1_Name") = ""
            .Fields("Printer1_Used") = 0
             .Fields("Printer2_Name") = ""
            .Fields("Printer2_Used") = 0
             .Fields("Printer3_Name") = ""
            .Fields("Printer3_Used") = 0
            .Update
            cboReceipt.Text = "Ch­a ®Þnh nghÜa"
            cboPrinter1.Text = "Ch­a ®Þnh nghÜa"
            chkPrinter1.Value = False
            cboPrinter2.Text = "Ch­a ®Þnh nghÜa"
            chkPrinter2.Value = False
            cboPrinter3.Text = "Ch­a ®Þnh nghÜa"
            chkPrinter3.Value = False
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - flgLocation_Click"
End Sub

Private Sub Form_Load()
On Error GoTo Handle
'    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    'KiÓm tra nÕu ch­a cã table Setup_Printer_Location th× t¹o míi
    If Not Check_Table_exist("Setup_Printer_Location") Then
        Call Create_Table_Printer_Location
        Exit Sub
    End If
    
    Set rsLocation = Open_Table(cnData, "Table_Diagram_sections")
    'Khëi t¹o m¸y in vµo c¸c combo printer
    Call Load_PrinterName
    'Khëi t¹o List Location
    Call Init_Location
    'Khëi t¹o Recoset May in ¶o
    Call Init_Printer_Location
    cboReceipt.ListIndex = 0
    cboPrinter1.ListIndex = 0
    cboPrinter2.ListIndex = 0
    cboPrinter3.ListIndex = 0
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - Form_Load"
End Sub

Public Sub Load_PrinterName()
On Error GoTo Handle
    Dim prt As Printer
    With cboPrinter1
        .Clear
        For Each prt In Printers
            .AddItem prt.DeviceName
        Next
    End With
    With cboPrinter2
        .Clear
        For Each prt In Printers
            .AddItem prt.DeviceName
        Next
    End With
    With cboPrinter3
        .Clear
        For Each prt In Printers
            .AddItem prt.DeviceName
        Next
    End With
    With cboReceipt
        .Clear
        For Each prt In Printers
            .AddItem prt.DeviceName
        Next
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""

End Sub

Public Sub Init_Location()
On Error GoTo Handle
Dim i As Integer
i = 1
 With flgLocation
        .Cols = 2
        .Rows = 5
        .ColWidth(0) = 1500
        .ColWidth(1) = 4500
        .TextMatrix(0, 0) = "M· KV"
        .ColAlignment(0) = 2
        .TextMatrix(0, 1) = "Tªn KV"
        .ColAlignment(1) = 2
    End With
        With rsLocation
            If .State <> 0 Then
                If .RecordCount = 0 Then
                    Exit Sub
                Else
                    .MoveFirst
                End If
            Else
                Exit Sub
            End If
            flgLocation.Rows = rsLocation.RecordCount + 1
            Do While Not .EOF
                flgLocation.TextMatrix(i, 0) = .Fields("Location_ID")
                flgLocation.TextMatrix(i, 1) = .Fields("Section_ID")
            .MoveNext
            i = i + 1
            Loop
        End With
       Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Init_Location"
End Sub

Public Sub Init_Printer_Location()
On Error GoTo Handle
Dim rsLocation_Print As New ADODB.Recordset
Set rsLocation_Print = Open_Table(cnData, "Setup_Printer_Location")
    With rsPrinter_Location
                If .State = 0 Then
                    .Fields.Append "Location_ID", adVarWChar, 2
                    .Fields.Append "Receipt_Name", adVarWChar, 100
                    .Fields.Append "Printer1_Name", adVarWChar, 100
                    .Fields.Append "Printer1_Used", adBoolean
                    .Fields.Append "Printer2_Name", adVarWChar, 100
                    .Fields.Append "Printer2_Used", adBoolean
                    .Fields.Append "Printer3_Name", adVarWChar, 100
                    .Fields.Append "Printer3_Used", adBoolean
                    .Open
                End If
                If rsLocation_Print.RecordCount > 0 Then
                    Do While Not rsLocation_Print.EOF
                        .addNew
                        .Fields("Location_ID") = rsLocation_Print.Fields("Location_ID")
                        .Fields("Receipt_Name") = rsLocation_Print.Fields("Receipt_Name")
                        .Fields("Printer1_Name") = rsLocation_Print.Fields("Printer1")
                        .Fields("Printer1_Used") = rsLocation_Print.Fields("Printer1_Used")
                        .Fields("Printer2_Name") = rsLocation_Print.Fields("Printer2")
                        .Fields("Printer2_Used") = rsLocation_Print.Fields("Printer2_Used")
                        .Fields("Printer3_Name") = rsLocation_Print.Fields("Printer3")
                        .Fields("Printer3_Used") = rsLocation_Print.Fields("Printer3_Used")
                        .Update
                    rsLocation_Print.MoveNext
                    Loop
                Else
                    Exit Sub
                End If
        End With
       Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Init_Location"
End Sub

Public Sub Save_Data()
On Error GoTo Handle
Dim rsLocation_Print As New ADODB.Recordset
    cnData.Execute "Delete  from Setup_Printer_Location"
    Set rsLocation_Print = Open_Table(cnData, "Setup_Printer_Location")
    With rsLocation_Print
        rsPrinter_Location.MoveFirst
        Do While Not rsPrinter_Location.EOF
            .addNew
            .Fields("Location_ID") = rsPrinter_Location.Fields("Location_ID")
            .Fields("Receipt_Name") = rsPrinter_Location.Fields("Receipt_Name")
            .Fields("Printer1") = rsPrinter_Location.Fields("Printer1_Name")
            .Fields("Printer1_Used") = rsPrinter_Location.Fields("Printer1_Used")
            .Fields("Printer2") = rsPrinter_Location.Fields("Printer2_Name")
            .Fields("Printer2_Used") = rsPrinter_Location.Fields("Printer2_Used")
            .Fields("Printer3") = rsPrinter_Location.Fields("Printer3_Name")
            .Fields("Printer3_Used") = rsPrinter_Location.Fields("Printer3_Used")
            .Update
        rsPrinter_Location.MoveNext
        Loop
    End With
    MsgBox "§· l­u thµnh c«ng!", vbInformation
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Save_Data"
End Sub
