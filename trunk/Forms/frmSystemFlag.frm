VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSystemFlag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cµi ®Æt Cê hÖ thèng"
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   10530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdClose 
      Cancel          =   -1  'True
      Height          =   945
      Left            =   5160
      TabIndex        =   23
      Tag             =   "L3"
      Top             =   8010
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   1667
      BTYPE           =   14
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
      BCOL            =   16578804
      BCOLO           =   16711680
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSystemFlag.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdSaveChange 
      Height          =   945
      Left            =   2850
      TabIndex        =   9
      Tag             =   "L2"
      Top             =   8010
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   1667
      BTYPE           =   14
      TX              =   "&Save change"
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
      BCOL            =   16578804
      BCOLO           =   16711680
      FCOL            =   16711680
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSystemFlag.frx":001C
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
      Height          =   7215
      Left            =   90
      TabIndex        =   0
      Top             =   750
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   12726
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   7
      TabHeight       =   882
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "SF1"
      TabPicture(0)   =   "frmSystemFlag.frx":0038
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "SF2 (M¸y in bÕp)"
      TabPicture(1)   =   "frmSystemFlag.frx":0054
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtValue(1)"
      Tab(1).Control(1)=   "Frame5"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "SF3 (In backup)"
      TabPicture(2)   =   "frmSystemFlag.frx":0070
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "SF4"
      TabPicture(3)   =   "frmSystemFlag.frx":008C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "SF5 "
      TabPicture(4)   =   "frmSystemFlag.frx":00A8
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame7(3)"
      Tab(4).Control(1)=   "Frame2(0)"
      Tab(4).ControlCount=   2
      TabCaption(5)   =   "SF6 (Receipt Options)"
      TabPicture(5)   =   "frmSystemFlag.frx":00C4
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame2(1)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "SF7"
      TabPicture(6)   =   "frmSystemFlag.frx":00E0
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame2(2)"
      Tab(6).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   6375
         Index           =   2
         Left            =   -74640
         TabIndex        =   28
         Top             =   720
         Width           =   9615
         Begin VB.TextBox txtValue 
            Alignment       =   2  'Center
            Height          =   495
            Index           =   6
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   31
            Top             =   360
            Width           =   1545
         End
         Begin VB.Frame Frame7 
            Height          =   5325
            Index           =   5
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   9375
            Begin VB.ListBox lstValue 
               Height          =   4860
               Index           =   6
               Left            =   60
               Style           =   1  'Checkbox
               TabIndex        =   30
               Top             =   240
               Width           =   9255
            End
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6375
         Index           =   1
         Left            =   -74700
         TabIndex        =   24
         Top             =   660
         Width           =   9615
         Begin VB.Frame Frame7 
            Height          =   5325
            Index           =   4
            Left            =   120
            TabIndex        =   26
            Top             =   960
            Width           =   9375
            Begin VB.ListBox lstValue 
               Height          =   4860
               Index           =   5
               Left            =   60
               Style           =   1  'Checkbox
               TabIndex        =   27
               Top             =   240
               Width           =   9255
            End
         End
         Begin VB.TextBox txtValue 
            Alignment       =   2  'Center
            Height          =   495
            Index           =   5
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   360
            Width           =   1545
         End
      End
      Begin VB.Frame Frame7 
         Height          =   5565
         Index           =   3
         Left            =   -74850
         TabIndex        =   21
         Top             =   1380
         Width           =   9735
         Begin VB.ListBox lstValue 
            Height          =   5160
            Index           =   4
            Left            =   60
            Style           =   1  'Checkbox
            TabIndex        =   22
            Top             =   240
            Width           =   9615
         End
      End
      Begin VB.TextBox txtValue 
         Alignment       =   2  'Center
         Height          =   495
         Index           =   1
         Left            =   -70680
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   840
         Width           =   1545
      End
      Begin VB.Frame Frame6 
         Height          =   6555
         Left            =   180
         TabIndex        =   5
         Top             =   540
         Width           =   9855
         Begin VB.Frame Frame7 
            Height          =   5535
            Index           =   0
            Left            =   60
            TabIndex        =   7
            Top             =   840
            Width           =   9735
            Begin VB.ListBox lstValue 
               Height          =   5160
               Index           =   0
               Left            =   60
               Style           =   1  'Checkbox
               TabIndex        =   10
               Top             =   240
               Width           =   9615
            End
         End
         Begin VB.TextBox txtValue 
            Alignment       =   2  'Center
            Height          =   495
            Index           =   0
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   330
            Width           =   1545
         End
      End
      Begin VB.Frame Frame5 
         Height          =   6615
         Left            =   -74940
         TabIndex        =   4
         Top             =   540
         Width           =   9855
         Begin VB.Frame Frame8 
            Height          =   5685
            Left            =   60
            TabIndex        =   15
            Top             =   720
            Width           =   9735
            Begin VB.ListBox lstValue 
               Height          =   5160
               Index           =   1
               Left            =   60
               Style           =   1  'Checkbox
               TabIndex        =   16
               Top             =   240
               Width           =   9615
            End
         End
      End
      Begin VB.Frame Frame4 
         Height          =   6615
         Left            =   -74940
         TabIndex        =   3
         Top             =   540
         Width           =   9855
         Begin VB.Frame Frame7 
            Height          =   5685
            Index           =   1
            Left            =   60
            TabIndex        =   17
            Top             =   840
            Width           =   9735
            Begin VB.ListBox lstValue 
               Height          =   5160
               Index           =   2
               Left            =   60
               Style           =   1  'Checkbox
               TabIndex        =   18
               Top             =   240
               Width           =   9615
            End
         End
         Begin VB.TextBox txtValue 
            Alignment       =   2  'Center
            Height          =   495
            Index           =   2
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   360
            Width           =   1545
         End
      End
      Begin VB.Frame Frame3 
         Height          =   6615
         Left            =   -74940
         TabIndex        =   2
         Top             =   540
         Width           =   9885
         Begin VB.Frame Frame7 
            Height          =   5685
            Index           =   2
            Left            =   60
            TabIndex        =   19
            Top             =   840
            Width           =   9735
            Begin VB.ListBox lstValue 
               Height          =   5160
               Index           =   3
               Left            =   60
               Style           =   1  'Checkbox
               TabIndex        =   20
               Top             =   240
               Width           =   9615
            End
         End
         Begin VB.TextBox txtValue 
            Alignment       =   2  'Center
            Height          =   495
            Index           =   3
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   13
            Top             =   360
            Width           =   1545
         End
      End
      Begin VB.Frame Frame2 
         Height          =   6495
         Index           =   0
         Left            =   -74910
         TabIndex        =   1
         Top             =   540
         Width           =   9855
         Begin VB.TextBox txtValue 
            Alignment       =   2  'Center
            Height          =   495
            Index           =   4
            Left            =   4320
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   360
            Width           =   1545
         End
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "cµi ®Æt cê hÖ thèng"
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
      Height          =   555
      Left            =   2640
      TabIndex        =   8
      Tag             =   "L1"
      Top             =   30
      Width           =   4425
   End
End
Attribute VB_Name = "frmSystemFlag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSystemFlag As New ADODB.Recordset
Dim rsPrint As New ADODB.Recordset

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSaveChange_Click()
Dim i As Integer
On Error GoTo Handle
    With rsSystemFlag
        For i = 1 To 7
            .Find "SF='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Fields("Data") = txtValue(i - 1).Text
                .Update
            Else
                .addNew
                .Fields("SF") = Format(i, "00")
                .Fields("Data") = txtValue(i - 1).Text
                .Update
            End If
        Next
    End With
    Call EnablePrint
    Call Load_SF_System
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  cmdSaveChange_Click"
End Sub

Private Sub Form_Activate()
On Error GoTo errHdl

    Dim DescArr() As String
    Dim ctrl As Control
    DescArr = LoadLanguage(LngFile, "#03:028:")
    If cmdClose.Font.name <> CurFont Then Call Set_Language(Me, CurFont)
    Me.Caption = DescArr(1)
    For Each ctrl In Me
    DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl

Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description

End Sub

Private Sub Form_Load()
On Error GoTo Handle
Dim i As Integer
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsSystemFlag = Open_Table(cnData, "SystemFlag")
    With rsSystemFlag
        .Find "SF=07", , adSearchForward, adBookmarkFirst
        If .EOF Then
            .addNew
            .Fields("SF") = "07"
            .Fields("Data") = "00"
            .Update
        End If
    End With
    If rsSystemFlag.RecordCount > 0 Then
        With rsSystemFlag
            For i = 1 To 7
                .Find "SF='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    txtValue(i - 1).Text = .Fields("Data")
                End If
            .MoveFirst
            Next
        End With
    End If
    Call Add_Flag_System
    For i = 0 To txtValue.count - 1 Step 1
        DoEvents
        AddValueForList txtValue(i).Text, lstValue(i)
    Next i
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub
Public Sub Add_Flag_System()
Dim i, j As Integer
On Error GoTo Handle
    Dim arrFlag() As String
    Dim iCount As Integer
    arrFlag = LoadLanguage(LngFile, "#02:007:")
    iCount = lstValue.count - 1
    For i = 0 To iCount
        lstValue(i).FontSize = 16
        Select Case i
            Case 0
                For j = 1 To 8
                    lstValue(i).AddItem arrFlag(j)
                Next j
            Case 1
                For j = 1 To 8
                    lstValue(i).AddItem arrFlag(j + 8)
                Next j
            Case 2
                For j = 1 To 8
                    lstValue(i).AddItem arrFlag(j + 16)
                Next j
            Case 3
                For j = 1 To 8
                    lstValue(i).AddItem arrFlag(j + 24)
                Next j
            Case 4
                For j = 1 To 8
                    lstValue(i).AddItem arrFlag(j + 32)
                Next j
            Case 5
                For j = 1 To 8
                    lstValue(i).AddItem arrFlag(j + 40)
                Next j
             Case 6
                For j = 1 To 8
                    lstValue(i).AddItem arrFlag(j + 48)
                Next j
        End Select
    Next i
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "Add_Flag_System"
End Sub
Private Sub LockTxtFlag()
Dim i As Integer
On Error GoTo errHdl

    For i = 0 To txtValue.count - 1 Step 1
    DoEvents
        txtValue(i).Locked = True
    Next i

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - LockTxtFlag "
End Sub
Private Sub lstValue_Click(Index As Integer)
Dim i As Integer
On Error GoTo errHdl

    Dim strflag As String
    'danh dau 1 co SF6
'    Select Case Index
'        Case 5
'            For i = 0 To lstValue(Index).ListCount - 1
'            DoEvents
'                 lstValue(Index).Selected(i) = False
'            Next i
'
'    End Select
'    If fLoad Then ' event is called directly by clicking on list, not call by another functions or subs
        strflag = ""
        For i = 0 To lstValue(Index).ListCount - 1
        DoEvents
            If lstValue(Index).Selected(i) Then
                strflag = strflag & "1"
            Else: strflag = strflag & "0"
            End If
        Next i
        txtValue(Index).Text = FillZeroForString(BinToHex(strflag), 2)
'    End If
    

Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
        & Me.name & " - lstFlag_Click "
End Sub

Public Sub EnablePrint()
On Error GoTo Handle
Dim i As Integer
Set rsPrint = Open_Table(cnData, "Printer_Mapping")
    With rsPrint
        For i = 1 To 8
        DoEvents
            .Find "PrinterName='" & Format(i, "00") & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Fields("Disabled") = Mid(Right("00000000" & HexToBin(txtValue(1)), 8), i, 1)
                .Update
            End If
        Next
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " EnablePrint "
End Sub
