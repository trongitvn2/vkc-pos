VERSION 5.00
Begin VB.Form frmRangeTable 
   Caption         =   "Thªm bµn míi"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   ClipControls    =   0   'False
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
   ScaleHeight     =   5175
   ScaleWidth      =   5880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraSapxep 
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Visible         =   0   'False
      Width           =   5655
      Begin VB.CheckBox chkAutoRange 
         Caption         =   "S¾p xÕp tù ®éng theo ch­¬ng tr×nh"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
      Begin VB.Frame fraSize 
         Caption         =   "Réng...........................Cao.....................Cì ch÷.............."
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Visible         =   0   'False
         Width           =   5415
         Begin VB.ComboBox cbofont 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   390
            ItemData        =   "frmRangeTable.frx":0000
            Left            =   3840
            List            =   "frmRangeTable.frx":0019
            TabIndex        =   15
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtHeight 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   1920
            TabIndex        =   14
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtwidth 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "BiÓu t­îng"
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
         TabIndex        =   8
         Top             =   1440
         Visible         =   0   'False
         Width           =   2295
         Begin VB.TextBox tblSymbol 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame fraRange 
         Caption         =   "Tõ sè                      §Õn sè"
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
         Left            =   2520
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   3015
         Begin VB.TextBox txtTo 
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
            Height          =   495
            Left            =   1800
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox txtFrom 
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
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   975
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000000&
            Visible         =   0   'False
            X1              =   720
            X2              =   2040
            Y1              =   120
            Y2              =   120
         End
      End
      Begin VB.OptionButton OptMulti 
         Caption         =   "Mét d·y bµn"
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
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   4095
      End
      Begin VB.OptionButton OptSingle 
         Caption         =   "1 Bµn"
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
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   3735
      End
   End
   Begin prjTouchScreen.MyButton cmdOK 
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   4200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "&§ång ý"
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
      BCOL            =   14215660
      BCOLO           =   16777215
      FCOL            =   16711680
      FCOLO           =   33023
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRangeTable.frx":0039
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdCancel 
      Height          =   855
      Left            =   3000
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      BTYPE           =   14
      TX              =   "&Tho¸t"
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
      BCOL            =   14215660
      BCOLO           =   16777215
      FCOL            =   16711680
      FCOLO           =   33023
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmRangeTable.frx":0055
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
Attribute VB_Name = "frmRangeTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Location As String
Dim rsTable As New ADODB.Recordset
Dim Width_Layout As Integer
Dim Height_Layout As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
On Error GoTo Handle
    If OptSingle.Value = True Then
        Unload Me
        With frmKeyboard
            .FormCallkeyboard = "Add_Table"
            .lblTitle.Caption = "Enter_Table_Name"
            .lblTableType.Visible = True
            .cboTbaleType.Visible = True
            .txtInput.PasswordChar = ""
            .Show vbModal
        End With
    ElseIf OptMulti.Value = True Then
        If txtwidth >= 1600 And txtHeight >= 1000 Then
            If CDbl("0" & txtTo.Text) - CDbl("0" & txtFrom.Text) > 56 Then
                MsgBox "Víi kÝch th­íc nµy, b¹n kh«ng thÓ thªm ®­îc "
            Else
                If chkAutoRange.Value = 1 Then
                    'Call AddTableRange(CInt(txtFrom.Text), CInt(txtTo.Text))
                    Call Auto_Range(txtwidth.Text, txtHeight.Text, CInt(txtTo.Text), CInt(txtFrom.Text))
                Else
                    Call AddTable(CInt(txtFrom.Text), CInt(txtTo.Text))
                End If
            End If
        Else
            If chkAutoRange.Value = 1 Then
'                Call AddTableRange(CInt(txtFrom.Text), CInt(txtTo.Text))
                Call Auto_Range(txtwidth.Text, txtHeight.Text, CInt(txtTo.Text), CInt(txtFrom.Text))
            Else
                Call AddTable(CInt(txtFrom.Text), CInt(txtTo.Text))
            End If
        End If
        Unload Me
    End If
    
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdOK_Click"

End Sub

Private Sub Form_Load()
On Error GoTo Handle
    OptSingle.Value = True
    txtwidth.Text = 1300
    txtHeight.Text = 900
    cbofont.ListIndex = 0
    Set rsTable = OpenCriticalTable("select * from Table_Diagram where Section_ID='" & Sec_ID & "'", cnData)
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Load"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rsTable = Nothing
End Sub

Private Sub OptMulti_Click()
On Error GoTo Handle
    fraRange.Visible = True
    Frame2.Visible = True
    Line1.Visible = True
    tblSymbol.SetFocus
    fraSapxep.Visible = True
    fraSize.Visible = True
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  OptSingle_Click"
End Sub

Private Sub OptSingle_Click()
On Error GoTo Handle
    fraRange.Visible = False
    Line1.Visible = False
    'OptSingle.Value = False
    'OptMulti.Value = True
    txtFrom.TabIndex = 1
    fraSapxep.Visible = False
    Frame2.Visible = False
    fraSize.Visible = False
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  OptSingle_Click"
End Sub





Private Sub tblSymbol_DblClick()
With frmKeyboard
        With frmKeyboard.txtInput
            .PasswordChar = ""
            .Text = tblSymbol.Text
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        .FormCallkeyboard = "Other"
        .Show vbModal
        tblSymbol.Text = .Let_Text_Input
       
    End With
   
End Sub

Private Sub tblSymbol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFrom.SetFocus
    End If
End Sub

Private Sub txtFrom_DblClick()
    With frmKeyboard
        With frmKeyboard.txtInput
        
        .PasswordChar = ""
        .Text = txtFrom.Text
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtFrom.Text = .Let_Text_Input
       
    End With
End Sub

Private Sub txtFrom_GotFocus()
     tblSymbol.SelStart = 0
     tblSymbol.SelLength = Len(tblSymbol.Text)
End Sub

Private Sub txtFrom_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii = 13 Then txtTo.SetFocus
    If KeyAscii < 32 Then Exit Sub
    Select Case KeyAscii
        Case 48 To 57, 46
        Case Else:   KeyAscii = 0
    End Select
    
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtFrom_KeyPress"
End Sub



Private Sub txtTo_DblClick()
With frmKeyboard
        With frmKeyboard.txtInput
            .PasswordChar = ""
            .Text = txtTo.Text
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtTo.Text = .Let_Text_Input
       
    End With
   
End Sub

Private Sub txtTo_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
     If KeyAscii = 13 Then cmdOk_Click
    If KeyAscii < 32 Then Exit Sub
    Select Case KeyAscii
        Case 48 To 57, 46
        Case Else:   KeyAscii = 0
    End Select
   
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtTo_KeyPress"
End Sub



Public Property Let Get_Location(ByVal vNewValue As Variant)
    Location = vNewValue
End Property

Public Sub AddTable(FromValue As Integer, Tovalue As Integer)
On Error GoTo Handle
    Dim i As Integer
    
    For i = FromValue To Tovalue
        With rsTable
            rsTable.Find "Table_Number='" & Trim(tblSymbol.Text) & i & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                rsTable.addNew
                rsTable!Store_ID = Store_ID
                rsTable!Section_ID = Sec_ID
                rsTable!Table_Number = Trim(tblSymbol.Text) & i
                
                rsTable!XPos = 1000
                rsTable!YPos = 1000
                rsTable!Height = txtHeight.Text
                rsTable!Width = txtwidth.Text
                rsTable!Cost_Center_Index = cbofont.Text
                rsTable!NumSeats = 1
                rsTable!ShapeType = 0
                rsTable.Update
                rsTable.Requery
            End If
        End With
    Next
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " AddTable"

End Sub

Public Sub AddTableRange(FromValue As Integer, Tovalue As Integer)
On Error GoTo Handle
    Dim i, j As Integer
    Dim w, h As Integer
    For i = FromValue To Tovalue
        With rsTable
            rsTable.Find "Table_Number='" & Trim(tblSymbol.Text) & i & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                rsTable.addNew
                rsTable!Store_ID = Store_ID
                rsTable!Section_ID = Sec_ID
                rsTable!Table_Number = Trim(tblSymbol.Text) & i
                j = j + 1
                Select Case j
                    Case 1 To 9
                        rsTable!YPos = 100
                        rsTable!XPos = 1430 * (j - 1) + 50
                    Case 10 To 18
                        rsTable!YPos = 1100
                        rsTable!XPos = 1430 * (j - 10) + 50
                    Case 19 To 27
                        rsTable!YPos = 2200
                        rsTable!XPos = 1430 * (j - 19) + 50
                    Case 28 To 36
                        rsTable!YPos = 3300
                        rsTable!XPos = 1430 * (j - 28) + 50
                    Case 37 To 46
                        rsTable!YPos = 4400
                        rsTable!XPos = 1430 * (j - 37) + 50
                    Case 47 To 56
                        rsTable!YPos = 5500
                        rsTable!XPos = 1430 * (j - 47) + 50
                    Case 57 To 66
                        rsTable!YPos = 6600
                        rsTable!XPos = 1430 * (j - 57) + 50
                    Case 67 To 76
                        rsTable!YPos = 7700
                        rsTable!XPos = 1430 * (j - 67) + 50
                    Case 77 To 86
                        rsTable!YPos = 8800
                        rsTable!XPos = 1430 * (j - 77) + 50
                    Case 78 To 87
                        rsTable!YPos = 9900
                        rsTable!XPos = 1430 * (j - 78) + 50
                End Select
                rsTable!Height = txtHeight.Text
                rsTable!Width = txtwidth.Text
                rsTable!Cost_Center_Index = cbofont.Text
                rsTable!NumSeats = 1
                rsTable!ShapeType = 0
                rsTable.Update
                rsTable.Requery
            End If
        End With
    Next
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " AddTable"

End Sub



Public Property Let Get_Width(ByVal vNewValue As Variant)
    Width_Layout = vNewValue
End Property


Public Property Let Get_Height(ByVal vNewValue As Variant)
    Height_Layout = vNewValue
End Property

Public Sub Auto_Range(Tablewidth As Integer, TableHeight As Integer, NumofTable As Integer, Start_num As Integer)
    On Error GoTo Handle
        Dim rows, cols As Integer
        Dim i, j As Integer
        cols = Int((Width_Layout - 500) / Tablewidth)
        rows = Int(NumofTable / cols) + 1
        Dim cap As Integer
        For i = 1 To rows
            For j = 0 To cols - 1
            cap = i * cols - cols + j + Start_num
                If cap > NumofTable Then Exit Sub
                With rsTable
                    rsTable.Find "Table_Number='" & Trim(tblSymbol.Text) & cap & "'", , adSearchForward, adBookmarkFirst
                    If .EOF Then
                        rsTable.addNew
                        rsTable!Store_ID = Store_ID
                        rsTable!Section_ID = Sec_ID
                        rsTable!Table_Number = Trim(tblSymbol.Text) & cap
                        If i = 1 Then
                            rsTable!YPos = i * TableHeight - TableHeight
                        Else
                            rsTable!YPos = i * TableHeight - TableHeight + i * 50
                        End If
                        If j = 0 Then
                            rsTable!XPos = (j + 1) * Tablewidth - Tablewidth
                        Else
                            rsTable!XPos = (j + 1) * Tablewidth - Tablewidth + (j + 1) * 50
                        End If
                        
                        rsTable!Height = TableHeight
                        rsTable!Width = Tablewidth
                        rsTable!Cost_Center_Index = cbofont.Text
                        rsTable!NumSeats = 1
                        rsTable!ShapeType = 0
                        rsTable.Update
                        rsTable.Requery
                    End If
                End With
            Next j
        Next i
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Auto_Range"
End Sub
