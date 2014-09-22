VERSION 5.00
Begin VB.Form frmWordObject 
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12360
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
   ScaleHeight     =   8205
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Xem"
      Height          =   1095
      Left            =   8640
      TabIndex        =   2
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtword 
      Height          =   6015
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "luu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8640
      TabIndex        =   0
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "frmWordObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public ungdung As Word.Application
'Public tailieu As Word.Document
'
'Private Sub Command1_Click()
'    tailieu.Content.Text = txtword.Text
'End Sub
'
'Private Sub Command2_Click()
'    tailieu.PrintPreview
'    ungdung.Visible = True
'    ungdung.Activate
'End Sub
'
'Private Sub Form_Load()
'On Error GoTo Handle
'    Set ungdung = CreateObject("Word.application")
'    Set tailieu = ungdung.Documents.Open(WorkingFolder & "\HD.doc")
'    tailieu.MailMerge.Application.Activate
'
'Exit Sub
'Handle:
'MsgBox Err.Number & Err.Description & Me.Name & " Form_Load"
'End Sub
