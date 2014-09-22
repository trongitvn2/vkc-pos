VERSION 5.00
Begin VB.Form frmMessage 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1320
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5520
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
   ScaleHeight     =   1320
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4740
      Top             =   1080
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   150
      TabIndex        =   0
      Tag             =   "L1"
      Top             =   240
      Width           =   5205
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Arrdesc() As String
Dim i As Integer

Private Sub Form_Activate()
On Error GoTo Handle
    Dim ctrl As Control
        i = 0
        Arrdesc = LoadLanguage(LngFile, "#02:008:")
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Arrdesc(Mid(ctrl.Tag, 2))
        Next ctrl
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.Name & "  Form_Load"
End Sub

Private Sub Timer1_Timer()
    i = i + 1
    If i = 2 Then Unload Me
End Sub
