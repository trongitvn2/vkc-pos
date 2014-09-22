VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmKeyboard 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bµn ph›m c∂m ¯ng"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13725
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
   Icon            =   "frmKeyboard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   13725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjTouchScreen.MyButton cmdUp 
      Height          =   855
      Left            =   12120
      TabIndex        =   53
      Top             =   4080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   ""
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
      BCOLO           =   12632256
      FCOL            =   12632256
      FCOLO           =   12632256
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmKeyboard.frx":000C
      PICN            =   "frmKeyboard.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   3300
      Left            =   0
      ScaleHeight     =   3240
      ScaleWidth      =   11235
      TabIndex        =   4
      Top             =   2280
      Width           =   11295
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   19
         Left            =   10125
         TabIndex        =   49
         Top             =   0
         Width           =   1095
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "p"
         Size            =   "1931;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   38
         Left            =   10920
         TabIndex        =   48
         Top             =   1560
         Visible         =   0   'False
         Width           =   1020
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "."
         Size            =   "1799;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   36
         Left            =   10440
         TabIndex        =   47
         Top             =   1680
         Visible         =   0   'False
         Width           =   1050
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   ","
         Size            =   "1852;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   35
         Left            =   8040
         TabIndex        =   46
         Top             =   1620
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "m"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   34
         Left            =   6960
         TabIndex        =   45
         Top             =   1620
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "n"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   33
         Left            =   5880
         TabIndex        =   44
         Top             =   1620
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "b"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   32
         Left            =   4800
         TabIndex        =   43
         Top             =   1620
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "v"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdEnter 
         Height          =   1610
         Index           =   48
         Left            =   9135
         TabIndex        =   42
         Top             =   1620
         Width           =   2085
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "Enter"
         PicturePosition =   196613
         Size            =   "3678;2840"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   315
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   31
         Left            =   3720
         TabIndex        =   41
         Top             =   1620
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "c"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   41
         Left            =   2280
         TabIndex        =   40
         Top             =   2430
         Width           =   6820
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "Space"
         Size            =   "12030;1402"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   30
         Left            =   2640
         TabIndex        =   39
         Top             =   1620
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "x"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCancel 
         Height          =   795
         Index           =   42
         Left            =   1080
         TabIndex        =   38
         Top             =   2430
         Width           =   1180
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "Cancel"
         Size            =   "2081;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   29
         Left            =   1560
         TabIndex        =   37
         Top             =   1620
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "z"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   39
         Left            =   0
         TabIndex        =   36
         Top             =   2430
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "Ctrl"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   51
         Left            =   0
         TabIndex        =   35
         Top             =   1620
         Width           =   1545
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "Shift"
         Size            =   "2725;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   28
         Left            =   9960
         TabIndex        =   34
         Top             =   810
         Width           =   1260
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "l"
         Size            =   "2222;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   27
         Left            =   8880
         TabIndex        =   33
         Top             =   810
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "k"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   26
         Left            =   7800
         TabIndex        =   32
         Top             =   810
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "j"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   25
         Left            =   6720
         TabIndex        =   31
         Top             =   810
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "h"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   24
         Left            =   5640
         TabIndex        =   30
         Top             =   810
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "g"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   23
         Left            =   4560
         TabIndex        =   29
         Top             =   810
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "f"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   22
         Left            =   3480
         TabIndex        =   28
         Top             =   810
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "d"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   21
         Left            =   2400
         TabIndex        =   27
         Top             =   810
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "s"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   20
         Left            =   1320
         TabIndex        =   26
         Top             =   810
         Width           =   1070
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "a"
         Size            =   "1887;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   50
         Left            =   0
         TabIndex        =   25
         Top             =   810
         Width           =   1310
         ForeColor       =   16711680
         BackColor       =   12632256
         VariousPropertyBits=   8388635
         Caption         =   "Caps Lock"
         Size            =   "2311;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   18
         Left            =   9000
         TabIndex        =   24
         Top             =   0
         Width           =   1110
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "o"
         Size            =   "1958;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   17
         Left            =   8010
         TabIndex        =   23
         Top             =   0
         Width           =   975
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "i"
         Size            =   "1720;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   16
         Left            =   7020
         TabIndex        =   22
         Top             =   0
         Width           =   975
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "u"
         Size            =   "1720;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   15
         Left            =   6030
         TabIndex        =   21
         Top             =   0
         Width           =   975
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "y"
         Size            =   "1720;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   14
         Left            =   5040
         TabIndex        =   20
         Top             =   0
         Width           =   975
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "t"
         Size            =   "1720;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   13
         Left            =   4035
         TabIndex        =   19
         Top             =   0
         Width           =   990
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "r"
         Size            =   "1746;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   12
         Left            =   3045
         TabIndex        =   18
         Top             =   0
         Width           =   975
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "e"
         Size            =   "1720;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   11
         Left            =   2040
         TabIndex        =   17
         Top             =   0
         Width           =   990
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "w"
         Size            =   "1746;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   10
         Left            =   1035
         TabIndex        =   16
         Top             =   0
         Width           =   990
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "q"
         Size            =   "1746;1402"
         FontName        =   ".VnArial"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdText 
         Height          =   795
         Index           =   37
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1020
         ForeColor       =   16711680
         BackColor       =   12632256
         Caption         =   "#"
         Size            =   "1799;1402"
         FontName        =   ".VnArialH"
         FontEffects     =   1073741825
         FontHeight      =   360
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   9255
   End
   Begin VB.ComboBox cboTbaleType 
      Height          =   390
      Left            =   8700
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   2535
   End
   Begin prjTouchScreen.MyButton cmdRight 
      Height          =   855
      Left            =   12840
      TabIndex        =   54
      Top             =   4920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   ""
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
      BCOLO           =   12632256
      FCOL            =   12632256
      FCOLO           =   12632256
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmKeyboard.frx":01B7
      PICN            =   "frmKeyboard.frx":01D3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmdLeft 
      Height          =   855
      Left            =   11400
      TabIndex        =   55
      Top             =   4920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   ""
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
      BCOLO           =   12632256
      FCOL            =   12632256
      FCOLO           =   12632256
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmKeyboard.frx":0361
      PICN            =   "frmKeyboard.frx":037D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin prjTouchScreen.MyButton cmddown 
      Height          =   855
      Left            =   12120
      TabIndex        =   56
      Top             =   4920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
      BTYPE           =   3
      TX              =   ""
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
      BCOLO           =   12632256
      FCOL            =   12632256
      FCOLO           =   12632256
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmKeyboard.frx":04DD
      PICN            =   "frmKeyboard.frx":04F9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   4
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      Value           =   0   'False
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   42
      Left            =   11280
      TabIndex        =   68
      Top             =   600
      Width           =   795
      ForeColor       =   16711680
      BackColor       =   12632256
      Caption         =   "1"
      Size            =   "1411;1411"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   43
      Left            =   12090
      TabIndex        =   67
      Top             =   600
      Width           =   795
      ForeColor       =   16711680
      BackColor       =   12632256
      Caption         =   "2"
      Size            =   "1411;1411"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   44
      Left            =   12885
      TabIndex        =   66
      Top             =   600
      Width           =   795
      ForeColor       =   16711680
      BackColor       =   12632256
      Caption         =   "3"
      Size            =   "1411;1411"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   45
      Left            =   11280
      TabIndex        =   65
      Top             =   1440
      Width           =   795
      ForeColor       =   16711680
      BackColor       =   12632256
      Caption         =   "4"
      Size            =   "1402;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   46
      Left            =   12090
      TabIndex        =   64
      Top             =   1440
      Width           =   795
      ForeColor       =   16711680
      BackColor       =   12632256
      Caption         =   "5"
      Size            =   "1402;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   47
      Left            =   12885
      TabIndex        =   63
      Top             =   1440
      Width           =   795
      ForeColor       =   16711680
      BackColor       =   12632256
      Caption         =   "6"
      Size            =   "1402;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   48
      Left            =   11280
      TabIndex        =   62
      Top             =   2280
      Width           =   795
      ForeColor       =   16711680
      BackColor       =   12632256
      Caption         =   "7"
      Size            =   "1402;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   49
      Left            =   12090
      TabIndex        =   61
      Top             =   2280
      Width           =   795
      ForeColor       =   16711680
      BackColor       =   12632256
      Caption         =   "8"
      Size            =   "1402;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   52
      Left            =   12885
      TabIndex        =   60
      Top             =   2280
      Width           =   795
      ForeColor       =   16711680
      BackColor       =   12632256
      Caption         =   "9"
      Size            =   "1402;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   53
      Left            =   11280
      TabIndex        =   59
      Top             =   3120
      Width           =   795
      ForeColor       =   16711680
      BackColor       =   12632256
      Caption         =   "0"
      Size            =   "1402;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   54
      Left            =   12090
      TabIndex        =   58
      Top             =   3120
      Width           =   795
      ForeColor       =   16711680
      BackColor       =   12632256
      Caption         =   ","
      Size            =   "1402;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   55
      Left            =   12885
      TabIndex        =   57
      Top             =   3120
      Width           =   795
      ForeColor       =   16711680
      BackColor       =   12632256
      Caption         =   "."
      Size            =   "1402;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   360
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "HÁ trÓ g‚ d u ti’ng vi÷t ki”u VNI"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   52
      Top             =   5595
      Width           =   11295
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   40
      Left            =   9405
      TabIndex        =   51
      Top             =   600
      Width           =   1845
      ForeColor       =   16777215
      BackColor       =   65280
      Caption         =   "Clear"
      Size            =   "3254;1402"
      FontName        =   ".VnArialH"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdBackSpace 
      Height          =   795
      Index           =   43
      Left            =   10080
      TabIndex        =   50
      Top             =   1455
      Width           =   1170
      ForeColor       =   12582912
      BackColor       =   8438015
      Caption         =   "Backs"
      Size            =   "2064;1402"
      FontName        =   ".VnArial"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   9
      Left            =   9060
      TabIndex        =   14
      Top             =   1455
      Width           =   1005
      BackColor       =   8438015
      Caption         =   "0"
      Size            =   "1773;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   8
      Left            =   8040
      TabIndex        =   13
      Top             =   1455
      Width           =   1005
      BackColor       =   8438015
      Caption         =   "9"
      Size            =   "1773;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   7
      Left            =   7020
      TabIndex        =   12
      Top             =   1455
      Width           =   1005
      BackColor       =   8438015
      Caption         =   "8"
      Size            =   "1773;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   6
      Left            =   6015
      TabIndex        =   11
      Top             =   1455
      Width           =   1005
      BackColor       =   8438015
      Caption         =   "7"
      Size            =   "1773;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   5
      Left            =   5010
      TabIndex        =   10
      Top             =   1455
      Width           =   1005
      BackColor       =   8438015
      Caption         =   "6"
      Size            =   "1773;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   4
      Left            =   4005
      TabIndex        =   9
      Top             =   1455
      Width           =   1005
      BackColor       =   8438015
      Caption         =   "5"
      Size            =   "1773;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   3
      Left            =   3000
      TabIndex        =   8
      Top             =   1455
      Width           =   1005
      BackColor       =   8438015
      Caption         =   "4"
      Size            =   "1773;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   2
      Left            =   1995
      TabIndex        =   7
      Top             =   1455
      Width           =   1005
      BackColor       =   8438015
      Caption         =   "3"
      Size            =   "1773;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   1
      Left            =   1005
      TabIndex        =   6
      Top             =   1455
      Width           =   1005
      BackColor       =   8438015
      Caption         =   "2"
      Size            =   "1773;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdText 
      Height          =   795
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   1455
      Width           =   1005
      BackColor       =   8438015
      Caption         =   "1"
      Size            =   "1773;1402"
      FontName        =   ".VnArial"
      FontEffects     =   1073741825
      FontHeight      =   315
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label lblTableType 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Loπi bµn:"
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   7110
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter you values:"
      BeginProperty Font 
         Name            =   ".VnArial"
         Size            =   15.75
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
      Top             =   90
      Width           =   6465
   End
End
Attribute VB_Name = "frmKeyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsTable As New ADODB.Recordset
Dim Text_Field As String
Dim FlagShift, FlagCaplock As Boolean
Dim Keyboard As String
Dim text_Input, Table_ID As String
Dim State_Type As Boolean

Private Sub cmdBackSpace_Click(Index As Integer)
On Error GoTo Handle
If txtInput.SelStart > 0 Then
    Dim txt1, txt2, txt As String
    Dim selpos As Integer
    txt1 = Left(txtInput.Text, txtInput.SelStart)
    selpos = txtInput.SelStart
    txt2 = Right(txtInput.Text, Len(txtInput.Text) - txtInput.SelStart)
    txt = Left(txt1, Len(txt1) - 1) & txt2
    txtInput.Text = txt
    txtInput.SetFocus
    txtInput.SelStart = selpos - 1

End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdEnter_Click(Index As Integer)
On Error GoTo Handle
    
    Select Case Keyboard
        
        Case "EditTable":
    '    If UserLevel = 1 Then
            If UCase(txtInput.Text) = UCase(UserID) & UCase(UserPass) Or UCase(txtInput.Text) = "131112" Then
                Unload Me
                frmEditTablePlan.Show vbModal
            Else
                MsgBox "Bπn nhÀp sai mÀt kh»u ho∆c bπn kh´ng c„ quy“n qu∂n trﬁ,li™n h÷ qu∂n l˝", vbOKOnly
            End If
    '    End If
        Case "Add_Section":
            Dim rsSection As New ADODB.Recordset
            Dim rsmax As New ADODB.Recordset
            'If Dir(WorkingFolder & "\Database.mdb", vbDirectory) <> "" Then
                Set rsSection = OpenCriticalTable("select Store_ID,Location_ID,Section_ID,PriceRate,VAT,Price_Level ,Service_Charge,TimeLevel,isTimer from Table_Diagram_Sections ", cnData)
                Set rsmax = OpenCriticalTable("select Max(Location_ID) as MaxID from Table_Diagram_Sections", cnData)
                With rsSection
                    .addNew
                    .Fields("Store_ID") = Store_ID
                    .Fields("Location_ID") = Format(CDbl("0" & rsmax.Fields("maxID")) + 1, "00")
                    .Fields("Section_ID") = txtInput.Text
                    .Fields("PriceRate") = "0"
                    .Fields("VAT") = 0
                    .Fields("Price_Level") = 1
                    .Fields("Service_Charge") = 0
                    .Fields("TimeLevel") = 0
                    .Fields("isTimer") = False
                    .Update
                End With
            'End If
            
        Case "Add_Table"
            Set rsTable = OpenCriticalTable("select * from Table_Diagram where Section_ID='" & Sec_ID & "'", cnData)
                rsTable.Find "Table_Number='" & Trim(txtInput.Text) & "'", , adSearchForward, adBookmarkFirst
                If Not rsTable.EOF Then
                    MsgBox " ß∑ tÂn tπi sË bµn nµy trong khu !", vbInformation
                    Exit Sub
                Else
                    rsTable.addNew
                    rsTable!Store_ID = Store_ID
                    rsTable!Section_ID = Sec_ID
                    rsTable!Table_Number = Trim(txtInput.Text)
                    rsTable!XPos = 1000
                    rsTable!YPos = 1000
                    rsTable!Height = 1000
                    rsTable!Width = 1400
                    rsTable!Cost_Center_Index = 14
                    rsTable!NumSeats = 1
                    Select Case cboTbaleType.ListIndex
                        Case 0: rsTable!ShapeType = 0
                        Case 1: rsTable!ShapeType = 5
                        Case 2: rsTable!ShapeType = 2
                    End Select
                    rsTable.Update
                    rsTable.Requery
                End If
        Case "DeleteSale"
            Me.Caption = "NhÀp mÀt kh»u"
            If UCase(txtInput.Text) = "131112" Then
                Unload Me
                frmDeleteSaleData.Show vbModal
            End If
           Case "Select_Station"
            Me.Caption = "NhÀp mÀt kh»u"
            If UCase(txtInput.Text) = "131112" Then
                Unload Me
                frmSelect_Station.Show vbModal
            End If
            
        Case "TakeOut"
            If txtInput.Text = "" Then
                MsgBox "Bπn ph∂i nhÀp th´ng tin ng≠Íi mua hµng, ho∆c sË bµn"
            Else
                Call Update_TakeOut(Trim(txtInput.Text))
            End If
        
        Case "SetColor"
            If UCase(txtInput.Text) = UCase(UserID) & UCase(UserPass) Or UCase(txtInput.Text) = "131112" Then
                Me.Caption = "NhÀp mÀt kh»u"
                Unload Me
                frmColorBox.Show vbModal
            Else
                MsgBox "NhÀp sai MÀt kh»u !", vbInformation
            End If
            
        Case "SystemFlag"
            If UCase(txtInput.Text) = UCase(UserID) & UCase(UserPass) Or UCase(txtInput.Text) = "131112" Or UCase(txtInput.Text) = "881507" Then
                Unload Me
                frmSystemFlag.Show vbModal
            Else
                MsgBox "NhÀp sai MÀt kh»u !", vbInformation
            End If
        Case "Employee"
            If UCase(txtInput.Text) = "131112" Then
                Unload Me
                frmemployee.Show vbModal
            Else
                MsgBox "Sai mÀt kh»u !"
            End If
        Case "Other"
            text_Input = txtInput.Text
        Case "AddPrint"
            Call AddPrint
        Case "EditName"
            text_Input = txtInput.Text
        Case "SaleReport"
           ' Set cnData = Get_Connection(BackupFolder & "\Database.mdb", "100881administrator")
            If UCase(txtInput.Text) = "131112" Then
                With frmSetup
                    .Show vbModal
                End With
            Else
                MsgBox "Sai mÀt kh»u Æ®ng nhÀp !"
            End If
        Case Else:
        
    End Select

    Unload Me
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "   cmdEnter_Click"
End Sub

Private Sub cmdLeft_Click()
    If txtInput.SelStart <> 0 Then
        txtInput.SelStart = txtInput.SelStart - 1
    End If
    txtInput.SetFocus
End Sub

Private Sub cmdRight_Click()
    If txtInput.SelStart <> 0 Then
        txtInput.SelStart = txtInput.SelStart + 1
    End If
    txtInput.SetFocus
End Sub

Private Sub cmdText_Click(Index As Integer)
On Error GoTo Handle
Dim strSpecial As String
strSpecial = "!@$%&*()-+"
Dim i As Integer
    Select Case Index
        Case 50:
            FlagCaplock = Not FlagCaplock
            If FlagCaplock = True Then
                For i = 10 To 38
                    cmdText(i).Caption = UCase(cmdText(i).Caption)
                Next
            Else
                For i = 10 To 38
                    cmdText(i).Caption = LCase(cmdText(i).Caption)
                Next
            End If
            'FlagCaplock = True
        Case 51:
            FlagShift = True
            For i = 10 To 38
                cmdText(i).Caption = UCase(cmdText(i).Caption)
            Next
            For i = 1 To 10
                cmdText(i - 1).Caption = Mid(strSpecial, i, 1)
            Next i
        Case 40:
            txtInput.Text = ""
        Case 41:
            Dim posSel As Integer
            Dim txt1, txt2, txt, alpha As String
            posSel = txtInput.SelStart
            alpha = Space(1)
            txt1 = Left(txtInput.Text, posSel)
            txt2 = Right(txtInput.Text, Len(txtInput.Text) - posSel)
            txt = txt1 & alpha & txt2
            
            txtInput.Text = txt
            txtInput.SetFocus
            txtInput.SelStart = posSel + 1
            
        Case Else:
            'txtInput.text = txtInput.text & cmdText(Index).Caption
            
            posSel = txtInput.SelStart
            alpha = cmdText(Index).Caption
            txt1 = Left(txtInput.Text, posSel)
            txt2 = Right(txtInput.Text, Len(txtInput.Text) - posSel)
           
            If posSel = 0 Then posSel = 1
            Select Case alpha
                Case "1"
                    If Mid(txt1, posSel, 1) = "a" Then
                        txt = Left(txt1, Len(txt1) - 1) & "∏" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "e" Then
                        txt = Left(txt1, Len(txt1) - 1) & "–" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "®" Then
                        txt = Left(txt1, Len(txt1) - 1) & "æ" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "©" Then
                        txt = Left(txt1, Len(txt1) - 1) & " " & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "o" Then
                        txt = Left(txt1, Len(txt1) - 1) & "„" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "´" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Ë" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "¨" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Ì" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "u" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Û" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "≠" Then
                        txt = Left(txt1, Len(txt1) - 1) & "¯" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "i" Then
                        txt = Left(txt1, Len(txt1) - 1) & "›" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "™" Then
                        txt = Left(txt1, Len(txt1) - 1) & "’" & txt2
                    Else
                        txt = txt1 & alpha & txt2
                    End If
                Case "2"
                    If Mid(txt1, posSel, 1) = "a" Then
                        txt = Left(txt1, Len(txt1) - 1) & "µ" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "®" Then
                        txt = Left(txt1, Len(txt1) - 1) & "ª" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "©" Then
                        txt = Left(txt1, Len(txt1) - 1) & "«" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "e" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Ã" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "™" Then
                        txt = Left(txt1, Len(txt1) - 1) & "“" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "o" Then
                        txt = Left(txt1, Len(txt1) - 1) & "ﬂ" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "´" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Â" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "¨" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Í" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "u" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Ô" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "≠" Then
                        txt = Left(txt1, Len(txt1) - 1) & "ı" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "i" Then
                        txt = Left(txt1, Len(txt1) - 1) & "◊" & txt2
                    Else
                        txt = txt1 & alpha & txt2
                    End If
                Case "3"
                    If Mid(txt1, posSel, 1) = "a" Then
                        txt = Left(txt1, Len(txt1) - 1) & "∂" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "®" Then
                        txt = Left(txt1, Len(txt1) - 1) & "º" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "©" Then
                        txt = Left(txt1, Len(txt1) - 1) & "»" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "e" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Œ" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "™" Then
                        txt = Left(txt1, Len(txt1) - 1) & "”" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "o" Then
                        txt = Left(txt1, Len(txt1) - 1) & "·" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "¨" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Î" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "´" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Ê" & txt2
                    
                    ElseIf Mid(txt1, posSel, 1) = "u" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Ò" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "≠" Then
                        txt = Left(txt1, Len(txt1) - 1) & "ˆ" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "i" Then
                        txt = Left(txt1, Len(txt1) - 1) & "ÿ" & txt2
                    Else
                        txt = txt1 & alpha & txt2
                    End If
                Case "4"
                    If Mid(txt1, posSel, 1) = "a" Then
                        txt = Left(txt1, Len(txt1) - 1) & "∑" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "®" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Ω" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "©" Then
                        txt = Left(txt1, Len(txt1) - 1) & "…" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "e" Then
                        txt = Left(txt1, Len(txt1) - 1) & "œ" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "™" Then
                        txt = Left(txt1, Len(txt1) - 1) & "‘" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "o" Then
                        txt = Left(txt1, Len(txt1) - 1) & "‚" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "´" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Á" & txt2
                        ElseIf Mid(txt1, posSel, 1) = "¨" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Ï" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "u" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Ú" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "≠" Then
                        txt = Left(txt1, Len(txt1) - 1) & "˜" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "i" Then
                        txt = Left(txt1, Len(txt1) - 1) & "‹" & txt2
                    
                    ElseIf Mid(txt1, posSel, 1) = "y" Then
                        txt = Left(txt1, Len(txt1) - 1) & "¸" & txt2
                    Else
                        txt = txt1 & alpha & txt2
                    End If
                Case "5"
                    If Mid(txt1, posSel, 1) = "a" Then
                        txt = Left(txt1, Len(txt1) - 1) & "π" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "®" Then
                        txt = Left(txt1, Len(txt1) - 1) & "∆" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "©" Then
                        txt = Left(txt1, Len(txt1) - 1) & "À" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "e" Then
                        txt = Left(txt1, Len(txt1) - 1) & "—" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "™" Then
                        txt = Left(txt1, Len(txt1) - 1) & "÷" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "o" Then
                        txt = Left(txt1, Len(txt1) - 1) & "‰" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "´" Then
                        txt = Left(txt1, Len(txt1) - 1) & "È" & txt2
                        ElseIf Mid(txt1, posSel, 1) = "¨" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Ó" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "u" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Ù" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "≠" Then
                        txt = Left(txt1, Len(txt1) - 1) & "˘" & txt2
                        
                    ElseIf Mid(txt1, posSel, 1) = "i" Then
                        txt = Left(txt1, Len(txt1) - 1) & "ﬁ" & txt2
                    Else
                        txt = txt1 & alpha & txt2
                    End If
                Case "6"
                    If Mid(txt1, posSel, 1) = "a" Then
                        txt = Left(txt1, Len(txt1) - 1) & "©" & txt2
                    
                    ElseIf Mid(txt1, posSel, 1) = "e" Then
                        txt = Left(txt1, Len(txt1) - 1) & "™" & txt2
                    
                    ElseIf Mid(txt1, posSel, 1) = "o" Then
                        txt = Left(txt1, Len(txt1) - 1) & "´" & txt2
                    Else
                        txt = txt1 & alpha & txt2
                    End If
                Case "7"
                    If Mid(txt1, posSel, 1) = "o" Then
                        txt = Left(txt1, Len(txt1) - 1) & "¨" & txt2
                    
                    ElseIf Mid(txt1, posSel, 1) = "u" Then
                        txt = Left(txt1, Len(txt1) - 1) & "≠" & txt2
                    Else
                        txt = txt1 & alpha & txt2
                    End If
                Case "8"
                    If Mid(txt1, posSel, 1) = "a" Then
                        txt = Left(txt1, Len(txt1) - 1) & "®" & txt2
                    Else
                        txt = txt1 & alpha & txt2
                    
                    End If
                Case "9"
                    If Mid(txt1, posSel, 1) = "d" Or Mid(txt1, posSel, 1) = "D" Then
                        txt = Left(txt1, Len(txt1) - 1) & "Æ" & txt2
                    ElseIf Mid(txt1, posSel, 1) = "D" Then
                        txt = Left(txt1, Len(txt1) - 1) & "ß" & txt2
                    Else
                    
                        txt = txt1 & alpha & txt2
                    End If
                    
                Case Else
                    txt = txt1 & alpha & txt2
            End Select
            
            
            txtInput.Text = txt
            txtInput.SetFocus
            txtInput.SelStart = posSel + 1
            '''''''''''''''''''''''''''''''''''''
            
            If FlagShift = True Then
                If FlagCaplock = False Then
                    For i = 10 To 38
                        cmdText(i).Caption = LCase(cmdText(i).Caption)
                    Next
                    
                End If
                strSpecial = "1234567890"
                For i = 1 To 10
                    cmdText(i - 1).Caption = Mid(strSpecial, i, 1)
                Next i
            End If
    End Select
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  cmdText_Click"
End Sub

Public Sub AddTableType()
    On Error GoTo Handle
    cboTbaleType.Clear
        With cboTbaleType
            .AddItem "Bµn ch˜ nhÀt", 0
            .AddItem "Bµn vu´ng", 1
            .AddItem "Bµn trﬂn", 2
        End With
        cboTbaleType.ListIndex = 0
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  AddTableType"
End Sub



Private Sub Form_Activate()
    text_Input = ""
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
'    text_Input = ""
'    Text_Box = ""
'    Text_Field = ""
    
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
        Call AddTableType
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   Form_Load"
End Sub

Public Property Get Text_Box() As Variant
    On Error GoTo Handle
        Text_Field = txtInput.Text
    Exit Property
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Property


Public Property Let Text_Box(ByVal vNewValue As Variant)
    On Error GoTo Handle
        Text_Box = Text_Field
        Text_Field = ""
    Exit Property
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Property

Private Sub Form_Unload(Cancel As Integer)
    FlagShift = False
    FlagCaplock = False
End Sub

Public Property Let FormCallkeyboard(ByVal vNewValue As Variant)
    Keyboard = vNewValue
End Property

Public Sub AddPrint()
    On Error GoTo Handle
    Dim rsPrint As New ADODB.Recordset
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Set rsPrint = Open_Table(cnData, "Friendly_Printers")
    With rsPrint
        .addNew
        .Fields("PrtID") = MaxID("Friendly_Printers")
        .Fields("PrinterName") = Trim(txtInput.Text)
        .Fields("Store_ID") = Store_ID
        .Update
    End With
    
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " AddPrint"
End Sub

Public Function MaxID(TableName As String) As String
Dim str As String
On Error GoTo Handle:
    Dim rs As New ADODB.Recordset
    Set rs = OpenCriticalTable("Select Max(PrtID)as maxPrtID from " & TableName, cnData)
    If rs.RecordCount > 0 Then
        If Not rs.EOF Then
            str = Format(CDbl(rs.Fields("maxPrtID")) + 1, "00")
        Else
            str = "01"
        End If
    End If
    MaxID = str
Exit Function
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  MaxID"

End Function

Public Sub Update_TakeOut(S As String)
On Error GoTo Handle
Dim MaxInvoice As Integer

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Table_ID = S
    MaxInvoice = GetMaxInvoice_Number
    'SaveSettingStr "SYSTEM", "MaxInvoice", MaxInvoice, myIniFile
    
    If gfUpdate_Invoice_Totals(MaxInvoice) = True Then
        If gfUpdate_Invoice_OnHold(MaxInvoice) = False Then Exit Sub
        If gfUpdate_Invoice_Notes(MaxInvoice) = False Then Exit Sub
    End If
    
'    With frmOrder
'        .Get_Table_ID = Table_ID
'        .Get_Secion = "TO"
'        .GetBill_Number = MaxInvoice
'        .Show vbModal
'    End With
'    currentBill = ""
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Update_TakeOut"
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function gfUpdate_Invoice_OnHold(Invoice As Integer) As Boolean
On Error GoTo Handle
    Dim rsinvoice_hold As New ADODB.Recordset
    gfUpdate_Invoice_OnHold = False
    Set rsinvoice_hold = OpenCriticalTable("select * from Invoice_OnHold ", cnData)
        With rsinvoice_hold
        .Find "OnHoldID='" & Table_ID & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                'Khong ghi xuong Invoice_onHold
                currentBill = .Fields("Invoice_Number")
            Else
                ' ghi xuong Invoice_onHold
                currentBill = Invoice
                .addNew
                .Fields("Invoice_Number") = Invoice
                .Fields("OnHoldID") = Table_ID
                .Fields("Cashier_ID") = UserID
                .Fields("Store_ID") = Store_ID
                .Fields("Occupied") = -1
                .Fields("Section_ID") = "AR"
                .Fields("Status") = 0
                .Update
            End If
    End With
gfUpdate_Invoice_OnHold = True
    
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " gfUpdate_Invoice_OnHold"
    gfUpdate_Invoice_OnHold = False
End Function
Public Function gfUpdate_Invoice_Totals(Invoice As Integer) As Boolean
On Error GoTo Handle
    Dim rsInvoice_Total As New ADODB.Recordset
    gfUpdate_Invoice_Totals = False
        Set rsInvoice_Total = OpenCriticalTable("select * from Invoice_Totals ", cnData)
        With rsInvoice_Total
        .Find "Invoice_Number='" & Invoice & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                Dim rscust As New ADODB.Recordset
                Set rscust = Open_Table(cnData, "Customer")
                    rscust.Find "CustNum='" & .Fields("CustNum") & "'", , adSearchForward, adBookmarkFirst
                    If Not rscust.EOF Then
                        CustNo(0) = .Fields("CustNum")
                        CustNo(1) = rscust!CustName
                        CustNo(2) = rscust!Acct_Balance
                        'Discount = CDbl("0" & rscust.Fields("Discount"))
                    End If
                'Discount = .Fields("Discount")
            Else
                ' ghi xuong Invoice_onHold,Invoice_Total,Invoice_Notes
                .addNew
                .Fields("Invoice_Number") = Invoice
                .Fields("Store_ID") = Store_ID
                .Fields("CustNum") = "101"
                .Fields("DateTime") = Date & Format(Now, "HH:mm:ss")
                .Fields("InvoiceNotesUsed") = -1
                .Fields("Status") = "O"
                .Fields("Station_ID") = "AR"
                .Fields("Cashier_ID") = UserID
                .Fields("Payment_MeThod") = "CA"
                .Fields("InvType") = 0
                .Fields("Orig_OnHoldID") = Trim(Table_ID)
                '.Fields("Tax_Rate_ID") = 0
                .Update
            End If
    End With
gfUpdate_Invoice_Totals = True
    
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " gfUpdate_Invoice_OnHold"
    gfUpdate_Invoice_Totals = False
End Function

Public Function gfUpdate_Invoice_Notes(Invoice As Integer) As Boolean
On Error GoTo Handle
    Dim rsInvoice_Notes As New ADODB.Recordset
    gfUpdate_Invoice_Notes = False
    Set rsInvoice_Notes = Open_Table(cnData, "Invoice_Totals_Notes")
        
        With rsInvoice_Notes
        .Find "Invoice_Number='" & Invoice & "'", , adSearchForward, adBookmarkFirst
            If .EOF Then
                ' ghi xuong Invoice_Notes
                .addNew
                .Fields("Invoice_Number") = Invoice
                .Fields("Store_ID") = Store_ID
                .Fields("OpenTime") = Date & Format(Now, "HH:mm:ss")
                .Fields("ClosingTime") = ""
                .Update
            End If
        End With

    
    gfUpdate_Invoice_Notes = True
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " gfUpdate_Invoice_OnHold"
    gfUpdate_Invoice_Notes = False
End Function
Private Sub txtInput_Change()
Dim Pos As Integer
Pos = txtInput.SelStart
    If State_Type = True Then
        txtInput.Text = Format(txtInput.Text, "#,##0")
        txtInput.SelStart = Pos 'Len(txtInput.Text) + 1
    Else
        txtInput.Text = txtInput.Text
        txtInput.SelStart = Pos
    End If
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmdEnter_Click(48)
    
End Sub


Public Property Get Let_Text_Input() As Variant
    Let_Text_Input = text_Input
    text_Input = ""
End Property


Public Property Let Let_state(ByVal vNewValue As Variant)
    State_Type = vNewValue
End Property
