VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmMenuSelect 
   Caption         =   "Chän mãn ¨n"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14010
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
   ScaleHeight     =   10470
   ScaleWidth      =   14010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSearch 
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   ".VnArial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6720
      TabIndex        =   86
      Text            =   "NhËp tªn mãn cÇn t×m"
      Top             =   0
      Width           =   5655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10365
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   5025
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   0
         ScaleHeight     =   3375
         ScaleWidth      =   5055
         TabIndex        =   56
         Top             =   7200
         Width           =   5055
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   705
            Left            =   15
            TabIndex        =   57
            Top             =   0
            Width           =   3900
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   0
            Left            =   0
            TabIndex        =   72
            Top             =   735
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "1"
            PicturePosition =   131072
            Size            =   "1720;1402"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   1
            Left            =   990
            TabIndex        =   71
            Top             =   735
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "2"
            PicturePosition =   131072
            Size            =   "1720;1402"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   2
            Left            =   1980
            TabIndex        =   70
            Top             =   735
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "3"
            PicturePosition =   131072
            Size            =   "1720;1402"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   3
            Left            =   0
            TabIndex        =   69
            Top             =   1560
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "4"
            PicturePosition =   131072
            Size            =   "1720;1402"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   4
            Left            =   990
            TabIndex        =   68
            Top             =   1560
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "5"
            PicturePosition =   131072
            Size            =   "1720;1402"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   5
            Left            =   1980
            TabIndex        =   67
            Top             =   1560
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "6"
            PicturePosition =   131072
            Size            =   "1720;1402"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   6
            Left            =   0
            TabIndex        =   66
            Top             =   2385
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "7"
            PicturePosition =   131072
            Size            =   "1720;1402"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   7
            Left            =   990
            TabIndex        =   65
            Top             =   2385
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "8"
            PicturePosition =   131072
            Size            =   "1720;1402"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   8
            Left            =   1980
            TabIndex        =   64
            Top             =   2385
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "9"
            PicturePosition =   131072
            Size            =   "1720;1402"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   9
            Left            =   2970
            TabIndex        =   63
            Top             =   735
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "0"
            PicturePosition =   131072
            Size            =   "1720;1402"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   10
            Left            =   2970
            TabIndex        =   62
            Top             =   1560
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "00"
            PicturePosition =   131072
            Size            =   "1720;1402"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   11
            Left            =   2970
            TabIndex        =   61
            Top             =   2385
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "."
            PicturePosition =   131072
            Size            =   "1720;1402"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   480
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   690
            Index           =   12
            Left            =   3960
            TabIndex        =   60
            Top             =   30
            Width           =   1125
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "Bks"
            PicturePosition =   131072
            Size            =   "1984;1217"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   285
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   13
            Left            =   3960
            TabIndex        =   59
            Top             =   735
            Width           =   1125
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "CLR"
            PicturePosition =   131072
            Size            =   "1984;1402"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   315
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1625
            Index           =   14
            Left            =   3960
            TabIndex        =   58
            Top             =   1560
            Width           =   1125
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "Enter"
            PicturePosition =   131072
            Size            =   "1984;2866"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   315
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         FillColor       =   &H008080FF&
         ForeColor       =   &H008080FF&
         Height          =   615
         Left            =   -10
         ScaleHeight     =   615
         ScaleWidth      =   5115
         TabIndex        =   53
         Top             =   6480
         Width           =   5120
         Begin VB.Label lblTotal 
            BackStyle       =   0  'Transparent
            Caption         =   "Tæng céng:"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            TabIndex        =   55
            Tag             =   "L5"
            Top             =   120
            Width           =   1800
         End
         Begin VB.Label lblTotalAmt 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   ".VnArialH"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   525
            Left            =   2400
            TabIndex        =   54
            Top             =   0
            Width           =   2535
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flgOrder 
         Height          =   5895
         Left            =   0
         TabIndex        =   73
         Top             =   0
         Width           =   5050
         _ExtentX        =   8916
         _ExtentY        =   10398
         _Version        =   393216
         Rows            =   16
         Cols            =   6
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   16777215
         ForeColorFixed  =   16711680
         ForeColorSel    =   65280
         BackColorBkg    =   16777215
         GridColorFixed  =   12632256
         Redraw          =   -1  'True
         ScrollTrack     =   -1  'True
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         ScrollBars      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArialH"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Customer:"
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
         Left            =   90
         TabIndex        =   77
         Tag             =   "L14"
         Top             =   10350
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
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
         Left            =   90
         TabIndex        =   76
         Tag             =   "L34"
         Top             =   10740
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSForms.CommandButton MyButton1 
         Height          =   615
         Left            =   2530
         TabIndex        =   75
         Top             =   5880
         Width           =   2550
         BackColor       =   8454143
         Size            =   "4498;1085"
         Picture         =   "frmMenuSelect.frx":0000
         FontName        =   ".VnArial"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdListdown 
         Height          =   615
         Left            =   0
         TabIndex        =   74
         Top             =   5880
         Width           =   2550
         BackColor       =   8454143
         Size            =   "4498;1085"
         Picture         =   "frmMenuSelect.frx":018F
         FontName        =   ".VnArial"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.CommandButton cmdObj 
      BackColor       =   &H000000FF&
      Height          =   855
      Index           =   0
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   960
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   1
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   2
      Left            =   8175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   3
      Left            =   9630
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   4
      Left            =   11085
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   5
      Left            =   12540
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   6
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   7
      Left            =   8175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   8
      Left            =   9630
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   9
      Left            =   11085
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   10
      Left            =   12540
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   11
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   2265
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   12
      Left            =   8175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   2265
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   13
      Left            =   9630
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   2265
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   14
      Left            =   11085
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   2265
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   15
      Left            =   12540
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   2265
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   16
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3105
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   17
      Left            =   8175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3105
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   18
      Left            =   9630
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3105
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   19
      Left            =   11085
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3105
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   20
      Left            =   12540
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3105
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   21
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   3945
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   22
      Left            =   8175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   3945
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   23
      Left            =   9630
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3945
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   24
      Left            =   11085
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   3945
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   25
      Left            =   12540
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3945
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   26
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4785
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   27
      Left            =   8175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4785
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   28
      Left            =   9630
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4785
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   29
      Left            =   11085
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4785
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   30
      Left            =   12540
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4785
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   31
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5625
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   32
      Left            =   8175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5625
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   33
      Left            =   9630
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5625
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   34
      Left            =   11085
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5625
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   35
      Left            =   12540
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5625
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   36
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6465
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   37
      Left            =   8175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6465
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   38
      Left            =   9630
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6465
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   39
      Left            =   11085
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6465
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   40
      Left            =   12540
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6465
      Width           =   1455
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   ".VnArialH"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   740
      Index           =   0
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   720
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1580
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   45
      Left            =   12540
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7305
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   44
      Left            =   11085
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7305
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   43
      Left            =   9630
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7305
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   42
      Left            =   8175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7305
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   41
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7305
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   50
      Left            =   12540
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8145
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   49
      Left            =   11085
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8145
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   48
      Left            =   9630
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8145
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   47
      Left            =   8175
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8145
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Index           =   46
      Left            =   6720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8145
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5760
      Top             =   7800
   End
   Begin MSForms.CommandButton cmdFilter 
      Height          =   615
      Left            =   12405
      TabIndex        =   87
      Top             =   0
      Width           =   1575
      BackColor       =   65280
      Caption         =   "Läc ..."
      Size            =   "2778;1085"
      FontName        =   ".VnArial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdDelete 
      Height          =   585
      Left            =   6720
      TabIndex        =   85
      Top             =   9840
      Width           =   2295
      ForeColor       =   16711680
      BackColor       =   8438015
      VariousPropertyBits=   8388635
      Caption         =   "Xoùa"
      Size            =   "4048;1032"
      FontName        =   "VNI-Times"
      FontEffects     =   1073741827
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdExtraPrice 
      Height          =   585
      Left            =   9120
      TabIndex        =   84
      Tag             =   "L26"
      Top             =   9840
      Width           =   2295
      ForeColor       =   16777215
      BackColor       =   12582912
      VariousPropertyBits=   8388635
      Caption         =   "Giaù môû"
      Size            =   "4048;1032"
      FontName        =   "VNI-Times"
      FontEffects     =   1073741827
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdEditName 
      Height          =   585
      Left            =   9120
      TabIndex        =   83
      Top             =   9240
      Width           =   2295
      ForeColor       =   16711680
      BackColor       =   8438015
      VariousPropertyBits=   8388635
      Caption         =   "Söûa teân moùn aên"
      Size            =   "4048;1032"
      FontName        =   "VNI-Times"
      FontEffects     =   1073741827
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCookingMessage 
      Height          =   600
      Left            =   11520
      TabIndex        =   82
      Top             =   9240
      Width           =   2520
      ForeColor       =   16777215
      BackColor       =   12582912
      VariousPropertyBits=   8388635
      Caption         =   "Chuù thích moùn"
      Size            =   "4445;1058"
      FontName        =   "VNI-Times"
      FontEffects     =   1073741827
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdExit 
      Height          =   600
      Left            =   11520
      TabIndex        =   81
      Tag             =   "L18"
      Top             =   9840
      Width           =   2520
      ForeColor       =   16777215
      BackColor       =   33023
      VariousPropertyBits=   8388635
      Caption         =   "Tho¸t"
      Size            =   "4445;1058"
      FontName        =   ".VnArial"
      FontEffects     =   1073741827
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmddown 
      Height          =   630
      Left            =   5040
      TabIndex        =   80
      Top             =   9735
      Width           =   1545
      ForeColor       =   16777215
      BackColor       =   255
      VariousPropertyBits=   8388635
      Caption         =   "trang dø¬i"
      Size            =   "2716;1111"
      FontName        =   ".VnArialH"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdUp 
      Height          =   630
      Left            =   5040
      TabIndex        =   79
      Top             =   0
      Width           =   1575
      ForeColor       =   16777215
      BackColor       =   255
      VariousPropertyBits=   8388635
      Caption         =   "trang trªn"
      Size            =   "2778;1111"
      FontName        =   ".VnArialH"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNewBalance 
      Height          =   585
      Left            =   6720
      TabIndex        =   78
      Tag             =   "L16"
      Top             =   9240
      Width           =   2295
      ForeColor       =   16777215
      BackColor       =   33023
      VariousPropertyBits=   8388635
      Caption         =   "Löu"
      Size            =   "4048;1032"
      FontName        =   "VNI-Times"
      FontEffects     =   1073741827
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmMenuSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDepartment As New ADODB.Recordset
Dim strLast As String
Dim Desarr() As String 'Array caption
Dim rsJoin As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim PluNo As String
Dim formCallme As Integer
Dim ArrCommand() As String
Dim arrLoaded() As String
Dim LineNum As Double
Dim LineDelete, S As String
Dim rsInventory As New ADODB.Recordset
Dim rsShowPLU As New ADODB.Recordset
Dim rslinedelete As New ADODB.Recordset
Dim rsReserve As New ADODB.Recordset
Dim Table_ID As String
Dim Discount_Status, reason_discount As String
Dim isOK As Boolean
Dim Totals As Double

Private Sub cmdAlpha_Click(Index As Integer)
On Error GoTo Handle
    Select Case Index
        Case 0 To 11:
                txtQty.Text = txtQty.Text & cmdAlpha(Index).Caption
        Case 13
            txtQty.Text = ""
        Case 14
            If txtQty.Text = "" Then
                txtQty.Text = "1"
            Else
                ConQty = txtQty.Text
            End If
            txtQty.Text = ""
        Case 12
            If Len(txtQty) > 0 Then
              txtQty.Text = Left(txtQty, Len(txtQty) - 1)
            End If
    End Select
Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdAlpha_Click" & vbCrLf
MsgBox Err.Number & Err.Description & Me.name & " cmdAlpha_Click"
End Sub

Private Sub cmdBtn_Click(Index As Integer)
    On Error Resume Next 'GoTo Handle '
    Dim rs As New ADODB.Recordset
    Dim rsLast As New ADODB.Recordset
    Dim bt As CommandButton
    Dim i As Integer
    Dim ctrl As Control
    Set rsShowPLU = Nothing
    'cnData.Execute "Delete  from SetupPLU"
    i = 1
    Unload cmdObj(1)
    Call addButton(cmdBtn(Index).top + 15, cmdBtn(Index).Left + cmdBtn(Index).Width)
    
    If cnData.State <> 0 Then
        Dim strSql As String
        strSql = "SELECT Inventory.ItemNum, Inventory.ItemName, Inventory.Std_Price1," & _
        "Inventory.Std_Price2,Inventory.Std_Price3,Inventory.HH_Price1,Inventory.HH_Price2," & _
        "Inventory.HH_Price3,Inventory.EV_Price1,Inventory.EV_Price2,Inventory.EV_Price3," & _
        "Inventory.Picture,Inventory.Modify_Number,Inventory.LimitPrice,Inventory.F1, Departments.GIndex," & _
        "Inventory.F2,Inventory.F3,Inventory.F4,Inventory.F5,Departments.Dept_ID" & _
        " FROM Departments INNER JOIN Inventory ON (Departments.Dept_ID = Inventory.Dept_ID)" & _
        " WHERE (((Departments.GIndex)=" & Index & ")) order by Inventory.ItemNum ASC"
        
        Set rsJoin = OpenCriticalTable(strSql, cnData)

        If strLast <> "" Then
        Set rsLast = OpenCriticalTable("SELECT Inventory.ItemNum, Inventory.ItemName," & _
                                        "Inventory.Std_Price1, Inventory.Std_Price2,Inventory.Std_Price3," & _
                                        "Inventory.HH_Price1,Inventory.HH_Price2,Inventory.HH_Price3," & _
                                        "Inventory.EV_Price1,Inventory.EV_Price2,Inventory.EV_Price3," & _
                                        "Inventory.Picture,Inventory.Modify_Number,Inventory.F1,Inventory.F2," & _
                                        "Inventory.F3,Inventory.F4,Inventory.F5, Departments.GIndex,Departments.Dept_ID" & _
                                        " FROM Departments INNER JOIN Inventory ON (Departments.Dept_ID = Inventory.Dept_ID)" & _
                                        " WHERE (((Departments.GIndex)=" & strLast & "))and Inventory.F4='10'", cnData)
        i = 1
        Do While i <= rsLast.RecordCount 'Not rsLast.EOF
            
            Unload cmdSub(i)
            i = i + 1
            rsLast.MoveNext
        Loop
        End If
        'Set rs = OpenCriticalTable("Select * from SetupPLU", cnData)
    End If
    'Gan cac ma hang can hien thi vao rsShowPLU
        i = 1
        Do While Not rsJoin.EOF
        
        If ArrayFlag(rsJoin.Fields("F4"), 4) = 1 Then
            With rsShowPLU
                If .State = 0 Then
                    .Fields.Append "Index", adInteger
                    .Fields.Append "ItemNo", adVarWChar, 20
                    .Fields.Append "ItemName", adVarWChar, 50
                    .Fields.Append "Std_Price1", adVarWChar, 10
                    .Fields.Append "Std_Price2", adVarWChar, 10
                    .Fields.Append "Std_Price3", adVarWChar, 10
                    .Fields.Append "HH_Price1", adVarWChar, 10
                    .Fields.Append "HH_Price2", adVarWChar, 10
                    .Fields.Append "HH_Price3", adVarWChar, 10
                    .Fields.Append "EV_Price1", adVarWChar, 10
                    .Fields.Append "EV_Price2", adVarWChar, 10
                    .Fields.Append "EV_Price3", adVarWChar, 10
                    .Fields.Append "Picture", adVarWChar, 225
                    .Fields.Append "Modifier_No", adVarWChar, 225
                    .Fields.Append "Color", adVarWChar, 12
                    .Fields.Append "F1", adVarWChar, 2
                    .Fields.Append "F2", adVarWChar, 2
                    .Fields.Append "F3", adVarWChar, 2
                    .Fields.Append "F4", adVarWChar, 2
                    .Fields.Append "F5", adVarWChar, 2
                    .Fields.Append "Dept_ID", adVarWChar, 3
                    .Open
                End If
                .addNew
                .Fields("Index") = i
                .Fields("ItemNo") = rsJoin.Fields("ItemNum")
                .Fields("ItemName") = rsJoin.Fields("ItemName")
                .Fields("Std_Price1") = rsJoin.Fields("Std_Price1")
                .Fields("Std_Price2") = rsJoin.Fields("Std_Price2")
                .Fields("Std_Price3") = rsJoin.Fields("Std_Price3")
                .Fields("HH_Price1") = rsJoin.Fields("HH_Price1")
                .Fields("HH_Price2") = rsJoin.Fields("HH_Price2")
                .Fields("HH_Price3") = rsJoin.Fields("HH_Price3")
                .Fields("EV_Price1") = rsJoin.Fields("EV_Price1")
                .Fields("EV_Price2") = rsJoin.Fields("EV_Price2")
                .Fields("EV_Price3") = rsJoin.Fields("EV_Price3")
                .Fields("Picture") = rsJoin.Fields("Picture")
                .Fields("Modifier_No") = rsJoin.Fields("Modify_Number")
                .Fields("Color") = rsJoin.Fields("LimitPrice")
                .Fields("F1") = rsJoin.Fields("F1")
                .Fields("F2") = rsJoin.Fields("F2")
                .Fields("F3") = rsJoin.Fields("F3")
                .Fields("F4") = rsJoin.Fields("F4")
                .Fields("F5") = rsJoin.Fields("F5")
                .Fields("Dept_ID") = rsJoin.Fields("Dept_ID")
                .Update
        End With
'    Else
        i = i + 1
    End If
    rsJoin.MoveNext
    'i = i + 1
    Loop
        Call LoadCommandSub(rsShowPLU, "ItemNo", "ItemName")
        strLast = Index
    If rsShowPLU.State = 1 And rsShowPLU.RecordCount > 0 Then rsShowPLU.MoveFirst
    Exit Sub
'Handle:
'    MsgBox Err.Number & Err.Description & Me.Name & "  cmdBtn_Click"
End Sub

Private Sub cmdCookingMessage_Click()
On Error GoTo Handle
Dim strKit_Desc As String
    LineNum = flgOrder.TextMatrix(flgOrder.Row, 5)
    With frmKit_Desc
        .Show vbModal
        strKit_Desc = .Get_Kit_Desc
    End With
    With rsTemp
        If rsTemp.State <> 0 Then rsTemp.MoveFirst
        'If LineNum <> 0 Then
            .Find "Line_Number='" & LineNum & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    .Fields("Kit_Desc") = "(" & strKit_Desc & ")"
                    .Update
                End If
        'End If
    End With
Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdCookingMessage_Click" & vbCrLf
MsgBox Err.Number & Err.Description & Me.name & " cmdCookingMessage_Click"
End Sub


Private Sub cmddelete_Click()
    On Error GoTo Handle
'    If fClick = False Then Exit Sub
    If LineDelete = "" Then
        MsgBox "Vui lßng chän mãn cÇn xãa"
    Else
        Dim ID As String
        If rsTemp.State <> 0 Then
            With rsTemp
                .Find "Line_Number='" & LineDelete & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    
                    .Delete adAffectCurrent
                End If
            End With
            Set rslinedelete = Nothing
            
            Call SetFLGRIDORDER(rsTemp)
            If rsTemp.RecordCount = 0 Then
                Call Set_flgOrder
            End If
        End If
    End If
    Exit Sub

Handle:
''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdDelete_Click" & vbCrLf
    MsgBox Err.Number & Err.Description & Me.name & "  cmdDelete_Click"
End Sub


Private Sub cmdDown_Click()
On Error GoTo Handle
    Dim ctrl As Control
    Dim i As Integer
    If LastIndex + (rsDepartment.RecordCount Mod 12) > UBound(ArrCommand) Then Exit Sub
    For i = UBound(arrLoaded) - 1 To 0 Step -1
    DoEvents
        Unload cmdBtn(arrLoaded(i))
    Next i
    If LastIndex = 0 Then LastIndex = 12
    Call LoadCommand(12, ArrCommand, rsDepartment)
    LastIndex = LastIndex + 12
    Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdDown_Click" & vbCrLf
    MsgBox Err.Number & Err.Description & " - " & "Vui lßng ®îi trong gi©y l¸t ®Ó load d÷ liÖu"
End Sub

Private Sub cmdEditName_Click()
Dim S, S1 As String
    On Error GoTo Handle
    If LineDelete = "" Then
        MsgBox "B¹n ph¶i chän mãn cÇn söa tªn !", vbInformation
        Exit Sub
    End If
        S1 = flgOrder.TextMatrix(flgOrder.Row, 1)
        With frmKeyboard
            .FormCallkeyboard = "EditName"
            .txtInput.PasswordChar = ""
            .txtInput.Text = S1
            .txtInput.SelLength = 32
            .Show vbModal
            S = .Let_Text_Input
        End With
        With rsTemp
            .Find "Line_Number=" & LineDelete, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Fields("PluName") = S
                .Update
            End If
        End With

        Call SetFLGRIDORDER(rsTemp)
    Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & " cmdEditName_Click"
    MsgBox Err.Number & Err.Description & Me.name & " cmdEditName_Click"
End Sub

Private Sub cmdeditprice_Click()
Dim S As Double
    On Error GoTo Handle
        With frmPhimso
            .FormCall = 3
            .Show vbModal
            S = .Return_Value
        End With
        
        With rsTemp
            .Find "Line_Number=" & LineDelete, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                'Print #fFile, "Söa gi¸:" & vbTab & .Fields("PluName") & vbTab & .Fields("Std_Price1") & "-->" & S & vbTab & Now
                .Fields("Std_Price1") = S
                .Fields("Amt") = .Fields("Qty") * S
                .Update
            End If
        End With
        Call SetFLGRIDORDER(rsTemp)
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdeditprice_Click"
End Sub


Private Sub cmdExit_Click()
On Error GoTo Handle
    isOK = False
    Unload Me
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdexit_Click"
End Sub
Private Sub cmdListDown_Click()
On Error GoTo Handle

With flgOrder
    If .Row < .Rows - 13 Then
    .Row = .Row + 13
    .TopRow = .Row
    Else
        .Row = .Rows - 1
        .TopRow = .Row
    End If
    .AllowBigSelection = True
    .ScrollBars = flexScrollBarVertical
    .SelectionMode = flexSelectionByRow
    .Move .Rows
    .ScrollTrack = True
  
End With
Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & vbCrLf
    MsgBox Err.Number & Err.Description
End Sub

Private Sub cmdNewBalance_Click()
    isOK = True
    Unload Me
End Sub

Private Sub cmdSub_Click(Index As Integer)
    On Error GoTo Handle
       Call cmdAlpha_Click(14)
        LineNum = LineNum + 1
        With rsTemp
            If .State = 0 Then
                .Fields.Append "TableNo", adVarWChar, 50
                .Fields.Append "Line_Number", adDouble
                .Fields.Append "Dept_ID", adVarWChar, 3
                .Fields.Append "PLUNo", adVarWChar, 20
                .Fields.Append "PLUName", adVarWChar, 50
                .Fields.Append "Qty", adDouble
                .Fields.Append "Price", adDouble
                .Fields.Append "Amt", adDouble
                .Fields.Append "F1", adVarWChar, 2
                .Fields.Append "F2", adVarWChar, 2
                .Fields.Append "F3", adVarWChar, 2
                .Open
            End If
            rsShowPLU.Find "Index=" & Index, , adSearchForward, adBookmarkFirst
                If Not rsShowPLU.EOF Then
                    .addNew
                    .Fields("Qty") = ConQty
                    .Fields("TableNo") = Table_ID
                    .Fields("Line_Number") = LineNum
                    .Fields("PluNo") = rsShowPLU.Fields("ItemNo")
                    .Fields("PLUName") = rsShowPLU.Fields("ItemName")
                    .Fields("F1") = rsShowPLU!F1
                    .Fields("F2") = rsShowPLU!F2
                    .Fields("F3") = rsShowPLU!F3
                    .Fields("Dept_ID") = rsShowPLU!Dept_ID
                    With frmPhimso
                        .lblTitle.Caption = "NhËp gi¸ b¸n:"
                        .FormCall = 3
                        .Show vbModal
                        ExtrasPrice = .Return_Value
                    End With
                    .Fields("Price") = ExtrasPrice
                    .Fields("Amt") = ConQty * ExtrasPrice
                    .Update
                End If
        End With
            Call SetFLGRIDORDER(rsTemp)
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   cmdSub_Click"
End Sub

Private Sub cmdUp_Click()
On Error GoTo Handle
Dim i As Integer
    If LastIndex < 12 Then Exit Sub
   
    For i = UBound(arrLoaded) - 1 To 0 Step -1
    DoEvents
        Unload cmdBtn(arrLoaded(i))
    Next i
    If LastIndex > 24 Then
        LastIndex = LastIndex - 24
    Else
         LastIndex = LastIndex - 12
    End If
    
    If LastIndex < 0 Then Exit Sub
    If LastIndex = 12 Then LastIndex = 0
    Call LoadCommand(12, ArrCommand, rsDepartment)
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & " - " & "Vui lßng ®îi gi©y l¸t ®Ó load d÷ liÖu"
End Sub


Private Sub flgOrder_Click()
    On Error GoTo Handle
        With rslinedelete
            If .State = 0 Then
                .Fields.Append "Line", adVarWChar, 3
                .Open
            End If
            LineDelete = flgOrder.TextMatrix(flgOrder.Row, 5)
            .Find "Line ='" & LineDelete & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    .Delete adAffectCurrent
'                    .Requery
                Else
                    .addNew
                    .Fields("Line") = LineDelete
                    .Update
                End If
        End With
        flgOrder.CellTextStyle = flexTextFlat
        flgOrder.SelectionMode = flexSelectionByRow
        'flgOrder.CellBackColor = vbWhite
        flgOrder.Highlight = flexHighlightWithFocus
        If flgOrder.CellBackColor = vbBlue Then
            flgOrder.CellBackColor = vbWhite
        Else
            flgOrder.CellBackColor = vbBlue
        End If
    Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " flgOrder_Click" & vbCrLf
    MsgBox Err.Number & Err.Description & Me.name & " flgOrder_Click"
End Sub

Private Sub Form_Activate()
 On Error GoTo Handle
        Dim ctrl As Control
        Desarr = LoadLanguage(LngFile, "#01:007:")
'        For Each ctrl In Me
'        DoEvents
'            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Desarr(Mid(ctrl.Tag, 2))
'        Next ctrl

    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Form_Activate"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then frmHelp.Show vbModal
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    Dim i  As Integer
    ConQty = 1
    LineNum = 0
    Set rsTemp = New ADODB.Recordset
    Desarr = LoadLanguage(LngFile, "#01:007:")
'        If cnData.State = 0 Then
'            Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'        End If
        Set rsDepartment = Open_Table(cnData, "Departments")
        
        ReDim Preserve ArrCommand(rsDepartment.RecordCount)
        Do While Not rsDepartment.EOF
        DoEvents
            ArrCommand(i) = rsDepartment.Fields("GIndex")
        rsDepartment.MoveNext
        i = i + 1
        Loop
        Call LoadCommand(12, ArrCommand, rsDepartment)
        'end
        LastIndex = 12
        
        Call addButton(cmdBtn(0).top + 15, cmdBtn(0).Left + 1670)
        
        If rsDepartment.State = 1 Then
            If rsDepartment.RecordCount > 0 Then
                rsDepartment.MoveFirst
                Call cmdBtn_Click(rsDepartment.Fields("GIndex"))
            End If
        End If
        With rsTemp
            If .State = 0 Then
                .Fields.Append "TableNo", adVarWChar, 50
                .Fields.Append "Sec_No", adVarWChar, 2
                .Fields.Append "Line_Number", adDouble
                .Fields.Append "Dept_ID", adVarWChar, 3
                .Fields.Append "PLUNo", adVarWChar, 20
                .Fields.Append "PLUName", adVarWChar, 50
                .Fields.Append "Qty", adDouble
                .Fields.Append "Price", adDouble
                .Fields.Append "Amt", adDouble
                .Fields.Append "F1", adVarWChar, 2
                .Fields.Append "F2", adVarWChar, 2
                .Fields.Append "F3", adVarWChar, 2
                .Open
            End If
            i = 0
            If rsReserve.State <> 0 Then
                If rsReserve.RecordCount > 0 Then rsReserve.MoveFirst
            Else
                
            End If
            Do While Not rsReserve.EOF
                .addNew
                .Fields("Qty") = rsReserve.Fields("Qty")
                .Fields("TableNo") = Table_ID
                .Fields("Line_Number") = i
                .Fields("PluNo") = rsReserve.Fields("PluNo")
                .Fields("PLUName") = rsReserve.Fields("PluName")
                .Fields("Price") = rsReserve.Fields("Price")
                .Fields("Amt") = rsReserve.Fields("Amt")
                .Update
            rsReserve.MoveNext
            i = i + 1
            Loop
        End With
        
        Call Set_flgOrder
        Call SetFLGRIDORDER(rsTemp)
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub

Public Sub Set_flgOrder()
    On Error GoTo Handle
        With flgOrder
            .Cols = 6
            .Rows = 20
            .ColWidth(0) = 0
            .ColWidth(1) = 2200
            .ColWidth(2) = 500
            .ColWidth(3) = 1150
            .ColWidth(4) = 1150
            .ColWidth(5) = 0
            .TextMatrix(0, 1) = Desarr(19) '"Tên món"
            .TextMatrix(0, 2) = Desarr(20) ' "Sô' luong"
            .TextMatrix(0, 3) = Desarr(21) '" D/Giá"
            .TextMatrix(0, 4) = Desarr(22) '"T/Tiên`"
            .ColAlignment(1) = 2
            .ColAlignment(2) = 4
            .ColAlignment(3) = 6
            .ColAlignment(4) = 6
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            .TextMatrix(1, 3) = ""
            .TextMatrix(1, 4) = ""
            
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Set_flgOrder"
End Sub

Public Sub LoadCommand(n As String, Arr() As String, rs As ADODB.Recordset)
On Error GoTo Handle 'Resume Next
Dim btn As CommandButton
Dim iIndex As Integer
iIndex = 1
Dim i As Integer
For i = 1 To n
DoEvents
If LastIndex + (rs.RecordCount Mod 12) <= UBound(Arr) Then
If Arr(i - 1 + LastIndex) = "" Then Exit Sub
    iIndex = Arr((i - 1) + LastIndex)
    Load cmdBtn(iIndex)
    ReDim Preserve arrLoaded(i)
    arrLoaded(i - 1) = iIndex
    With cmdBtn(iIndex)
        If i = 1 Then
            .top = cmdUp.top + cmdUp.Height + 10
        Else
            .top = cmdBtn(iIndex - 1).top + cmdBtn(iIndex - 1).Height '+ 10
        End If

            rs.Find "GIndex='" & Arr(i - 1 + LastIndex) & "'", , adSearchForward, adBookmarkFirst
            If Not rs.EOF Then
                .Caption = rs.Fields("Description")
                .BackColor = rs.Fields("ColorDept")
            End If

        .Visible = True
        .Height = 790
        .Width = 1580

    End With
    Set btn = Nothing
End If
Next
Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " LoadCommand" & vbCrLf
    MsgBox Err.Number & Err.Description & Me.name & " LoadCommand"
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub LoadCommandSub(rs As ADODB.Recordset, strTenfield As String, strTenfield2 As String)
On Error GoTo Handle
Dim btn As CommandButton
Dim Index, i, j As Integer
j = 1
If rs.State <> 0 Then
    rs.MoveFirst
Else
    For i = 1 To 50
    DoEvents
'        cmdSub(i).Caption = ""
        cmdSub(i).Visible = False
    Next i
    Exit Sub
End If
For i = 1 To 50
DoEvents
    'Set cmdSub(i).PictureNormal = Nothing
    cmdSub(i).Picture = Nothing
    cmdSub(i).Caption = ""
Next i
    Do While Not rs.EOF
        If j > 50 Then Exit Sub
            With cmdSub(j)
                If Not rs.EOF Then
                    .Tag = rs.Fields("" & strTenfield & "")
                    .Caption = rs.Fields("" & strTenfield2 & "") '& vbCrLf & Format(rs.Fields("Std_Price1"), "#,##0")
                    .Font.Size = 10
                    .BackColor = HexToDec(rs.Fields("Color"))
                    If Dir(rs.Fields("Picture"), vbDirectory) <> "" Then
                      .Picture = LoadPicture(rs.Fields("Picture"))
                    End If
                    .Visible = True
                End If
            
            End With
        rs.MoveNext
        j = j + 1
    Loop
    For i = j + 1 To 50
    DoEvents
        cmdSub(i).Visible = False
    Next i
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  LoadCommandSub"
End Sub

Public Sub addButton(bttop As Integer, btleft As Integer)
    On Error GoTo Handle
    Load cmdObj(1)
        With cmdObj(1)
            .top = bttop + 15
            .Left = btleft
            .Height = 700
            .Width = 100
            .Visible = True
            .BackColor = vbRed
            
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "   addButton"
End Sub

Public Sub SetFLGRIDORDER(rs As ADODB.Recordset)
On Error GoTo Handle
        Dim incount As Integer
        Totals = 0
        If rs.RecordCount = 0 Then Exit Sub
        rs.MoveFirst
        With rs
            .Sort = "Line_Number DeSC"
            Do While Not .EOF
                incount = incount + 1
                flgOrder.Rows = rs.RecordCount + 1
                With flgOrder
                    .TextMatrix(incount, 1) = rs!PluName
                    .TextMatrix(incount, 2) = rs!Qty
                    .TextMatrix(incount, 3) = Format(rs!Price, formatNum)
                    .TextMatrix(incount, 4) = Format(rs!Amt, formatNum)
                    .TextMatrix(incount, 5) = rs!Line_Number
                    .RowHeight(incount) = 500
                    .CellAlignment = vbAlignTop
                    Totals = Totals + rs!Amt

                End With
            rs.MoveNext
            Loop
        End With
        If rs.RecordCount = 0 Then
            With flgOrder
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
                .TextMatrix(1, 3) = ""
                .TextMatrix(1, 4) = ""
                .TextMatrix(1, 5) = ""
            End With
        End If
        flgOrder.Row = flgOrder.Rows - 1
        If flgOrder.Row > 14 Then flgOrder.TopRow = flgOrder.Rows - 1
        lblTotalAmt.Caption = Format(Totals, "#,##0")
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "SetFLGRIDORDER"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set rsReserve = Nothing
    CloseRecordset rsReserve
    CloseRecordset rsDepartment
    CloseRecordset rsInventory
    CloseRecordset rsJoin
    CloseRecordset rsShowPLU
End Sub

Private Sub MyButton1_Click()
On Error GoTo Handle
With flgOrder
    If .Row >= 13 Then
    .Row = .Row - 13
    .TopRow = .Row
    Else
        .Row = 1
        .TopRow = .Row
    End If
'    .SetFocus
    .AllowBigSelection = True
    .ScrollBars = flexScrollBarVertical
    .SelectionMode = flexSelectionByRow
    .Move .Rows
    .ScrollTrack = True
End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description
End Sub


Public Property Let Get_Table_ID(ByVal vNewValue As Variant)
    Table_ID = vNewValue
End Property

Private Sub txtQty_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii < 32 And KeyAscii <> 13 Then Exit Sub
    Select Case KeyAscii
        Case 48 To 57, 46
        Case 13
            Dim ID As String
            If txtQty.Text = "" Then Exit Sub
            ID = TrimSpecialChar(txtQty.Text)
            txtQty.Text = ""
        Case Else:   KeyAscii = 0
    End Select
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtQty_KeyPress"
End Sub

Public Function TrimSpecial(ByVal str As String) As String
Dim i
Dim S As String
For i = 1 To Len(str)
    If Mid(str, i, 1) <> "," Then
        S = S & Mid(str, i, 1)
    Else
        S = S & "."
    End If
Next i
TrimSpecial = S
End Function

Public Sub cmdFilter_Click()
On Error Resume Next 'GoTo Handle '
    Dim rs As New ADODB.Recordset
    Dim rsLast As New ADODB.Recordset
    Dim bt As CommandButton
    Dim i As Integer
    Dim ctrl As Control
    Set rsShowPLU = Nothing
    'cnData.Execute "Delete  from SetupPLU"
    i = 1
'    If cnData.State <> 0 Then
        Dim strSql As String
        strSql = "SELECT Inventory.ItemNum, Inventory.ItemName, Inventory.Std_Price1," & _
           " Inventory.Std_Price2, Inventory.Std_Price3, Inventory.HH_Price1," & _
           " Inventory.HH_Price2, Inventory.HH_Price3, Inventory.EV_Price1," & _
           " Inventory.EV_Price2, Inventory.EV_Price3, Inventory.Picture," & _
           " Inventory.Modify_Number, Inventory.LimitPrice, Inventory.F1," & _
           " Departments.Index, Inventory.F2, Inventory.F3, Inventory.F4," & _
           " Inventory.F5" & _
            " FROM Departments INNER JOIN Inventory ON Departments.Dept_ID" & _
            " = Inventory.Dept_ID" & _
            " where INSTR(ItemName,""" & Trim(txtSearch.Text) & """)>0" & _
            " ORDER BY Inventory.ItemNum"
            
        Set rsJoin = OpenCriticalTable(strSql, cnData)

        If strLast <> "" Then
        Set rsLast = OpenCriticalTable("SELECT Inventory.ItemNum, Inventory.ItemName," & _
                                        "Inventory.Std_Price1, Inventory.Std_Price2,Inventory.Std_Price3," & _
                                        "Inventory.HH_Price1,Inventory.HH_Price2,Inventory.HH_Price3," & _
                                        "Inventory.EV_Price1,Inventory.EV_Price2,Inventory.EV_Price3," & _
                                        "Inventory.Picture,Inventory.Modify_Number,Inventory.F1,Inventory.F2," & _
                                        "Inventory.F3,Inventory.F4,Inventory.F5, Departments.Index,Departments.Dept_ID" & _
                                        " FROM Departments INNER JOIN Inventory ON (Departments.Dept_ID = Inventory.Dept_ID)" & _
                                        " WHERE (((Departments.Index)=" & strLast & "))and Inventory.F4='10'", cnData)
        i = 1
        Do While i <= rsLast.RecordCount 'Not rsLast.EOF
            
            Unload cmdSub(i)
            i = i + 1
            rsLast.MoveNext
        Loop
        
    End If
    'Gan cac ma hang can hien thi vao rsShowPLU
        i = 1
        If rsJoin.RecordCount > 0 Then rsJoin.MoveFirst
        Do While Not rsJoin.EOF
        If ArrayFlag(rsJoin.Fields("F4"), 4) = 1 Then
            With rsShowPLU
                If .State = 0 Then
                    .Fields.Append "Index", adInteger
                    .Fields.Append "ItemNo", adVarWChar, 20
                    .Fields.Append "ItemName", adVarWChar, 50
                    .Fields.Append "Std_Price1", adVarWChar, 10
                    .Fields.Append "Std_Price2", adVarWChar, 10
                    .Fields.Append "Std_Price3", adVarWChar, 10
                    .Fields.Append "HH_Price1", adVarWChar, 10
                    .Fields.Append "HH_Price2", adVarWChar, 10
                    .Fields.Append "HH_Price3", adVarWChar, 10
                    .Fields.Append "EV_Price1", adVarWChar, 10
                    .Fields.Append "EV_Price2", adVarWChar, 10
                    .Fields.Append "EV_Price3", adVarWChar, 10
                    .Fields.Append "Picture", adVarWChar, 225
                    .Fields.Append "Modifier_No", adVarWChar, 225
                    .Fields.Append "Color", adVarWChar, 12
                    .Fields.Append "F1", adVarWChar, 2
                    .Fields.Append "F2", adVarWChar, 2
                    .Fields.Append "F3", adVarWChar, 2
                    .Fields.Append "F4", adVarWChar, 2
                    .Fields.Append "F5", adVarWChar, 2
                    .Fields.Append "Dept_ID", adVarWChar, 3
                    .Open
                End If
                .addNew
                .Fields("Index") = i
                .Fields("ItemNo") = rsJoin.Fields("ItemNum")
                .Fields("ItemName") = rsJoin.Fields("ItemName")
                .Fields("Std_Price1") = rsJoin.Fields("Std_Price1")
                .Fields("Std_Price2") = rsJoin.Fields("Std_Price2")
                .Fields("Std_Price3") = rsJoin.Fields("Std_Price3")
                .Fields("HH_Price1") = rsJoin.Fields("HH_Price1")
                .Fields("HH_Price2") = rsJoin.Fields("HH_Price2")
                .Fields("HH_Price3") = rsJoin.Fields("HH_Price3")
                .Fields("EV_Price1") = rsJoin.Fields("EV_Price1")
                .Fields("EV_Price2") = rsJoin.Fields("EV_Price2")
                .Fields("EV_Price3") = rsJoin.Fields("EV_Price3")
                .Fields("Picture") = rsJoin.Fields("Picture")
                .Fields("Modifier_No") = rsJoin.Fields("Modify_Number")
                .Fields("Color") = rsJoin.Fields("LimitPrice")
                .Fields("F1") = rsJoin.Fields("F1")
                .Fields("F2") = rsJoin.Fields("F2")
                .Fields("F3") = rsJoin.Fields("F3")
                .Fields("F4") = rsJoin.Fields("F4")
                .Fields("F5") = rsJoin.Fields("F5")
                .Fields("Dept_ID") = rsJoin.Fields("Dept_ID")
                .Update
        End With
'    Else
        i = i + 1
    End If
    rsJoin.MoveNext
    'i = i + 1
    Loop
        Call LoadCommandSub(rsShowPLU, "ItemNo", "ItemName")
        
    Exit Sub
End Sub


Private Sub txtSearch_DblClick()
On Error GoTo Handle
    With frmKeyboard
        .txtInput.PasswordChar = ""
        .FormCallkeyboard = "Other"
        .Show vbModal
        txtSearch.Text = .Let_Text_Input
        cmdFilter_Click
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtSearch_DblClick"
End Sub

Private Sub txtSearch_GotFocus()
On Error GoTo Handle
    With txtSearch
        .SelStart = 0
        .SelLength = Len(.Text)
        .SelText = ""
    End With
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtSearch_GotFocus"
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    On Error GoTo Handle
        If KeyAscii = 13 Then Call cmdFilter_Click
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtSearch_KeyPress"
End Sub

Private Sub txtSearch_LostFocus()
On Error GoTo Handle
    With txtSearch
        .Text = "NhËp tªn mãn cÇn t×m"
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtSearch_LostFocus"
End Sub

Public Property Let Get_Records(ByVal vNewValue As Variant)
    Set rsReserve = vNewValue
End Property

Public Property Get return_Recordset() As Variant
    Set return_Recordset = rsTemp
End Property


Public Property Get Let_OK() As Variant
    Let_OK = isOK
End Property

