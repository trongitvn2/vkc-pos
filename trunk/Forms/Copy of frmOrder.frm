VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmOrder 
   BackColor       =   &H00404000&
   ClientHeight    =   11400
   ClientLeft      =   225
   ClientTop       =   120
   ClientWidth     =   15240
   ControlBox      =   0   'False
   FillColor       =   &H00008000&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11643.85
   ScaleMode       =   0  'User
   ScaleWidth      =   15960.92
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSDataGridLib.DataGrid dtgFind 
      Height          =   6735
      Left            =   8160
      TabIndex        =   141
      Top             =   840
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   24
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   ".VnArial NarrowH"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picwait 
      BackColor       =   &H000000FF&
      Height          =   735
      Left            =   5520
      ScaleHeight     =   675
      ScaleWidth      =   5955
      TabIndex        =   135
      Top             =   4800
      Visible         =   0   'False
      Width           =   6015
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "�ang x� l�......."
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   0
         TabIndex        =   136
         Top             =   0
         Width           =   6015
      End
   End
   Begin VB.Frame fraEdit 
      BackColor       =   &H00808080&
      Height          =   10080
      Left            =   8160
      TabIndex        =   121
      Top             =   480
      Visible         =   0   'False
      Width           =   10470
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H000080FF&
         Caption         =   "&��ng"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   123
         Tag             =   "L45"
         Top             =   8520
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "l�a ch�n ch�c n�ng"
         BeginProperty Font 
            Name            =   ".VnArialH"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   122
         Tag             =   "L34"
         Top             =   210
         Width           =   10245
      End
      Begin MSCommLib.MSComm MSCom 
         Left            =   1320
         Top             =   8640
         _ExtentX        =   1005
         _ExtentY        =   1005
         _Version        =   393216
         DTREnable       =   -1  'True
      End
      Begin MSForms.CommandButton cmdVoidTran 
         Height          =   1260
         Left            =   2160
         TabIndex        =   140
         Tag             =   "L12"
         Top             =   3960
         Width           =   1830
         ForeColor       =   16777215
         BackColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "S� kh�ch"
         Size            =   "3228;2222"
         FontName        =   ".VnArial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdOrderMan 
         Height          =   1260
         Left            =   4020
         TabIndex        =   139
         Top             =   3960
         Width           =   1830
         ForeColor       =   16777215
         BackColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "Nh�n vi�n ph�c v�"
         Size            =   "3228;2222"
         FontName        =   ".VnArial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdCustomer 
         Height          =   1260
         Left            =   5880
         TabIndex        =   138
         Tag             =   "L14"
         Top             =   3960
         Width           =   1830
         ForeColor       =   16777215
         BackColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "Kh�ch h�ng"
         Size            =   "3228;2222"
         FontName        =   ".VnArial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSendKP 
         Height          =   1260
         Left            =   7755
         TabIndex        =   134
         Tag             =   "L43"
         Top             =   2640
         Visible         =   0   'False
         Width           =   1830
         ForeColor       =   16777215
         BackColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "G�i b�p"
         Size            =   "3228;2222"
         FontName        =   ".VnArial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdTachmon 
         Height          =   1260
         Left            =   285
         TabIndex        =   133
         Tag             =   "L31"
         Top             =   2640
         Width           =   1830
         ForeColor       =   16777215
         BackColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "Chuy�n m�n"
         Size            =   "3228;2222"
         FontName        =   ".VnArial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdVAT 
         Height          =   1260
         Left            =   5880
         TabIndex        =   132
         Tag             =   "L48"
         Top             =   1320
         Width           =   1830
         ForeColor       =   16777215
         BackColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "Thu� VAT"
         Size            =   "3228;2222"
         FontName        =   ".VnArial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdReceiveMoney 
         Height          =   1260
         Left            =   4020
         TabIndex        =   131
         Tag             =   "L46"
         Top             =   1320
         Width           =   1830
         ForeColor       =   16777215
         BackColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "Ph� thu ti�n m�t"
         Size            =   "3228;2222"
         FontName        =   ".VnArial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdServiceCharge 
         Height          =   1260
         Left            =   2145
         TabIndex        =   130
         Tag             =   "L13"
         Top             =   1320
         Width           =   1830
         ForeColor       =   16777215
         BackColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "% Ph� ph�c v�"
         Size            =   "3228;2222"
         FontName        =   ".VnArial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdeditprice 
         Height          =   1260
         Left            =   285
         TabIndex        =   129
         Tag             =   "L49"
         Top             =   1320
         Width           =   1830
         ForeColor       =   16777215
         BackColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "S�a gi�"
         Size            =   "3228;2222"
         FontName        =   ".VnArial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdLanguageSelection 
         Height          =   1260
         Left            =   7755
         TabIndex        =   128
         Tag             =   "L47"
         Top             =   1320
         Width           =   1830
         ForeColor       =   16777215
         BackColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "L�a ch�n ng�n ng�"
         Size            =   "3228;2222"
         FontName        =   ".VnArial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdPrice3 
         Height          =   1260
         Left            =   5880
         TabIndex        =   127
         Tag             =   "L37"
         Top             =   2640
         Width           =   1830
         ForeColor       =   16777215
         BackColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "Gi� 3"
         Size            =   "3228;2222"
         FontName        =   ".VnArial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdPrice2 
         Height          =   1260
         Left            =   4020
         TabIndex        =   126
         Tag             =   "L36"
         Top             =   2640
         Width           =   1830
         ForeColor       =   16777215
         BackColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "Gi� 2"
         Size            =   "3228;2222"
         FontName        =   ".VnArial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton Price1 
         Height          =   1260
         Left            =   2145
         TabIndex        =   125
         Top             =   2640
         Width           =   1830
         ForeColor       =   16777215
         BackColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "Gi� 1"
         Size            =   "3228;2222"
         FontName        =   ".VnArial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   1260
         Left            =   285
         TabIndex        =   124
         Top             =   3960
         Width           =   1830
         ForeColor       =   16777215
         BackColor       =   12582912
         VariousPropertyBits=   8388635
         Caption         =   "Tr� ti�n c�c"
         Size            =   "3228;2222"
         FontName        =   ".VnArial"
         FontHeight      =   240
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
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
      Height          =   495
      Left            =   8160
      TabIndex        =   119
      Text            =   "Nh�p t�n m�n c�n t�m"
      Top             =   0
      Width           =   5655
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   7200
      Top             =   9360
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   46
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   94
      Top             =   9105
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   47
      Left            =   9615
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   93
      Top             =   9105
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   48
      Left            =   11070
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   9105
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   49
      Left            =   12525
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   9105
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   50
      Left            =   13980
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   90
      Top             =   9105
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   41
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   89
      Top             =   8145
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   42
      Left            =   9615
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   88
      Top             =   8145
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   43
      Left            =   11070
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   87
      Top             =   8145
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   44
      Left            =   12525
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   86
      Top             =   8145
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   45
      Left            =   13980
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   8145
      Width           =   1455
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Command2"
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   740
      Index           =   0
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   83
      Top             =   720
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1580
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   40
      Left            =   13980
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   7185
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   39
      Left            =   12525
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   7185
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   38
      Left            =   11070
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   7185
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   37
      Left            =   9610
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   7185
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   36
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   7185
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   35
      Left            =   13980
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   6225
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   34
      Left            =   12525
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   6225
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   33
      Left            =   11070
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6225
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   32
      Left            =   9610
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6225
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   31
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6225
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   30
      Left            =   13980
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   5265
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   29
      Left            =   12525
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   5265
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   28
      Left            =   11070
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   5265
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   27
      Left            =   9610
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   5265
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   26
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   5265
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   25
      Left            =   13980
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   4305
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   24
      Left            =   12525
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   4305
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   23
      Left            =   11070
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4305
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   22
      Left            =   9610
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4305
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   21
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   4305
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   20
      Left            =   13980
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3345
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   19
      Left            =   12525
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3345
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   18
      Left            =   11070
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3345
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   17
      Left            =   9610
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3345
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   16
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3345
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   15
      Left            =   13980
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2385
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   14
      Left            =   12525
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2385
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   13
      Left            =   11070
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2385
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   12
      Left            =   9610
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2385
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   11
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2385
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   10
      Left            =   13980
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   9
      Left            =   12525
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   8
      Left            =   11070
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   7
      Left            =   9610
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   6
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   5
      Left            =   13980
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   4
      Left            =   12525
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   3
      Left            =   11070
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   2
      Left            =   9610
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdSub 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   ".VnArial NarrowH"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   950
      Index           =   1
      Left            =   8160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton cmdObj 
      BackColor       =   &H000000FF&
      Height          =   855
      Index           =   0
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   135
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
      Height          =   11565
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6470
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         FillColor       =   &H008080FF&
         ForeColor       =   &H008080FF&
         Height          =   1455
         Left            =   -10
         ScaleHeight     =   1455
         ScaleWidth      =   5115
         TabIndex        =   99
         Top             =   6960
         Width           =   5120
         Begin VB.Frame Frame1 
            BackColor       =   &H00404000&
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   0
            TabIndex        =   100
            Top             =   -120
            Width           =   2565
            Begin VB.Label lblAdj2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "0%"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1635
               TabIndex        =   108
               Top             =   960
               Width           =   855
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Gi�m % T.U�ng:"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   375
               Left            =   0
               TabIndex        =   107
               Tag             =   " L40"
               Top             =   960
               Width           =   1560
            End
            Begin VB.Label lblAdj1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "0%"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1635
               TabIndex        =   106
               Top             =   600
               Width           =   855
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Gi�m % T.�n:"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   375
               Left            =   120
               TabIndex        =   105
               Tag             =   "L39"
               Top             =   600
               Width           =   1440
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00404000&
               Caption         =   "Gi�m T�ng H�:"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   375
               Left            =   120
               TabIndex        =   104
               Tag             =   "L9"
               Top             =   240
               Width           =   1440
            End
            Begin VB.Label lblDiscount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00404000&
               Caption         =   "0%"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   1515
               TabIndex        =   103
               Top             =   240
               Width           =   975
            End
            Begin VB.Label lblPersonNum 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   375
               Left            =   1590
               TabIndex        =   102
               Top             =   1320
               Width           =   885
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "S� kh�ch:"
               BeginProperty Font 
                  Name            =   ".VnArial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H0000FF00&
               Height          =   375
               Left            =   0
               TabIndex        =   101
               Tag             =   "L12"
               Top             =   1320
               Width           =   1560
            End
         End
         Begin VB.Label lblCustomer 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ABC"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   555
            Left            =   2610
            TabIndex        =   111
            Top             =   1020
            Width           =   2445
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblTotalAmt 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   645
            Left            =   2520
            TabIndex        =   110
            ToolTipText     =   "Click here for details !!"
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label lblTotal 
            BackStyle       =   0  'Transparent
            Caption         =   "Total:"
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
            Left            =   2655
            TabIndex        =   109
            Tag             =   "L5"
            Top             =   105
            Width           =   1800
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   0
         ScaleHeight     =   3255
         ScaleWidth      =   5055
         TabIndex        =   16
         Top             =   8400
         Width           =   5055
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   740
            Left            =   740
            TabIndex        =   0
            Top             =   0
            Width           =   1980
         End
         Begin MSForms.CommandButton cmdCustSelect 
            Height          =   735
            Left            =   4500
            TabIndex        =   142
            ToolTipText     =   "Kh�ch h�ng"
            Top             =   0
            Width           =   600
            ForeColor       =   255
            BackColor       =   16744576
            VariousPropertyBits=   8388635
            PicturePosition =   65543
            Size            =   "1058;1296"
            Picture         =   "frmOrder.frx":26EFB
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   405
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAdd 
            Height          =   735
            Left            =   2740
            TabIndex        =   118
            Top             =   0
            Width           =   735
            ForeColor       =   255
            BackColor       =   65280
            Caption         =   "+"
            Size            =   "1296;1296"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   405
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdMinus 
            Height          =   735
            Left            =   0
            TabIndex        =   117
            Top             =   0
            Width           =   735
            ForeColor       =   255
            BackColor       =   65280
            Caption         =   "-"
            Size            =   "1296;1296"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   405
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1530
            Index           =   14
            Left            =   3960
            TabIndex        =   71
            Top             =   1560
            Width           =   1125
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "Enter"
            PicturePosition =   131072
            Size            =   "1984;2699"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   315
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   13
            Left            =   3960
            TabIndex        =   70
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
            Height          =   735
            Index           =   12
            Left            =   3480
            TabIndex        =   69
            Top             =   0
            Width           =   1005
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "Bks"
            PicturePosition =   131072
            Size            =   "1773;1296"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   285
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   705
            Index           =   11
            Left            =   2970
            TabIndex        =   68
            Top             =   2385
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "."
            PicturePosition =   131072
            Size            =   "1720;1244"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   480
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   795
            Index           =   10
            Left            =   2970
            TabIndex        =   67
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
            Index           =   9
            Left            =   2970
            TabIndex        =   66
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
            Height          =   705
            Index           =   8
            Left            =   1980
            TabIndex        =   65
            Top             =   2385
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "9"
            PicturePosition =   131072
            Size            =   "1720;1244"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   705
            Index           =   7
            Left            =   990
            TabIndex        =   64
            Top             =   2385
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "8"
            PicturePosition =   131072
            Size            =   "1720;1244"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   705
            Index           =   6
            Left            =   0
            TabIndex        =   63
            Top             =   2385
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   12632256
            VariousPropertyBits=   8388635
            Caption         =   "7"
            PicturePosition =   131072
            Size            =   "1720;1244"
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
            TabIndex        =   62
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
            Index           =   4
            Left            =   990
            TabIndex        =   61
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
            Index           =   3
            Left            =   0
            TabIndex        =   60
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
            Index           =   2
            Left            =   1980
            TabIndex        =   59
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
            Index           =   1
            Left            =   990
            TabIndex        =   58
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
            Index           =   0
            Left            =   0
            TabIndex        =   57
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
      End
      Begin VB.PictureBox pictFunction 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   11535
         Left            =   5090
         ScaleHeight     =   11535
         ScaleWidth      =   1395
         TabIndex        =   15
         Top             =   0
         Width           =   1390
         Begin MSForms.CommandButton cmdAdjustment2 
            Height          =   825
            Left            =   0
            TabIndex        =   137
            Tag             =   "L40"
            Top             =   9720
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "Gi�m % Th�c u�ng"
            Size            =   "2355;1455"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdAdjustment1 
            Height          =   825
            Left            =   0
            TabIndex        =   116
            Tag             =   "L39"
            Top             =   8900
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "Gi�m % Th�c �n"
            Size            =   "2355;1455"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdItemInfor 
            Height          =   1000
            Left            =   0
            TabIndex        =   115
            Top             =   10485
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   255
            VariousPropertyBits=   8388635
            Caption         =   "Hi�u ch�nh m�n"
            Size            =   "2355;1764"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdReSendKP 
            Height          =   825
            Left            =   0
            TabIndex        =   84
            Tag             =   "L41"
            Top             =   8080
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "Nh�c m�n"
            Size            =   "2355;1455"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdItemDiscount 
            Height          =   825
            Left            =   0
            TabIndex        =   80
            Tag             =   "L11"
            Top             =   7280
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "Gi�m % m�n"
            Size            =   "2355;1455"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdEditName 
            Height          =   825
            Left            =   0
            TabIndex        =   79
            Tag             =   "L42"
            Top             =   6460
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "S�a t�n m�n"
            Size            =   "2355;1455"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdEditQuantity 
            Height          =   825
            Left            =   0
            TabIndex        =   78
            Tag             =   "L27"
            Top             =   5650
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "S�a sai SL"
            Size            =   "2355;1455"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdExtraPrice 
            Height          =   945
            Left            =   0
            TabIndex        =   77
            Tag             =   "L26"
            Top             =   4725
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "Gi� m�"
            Size            =   "2355;1667"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdCookingMessage 
            Height          =   945
            Left            =   0
            TabIndex        =   76
            Tag             =   "L38"
            Top             =   3775
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "Th�ng tin ch� d�n b�p"
            Size            =   "2355;1667"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdGopban 
            Height          =   945
            Left            =   0
            TabIndex        =   75
            Tag             =   "L24"
            Top             =   2835
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "G�p b�n"
            Size            =   "2355;1667"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdTranferTable 
            Height          =   945
            Left            =   0
            TabIndex        =   74
            Tag             =   "L15"
            Top             =   1890
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "Chuy�n b�n"
            Size            =   "2355;1667"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdDiscount 
            Height          =   945
            Left            =   0
            TabIndex        =   73
            Tag             =   "L9"
            Top             =   940
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "Gi�m %"
            Size            =   "2355;1667"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdDelete 
            Height          =   945
            Left            =   0
            TabIndex        =   72
            Tag             =   "L8"
            Top             =   0
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "X�a"
            Size            =   "2355;1667"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   735
         Left            =   0
         ScaleHeight     =   675
         ScaleWidth      =   5040
         TabIndex        =   6
         Top             =   0
         Width           =   5100
         Begin VB.Label lblBillNo 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   0
            TabIndex        =   14
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label lblCashierName 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            Caption         =   "Administrator"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   3510
            TabIndex        =   13
            Top             =   360
            Width           =   1545
         End
         Begin VB.Label lblStationName 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            Caption         =   "A"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   2310
            TabIndex        =   12
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblTableNo 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            Caption         =   "1"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   1170
            TabIndex        =   11
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblTable 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            Caption         =   "B�n s�"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   400
            Left            =   1170
            TabIndex        =   10
            Tag             =   "L2"
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label lblStation 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            Caption         =   "Khu v�c"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   400
            Left            =   2310
            TabIndex        =   9
            Tag             =   "L3"
            Top             =   0
            Width           =   1245
         End
         Begin VB.Label lblNhanVien 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            Caption         =   "Nh�n vi�n"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   400
            Left            =   3510
            TabIndex        =   8
            Tag             =   "L4"
            Top             =   0
            Width           =   1545
         End
         Begin VB.Label lblBill 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            Caption         =   "S� H�"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   400
            Left            =   0
            TabIndex        =   7
            Tag             =   "L1"
            Top             =   0
            Width           =   1185
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flgOrder 
         Height          =   5670
         Left            =   0
         TabIndex        =   5
         Top             =   720
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   10001
         _Version        =   393216
         Rows            =   12
         Cols            =   6
         BackColor       =   16777215
         ForeColor       =   0
         BackColorFixed  =   16777215
         ForeColorFixed  =   16711680
         BackColorSel    =   16776960
         ForeColorSel    =   16777215
         BackColorBkg    =   16777215
         GridColorFixed  =   12632256
         WordWrap        =   -1  'True
         Redraw          =   -1  'True
         ScrollTrack     =   -1  'True
         FocusRect       =   2
         HighLight       =   0
         FillStyle       =   1
         GridLinesFixed  =   1
         ScrollBars      =   0
         SelectionMode   =   1
         MergeCells      =   4
         PictureType     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial NarrowH"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSForms.CommandButton cmdListdown 
         Height          =   615
         Left            =   0
         TabIndex        =   82
         Top             =   6360
         Width           =   2550
         BackColor       =   8454143
         Size            =   "4498;1085"
         Picture         =   "frmOrder.frx":27B4D
         FontName        =   ".VnArial"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdListUp 
         Height          =   615
         Left            =   2550
         TabIndex        =   81
         Top             =   6360
         Width           =   2600
         BackColor       =   8454143
         Size            =   "4586;1085"
         FontName        =   ".VnArial"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
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
         TabIndex        =   3
         Tag             =   "L34"
         Top             =   10740
         Visible         =   0   'False
         Width           =   1695
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
         TabIndex        =   2
         Tag             =   "L14"
         Top             =   10350
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin MSForms.CommandButton cmdFilter 
      Height          =   495
      Left            =   13850
      TabIndex        =   120
      Tag             =   "L44"
      Top             =   0
      Width           =   1575
      BackColor       =   65280
      Caption         =   "L�c ..."
      Size            =   "2778;873"
      FontName        =   ".VnArial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdNewBalance 
      Height          =   1440
      Left            =   8040
      TabIndex        =   114
      Tag             =   "L16"
      Top             =   10080
      Width           =   1830
      ForeColor       =   16777215
      BackColor       =   33023
      VariousPropertyBits=   8388635
      Caption         =   "T�m t�nh"
      Size            =   "3228;2540"
      FontName        =   ".VnArial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdUp 
      Height          =   630
      Left            =   6480
      TabIndex        =   113
      Top             =   0
      Width           =   1575
      ForeColor       =   16777215
      BackColor       =   255
      VariousPropertyBits=   8388635
      Caption         =   "trang tr�n"
      Size            =   "2778;1111"
      FontName        =   ".VnArialH"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmddown 
      Height          =   630
      Left            =   6480
      TabIndex        =   112
      Top             =   10095
      Width           =   1545
      ForeColor       =   16777215
      BackColor       =   255
      VariousPropertyBits=   8388635
      Caption         =   "trang d��i"
      Size            =   "2716;1111"
      FontName        =   ".VnArialH"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdFunctionKey 
      Height          =   795
      Left            =   6480
      TabIndex        =   98
      Tag             =   "L28"
      Top             =   10725
      Width           =   1545
      ForeColor       =   16777215
      BackColor       =   16711680
      VariousPropertyBits=   8388635
      Caption         =   "Ph�m ch�c n�ng"
      Size            =   "2716;1402"
      FontName        =   ".VnArial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdBufferPrint 
      Height          =   1440
      Left            =   9840
      TabIndex        =   97
      Tag             =   "L25"
      Top             =   10080
      Width           =   1830
      ForeColor       =   16777215
      BackColor       =   128
      VariousPropertyBits=   8388635
      Caption         =   "In T�m t�nh"
      PicturePosition =   131072
      Size            =   "3228;2540"
      FontName        =   ".VnArial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdOtherPayment 
      Height          =   1440
      Left            =   11670
      TabIndex        =   96
      Tag             =   "L17"
      Top             =   10080
      Width           =   1815
      ForeColor       =   16777215
      BackColor       =   16711680
      VariousPropertyBits=   8388635
      Caption         =   "Thanh to�n"
      PicturePosition =   327683
      Size            =   "3201;2540"
      FontName        =   ".VnArial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdExit 
      Cancel          =   -1  'True
      Height          =   1440
      Left            =   13485
      TabIndex        =   95
      Tag             =   "L18"
      Top             =   10080
      Width           =   1920
      ForeColor       =   16777215
      BackColor       =   255
      VariousPropertyBits=   8388635
      Caption         =   "Tho�t"
      Size            =   "3387;2540"
      FontName        =   ".VnArial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Menu mnuList 
      Caption         =   "Chi ti�t order"
      Visible         =   0   'False
      Begin VB.Menu mnuDetails 
         Caption         =   "Chi ti�t Order"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "��ng"
      End
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDepartment As New ADODB.Recordset
Public strLast As String
Dim Desarr() As String 'Array caption
Dim rsSetupPLU As New ADODB.Recordset
Dim rsJoin As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim rsInvoice_Total As New ADODB.Recordset
Dim rsInvoice_Items As New ADODB.Recordset
Dim rsInvoice_Note As New ADODB.Recordset
Dim rsSystem As New ADODB.Recordset
Dim isLoaded, rightdelete As Boolean
Dim TotalAmt As Double
Dim PluNo As String
Dim formCallme As Integer
Dim ArrCommand() As String
Dim arrLoaded() As String
Dim PriceRate As Integer
Dim rsLocation As New ADODB.Recordset
Dim LocationFlag As Integer
Dim PriceFlag As Integer
Dim rsPriceTime As New ADODB.Recordset
Dim LineNum As Double
Dim LineDelete, S As String
Dim isExtrasPrice As Boolean
Dim rsDelete As New ADODB.Recordset
Dim rsInventory As New ADODB.Recordset
Dim arrSelect() As String
Dim rsShowPLU As New ADODB.Recordset
Dim strBill As Double
Dim service_Charge As Integer
Dim MoneyAmount As Double
Dim Adjtotal1, Adjtotal2, Adjtotal3, Adjtotal4 As Double
Dim blnPrice As Integer
Dim blnAutoselect_Price, lblAutoConsolidate As Boolean
Dim iset As Boolean
Dim blnEditQty As Boolean
Dim Item_Order_State As Boolean
Dim Adj1, Adj2 As Integer
Dim AllowDelete As Boolean
Dim rslinedelete As New ADODB.Recordset
Dim fClick As Boolean
Dim Discount As Integer
Dim Table_ID As String
Dim Personal, printcount As Integer
Dim PrintKit() As String
Dim diemtichluy As Double
Dim Emp_ID As String
Dim Discount_Status, reason_discount As String
Dim MeUnload As Boolean
Dim Dept_Index As Integer
Dim rsFind As New ADODB.Recordset
Dim isCust As Boolean
Dim delete_ordered As Boolean
Dim rsNew As New ADODB.Recordset


Private Sub cmdAdd_Click()
    On Error GoTo Handle
    Dim Qty_Adj As Integer
    'option 19-12-2013
Dim ID As String
If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "editquantity") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "editquantity") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
'end options

    If check_IsPrint(lblBillNo.Caption) = True Then Exit Sub
        If LineDelete = "" Then
            txtQty.Text = ""
            Exit Sub
        End If
        If txtQty.Text = "" Then
            Qty_Adj = 1
        Else
            Qty_Adj = Val("0" & txtQty.Text)
        End If
        With rsTemp
            .Find "Line_Number='" & LineDelete & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If rsInventory.State <> 0 Then rsInventory.MoveFirst
                rsInventory.Find "ItemNum='" & rsTemp.Fields("PluNo") & "'", , adSearchForward, adBookmarkFirst
                If Not rsInventory.EOF Then
                    If ArrayFlag(rsInventory.Fields("F1"), 1) = 1 Then
                        .Fields("Qty") = .Fields("Qty") + Qty_Adj
                        .Fields("Amt") = .Fields("Qty") * .Fields("Std_Price1")
                        If .Fields("Quanburned") <> .Fields("Qty") Then
                            .Fields("Status") = 0
                        End If
                        .Update
                    End If
                End If
            End If
        End With
        Call SetFLGRIDORDER(rsTemp)
        txtQty.Text = ""
       ' LineDelete = ""
'        blnEditQty = False
    Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdAdd_Click" & vbCrLf
    MsgBox Err.Number & Err.Description & Me.name & "  cmdAdd_Click"
End Sub

Private Sub cmdAdjustment1_Click()
On Error GoTo Handle
Dim i As Integer
Dim adj(1) As String
Dim AutoAdj As Boolean
Dim ID As String
            iset = True
'option 19-12-2013
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "adj1") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "adj1") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
'end options

            If check_IsPrint(lblBillNo.Caption) Then
             If ArrayFlag(SF(4), 8) = 0 Then Exit Sub
                    If UserLevel = 1 Then AllowDelete = True
                    If Not AllowDelete Then
                        With frmPassword
                            .FormActionKey = "Others"
                            .Show vbModal
                            ID = .return_Pass
                            If Not .Return_right Then Exit Sub
                        End With
                        If Return_right(ID, "Delete") Or UCase(ID) = "131112" Then AllowDelete = True
                    Else
                        ID = UserID
                    End If
                     GoTo 1
            Else
                AllowDelete = True
1:
                If AllowDelete = False Then Exit Sub
                    With frmPhimso
                        .lblTitle.Caption = "Nh�p % gi�m cho th�c �n:"
                        .FormCall = 3
                        .Show vbModal
                        Adj1 = .Return_Value
                    End With
                    
                     If ArrayFlag(SF(6), 3) = 0 Then
                        GoTo GiamTA
                    Else
                        With frmPro_Reason
                            .Show vbModal
                            reason_discount = .Let_Reason
                            If .Let_OK_Cancel = True Then
                                GoTo GiamTA
                            Else
                                Adj1 = 0
                                Exit Sub
                            End If
                        End With
                    End If
GiamTA:
                        If Adj1 > 100 Then
                            MsgBox "S� % gi�m kh�ng ���c l�n h�n 100%"
                            Adj1 = 0
                        End If
                        If ArrayFlag(SF(4), 1) = 1 Then
                            adj(0) = 1
                        Else
                            adj(0) = 0
                        End If
                        If ArrayFlag(SF(4), 5) = 1 Then AutoAdj = True
                            If AutoAdj = True Then Exit Sub
                                'Lay cac gia tri giam mon
                                If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                                    With rsInvoice_Total
                                        .Find "Invoice_Number=" & CDbl(lblBillNo.Caption), , adSearchForward, adBookmarkFirst
                                        If Not .EOF Then
                                            If .Fields("Adjustment1") <> 0 Then Adjtotal1 = 0
                                            If .Fields("Adjustment2") <> 0 Then Adjtotal2 = 0
                                            
                                            Call Get_Adjustment_Value_lastest(rsTemp, Adj1, 0)
                    '                        Call Confirm_Negative
                    
                                            If adj(0) = 1 Then
                                                !Adjustment1 = -Abs(Adjtotal1)
                                            Else
                                                !Adjustment1 = Adjtotal1
                                            End If
                                            rsInvoice_Total.Fields("Pro_Desc") = reason_discount
                                            rsInvoice_Total.Update
                                        End If
                                    End With
                    lblAdj1.Caption = Adj1 & "%"
                    'Print #fFile, "Gi�m % th�c �n:" & Adj1 & "%" & vbTab & Now
            End If
        AllowDelete = False
Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdAdjustment1_Click" & vbCrLf
MsgBox Err.Number & Err.Description & Me.name & " cmdAdjustment1_Click"
End Sub

Private Sub cmdAdjustment2_Click()
Dim i As Integer
Dim adj(1) As String
Dim AutoAdj As Boolean
'fraEdit.Visible = False
   Dim ID As String
    iset = True
    
'option 19-12-2013
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "adj2") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "adj2") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
'end options

    If check_IsPrint(lblBillNo.Caption) Then
     If ArrayFlag(SF(4), 8) = 0 Then Exit Sub
            If UserLevel = 1 Then AllowDelete = True
            If Not AllowDelete Then
                With frmPassword
                    .FormActionKey = "Others"
                    .Show vbModal
                    ID = .return_Pass
                    If Not .Return_right Then Exit Sub
                End With
                If Return_right(ID, "Delete") Or UCase(ID) = "131112" Then AllowDelete = True
            Else
                ID = UserID
            End If
             GoTo 1
    Else
        AllowDelete = True
1:
        If AllowDelete = False Then Exit Sub
            With frmPhimso
                .lblTitle.Caption = "Nh�p % gi�m cho th�c u�ng:"
                .FormCall = 3
                .Show vbModal
                Adj2 = .Return_Value
            End With
            
            If ArrayFlag(SF(6), 3) = 0 Then
                        GoTo GiamTU
                    Else
                        With frmPro_Reason
                            .Show vbModal
                            reason_discount = .Let_Reason
                            If .Let_OK_Cancel = True Then
                                GoTo GiamTU
                            Else
                                Adj2 = 0
                                Exit Sub
                            End If
                        End With
                    End If
GiamTU:
            
            If Adj2 > 100 Then
                MsgBox "S� % gi�m kh�ng ���c l�n h�n 100%"
                Adj2 = 0
            End If
                If cnData.State = 0 Then
                    Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
                End If
                  'Lay cac gia tri giam mon
                  If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
                  Call Get_Adjustment_Value_lastest(rsTemp, 0, Adj2)
        '                Call Confirm_Negative
                'Set rsInvoice_Total = OpenCriticalTable("select * from Invoice_Totals", cnData)
                    With rsInvoice_Total
                        .Find "Invoice_Number=" & CDbl(lblBillNo.Caption), , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                          !Adjustment2 = Adjtotal2
                          !Pro_Desc = reason_discount
                          rsInvoice_Total.Update
                        End If
                    End With
        '  End If
        lblAdj2.Caption = Adj2 & "%"
        'Print #fFile, "Gi�m % th�c u�ng:" & Adj2 & "%" & vbTab & Now
    End If
    AllowDelete = False
Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdAdjustment2_Click" & vbCrLf
MsgBox Err.Number & Err.Description & Me.name & " cmdAdjustment2_Click"

End Sub

Private Sub cmdAlpha_Click(Index As Integer)
On Error GoTo Handle
    Select Case Index
        Case 0 To 11:
            
                txtQty.Text = txtQty.Text & cmdAlpha(Index).Caption
        Case 13
            txtQty.Text = ""
            blnEditQty = False
            isExtrasPrice = False
        Case 14
            If txtQty.Text = "" Then txtQty.Text = "1"
            If ConQty = 1 Then
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
    Dept_Index = Index
    'cnData.Execute "delete * from SetupPLU"
    i = 1
    Unload cmdObj(1)
    Call addButton(cmdBtn(Index).top + 15, cmdBtn(Index).Left + cmdBtn(Index).Width)
    
    If cnData.State <> 0 Then
        Dim strSql As String
        strSql = "SELECT Inventory.ItemNum, Inventory.ItemName, Inventory.Std_Price1," & _
        "Inventory.Std_Price2,Inventory.Std_Price3,Inventory.HH_Price1,Inventory.HH_Price2," & _
        "Inventory.HH_Price3,Inventory.EV_Price1,Inventory.EV_Price2,Inventory.EV_Price3," & _
        "Inventory.Picture,Inventory.Modify_Number,Inventory.LimitPrice,Inventory.F1, Departments.Index," & _
        "Inventory.F2,Inventory.F3,Inventory.F4,Inventory.F5,Departments.Dept_ID" & _
        " FROM Departments INNER JOIN Inventory ON (Departments.Dept_ID = Inventory.Dept_ID)" & _
        " WHERE (((Departments.Index)=" & Index & ")) order by " & Right(Sort_By, Len(Sort_By) - InStr(Sort_By, ","))
        
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
                    .Fields.Append "ItemName", adVarWChar, 100
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

Private Sub cmdBufferPrint_Click()
   On Error GoTo Handle
   'options 19-12-2013
 Dim ID As String
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "bufferPrint") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "bufferPrint") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
'end options


    Picwait.Visible = True
    'Me.Enabled = False
    If MeUnload = False Then
            MeUnload = True
            cmdNewBalance.Enabled = False
            cmdOtherPayment.Enabled = False
            Call Update_Invoice_Total_Isprint(CDbl(lblBillNo.Caption))
            Call NewBalance
            Call Add_OrderMan
    '        Call AddDatato_Deleted_Items
            iset = False
         With frmShowBillBalance
            .Get_Style = formCallme
            .GetBill = CDbl(lblBillNo.Caption)
            .Show vbModal
        End With
    End If
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdBufferPrint_Click" & vbCrLf
    'MsgBox Err.Number & Err.Description & Me.Name & " cmdIntamtinh_Click"
    Unload Me
    Exit Sub
End Sub


Private Sub cmdClose_Click()
    fraEdit.Visible = False
End Sub

Private Sub cmdCookingMessage_Click()
On Error GoTo Handle
Dim strKit_Desc As String
    'LineDelete = flgOrder.TextMatrix(flgOrder.Row, 5)
    iset = False
    If LineDelete = "" Then LineDelete = flgOrder.Row
    With frmKit_Desc
        .Show vbModal
        strKit_Desc = .Get_Kit_Desc
    End With
    With rsTemp
        If rsTemp.State <> 0 Then rsTemp.MoveFirst
        'If LineNum <> 0 Then
            .Find "Line_Number='" & LineDelete & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    .Fields("Kit_Desc") = "(" & strKit_Desc & ")"
                    .Update
                End If
        'End If
    End With
    LineDelete = ""
Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdCookingMessage_Click" & vbCrLf
MsgBox Err.Number & Err.Description & Me.name & " cmdCookingMessage_Click"
End Sub

Private Sub cmdCustomer_Click()
    fraEdit.Visible = False
    iset = False
    With frmFindCustomer
        .Get_State = 1
        .FormCall = "CustomerSelect"
        .get_Amount = TotalAmt
        .Show vbModal
    End With
    lblCustomer.Caption = CustNo(1)
End Sub

Private Sub cmdCustSelect_Click()
    isCust = True
    txtQty.SetFocus
End Sub

Private Sub cmdCustSelect_DblClick(Cancel As MSForms.ReturnBoolean)
    cmdCustomer_Click
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo Handle
Dim ID As String
iset = False
If LineDelete = "" Then
    If Label3.BackColor = vbYellow Then
        lblTotalAmt.Caption = Format(TotalAmt, "#,##0")
        Discount = 0
        lblDiscount.Caption = "0%"
        Label3.BackColor = &H404000
        lblDiscount.BackColor = &H404000
        CustNo(0) = "101"
        lblCustomer.Caption = ""
        Exit Sub
    Else
        MsgBox "B�n ph�i ch�n m�n c�n x�a!", vbInformation
    Exit Sub
    End If
End If
If check_IsPrint(lblBillNo.Caption) Then
    If ArrayFlag(SF(4), 8) = 0 Then Exit Sub
        If UserLevel = 1 Or rightdelete = True Then
            AllowDelete = True
        Else
           With frmPassword
                .FormActionKey = "Others"
                .Show vbModal
                ID = .return_Pass
                If Not .Return_right Then Exit Sub
            End With
            If Return_right(ID, "Delete") Or UCase(ID) = "131112" Then
                AllowDelete = True
                UserID = Left(ID, 2)
            Else
                Exit Sub
            End If
        End If
        GoTo 1
Else

AllowDelete = True
1:
                With rsDelete
                    If .State = 0 Then
                        .Fields.Append "TableNo", adVarWChar, 50
                        .Fields.Append "BillNo", adDouble
                        .Fields.Append "Sec_No", adVarWChar, 2
                        .Fields.Append "LineNum", adVarWChar, 2
                        .Fields.Append "PLUNo", adVarWChar, 20
                        .Fields.Append "PLUName", adVarWChar, 100
                        .Fields.Append "Qty", adDouble
                        .Fields.Append "Std_Price1", adDouble
                        .Fields.Append "Amt", adDouble
                        .Fields.Append "F2", adVarWChar, 2
                        .Fields.Append "Cashier_ID", adVarWChar, 25
                        .Fields.Append "DateTime", adVarWChar, 50
                        .Fields.Append "Ordered", adBoolean
                        .Fields.Append "Reason", adVarWChar, 200
                        .Fields.Append "Kit_Desc", adVarWChar, 250
                        .Fields("Kit_Desc").Attributes = adColNullable
                        .Fields.Append "Line_Disc", adDouble
                        .Fields("Line_Disc").Attributes = adColNullable
                        .Fields.Append "Line_Disc_Desc", adVarWChar, 200
                        .Fields("Line_Disc_Desc").Attributes = adColNullable
                        .Fields.Append "PrintCount", adDouble
                        .Fields("PrintCount").Attributes = adColNullable
                        .Open
                    End If
                End With
                If rsTemp.State <> 0 Then
                    With rsTemp
        '                Do While Not rslinedelete.EOF
                            .Find "Line_Number='" & LineDelete & "'", , adSearchForward, adBookmarkFirst
                            If Not .EOF Then
                            If Not AllowDelete Then Exit Sub
                                'In du lieu xoa xuong file Log
                                'Print #fFile, "X�a m�n " & vbTab & Now
                                'Print #fFile, vbTab & .Fields("TableNo") & vbTab & .Fields("PluName") & vbTab & .Fields("Qty") & vbTab & .Fields("Std_Price1")
                                ' Gan du lieu xoa vao bang du lieu xoa
                                rsDelete.addNew
                                rsDelete!TableNo = .Fields("TableNo")
                                rsDelete!BillNO = CDbl("0" & lblBillNo.Caption)
                                rsDelete!Sec_No = .Fields("Sec_No")
                                rsDelete!LineNum = .Fields("Line_Number")
                                rsDelete!PluNo = .Fields("PluNo")
                                rsDelete!PluName = .Fields("PluName")
                                rsDelete!Qty = .Fields("Qty")
                                rsDelete!Std_Price1 = .Fields("Std_Price1")
                                rsDelete!printcount = printcount
            '                    rsDelete!Amt = .Fields("Amt")
                                rsInventory.Find "ItemNum='" & .Fields("PluNo") & "'", , adSearchForward, adBookmarkFirst
                                If Not rsInventory.EOF Then
                                    rsDelete!F2 = rsInventory.Fields("F2")
                                End If
                                If UCase(ID) = "131112" Then
                                    rsDelete!cashier_ID = "131112"
                                Else
                                    rsDelete!cashier_ID = Left(ID, 2)
                                End If
                                'rsDelete!DateTime = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") & Format(Now, "HH:mm:ss")
                                rsDelete!DateTime = DateDefault & Format(Now, "HH:mm:ss")
                                If .Fields("Status") = True Then
                                'update de sua xoa sau khi luu
                                If UserLevel = 1 Then delete_ordered = True
                                If Not delete_ordered Then Exit Sub
                                
                                ' end
                                    rsDelete!Ordered = 1
                                    iset = False
                                    With frmReason
                                        .Show vbModal
                                        rsDelete!reason = .GetReason
                                    End With
                                    If frmReason.GetReason = "" Then
                                        rsDelete!Ordered = 0
                                        AllowDelete = False
                                        rightdelete = False
                                        Exit Sub
                                    End If
                                Else
                                    rsDelete!Ordered = 0
                                End If
                                rsDelete!Kit_Desc = .Fields("Kit_Desc")
                                rsDelete!Line_Disc = .Fields("Line_Disc")
                                rsDelete!Line_Disc_Desc = .Fields("Line_Disc_Desc")
                                rsDelete.Update
                                ' Xoa du lieu hien tai
                                .Delete adAffectCurrent
                            End If
                    End With
                    Set rslinedelete = Nothing
                    Call SetFLGRIDORDER(rsTemp)
                    flgOrder.BackColor = -2147483643
                    lblTotalAmt.Caption = Format(TotalAmt - TotalAmt * Discount / 100, formatNum)
                    PluNo = ""
                    If rsTemp.RecordCount = 0 Then
                        Call Set_flgOrder
                    End If
        '        End If

            AllowDelete = False
            rightdelete = False
        End If
End If
    Exit Sub

Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdDelete_Click"
End Sub

Private Sub cmdDiscount_Click()
    On Error GoTo Handle
    Dim ID As String
    
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "discount") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "discount") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
            iset = True
            If check_IsPrint(lblBillNo.Caption) Then
                    If UserLevel = 1 Then AllowDelete = True
                    If Not AllowDelete Then
                        With frmPassword
                            .FormActionKey = "Others"
                            .Show vbModal
                            ID = .return_Pass
                            If Not .Return_right Then Exit Sub
                        End With
                        If Return_right(ID, "Delete") Or UCase(ID) = "131112" Then AllowDelete = True
                    Else
                        ID = UserID
                    End If
                     GoTo 1
            Else
            AllowDelete = True
1:
                If AllowDelete = False Then Exit Sub
               If ArrayFlag(SF(6), 3) = 0 Then
                    If txtQty.Text <> "" Then
                        Discount = CDbl("0" & txtQty.Text)
                    Else
                        Discount = getDiscount
                    End If
                    If Discount > 100 Then
                        MsgBox "S� % gi�m ph�i nho h�n ho�c b�ng 100%"
                        Discount = 0
                        Exit Sub
                    End If
                        lblTotalAmt.Caption = Format(CDbl(lblTotalAmt.Caption) - CDbl(lblTotalAmt.Caption) * Discount / 100, formatNum)
                        lblDiscount.Caption = Discount & "%"
                        txtQty.Text = ""
                        'Print #fFile, "Gi�m:" & Discount & "%" & vbTab & Now & vbTab & Table_ID & vbTab & lblBillNo.Caption & ":" & userName
            Else
            ' Update khuyen mai theo yeu cau
                If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
                If Not Check_Table_exist("Promotion") Then Call Create_Table_Discount
                With frmDiscount
                    .Get_Total = TotalAmt 'lblTotalAmt.Caption
                    .Show vbModal
                    If .Let_OK Then
                        Discount = .Let_Value
                        lblTotalAmt.Caption = Format(CDbl(lblTotalAmt.Caption) - CDbl(lblTotalAmt.Caption) * Discount / 100, formatNum)
                        lblDiscount.Caption = Discount & "%"
                        Discount_Status = .Let_Discount_Status
                        reason_discount = .Let_Reason_Discount
                        'Print #fFile, "Gi�m:" & Discount & "%" & vbTab & reason_discount & vbTab & Now
                    End If
                End With
            End If
        End If
        AllowDelete = False
    Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdDiscount_Click" & vbCrLf
    MsgBox Err.Number & Err.Description & Me.name & "  cmdDiscount_Click"
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
    MsgBox Err.Number & Err.Description & " - " & "Vui l�ng ��i trong gi�y l�t �� load d� li�u"
End Sub

Private Sub cmdEditName_Click()
Dim S, S1, ID As String
    On Error GoTo Handle
    'option 19-12-2013
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "editname") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "editname") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
'end options
    If LineDelete = "" Then
        MsgBox "B�n ph�i ch�n m�n c�n s�a t�n !", vbInformation
        Exit Sub
    End If
        iset = False
        S1 = flgOrder.TextMatrix(1, 1)
        With rsTemp
         .Find "Line_Number=" & LineDelete, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                    If .Fields("Status") = 0 Then
                        With frmKeyboard
                            .FormCallkeyboard = "EditName"
                            .txtInput.PasswordChar = ""
                            .txtInput.Text = S1
                            .txtInput.SelLength = 999
                            .Show vbModal
                            S = .Let_Text_Input
                            If Len(S) > 100 Then
                                MsgBox " T�n m�n kh�ng ���c v��t qu� 100 k� t�"
                                Exit Sub
                            End If
                            
                        End With
                        If ArrayFlag(.Fields("F1"), 5) = 0 Then
                            .Fields("PluName") = S
                            .Update
                        Else
                            MsgBox "M� h�ng n�y d�ng kho n�n kh�ng th� s�a t�n --> C� �i�u khi�n --> " & vbCrLf & " Ch�n PF1 --> b� d�u check � � e ( Kh�ng cho ph�p s�a t�n m�n)"
                        End If
                    Else
                        MsgBox "M�n n�y �� order kh�ng ���c s�a t�n m�n", vbInformation
                    End If
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
    Dim ID As String
    
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "editprice") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "editprice") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:

    iset = False
    If check_IsPrint(lblBillNo.Caption) Then Exit Sub
        With frmPhimso
            .lblTitle.Caption = "Nh�p gi� b�n:"
            .FormCall = 3
            .Show vbModal
            S = .Return_Value
        End With
        
        With rsTemp
            .Find "Line_Number=" & LineDelete, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                'Print #fFile, "S�a gi�:" & vbTab & .Fields("PluName") & vbTab & .Fields("Std_Price1") & "-->" & S & vbTab & Now
                .Fields("Std_Price1") = S
                .Fields("Amt") = .Fields("Qty") * S
                .Update
            End If
        End With
        Call SetFLGRIDORDER(rsTemp)
        fraEdit.Visible = False
    Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & " cmdeditprice_Click"
    MsgBox Err.Number & Err.Description & Me.name & " cmdeditprice_Click"
End Sub

Private Sub cmdEditQuantity_Click()
On Error GoTo Handle
Dim ID As String
    iset = True
     'options 19-12-2013
   
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "editquantity") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "editquantity") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
'end options

If check_IsPrint(lblBillNo.Caption) Then
    If ArrayFlag(SF(4), 8) = 0 Then Exit Sub
        If UserLevel = 1 Or rightdelete = True Then
            AllowDelete = True
        Else
           With frmPassword
                .FormActionKey = "Others"
                .Show vbModal
                ID = .return_Pass
                If Not .Return_right Then Exit Sub
            End With
            If Return_right(ID, "Delete") Or UCase(ID) = "131112" Then
                AllowDelete = True
                UserID = Left(ID, 2)
            Else
                Exit Sub
            End If
        End If
        GoTo 1
Else

AllowDelete = True
1:
        
            Call cmdAlpha_Click(14)
            blnEditQty = True
            If LineDelete = "" Then
                If blnEditQty = False Then ConQty = 1
                Exit Sub
            End If
            '22/8/2012
            If Not AllowDelete Then Exit Sub
            'end 22/8/2012
            With rsTemp
                .Find "Line_Number='" & LineDelete & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    If rsInventory.State <> 0 Then rsInventory.MoveFirst
                    rsInventory.Find "ItemNum='" & rsTemp.Fields("PluNo") & "'", , adSearchForward, adBookmarkFirst
                    If Not rsInventory.EOF Then
                        If ArrayFlag(rsInventory.Fields("F1"), 1) = 1 Then
                            If .Fields("Qty") < ConQty Then
                                .Fields("Qty") = ConQty
                            Else
                                .Fields("Qty") = .Fields("Qty") - ConQty
                                With rsDelete
                                    If .State = 0 Then
                                        .Fields.Append "TableNo", adVarWChar, 50
                                        .Fields.Append "BillNo", adDouble
                                        .Fields.Append "Sec_No", adVarWChar, 2
                                        .Fields.Append "LineNum", adVarWChar, 2
                                        .Fields.Append "PLUNo", adVarWChar, 20
                                        .Fields.Append "PLUName", adVarWChar, 100
                                        .Fields.Append "Qty", adDouble
                                        .Fields.Append "Std_Price1", adDouble
                                        .Fields.Append "Amt", adDouble
                                        .Fields.Append "F2", adVarWChar, 2
                                        .Fields.Append "Cashier_ID", adVarWChar, 25
                                        .Fields.Append "DateTime", adVarWChar, 50
                                        .Fields.Append "Ordered", adBoolean
                                        .Fields.Append "Reason", adVarWChar, 200
                                        .Fields.Append "Kit_Desc", adVarWChar, 250
                                        .Fields("Kit_Desc").Attributes = adColNullable
                                        .Fields.Append "Line_Disc", adDouble
                                        .Fields("Line_Disc").Attributes = adColNullable
                                        .Fields.Append "Line_Disc_Desc", adVarWChar, 250
                                        .Fields("Line_Disc_Desc").Attributes = adColNullable
                                        .Fields.Append "PrintCount", adDouble
                                        .Fields("PrintCount").Attributes = adColNullable
                                        .Open
                                    End If
                                        .Find "LineNum='" & LineDelete & "'", , adSearchForward, adBookmarkFirst
                                        If Not .EOF Then
                                            !Qty = !Qty + ConQty
                                        Else
                                            .addNew
                                            !TableNo = rsTemp.Fields("TableNo")
                                            !BillNO = CDbl("0" & lblBillNo.Caption)
                                            !Sec_No = rsTemp.Fields("Sec_No")
                                            !LineNum = rsTemp.Fields("Line_Number")
                                            !PluNo = rsTemp.Fields("PluNo")
                                            !PluName = rsTemp.Fields("PluName")
                                            !Qty = ConQty
                                            !Std_Price1 = rsTemp.Fields("Std_Price1")
                                            !Amt = rsTemp.Fields("Amt")
                                            !printcount = printcount
                                            rsInventory.Find "ItemNum='" & rsTemp.Fields("PluNo") & "'", , adSearchForward, adBookmarkFirst
                                            If Not rsInventory.EOF Then
                                                !F2 = rsInventory.Fields("F2")
                                            End If
                                            '!cashier_ID = UserID
                                            If UCase(ID) = "131112" Then
                                                !cashier_ID = "131112"
                                            Else
                                                !cashier_ID = Left(ID, 2)
                                            End If
                                            !DateTime = DateDefault & Format(Now, "HH:mm:ss")
                                            If rsTemp.Fields("Status") = True Then
                                                !Ordered = True
                '                                frmReason.Show vbModal
                '                                !Reason = frmReason.GetReason
                                            Else
                                                rsDelete!Ordered = False
                                            End If
                                            !Kit_Desc = rsTemp.Fields("Kit_Desc")
                                            !Line_Disc = rsTemp.Fields("Line_Disc")
                                            !Line_Disc_Desc = rsTemp.Fields("Line_Disc_Desc")
                                            .Update
                                            'Ghi du lieu xuong file Log
                                            'Print #fFile, "S�a sai s� l��ng " & vbTab & Now
                                            'Print #fFile, vbTab & .Fields("PluName") & vbTab & "SL C�:" & .Fields("Qty") & vbTab & "SL m�i:" & ConQty
        '                                    .Requery
                                        End If
                                End With
                            End If
                            '.Fields("Amt") = ConQty * .Fields("Std_Price1")
                            .Fields("Amt") = .Fields("Qty") * .Fields("Std_Price1")
                            If .Fields("Quanburned") <> .Fields("Qty") Then
                                .Fields("Status") = 0
                            End If
                            .Update
                            If .Fields("Qty") = 0 Then
                                .Delete adAffectCurrent
        '                        .Requery
                            End If
                        End If
                    End If
                End If
            End With
            Call SetFLGRIDORDER(rsTemp)
            ConQty = 1
            LineDelete = ""
            blnEditQty = False
        
    End If
    AllowDelete = False
    rightdelete = False
Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & " cmdEditQuantity_Click"
MsgBox Err.Number & Err.Description & Me.name & "  cmdEditQuantity_Click"
End Sub

Private Sub cmdExit_Click()
On Error GoTo Handle
Dim ans As Integer
Dim BillNo_Cancel As Double

    BillNo_Cancel = CDbl("0" & Trim(lblBillNo.Caption))
    If formCallme = 1 Then
        Unload Me
    Else
        If TotalAmt <> 0 Then
            ans = MsgBox("Giao d�ch �ang th�c hi�n, B�n c� mu�n l�u kh�ng?", vbYesNoCancel)
            If ans = vbYes Then
                Call cmdNewBalance_Click
            ElseIf ans = vbNo Then
                Set rsTemp = Nothing
                'Call delete_Bill_Null(lblBillNo.Caption)
                cnData.Execute "Update Invoice_Totals set InvoiceNotesUsed =false "
                Unload Me
            Else
                Exit Sub
            End If
        Else
            If rsTemp.State <> 0 Then
                Call NewBalance
                If BillNo_Cancel = GetMaxInvoice_Number Then
                    With rsDelete
                        .Find "Invoice_Number=" & BillNo_Cancel, , adSearchBackward, adBookmarkFirst
                            If Not .EOF Then
                               Call Update_Cancel_Bill(BillNo_Cancel)
                            Else
                                Call delete_Bill_Null(BillNo_Cancel)
                            End If
                    End With
                Else
                    Call Update_Cancel_Bill(BillNo_Cancel)
                End If
            Else
                If BillNo_Cancel = GetMaxInvoice_Number - 1 Then
                    Call delete_Bill_Null(BillNo_Cancel)
                Else
                    Call Update_Cancel_Bill(BillNo_Cancel)
                End If
            End If
            Unload Me
        End If
    End If
Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & " cmdexit_Click"
    MsgBox Err.Number & Err.Description & Me.name & " cmdexit_Click"
End Sub

Private Sub cmdExtraPrice_Click()
On Error GoTo Handle
'options 19-12-2013
 Dim ID As String
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "extraPrice") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "extraPrice") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
'end options
    isExtrasPrice = True
    iset = False
    With frmPhimso
        .lblTitle.Caption = "Nh�p gi� b�n:"
        .FormCall = 3
        .Show vbModal
        ExtrasPrice = .Return_Value
    End With
    iset = True
Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & " cmdExtraPrice_Click"
MsgBox Err.Number & Err.Description & Me.name & " cmdExtraPrice_Click"
End Sub


Private Sub cmdFunctionkey_Click()
    With fraEdit
        .top = pictFunction.top - 100
        .Left = pictFunction.Left
        .Visible = True
    End With
End Sub

Private Sub cmdGopban_Click()
On Error GoTo Handle
    Dim Table As String
    'options 19-12-2013
 Dim ID As String
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "joint_table") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "joint_table") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
'end options
    If check_IsPrint(lblBillNo.Caption) Then Exit Sub
    Picwait.Visible = True
    Me.Enabled = False
     If MsgBox("B�n c� ch�c ch�n mu�n g�p b�n kh�ng???", vbYesNo) = vbYes Then
        Table = lblTableNo.Caption
        currentBill = lblBillNo.Caption
        cmdNewBalance_Click
        With frmTablePlan
            .GetBillTranfer = CDbl(currentBill)
            .GetLocationTranfer = Sec_ID
            .GetTableTranfer = Table
            .FormState = 3
        End With
    
Else
    Picwait.Visible = False
    Me.Enabled = True
End If
Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & " cmdTranferTable_Click"
MsgBox Err.Number & Err.Description & Me.name & "  cmdTranferTable_Click"

End Sub



Private Sub cmdItemDiscount_Click()
Dim PriceDiscount As Double
On Error GoTo Handle
Dim ID As String
iset = True

 'option 19-12-2013
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "discount_item") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "discount_item") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
'end options

If check_IsPrint(lblBillNo.Caption) Then
 'If ArrayFlag(SF(4), 8) = 0 Then Exit Sub
        If UserLevel = 1 Then AllowDelete = True
        If Not AllowDelete Then
            With frmPassword
                .FormActionKey = "Others"
                .Show vbModal
                ID = .return_Pass
                If Not .Return_right Then Exit Sub
            End With
            If Return_right(ID, "Delete") Or UCase(ID) = "131112" Then AllowDelete = True
        Else
            ID = UserID
        End If
         GoTo 1
Else
    AllowDelete = True
1:
    If AllowDelete = False Then Exit Sub
    With frmPhimso
      .lblTitle.Caption = "Nh�p % gi�m cho m�n:"
        .FormCall = 3
        .cmdfree.Visible = True
        .Show vbModal
        LineDiscount = .Return_Value
        If LineDiscount > 100 Then
            MsgBox "Gi� tr� gi�m kh�ng th� l�n h�n 100%"
            LineDiscount = 0
            Exit Sub
        End If
    End With
    If LineDelete = "" Then Exit Sub
    With rsTemp
        .Find "Line_Number='" & LineDelete & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
'                PriceDiscount = .Fields("Std_Price1") - .Fields("Std_Price1") * LineDiscount / 100
'                .Fields("Std_Price1") = PriceDiscount
'                .Fields("Amt") = PriceDiscount * .Fields("Qty")
                If ArrayFlag(SF(6), 3) = 0 Then
                    rsTemp.Fields("amt") = rsTemp.Fields("Qty") * rsTemp.Fields("Std_Price1")
                    rsTemp.Fields("amt") = rsTemp.Fields("amt") - rsTemp.Fields("amt") * LineDiscount / 100
                    rsTemp.Fields("Line_Disc") = LineDiscount
                    rsTemp.Fields("Line_Disc_Desc") = ""
                    rsTemp.Update
                Else
                    With frmPro_Reason
                        .Show vbModal
                        If .Let_OK_Cancel = True Then
                            'Print #fFile, "Gi�m % m�n:" & LineDiscount & "%" & vbTab & rsTemp.Fields("PluName") & vbTab & Now
                            rsTemp.Fields("amt") = rsTemp.Fields("Qty") * rsTemp.Fields("Std_Price1")
                            rsTemp.Fields("amt") = rsTemp.Fields("amt") - rsTemp.Fields("amt") * LineDiscount / 100
                            rsTemp.Fields("Line_Disc") = LineDiscount
                            rsTemp.Fields("Line_Disc_Desc") = .Let_Reason
                            rsTemp.Update
                        End If
                    End With
                End If
            End If
    End With
    Call SetFLGRIDORDER(rsTemp)
End If
AllowDelete = False
Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdItemDiscount_Click" & vbCrLf
MsgBox Err.Number & Err.Description & Me.name & "cmdItemDiscount_Click"
End Sub

Private Sub cmdLanguageSelection_Click()
iset = False
    fraEdit.Visible = False
    frmLanguageSelection.Show vbModal
'    isLoaded = False
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
'    .SetFocus
    .AllowBigSelection = True
    .ScrollBars = flexScrollBarVertical
    .SelectionMode = flexSelectionByRow
    .Move .Rows
    .ScrollTrack = True
    '.CellBackColor = vbBlue
    
End With
Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & vbCrLf
    MsgBox Err.Number & Err.Description
End Sub

Private Sub cmdMinus_Click()
On Error GoTo Handle
'option 19-12-2013
Dim ID As String
If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "editquantity") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "editquantity") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
'end options

If check_IsPrint(lblBillNo.Caption) = True Then Exit Sub
    blnEditQty = True
    If txtQty.Text = "" Then
        MsgBox "Nh�p s� l��ng c�n tr� !", vbInformation
        Exit Sub
    End If
    With rsTemp
        If LineDelete = "" Then Exit Sub
        .Find "Line_Number='" & LineDelete & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If rsInventory.State <> 0 Then rsInventory.MoveFirst
            rsInventory.Find "ItemNum='" & rsTemp.Fields("PluNo") & "'", , adSearchForward, adBookmarkFirst
            If Not rsInventory.EOF Then
                If ArrayFlag(rsInventory.Fields("F1"), 1) = 1 Then
                    If Val(txtQty.Text) > .Fields("Qty") Then
                         txtQty.Text = ""
                        Exit Sub
                    End If
                    .Fields("Qty") = .Fields("Qty") - Val("0" & txtQty.Text)
                    With rsDelete
                        If .State = 0 Then
                            '''''''''''''''''''''''''''''''''''''
                            .Fields.Append "TableNo", adVarWChar, 50
                            .Fields.Append "BillNo", adDouble
                            .Fields.Append "Sec_No", adVarWChar, 2
                            .Fields.Append "LineNum", adVarWChar, 2
                            .Fields.Append "PLUNo", adVarWChar, 20
                            .Fields.Append "PLUName", adVarWChar, 100
                            .Fields.Append "Qty", adDouble
                            .Fields.Append "Std_Price1", adDouble
                            .Fields.Append "Amt", adDouble
                            .Fields.Append "F2", adVarWChar, 2
                            .Fields.Append "Cashier_ID", adVarWChar, 25
                            .Fields.Append "DateTime", adVarWChar, 50
                            .Fields.Append "Ordered", adBoolean
                            .Fields.Append "Reason", adVarWChar, 200
                            .Fields.Append "Kit_Desc", adVarWChar, 250
                            .Fields("Kit_Desc").Attributes = adColNullable
                            .Fields.Append "Line_Disc", adDouble
                            .Fields("Line_Disc").Attributes = adColNullable
                            .Fields.Append "Line_Disc_Desc", adVarWChar, 250
                            .Fields("Line_Disc_Desc").Attributes = adColNullable
                            .Fields.Append "PrintCount", adDouble
                            .Fields("PrintCount").Attributes = adColNullable
                            '''''''''''''''''''''''''''''''''
                            
                            .Open
                        End If
                            .addNew
                            !TableNo = rsTemp.Fields("TableNo")
                            !BillNO = CDbl("0" & lblBillNo.Caption)
                            !Sec_No = rsTemp.Fields("Sec_No")
                            !PluNo = rsTemp.Fields("PluNo")
                            !PluName = rsTemp.Fields("PluName")
                            !Qty = Val("0" & txtQty.Text)
                            !Std_Price1 = rsTemp.Fields("Std_Price1")
                            !Amt = rsTemp.Fields("Amt")
                            !printcount = printcount
                            rsInventory.Find "ItemNum='" & rsTemp.Fields("PluNo") & "'", , adSearchForward, adBookmarkFirst
                            If Not rsInventory.EOF Then
                                !F2 = rsInventory.Fields("F2")
                            End If
                            !cashier_ID = UserID
                            !DateTime = DateDefault & Format(Now, "HH:mm:ss")
                            If rsTemp.Fields("Status") = True Then
                                !Ordered = 1
'                                frmReason.Show vbModal
'                                !Reason = frmReason.GetReason
                            Else
                                rsDelete!Ordered = 0
                            End If
                            !Kit_Desc = rsTemp.Fields("Kit_Desc")
                            !Line_Disc = rsTemp.Fields("Line_Disc")
                            !Line_Disc_Desc = rsTemp.Fields("Line_Disc_Desc")
                            .Update
                        End With
                    .Fields("Amt") = .Fields("Qty") * .Fields("Std_Price1")
                    If .Fields("Quanburned") <> .Fields("Qty") Then
                        .Fields("Status") = 0
                    End If
                    .Update
                    If .Fields("Qty") = 0 Then
                        .Delete adAffectCurrent
                    End If
                End If
            End If
        End If
    End With
    Call SetFLGRIDORDER(rsTemp)
    txtQty.Text = ""
    LineDelete = ""
    blnEditQty = False
Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdEditQuantity_Click" & vbCrLf
MsgBox Err.Number & Err.Description & Me.name & "  cmdEditQuantity_Click"
End Sub

Private Sub cmdNewBalance_Click()
On Error GoTo Handle
    Picwait.Visible = True
    'Me.Enabled = False
    If MeUnload = False Then
         MeUnload = True
         cmdBufferPrint.Enabled = False
         cmdOtherPayment.Enabled = False
         Call NewBalance
'        'Print #fFile, "��ng b�n:" & Table_ID & vbTab & Now & vbTab & ":" & userName
'        'Print #fFile, "++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++"
        Unload Me
    End If
Exit Sub
Handle:
    DoEvents
    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & " cmdNewBalance_Click"
    MsgBox Err.Number & Err.Description & Me.name & "  cmdNewBalance_Click"
End Sub

Private Sub cmdOtherPayment_Click()
On Error GoTo Handle
Dim rsPer As New ADODB.Recordset
Dim blnPer As Boolean
  'options 19-12-2013
 Dim ID As String
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "payment") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "payment") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
'end options

Picwait.Visible = True
Me.Enabled = False
If MeUnload = False Then
    MeUnload = True
    cmdNewBalance.Enabled = False
    cmdBufferPrint.Enabled = False
    cmdOtherPayment.MousePointer = vbArrowHourglass
    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
        Set rsPer = Open_Table(cnData, "Invoice_Totals_Person_Mapping")
        If ArrayFlag(SF(3), 4) = 1 Then
            blnPer = True
        End If
    With rsPer
        If .State <> 0 Then
            If .RecordCount > 0 Then
                .MoveFirst
            End If
        End If
        .Find "Invoice_Number=" & CDbl("0" & lblBillNo.Caption), , adSearchForward, adBookmarkFirst
        If .EOF Then
            If blnPer = True Then
                MsgBox "B�n ph�i nh�p s� kh�ch !!!", vbInformation
                Exit Sub
            Else
                With rsPer
                    .addNew
                    .Fields("Invoice_Number") = CDbl("0" & lblBillNo.Caption)
                    .Fields("Store_ID") = Store_ID
                    .Fields("SeatNum") = 0
                    .Update
                End With
            End If
        End If
    End With
        'bo tam thoi dong giam % an uong
        Call Get_Adjustment_Value(rsTemp)
        Call NewBalance
        With frmCash
        iset = False
            If Discount > 0 Then
                .GetTotals = TotalAmt - TotalAmt * Discount / 100 + TotalAmt * service_Charge / 100 + Adjtotal1 + Adjtotal2 + Adjtotal3 + Adjtotal4 + MoneyAmount
            Else
                .GetTotals = TotalAmt + TotalAmt * service_Charge / 100 + Adjtotal1 + Adjtotal2 + Adjtotal3 + Adjtotal4 + MoneyAmount
            End If
            .GetTotal = TotalAmt '+ TotalAmt * service_Charge / 100 + TotalAmt * VAT / 100 + Adjtotal1 + Adjtotal2 + Adjtotal3 + Adjtotal4 '- Karaoke_Amt
            If CustNo(0) = "" Or CustNo(0) = "101" Then
                .GetCustomer = "101"
                .Get_Diem = 0
            Else
                .GetCustomer = CustNo(0)
                .Get_Diem = diemtichluy
            End If
            .GetBillNo = CDbl("0" & lblBillNo.Caption)
            .Show vbModal
        End With
            Set rsDelete = Nothing
        Unload Me
    End If
Exit Sub
Handle:
'Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " Othe_Payment_Click" & vbCrLf
MsgBox Err.Number & Err.Description & Me.name & ""

End Sub

Private Sub cmdPrice2_Click()
    'fraEdit.Visible = False
    blnPrice = 2
End Sub

Private Sub cmdPrice3_Click()
    blnPrice = 3
    fraEdit.Visible = False
End Sub

Private Sub cmdReceiveMoney_Click()
    On Error GoTo Handle
    fraEdit.Visible = False
    Dim ID As String
    
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "money") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "money") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:

    iset = False
    If check_IsPrint(lblBillNo.Caption) Then
     If ArrayFlag(SF(4), 8) = 0 Then Exit Sub
            If UserLevel = 1 Then AllowDelete = True
            If Not AllowDelete Then
                With frmPassword
                    .FormActionKey = "Others"
                    .Show vbModal
                    ID = .return_Pass
                    If Not .Return_right Then Exit Sub
                End With
                If Return_right(ID, "Delete") Or UCase(ID) = "131112" Then AllowDelete = True
            Else
                ID = UserID
            End If
             GoTo 1
    Else
        AllowDelete = True
1:
        If AllowDelete = False Then Exit Sub
        With frmPhimso
            .lblTitle.Caption = "Nh�p ti�n ph� thu:"
            .cmdAdd.Visible = True
            .cmdMinus.Visible = True
            .cmdAlpha(14).Visible = False
            .FormCall = 3
            .Show vbModal
            MoneyAmount = .Return_Value
        End With
        'Print #fFile, "Ph� thu ti�n m�t" & MoneyAmount & "%" & vbTab & Now
        lblTotalAmt.Caption = Format(CDbl(lblTotalAmt.Caption) + MoneyAmount, formatNum)
    End If
    AllowDelete = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdReceiveMoney_Click"

End Sub

Private Sub cmdReSendKP_Click()
On Error GoTo Handle
fraEdit.Visible = False
Dim ReQty As Double
If rslinedelete.State = 0 Then
    MsgBox "B�n ch�n m�n c�n nh�c !", vbInformation
    Exit Sub
End If
If check_already_exit_Invoicce_Number_Pending(lblBillNo.Caption) Then cnData.Execute "Delete * from Pending_Orders where Invoice_Number=" & lblBillNo.Caption
If rsTemp.State <> 0 And rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
iset = False
    With rsTemp
        rslinedelete.MoveFirst
        Do While Not rslinedelete.EOF
            .Find "Line_Number=" & rslinedelete.Fields("Line"), , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If .Fields("Qty") <> 1 Then
                    With frmPhimso
                        .lblTitle.Caption = "Nh�p s� l��ng c�n nh�c:"
                        .FormCall = 3
                        .Show vbModal
                        ReQty = .Return_Value
                    End With
                Else
                    ReQty = .Fields("Qty")
                End If
                If ReQty > .Fields("Qty") Then
                    MsgBox "Kh�ng th� nh�c s� l��ng nhi�u h�n hi�n t�i", vbInformation
                    Exit Sub
                ElseIf ReQty = 0 Then
                    Exit Sub
                End If
                    .Fields("QuanBurned") = .Fields("Qty") - ReQty
                    .Fields("Status") = 0
                    .Fields("Resend") = 1
                    .Update
            End If
        rslinedelete.MoveNext
        Loop
    End With
iset = True
Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdReSendKP_Click" & vbCrLf
    MsgBox Err.Number & Err.Description & Me.name & "cmdReSendKP_Click"
End Sub

Private Sub cmdSendKP_Click()
    On Error GoTo Handle
        fraEdit.Visible = False
        Call SendtoKP
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdSendKP_Click"
End Sub

Private Sub cmdServiceCharge_Click()
    On Error GoTo Handle
    fraEdit.Visible = False
    Dim ID As String
    
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "service_charge") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "service_charge") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:

    iset = False
    If check_IsPrint(lblBillNo.Caption) Then
     If ArrayFlag(SF(4), 8) = 0 Then Exit Sub
            If UserLevel = 1 Then AllowDelete = True
            If Not AllowDelete Then
                With frmPassword
                    .FormActionKey = "Others"
                    .Show vbModal
                    ID = .return_Pass
                    If Not .Return_right Then Exit Sub
                End With
                If Return_right(ID, "Delete") Or UCase(ID) = "131112" Then AllowDelete = True
            Else
                ID = UserID
            End If
             GoTo 1
    Else
        AllowDelete = True
1:
        If AllowDelete = False Then Exit Sub
        With frmPhimso
            .lblTitle.Caption = "Nh�p % ph� ph�c v�:"
            .FormCall = 3
            .Show vbModal
            service_Charge = .Return_Value
            If service_Charge > 100 Then
                MsgBox "% kh�ng v��t qu� gi�i h�n 100%"
                service_Charge = 0
            End If
            'Print #fFile, "Ph� ph�c v�:" & service_Charge & "%" & vbTab & Now
        End With
        lblTotalAmt.Caption = Format(CDbl(lblTotalAmt.Caption) + CDbl(lblTotalAmt.Caption) * service_Charge / 100, formatNum)
    End If
    AllowDelete = False
    Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdServiceCharge_Click" & vbCrLf
    MsgBox Err.Number & Err.Description & Me.name & "  cmdServiceCharge_Click"
End Sub

Private Sub cmdSub_Click(Index As Integer)
    On Error GoTo Handle
    If Item_Order_State Then
        Dim rsItem As New ADODB.Recordset
        Dim S As String
        Dim blnMenuOut As Boolean
        iset = True
        If check_IsPrint(lblBillNo.Caption) And ArrayFlag(SF(4), 8) = 0 Then Exit Sub
            'Lay so luong nhap
            Call cmdAlpha_Click(14)
    
            LineNum = LineNum + 1
            LineDelete = LineNum
            Dim str As String
        rsShowPLU.Find "Index=" & Index, , adSearchForward, adBookmarkFirst
        If Not rsShowPLU.EOF Then
            If ArrayFlag(rsShowPLU.Fields("F4"), 5) = 1 Then
                Call Update_OrderMan
            End If
            If ArrayFlag(rsShowPLU.Fields("F1"), 4) = 1 Then
                Call cmdExtraPrice_Click
            End If
            If ArrayFlag(rsShowPLU.Fields("F1"), 2) = 1 Then
                blnMenuOut = True
                Call cmdExtraPrice_Click
                iset = False
                With frmKeyboard
                    .FormCallkeyboard = "EditName"
                    .txtInput.PasswordChar = ""
                    .txtInput.SelLength = 32
                    .Show vbModal
                    S = .Let_Text_Input
                End With
            End If
            With rsTemp
                If .State = 0 Then
                    .Fields.Append "TableNo", adVarWChar, 50
                    .Fields.Append "Sec_No", adVarWChar, 2
                    .Fields.Append "Line_Number", adDouble
                    .Fields.Append "Dept_ID", adVarWChar, 3
                    .Fields.Append "PLUNo", adVarWChar, 20
                    .Fields.Append "PLUName", adVarWChar, 100
                    .Fields.Append "Qty", adDouble
                    .Fields.Append "Std_Price1", adDouble
                    .Fields.Append "Amt", adDouble
                    .Fields.Append "F1", adVarWChar, 2
                    .Fields.Append "F2", adVarWChar, 2
                    .Fields.Append "F3", adVarWChar, 2
                    .Fields.Append "Resend", adBoolean
                    .Fields.Append "Status", adBoolean
                    .Fields.Append "Quanburned", adDouble
                    .Fields.Append "Kit_Desc", adVarWChar, 250
                    .Fields("Kit_Desc").Attributes = adColNullable
                    .Fields.Append "Modifier_No", adVarWChar, 250
                    .Fields("Modifier_No").Attributes = adColNullable
                    .Fields.Append "Line_Disc", adDouble
                    .Fields.Append "Line_Disc_Desc", adVarWChar, 250
                    .Fields("Line_Disc_Desc").Attributes = adColNullable
                    .Fields.Append "TimeOrder", adVarWChar, 250
                    .Fields("TimeOrder").Attributes = adColNullable
                    .Open
    
                End If
                'If lblAutoConsolidate Then
                If .State = 1 And .RecordCount > 0 Then
                    .MoveFirst
                Else
                    GoTo 1
                End If
                .Find "PluNo='" & Trim(rsShowPLU.Fields("ItemNo")) & "'", , adSearchForward, adBookmarkFirst
                If .EOF Then
1:                  .addNew
                    .Fields("Qty") = ConQty
                Else
                    If lblAutoConsolidate = True Then
                        If .Fields("Status") = True Then
                            .Fields("Quanburned") = .Fields("Qty")
                        End If
                    
                        If blnEditQty = True Then ConQty = -ConQty
                        .Fields("Qty") = .Fields("Qty") + ConQty
                    Else
                        .addNew
                        If .Fields("Status") = True Then
                            .Fields("Quanburned") = .Fields("Qty")
                        End If
                        If blnEditQty = True Then ConQty = -ConQty
                        .Fields("Qty") = ConQty
                    End If
    '                Neu sua sai so luong bang 0 thi xoa luon record
                    If .Fields("Qty") = 0 Then
                        With rsDelete
                            If .State = 0 Then
                                .Fields.Append "TableNo", adVarWChar, 50
                                .Fields.Append "BillNo", adDouble
                                .Fields.Append "Sec_No", adVarWChar, 2
                                .Fields.Append "PLUNo", adVarWChar, 20
                                .Fields.Append "PLUName", adVarWChar, 100
                                .Fields.Append "Qty", adDouble
                                .Fields.Append "Std_Price1", adDouble
                                .Fields.Append "Amt", adDouble
                                .Fields.Append "F2", adVarWChar, 2
                                .Fields.Append "Cashier_ID", adVarWChar, 25
                                .Fields.Append "DateTime", adVarWChar, 50
                                .Fields.Append "Ordered", adBoolean
                                .Fields.Append "Reason", adVarWChar, 200
                                .Fields.Append "Kit_Desc", adVarWChar, 250
                                .Fields("Kit_Desc").Attributes = adColNullable
                                .Fields.Append "Line_Disc", adDouble
                                .Fields("Line_Disc").Attributes = adColNullable
                                .Fields.Append "Line_Disc_Desc", adVarWChar, 250
                                .Fields("Line_Disc_Desc").Attributes = adColNullable
                                .Open
                            End If
                        End With
                                    ' Gan du lieu xoa vao bang du lieu xoa
                        rsDelete.addNew
                        rsDelete!TableNo = .Fields("TableNo")
                        rsDelete!BillNO = CDbl("0" & lblBillNo.Caption)
                        rsDelete!Sec_No = .Fields("Sec_No")
                        rsDelete!PluNo = .Fields("PluNo")
                        rsDelete!PluName = .Fields("PluName")
                        rsDelete!Qty = -ConQty
                        rsDelete!Std_Price1 = .Fields("Std_Price1")
    '                    rsDelete!Amt = .Fields("Amt")
                        rsInventory.Find "ItemNum='" & .Fields("PluNo") & "'", , adSearchForward, adBookmarkFirst
                        If Not rsInventory.EOF Then
                            rsDelete!F2 = rsInventory.Fields("F2")
                        End If
                        rsDelete!cashier_ID = UserID
                        rsDelete!DateTime = DateDefault & Format(Now, "HH:mm:ss")
                        rsDelete!Ordered = 1
                        rsDelete!reason = " "
                        rsDelete!Kit_Desc = .Fields("Kit_Desc")
                        rsDelete!Line_Disc = .Fields("Line_Disc")
                        rsDelete!Line_Disc_Desc = .Fields("Line_Disc_Desc")
                        rsDelete.Update
                            'end
                        ' Xoa du lieu hien tai
                      .Delete adAffectCurrent
                      GoTo 2
                    End If
                End If
                .Fields("Sec_No") = Sec_ID
                .Fields("TableNo") = Table_ID
                .Fields("Line_Number") = LineNum
                .Fields("PluNo") = rsShowPLU.Fields("ItemNo")
                .Fields("TimeOrder") = Format(Now, "HH:mm:ss")
                If blnMenuOut = True Then
                    If S = "" Then S = "M�n ngo�i menu"
                    .Fields("PluName") = S
                Else
                    If ArrayFlag(rsShowPLU!F1, 6) = 1 Then
                    Dim isOK As Boolean
                        With frmKit_Desc
                            .txtKit_Desc = rsShowPLU.Fields("ItemName")
                            .txtKit_Desc.SelStart = Len(.txtKit_Desc.Text)
                            .Show vbModal
                            isOK = .Let_OK
                            If isOK = True Then
                                S = .Let_Kit_Des
                             Else
                               S = rsShowPLU.Fields("ItemName")
                             End If
                        End With
                        .Fields("PluName") = S
                    Else
                        .Fields("PluName") = rsShowPLU.Fields("ItemName")
                    End If
                End If
                '.Fields("Qty") = ConQty
                .Fields("F1") = rsShowPLU!F1
                .Fields("F2") = rsShowPLU!F2
                .Fields("F3") = rsShowPLU!F3
                .Fields("Dept_ID") = rsShowPLU!Dept_ID
                .Fields("Status") = False
                .Fields("Resend") = 0
                If isExtrasPrice = False Then
                    If PriceFlag = 1 Then
                        If LocationFlag = 1 Then
                            If rsPriceTime.RecordCount > 0 Then
                                rsPriceTime.Find "ID='" & 1 & "'", , adSearchBackward, adBookmarkFirst
                                If Not rsPriceTime.EOF Then
                                    If Format(Now, "HH:mm:ss") >= Format(rsPriceTime.Fields("StartTime"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(rsPriceTime.Fields("EndTime"), "HH:mm:ss") Then
                                        If rsShowPLU!Std_Price1 = 0 Then
                                            If ArrayFlag(rsShowPLU!F3, 6) = 1 Then
                                                If blnAutoselect_Price = True Then
                                                    If blnPrice = 1 Then
                                                        .Fields("Std_Price1") = rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100)
                                                    ElseIf blnPrice = 2 Then
                                                        .Fields("Std_Price1") = rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100)
                                                    Else
                                                        .Fields("Std_Price1") = rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100)
                                                    End If
                                                Else
                                                    If blnPrice = 1 Then
                                                        .Fields("Std_Price1") = rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100)
                                                    ElseIf blnPrice = 2 Then
                                                        .Fields("Std_Price1") = rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100)
                                                    Else
                                                        .Fields("Std_Price1") = rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100)
                                                    End If
                                                End If
                                            Else
                                                MsgBox " Kh�ng cho ph�p b�n gi� b�ng 0"
                                                .Delete adAffectCurrent
                                                GoTo 2
                                            End If
                                        Else
                                            If blnAutoselect_Price = True Then
                                                If blnPrice = 1 Then
                                                    .Fields("Std_Price1") = rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100)
                                                ElseIf blnPrice = 2 Then
                                                    .Fields("Std_Price1") = rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100)
                                                Else
                                                    .Fields("Std_Price1") = rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100)
                                                End If
                                            Else
                                                If blnPrice = 1 Then
                                                    .Fields("Std_Price1") = rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100)
                                                ElseIf blnPrice = 2 Then
                                                    .Fields("Std_Price1") = rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100)
                                                Else
                                                    .Fields("Std_Price1") = rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                rsPriceTime.MoveFirst
                                rsPriceTime.Find "ID='" & 2 & "'", , adSearchForward, adBookmarkFirst
                                If Not rsPriceTime.EOF Then
                                    If Format(Now, "HH:mm:ss") >= Format(rsPriceTime.Fields("StartTime"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(rsPriceTime.Fields("EndTime"), "HH:mm:ss") Then
                                        If rsShowPLU!HH_Price1 = 0 Then
                                            If ArrayFlag(rsShowPLU!F3, 6) = 1 Then
                                            If blnAutoselect_Price = True Then
                                                If blnPrice = 1 Then
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price1 + rsShowPLU!HH_Price1 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!HH_Price1 + rsShowPLU!HH_Price1 * PriceRate / 100)
                                                ElseIf blnPrice = 2 Then
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price2 + rsShowPLU!HH_Price2 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!HH_Price2 + rsShowPLU!HH_Price2 * PriceRate / 100)
                                                Else
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price3 + rsShowPLU!HH_Price3 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!HH_Price3 + rsShowPLU!HH_Price3 * PriceRate / 100)
                                                End If
                                            Else
                                                If blnPrice = 1 Then
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price1 + rsShowPLU!HH_Price1 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!HH_Price1 + rsShowPLU!HH_Price1 * PriceRate / 100)
                                                ElseIf blnPrice = 2 Then
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price2 + rsShowPLU!HH_Price2 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!HH_Price2 + rsShowPLU!HH_Price2 * PriceRate / 100)
                                                Else
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price3 + rsShowPLU!HH_Price3 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!HH_Price3 + rsShowPLU!HH_Price3 * PriceRate / 100)
                                                End If
                                            End If
                                            Else
                                                MsgBox " Kh�ng cho ph�p b�n gi� b�ng 0"
                                                .Delete adAffectCurrent
                                                GoTo 2
                                            End If
                                        Else
                                            If blnAutoselect_Price = True Then
                                                If blnPrice = 1 Then
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price1 + rsShowPLU!HH_Price1 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!HH_Price1 + rsShowPLU!HH_Price1 * PriceRate / 100)
                                                ElseIf blnPrice = 2 Then
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price2 + rsShowPLU!HH_Price2 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!HH_Price2 + rsShowPLU!HH_Price2 * PriceRate / 100)
                                                Else
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price3 + rsShowPLU!HH_Price3 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!HH_Price3 + rsShowPLU!HH_Price3 * PriceRate / 100)
                                                End If
                                            Else
                                                If blnPrice = 1 Then
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price1 + rsShowPLU!HH_Price1 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!HH_Price1 + rsShowPLU!HH_Price1 * PriceRate / 100)
                                                ElseIf blnPrice = 2 Then
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price2 + rsShowPLU!HH_Price2 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!HH_Price2 + rsShowPLU!HH_Price2 * PriceRate / 100)
                                                Else
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price3 + rsShowPLU!HH_Price3 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!HH_Price3 + rsShowPLU!HH_Price3 * PriceRate / 100)
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                rsPriceTime.MoveFirst
                                rsPriceTime.Find "ID='" & 3 & "'", , adSearchForward, adBookmarkFirst
                                If Not rsPriceTime.EOF Then
                                    If Format(Now, "HH:mm:ss") >= Format(rsPriceTime.Fields("StartTime"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(rsPriceTime.Fields("EndTime"), "HH:mm:ss") Then
                                        If rsShowPLU!EV_Price1 = 0 Then
                                            If ArrayFlag(rsShowPLU!F3, 6) = 1 Then
                                                If blnAutoselect_Price = True Then
                                                    If blnPrice = 1 Then
                                                        .Fields("Std_Price1") = rsShowPLU!EV_Price1 + rsShowPLU!EV_Price1 * PriceRate / 100
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!EV_Price1 + rsShowPLU!EV_Price1 * PriceRate / 100)
                                                    ElseIf blnPrice = 2 Then
                                                        .Fields("Std_Price1") = rsShowPLU!EV_Price2 + rsShowPLU!EV_Price2 * PriceRate / 100
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!EV_Price2 + rsShowPLU!EV_Price2 * PriceRate / 100)
                                                    Else
                                                        .Fields("Std_Price1") = rsShowPLU!EV_Price3 + rsShowPLU!EV_Price3 * PriceRate / 100
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!EV_Price3 + rsShowPLU!EV_Price3 * PriceRate / 100)
                                                    End If
                                                Else
                                                    If blnPrice = 1 Then
                                                        .Fields("Std_Price1") = rsShowPLU!EV_Price1 + rsShowPLU!EV_Price1 * PriceRate / 100
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!EV_Price1 + rsShowPLU!EV_Price1 * PriceRate / 100)
                                                    ElseIf blnPrice = 2 Then
                                                        .Fields("Std_Price1") = rsShowPLU!EV_Price2 + rsShowPLU!EV_Price2 * PriceRate / 100
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!EV_Price2 + rsShowPLU!EV_Price2 * PriceRate / 100)
                                                    Else
                                                        .Fields("Std_Price1") = rsShowPLU!EV_Price3 + rsShowPLU!EV_Price3 * PriceRate / 100
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!EV_Price3 + rsShowPLU!EV_Price3 * PriceRate / 100)
                                                    End If
                                                End If
                                            Else
                                                MsgBox " Kh�ng cho ph�p b�n gi� b�ng 0"
                                                .Delete adAffectCurrent
                                                GoTo 2
                                            End If
                                        Else
                                            If blnAutoselect_Price = True Then
                                                If blnPrice = 1 Then
                                                    .Fields("Std_Price1") = rsShowPLU!EV_Price1 + rsShowPLU!EV_Price1 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!EV_Price1 + rsShowPLU!EV_Price1 * PriceRate / 100)
                                                ElseIf blnPrice = 2 Then
                                                    .Fields("Std_Price1") = rsShowPLU!EV_Price2 + rsShowPLU!EV_Price2 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!EV_Price2 + rsShowPLU!EV_Price2 * PriceRate / 100)
                                                Else
                                                    .Fields("Std_Price1") = rsShowPLU!EV_Price3 + rsShowPLU!EV_Price3 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!EV_Price3 + rsShowPLU!EV_Price3 * PriceRate / 100)
                                                End If
                                            Else
                                                If blnPrice = 1 Then
                                                    .Fields("Std_Price1") = rsShowPLU!EV_Price1 + rsShowPLU!EV_Price1 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!EV_Price1 + rsShowPLU!EV_Price1 * PriceRate / 100)
                                                ElseIf blnPrice = 2 Then
                                                    .Fields("Std_Price1") = rsShowPLU!EV_Price2 + rsShowPLU!EV_Price2 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!EV_Price2 + rsShowPLU!EV_Price2 * PriceRate / 100)
                                                Else
                                                    .Fields("Std_Price1") = rsShowPLU!EV_Price3 + rsShowPLU!EV_Price3 * PriceRate / 100
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!EV_Price3 + rsShowPLU!EV_Price3 * PriceRate / 100)
                                                End If
                                            End If
                                        End If
                                        
                                    End If
                                End If
                            End If
                        Else
                            If rsPriceTime.RecordCount > 0 Then
                                rsPriceTime.Find "ID='" & 1 & "'", , adSearchBackward, adBookmarkFirst
                                If Not rsPriceTime.EOF Then
                                    If Format(Now, "HH:mm:ss") >= Format(rsPriceTime.Fields("StartTime"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(rsPriceTime.Fields("EndTime"), "HH:mm:ss") Then
                                        If rsShowPLU!Std_Price1 = 0 Then
                                            If ArrayFlag(rsShowPLU!F3, 6) = 1 Then
                                                If blnAutoselect_Price = True Then
                                                    If blnPrice = 1 Then
                                                        .Fields("Std_Price1") = rsShowPLU!Std_Price1
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price1
                                                    ElseIf blnPrice = 2 Then
                                                        .Fields("Std_Price1") = rsShowPLU!Std_Price2
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price2
                                                    Else
                                                        .Fields("Std_Price1") = rsShowPLU!Std_Price3
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price3
                                                    End If
                                                Else
                                                    If blnPrice = 1 Then
                                                        .Fields("Std_Price1") = rsShowPLU!Std_Price1
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price1
                                                    ElseIf blnPrice = 2 Then
                                                        .Fields("Std_Price1") = rsShowPLU!Std_Price2
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price2
                                                    Else
                                                        .Fields("Std_Price1") = rsShowPLU!Std_Price3
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price3
                                                    End If
                                                End If
                                            Else
                                                MsgBox " Kh�ng cho ph�p b�n gi� b�ng 0"
                                                .Delete adAffectCurrent
                                                GoTo 2
                                            End If
                                        Else
                                        If blnAutoselect_Price = True Then
                                            If blnPrice = 1 Then
                                                .Fields("Std_Price1") = rsShowPLU!Std_Price1
                                                .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price1
                                            ElseIf blnPrice = 2 Then
                                                .Fields("Std_Price1") = rsShowPLU!Std_Price2
                                                .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price2
                                            Else
                                                .Fields("Std_Price1") = rsShowPLU!Std_Price3
                                                .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price3
                                            End If
                                        Else
                                            If blnPrice = 1 Then
                                                .Fields("Std_Price1") = rsShowPLU!Std_Price1
                                                .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price1
                                            ElseIf blnPrice = 2 Then
                                                .Fields("Std_Price1") = rsShowPLU!Std_Price2
                                                .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price2
                                            Else
                                                .Fields("Std_Price1") = rsShowPLU!Std_Price3
                                                .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price3
                                            End If
                                        End If
                                        End If
                                    End If
                                End If
                                rsPriceTime.MoveFirst
                                rsPriceTime.Find "ID='" & 2 & "'", , adSearchForward, adBookmarkFirst
                                If Not rsPriceTime.EOF Then
                                    If Format(Now, "HH:mm:ss") >= Format(rsPriceTime.Fields("StartTime"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(rsPriceTime.Fields("EndTime"), "HH:mm:ss") Then
                                        If rsShowPLU!HH_Price1 = 0 Then
                                            If ArrayFlag(rsShowPLU!F3, 6) = 1 Then
                                                If blnAutoselect_Price = True Then
                                                    If blnPrice = 1 Then
                                                        .Fields("Std_Price1") = rsShowPLU!HH_Price1
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!HH_Price1
                                                    ElseIf blnPrice = 2 Then
                                                        .Fields("Std_Price1") = rsShowPLU!HH_Price2
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!HH_Price2
                                                    Else
                                                        .Fields("Std_Price1") = rsShowPLU!HH_Price3
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!HH_Price3
                                                    End If
                                                Else
                                                    If blnPrice = 1 Then
                                                        .Fields("Std_Price1") = rsShowPLU!HH_Price1
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!HH_Price1
                                                    ElseIf blnPrice = 2 Then
                                                        .Fields("Std_Price1") = rsShowPLU!HH_Price2
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!HH_Price2
                                                    Else
                                                        .Fields("Std_Price1") = rsShowPLU!HH_Price3
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!HH_Price3
                                                    End If
                                                End If
                                            Else
                                                MsgBox " Kh�ng cho ph�p b�n gi� b�ng 0"
                                                .Delete adAffectCurrent
                                                GoTo 2
                                            End If
                                        Else
                                            If blnAutoselect_Price = True Then
                                                If blnPrice = 1 Then
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price1
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!HH_Price1
                                                ElseIf blnPrice = 2 Then
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price2
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!HH_Price2
                                                Else
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price3
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!HH_Price3
                                                End If
                                            Else
                                                If blnPrice = 1 Then
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price1
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!HH_Price1
                                                ElseIf blnPrice = 2 Then
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price2
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!HH_Price2
                                                Else
                                                    .Fields("Std_Price1") = rsShowPLU!HH_Price3
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!HH_Price3
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                                rsPriceTime.MoveFirst
                                rsPriceTime.Find "ID='" & 3 & "'", , adSearchForward, adBookmarkFirst
                                If Not rsPriceTime.EOF Then
                                    If Format(Now, "HH:mm:ss") >= Format(rsPriceTime.Fields("StartTime"), "HH:mm:ss") And Format(Now, "HH:mm:ss") <= Format(rsPriceTime.Fields("EndTime"), "HH:mm:ss") Then
                                        If rsShowPLU!EV_Price1 = 0 Then
                                            If ArrayFlag(rsShowPLU!F3, 6) = 1 Then
                                                If blnAutoselect_Price = True Then
                                                    If blnPrice = 1 Then
                                                        .Fields("Std_Price1") = rsShowPLU!EV_Price1
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!EV_Price1
                                                    ElseIf blnPrice = 2 Then
                                                        .Fields("Std_Price1") = rsShowPLU!EV_Price2
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!EV_Price2
                                                    Else
                                                        .Fields("Std_Price1") = rsShowPLU!EV_Price3
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!EV_Price3
                                                    End If
                                                Else
                                                    If blnPrice = 1 Then
                                                        .Fields("Std_Price1") = rsShowPLU!EV_Price1
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!EV_Price1
                                                    ElseIf blnPrice = 2 Then
                                                        .Fields("Std_Price1") = rsShowPLU!EV_Price2
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!EV_Price2
                                                    Else
                                                        .Fields("Std_Price1") = rsShowPLU!EV_Price3
                                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!EV_Price3
                                                    End If
                                                End If
                                            Else
                                                MsgBox " Kh�ng cho ph�p b�n gi� b�ng 0"
                                                .Delete adAffectCurrent
                                                GoTo 2
                                            End If
                                        Else
                                            If blnAutoselect_Price = True Then
                                                If blnPrice = 1 Then
                                                    .Fields("Std_Price1") = rsShowPLU!EV_Price1
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!EV_Price1
                                                ElseIf blnPrice = 2 Then
                                                    .Fields("Std_Price1") = rsShowPLU!EV_Price2
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!EV_Price2
                                                Else
                                                    .Fields("Std_Price1") = rsShowPLU!EV_Price3
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!EV_Price3
                                                End If
                                            Else
                                                If blnPrice = 1 Then
                                                    .Fields("Std_Price1") = rsShowPLU!EV_Price1
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!EV_Price1
                                                ElseIf blnPrice = 2 Then
                                                    .Fields("Std_Price1") = rsShowPLU!EV_Price2
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!EV_Price2
                                                Else
                                                    .Fields("Std_Price1") = rsShowPLU!EV_Price3
                                                    .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!EV_Price3
                                                End If
                                            End If
                                        End If
                                        
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If LocationFlag = 1 Then
                            If rsShowPLU!Std_Price1 = 0 Then
                                If ArrayFlag(rsShowPLU!F3, 6) = 1 Then
                                    If blnAutoselect_Price = True Then
                                        If blnPrice = 1 Then
                                            .Fields("Std_Price1") = rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100
                                            .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100)
                                        ElseIf blnPrice = 2 Then
                                            .Fields("Std_Price1") = rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100
                                            .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100)
                                        Else
                                            .Fields("Std_Price1") = rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100
                                            .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100)
                                        End If
                                    Else
                                        If blnPrice = 1 Then
                                            .Fields("Std_Price1") = rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100
                                            .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100)
                                        ElseIf blnPrice = 2 Then
                                            .Fields("Std_Price1") = rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100
                                            .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100)
                                        Else
                                            .Fields("Std_Price1") = rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100
                                            .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100)
                                        End If
                                    End If
                                Else
                                    MsgBox " Kh�ng cho ph�p b�n gi� b�ng 0"
                                    .Delete adAffectCurrent
                                    GoTo 2
                                End If
                            Else
                                If blnAutoselect_Price = True Then
                                    If blnPrice = 1 Then
                                        .Fields("Std_Price1") = rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100
                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100)
                                    ElseIf blnPrice = 2 Then
                                        .Fields("Std_Price1") = rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100
                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100)
                                    Else
                                        .Fields("Std_Price1") = rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100
                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100)
                                    End If
                                Else
                                    If blnPrice = 1 Then
                                        .Fields("Std_Price1") = rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100
                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price1 + rsShowPLU!Std_Price1 * PriceRate / 100)
                                    ElseIf blnPrice = 2 Then
                                        .Fields("Std_Price1") = rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100
                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price2 + rsShowPLU!Std_Price2 * PriceRate / 100)
                                    Else
                                        .Fields("Std_Price1") = rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100
                                        .Fields("Amt") = .Fields("Amt") + ConQty * (rsShowPLU!Std_Price3 + rsShowPLU!Std_Price3 * PriceRate / 100)
                                    End If
                                End If
                            End If
                        Else
                            If CDbl("0" & rsShowPLU!Std_Price1) = 0 Then
                                If ArrayFlag(rsShowPLU!F3, 6) = 1 Then
                                    If blnAutoselect_Price = True Then
                                        If blnPrice = 1 Then
                                            .Fields("Std_Price1") = rsShowPLU!Std_Price1
                                            .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price1
                                        ElseIf blnPrice = 2 Then
                                            .Fields("Std_Price1") = rsShowPLU!Std_Price2
                                            .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price2
                                        Else
                                            .Fields("Std_Price1") = rsShowPLU!Std_Price3
                                            .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price3
                                        End If
                                    Else
                                        If blnPrice = 1 Then
                                            .Fields("Std_Price1") = rsShowPLU!Std_Price1
                                            .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price1
                                        ElseIf blnPrice = 2 Then
                                            .Fields("Std_Price1") = rsShowPLU!Std_Price2
                                            .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price2
                                        Else
                                            .Fields("Std_Price1") = rsShowPLU!Std_Price3
                                            .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price3
                                        End If
                                    End If
                                Else
                                    MsgBox " Kh�ng cho ph�p b�n gi� b�ng 0"
                                    .Delete adAffectCurrent
                                    GoTo 2
                                End If
                            Else
                                If blnAutoselect_Price = True Then
                                    If blnPrice = 1 Then
                                        .Fields("Std_Price1") = rsShowPLU!Std_Price1
                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price1
                                    ElseIf blnPrice = 2 Then
                                        .Fields("Std_Price1") = rsShowPLU!Std_Price2
                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price2
                                    Else
                                        .Fields("Std_Price1") = rsShowPLU!Std_Price3
                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price3
                                    End If
                                Else
                                    If blnPrice = 1 Then
                                        .Fields("Std_Price1") = rsShowPLU!Std_Price1
                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price1
                                    ElseIf blnPrice = 2 Then
                                        .Fields("Std_Price1") = rsShowPLU!Std_Price2
                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price2
                                    Else
                                        .Fields("Std_Price1") = rsShowPLU!Std_Price3
                                        .Fields("Amt") = .Fields("Amt") + ConQty * rsShowPLU!Std_Price3
                                    End If
                                End If
                            End If
                            
                        End If
                    End If
                Else
                    If ExtrasPrice = 0 Then
                        If ArrayFlag(rsShowPLU!F3, 6) = 1 Then
                            .Fields("Std_Price1") = ExtrasPrice
                            .Fields("Amt") = .Fields("Amt") + ConQty * ExtrasPrice
                        Else
                            MsgBox " Kh�ng cho ph�p b�n gi� b�ng 0"
                            .Delete adAffectCurrent
                            GoTo 2
                        End If
                    Else
                        If ArrayFlag(rsShowPLU.Fields("F3"), 7) = 1 Then ExtrasPrice = -ExtrasPrice
                        .Fields("Std_Price1") = ExtrasPrice
                        .Fields("Amt") = .Fields("Amt") + ConQty * ExtrasPrice
                    End If
                End If
                Dim strLine As String
                strLine = .Fields("PluName") & Space(3) & .Fields("Qty") & Space(1) & Format(.Fields("Amt"), "#,##0")
                .Fields("Line_Disc") = 0
                .Fields("Line_Disc_Desc") = ""
               
                ' Ghi thong tin mon order xuong file log
                'Print #fFile, vbTab & .Fields("PluNo") & vbTab & .Fields("PluName") & vbTab & .Fields("Qty") & vbTab & .Fields("Std_Price1") & vbTab & .Fields("Qty") * .Fields("Std_Price1") & vbTab & Now
                .Update
                'If ArrayFlag(SF(6), 4) = 1 Then Call Display_Sale(strLine)
            End With
            
2:            Call SetFLGRIDORDER(rsTemp)
            ConQty = 1
            blnEditQty = False
            txtQty.Text = ""
            isExtrasPrice = False
'            LineDelete = ""
            ExtrasPrice = 0
            If Not blnAutoselect_Price Then
                blnPrice = 1
            End If
            'SetColorFlexGrid flgOrder, 1, 1, flgOrder.Cols
            'lblTotalAmt.Caption = Format(TotalAmt - TotalAmt * Discount / 100 + tota, formatNum)
             If ArrayFlag(SF(6), 4) = 1 Then Call Display_Sale("", lblTotalAmt.Caption)
             
        End If
    Else
        With frmItem_Details
            .Get_Item_Code = cmdSub(Index).Tag
            .Show vbModal
        End With
       Call cmdBtn_Click(Dept_Index)
    End If
    Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdSub_Click" & vbCrLf
    MsgBox Err.Number & Err.Description & Me.name & "   cmdSub_Click"
    Item_Order_State = True
End Sub

Private Sub cmdTachmon_Click()
'Exit Sub
On Error GoTo Handle
Dim OK As Boolean
fraEdit.Visible = False
Dim ID As String
    
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "split_items") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "split_items") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
       GoTo OK
    End If
OK:

iset = False
If check_IsPrint(lblBillNo.Caption) Then Exit Sub
If fClick = False Then
    MsgBox "B�n ph�i ch�n m�n c�n chuy�n �i!"
    Exit Sub
End If
    With rsTranfer
        If .State = 0 Then
            .Fields.Append "PLUNo", adVarWChar, 20
            .Fields.Append "PLUName", adVarWChar, 100
            .Fields.Append "Qty", adDouble
            .Fields.Append "Std_Price1", adDouble
            .Fields.Append "Amt", adDouble
            .Fields.Append "F2", adVarWChar, 2
            .Fields.Append "Cashier_ID", adVarWChar, 25
            .Fields.Append "DateTime", adVarWChar, 50
            .Fields.Append "Ordered", adBoolean
            '.Fields.Append "Resend", adBoolean
            .Fields("Ordered").Attributes = adColNullable
            .Fields.Append "Reason", adVarWChar, 200
            .Fields("Reason").Attributes = adColNullable
            .Fields.Append "Kit_Desc", adVarWChar, 250
            .Fields("Kit_Desc").Attributes = adColNullable
            .Fields.Append "Line_Disc", adDouble
            .Fields.Append "Line_Disc_Desc", adVarWChar, 250
            .Fields("Line_Disc_Desc").Attributes = adColNullable
            .Open
        End If
    End With
    If rslinedelete.State <> 0 Then
        rslinedelete.MoveFirst
    Else
        Exit Sub
    End If
    If rsTemp.RecordCount > 0 Then
    With rsTemp
    iset = False
        Do While Not rslinedelete.EOF
            .Find "Line_Number=" & rslinedelete.Fields("Line"), , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If CDbl("0" & .Fields("Qty")) > 1 Then
                    With frmQtyTranfer
                        .Show vbModal
                        OK = .GetOK
                    End With
                    If OK = False Then Exit Sub
                    If qtyTran > .Fields("Qty") Then
                        MsgBox "Kh�ng cho ph�p s� l��ng chuy�n l�n h�n hi�n t�i", vbInformation
                        Exit Sub
                    ElseIf qtyTran = .Fields("Qty") Then
                        rsTranfer.addNew
                        rsTranfer!PluNo = .Fields("PluNo")
                        rsTranfer!PluName = .Fields("PluName")
                        rsTranfer!Qty = qtyTran
                        rsTranfer!Std_Price1 = .Fields("Std_Price1")
                        rsTranfer!Amt = .Fields("Amt")
                        rsTranfer!Kit_Desc = .Fields("Kit_Desc")
                        rsTranfer!Line_Disc = .Fields("Line_Disc")
                        rsTranfer!Line_Disc_Desc = .Fields("Line_Disc_Desc")
                        'rsTranfer!Resend = .Fields("Resend")
                        rsInventory.Find "ItemNum='" & .Fields("PluNo") & "'", , adSearchForward, adBookmarkFirst
                        If Not rsInventory.EOF Then
                            rsTranfer!F2 = rsInventory.Fields("F2")
                        End If
                        rsTranfer!cashier_ID = UserID
                        rsTranfer!DateTime = DateDefault & Format(Now, "HH:mm:ss")
                        If .Fields("Status") = True Then
                            rsTranfer!Ordered = 1
                        End If
                        rsTranfer.Update
                        .Delete adAffectCurrent
                        GoTo 2
                    Else
                        .Fields("Qty") = .Fields("qty") - qtyTran
                        .Update
                    End If
                    'Update tranfer
                    rsTranfer.addNew
                    rsTranfer!PluNo = .Fields("PluNo")
                    rsTranfer!PluName = .Fields("PluName")
                    rsTranfer!Qty = qtyTran
                    rsTranfer!Std_Price1 = .Fields("Std_Price1")
                    rsTranfer!Amt = .Fields("Amt")
                    'rsTranfer!Resend = True
                    rsTranfer!Kit_Desc = .Fields("Kit_Desc")
                    rsTranfer!Line_Disc = .Fields("Line_Disc")
                    rsTranfer!Line_Disc_Desc = .Fields("Line_Disc_Desc")
                    
                    rsInventory.Find "ItemNum='" & .Fields("PluNo") & "'", , adSearchForward, adBookmarkFirst
                    If Not rsInventory.EOF Then
                        rsTranfer!F2 = rsInventory.Fields("F2")
                    End If
                    rsTranfer!cashier_ID = UserID
                    rsTranfer!DateTime = DateDefault & Format(Now, "HH:mm:ss")
                    If .Fields("Status") = True Then
                        rsTranfer!Ordered = 1
                    End If
                    rsTranfer.Update
                    GoTo 2
                Else
                    rsTranfer.addNew
                    rsTranfer.Fields("PluNo") = .Fields("PluNo")
                    rsTranfer.Fields("PluName") = .Fields("PluName")
                    rsTranfer.Fields("Qty") = .Fields("Qty")
                    rsTranfer.Fields("Std_Price1") = .Fields("Std_Price1")
                    rsTranfer.Fields("Amt") = .Fields("Amt")
                    rsTranfer.Fields("Kit_Desc") = .Fields("Kit_Desc")
                    rsTranfer.Fields("Line_Disc") = .Fields("Line_Disc")
                    rsTranfer.Fields("Line_Disc_Desc") = .Fields("Line_Disc_Desc")
                    rsInventory.Find "ItemNum='" & .Fields("PluNo") & "'", , adSearchForward, adBookmarkFirst
                    If Not rsInventory.EOF Then
                        rsTranfer.Fields("F2") = rsInventory.Fields("F2")
                    End If
                    rsTranfer.Fields("Cashier_ID") = UserID
                    rsTranfer.Fields("DateTime") = DateDefault & Format(Now, "HH:mm:ss")
                    If .Fields("Status") = True Then
                        rsTranfer.Fields("Ordered") = 1
                    End If
                    rsTranfer.Update
                .Delete adAffectCurrent
                End If
            End If
            rslinedelete.MoveNext
        Loop
        Set rslinedelete = Nothing
        LineDelete = ""
   End With
End If
2:
currentBill = lblBillNo.Caption
    cmdNewBalance_Click
    With frmTablePlan
        .GetBillTranfer = CDbl(currentBill)
        .FormState = 4
        .GetLocationTranfer = Sec_ID
        .GetTableTranfer = Table_ID
    End With
Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " cmdTachmon_Click" & vbCrLf
    MsgBox Err.Number & Err.Description & Me.name & " cmdTachmon_Click"
End Sub

Private Sub cmdTranferTable_Click()
On Error GoTo Handle
 Dim Table As String
 'options 19-12-2013
 Dim ID As String
    
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "tabletranffer") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "tabletranffer") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
'end options

 Picwait.Visible = True
 Me.Enabled = False
 If MsgBox("B�n c� ch�c ch�n chuy�n b�n kh�ng???", vbYesNo) = vbYes Then
    currentBill = lblBillNo.Caption
    Table = lblTableNo.Caption
    Call cmdNewBalance_Click
    With frmTablePlan
        .FormState = 2
        .GetLocationTranfer = Sec_ID
        .GetTableTranfer = Table
        .GetBillTranfer = CDbl(currentBill)
    End With
Else
    Picwait.Visible = False
    Me.Enabled = True
End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  cmdTranferTable_Click"
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
    MsgBox Err.Number & Err.Description & " - " & "Vui l�ng ��i gi�y l�t �� load d� li�u"
End Sub

Private Sub cmdVAT_Click()
 On Error GoTo Handle
    fraEdit.Visible = False
    Dim ID As String
    iset = False
    If check_IsPrint(lblBillNo.Caption) Then
        If ArrayFlag(SF(4), 8) = 0 Then Exit Sub
            If UserLevel = 1 Then AllowDelete = True
            If Not AllowDelete Then
                With frmPassword
                    .FormActionKey = "Others"
                    .Show vbModal
                    ID = .return_Pass
                    If Not .Return_right Then Exit Sub
                End With
                If Return_right(ID, "Delete") Or UCase(ID) = "131112" Then AllowDelete = True
            Else
                ID = UserID
            End If
             GoTo 1
    Else
        AllowDelete = True
1:
        If AllowDelete = False Then Exit Sub
            With frmPhimso
                .lblTitle.Caption = "Nh�p % VAT:"
                .FormCall = 3
                .Show vbModal
                VAT = .Return_Value
            End With
    '        'Print #fFile, "Thu� VAT:" & VAT & "%" & vbTab & Now
            If VAT > 100 Then
                MsgBox "Thu� VAT kh�ng ���c l�n h�n 100%"
                VAT = 0
            End If
            lblTotalAmt.Caption = Format(CDbl(lblTotalAmt.Caption) + CDbl(lblTotalAmt.Caption) * VAT / 100, formatNum)
        End If
        AllowDelete = False
    Exit Sub
Handle:
'    'Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " VAT_Click" & vbCrLf
    MsgBox Err.Number & Err.Description & Me.name & "  VAT_Click"
End Sub

Private Sub cmdVoidTran_Click()
On Error GoTo Handle
    iset = False
    fraEdit.Visible = False
    With frmPhimso
        .lblTitle.Caption = "Nh�p s� kh�ch:"
        .FormCall = 3
        .Show vbModal
        Personal = .Return_Value
    End With
    With rsInvoice_Total
        .Find "Invoice_Number=" & lblBillNo.Caption, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Fields("Personals") = Val("0" & .Fields("Personals")) + Personal
            .Update
        End If
        
    End With
Exit Sub
Handle:
    MsgBox Err.Description & Me.name & " " & "cmdVoidTran_Click"
End Sub

Private Sub cmdItemInfor_Click()
Dim ID As String
    
    If UserLevel = 1 Or UserID = "131112" Then GoTo OK
    
    If Not Get_Right(UserID, "item_infor") Then
        With frmPassword
            .FormActionKey = "Others"
            .Show vbModal
            ID = .return_Pass
            If Not .Return_right Then Exit Sub
        End With
        If Get_Right(ID, "item_infor") Then
            GoTo OK
        Else
            Exit Sub
        End If
    Else
        GoTo OK
    End If
OK:
If cmdItemInfor.Caption = "Hi�u ch�nh m�n" Then
    Item_Order_State = False
    cmdItemInfor.BackColor = &HFF00&
    cmdItemInfor.Caption = "Ho�n t�t"
ElseIf cmdItemInfor.Caption = "Ho�n t�t" Then
    Item_Order_State = True
    cmdItemInfor.BackColor = &HFF&
    cmdItemInfor.Caption = "Hi�u ch�nh m�n"
End If
End Sub

Private Sub CommandButton1_Click()
    fraEdit.Visible = False
    frmReserve.Show vbModal
End Sub





Private Sub dtgFind_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    Call txtQty_KeyPress(14)
    If KeyAscii = 27 Then
        dtgFind.Visible = False
    ElseIf KeyAscii = 13 Then
    LineNum = LineNum + 1
        With rsFind
            If .RecordCount = 0 Then
                dtgFind.Visible = False
                Exit Sub
            End If
             With rsTemp
                If .State = 0 Then
                    .Fields.Append "TableNo", adVarWChar, 50
                    .Fields.Append "Sec_No", adVarWChar, 2
                    .Fields.Append "Line_Number", adDouble
                    .Fields.Append "Dept_ID", adVarWChar, 3
                    .Fields.Append "PLUNo", adVarWChar, 20
                    .Fields.Append "PLUName", adVarWChar, 100
                    .Fields.Append "Qty", adDouble
                    .Fields.Append "Std_Price1", adDouble
                    .Fields.Append "Amt", adDouble
                    .Fields.Append "F1", adVarWChar, 2
                    .Fields.Append "F2", adVarWChar, 2
                    .Fields.Append "F3", adVarWChar, 2
                    .Fields.Append "Resend", adBoolean
                    .Fields.Append "Status", adBoolean
                    .Fields.Append "Quanburned", adDouble
                    .Fields.Append "Kit_Desc", adVarWChar, 250
                    .Fields("Kit_Desc").Attributes = adColNullable
                    .Fields.Append "Modifier_No", adVarWChar, 250
                    .Fields("Modifier_No").Attributes = adColNullable
                    .Fields.Append "Line_Disc", adDouble
                    .Fields.Append "Line_Disc_Desc", adVarWChar, 250
                    .Fields("Line_Disc_Desc").Attributes = adColNullable
                    .Fields.Append "TimeOrder", adVarWChar, 250
                    .Fields("TimeOrder").Attributes = adColNullable
                    .Open
                End If
                .addNew
                .Fields("Sec_No") = Sec_ID
                .Fields("TableNo") = Table_ID
                .Fields("Line_Number") = LineNum
                .Fields("PluNo") = rsFind.Fields("ItemNum")
                .Fields("PluName") = rsFind.Fields("ItemName")
                .Fields("Qty") = ConQty
                .Fields("F1") = rsShowPLU!F1
                .Fields("F2") = rsShowPLU!F2
                .Fields("F3") = rsShowPLU!F3
                .Fields("Dept_ID") = rsShowPLU!Dept_ID
                .Fields("Status") = False
                .Fields("Resend") = 0
                If blnPrice = 1 Then
                    .Fields("Std_Price1") = rsFind.Fields("Std_Price1")
                    .Fields("Amt") = .Fields("Amt") + ConQty * rsFind!Std_Price1
                ElseIf blnPrice = 2 Then
                    .Fields("Std_Price1") = rsFind.Fields("Std_Price2")
                    .Fields("Amt") = .Fields("Amt") + ConQty * rsFind!Std_Price2
                ElseIf blnPrice = 3 Then
                     .Fields("Std_Price1") = rsFind.Fields("Std_Price3")
                    .Fields("Amt") = .Fields("Amt") + ConQty * rsFind!Std_Price3
                End If
                .Update
            End With
        End With
    ElseIf KeyAscii = 9 Then
        dtgFind.Visible = False
    End If
    Call SetFLGRIDORDER(rsTemp)
    dtgFind.Visible = False
    Delay (100)
    dtgFind.Visible = False
    txtSearch.Text = ""
    txtQty.SetFocus
    ConQty = 1
    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " "
End Sub


Private Sub flgOrder_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuList, 0
    End If
End Sub


Private Sub Label3_Click()
    Label3.BackColor = vbYellow
    lblDiscount.BackColor = vbYellow
End Sub

Private Sub Label4_Click()
    Call cmdAdjustment1_Click
End Sub

Private Sub Label6_Click()
    Call cmdAdjustment2_Click
End Sub

Private Sub lblDiscount_Click()
    Label3.BackColor = vbYellow
    lblDiscount.BackColor = vbYellow
End Sub

Private Sub lblTotalAmt_Click()
On Error GoTo Handle
    With frmTotalDetails
        .Get_Adj1Per = Adj1
        .Get_Adj1 = Adjtotal1
        .Get_Adj2 = Adjtotal2
        .Get_Adj2Per = Adj2
        .Get_DiscountPer = Discount
        .Get_Money = MoneyAmount
        .Get_Sercharge = service_Charge
        .Get_Total = TotalAmt
        .Get_Table = Table_ID
        .Get_Bill = lblBillNo.Caption
        .Get_VAT = VAT
        .Show vbModal
    End With
Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & userName
    MsgBox Err.Number & Err.Description & Me.name
End Sub

Private Sub mnuClose_Click()
    Call cmdNewBalance_Click
End Sub

Private Sub mnuDetails_Click()
On Error GoTo Handle
    With frmDetailsOrder
        .Get_Recordset = rsTemp
        .LetBill = lblBillNo.Caption
        .LetTable = Table_ID
        .Show vbModal
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " mnuDetails_Click"
End Sub

'Private Sub MSCom_OnComm()
'On Error GoTo Handle
'Dim EventMsg$, ErrorMsg$
'Select Case MSCom.CommEvent
'    Case comEvReceive
'        Dim buffer As Variant
'        buffer = MSCom.input
'        Display_Sale ("",buffer)
'        EventMsg$ = "Receive"
'    Case comEvSend
'        EventMsg$ = "Send"
'    Case comEvCTS
'        EventMsg$ = "Change in CTS Detected"
'    Case comEvDSR
'        EventMsg$ = "Change in DSR Detected"
'    Case comEvCD
'        EventMsg$ = "Change in CD Detected"
'    Case comEvRing
'        EventMsg$ = "The Phone is Ringing"
'    Case comEvEOF
'        EventMsg$ = "End Of File Detected"
'    Case comBreak
'        EventMsg$ = "Break received"
'    Case comCDTO
'        EventMsg$ = "Carrier Detect Time Out"
'    Case comCTSTO
'        EventMsg$ = "CTS Time Out"
'    Case comDCB
'        EventMsg$ = "Error retrieving DCB"
'    Case comDSRTO
'        EventMsg$ = "DSR TimeOut"
'    Case comFrame
'        EventMsg$ = "Framing Error"
'    Case comOverrun
'        EventMsg$ = "Over Run Error"
'    Case comRxOver
'        EventMsg$ = "Receive Buffer Overflow"
'    Case comRxParity
'        EventMsg$ = "Parity Error"
'    Case comTxFull
'        EventMsg$ = "Transmit Buffer Full"
'    Case Else
'        EventMsg$ = "Unknown Error or Event"
'
'End Select
'
'Exit Sub
'Handle:
'    MsgBox Err.Number & Err.Description & Me.name
'End Sub
Private Sub Price1_Click()
    On Error GoTo Handle
        blnPrice = 1
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - "
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
        fClick = True
    Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " flgOrder_Click" & vbCrLf
    MsgBox Err.Number & Err.Description & Me.name & " flgOrder_Click"
End Sub

Private Sub Form_Activate()
 On Error GoTo Handle
        Dim ctrl As Control
        iset = True
'        If isLoaded = True Then Exit Sub
'        isLoaded = True
'        If cmdOtherPayment.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
        Desarr = LoadLanguage(LngFile, "#01:007:")
        'Me.Caption = Desarr(23)
        For Each ctrl In Me
        DoEvents
            If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = Desarr(Mid(ctrl.Tag, 2))
        Next ctrl
        lblCustomer.Caption = CustNo(1)
        'lblCustBalance.Caption = CustNo(2)
        
        lblPersonNum.Caption = Personal
        
        lblDiscount.Caption = Discount & "%"
        lblCustomer.Caption = CustNo(1)
        lblAdj1.Caption = Adj1 & "%"
        lblAdj2.Caption = Adj2 & "%"
        Set rsInvoice_Total = OpenCriticalTable("select * from Invoice_totals", cnData)
        With rsInvoice_Total
            .Find "Invoice_Number=" & lblBillNo.Caption, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                Dim rsuser As New ADODB.Recordset
                Set rsuser = Open_Table(cnData, "Employee")
                With rsuser
                    .Find "Cashier_ID='" & rsInvoice_Total.Fields("orderMan") & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        lblCashierName.Caption = .Fields("EmpName")
                    Else
                        lblCashierName.Caption = userName
                    End If
                End With
            End If
        
        End With
        
'        If UserLevel <> 1 Then Call CheckRight
    
    Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  Form_Activate"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        frmHelp.Show vbModal
    ElseIf KeyCode = vbKeyF2 Then
        Call cmdNewBalance_Click
    ElseIf KeyCode = vbKeyF3 And cmdBufferPrint.Enabled = True Then
        Call cmdBufferPrint_Click
    ElseIf KeyCode = vbKeyF4 And cmdOtherPayment.Enabled = True Then
        Call cmdOtherPayment_Click
    ElseIf KeyCode = vbKeyDelete And cmddelete.Enabled = True Then
        Call cmdDelete_Click
    ElseIf KeyCode = vbKeyPageUp Then
        Call cmdUp_Click
    ElseIf KeyCode = vbKeyPageDown Then
        Call cmdDown_Click
    End If
    If Shift = 2 Then
        If KeyCode = vbKeyF Then
            txtSearch.Text = ""
            txtSearch.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
    Dim i  As Integer
    Dim LineAmount_Discount As Double
    Item_Order_State = True
    MeUnload = False
    ConQty = 1
    LineAmount_Discount = 0
    LineNum = 0
    isExtrasPrice = False
    blnEditQty = False
    strLast = ""
    ' Mo com Customer_Display
    If ArrayFlag(SF(6), 4) = 1 Then Call Open_Port
'    Set rsTemp = Nothing
    Set rsTemp = New ADODB.Recordset
    lblAutoConsolidate = False
    Desarr = LoadLanguage(LngFile, "#01:007:")
   
        If cnData.State = 0 Then
            Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
        End If
        If cnData.State <> 0 Then
            Set rsInventory = Open_Table(cnData, "Inventory")
            Set rsDepartment = OpenCriticalTable("SELECT Index,Dept_ID, Description,ColorDept from Departments order by Index;", cnData)
            Set rsInvoice_Total = Open_Table(cnData, "Invoice_Totals")
            Set rsInvoice_Items = Open_Table(cnData, "Invoice_Itemized")
            Set rsLocation = Open_Table(cnData, "Table_Diagram_Sections")
            Set rsPriceTime = Open_Table(cnData, "PeriodPrice")
            Set rsSystem = Open_Table(cnData, "SystemFlag")
            Set rsInvoice_Note = Open_Table(cnData, "Invoice_Totals_Notes")

        End If

        'Can modify 18/11/2007
        'lay muc gia quy dinh va ty le gia' gia tang theo khu vuc
        Call Get_Charge(strBill)
        Call GetAutoPrice
        'modify 07/03/2011 : lay giam tu dong
        If ArrayFlag(SF(4), 2) = 1 Then
            Call get_Discount_Auto
        End If
        'end
        
        ReDim Preserve ArrCommand(rsDepartment.RecordCount)
        Do While Not rsDepartment.EOF
        DoEvents
            'ArrCommand(i) = rsDepartment.NextRecordset(11)
            ArrCommand(i) = rsDepartment.Fields("Index")
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
                Call cmdBtn_Click(rsDepartment.Fields("Index"))
            End If
        End If
        
        Call Set_flgOrder
        
        lblTableNo.Caption = Table_ID
        If Table_ID = "" Then Exit Sub
        With rsLocation
            .Find "Location_ID='" & Sec_ID & "'", , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    lblStationName.Caption = .Fields("Section_ID")
                Else
                    lblStationName.Caption = Sec_ID
                End If
        End With
        
        
       
        If currentBill = "" Then
            currentBill = CDbl(GetSettingStr("SYSTEM", "MaxInvoice", True, myIniFile)) + 1
            SaveSettingStr "SYSTEM", "MaxInvoice", currentBill, myIniFile
        End If
        
        lblBillNo.Caption = strBill
        
        Dim strBalance As String
        If ArrayFlag(SF(6), 1) = 1 Then
            strBalance = "SELECT ItemNum AS PluNo, sum(Quantity) AS Qty,LineNum," & _
                        " PricePer AS Std_Price1, DiffItemName AS PluName," & _
                        " Kit_Description as Kit_Desc,LineDisc,Line_Disc_Desc,TimeOrder " & _
                        " From Invoice_Itemized" & _
                        " WHERE Invoice_Number=" & strBill & _
                        " group by ItemNum, PricePer , DiffItemName,Kit_Description," & _
                        " LineDisc,LineNum,Line_Disc_Desc,TimeOrder" & _
                        " ORDER BY LineNum"

        Else
            strBalance = "SELECT ItemNum AS PluNo, sum(Quantity) AS Qty, PricePer AS Std_Price1," & _
                        " DiffItemName AS PluName," & _
                        " Kit_Description as Kit_Desc,LineDisc,Line_Disc_Desc  " & _
                        " From Invoice_Itemized" & _
                        " WHERE Invoice_Number=" & strBill & _
                        " group by ItemNum, PricePer , DiffItemName," & _
                        " Kit_Description,LineDisc,Line_Disc_Desc " & _
                        " ORDER BY ItemNum"
        End If
        Set rsNew = OpenCriticalTable(strBalance, cnData)
        Dim rsTempDelete As New ADODB.Recordset
       
       If rsNew.RecordCount > 0 Then
       i = 1
       Dim SubTotal As Double
       Do While Not rsNew.EOF
       DoEvents
        With rsTemp
            If .State = 0 Then
                .Fields.Append "TableNo", adVarWChar, 50
                .Fields.Append "Sec_No", adVarWChar, 2
                .Fields.Append "Line_Number", adDouble
                .Fields.Append "Dept_ID", adVarWChar, 3
                .Fields.Append "PLUNo", adVarWChar, 20
                .Fields.Append "PLUName", adVarWChar, 100
                .Fields.Append "Qty", adDouble
                .Fields.Append "Std_Price1", adDouble
                .Fields.Append "Amt", adDouble
                .Fields.Append "F1", adVarWChar, 2
                .Fields.Append "F2", adVarWChar, 2
                .Fields.Append "F3", adVarWChar, 2
                .Fields.Append "Resend", adBoolean
                .Fields.Append "Status", adBoolean
                .Fields.Append "Quanburned", adDouble
                .Fields.Append "Kit_Desc", adVarWChar, 250
                .Fields("Kit_Desc").Attributes = adColNullable
                .Fields.Append "Modifier_No", adVarWChar, 250
                .Fields("Modifier_No").Attributes = adColNullable
                .Fields.Append "Line_Disc", adDouble
                .Fields.Append "Line_Disc_Desc", adVarWChar, 250
                .Fields.Append "TimeOrder", adVarWChar, 250
                .Fields("TimeOrder").Attributes = adColNullable
                .Open
            End If
            .addNew
            .Fields("Sec_No") = Sec_ID
            .Fields("TableNo") = Table_ID
            If ArrayFlag(SF(6), 1) = 1 Then
                .Fields("Line_Number") = rsNew!LineNum
            Else
                .Fields("Line_Number") = LineNum
            End If
            .Fields("PluNo") = rsNew!PluNo
            .Fields("PluName") = rsNew!PluName
            .Fields("Qty") = rsNew!Qty
            .Fields("Std_Price1") = rsNew!Std_Price1
            .Fields("Amt") = rsNew!Qty * rsNew!Std_Price1
            .Fields("Amt") = .Fields("Amt") - .Fields("Amt") * rsNew!LineDisc / 100
            .Fields("Status") = 1
            .Fields("Resend") = 0
            .Fields("Quanburned") = rsNew!Qty
            .Fields("Kit_Desc") = " " & rsNew!Kit_Desc
            .Fields("Line_Disc") = " " & rsNew!LineDisc
            .Fields("Line_Disc_Desc") = " " & rsNew!Line_Disc_Desc
            .Fields("TimeOrder") = " " '& rsNew!TimeOrder
            LineAmount_Discount = rsNew.Fields("LineDisc") * rsNew.Fields("Qty") * rsNew.Fields("Std_Price1") / 100
            rsInventory.Find "ItemNum='" & rsNew!PluNo & "'", , adSearchForward, adBookmarkFirst
            If Not rsInventory.EOF Then
                .Fields("F1") = rsInventory!F1
                .Fields("F2") = rsInventory!F2
                .Fields("F3") = rsInventory!F3
                .Fields("Dept_ID") = rsInventory!Dept_ID
                .Fields("Modifier_No") = rsInventory!Modify_Number
            End If
            .Update
        End With
        SubTotal = SubTotal + CDbl(rsNew!Qty * rsNew!Std_Price1) - LineAmount_Discount
        LineNum = LineNum + 1
        rsNew.MoveNext
        i = i + 1
        Loop
            Call SetFLGRIDORDER(rsTemp)
        End If
        'Giam di phan Discount
        If Discount > 0 Then
            SubTotal = SubTotal - SubTotal * Discount / 100
        End If

        Call Get_AdjValue(strBill)
        
        SubTotal = SubTotal + SubTotal * service_Charge / 100
        SubTotal = SubTotal + Adjtotal1 + Adjtotal2 + Adjtotal3 + Adjtotal4 + MoneyAmount
        SubTotal = SubTotal + SubTotal * VAT / 100
        lblTotalAmt.Caption = Format(SubTotal, formatNum)
        Call GetAllowChangPrice
        If ArrayFlag(SF(6), 4) = 1 Then Call Display_Sale("", Format(lblTotalAmt.Caption, "#,##0"))
    Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " Form_Load" & vbCrLf
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub

Private Sub cmdSoluong_Click()
    On Error GoTo Handle
        iset = False
        With frmPhimso
            .lblTitle.Caption = "Nh�p s� l��ng:"
            .FormCall = 1
            .Show vbModal
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdSoluong_Click"
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
            .TextMatrix(0, 1) = Desarr(19) '"T�n m�n"
            .TextMatrix(0, 2) = Desarr(20) ' "S�' luong"
            .TextMatrix(0, 3) = Desarr(21) '" D/Gi�"
            .TextMatrix(0, 4) = Desarr(22) '"T/Ti�n`"
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
'Public Sub LoadCommand(n As String, rs As ADODB.Recordset, strTenfield As String)
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

            rs.Find "Index='" & Arr(i - 1 + LastIndex) & "'", , adSearchForward, adBookmarkFirst
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
    cmdSub(i).Picture = Nothing
    cmdSub(i).Caption = ""
Next i
    Do While Not rs.EOF
        If j > 50 Then Exit Sub
            With cmdSub(j)
                If Not rs.EOF Then
                    .Tag = rs.Fields("" & strTenfield & "")
                    If ArrayFlag(SF(3), 6) = 1 Then
                        If blnPrice = 1 Then
                            .Caption = rs.Fields("" & strTenfield2 & "") & vbCrLf & Format(rs.Fields("Std_Price1"), "#,##0") '& Chr(13) & rs.Fields("" & strTenfield2 & "")
                            .Font.Size = 10
                        ElseIf blnPrice = 2 Then
                            .Caption = rs.Fields("" & strTenfield2 & "") & vbCrLf & Format(rs.Fields("Std_Price2"), "#,##0") '& Chr(13) & rs.Fields("" & strTenfield2 & "")
                            .Font.Size = 10
                        ElseIf blnPrice = 3 Then
                            .Caption = rs.Fields("" & strTenfield2 & "") & vbCrLf & Format(rs.Fields("Std_Price3"), "#,##0") '& Chr(13) & rs.Fields("" & strTenfield2 & "")
                            .Font.Size = 10
                        End If
                    Else
                        .Caption = rs.Fields("" & strTenfield2 & "")
                        .Font.Size = 11
                    End If
                    If rs.Fields("Color") <> "" Then
                        .BackColor = HexToDec(rs.Fields("Color"))
                    Else
                        .BackColor = vbRed
                    End If
                    If Dir(rs.Fields("Picture") & "", vbDirectory) <> "" Then
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
    ''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " LoadCommand" & vbCrLf
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

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
    MeUnload = True
    Clipboard.Clear
    CloseRecordset rsTemp
    isLoaded = False
    CloseRecordset rsSetupPLU
    TotalAmt = 0
    VAT = 0
    service_Charge = 0
    txtPhimso = ""
    formCallme = 0
    ReDim Preserve ArrCommand(0)
    For i = 0 To 2
        CustNo(i) = ""
    Next
    Table_ID = ""
    LastIndex = 0
    Discount = 0
    Adj1 = 0
    Adj2 = 0
    Adjtotal1 = 0
    Adjtotal2 = 0
    diemtichluy = 0
    LineNum = 0
    Personal = 0
    LineDelete = ""
    CloseRecordset rsSystem
    CloseRecordset rsInventory
    CloseRecordset rsInvoice_Items
    CloseRecordset rsInvoice_Note
    CloseRecordset rsDepartment
    CloseRecordset rsInvoice_Total
    CloseRecordset rsDelete
    CloseRecordset rsJoin
    CloseRecordset rsLocation
    CloseRecordset rsPriceTime
    CloseRecordset rsShowPLU
    CloseRecordset rslinedelete
    Dept_Index = 0
    CloseRecordset rsFind
End Sub

Public Sub SetFLGRIDORDER(rs As ADODB.Recordset)
On Error GoTo Handle
        Dim incount As Integer
        TotalAmt = 0
        rs.MoveFirst
        With rs
            .Sort = "Line_Number DeSC"
            Do While Not .EOF
                incount = incount + 1
                flgOrder.Rows = rs.RecordCount + 1
                With flgOrder
                    .TextMatrix(incount, 1) = rs!PluName
                    .TextMatrix(incount, 2) = rs!Qty
                    .TextMatrix(incount, 3) = Format(rs!Std_Price1, formatNum)
                    .TextMatrix(incount, 4) = Format(rs!Amt, formatNum)
                    .TextMatrix(incount, 5) = rs!Line_Number
                    If rs.Fields("Status") = False Then
                        .Row = incount
                        .CellBackColor = vbGreen
                    End If
                    '.RowHeight(incount) = 500
                    .CellAlignment = vbAlignTop
                End With
                TotalAmt = TotalAmt + rs!Amt '- rs!Amt * rs!Line_Disc / 100
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
    If Discount > 0 Then
        lblTotalAmt.Caption = Format(TotalAmt - TotalAmt * Discount / 100 + TotalAmt * service_Charge / 100 + TotalAmt * VAT / 100, formatNum)
    Else
        lblTotalAmt.Caption = Format(TotalAmt + TotalAmt * service_Charge / 100 + TotalAmt * VAT / 100, formatNum)
    End If
Exit Sub
Handle:
 ''Print #fFile, Now & vbTab & Err.Number & Err.Description & vbTab & Me.Name & vbTab & " SetFLGRIDORDER" & vbCrLf
MsgBox Err.Number & Err.Description & Me.name & "SetFLGRIDORDER"
End Sub

Public Property Get GetPaymentTotal() As Variant
    GetPaymentTotal = TotalAmt
End Property

Public Property Let GetPaymentTotal(ByVal vNewValue As Variant)
    TotalAmt = vNewValue
End Property


Public Sub GetAllowChangPrice()
    On Error GoTo Handle
        LocationFlag = ArrayFlag(SF(0), 3)
        PriceFlag = ArrayFlag(SF(0), 4)
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "GetAllowChangPrice"
End Sub

Public Sub NewBalance()
On Error Resume Next 'GoTo Handle
Dim i As Integer
Dim j As Integer
Dim LineAmount_Discount As Double
Dim adj(4) As String
diemtichluy = 0
LineAmount_Discount = 0
For i = 1 To 4
DoEvents
    If ArrayFlag(SF(4), i) = 1 Then
        adj(i - 1) = 1
    Else
        adj(i - 1) = 0
    End If
Next i
       
    Dim dblTotal As Double
    i = 0
    j = 1
    dblTotal = 0
    If rsTemp.State <> 0 Then
        With rsInvoice_Items
        If .State = 0 Then Set rsInvoice_Items = Open_Table(cnData, "Invoice_Itemized")
            .Find "Invoice_Number=" & lblBillNo.Caption, , adSearchForward, adBookmarkFirst
            cnData.Execute "delete * from Invoice_Itemized where Invoice_Number=" & lblBillNo.Caption
            If rsTemp.State <> 0 Then
                If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
            Else
                Exit Sub
            End If
                rsTemp.Sort = "Line_Number ASC"
                Do While Not rsTemp.EOF
                    .addNew
                    .Fields("Invoice_Number") = lblBillNo.Caption
                    .Fields("LineNum") = i
                    .Fields("ItemNum") = rsTemp.Fields("PluNo")
                    .Fields("DiffItemName") = rsTemp.Fields("PluName")
                    .Fields("Quantity") = rsTemp.Fields("Qty")
                    .Fields("PricePer") = rsTemp.Fields("Std_Price1")
                    .Fields("Amt") = rsTemp.Fields("Amt")
                    .Fields("Store_ID") = Store_ID
                    .Fields("Returned") = 0
                    .Fields("Kit_Description") = rsTemp.Fields("Kit_Desc")
                    .Fields("LineDisc") = rsTemp.Fields("Line_Disc")
                    .Fields("Line_Disc_Desc") = Trim(Left(rsTemp.Fields("Line_Disc_Desc"), 200))
                    .Fields("TimeOrder") = Trim(rsTemp.Fields("TimeOrder"))
                    LineAmount_Discount = rsTemp.Fields("Line_Disc") * rsTemp.Fields("Qty") * rsTemp.Fields("Std_Price1") / 100
                    .Update
                    dblTotal = dblTotal + CDbl(rsTemp.Fields("Qty") * rsTemp.Fields("Std_Price1")) - LineAmount_Discount
                    rsTemp.MoveNext
                i = i + 1
                Loop
        End With
        If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
            With rsInvoice_Total
                If .State = 0 Then Set rsInvoice_Total = OpenCriticalTable("select * from Invoice_Totals", cnData)
                .Find "Invoice_Number=" & CDbl(lblBillNo.Caption), , adSearchForward, adBookmarkFirst
                If Not .EOF Then
                    If CustNo(0) = "" Then CustNo(0) = "101"
                    !CustNum = CustNo(0)
                    !Total_Price = CDbl("0" & dblTotal)
                    !Discount = Discount
                    !Tax_Rate_ID = CInt("0" & Discount_Status)
                    !service_Charge = service_Charge
                    !VATFee = VAT
                    !AddMoney = MoneyAmount
                    !OrderMan = Emp_ID
                    If formCallme <> 1 Then
                        !cashier_ID = UserID
                    End If
                    
                    !Adj1Rate = Adj1
                    !Adj2Rate = Adj2
                    Call Get_Adjustment_Value_lastest(rsTemp, Adj1, Adj2)
                    !Adjustment1 = Adjtotal1
                    !Adjustment2 = Adjtotal2
                    !Adjustment3 = Adjtotal3
                    !Adjustment4 = Adjtotal4
                    !Total_Tax1 = dblTotal - dblTotal * Discount / 100 + dblTotal * service_Charge / 100 + Adjtotal1 + Adjtotal2 + Adjtotal3 + Adjtotal4 + MoneyAmount
                    !Grand_Total = !Total_Tax1 + !Total_Tax1 * VAT / 100
                    .Fields("InvoiceNotesUsed") = False
                    .Fields("Pro_Desc") = reason_discount
                    .Update
'                    .Requery
                End If
            End With
            Call UpdatePerson(CDbl(lblBillNo.Caption))
            Call AddDatato_Deleted_Items
      Call SendtoKP
    Else
        With rsInvoice_Total
            If .State = 0 Then Set rsInvoice_Total = OpenCriticalTable("select * from Invoice_Totals", cnData)
            .Find "Invoice_Number=" & CDbl(lblBillNo.Caption), , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                !InvoiceNotesUsed = False
                .Update
            End If
        End With
    End If
    isLoaded = False
Exit Sub
'Handle:
'    DoEvents
'    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & " NewBalance"
'    MsgBox Err.Number & Err.Description & Me.Name & " NewBalance"
End Sub

Public Sub SendtoKP()
    On Error GoTo Handle
    Dim i As Integer
    Dim j As Integer
    Dim rsPendingOrder As New ADODB.Recordset
    Dim rsPendingMaster As New ADODB.Recordset
    Dim ischange As Boolean
    If cnData.State = 0 Then
        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    End If
    If Not Check_Table_exist("Pending_Orders") Then
        Call Create_Pending_Orders
    End If
    
    If Not Check_Table_exist("Pending_Orders_Items") Then
        Call Create_Pending_Orders_item
    End If
    
    Set rsPendingMaster = Open_Table(cnData, "Pending_Orders")
    Set rsPendingOrder = Open_Table(cnData, "Pending_Orders_Items")
    
    If rsTemp.State <> 0 And rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
    End If
    j = 0
    For i = 1 To 4
      DoEvents
        If Check_Exist_Printer(i) Then
            If rsTemp.RecordCount = 0 Then GoTo 1
                Do While Not rsTemp.EOF
                    If rsTemp.Fields("Status") = False Then
                        If ArrayFlag(rsTemp.Fields("F2"), i) = 1 Then
                        j = j + 1
                            With rsPendingMaster
                                .Find "Invoice_Number='" & lblBillNo.Caption & "'", , adSearchForward, adBookmarkFirst
                                If .EOF Then
                                    .addNew
                                    .Fields("Invoice_Number") = lblBillNo.Caption
                                    .Fields("Store_ID") = Store_ID
                                    .Fields("Station_ID") = Sec_ID
                                    .Fields("Cashier_ID") = UserID
                                    .Fields("OnHoldID") = rsTemp!TableNo
                                    .Fields("Resend") = rsTemp!Resend
                                    .Fields("Personal") = Personal
                                    .Update
                                    ischange = True
                                End If
                            End With
                            'cnData.Execute "Delete * from Pending_Orders_Items where Invoice_Number=" & lblBillNo.Caption
                            With rsPendingOrder
                                .addNew
                                .Fields("Invoice_Number") = lblBillNo.Caption 'rsTemp!Invoice_Number
                                .Fields("ItemNo") = rsTemp!PluNo
                                .Fields("ItemName") = rsTemp!PluName
                                .Fields("Quan") = rsTemp!Qty
                                .Fields("Store_ID") = Store_ID
                                .Fields("Price") = rsTemp!Std_Price1
                                .Fields("LineNum") = rsTemp!Line_Number
                                .Fields("QuanBurned") = rsTemp!Quanburned
                                .Fields("Kit_Desc") = rsTemp!Kit_Desc
                                .Fields("PrintID") = Format(i, "00")
                                .Fields("Count") = j
        '                        .Fields("TimeOrder") = Format(Now, "HH:mm:ss")
                                .Update
                                ischange = True
                            End With
                            
                        End If
        '            If isticket = True Then Call PrintOrder(Format(i, "00"))
                    End If
                rsTemp.MoveNext
                Loop
        ' Goi du lieu xoa sau khi order xuong bep
1:            If rsDelete.State <> 0 Then
                rsDelete.MoveFirst
                Do While Not rsDelete.EOF
                DoEvents
                    If rsDelete!Ordered = True Then
                        If ArrayFlag(rsDelete.Fields("F2"), i) = 1 Then
                            With rsPendingMaster
                                .Find "Invoice_Number=" & lblBillNo.Caption, , adSearchForward, adBookmarkFirst
                                If .EOF Then
                                    .addNew
                                    .Fields("Invoice_Number") = lblBillNo.Caption
                                    .Fields("Store_ID") = Store_ID
                                    .Fields("Station_ID") = Sec_ID
                                    .Fields("Cashier_ID") = UserID
                                    .Fields("OnHoldID") = Table_ID 'Trim(lblTableNo.Caption)
                                    .Fields("Resend") = 0
                                    .Fields("Personal") = Personal
                                    .Update
                                    ischange = True
                                End If
                            End With
                            With rsPendingOrder
                                .Find "ItemNo='" & rsDelete!PluNo & "'", , adSearchForward, adBookmarkFirst
                                If .EOF Then
                                        .addNew
                                        .Fields("Invoice_Number") = lblBillNo.Caption 'rsTemp!Invoice_Number
                                        .Fields("ItemName") = rsDelete!PluName
                                        .Fields("ItemNo") = rsDelete!PluNo
                                        .Fields("Quan") = 0
                                        .Fields("Store_ID") = Store_ID
                                        .Fields("Price") = rsDelete!Std_Price1
                                        .Fields("LineNum") = i
                                        .Fields("QuanBurned") = rsDelete!Qty
                                        .Fields("Kit_Desc") = rsDelete!Kit_Desc & "-" & rsDelete!reason
                                        .Fields("PrintID") = Format(i, "00")
                                        .Update
                                        ischange = True
                                Else
                                    'Truong hop in 1 mon o 2 may in
        '                            .MoveLast
        '                            .Fields ("PrintID") <> Format(i, "00")
                                   If .Fields("ItemName") <> rsDelete!PluName Or .Fields("Price") <> rsDelete!Std_Price1 Then
                                        .addNew
                                        .Fields("Invoice_Number") = lblBillNo.Caption 'rsTemp!Invoice_Number
                                        .Fields("ItemName") = rsDelete!PluName
                                        .Fields("ItemNo") = rsDelete!PluNo
                                        .Fields("Quan") = 0
                                        .Fields("Store_ID") = Store_ID
                                        .Fields("Price") = rsDelete!Std_Price1
                                        .Fields("LineNum") = i
                                        .Fields("QuanBurned") = rsDelete!Qty
                                        .Fields("Kit_Desc") = rsDelete!Kit_Desc & "-" & rsDelete!reason
                                        .Fields("PrintID") = Format(i, "00")
                                        .Update
                                        ischange = True
                                    End If
                                End If
                            End With
                        End If
                    End If
                rsDelete.MoveNext
                Loop
            End If
        
            If rsDelete.State <> 0 Then
                If rsDelete.RecordCount > 0 Then
                    rsDelete.MoveFirst
                End If
            End If
            If rsTemp.State > 0 And rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
        End If
    Next i
    'Set rsDelete = Nothing
    If ischange = True Then Call PrintOrder '(Format(i, "00"))
    Exit Sub
Handle:
    DoEvents
    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & " SendtoKP"
    'Exit Sub
    If Err.Number = 0 Then
        Exit Sub
    Else
        MsgBox Err.Number & Err.Description & "L�i khi thao t�c in Order ! Vui l�ng nh�n OK �� ti�p t�c !"
    End If
End Sub

'27-7-2012
'Procedure Purpose: Send information to KP
'Write by:Vu Khac Can
Public Sub PrintOrder() '(Printer_ID As String)
On Error Resume Next
            Dim SQL As String
            Dim i As Integer
            Dim Printer_ID As String
            Dim Printer_Name As String
            Dim rsisPrint As New ADODB.Recordset
            Dim iReport As New CRAXDDRT.Report
            Dim cmd As New ADODB.Command
            Dim count, Countdown As Integer
            If ArrayFlag(SF(0), 6) = 1 Then
            count = Get_record_No("Pending_Orders_Items", lblBillNo.Caption)
            Countdown = 1
            For i = 1 To 4
                Printer_ID = Right("00" & i, 2)
                If Check_Exist_Printer(i) Then
                    If ArrayFlag(SF(6), 5) = 1 Then
                        Printer_Name = Get_Printer_Order(Sec_ID, Printer_ID)
                    Else
                        Printer_Name = Get_Friend_Print(Printer_ID)
                    End If
                    If Printer_Name = "" Then GoTo 1 'Printer_Name = Printer.DeviceName '
                    Do While Countdown <= count
                    SQL = "SELECT DISTINCTROW Pending_Orders.Invoice_Number, Pending_Orders.Personal,Pending_Orders.Station_ID," & _
                          "Pending_Orders.Store_ID, Pending_Orders.Cashier_ID, Pending_Orders." & _
                          "OnHoldID, Pending_Orders_Items.ItemName,Pending_Orders_Items.ItemNo,Pending_Orders_Items.Kit_Desc, Pending_Orders_Items.Quan," & _
                          "Pending_Orders_Items.IsModifier, " & _
                          "Pending_Orders_Items.Price, Pending_Orders_Items.QuanBurned, " & _
                          "Pending_Orders_Items.LineNum,( Pending_Orders_Items.Quan-Pending_Orders_Items.QuanBurned) as SendKP" & _
                          " FROM Pending_Orders INNER JOIN Pending_Orders_Items ON Pending_Orders.Invoice_Number = Pending_Orders_Items.Invoice_Number" & _
                          " where Pending_Orders.Resend=0 and PrintID='" & Printer_ID & "' and Pending_Orders.Invoice_Number=" & CDbl("0" & lblBillNo.Caption) & " and Pending_Orders_Items.Count=" & Countdown
        
                     Set rsisPrint = OpenCriticalTable(SQL, cnData)
                     If rsisPrint.RecordCount = 0 Then GoTo 1
                     
                     Call Add_KP_Items(SQL, Printer_ID)
                    
                     
                    Set iReport = Nothing
                    Set crNewBalance = Nothing
                    Set crNewBalance58 = Nothing
                    
                        cmd.ActiveConnection = cnData
                        cmd.CommandText = SQL
                        cmd.Execute
                    If OrderType = "80" Then
                        Set iReport = crNewBalance
                    ElseIf OrderType = "58" Then
                        Set iReport = crNewBalance58
                    Else
                        Set iReport = crNewBalance
                    End If
                    
                    With iReport
                        .Database.AddADOCommand cnData, cmd
                        .KP.SetText Printer_Name
                        .Location.SetUnboundFieldSource "{ado.Station_ID}"
                        .Table.SetUnboundFieldSource "{ado.OnHoldID}"
                        .Cashier.SetUnboundFieldSource "{ado.Cashier_ID}"
                        .Items.SetUnboundFieldSource "{ado.ItemName}"
                        .ItemNum.SetUnboundFieldSource "{ado.ItemNo}"
                        .Qty.SetUnboundFieldSource "{ado.SendKP}"
                        .Price.SetUnboundFieldSource "{ado.Price}"
                        .txtsokhach.SetUnboundFieldSource "{ado.Personal}"
                        .txtKitDesc.SetUnboundFieldSource "{ado.Kit_Desc}"
                        .txtBillNo.SetUnboundFieldSource "{ado.Invoice_Number}"
                        'canh le
                        .TopMargin = TopAlign
                        .BottomMargin = BottomAlign
                        .LeftMargin = LeftAlign
                        .RightMargin = RightAlign
                        
                        With .Qty
                            '.DecimalPlaces = DecimalQtyNumber
                            .DecimalSymbol = DecimalMark
                            .ThousandsSeparators = True
                            .ThousandSymbol = DigitGroupMark
                        End With
                    
                    'An cot gia
                        If ArrayFlag(SF(3), 5) = 0 Then
                            .Price.Suppress = True
                            .Items.HorAlignment = crLeftAlign
                        End If
                        iset = False
                        With frmShowSendKP
                            .Report = iReport
                            .Get_ID = Printer_ID
                            .GetPrinter = Printer_Name
                            '.GetPrinter1 = Printer_Name1
                            .Show vbModal
                        End With
                    End With
                    Countdown = Countdown + 1
                Loop
                
                     cnData.Execute "Delete * from Pending_Orders_Items where Invoice_Number =" & lblBillNo.Caption & " and printID='" & Printer_ID & "'"
                End If
1:
        'In phieu nhac mon xuong bep
        Call Printe_Resend_Item(Printer_ID, Printer_Name, Printer_Name)
        Next i
       
        cnData.Execute "Delete * from Pending_Orders where Invoice_Number=" & lblBillNo.Caption
        cnData.Execute "Delete * from Pending_Orders_Items where Invoice_Number=" & lblBillNo.Caption ' where printID='" & Printer_ID & "'"
    Else
          For i = 1 To 4
                Printer_ID = Right("00" & i, 2)
                If Check_Exist_Printer(i) Then
                    If ArrayFlag(SF(6), 5) = 1 Then
                        Printer_Name = Get_Printer_Order(Sec_ID, Printer_ID)
                    Else
                        Printer_Name = Get_Friend_Print(Printer_ID)
                    End If
                    If Printer_Name = "" Then GoTo 2 'Printer_Name = Printer.DeviceName '
                    SQL = "SELECT DISTINCTROW Pending_Orders.Invoice_Number, Pending_Orders.Personal,Pending_Orders.Station_ID," & _
                          "Pending_Orders.Store_ID, Pending_Orders.Cashier_ID, Pending_Orders." & _
                          "OnHoldID, Pending_Orders_Items.ItemName,Pending_Orders_Items.ItemNo,Pending_Orders_Items.Kit_Desc, Pending_Orders_Items.Quan," & _
                          "Pending_Orders_Items.IsModifier, " & _
                          "Pending_Orders_Items.Price, Pending_Orders_Items.QuanBurned, " & _
                          "Pending_Orders_Items.LineNum,( Pending_Orders_Items.Quan-Pending_Orders_Items.QuanBurned) as SendKP" & _
                          " FROM Pending_Orders INNER JOIN Pending_Orders_Items ON Pending_Orders.Invoice_Number = Pending_Orders_Items.Invoice_Number" & _
                          " where Pending_Orders.Resend=0 and PrintID='" & Printer_ID & "' and Pending_Orders.Invoice_Number=" & CDbl("0" & lblBillNo.Caption)
        
                     Set rsisPrint = OpenCriticalTable(SQL, cnData)
                     If rsisPrint.RecordCount = 0 Then GoTo 2
                     
                     Call Add_KP_Items(SQL, Printer_ID)
                    
                     
                    Set iReport = Nothing
                    Set crNewBalance = Nothing
                    Set crNewBalance58 = Nothing
                    
                        cmd.ActiveConnection = cnData
                        cmd.CommandText = SQL
                        cmd.Execute
                    If OrderType = "80" Then
                        Set iReport = crNewBalance
                    ElseIf OrderType = "58" Then
                        Set iReport = crNewBalance58
                    Else
                        Set iReport = crNewBalance
                    End If
                    
                    With iReport
                        .Database.AddADOCommand cnData, cmd
                        .KP.SetText Printer_Name
                        .Location.SetUnboundFieldSource "{ado.Station_ID}"
                        .Table.SetUnboundFieldSource "{ado.OnHoldID}"
                        .Cashier.SetUnboundFieldSource "{ado.Cashier_ID}"
                        .Items.SetUnboundFieldSource "{ado.ItemName}"
                        .ItemNum.SetUnboundFieldSource "{ado.ItemNo}"
                        .Qty.SetUnboundFieldSource "{ado.SendKP}"
                        .Price.SetUnboundFieldSource "{ado.Price}"
                        .txtsokhach.SetUnboundFieldSource "{ado.Personal}"
                        .txtKitDesc.SetUnboundFieldSource "{ado.Kit_Desc}"
                        .txtBillNo.SetUnboundFieldSource "{ado.Invoice_Number}"
                        'canh le
                        .TopMargin = TopAlign
                        .BottomMargin = BottomAlign
                        .LeftMargin = LeftAlign
                        .RightMargin = RightAlign
                        
                        With .Qty
                            '.DecimalPlaces = DecimalQtyNumber
                            .DecimalSymbol = DecimalMark
                            .ThousandsSeparators = True
                            .ThousandSymbol = DigitGroupMark
                        End With
                    
                    'An cot gia
                        If ArrayFlag(SF(3), 5) = 0 Then
                            .Price.Suppress = True
                            .Items.HorAlignment = crLeftAlign
                        End If
                        iset = False
                        With frmShowSendKP
                            .Report = iReport
                            .Get_ID = Printer_ID
                            .GetPrinter = Printer_Name
                            '.GetPrinter1 = Printer_Name1
                            .Show vbModal
                        End With
                    End With
                
                     cnData.Execute "Delete * from Pending_Orders_Items where Invoice_Number =" & lblBillNo.Caption & " and printID='" & Printer_ID & "'"
                End If
2:
        'In phieu nhac mon xuong bep
        Call Printe_Resend_Item(Printer_ID, Printer_Name, Printer_Name)
        Next i
       
        cnData.Execute "Delete * from Pending_Orders where Invoice_Number=" & lblBillNo.Caption
        cnData.Execute "Delete * from Pending_Orders_Items where Invoice_Number=" & lblBillNo.Caption ' where printID='" & Printer_ID & "'"
    End If
Exit Sub
Handle:
    DoEvents
    cnData.Execute "Delete * from Pending_Orders_Items where printID='" & Printer_ID & "'"
    MsgBox Err.Number & Err.Description & Me.name & " PrintOrder "
End Sub

Public Sub AddDatato_Deleted_Items()
    On Error GoTo Handle
    Dim rsDelete_Items As New ADODB.Recordset
    If cnData.State = 0 Then Exit Sub
    Set rsDelete_Items = Open_Table(cnData, "Items_Deleted")
    'Update 05/12/2011
    If Check_Field_Exist(rsDelete_Items, "Line_Disc") = False Then
        cnData.Execute "ALTER TABLE Items_Deleted ADD COLUMN Line_Disc double, Line_Disc_Desc char"
    End If
    If rsDelete.State = 0 Then Exit Sub
        If rsDelete.RecordCount > 0 Then
            rsDelete.MoveFirst
            Do While Not rsDelete.EOF
            DoEvents
                With rsDelete_Items
                    .addNew
                    .Fields("Sec_ID") = rsDelete.Fields("Sec_No")
                    .Fields("Invoice_Num") = rsDelete!BillNO
                    .Fields("Table_ID") = rsDelete.Fields("TableNo")
                    .Fields("Cashier_ID") = rsDelete.Fields("Cashier_ID")
                    .Fields("PluNo") = rsDelete.Fields("PluNo")
                    .Fields("Quantity") = rsDelete.Fields("Qty")
                    .Fields("Price") = rsDelete.Fields("Std_Price1")
                    .Fields("Amount") = rsDelete.Fields("Amt")
                    .Fields("DateTime") = rsDelete.Fields("DateTime")
                    .Fields("Ordered") = rsDelete.Fields("Ordered")
                    .Fields("Reason") = rsDelete!reason
                    .Fields("PrintCount") = rsDelete!printcount
                    .Fields("Line_Disc") = rsDelete!Line_Disc
                    .Fields("Line_Disc_Desc") = Left(rsDelete!Line_Disc_Desc, 200)
                    .Update
                End With
            rsDelete.MoveNext
            Loop
        End If
    
    Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & " AddDatato_Deleted_Items"
    MsgBox Err.Number & Err.Description & Me.name & "  AddDatato_Deleted_Items "
End Sub

Public Sub delete_Bill_Null(ByVal S As String)
On Error GoTo Handle
Dim rsOnHold As New ADODB.Recordset
'Dim rsInvoice_Notes As New ADODB.Recordset
If cnData.State <> 0 Then
    Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
End If
Set rsOnHold = Open_Table(cnData, "Invoice_OnHold")
'Set rsInvoice_Notes = Open_Table(cnData, "Invoice_Totals_Notes")
'Xoa invoice tam trong Invoice_Notes
    With rsInvoice_Note
        .Find "Invoice_Number=" & S, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Delete adAffectCurrent
            '.Requery
        End If
    End With
'Xoa Ban tam trong Table_OnHold
    With rsOnHold
        .Find "Invoice_Number=" & S, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            .Delete adAffectCurrent
            .Requery
        End If
    End With
'Xoa invoice tam trong Invoice_totals
'    Delay (1000)
'    With rsInvoice_Total
'    .Find "Invoice_Number=" & S, , adSearchForward, adBookmarkFirst
'    If Not .EOF Then
'        .Delete adAffectCurrent
'        .Requery
'    End If
'    End With

Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Public Property Let Get_Secion(ByVal vNewValue As Variant)
    Sec_ID = vNewValue
End Property

Private Sub Label2_Click()
    Call cmdVoidTran_Click
End Sub

Private Sub lblPersonNum_Click()
    Call cmdVoidTran_Click
End Sub

Private Sub cmdListUp_Click()
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

Public Sub Add_OrderMan()
On Error GoTo Handle
    iset = False
    Dim strEmp_ID As String
    
    If ArrayFlag(SF(3), 2) = 1 Then
        With rsInvoice_Total
            .Find "Invoice_Number=" & lblBillNo.Caption, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If .Fields("OrderMan") & " " = " " Then
                    With frmOrderMan
                        .Show vbModal
                        strEmp_ID = .Let_Emp
                    End With
                    .Fields("OrderMan") = strEmp_ID
                    .Update
                End If
                
            End If
        End With
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Public Property Let GetBill_Number(ByVal vNewValue As Variant)
    strBill = vNewValue
End Property

Public Property Let Get_Table_ID(ByVal vNewValue As Variant)
    Table_ID = vNewValue
End Property

Public Sub Get_Charge(ByVal Bill As Double)
On Error GoTo Handle
    With rsInvoice_Total
        If .State = 1 And .RecordCount > 0 Then
            .MoveFirst
        Else
            Exit Sub
        End If
        .Find "Invoice_Number=" & Bill, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            service_Charge = CInt(.Fields("Service_Charge"))
            VAT = CInt(.Fields("VATFee"))
            MoneyAmount = CDbl("0" & .Fields("AddMoney"))
            Personal = CInt("0" & .Fields("Personals"))
            printcount = CInt("0" & .Fields("InvType"))
            Emp_ID = "" & .Fields("OrderMan")
            If Val("0" & Adj1) = 0 Then Adj1 = CInt("0" & .Fields("Adj1Rate"))
            If Val("0" & Adj2) = 0 Then Adj2 = CInt("0" & .Fields("Adj2Rate"))
            If Val("0" & Discount) = 0 Then Discount = CInt("0" & .Fields("Discount"))
        End If
    End With
Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & " Get_ServiceCharge"
    MsgBox Err.Number & Err.Description & Me.name & " Get_ServiceCharge"
End Sub
'Lay so tien mat phu thu
Public Function Get_Money(ByVal Bill As Integer) As Double
On Error GoTo Handle
Dim Temp As Double
    With rsInvoice_Total
        If .State = 1 And .RecordCount > 0 Then
            .MoveFirst
        Else
            
            Exit Function
        End If
        .Find "Invoice_Number=" & Bill, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            Temp = CDbl(.Fields("AddMoney"))
        Else
            Temp = 0
        End If
    End With
    Get_Money = Temp
Exit Function
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & " Get_Money"
    MsgBox Err.Number & Err.Description & Me.name & " Get_Money"
End Function

Public Sub Get_AdjValue(ByVal Bill As Double)
On Error GoTo Handle
    With rsInvoice_Total
        If .State = 1 And .RecordCount > 0 Then
            .MoveFirst
        Else
            
            Exit Sub
        End If
        .Find "Invoice_Number=" & Bill, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            Adjtotal1 = .Fields("Adjustment1")
            Adjtotal2 = .Fields("Adjustment2")
            Adjtotal3 = .Fields("Adjustment3")
            Adjtotal4 = .Fields("Adjustment4")
        Else
            Adjtotal1 = 0
            Adjtotal2 = 0
            Adjtotal3 = 0
            Adjtotal4 = 0
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Get_AdjValue"
End Sub

Public Sub Get_Adjustment_Value(rs As Recordset)
On Error GoTo Handle
Dim rsAdjustment As New ADODB.Recordset
Dim rsDept As New ADODB.Recordset
    'AdjTotal1 = 0: AdjTotal2 = 0: AdjTotal3 = 0: AdjTotal4 = 0:
    
    If cnData.State = 0 Then
        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    End If
    
    Set rsAdjustment = Open_Table(cnData, "Adjustment")
    Set rsDept = Open_Table(cnData, "Departments")
    
    With rsDept
        If .State = 1 And .RecordCount > 0 Then
            .MoveFirst
        Else
            Exit Sub
        End If
        Do While Not rs.EOF
        DoEvents
            If rs.Fields("Status") = False Then
            .Find "Dept_ID='" & rs.Fields("Dept_ID") & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If Left(Right("00000000" & HexToBin(.Fields("F")), 8), 1) = 1 Then
                    With rsAdjustment
                        .Find "AdjNo='01'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            Adjtotal1 = Adjtotal1 + rs.Fields("Amt") * CDbl("0" & .Fields("AdjRate")) / 100
                        End If
                    End With
                ElseIf Mid(Right("00000000" & HexToBin(.Fields("F")), 8), 2, 1) = 1 Then
                    With rsAdjustment
                        .Find "AdjNo='02'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            Adjtotal2 = Adjtotal2 + rs.Fields("Amt") * CDbl("0" & .Fields("AdjRate")) / 100
                        End If
                    End With
                ElseIf Mid(Right("00000000" & HexToBin(.Fields("F")), 8), 3, 1) = 1 Then
                    With rsAdjustment
                        .Find "AdjNo='03'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            Adjtotal3 = Adjtotal3 + rs.Fields("Amt") * CDbl("0" & .Fields("AdjRate")) / 100
                        End If
                    End With
                ElseIf Mid(Right("00000000" & HexToBin(.Fields("F")), 8), 4, 1) = 1 Then
                    With rsAdjustment
                        .Find "AdjNo='04'", , adSearchForward, adBookmarkFirst
                        If Not .EOF Then
                            Adjtotal4 = Adjtotal4 + rs.Fields("Amt") * CDbl("0" & .Fields("AdjRate")) / 100
                        End If
                    End With
                End If
            End If
            End If
        rs.MoveNext
        Loop
    End With

Exit Sub
Handle:
''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub
Public Sub Get_Adjustment_Value_lastest(rs As Recordset, ByVal Adj1Rate As Integer, ByVal Adj2Rate As Integer)
On Error GoTo Handle
Dim rsAdjustment As New ADODB.Recordset
Dim rsDept As New ADODB.Recordset
    Adjtotal1 = 0: Adjtotal2 = 0: Adjtotal3 = 0: Adjtotal4 = 0:
    
    If cnData.State = 0 Then
        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    End If
    
    Set rsAdjustment = Open_Table(cnData, "Adjustment")
    Set rsDept = Open_Table(cnData, "Departments")
    
    With rsDept
        If .State = 1 And .RecordCount > 0 Then
            .MoveFirst
        Else
            Exit Sub
        End If
        Do While Not rs.EOF
        DoEvents
            .Find "Dept_ID='" & rs.Fields("Dept_ID") & "'", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If ArrayFlag(.Fields("F"), 1) = 1 Then
                    Adjtotal1 = Adjtotal1 - rs.Fields("Amt") * Adj1Rate / 100
                ElseIf ArrayFlag(.Fields("F"), 2) = 1 Then
                    Adjtotal2 = Adjtotal2 - rs.Fields("Amt") * Adj2Rate / 100
                End If
            End If
            
        rs.MoveNext
        Loop
    End With

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Public Sub CheckRight()
On Error GoTo Handle
    Dim res As New ADODB.Recordset
        Set res = LoadPasswordData
        With MyRight
            res.MoveFirst
            Do While Not res.EOF
                If StrComp(res.Fields("ID"), UserID, 1) = 0 Then
                    .FullRight = res.Fields("Right")
                    .Banhang = RightDeCode(Mid(.FullRight, 65, 64))
                    .Danhmuc = RightDeCode(Mid(.FullRight, 129, 64))
                    Exit Do
                End If
                res.MoveNext
            Loop
            If Mid(.Banhang, 2, 1) = 1 Then
                AllowDelete = True
'                rightdelete = True
            Else
                AllowDelete = False
'                rightdelete = False
            End If
            If Mid(.Banhang, 3, 1) = 0 Then
                  cmdDiscount.Enabled = False
            Else: cmdDiscount.Enabled = True
            End If
            If Mid(.Banhang, 4, 1) = 0 Then
                  cmdeditprice.Enabled = False
            Else: cmdeditprice.Enabled = True
            End If
            If Mid(.Banhang, 5, 1) = 0 Then
                  cmdItemDiscount.Enabled = False
            Else: cmdItemDiscount.Enabled = True
            End If
            If Mid(.Banhang, 6, 1) = 0 Then
                  cmdExtraPrice.Enabled = False
            Else: cmdExtraPrice.Enabled = True
            End If
            If Mid(.Banhang, 7, 1) = 0 Then
                  cmdEditQuantity.Enabled = False
            Else: cmdEditQuantity.Enabled = True
            End If
            If Mid(.Banhang, 8, 1) = 0 Then
                  cmdBufferPrint.Enabled = False
            Else: cmdBufferPrint.Enabled = True
            End If
            If Mid(.Banhang, 9, 1) = 0 Then
                  cmdTranferTable.Enabled = False
            Else: cmdTranferTable.Enabled = True
            End If
            If Mid(.Banhang, 10, 1) = 0 Then
                  cmdGopban.Enabled = False
            Else: cmdGopban.Enabled = True
            End If
            If Mid(.Banhang, 11, 1) = 0 Then
                  cmdTachmon.Enabled = False
            Else: cmdTachmon.Enabled = True
            End If
            If Mid(.Banhang, 12, 1) = 0 Then
                  cmdOtherPayment.Enabled = False
            Else: cmdOtherPayment.Enabled = True
            End If
            If Mid(.Banhang, 13, 1) = 0 Then
                  cmdAdjustment1.Enabled = False
            Else: cmdAdjustment1.Enabled = True
            End If
            If Mid(.Banhang, 14, 1) = 0 Then
                  cmdAdjustment2.Enabled = False
            Else: cmdAdjustment2.Enabled = True
            End If
            If Mid(.Banhang, 16, 1) = 0 Then
                  cmdEditName.Enabled = False
            Else: cmdEditName.Enabled = True
            End If
            If Mid(.Banhang, 17, 1) = 0 Then
                  cmdItemInfor.Enabled = False
            Else: cmdItemInfor.Enabled = True
            End If
            If Mid(.Banhang, 18, 1) = 0 Then
                  cmdReceiveMoney.Enabled = False
            Else: cmdReceiveMoney.Enabled = True
            End If
            
            If Mid(.Banhang, 19, 1) = 0 Then
                  rightdelete = False
            Else: rightdelete = True
            End If
            
            
            If Mid(.Banhang, 26, 1) = 0 Then
                  delete_ordered = False
            Else: delete_ordered = True
            End If
        End With
    CloseRecordset res
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " CheckRight"
End Sub

Public Sub Add_KP_Items(strSql As String, Printer_ID As String)
On Error GoTo Handle
Dim rsKP_Master As New ADODB.Recordset
Dim rsKP_Items As New ADODB.Recordset
Dim rsIs_Printed As New ADODB.Recordset
Dim rsMax_Line As New ADODB.Recordset
Dim Max_line As Integer
    If cnData.State = 0 Then Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
    Set rsKP_Master = Open_Table(cnData, "Kitchen_Order_Master")
    Set rsKP_Items = Open_Table(cnData, "Kitchen_Order_Items")
    Set rsIs_Printed = OpenCriticalTable(strSql, cnData)
    With rsIs_Printed
        Do While Not .EOF
        DoEvents
            With rsKP_Master
            .Find "Invoice_Number=" & rsIs_Printed.Fields("Invoice_Number"), , adSearchForward, adBookmarkFirst
                If .EOF Then
                    .addNew
                    .Fields("Invoice_Number") = rsIs_Printed.Fields("Invoice_Number")
                    .Fields("Station_ID") = rsIs_Printed.Fields("Station_ID")
                    .Fields("Store_ID") = rsIs_Printed.Fields("Store_ID")
                    .Fields("Cashier_ID") = rsIs_Printed.Fields("Cashier_ID")
                    .Fields("Table_ID") = rsIs_Printed.Fields("onHoldID")
                    .Update
                End If
            End With
            Set rsMax_Line = OpenCriticalTable("select max(LineNum)as Max_line from [Kitchen_Order_Items] where [Kitchen_Order_Items].invoice_number=" & rsIs_Printed.Fields("Invoice_Number"), cnData)
            If Not rsMax_Line.EOF And rsMax_Line.RecordCount > 0 Then
                Max_line = CDbl("0" & rsMax_Line.Fields("Max_Line")) + 1
            Else
                Max_line = 1
            End If
            With rsKP_Items
                .addNew
                .Fields("Invoice_Number") = rsIs_Printed.Fields("Invoice_Number")
                .Fields("ItemNum") = rsIs_Printed.Fields("ItemNo")
                .Fields("ItemName") = rsIs_Printed.Fields("ItemName")
                .Fields("Quantity") = rsIs_Printed.Fields("Quan") - rsIs_Printed.Fields("QuanBurned")
                .Fields("Price") = rsIs_Printed.Fields("Price")
                .Fields("Printer_ID") = Printer_ID
                .Fields("LineNum") = Max_line 'rsIs_Printed.Fields("LineNum")
                .Fields("Kit_Desc") = rsIs_Printed.Fields("Kit_Desc")
                .Fields("Send_KP_Date") = DateDefault
                .Fields("Send_KP_Time") = Format(Now, "HH:mm:ss")
                .Update
            End With
        .MoveNext
        Loop
    End With
    
Exit Sub
Handle:
    DoEvents
    MsgBox Err.Number & Err.Description & Me.name & " Add_KP_Items"
End Sub
Public Function Check_Backup_Printer(Print_ID As String) As Boolean
On Error GoTo Handle

    If ArrayFlag(SF(2), CDbl(Print_ID)) = 1 Then Check_Backup_Printer = True
    
Exit Function
Handle:
    Check_Backup_Printer = False
    MsgBox Err.Number & Err.Description & Me.name & "  Check_Backup_Printer"
End Function

Public Sub GetAutoPrice()
On Error GoTo Handle
    If ArrayFlag(SF(0), 3) = 1 Then
        blnAutoselect_Price = True
    Else
        blnAutoselect_Price = False
    End If
    If ArrayFlag(SF(3), 7) = 1 Then
        lblAutoConsolidate = True
    Else
        lblAutoConsolidate = False
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & ""
End Sub


Public Property Let FormCall(ByVal vNewValue As Variant)
    formCallme = vNewValue
End Property

Public Function Check_Exist_Printer(i As Integer) As Boolean
On Error GoTo Handle
Dim isExist As Boolean
Check_Exist_Printer = False
    
    If ArrayFlag(SF(1), i) = 1 Then isExist = True
    Check_Exist_Printer = isExist
Exit Function
Handle:
    Check_Exist_Printer = False
    MsgBox Err.Number & Err.Description & Me.name & " Check_Exist_Printer"
End Function


Public Function get_AmountKar(BillNO As String) As Double
On Error GoTo Handle
    'Dim rsInvoice_Notes As New ADODB.Recordset
    Dim Value_return As Double
    Value_return = 0
    'Set rsInvoice_Notes = Open_Table(cnData, "Invoice_Totals_Notes")
    With rsInvoice_Note
        If Not .EOF And .RecordCount > 0 Then .MoveFirst
        .Find "Invoice_Number=" & BillNO, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            Value_return = CDbl("0" & .Fields("Karaoke_Amount"))
        End If
    
    End With
    get_AmountKar = Value_return
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " get_AmountKar"
End Function

Public Sub Delete_Invoice_Onhold(iInvoice_num As Integer)
    On Error GoTo Handle
    Dim rsinvoice_hold As New ADODB.Recordset
    Set rsinvoice_hold = Open_Table(cnData, "Invoice_OnHold")
     With rsinvoice_hold
     If .State = 1 And .RecordCount > 0 Then
        .MoveFirst
     Else
        Exit Sub
     End If
      .Find "Invoice_Number=" & iInvoice_num, , adSearchForward, adBookmarkFirst
          If Not .EOF Then
              .Delete adAffectCurrent
              .Requery
          End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Delete_Invoice_Onhold"
End Sub

Public Sub Update_Invoice_Notes(invoice_num As Integer)
On Error GoTo Handle
Dim rsInvoice_Note As New ADODB.Recordset
Set rsInvoice_Note = Open_Table(cnData, "Invoice_Totals_Notes")
    With rsInvoice_Note
        .Find "Invoice_Number=" & invoice_num, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If UCase(.Fields("ClosingTime")) = "C" Then
                .Fields("ClosingTime") = DateDefault & Format(Now, "HH:mm:ss")
                .Update
            End If
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Update_Invoice_Notes"
End Sub

Public Sub Update_Payment(Invoice As Integer)
On Error GoTo Handle
    With rsInvoice_Total
        .Find "Invoice_Number=" & Invoice, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            !Payment_Method = "CA"
            !Status = "C"
            If CDbl(Val(Replace(txtQty.Text, ",", ""))) >= CDbl(!Total_Price) Then
                !Amt_Tendered = CDbl(Val(Replace(txtQty.Text, ",", "")))
                !Amt_Change = CDbl(Val(Replace(txtQty.Text, ",", ""))) - !Grand_Total
                .Update
            Else
                !Amt_Tendered = !Grand_Total
                !Amt_Change = 0
            End If
            .Update
'            .Requery
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Update_Payment"
End Sub

Public Sub Update_Invoice_Total_Isprint(Invoice As Double)
    On Error GoTo Handle
        With rsInvoice_Total
            .Find "Invoice_Number=" & Invoice, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                If formCallme <> 1 Then
                    .Fields("Status") = "P"
                    .Fields("InvType") = CInt("0" & .Fields("InvType")) + 1
                Else
                    .Fields("Status") = "C"
                End If
                .Update
'                .Requery
            End If
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " Update_Invoice_Total_Isprint"
End Sub

Public Sub Update_OrderMan()
On Error GoTo Handle
    iset = False
    With rsInvoice_Total
        .Find "Invoice_Number=" & lblBillNo.Caption, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
                With frmOrderMan
                    .Show vbModal
                    Emp_ID = .Let_Emp
                End With
        .Fields("OrderMan") = Emp_ID
        .Update
        End If
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Update_OrderMan"
End Sub

Public Property Get Get_Discount() As Variant
    Get_Discount = Discount
End Property

Public Property Let Get_Discount(ByVal vNewValue As Variant)
    Discount = vNewValue
End Property

Private Sub cmdOrderMan_Click()
    iset = False
    Call Update_OrderMan
    fraEdit.Visible = False
End Sub

Public Sub Printe_Resend_Item(printerID As String, PrinterName As String, printername1 As String)
On Error GoTo Handle
    Dim iReport As New CRAXDDRT.Report
    Dim cmd As New ADODB.Command
    Dim SQL As String
    Dim rsisPrint As New ADODB.Recordset
    
    SQL = "SELECT Pending_Orders.Invoice_Number, Pending_Orders.Station_ID," & _
              "Pending_Orders.Store_ID, Pending_Orders.Cashier_ID, Pending_Orders." & _
              "OnHoldID,Pending_Orders_Items.ItemNo, Pending_Orders_Items.ItemName,Pending_Orders_Items.Kit_Desc, Pending_Orders_Items.Quan," & _
              "Pending_Orders_Items.IsModifier, " & _
              "Pending_Orders_Items.Price, Pending_Orders_Items.QuanBurned, " & _
              "Pending_Orders_Items.LineNum,( Pending_Orders_Items.Quan-Pending_Orders_Items.QuanBurned) as SendKP" & _
              " FROM Pending_Orders INNER JOIN Pending_Orders_Items ON Pending_Orders.Invoice_Number = Pending_Orders_Items.Invoice_Number " & _
              " where Pending_Orders.Resend=true and Pending_Orders_Items.PrintID='" & printerID & "'"
         
         Set rsisPrint = OpenCriticalTable(SQL, cnData)
         If rsisPrint.RecordCount = 0 Then
            Set rsisPrint = Nothing
            Exit Sub
         End If
         
        Set iReport = Nothing
        Set crResendKP = Nothing
            cmd.ActiveConnection = cnData
            cmd.CommandText = SQL
            cmd.Execute
        Set iReport = crResendKP
        With iReport
            .Database.AddADOCommand cnData, cmd
            .Location.SetUnboundFieldSource "{ado.Station_ID}"
            .Table.SetUnboundFieldSource "{ado.OnHoldID}"
            .Cashier.SetUnboundFieldSource "{ado.Cashier_ID}"
            .ItemNum.SetUnboundFieldSource "{ado.ItemNo}"
            .Items.SetUnboundFieldSource "{ado.ItemName}"
            .Qty.SetUnboundFieldSource "{ado.SendKP}"
            .Price.SetUnboundFieldSource "{ado.Price}"
            .txtKitDesc.SetUnboundFieldSource "{ado.Kit_Desc}"
            .txtBillNo.SetUnboundFieldSource "{ado.Invoice_Number}"
            With .Qty
                .DecimalPlaces = DecimalQtyNumber
                .DecimalSymbol = DecimalMark
                .ThousandsSeparators = True
                .ThousandSymbol = DigitGroupMark
            End With
        End With
        iset = False
        With frmShowSendKP
            .Report = iReport
            .Get_ID = printerID
            .GetPrinter = PrinterName
           ' .GetPrinter1 = printername1
            .Show vbModal
        End With
    cnData.Execute "Delete * from Pending_Orders_Items where printID='" & printerID & "'"
Exit Sub
Handle:
Exit Sub
''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & " Printe_Resend_Item"
MsgBox Err.Number & Err.Description & Me.name & " Printe_Resend_Item"
End Sub

Public Function getDiscount() As Integer
    On Error GoTo Handle
        Dim Result As Integer
        Dim rsAdjust As New ADODB.Recordset
        Set rsAdjust = Open_Table(cnData, "Adjustment")
        
        With rsAdjust
            .Find "AdjNo=05", , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                Result = .Fields("AdjRate")
            End If
        End With
        getDiscount = Result
        
    Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " getDiscount"
    getDiscount = 0
End Function


Private Sub txtQty_KeyPress(KeyAscii As Integer)
On Error GoTo Handle
    If KeyAscii < 32 And KeyAscii <> 13 Then Exit Sub
    Select Case KeyAscii
        Case 48 To 57, 45, 46
        Case 13
            If Len(txtQty.Text) > 3 Then
                If isCust = True Then
                    Dim ID As String
                    If txtQty.Text = "" Then Exit Sub
                    ID = TrimSpecialChar(txtQty.Text)
                    txtQty.Text = ""
                    iset = False
                    With frmPrintCust
                        .Get_CustID = ID
                        .Get_Total = TotalAmt
                        .Show vbModal
                    End With
                    'lblDiscount.Caption = Discount & "%"
                    lblCustomer.Caption = CustNo(1)
                    lblTotalAmt.Caption = Format(TotalAmt - TotalAmt * Discount / 100, "#,##0")
                    isCust = False
                Else
                    Exit Sub
                End If
            Else
                Call cmdAlpha_Click(14)
                txtSearch.Text = ""
                txtSearch.SetFocus
        End If
        Case Else:   KeyAscii = 0
    End Select
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " txtQty_KeyPress"
End Sub

Public Function check_IsPrint(BillNO As Double) As Boolean
On Error GoTo Handle
Dim blnPrint As Boolean

    With rsInvoice_Total
        .Find "Invoice_Number=" & BillNO, , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            If .Fields("Status") = "P" Then
                    blnPrint = True
            End If
        End If
    End With
check_IsPrint = blnPrint
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & Me.name
    check_IsPrint = False
End Function


Public Sub get_Discount_Auto()
On Error GoTo Handle
    Dim i As Integer
    Dim rsdiscount As New ADODB.Recordset
    Dim rsMiss As New ADODB.Recordset
    
    Set rsdiscount = Open_Table(cnData, "Adjustment")
    Set rsMiss = Open_Table(cnData, "MismatchTable")
    
    With rsMiss
        Do While Not .EOF
        If .Fields(0) = "1" Then
            If gfCONVERT_DATE_TO_STRING(.Fields(2)) <= Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") And Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") <= gfCONVERT_DATE_TO_STRING(.Fields(3)) Then
                With rsdiscount
                    .Find "AdjNo='" & Format("5", "00") & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        If Discount = 0 Then
                            Discount = .Fields("AdjRate")
                        End If
                    End If
                End With
            End If
        ElseIf .Fields(0) = "2" Then
            If gfCONVERT_DATE_TO_STRING(.Fields(2)) <= Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") And Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") <= gfCONVERT_DATE_TO_STRING(.Fields(3)) Then
                With rsdiscount
                    .Find "AdjNo='" & Format("1", "00") & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        If Adj1 = 0 Then Adj1 = .Fields("AdjRate")
                    End If
                End With
            End If
        ElseIf .Fields(0) = "3" Then
            If gfCONVERT_DATE_TO_STRING(.Fields(2)) <= Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") And Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00") <= gfCONVERT_DATE_TO_STRING(.Fields(3)) Then
                With rsdiscount
                    .Find "AdjNo='" & Format("2", "00") & "'", , adSearchForward, adBookmarkFirst
                    If Not .EOF Then
                        If Adj2 = 0 Then Adj2 = .Fields("AdjRate")
                    End If
                End With
            End If
        End If
            
       .MoveNext
       Loop
    End With
    
Exit Sub
Handle:
    ''Print #fFile, Now & vbTab & Err.Number & vbTab & Err.Description & vbTab & Me.Name & vbTab & " get_Discount_Auto"
    MsgBox Err.Number & Err.Description & Me.name & " get_Discount_Auto"
End Sub
Public Sub UpdatePerson(invoice_num As Double)
    On Error GoTo Handle
        Dim rsTotal_person As New ADODB.Recordset
        Set rsTotal_person = Open_Table(cnData, "Invoice_Totals_Person_Mapping")
        With rsTotal_person
            .Find "Invoice_Number=" & invoice_num, , adSearchForward, adBookmarkFirst
            If .EOF Then
                .addNew
                .Fields("Invoice_Number") = invoice_num
                .Fields("Store_ID") = Store_ID
                .Fields("SeatNum") = Personal
                .Update
            End If
            
        End With
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " -UpdatePerson"
End Sub

Public Sub Update_Cancel_Bill(ByVal BillCancel As Double)
    On Error GoTo Handle
    Dim rsinvoice_hold As New ADODB.Recordset
    Set rsinvoice_hold = Open_Table(cnData, "Invoice_OnHold")
        With rsInvoice_Note
            .Find "Invoice_Number=" & BillCancel, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Fields("ClosingTime") = DateDefault & Format(Now, "HH:mm:ss")
                .Update
            End If
        End With
        With rsInvoice_Total
            .Find "Invoice_Number=" & BillCancel, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Fields("Status") = "CO"
                .Update
            End If
        End With
        With rsinvoice_hold
            .Find "Invoice_Number=" & BillCancel, , adSearchForward, adBookmarkFirst
            If Not .EOF Then
                .Delete adAffectCurrent
                .Requery
            End If
        End With
        
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " -Update_Cancel_Bill"
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

Public Sub Display_Sale(name As String, strDisplay As Variant)
On Error GoTo Handle
Dim i
'    If CDbl(strDisplay) = 0 Then
'        strDisplay = "00000.000"
'    Else
'        strDisplay = Right("000000000" & TrimSpecial(strDisplay), 9)
'    End If
    With MSCom
        If .PortOpen = True Then  'And .CTSHolding = False
            .RThreshold = 1
            .SThreshold = 0
            .InputMode = comInputModeText
'            For i = 1 To Len(strDisplay)
                
                .output = Chr$(13)
                .output = " Total:" & Format(strDisplay, "#,##0")
'                .output = strDisplay
'                .output = Mid(strDisplay, i, 1)
'            Next
                
        End If
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - Display_Sale"
End Sub

Public Sub Open_Port()
On Error GoTo Handle
Dim CommPort As String
Dim setting, HandShaking As String
CommPort = GetSettingStr("Properties", "ComPortNumber", "", myIniFile)
If CommPort <> "" Then MSCom.CommPort = CommPort
setting = GetSettingStr("Properties", "setting", "", myIniFile)
If setting <> "" Then MSCom.settings = setting
    With MSCom
        If .PortOpen = False Then .PortOpen = True
        '.CTSHolding = False
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - Open_Port"
End Sub


Public Sub cmdFilter_Click()
On Error Resume Next 'GoTo Handle '
    Dim rs As New ADODB.Recordset
    Dim rsLast As New ADODB.Recordset
    Dim bt As CommandButton
    Dim i As Integer
    Dim ctrl As Control
    Set rsShowPLU = Nothing
    
    'cnData.Execute "delete * from SetupPLU"
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
    'cap nhap lai thong tin
'    With txtSearch
'        .Text = "Nh�p t�n m�n c�n t�m"
'    End With
    
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
        .SelStart = Len(.Text)
        .SelLength = Len(.Text)
        .SelText = ""
    End With
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtSearch_GotFocus"
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Handle
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
        Set rsFind = OpenCriticalTable(strSql, cnData)
        With dtgFind
            Set .DataSource = rsFind
            .Columns(0).Width = 0
            .Columns(1).Width = 3000
            .Columns(2).Width = 1500
            .Columns(3).Width = 1500
            
            .Visible = True
            .SetFocus
        End With
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtSearch_KeyDown"
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    On Error GoTo Handle
        If KeyAscii = 13 Then
            Call cmdFilter_Click
            dtgFind.Visible = False
        End If
   ' Call cmdFilter_Click
        txtSearch.SetFocus
        If KeyAscii = vbKeyEscape Then dtgFind.Visible = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " txtSearch_KeyPress"
End Sub

Public Function Get_Friend_Print(ByVal prtID As String) As String
On Error GoTo Handle
Dim rsFriend_Printer As New ADODB.Recordset
Dim strSql As String
Dim PrintName As String
strSql = "SELECT Friendly_Printers.PrtID, Friendly_Printers.PrinterName, Printer_Mapping.Details" & _
                " FROM Friendly_Printers INNER JOIN Printer_Mapping ON Friendly_Printers.PrtID = Printer_Mapping.PrinterName" & _
                " GROUP BY Friendly_Printers.PrtID, Friendly_Printers.PrinterName, Printer_Mapping.Details"
    Set rsFriend_Printer = OpenCriticalTable(strSql, cnData)
    With rsFriend_Printer
        If .State <> 0 And .RecordCount > 0 Then .MoveFirst
        .Find "PrtID='" & prtID & "'", , adSearchForward, adBookmarkFirst
        If Not .EOF Then
            PrintName = .Fields("Details")
        End If
    End With
    Get_Friend_Print = PrintName
Exit Function
Handle:
    MsgBox Err.Number & Err.Description & " - Get_Friend_Print"
End Function

Public Function check_already_exit_Invoicce_Number_Pending(invoice_num As Double) As Boolean
On Error GoTo Handle
Dim is_Already As Boolean
is_Already = False
Dim rspending_orders As New ADODB.Recordset
Set rspending_orders = OpenCriticalTable("select * from Pending_Orders where Invoice_Number=" & invoice_num, cnData)
With rspending_orders
    If .EOF Then
        is_Already = False
    Else
        is_Already = True
    End If
End With
check_already_exit_Invoicce_Number_Pending = is_Already
Exit Function
Handle:
MsgBox Err.Number & Err.Description & " - check_already_exit_Invoicce_Number_Pending"
End Function

Public Property Let Get_Price_Level(ByVal vNewValue As Variant)
    blnPrice = vNewValue
End Property

Public Property Let Get_VAT(ByVal vNewValue As Variant)
    VAT = vNewValue
End Property


Public Property Let Get_Service(ByVal vNewValue As Variant)
    service_Charge = vNewValue
End Property

Public Property Let Get_PriceRate(ByVal vNewValue As Variant)
    PriceRate = vNewValue
End Property

Public Property Get Get_Adj1() As Variant
    Adj1 = Get_Adj1
End Property

Public Property Get Get_Adj2() As Variant
    Adj2 = Get_Adj2
End Property

Public Property Let Get_Adj1(ByVal vNewValue As Variant)
    Adj1 = vNewValue
End Property


Public Property Let Get_Adj2(ByVal vNewValue As Variant)
    Adj2 = vNewValue
End Property


Public Property Let Get_CustID(ByVal vNewValue As Variant)
CustNo(0) = vNewValue
End Property

Public Property Let Get_Record_Ordered(ByVal vNewValue As Variant)
    rsNew = vNewValue
End Property