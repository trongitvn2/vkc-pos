VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmHelp 
   Caption         =   "Gi�p ��"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   ClipControls    =   0   'False
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
   MinButton       =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjTouchScreen.MyButton cmdClose 
      Cancel          =   -1  'True
      Height          =   1335
      Left            =   12720
      TabIndex        =   62
      Top             =   9720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2355
      BTYPE           =   5
      TX              =   "��&ng"
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
      BCOL            =   255
      BCOLO           =   255
      FCOL            =   16777215
      FCOLO           =   16777215
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmHelp.frx":0000
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
      Caption         =   "H��ng d�n thao t�c"
      ForeColor       =   &H00FF0000&
      Height          =   5295
      Left            =   6600
      TabIndex        =   53
      Top             =   120
      Width           =   8535
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4575
         Left            =   240
         TabIndex        =   54
         Top             =   480
         Width           =   8055
      End
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
      TabIndex        =   0
      Top             =   0
      Width           =   6470
      Begin VB.PictureBox Picture1 
         Height          =   735
         Left            =   0
         ScaleHeight     =   675
         ScaleWidth      =   5040
         TabIndex        =   39
         Top             =   0
         Width           =   5100
         Begin VB.Label lblBill 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            Caption         =   "S� H�"
            ForeColor       =   &H00FFFFFF&
            Height          =   400
            Left            =   0
            TabIndex        =   47
            Tag             =   "L1"
            Top             =   0
            Width           =   1185
         End
         Begin VB.Label lblNhanVien 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            Caption         =   "Nh�n vi�n"
            ForeColor       =   &H00FFFFFF&
            Height          =   400
            Left            =   3510
            TabIndex        =   46
            Tag             =   "L4"
            Top             =   0
            Width           =   1545
         End
         Begin VB.Label lblStation 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            Caption         =   "Khu v�c"
            ForeColor       =   &H00FFFFFF&
            Height          =   400
            Left            =   2310
            TabIndex        =   45
            Tag             =   "L3"
            Top             =   0
            Width           =   1245
         End
         Begin VB.Label lblTable 
            Alignment       =   2  'Center
            BackColor       =   &H00404000&
            Caption         =   "B�n s�"
            ForeColor       =   &H00FFFFFF&
            Height          =   400
            Left            =   1170
            TabIndex        =   44
            Tag             =   "L2"
            Top             =   0
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
            TabIndex        =   43
            Top             =   360
            Width           =   1215
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
            TabIndex        =   42
            Top             =   360
            Width           =   1215
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
            TabIndex        =   41
            Top             =   360
            Width           =   1545
         End
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
            TabIndex        =   40
            Top             =   360
            Width           =   1185
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00404000&
         BorderStyle     =   0  'None
         FillColor       =   &H008080FF&
         ForeColor       =   &H008080FF&
         Height          =   1215
         Left            =   0
         ScaleHeight     =   1215
         ScaleWidth      =   5115
         TabIndex        =   31
         Top             =   5880
         Width           =   5120
         Begin VB.Label lblCustomer 
            BackStyle       =   0  'Transparent
            Caption         =   "ABC"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   2490
            TabIndex        =   38
            Top             =   780
            Width           =   2445
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Gi�m %"
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
            Left            =   10
            TabIndex        =   37
            Tag             =   "L9"
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblDiscount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   1515
            TabIndex        =   36
            Top             =   390
            Width           =   3375
         End
         Begin VB.Label lblTotalAmt 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   20.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   525
            Left            =   1485
            TabIndex        =   35
            Top             =   0
            Width           =   3255
         End
         Begin VB.Label lblTotal 
            BackStyle       =   0  'Transparent
            Caption         =   "T�ng c�ng:"
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
            Left            =   10
            TabIndex        =   34
            Tag             =   "L5"
            Top             =   15
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "S� kh�ch:"
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
            Left            =   30
            TabIndex        =   33
            Top             =   750
            Width           =   1185
         End
         Begin VB.Label lblPersonNum 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   ".VnArial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   1260
            TabIndex        =   32
            Top             =   750
            Width           =   765
         End
      End
      Begin VB.PictureBox pictFunction 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Height          =   11535
         Left            =   5090
         ScaleHeight     =   11535
         ScaleWidth      =   1395
         TabIndex        =   18
         Top             =   0
         Width           =   1390
         Begin MSForms.CommandButton cmdDelete 
            Height          =   945
            Left            =   0
            TabIndex        =   30
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
         Begin MSForms.CommandButton cmdDiscount 
            Height          =   945
            Left            =   0
            TabIndex        =   29
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
         Begin MSForms.CommandButton cmdTranferTable 
            Height          =   945
            Left            =   0
            TabIndex        =   28
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
         Begin MSForms.CommandButton cmdGopban 
            Height          =   945
            Left            =   0
            TabIndex        =   27
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
         Begin MSForms.CommandButton cmdCookingMessage 
            Height          =   945
            Left            =   0
            TabIndex        =   26
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
         Begin MSForms.CommandButton cmdExtraPrice 
            Height          =   945
            Left            =   0
            TabIndex        =   25
            Tag             =   "L26"
            Top             =   4730
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
         Begin MSForms.CommandButton cmdEditQuantity 
            Height          =   945
            Left            =   0
            TabIndex        =   24
            Tag             =   "L27"
            Top             =   5665
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "S�a sai SL"
            Size            =   "2355;1667"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdEditName 
            Height          =   945
            Left            =   0
            TabIndex        =   23
            Top             =   6610
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "S�a t�n m�n"
            Size            =   "2355;1667"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdItemDiscount 
            Height          =   945
            Left            =   0
            TabIndex        =   22
            Tag             =   "L11"
            Top             =   8500
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "Gi�m % m�n"
            Size            =   "2355;1667"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdPrice2 
            Height          =   1065
            Left            =   0
            TabIndex        =   21
            Tag             =   "L36"
            Top             =   9450
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "Gi� 2"
            Size            =   "2355;1879"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdPrice3 
            Height          =   1020
            Left            =   0
            TabIndex        =   20
            Tag             =   "L37"
            Top             =   10515
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "Gi� 3"
            Size            =   "2355;1799"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdReSendKP 
            Height          =   945
            Left            =   0
            TabIndex        =   19
            Top             =   7560
            Width           =   1335
            ForeColor       =   16777215
            BackColor       =   12582912
            VariousPropertyBits=   8388635
            Caption         =   "Nh�c m�n"
            Size            =   "2355;1667"
            FontName        =   ".VnArial"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         Height          =   3855
         Left            =   0
         ScaleHeight     =   3855
         ScaleWidth      =   5055
         TabIndex        =   1
         Top             =   7680
         Width           =   5055
         Begin VB.TextBox txtQty 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
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
            TabIndex        =   2
            Top             =   5
            Width           =   3900
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1035
            Index           =   0
            Left            =   0
            TabIndex        =   17
            Top             =   740
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "1"
            PicturePosition =   131072
            Size            =   "1720;1826"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1035
            Index           =   1
            Left            =   990
            TabIndex        =   16
            Top             =   740
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "2"
            PicturePosition =   131072
            Size            =   "1720;1826"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1035
            Index           =   2
            Left            =   1980
            TabIndex        =   15
            Top             =   740
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "3"
            PicturePosition =   131072
            Size            =   "1720;1826"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1035
            Index           =   3
            Left            =   0
            TabIndex        =   14
            Top             =   1800
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "4"
            PicturePosition =   131072
            Size            =   "1720;1826"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1035
            Index           =   4
            Left            =   990
            TabIndex        =   13
            Top             =   1800
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "5"
            PicturePosition =   131072
            Size            =   "1720;1826"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1035
            Index           =   5
            Left            =   1980
            TabIndex        =   12
            Top             =   1800
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "6"
            PicturePosition =   131072
            Size            =   "1720;1826"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1035
            Index           =   6
            Left            =   0
            TabIndex        =   11
            Top             =   2860
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "7"
            PicturePosition =   131072
            Size            =   "1720;1826"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1035
            Index           =   7
            Left            =   990
            TabIndex        =   10
            Top             =   2860
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "8"
            PicturePosition =   131072
            Size            =   "1720;1826"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1035
            Index           =   8
            Left            =   1980
            TabIndex        =   9
            Top             =   2860
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "9"
            PicturePosition =   131072
            Size            =   "1720;1826"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1035
            Index           =   9
            Left            =   2970
            TabIndex        =   8
            Top             =   740
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "0"
            PicturePosition =   131072
            Size            =   "1720;1826"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1035
            Index           =   10
            Left            =   2970
            TabIndex        =   7
            Top             =   1800
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "00"
            PicturePosition =   131072
            Size            =   "1720;1826"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   435
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1035
            Index           =   11
            Left            =   2970
            TabIndex        =   6
            Top             =   2860
            Width           =   975
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "."
            PicturePosition =   131072
            Size            =   "1720;1826"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   480
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   705
            Index           =   12
            Left            =   3960
            TabIndex        =   5
            Top             =   0
            Width           =   1125
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "Bks"
            PicturePosition =   131072
            Size            =   "1984;1244"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   285
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   1035
            Index           =   13
            Left            =   3960
            TabIndex        =   4
            Top             =   735
            Width           =   1125
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "CLR"
            PicturePosition =   131072
            Size            =   "1984;1826"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   315
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton cmdAlpha 
            Height          =   2050
            Index           =   14
            Left            =   3960
            TabIndex        =   3
            Top             =   1800
            Width           =   1125
            ForeColor       =   16711680
            BackColor       =   8421504
            VariousPropertyBits=   8388635
            Caption         =   "Enter"
            PicturePosition =   131072
            Size            =   "1984;3616"
            FontName        =   ".VnArial"
            FontEffects     =   1073741825
            FontHeight      =   315
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin MSFlexGridLib.MSFlexGrid flgOrder 
         Height          =   5190
         Left            =   0
         TabIndex        =   48
         Top             =   720
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   9155
         _Version        =   393216
         Rows            =   16
         Cols            =   6
         BackColor       =   14737632
         ForeColor       =   0
         BackColorFixed  =   14737632
         ForeColorFixed  =   16711680
         ForeColorSel    =   16777088
         BackColorBkg    =   14737632
         GridColor       =   4210752
         GridColorFixed  =   4210752
         WordWrap        =   -1  'True
         Redraw          =   -1  'True
         ScrollTrack     =   -1  'True
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial NarrowH"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label label 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Customer:"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   90
         TabIndex        =   52
         Tag             =   "L14"
         Top             =   10350
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   90
         TabIndex        =   51
         Tag             =   "L34"
         Top             =   10740
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSForms.CommandButton MyButton1 
         Height          =   615
         Left            =   2550
         TabIndex        =   50
         Top             =   7080
         Width           =   2535
         BackColor       =   8454143
         Size            =   "4471;1085"
         Picture         =   "frmHelp.frx":001C
         FontName        =   ".VnArial"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdListdown 
         Height          =   615
         Left            =   10
         TabIndex        =   49
         Top             =   7080
         Width           =   2535
         BackColor       =   8454143
         Size            =   "4471;1085"
         Picture         =   "frmHelp.frx":01AB
         FontName        =   ".VnArial"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin MSForms.CommandButton cmdTachmon 
      Height          =   1140
      Left            =   8700
      TabIndex        =   61
      Top             =   6840
      Width           =   2070
      ForeColor       =   16777215
      BackColor       =   12582912
      VariousPropertyBits=   8388635
      Caption         =   "Chuy�n m�n"
      Size            =   "3651;2011"
      FontName        =   ".VnArial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdAdjustment2 
      Height          =   1140
      Left            =   6600
      TabIndex        =   60
      Top             =   5640
      Width           =   2070
      ForeColor       =   16777215
      BackColor       =   12582912
      VariousPropertyBits=   8388635
      Caption         =   "Gi�m % Th�c u�ng"
      Size            =   "3651;2011"
      FontName        =   ".VnArial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdAdjustment1 
      Height          =   1140
      Left            =   8700
      TabIndex        =   59
      Top             =   5640
      Width           =   2070
      ForeColor       =   16777215
      BackColor       =   12582912
      VariousPropertyBits=   8388635
      Caption         =   "Gi�m % Th�c �n"
      Size            =   "3651;2011"
      FontName        =   ".VnArial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdVAT 
      Height          =   1140
      Left            =   10815
      TabIndex        =   58
      Top             =   6840
      Width           =   2070
      ForeColor       =   16777215
      BackColor       =   12582912
      VariousPropertyBits=   8388635
      Caption         =   "Thu� VAT"
      Size            =   "3651;2011"
      FontName        =   ".VnArial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdReceiveMoney 
      Height          =   1140
      Left            =   12945
      TabIndex        =   57
      Top             =   5640
      Width           =   2070
      ForeColor       =   16777215
      BackColor       =   12582912
      VariousPropertyBits=   8388635
      Caption         =   "Ph� thu ti�n m�t"
      Size            =   "3651;2011"
      FontName        =   ".VnArial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdServiceCharge 
      Height          =   1140
      Left            =   10815
      TabIndex        =   56
      Top             =   5640
      Width           =   2070
      ForeColor       =   16777215
      BackColor       =   12582912
      VariousPropertyBits=   8388635
      Caption         =   "% Ph� ph�c v�"
      Size            =   "3651;2011"
      FontName        =   ".VnArial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdeditprice 
      Height          =   1140
      Left            =   6600
      TabIndex        =   55
      Top             =   6840
      Width           =   2070
      ForeColor       =   16777215
      BackColor       =   12582912
      VariousPropertyBits=   8388635
      Caption         =   "S�a gi�"
      Size            =   "3651;2011"
      FontName        =   ".VnArial"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCookingMessage_Click()
    lblDescription.Caption = "Ch� d�n ch� bi�n" & vbCrLf & "Sau khi order m�n --> B�m th�ng tin Ch� d�n b�p --> G� ho�c ch�n ch� d�n --> ��ng �. Ch� d�n s� ���c in k�m v�i phi�u g�i m�n" & _
    vbCrLf & "V� d�:" & vbCrLf & Space(5) & "Kh�ch g�i m�n: Cafe �� (Kh�ng ���ng)" & vbCrLf & "Thao t�c:" & _
    vbCrLf & Space(5) & "B�m Cafe �� --> Th�ng tin ch� d�n b�p --> G� b�n ph�m ho�c ch�n Kh�ng ���ng --> ��ng �"
End Sub

Private Sub cmdDelete_Click()
    lblDescription.Caption = "X�a m�n" & vbCrLf & Space(5) & "Ch�n M�n c�n x�a b�n Chi ti�t m�n trong b�n (b�n tay tr�i m�n h�nh --> B�m ch�n X�a" & _
        vbCrLf & " N�u m�n �� ���c l�u, Vui l�ng nh�p l� do x�a trong c�a s� L� do --> OK"
End Sub

Private Sub cmdDiscount_Click()
    lblDescription.Caption = "Gi�m % T�ng H�a ��n" & vbCrLf & "Nh�p s� % c�n gi�m tr�n ph�m s� l��ng --> B�m Gi�m %" & _
        vbCrLf & " Ch� �:" & vbCrLf & Space(5) & " N�u s� % gi�m ���c nh�p sai th� l�p l�i thao t�c c�, ch��ng tr�nh ch� ghi nh�n l�i s� % gi�m sau c�ng" & _
        vbCrLf & "V� d�: Gi�m 20% m� l� �� nh�p 10% r�i th� v�n b�m 20 -->Gi�m %"
End Sub

Private Sub cmdEditName_Click()
lblDescription.Caption = "S�a t�n m�n ho�c m�n ngo�i th�c ��n" & vbCrLf & Space(5) & _
    "Ch�n m�n c�n s�a t�n (b�n chi ti�t b�n tr�i m�n h�nh) --> B�m S�a t�n m�n --> Nh�p t�n m�n m�i --> Enter" & _
    vbCrLf & "V� d�:" & vbCrLf & Space(5) & "Kh�ch g�i m�n: B� x�o c� h�nh m� trong th�c ��n ch� c� B� x�o b�ng c�i" & _
    vbCrLf & vbCrLf & "Thao t�c:" & vbCrLf & Space(5) & "B�m B� x�o B�ng c�i --> Ch�n m�n B� x�o b�ng c�i b�n List b�n tr�i --> B�m S�a t�n m�n -->Nh�p B� x�o c� h�nh --> Enter"
    
End Sub

Private Sub cmdEditQuantity_Click()
    lblDescription.Caption = "S�a sai s� l��ng ho�c tr� m�n" & vbCrLf & Space(5) & _
    "Ch�n m�n c�n tr� ho�c s�a s� l��ng(list b�n tr�i m�n h�nh) --> Nh�p s� l��ng c�n s�a (tr�)-->B�m S�a sai s� l��ng" & _
    vbCrLf & "V� d�:" & vbCrLf & Space(5) & "Kh�ch g�i 20 chai Ken, khi t�nh ti�n kh�ch tr� l�i 5 chai" & vbCrLf & "Thao t�c:" & vbCrLf & Space(5) & _
    "Ch�n d�ng 20 chai ken b�n list --> Nh�p 5 v� � s� l��ng --> B�m S�a sai s� l��ng"
    
End Sub

Private Sub cmdExtraPrice_Click()
    lblDescription.Caption = "Gi� ngo�i th�c ��n" & vbCrLf & "B�m Gi� m� --> Nh�p gi� --> ��ng � --> B�m m�n" & _
    vbCrLf & "V� d�:" & vbCrLf & Space(5) & "Ly Cafe �� gi� ch�nh th�c 15,000 mu�n b�n 17,000" & vbCrLf & "Thao t�c:" & vbCrLf & Space(5) & _
    "B�m Gi� m� --> Nh�p 17000 -->��ng � --> B�m Cafe ��"
End Sub

Private Sub cmdGopban_Click()
    lblDescription.Caption = "G�p b�n" & vbCrLf & " M� b�n c�n g�p --> B�m G�p b�n --> Ch�n b�n c�n chuy�n ��n (B�n �� c�)"
End Sub

Private Sub cmdTranferTable_Click()
    lblDescription.Caption = "Chuy�n b�n" & vbCrLf & " M� b�n c�n chuy�n �i --> B�m Chuy�n b�n --> Ch�n b�n c�n chuy�n ��n (B�n tr�ng)"
End Sub
