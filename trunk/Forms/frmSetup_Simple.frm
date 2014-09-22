VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSetup_Simple 
   Caption         =   "Setup"
   ClientHeight    =   9885
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   11610
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
   Icon            =   "frmSetup_Simple.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9885
   ScaleWidth      =   11610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraReport 
      Caption         =   "B¸o c¸o"
      ForeColor       =   &H00FF0000&
      Height          =   7215
      Left            =   2760
      TabIndex        =   39
      Top             =   480
      Width           =   8895
      Begin VB.Frame fraBaocaobanhang 
         Caption         =   "B¸o c¸o b¸n hµng"
         ForeColor       =   &H00FF0000&
         Height          =   6255
         Left            =   2610
         TabIndex        =   74
         Top             =   840
         Width           =   6015
         Begin VB.ComboBox cboSaleFilter 
            BeginProperty Font 
               Name            =   "VNI-Times"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2460
            TabIndex        =   76
            Top             =   420
            Width           =   3435
         End
         Begin VB.ComboBox cboSalesort 
            BeginProperty Font 
               Name            =   "VNI-Times"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2430
            TabIndex        =   75
            Top             =   1170
            Width           =   3465
         End
         Begin prjTouchScreen.MyProgressBar prbBanhang 
            Height          =   315
            Left            =   120
            TabIndex        =   77
            Top             =   5640
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   556
         End
         Begin MSForms.CommandButton cmdHourly 
            Height          =   825
            Left            =   4080
            TabIndex        =   92
            Tag             =   "L21"
            Top             =   2640
            Width           =   1725
            ForeColor       =   16711680
            BackColor       =   16761024
            VariousPropertyBits=   8388635
            Caption         =   "Baùo caùo theo giôø"
            Size            =   "3043;1455"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdDeletedItems 
            Height          =   825
            Left            =   4080
            TabIndex        =   91
            Tag             =   "L60"
            Top             =   1680
            Width           =   1725
            ForeColor       =   16711680
            BackColor       =   16761024
            VariousPropertyBits=   8388635
            Caption         =   "Baùo caùo moùn xoùa"
            Size            =   "3043;1455"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdCashierRepot 
            Height          =   825
            Left            =   2160
            TabIndex        =   90
            Tag             =   "L22"
            Top             =   1680
            Width           =   1725
            ForeColor       =   16711680
            BackColor       =   16761024
            VariousPropertyBits=   8388635
            Caption         =   "Baùo caùo ca"
            Size            =   "3043;1455"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdNotPaymented 
            Height          =   825
            Left            =   2160
            TabIndex        =   89
            Tag             =   "L20"
            Top             =   2640
            Width           =   1725
            ForeColor       =   16711680
            BackColor       =   16761024
            VariousPropertyBits=   8388635
            Caption         =   "Baùo caùo baøn chöa thu"
            Size            =   "3043;1455"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdTable 
            Height          =   825
            Left            =   240
            TabIndex        =   88
            Tag             =   "L19"
            Top             =   4560
            Width           =   1725
            ForeColor       =   16711680
            BackColor       =   16761024
            VariousPropertyBits=   8388635
            Caption         =   "Baùo caùo toång baøn"
            Size            =   "3043;1455"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdTongbill 
            Height          =   825
            Left            =   240
            TabIndex        =   87
            Tag             =   "L18"
            Top             =   3600
            Width           =   1725
            ForeColor       =   16711680
            BackColor       =   16761024
            VariousPropertyBits=   8388635
            Caption         =   "Baùo caùo toång HÑ"
            Size            =   "3043;1455"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdBaocaochitiet 
            Height          =   825
            Left            =   240
            TabIndex        =   86
            Tag             =   "L17"
            Top             =   2640
            Width           =   1725
            ForeColor       =   16711680
            BackColor       =   16761024
            VariousPropertyBits=   8388635
            Caption         =   "Baùo caùo chi tieát"
            Size            =   "3043;1455"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdReportGroup 
            Height          =   825
            Left            =   240
            TabIndex        =   85
            Tag             =   "L16"
            Top             =   1680
            Width           =   1725
            ForeColor       =   16711680
            BackColor       =   16761024
            VariousPropertyBits=   8388635
            Caption         =   "Baùo caùo theo nhoùm"
            Size            =   "3043;1455"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdBaocaotonghop 
            Height          =   825
            Left            =   240
            TabIndex        =   84
            Tag             =   "L15"
            Top             =   720
            Width           =   1725
            ForeColor       =   16711680
            BackColor       =   16761024
            VariousPropertyBits=   8388635
            Caption         =   "Baùo caùo toång hôïp"
            Size            =   "3043;1455"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label2 
            Caption         =   "Chia theo:"
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2160
            TabIndex        =   83
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label3 
            Caption         =   "S¾p xÕp:"
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2220
            TabIndex        =   82
            Top             =   900
            Width           =   1815
         End
         Begin MSForms.CommandButton CommandButton2 
            Height          =   825
            Left            =   2160
            TabIndex        =   81
            Top             =   3600
            Width           =   1725
            ForeColor       =   16711680
            BackColor       =   16761024
            VariousPropertyBits=   8388635
            Caption         =   "Baùo caùo theo nhaân vieân PV"
            Size            =   "3043;1455"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdItemDiscount 
            Height          =   825
            Left            =   4080
            TabIndex        =   80
            Top             =   3600
            Width           =   1725
            ForeColor       =   16711680
            BackColor       =   16761024
            VariousPropertyBits=   8388635
            Caption         =   "Baùo chi tieát giaûm moùn"
            Size            =   "3043;1455"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdMixmatchReport 
            Height          =   825
            Left            =   2160
            TabIndex        =   79
            Top             =   4560
            Width           =   1725
            ForeColor       =   16711680
            BackColor       =   16761024
            VariousPropertyBits=   8388635
            Caption         =   "BC Chi tieát chieát khaáu"
            Size            =   "3043;1455"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdItemCard 
            Height          =   825
            Left            =   4080
            TabIndex        =   78
            Top             =   4560
            Width           =   1725
            ForeColor       =   16711680
            BackColor       =   16761024
            Caption         =   "Baûn keâ chi tieát"
            Size            =   "3043;1455"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin VB.Frame fraKitchen 
         Caption         =   "B¶n kª order "
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
         Height          =   6255
         Left            =   2610
         TabIndex        =   52
         Top             =   840
         Width           =   6015
         Begin VB.ComboBox cboKit_Sort 
            BeginProperty Font 
               Name            =   "VNI-Times"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2880
            TabIndex        =   55
            Text            =   "cbo_sort"
            Top             =   3120
            Width           =   3015
         End
         Begin VB.ComboBox cboKit_Filter 
            BeginProperty Font 
               Name            =   "VNI-Times"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2880
            TabIndex        =   54
            Text            =   "cbo_sort"
            Top             =   2040
            Width           =   3015
         End
         Begin VB.ComboBox cboKit_Printer 
            BeginProperty Font 
               Name            =   ".VnArial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2880
            TabIndex        =   53
            Text            =   "cboKit_Printer"
            Top             =   840
            Width           =   3015
         End
         Begin prjTouchScreen.MyProgressBar MyProgressBar2 
            Height          =   315
            Left            =   120
            TabIndex        =   56
            Top             =   5760
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   556
         End
         Begin VB.Label lblKit_Sort 
            Caption         =   "S¾p xÕp"
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
            Height          =   375
            Left            =   2880
            TabIndex        =   61
            Top             =   2760
            Width           =   1935
         End
         Begin VB.Label lbl_Kit_filter 
            Caption         =   "Chia theo"
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
            Height          =   375
            Left            =   2880
            TabIndex        =   60
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label lblKit_Printer 
            Caption         =   "Chän m¸y in"
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
            Height          =   375
            Left            =   2880
            TabIndex        =   59
            Top             =   360
            Width           =   2535
         End
         Begin MSForms.CommandButton cmdKit_General 
            Height          =   975
            Left            =   240
            TabIndex        =   58
            Tag             =   "L91"
            Top             =   1080
            Width           =   2415
            ForeColor       =   16711680
            BackColor       =   16761024
            Caption         =   "Baûng keâ toång hôïp"
            Size            =   "4260;1720"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdOrderDetail80 
            Height          =   975
            Left            =   240
            TabIndex        =   57
            Tag             =   "L92"
            Top             =   2400
            Width           =   2415
            ForeColor       =   16711680
            BackColor       =   16761024
            Caption         =   "In baûng keâ Order"
            Size            =   "4260;1720"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin VB.Frame fra80 
         Height          =   6255
         Left            =   2610
         TabIndex        =   40
         Top             =   840
         Width           =   6015
         Begin VB.ComboBox cbo80Filter 
            BeginProperty Font 
               Name            =   "VNI-Times"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2970
            TabIndex        =   42
            Top             =   660
            Width           =   2475
         End
         Begin VB.ComboBox cbo80Sort 
            BeginProperty Font 
               Name            =   "VNI-Times"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2940
            TabIndex        =   41
            Top             =   1410
            Width           =   2505
         End
         Begin MSForms.CommandButton cmdEmp_Totals 
            Height          =   855
            Left            =   3000
            TabIndex        =   94
            Top             =   5040
            Width           =   2415
            ForeColor       =   16711680
            BackColor       =   16761024
            VariousPropertyBits=   8388635
            Caption         =   "Baùo caùo doanh soá  nhaân vieân phuïc vuï"
            Size            =   "4260;1508"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdMainGroup80 
            Height          =   855
            Left            =   240
            TabIndex        =   73
            Top             =   1200
            Width           =   2415
            ForeColor       =   16711680
            BackColor       =   16761024
            Caption         =   "Baùo caùo nhoùm chính"
            Size            =   "4260;1508"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdReport_Location 
            Height          =   855
            Left            =   3000
            TabIndex        =   67
            Top             =   2160
            Width           =   2415
            ForeColor       =   16711680
            BackColor       =   16761024
            Caption         =   "Doanh thu theo khu"
            Size            =   "4260;1508"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin VB.Label Label8 
            Caption         =   "Läc theo:"
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   2970
            TabIndex        =   51
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "S¾p xÕp:"
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   2880
            TabIndex        =   50
            Top             =   1140
            Width           =   915
         End
         Begin MSForms.CommandButton cmdGeneral80 
            Height          =   855
            Left            =   240
            TabIndex        =   49
            Tag             =   "L15"
            Top             =   240
            Width           =   2415
            ForeColor       =   16711680
            BackColor       =   16761024
            Caption         =   "Baùo caùo toång hôïp"
            Size            =   "4260;1508"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdDetail80 
            Height          =   855
            Left            =   240
            TabIndex        =   48
            Tag             =   "L17"
            Top             =   2160
            Width           =   2415
            ForeColor       =   16711680
            BackColor       =   16761024
            Caption         =   "Baùo caùo chi tieát"
            Size            =   "4260;1508"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdBill80 
            Height          =   855
            Left            =   240
            TabIndex        =   47
            Tag             =   "L18"
            Top             =   3120
            Width           =   2415
            ForeColor       =   16711680
            BackColor       =   16761024
            Caption         =   "Baùo caùo toång phieáu"
            Size            =   "4260;1508"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdBanchuathu80 
            Height          =   855
            Left            =   240
            TabIndex        =   46
            Tag             =   "L20"
            Top             =   4080
            Width           =   2415
            ForeColor       =   16711680
            BackColor       =   16761024
            Caption         =   "Baùo caùo baøn chöa thu"
            Size            =   "4260;1508"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdCashier80 
            Height          =   855
            Left            =   3000
            TabIndex        =   45
            Top             =   4080
            Width           =   2415
            ForeColor       =   16711680
            BackColor       =   16761024
            Caption         =   "Baùo caùo ca"
            Size            =   "4260;1508"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton CmdDetailEmployee 
            Height          =   855
            Left            =   240
            TabIndex        =   44
            Tag             =   "L93"
            Top             =   5040
            Width           =   2415
            ForeColor       =   16711680
            BackColor       =   16761024
            VariousPropertyBits=   8388635
            Caption         =   "Baùo caùo chi tieát theo nhaân vieân phuïc vuï"
            Size            =   "4260;1508"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
         Begin MSForms.CommandButton cmdBillList 
            Height          =   855
            Left            =   3000
            TabIndex        =   43
            Top             =   3120
            Width           =   2415
            ForeColor       =   16711680
            BackColor       =   16761024
            Caption         =   "Danh saùch hoùa ñôn"
            Size            =   "4260;1508"
            FontName        =   "VNI-Times"
            FontHeight      =   240
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
         End
      End
      Begin MSComCtl2.DTPicker dtpFromDate 
         Height          =   495
         Left            =   4020
         TabIndex        =   62
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   18022401
         UpDown          =   -1  'True
         CurrentDate     =   36892
      End
      Begin MSComCtl2.DTPicker dtpToDate 
         Height          =   495
         Left            =   6780
         TabIndex        =   63
         Top             =   210
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   18022401
         UpDown          =   -1  'True
         CurrentDate     =   36892
      End
      Begin prjTouchScreen.MyButton cmdReport80 
         Height          =   1335
         Left            =   330
         TabIndex        =   64
         Top             =   2805
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   2355
         BTYPE           =   1
         TX              =   "In Baùo caùo 58mm"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "VNI-Times"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421376
         BCOLO           =   33023
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSetup_Simple.frx":000C
         UMCOL           =   0   'False
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdSaleReport 
         Height          =   1335
         Left            =   330
         TabIndex        =   96
         Tag             =   "L6"
         Top             =   1200
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   2355
         BTYPE           =   1
         TX              =   "Baùo caùo baùn haøng A4"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "VNI-Times"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421376
         BCOLO           =   33023
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSetup_Simple.frx":0028
         UMCOL           =   0   'False
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin prjTouchScreen.MyButton cmdKitchen_List 
         Height          =   1335
         Left            =   330
         TabIndex        =   97
         Tag             =   "L87"
         Top             =   4440
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   2355
         BTYPE           =   1
         TX              =   "Baûn keâ phieáu order"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "VNI-Times"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   8421376
         BCOLO           =   33023
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSetup_Simple.frx":0044
         UMCOL           =   0   'False
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         Value           =   0   'False
      End
      Begin VB.Label lblDenngay 
         Caption         =   "§Õn ngµy:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5580
         TabIndex        =   66
         Top             =   300
         Width           =   1185
      End
      Begin VB.Label lblFromdate 
         Caption         =   "Tõ ngµy:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   65
         Top             =   330
         Width           =   1005
      End
   End
   Begin VB.Frame fraEditName 
      Caption         =   "Hieäu chænh teân"
      BeginProperty Font 
         Name            =   "VNI-Times"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7215
      Left            =   2760
      TabIndex        =   31
      Top             =   480
      Width           =   8895
      Begin MSForms.CommandButton cmdLocationName 
         Height          =   1185
         Left            =   600
         TabIndex        =   33
         Tag             =   "L80"
         Top             =   600
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Söûa teân Khu vöïc"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdPrinterName 
         Height          =   1185
         Left            =   3600
         TabIndex        =   32
         Tag             =   "L79"
         Top             =   600
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Söûa teân maùy in"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame fraCashier 
      Caption         =   "Cashier Maintenance"
      ForeColor       =   &H00FF0000&
      Height          =   7215
      Left            =   2760
      TabIndex        =   2
      Top             =   480
      Width           =   8895
      Begin MSForms.CommandButton cmdAttenden 
         Height          =   1185
         Left            =   600
         TabIndex        =   68
         Top             =   5520
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Chaám coâng"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSalary 
         Height          =   1185
         Left            =   6000
         TabIndex        =   18
         Tag             =   "L102"
         Top             =   3840
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Baûng möùc löông"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdEmployee 
         Height          =   1185
         Left            =   6000
         TabIndex        =   17
         Tag             =   "L99"
         Top             =   2160
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Danh saùch nhaân vieân"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdInvoiceHoldList 
         Height          =   1185
         Left            =   6120
         TabIndex        =   16
         Tag             =   "L105"
         Top             =   5520
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Chia nh©n viªn theo khu vùc"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdJobCode 
         Height          =   1185
         Left            =   3360
         TabIndex        =   15
         Tag             =   "L101"
         Top             =   3840
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Baûng coâng vieäc"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdChangeID 
         Height          =   1185
         Left            =   6000
         TabIndex        =   14
         Tag             =   "L96"
         Top             =   600
         Visible         =   0   'False
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Ñoåi maõ soá ñaêng nhaäp"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdInvoiceList 
         Height          =   1185
         Left            =   3360
         TabIndex        =   13
         Tag             =   "L104"
         Top             =   5520
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Danh saùch HÑ Thanh Toaùn"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdChangpass 
         Height          =   1185
         Left            =   3360
         TabIndex        =   12
         Tag             =   "L95"
         Top             =   600
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Ñoåi maät khaåu"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmddept 
         Height          =   1185
         Left            =   3360
         TabIndex        =   11
         Tag             =   "L98"
         Top             =   2160
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Phoøng ban"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdShift 
         Height          =   1185
         Left            =   600
         TabIndex        =   10
         Tag             =   "L100"
         Top             =   3840
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Ca laøm vieäc"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdRightSelection 
         Height          =   1185
         Left            =   600
         TabIndex        =   9
         Tag             =   "L97"
         Top             =   2160
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Phaân quyeàn"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdAddCashier 
         Height          =   1185
         Left            =   600
         TabIndex        =   8
         Tag             =   "L94"
         Top             =   600
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Danh saùch Ngöôøi duøng"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame fraList 
      Caption         =   "Cµi ®Æt danh môc"
      ForeColor       =   &H00FF0000&
      Height          =   7215
      Left            =   2760
      TabIndex        =   1
      Top             =   480
      Width           =   8895
      Begin MSForms.CommandButton cmdCust_Type 
         Height          =   1185
         Left            =   240
         TabIndex        =   95
         Top             =   1800
         Width           =   2580
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Danh muïc nhoùm khaùch haøng"
         Size            =   "4551;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmddiscountReason 
         Height          =   1185
         Left            =   3120
         TabIndex        =   72
         Top             =   4560
         Width           =   2580
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Danh muïc lyù do chieát khaáu"
         Size            =   "4551;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdChietkhau 
         Height          =   1185
         Left            =   240
         TabIndex        =   71
         Top             =   4560
         Width           =   2580
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Danh muïc chieát khaáu"
         Size            =   "4551;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdInstruction 
         Height          =   1185
         Left            =   240
         TabIndex        =   38
         Tag             =   "L75"
         Top             =   3120
         Width           =   2580
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Thoâng tin chuù thích cheá bieán"
         Size            =   "4551;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdGiftCard 
         Height          =   1185
         Left            =   6000
         TabIndex        =   34
         Tag             =   "L42"
         Top             =   3120
         Width           =   2580
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Phieáu quaø taëng"
         Size            =   "4551;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdMedia 
         Height          =   1185
         Left            =   3120
         TabIndex        =   24
         Tag             =   "L53"
         Top             =   3120
         Width           =   2580
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Danh muïc ngoaïi teä"
         Size            =   "4551;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdCustomer 
         Height          =   1185
         Left            =   3120
         TabIndex        =   23
         Tag             =   "L47"
         Top             =   1800
         Width           =   2580
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Danh muïc khaùch haøng"
         Size            =   "4551;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdVendor 
         Height          =   1185
         Left            =   6000
         TabIndex        =   22
         Tag             =   "L48"
         Top             =   1800
         Width           =   2580
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Danh muïc nhaø cung caáp"
         Size            =   "4551;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdMaingroup 
         Height          =   1185
         Left            =   240
         TabIndex        =   21
         Tag             =   "L43"
         Top             =   480
         Width           =   2580
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Danh muïc nhoùm chính"
         Size            =   "4551;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdGroup 
         Height          =   1185
         Left            =   3120
         TabIndex        =   20
         Tag             =   "L44"
         Top             =   480
         Width           =   2580
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Danh muïc nhoùm haøng"
         Size            =   "4551;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSKU 
         Height          =   1185
         Left            =   6000
         TabIndex        =   19
         Tag             =   "L45"
         Top             =   480
         Width           =   2580
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Danh muïc haøng hoùa"
         Size            =   "4551;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.Frame fraSys 
      Caption         =   "Cµi ®Æt hÖ thèng"
      ForeColor       =   &H00FF0000&
      Height          =   7215
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   8895
      Begin MSForms.CommandButton cmdVAT 
         Height          =   1185
         Left            =   600
         TabIndex        =   93
         Tag             =   "L110"
         Top             =   5760
         Visible         =   0   'False
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Keát xuaát HÑ VAT"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdConnect 
         Height          =   1185
         Left            =   6240
         TabIndex        =   70
         Tag             =   "L109"
         Top             =   3960
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Keát noái döõ lieäu maùy con"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdTimerSetup 
         Height          =   1185
         Left            =   6240
         TabIndex        =   37
         Tag             =   "L54"
         Top             =   480
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Caøi ñaët hieån thò"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdInvoiceSetup 
         Height          =   1185
         Left            =   3360
         TabIndex        =   36
         Tag             =   "L52"
         Top             =   480
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Thoâng tin ñaàu cuoái bill"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdLayout 
         Height          =   1185
         Left            =   600
         TabIndex        =   35
         Tag             =   "L106"
         Top             =   3960
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Caøi ñaët giao dieän"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdPriceRate 
         Height          =   1185
         Left            =   6210
         TabIndex        =   29
         Tag             =   "L41"
         Top             =   2160
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Chính saùch giaù baùn"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdAdjustment 
         Height          =   1185
         Left            =   3330
         TabIndex        =   28
         Tag             =   "L37"
         Top             =   2160
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Möùc giaûm %"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSetupPrint 
         Height          =   1185
         Left            =   570
         TabIndex        =   27
         Tag             =   "L38"
         Top             =   2160
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Maëc ñònh maùy in"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdTaxRate 
         Height          =   1185
         Left            =   3360
         TabIndex        =   26
         Tag             =   "L39"
         Top             =   3960
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Tyû leä thueá"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdSystemFlag 
         Height          =   1185
         Left            =   570
         TabIndex        =   25
         Tag             =   "L31"
         Top             =   480
         Width           =   2100
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Caáu hình heä thoáng"
         Size            =   "3704;2090"
         FontName        =   "VNI-Times"
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin MSForms.CommandButton cmdTichluy 
      Height          =   855
      Left            =   9840
      TabIndex        =   69
      Top             =   7560
      Visible         =   0   'False
      Width           =   1695
      ForeColor       =   255
      BackColor       =   14737632
      VariousPropertyBits=   8388635
      Caption         =   "BC Tích luõy"
      Size            =   "2990;1508"
      FontName        =   "VNI-Times"
      FontHeight      =   240
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdEditName 
      Height          =   1245
      Left            =   120
      TabIndex        =   30
      Tag             =   "L90"
      Top             =   5280
      Width           =   2130
      ForeColor       =   16711680
      BackColor       =   14737632
      VariousPropertyBits=   8388635
      Caption         =   "HIEÄU CHÆNH TEÂN"
      Size            =   "3757;2205"
      FontName        =   "VNI-Times"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdReport 
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Tag             =   "L4"
      Top             =   6840
      Width           =   2130
      ForeColor       =   16711680
      BackColor       =   14737632
      Caption         =   "BAÙO CAÙO"
      Size            =   "3757;2143"
      FontName        =   "VNI-Times"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdAdministrative 
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Tag             =   "L3"
      Top             =   3720
      Width           =   2130
      ForeColor       =   16711680
      BackColor       =   14737632
      VariousPropertyBits=   8388635
      Caption         =   "CAØI ÑAËT DANH MUÏC"
      Size            =   "3757;2143"
      FontName        =   "VNI-Times"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdSetup 
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Tag             =   "L2"
      Top             =   2160
      Width           =   2130
      ForeColor       =   16711680
      BackColor       =   14737632
      VariousPropertyBits=   8388635
      Caption         =   "CAÁU HÌNH HEÄ THOÁNG"
      Size            =   "3757;2143"
      FontName        =   "VNI-Times"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdCashier 
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Tag             =   "L1"
      Top             =   600
      Width           =   2130
      ForeColor       =   16711680
      BackColor       =   14737632
      VariousPropertyBits=   8388635
      Caption         =   "NHAÂN VIEÂN"
      Size            =   "3757;2143"
      FontName        =   "VNI-Times"
      FontHeight      =   285
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdDone 
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Tag             =   "L59"
      Top             =   8400
      Width           =   11415
      ForeColor       =   16711680
      BackColor       =   14737632
      Caption         =   "HOAØN TAÁT"
      Size            =   "20135;2566"
      FontName        =   "VNI-Times"
      FontHeight      =   405
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "frmSetup_Simple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iReport As CRAXDDRT.Report
Dim DescArr() As String
Dim DescArrReport() As String
Dim Printer_ID As String
Dim strDateTime As String

Private Sub cboKit_Printer_Change()
On Error GoTo Handle
   Printer_ID = Format(cboKit_Printer.ListIndex, "00")
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "cboKit_Printer_Change"
End Sub

Private Sub cboKit_Printer_Click()
    Call cboKit_Printer_Change
End Sub

Private Sub cmd_Click()
    frmStock_List.Show vbModal
End Sub

Private Sub cmdAttenden_Click()
    'frmAttendent.Show vbModal
    frmChamcong.Show vbModal
End Sub

Private Sub cmdBanchuathu80_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim CRReport As CRAXDDRT.Report
    
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = "SELECT Invoice_Totals_Notes.Invoice_Number, Invoice_Totals.Orig_OnHoldID," & _
    "  Left([DateTime],8) AS DateOpen, Invoice_Totals.Grand_Total," & _
    " Invoice_Totals.Station_ID, Invoice_Totals.Status, Invoice_Totals.Cashier_ID" & _
    " FROM Invoice_Totals INNER JOIN Invoice_Totals_Notes ON (Invoice_Totals.Store_ID = Invoice_Totals_Notes.Store_ID)" & _
    " AND (Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number)" & _
    " where Invoice_Totals.Status = 'O'  or Invoice_Totals.Status= 'P' and " & _
    " Left(Invoice_Totals.[DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' and Left(Invoice_Totals.[DateTime],8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'" & _
    " order by Invoice_Totals.Invoice_Number"
    Set crBanchuathu80 = Nothing
    Set crBanchuathu58 = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    If ReceiptType = 80 Then
        Set CRReport = crBanchuathu80
    Else
        Set CRReport = crBanchuathu58
    End If
        
    With CRReport
        .Database.AddADOCommand cnData, cmd
        .txtTableNo.SetUnboundFieldSource "{ado.Orig_OnHoldID}"
        .txtBillNo.SetUnboundFieldSource "{ado.Invoice_Number}"
        .txtserver.SetUnboundFieldSource "{ado.Station_ID}"
        .txtAmt.SetUnboundFieldSource "{ado.Grand_Total}"
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        'canh le
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign
        With .txtAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtsumAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    Set iReport = CRReport
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
Exit Sub
MsgBox Err.Number & Err.Description & Me.name & " cmdTongbill_Click"
End Sub

Private Sub cmdBangke_Click()
    frmList_Detail.Show vbModal
End Sub

Private Sub cmdBill80_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim CRReport As CRAXDDRT.Report
    
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop
'
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    
    SQL = "SELECT right('0000' & Invoice_Totals.Invoice_Number,4)as billNo,Invoice_Totals.Orig_OnHoldID as TableNo," & _
        "  Invoice_Totals.Payment_Method,Invoice_Totals.InvType,Invoice_Totals.Store_ID, Invoice_Totals.Station_ID," & _
        " Invoice_Totals.CustNum, Left([DateTime],8) AS DateInvoice, Invoice_Totals.Total_Cost, Invoice_Totals.Total_Price," & _
        " Invoice_Totals.Adjustment1,Invoice_Totals.Adjustment2,(Invoice_Totals.Total_Tax1*Invoice_Totals.VATFee/100)as VAT, " & _
        " Invoice_Totals.Discount, Invoice_Totals.Grand_Total, Invoice_Totals.AddMoney, Invoice_Totals.Cashier_ID," & _
        " (Invoice_Totals.Total_Price*Invoice_Totals.Service_Charge/100)as Service, Invoice_Totals.Personals" & _
      " from Invoice_Totals " & _
      " WHERE Left([DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And Left([DateTime], 8) <= '" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'" & _
      " and Invoice_Totals.Status<>'CO' and  Invoice_Totals.Status<>'O' and left(Invoice_Totals.Status,1)<>'T' " & _
      " GROUP BY Invoice_Totals.Invoice_Number, Invoice_Totals.Orig_OnHoldID,Invoice_Totals.Store_ID,Invoice_Totals.InvType,Invoice_Totals.Payment_Method, Invoice_Totals.Station_ID, Invoice_Totals.CustNum, Left([DateTime],8), Invoice_Totals.Total_Cost, Invoice_Totals.Total_Price," & _
      " Invoice_Totals.Discount, Invoice_Totals.Grand_Total, Invoice_Totals.AddMoney, Invoice_Totals.Cashier_ID," & _
      " Invoice_Totals.Personals,(Invoice_Totals.Total_Tax1*Invoice_Totals.VATFee/100),(Invoice_Totals.Total_Price*Invoice_Totals.Service_Charge/100),Adjustment2,Adjustment1"
    
    Set crBill80 = Nothing
    Set crBill58 = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    If ReceiptType = 80 Then
        Set CRReport = crBill80
    Else
        Set CRReport = crBill58
    End If
        
    With CRReport
        .Database.AddADOCommand cnData, cmd
        .txtBillDate.SetUnboundFieldSource "{ado.DateInvoice}"
        .txtBillNo.SetUnboundFieldSource "{ado.BillNo}"
        .TxtTotal.SetUnboundFieldSource "{ado.Total_Price}"
        .txtDiscount.SetUnboundFieldSource "{ado.Discount}"
        .txtTotalAmt.SetUnboundFieldSource "{ado.Grand_Total}"
'        .txtSokhach.SetUnboundFieldSource "{ado.SeatNum}"
        .txtReceiveMoney.SetUnboundFieldSource "{ado.AddMoney}"
        'canh le
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign
        If cboSaleFilter.ListIndex = 1 Then
            .txtStoreID.SetUnboundFieldSource "{ado.Store_ID}"
            .Section3.Suppress = False
            .Section4.Suppress = False
            .Section11.Suppress = True
            .Section5.Suppress = True
        ElseIf cboSaleFilter.ListIndex = 2 Then
            .txtStationID.SetUnboundFieldSource "{ado.Station_ID}"
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section11.Suppress = False
            .Section5.Suppress = False
        Else
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section5.Suppress = True
            .Section11.Suppress = True
        End If
        
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        
        With .txtTotalStore
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtDisAmountStore
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtDisAmountStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalAmtStore
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalAmtStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        
        End With
        With .txtAmtDis
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .TxtTotal
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtDiscount
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Field1
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
       
       With .txtSumTotalAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
    End With
    Set iReport = CRReport
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub

Handle:
Exit Sub
MsgBox Err.Number & Err.Description & Me.name & " cmdTongbill_Click"
End Sub

Private Sub cmdBillList_Click()
On Error GoTo errHdl
    Dim CRReport As New CRAXDDRT.Report
    Dim SQL As String
    Dim cmd As New ADODB.Command
    Dim FromDate, ToDate As String
    
    FromDate = Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00")
    ToDate = Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00")

        SQL = "SELECT Invoice_Totals.Invoice_Number, Invoice_Totals.CustNum," & _
        " Invoice_Totals.Discount, Invoice_Totals.Total_Price,left(Invoice_Totals_Notes.ClosingTime,8) as DateInvoice," & _
        " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adjustment1," & _
        " Invoice_Totals.Adjustment2,Invoice_Totals.Adjustment3," & _
        " Invoice_Totals.Adjustment4,Invoice_Totals.AddMoney,substring(Invoice_Totals_Notes.ClosingTime,9,8)as TimeInvoice," & _
        " Invoice_Totals.Grand_Total, Invoice_Totals.Amt_Tendered," & _
        " Invoice_Totals.Cashier_ID,Invoice_Totals.OrderMan," & _
        " Invoice_Totals.Station_ID,Invoice_Totals.Payment_Method,Invoice_Itemized.ItemNum, " & _
        " Sum(Invoice_Itemized.Quantity) AS Qty, Invoice_Itemized.PricePer," & _
        " sum(Invoice_Itemized.Quantity*Invoice_Itemized.PricePer) as Amt, " & _
        " Invoice_Itemized.DiffItemName ,Invoice_Totals.Orig_OnHoldID,MainGroup.GroupNo " & _
        " FROM  ((Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) INNER JOIN " & _
        " (Inventory INNER JOIN (Departments INNER JOIN MainGroup ON Departments.MainGroup = MainGroup.GroupNo) ON Inventory.Dept_ID = Departments.Dept_ID) ON Invoice_Itemized.ItemNum = Inventory.ItemNum)" & _
        " INNER JOIN Invoice_Totals_Notes ON Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number " & _
        " Where left(Invoice_Totals.DateTime,8)>='" & FromDate & "' and left(Invoice_Totals.DateTime,8)<='" & ToDate & "'" & _
        " group by Invoice_Itemized.ItemNum,Invoice_Totals.Invoice_Number," & _
        " Invoice_Totals.CustNum,Invoice_Totals.Discount," & _
        " Invoice_Totals.Total_Price,Invoice_Totals.Grand_Total,left(Invoice_Totals_Notes.ClosingTime,8)," & _
        " Invoice_Totals.Amt_Tendered,Invoice_Totals.Amt_Change,substring(Invoice_Totals_Notes.ClosingTime,9,8)," & _
        " Invoice_Totals.Cashier_ID, Invoice_Totals.Station_ID,Invoice_Totals.OrderMan," & _
        " Invoice_Itemized.PricePer, Invoice_Itemized.DiffItemName ,Invoice_Totals.Orig_OnHoldID,Invoice_Totals.Payment_Method, " & _
        " Invoice_Totals.Service_Charge,Invoice_Totals.VATFee,Invoice_Totals.Adjustment1,Invoice_Totals.Adjustment2,Invoice_Totals.Adjustment3," & _
        " Invoice_Totals.Adjustment4,Invoice_Totals.AddMoney,MainGroup.GroupNo,  substring([ClosingTime],9,8)" & _
        " order by Invoice_Itemized.PricePer"
   
    Set crBillList = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crBillList
        .Database.AddADOCommand cnData, cmd
        .txtPluCode.SetUnboundFieldSource "{ado.DiffItemName}"
        .PluName.SetUnboundFieldSource "{ado.ItemNum}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtCost.SetUnboundFieldSource "{ado.PricePer}"
        .txtAmt.SetUnboundFieldSource "{ado.Amt}"
        .txtBillNo.SetUnboundFieldSource "{ado.Invoice_Number}"
        .txtCashier.SetUnboundFieldSource "{ado.Cashier_ID}"
'        .txtTenderAmt.SetUnboundFieldSource "{ado.Amt_Tendered}"
        .txtNVPV.SetUnboundFieldSource "{ado.Orderman}"
        .txtDiscount.SetUnboundFieldSource "{ado.Discount}"
        .txtTable.SetUnboundFieldSource "{ado.Orig_OnHoldID}"
        .txtserver.SetUnboundFieldSource "{ado.Station_ID}"
'        .txtMethod.SetUnboundFieldSource "{ado.Payment_Method}"
        .txtPayment.SetUnboundFieldSource "{ado.Grand_Total}"
        .txtSerCharge.SetUnboundFieldSource "{ado.Service_Charge}"
'        .txtVAT.SetUnboundFieldSource "{ado.VATFee}"
        .txtAdjustment1.SetUnboundFieldSource "{ado.Adjustment1}"
        .txtAdjustment2.SetUnboundFieldSource "{ado.Adjustment2}"
        .txtReceiveMoney.SetUnboundFieldSource "{ado.AddMoney}"
        .txtDate.SetUnboundFieldSource "{ado.DateInvoice}"
        .txtTime.SetUnboundFieldSource "{ado.TimeInvoice}"
        .txtCustomer.SetUnboundFieldSource "{ado.CustNum}"
        .txtMaingroup.SetUnboundFieldSource "{ado.GroupNo}"
        With .txtReceiveMoney
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAdjustment1
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAdjustment2
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtPayment
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .AmtDisCount
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtCost
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAmtCharge
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtsumAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtPayment
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtSumMainGroup
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    
    Set iReport = crBillList
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = iReport
        .Show vbModal
    End With
Exit Sub
errHdl:
Exit Sub
    MsgBox Err.Number & " - cmdBillList_Click - " & Err.Description
End Sub

Private Sub cmdCalculate_Click()
    frmCal_TonTemp.Show vbModal
End Sub

Private Sub cmdCashier80_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim CRReport As CRAXDDRT.Report
    Dim fTime, tTime As String
    With frmTimeReport
        .Show vbModal
        fTime = .GetFTime
        tTime = .GetTTime
    End With
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop
'
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
     If UserLevel <> 1 Then
        SQL = "SELECT Count(Invoice_Totals.Invoice_Number) AS [Transaction]," & _
        " Count(Invoice_Totals.Discount) AS CountDis, " & _
        " Sum(Invoice_Totals.Total_Price) AS SumTP,Sum(Invoice_Totals.Adjustment1) AS Adj1,Sum(Invoice_Totals.Adjustment2) AS Adj2," & _
        " Sum(Invoice_Totals.Grand_Total) AS sumGT, " & _
        " Sum(Invoice_Totals.Adjustment3) AS Adj3,Sum(Invoice_Totals.Adjustment4) AS Adj4,Sum(Invoice_Totals.Adjustment5) AS Adj5,Sum(Invoice_Totals.Adjustment6) AS Adj6," & _
        " Invoice_Totals.Cashier_ID," & _
        " sum(Invoice_Totals.Discount*Invoice_Totals.Total_Price/100) AS AmtDis," & _
        " sum(Invoice_Totals.Service_Charge*Invoice_Totals.Total_Price/100) AS AmtSer," & _
        " sum(Invoice_Totals.VATFee*Invoice_Totals.Total_Tax1/100) AS AmtVAT," & _
        " sum(Invoice_Totals.AddMoney) AS AmtReceive,Left([DateTime],8)as Datesale " & _
        " From Invoice_Totals" & _
        " where Left([DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And Left([DateTime],8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "' and Invoice_Totals.Cashier_ID='" & UserID & "' and Invoice_Totals.status='C'" & _
        " GROUP BY Invoice_Totals.Cashier_ID, Left([DateTime],8)"
    Else
        SQL = "SELECT Count(Invoice_Totals.Invoice_Number) AS [Transaction]," & _
        " Count(Invoice_Totals.Discount) AS CountDis, " & _
        " Sum(Invoice_Totals.Total_Price) AS SumTP,Sum(Invoice_Totals.Adjustment1) AS Adj1,Sum(Invoice_Totals.Adjustment2) AS Adj2, Sum(Invoice_Totals.Grand_Total)" & _
        " AS sumGT, Invoice_Totals.Cashier_ID," & _
         " Sum(Invoice_Totals.Adjustment3) AS Adj3,Sum(Invoice_Totals.Adjustment4) AS Adj4,Sum(Invoice_Totals.Adjustment5) AS Adj5,Sum(Invoice_Totals.Adjustment6) AS Adj6," & _
        " sum(Invoice_Totals.Discount*Invoice_Totals.Total_Price/100) AS AmtDis," & _
        " sum(Invoice_Totals.Service_Charge*Invoice_Totals.Total_Price/100) AS AmtSer," & _
        " sum(Invoice_Totals.VATFee*Invoice_Totals.Total_Tax1/100) AS AmtVAT," & _
        " sum(Invoice_Totals.AddMoney) AS AmtReceive,Left([DateTime],8)as Datesale " & _
        " From Invoice_Totals" & _
        " where Left([DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And Left([DateTime],8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "' and Invoice_Totals.status='C'" & _
        " GROUP BY Invoice_Totals.Cashier_ID, Left([DateTime],8)"
    End If
    Set crCashierReport80 = Nothing
    Set crCashierReport58 = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
     If ReceiptType = 80 Then
        Set CRReport = crCashierReport80
    Else
        Set CRReport = crCashierReport58
    End If
        
    With CRReport
        .Database.AddADOCommand cnData, cmd
        .Transaction.SetUnboundFieldSource "{ado.Transaction}"
        .Total.SetUnboundFieldSource "{ado.SumTP}"
        .CountDis.SetUnboundFieldSource "{ado.CountDis}"
        .txtReceiveMoney.SetUnboundFieldSource "{ado.AmtReceive}"
        .txtVAT.SetUnboundFieldSource "{ado.AmtVAT}"
        .DisAmt.SetUnboundFieldSource "{ado.AmtDis}"
        .Datesale.SetUnboundFieldSource "{ado.Datesale}"
        .CashierID.SetUnboundFieldSource "{ado.Cashier_ID}"
        .Adjustment1.SetUnboundFieldSource "{ado.Adj1}"
        .Adjustment2.SetUnboundFieldSource "{ado.Adj2}"
        .Adjustment3.SetUnboundFieldSource "{ado.Adj3}"
        .Adjustment4.SetUnboundFieldSource "{ado.Adj4}"
        .Adjustment5.SetUnboundFieldSource "{ado.Adj5}"
        .Adjustment6.SetUnboundFieldSource "{ado.Adj6}"
        .Sercharge.SetUnboundFieldSource "{ado.AmtSer}"
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        'canh le
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign
        
        With .Total
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .DisAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtReceiveMoney
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .AmtTotal
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
    End With
    Set iReport = CRReport
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = iReport
        .Show vbModal
    End With
    Unload frmShowCashierReport
    prbBanhang.Value = 0
Exit Sub

Handle:
Exit Sub
MsgBox Err.Number & Err.Description & Me.name & " cmdCashierRepot_Click"

End Sub

Private Sub cmdChietkhau_Click()
    frmMixmatch.Show vbModal
End Sub

'Private Sub cmdConnect_Click()
'    frmConnectClient.Show vbModal
'End Sub

Private Sub cmdCust_Type_Click()
    On Error GoTo Handle
    If Check_Table_exist("Customer_Type") Then
        frmCustom_type.Show vbModal
    Else
        Call Create_Customer_Type
        MsgBox "You must open the Database, open Customer table with Design View mode, Rename Discount field by Cust_Type with data length=20"
    End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdCust_Type_Click"
End Sub

Private Sub cmdDebt_Click()
    MsgBox "Please wait to Complete, thanks!"
End Sub

Private Sub cmdDetail80_Click()
On Error GoTo Handle
If cbo80Filter.ListIndex <> 3 Then
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL, SQL2, SQL3, SQL4, SQL5, SQLSort As String
    Dim CRReport As CRAXDDRT.Report

    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop
    Select Case cbo80Sort.ListIndex
        Case 0: SQLSort = " Order by Invoice_Itemized.ItemNum  ASC"
        Case 1: SQLSort = " Order by Invoice_Itemized.DiffItemName  ASC"
        Case 2: SQLSort = " Order by sum(Invoice_Itemized.Quantity)  DESC"
        Case Else: SQLSort = " Order by Invoice_Itemized.ItemNum  ASC"
    End Select
    
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    'Khong danh cho karaoke
    SQL = "SELECT Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime" & _
          " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number " & _
          " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
          " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'" & _
          " GROUP BY Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer" & SQLSort

    SQL2 = "SELECT Invoice_Itemized.ItemNum, Invoice_Totals.Store_ID, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime" & _
           " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number" & _
            " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
            " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'" & _
          " GROUP BY Invoice_Itemized.ItemNum, Invoice_Totals.Store_ID, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer" & SQLSort

    SQL3 = "SELECT Invoice_Itemized.ItemNum, Invoice_Totals.Station_ID, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime" & _
            " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number" & _
        " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
        " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'" & _
        " GROUP BY Invoice_Itemized.ItemNum, Invoice_Totals.Station_ID, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer" & SQLSort

    SQL4 = "SELECT Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime, Inventory.Dept_ID, Departments.Description" & _
            " FROM Departments INNER JOIN (Inventory INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Departments.Dept_ID = Inventory.Dept_ID" & _
            " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
            " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'" & _
            " GROUP BY Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer, Inventory.Dept_ID, Departments.Description" & SQLSort
            
    SQL5 = "SELECT Invoice_Totals.Cashier_ID,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime, Inventory.Dept_ID, Departments.Description" & _
        " FROM Departments INNER JOIN (Inventory INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Departments.Dept_ID = Inventory.Dept_ID" & _
        " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
        " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'" & _
        " GROUP BY Invoice_Totals.Cashier_ID,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer, Inventory.Dept_ID, Departments.Description" & SQLSort
    Set crDetail80 = Nothing
    Set crDetail58 = Nothing
        cmd.ActiveConnection = cnData
        Select Case cbo80Filter.ListIndex
            Case 0
                cmd.CommandText = SQL
            Case 1
                cmd.CommandText = SQL2
            Case 2
                cmd.CommandText = SQL3
            Case 3:
                cmd.CommandText = SQL4
            Case 4:
                cmd.CommandText = SQL5
        End Select
        cmd.Execute
     If ReceiptType = 80 Then
        Set CRReport = crDetail80
    Else
        Set CRReport = crDetail58
    End If
    
    With CRReport
        .Database.AddADOCommand cnData, cmd
        .txtPluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .txtPluName.SetUnboundFieldSource "{ado.DiffItemName}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtCost.SetUnboundFieldSource "{ado.Price}"

'canh le

        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign

        If cbo80Filter.ListIndex = 1 Then
            .txtStoreID.SetUnboundFieldSource "{ado.Store_ID}"
            .Section1.Suppress = False
            .Section2.Suppress = False
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section5.Suppress = True
            .Section13.Suppress = True
        ElseIf cbo80Filter.ListIndex = 2 Then
            .txtStationID.SetUnboundFieldSource "{ado.Station_ID}"
            .Section3.Suppress = False
            .Section4.Suppress = False
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section5.Suppress = True
            .Section13.Suppress = True
        ElseIf cbo80Filter.ListIndex = 3 Then
            .txtGroup.SetUnboundFieldSource "{ado.Description}"
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section5.Suppress = False
            .Section13.Suppress = True
        ElseIf cbo80Filter.ListIndex = 4 Then
            .txtGroup.SetUnboundFieldSource "{ado.Description}"
            .txtCashier.SetUnboundFieldSource "{ado.Cashier_ID}"
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section5.Suppress = False
            .Section13.Suppress = False
        Else
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section5.Suppress = True
            .Section13.Suppress = True
        End If

        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtCost
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtStore
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtSumStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
    End With
    Set iReport = CRReport
With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = iReport
        .Show vbModal
    End With
    Unload frmShowCashierReport
    prbBanhang.Value = 0
Else

'
'theo nhom
    With frmShowCashierReport
        .filter_report = cbo80Sort.ListIndex
        .Let_Fromdate = Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00")
        .Let_Todate = Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00")
        .Show vbModal
    End With
    prbBanhang.Value = 0
End If
Exit Sub
Handle:
Exit Sub
MsgBox Err.Number & Err.Description & Me.name & " cmdBaocaochitiet_Click"
End Sub

Private Sub CmdDetailEmployee_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
    
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    'Khong danh cho karaoke
    SQL = "SELECT Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price," & _
    " Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime, Employee.EmpName" & _
    " FROM Employee INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) ON Employee.Cashier_ID = Invoice_Totals.OrderMan" & _
     " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
    " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'" & _
    " GROUP BY Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer, Employee.EmpName"

    Set crDetailOrder80 = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crDetailOrder80
        .Database.AddADOCommand cnData, cmd
        .txtOrder.SetUnboundFieldSource "{ado.EmpName}"
        .txtPluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .txtPluName.SetUnboundFieldSource "{ado.DiffItemName}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtCost.SetUnboundFieldSource "{ado.Price}"
        
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtCost
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        
    End With
    Set iReport = crDetailOrder80
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
Exit Sub
MsgBox Err.Number & Err.Description & Me.name & " cmdBaocaochitiet nhan vien_Click"
End Sub

Private Sub cmddiscountReason_Click()
    frmDiscount_reason_list.Show vbModal
End Sub

Private Sub cmdEditName_Click()
On Error GoTo Handle
        fraCashier.Visible = False
        fraSys.Visible = False
        fraList.Visible = False
        fraReport.Visible = False
        fraEditName.Visible = True
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdCashier_Click"

End Sub

Private Sub cmdGeneral80_Click()
On Error GoTo Handle
'Goi du lieu ban hang vao trong bao cao tong hop
    Call Get_Data_Report_General(gfCONVERT_DATE_TO_STRING(dtpFromDate), gfCONVERT_DATE_TO_STRING(dtpToDate))
    '''''''''''''''''''''''''''''''''
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim CRReport As CRAXDDRT.Report
    
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop
    
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = "SELECT * from RP_General"
    
    Set crGeneral80 = Nothing
    Set crGeneral58 = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    If ReceiptType = 80 Then
        Set CRReport = crGeneral80
    Else
        Set CRReport = crGeneral58
    End If
    
    With CRReport
        .Database.AddADOCommand cnData, cmd
        .txtDescription.SetUnboundFieldSource "{ado.Description}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtPrice.SetUnboundFieldSource "{ado.AVG_Price}"
        .txtAmt.SetUnboundFieldSource "{ado.Amount}"
        .txtCountTrans.SetUnboundFieldSource "{ado.CountTrans}"
        .txtCountDis.SetUnboundFieldSource "{ado.CountDist}"
        .txtDisAmt.SetUnboundFieldSource "{ado.AmountDist}"
        .txtCountDel.SetUnboundFieldSource "{ado.CountDeleteOrdered}"
        .txtDeleteAmt.SetUnboundFieldSource "{ado.AmountDeleteOrdered}"
        .txtCountDelNot.SetUnboundFieldSource "{ado.CountDelete}"
        .txtDelNotAmt.SetUnboundFieldSource "{ado.AmountCountDelete}"
        .txtCountReceipt.SetUnboundFieldSource "{ado.CountReceipt}"
        .txtReceiptAmt.SetUnboundFieldSource "{ado.AmountReceipt}"
        .txtCountExpense.SetUnboundFieldSource "{ado.CountPayouts}"
        .txtExpenseAmt.SetUnboundFieldSource "{ado.AmountPayouts}"
        ''''''
        .txtCountCA.SetUnboundFieldSource "{ado.CountCA}"
        .txtAmountCA.SetUnboundFieldSource "{ado.AmountCA}"
        
        'Giam % mon
        .txtCountLineDisc.SetUnboundFieldSource "{ado.KarDiscountCount}"
        .txtAmountLineDisc.SetUnboundFieldSource "{ado.KarDiscountAmount}"
        
        
        .txtCountOA.SetUnboundFieldSource "{ado.CountOA}"
        .txtAmountOA.SetUnboundFieldSource "{ado.AmountOA}"
        
        .txtCountCheck.SetUnboundFieldSource "{ado.CountCheck}"
        .txtAmountCheck.SetUnboundFieldSource "{ado.AmountCheck}"
        
        .txtCountCredit.SetUnboundFieldSource "{ado.CountCredit}"
        .txtAmountCredit.SetUnboundFieldSource "{ado.AmountCredit}"
        
        .txtCountROA.SetUnboundFieldSource "{ado.CountROA}"
        .txtAmountROA.SetUnboundFieldSource "{ado.AmountROA}"
        
        .txtCountGiftCard.SetUnboundFieldSource "{ado.CountGC}"
        .txtAmountGiftCard.SetUnboundFieldSource "{ado.AmountGC}"
        
        .txtCountOpen.SetUnboundFieldSource "{ado.CountOpen}"
        .txtAmountOpen.SetUnboundFieldSource "{ado.AmountOpen}"
        
        .txtCountSer.SetUnboundFieldSource "{ado.Service_Charge_Count}"
        
        .txtServAmt.SetUnboundFieldSource "{ado.Service_Charge_Amt}"
        
        .txtVATCount.SetUnboundFieldSource "{ado.VAT_Count}"
        .txtVATAmount.SetUnboundFieldSource "{ado.VAT_Amt}"
        
        .txtsokhach.SetUnboundFieldSource "{ado.Personal}"
        
       .adjAmt1.SetUnboundFieldSource "{ado.Adjustment1}"
        .CountAdj1.SetUnboundFieldSource "{ado.CountAdj1}"
        
        .AdjAmt2.SetUnboundFieldSource "{ado.Adjustment2}"
        .CountAdj2.SetUnboundFieldSource "{ado.CountAdj2}"
        
        .AdjAmt3.SetUnboundFieldSource "{ado.Adjustment3}"
        .CountAdj3.SetUnboundFieldSource "{ado.CountAdj3}"
        
        .AdjAmt4.SetUnboundFieldSource "{ado.Adjustment4}"
        .CountAdj4.SetUnboundFieldSource "{ado.CountAdj4}"
        
        .AdjAmt5.SetUnboundFieldSource "{ado.Adjustment5}"
        .CountAdj5.SetUnboundFieldSource "{ado.CountAdj5}"
        
        .AdjAmt6.SetUnboundFieldSource "{ado.Adjustment6}"
        .CountAdj6.SetUnboundFieldSource "{ado.CountAdj6}"
        '''
        ''''
        .txtCountReceiveMoney.SetUnboundFieldSource "{ado.CountReceive}"
        .txtAmountReceiveMoney.SetUnboundFieldSource "{ado.AmountReceive}"
        '''''
        .txtCountReserve.SetUnboundFieldSource "{ado.CountReserve}"
        .txtAmountReserve.SetUnboundFieldSource "{ado.AmountReserve}"


        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        'Gan caption
        .txtTitle.SetText DescArrReport(1)
        .lblFromdate.SetText DescArrReport(40)
        .lblToDate.SetText DescArrReport(41)
        .lblTongcong.SetText DescArrReport(62)
        .lblTransaction.SetText DescArrReport(3)
        .lblPhuthu.SetText DescArrReport(20)
        .lblDiscount.SetText DescArrReport(6)
        .lblService.SetText DescArrReport(52)
        .lblAdj1.SetText DescArrReport(7)
        .lblAdj2.SetText DescArrReport(8)
        .lblCash.SetText DescArrReport(12)
        .lblBalance.SetText DescArrReport(9)
        .lblCheck.SetText DescArrReport(10)
        .lblCredit.SetText DescArrReport(11)
        .lblNotPay.SetText DescArrReport(14)
        .lblCorrection.SetText DescArrReport(4)
        .lblVoid.SetText DescArrReport(5)
        .lblReceipt.SetText DescArrReport(54)
        .lblPay.SetText DescArrReport(55)
        .lblIndraw.SetText DescArrReport(21)
        .lblCashindrawer.SetText DescArrReport(13)
        .lblVAT.SetText DescArrReport(61)
        .lblGiftCard.SetText DescArrReport(63)
        'canh le
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign

        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtPrice
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtsumAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtDisAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtDeleteAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtDelNotAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtReceiptAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtExpenseAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .TxtTotal
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAmountCA
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAmountOA
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAmountCheck
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAmountCredit
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        
        With .txtAmountOpen
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtCashInDrawer
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAmount
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtServAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .AdjAmt2
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .adjAmt1
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .AdjAmt3
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        
        With .AdjAmt4
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .AdjAmt5
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .AdjAmt6
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    
    Set iReport = CRReport
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
Exit Sub
MsgBox Err.Number & Err.Description & Me.name & " cmdBaocaotonghop_Click"
End Sub

Private Sub cmdInvoiceHoldList_Click()
    frmLocation_Cashier.Show vbModal
End Sub

Private Sub cmdItemCard_Click()
    frmSochitiet.Show vbModal
End Sub

Private Sub cmdItemDiscount_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL, SQL2, SQL3, SQL4, SQL5 As String
    Dim RptID As Integer

    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    'Khong danh cho karaoke
    SQL = "SELECT right(Invoice_Totals.Invoice_Number,4) as Invoice_Num,Invoice_Totals.Orig_OnHoldID,Invoice_Itemized.ItemNum,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc, Invoice_Itemized.DiffItemName," & _
          " Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price," & _
          " Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime" & _
          " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number " & _
          " WHERE Invoice_Itemized.LineDisc>0 and (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
          " GROUP BY Invoice_Totals.Invoice_Number,Invoice_Totals.Orig_OnHoldID,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc"

    SQL2 = "SELECT right(Invoice_Totals.Invoice_Number,4) as Invoice_Num,Invoice_Totals.Orig_OnHoldID,Invoice_Itemized.ItemNum,Invoice_Itemized.LineDisc, Invoice_Totals.Store_ID," & _
           " Invoice_Itemized.DiffItemName,Invoice_Itemized.Line_Disc_Desc, Sum(Invoice_Itemized.Quantity) AS Qty," & _
           " Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime" & _
           " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number" & _
            " WHERE Invoice_Itemized.LineDisc>0 and(((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
          " GROUP BY right(Invoice_Totals.Invoice_Number,4),Invoice_Totals.Orig_OnHoldID,Invoice_Itemized.ItemNum, Invoice_Totals.Store_ID, Invoice_Itemized.DiffItemName," & _
          " Invoice_Itemized.PricePer,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc"

    SQL3 = "right(Invoice_Totals.Invoice_Number,4) as invoice_Num,Invoice_Totals.Orig_OnHoldID,SELECT Invoice_Itemized.ItemNum,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc, Invoice_Totals.Station_ID, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime" & _
            " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number" & _
        " WHERE Invoice_Itemized.LineDisc>0 and(((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
        " GROUP BY right(Invoice_Totals.Invoice_Number,4),Invoice_Totals.Orig_OnHoldID,Invoice_Itemized.ItemNum,Invoice_Itemized.LineDisc, Invoice_Totals.Station_ID, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer,Invoice_Itemized.Line_Disc_Desc"

    SQL4 = "SELECT right(Invoice_Totals.Invoice_Number,4) as invoice_Num,Invoice_Totals.Orig_OnHoldID,Invoice_Itemized.ItemNum,Invoice_Itemized.LineDiscInvoice_Itemized.Line_Disc_Desc, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime, Inventory.Dept_ID, Departments.Description" & _
            " FROM Departments INNER JOIN (Inventory INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Departments.Dept_ID = Inventory.Dept_ID" & _
            " WHERE Invoice_Itemized.LineDisc>0 and(((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
            " GROUP BY right(Invoice_Totals.Invoice_Number,4),Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc,Invoice_Itemized.Line_Disc_Desc, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer, Inventory.Dept_ID, Departments.Description"
    SQL5 = "SELECT right(Invoice_Totals.Invoice_Number,4) as invoice_Num,Invoice_Totals.Orig_OnHoldID,Invoice_Totals.Cashier_ID,Invoice_Itemized.ItemNum,Invoice_Itemized.LineDisc, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime, Inventory.Dept_ID, Departments.Description" & _
            " FROM Departments INNER JOIN (Inventory INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Departments.Dept_ID = Inventory.Dept_ID" & _
            " WHERE Invoice_Itemized.LineDisc>0 and (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
            " GROUP BY right(Invoice_Totals.Invoice_Number,4),Invoice_Totals.Orig_OnHoldID,Invoice_Totals.Cashier_ID,Invoice_Itemized.ItemNum,Invoice_Itemized.LineDisc,Invoice_Itemized.Line_Disc_Desc, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer, Inventory.Dept_ID, Departments.Description"
    Set crItemDiscountDetails = Nothing
        cmd.ActiveConnection = cnData
        Select Case cboSaleFilter.ListIndex
            Case 0
                cmd.CommandText = SQL
            Case 1
                cmd.CommandText = SQL2
            Case 2
                cmd.CommandText = SQL3
            Case 3:
                cmd.CommandText = SQL4
            Case 4:
                cmd.CommandText = SQL5
        End Select
        cmd.Execute
    With crItemDiscountDetails
        .Database.AddADOCommand cnData, cmd
        .txtInvoice.SetUnboundFieldSource "{ado.Invoice_Num}"
        .txtTableNo.SetUnboundFieldSource "{ado.Orig_OnHoldID}"
        .txtPluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .txtPluName.SetUnboundFieldSource "{ado.DiffItemName}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtCost.SetUnboundFieldSource "{ado.Price}"
        .txtDisc.SetUnboundFieldSource "{ado.LineDisc}"
        .txtLineDiscDesc.SetUnboundFieldSource "{ado.Line_Disc_Desc}"
        If cboSaleFilter.ListIndex = 1 Then
            .txtStoreID.SetUnboundFieldSource "{ado.Store_ID}"
            .Section1.Suppress = False
            .Section2.Suppress = False
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section5.Suppress = True
            .Section12.Suppress = True
        ElseIf cboSaleFilter.ListIndex = 2 Then
            .txtStationID.SetUnboundFieldSource "{ado.Station_ID}"
            .Section3.Suppress = False
            .Section4.Suppress = False
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section5.Suppress = True
            .Section12.Suppress = True
        ElseIf cboSaleFilter.ListIndex = 3 Then
            .txtGroup.SetUnboundFieldSource "{ado.Description}"
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section5.Suppress = False
            .Section12.Suppress = True
        ElseIf cboSaleFilter.ListIndex = 4 Then
            .txtGroup.SetUnboundFieldSource "{ado.Description}"
            .txtCashier.SetUnboundFieldSource "{ado.Cashier_ID}"
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section5.Suppress = False
            .Section12.Suppress = False
        Else
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section5.Suppress = True
            .Section12.Suppress = True
        End If

        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtCost
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtStore
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtSumStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    Set iReport = crItemDiscountDetails
    With frmShowReport
        .Report_Number = 2
        .Get_fDate = Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00")
        .Get_tDate = Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00")
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmd_Baocao giam chi tiet mon"
End Sub

Private Sub cmdItemsforCust_Click()
    MsgBox "Please wait to Complete, thanks!"
End Sub

Private Sub cmdKit_General_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL, sql1 As String
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Select Case cboKit_Printer.ListIndex
        Case 0
            SQL = "SELECT Kitchen_Order_Master.Invoice_Number, Kitchen_Order_Master.Station_ID," & _
            " Kitchen_Order_Master.Store_ID, Kitchen_Order_Master.Cashier_ID, " & _
            " Kitchen_Order_Master.Table_ID,  Kitchen_Order_Items.Send_KP_Date," & _
            " Kitchen_Order_Items.Send_KP_Time, Kitchen_Order_Items.ItemName, Kitchen_Order_Items.Quantity," & _
            " Kitchen_Order_Items.Price,Kitchen_Order_Items.Price*Kitchen_Order_Items.Quantity as Amt, Kitchen_Order_Items.LineNum, Kitchen_Order_Items.Kit_Desc " & _
            " FROM Kitchen_Order_Master INNER JOIN Kitchen_Order_Items ON Kitchen_Order_Master.Invoice_Number = Kitchen_Order_Items.Invoice_Number" & _
            " where Send_KP_Date>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "' and Send_KP_Date<='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'"
        Case Else
        sql1 = "SELECT Kitchen_Order_Master.Invoice_Number, Kitchen_Order_Master.Station_ID," & _
            " Kitchen_Order_Master.Store_ID, Kitchen_Order_Master.Cashier_ID, " & _
            " Kitchen_Order_Master.Table_ID,  Kitchen_Order_Items.Send_KP_Date," & _
            " Kitchen_Order_Items.Send_KP_Time, Kitchen_Order_Items.ItemName, Kitchen_Order_Items.Quantity," & _
            " Kitchen_Order_Items.Price,Kitchen_Order_Items.Price*Kitchen_Order_Items.Quantity as Amt, Kitchen_Order_Items.LineNum, Kitchen_Order_Items.Kit_Desc " & _
            " FROM Kitchen_Order_Master INNER JOIN Kitchen_Order_Items ON Kitchen_Order_Master.Invoice_Number = Kitchen_Order_Items.Invoice_Number" & _
            " where Kitchen_Order_Items.Printer_ID='" & Trim(Printer_ID) & "' and Send_KP_Date>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "' and Send_KP_Date<='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'"
       
    End Select
    
    Set crSend_KP_Detail = Nothing
        cmd.ActiveConnection = cnData
        Select Case cboKit_Printer.ListIndex
        Case 0
            cmd.CommandText = SQL
        Case Else
            cmd.CommandText = sql1
        End Select
        cmd.Execute
    With crSend_KP_Detail
        .Database.AddADOCommand cnData, cmd
        
        .txtInvoiceNum.SetUnboundFieldSource "{ado.Invoice_Number}"
        .txtTableID.SetUnboundFieldSource "{ado.Table_ID}"
        .txtCashierID.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtPluName.SetUnboundFieldSource "{ado.ItemName}"
        .txtCost.SetUnboundFieldSource "{ado.Price}"
        .txtQty.SetUnboundFieldSource "{ado.Quantity}"
        .txtAmt.SetUnboundFieldSource "{ado.Amt}"
        .txtTime.SetUnboundFieldSource "{ado.Send_KP_Time}"
'        .txtStoreID.SetUnboundFieldSource "{ado.store_ID}"
        .txtKitDesc.SetUnboundFieldSource "{ado.Kit_Desc}"
        
        If cboKit_Filter.ListIndex = 1 Then
            .txtStoreID.SetUnboundFieldSource "{ado.Store_ID}"
            .Section1.Suppress = False
            .Section2.Suppress = False
            .Section3.Suppress = True
            .Section4.Suppress = True
        ElseIf cboKit_Filter.ListIndex = 2 Then
            .txtStationID.SetUnboundFieldSource "{ado.Station_ID}"
            .Section3.Suppress = False
            .Section4.Suppress = False
            .Section1.Suppress = True
            .Section2.Suppress = True
        Else
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section3.Suppress = True
            .Section4.Suppress = True
        End If
        
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtCost
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtStore
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtSumStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
    End With
    Set iReport = crSend_KP_Detail
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdBaocaochitiet_Click"
End Sub

Private Sub cmdKitchen_List_Click()
On Error GoTo Handle
        fraBaocaobanhang.Visible = False
        fraKitchen.Visible = True
        fra80.Visible = False

        'Call addPrinter
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdSaleReport_Click"

End Sub



Private Sub cmdLevelMaterial_Click()

  On Error GoTo errHdl
    Dim cmdMaterial As New ADODB.Command
    Dim strSql As String
    strSql = "SELECT Inventory.ItemNum, Inventory.ItemName,Inventory.Std_Price1, SetMLink.SMPLUCode," & _
            " SetMLink.StockRate/1000 as Rate, SetMPLU.PLUName," & _
            " SetMPLU.Cost, SetMPLU.Unit" & _
            " FROM SetMPLU INNER JOIN (Inventory INNER JOIN SetMLink ON " & _
            " Inventory.ItemNum = SetMLink.PLUCode) ON SetMPLU.PLUCode = SetMLink.SMPLUCode"
    
    
    With cmdMaterial
        .ActiveConnection = cnData
        .CommandType = adCmdText
        .CommandText = strSql
        .Execute
    End With
    Dim sReport As New crMaterial
    With sReport
        .Database.AddADOCommand cnData, cmdMaterial
        .PluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .PluName.SetUnboundFieldSource "{ado.ItemName}"
        .Price.SetUnboundFieldSource "{ado.Std_Price1}"
        .SMPluCode.SetUnboundFieldSource "{ado.SMPluCode}"
        .SMPluName.SetUnboundFieldSource "{ado.PLUName}"
        .SMUnit.SetUnboundFieldSource "{ado.Unit}"
        .SMStockRate.SetUnboundFieldSource "{ado.Rate}"
        .Cost.SetUnboundFieldSource "{ado.Cost}"
'        .ReportTitle = DescArr(16)
'        .lblItemcode.SetText DescArr(17)
'        .lblItemName.SetText DescArr(18)
'        .lblUnit.SetText DescArr(19)
'        .lblStockRate.SetText DescArr(20)
'        .lblPrice.SetText DescArr(21)
'        .lblTotals.SetText DescArr(22)
        
    End With
    With frmShowReport
        .Report = sReport
        .Show vbModal, Me
    End With
    Set sReport = Nothing
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " - " & Err.Description & vbCrLf _
        & Me.name & " - Material_Report "

End Sub

Private Sub cmdLayout_Click()
    frmLayout.Show vbModal
End Sub

'Private Sub cmdLock_Click()
'On Error GoTo Handle
'    frmLockBook.Show vbModal
'Exit Sub
'Handle:
'    MsgBox Err.Number & Err.Description & Me.name & " cmdLock_Click"
'End Sub

Private Sub cmdMixmatchReport_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
    
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
'
SQL = "SELECT right( Invoice_Totals.Invoice_Number,4) as invoice_num, Invoice_Totals.Orig_OnHoldID,Invoice_Totals.CustNum, Invoice_Totals.Total_Price, Invoice_Totals.Tax_Rate_ID, (Invoice_Totals.Discount*Invoice_Totals.Total_Tax1)/100 as Amt_Disc,Invoice_Totals.Adjustment1, Invoice_Totals.Adjustment2,Invoice_Totals.Cashier_ID,Invoice_Totals.Pro_Desc" & _
      " from Invoice_Totals " & _
      " WHERE Total_Price<>Total_Tax1 and(((Invoice_Totals.Status)='C')) and Left([DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And Left([DateTime], 8) <= '" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'" & _
      " GROUP BY right(Invoice_Totals.Invoice_Number,4),Invoice_Totals.Orig_OnHoldID,Invoice_Totals.CustNum, Invoice_Totals.Total_Price, Invoice_Totals.Tax_Rate_ID, (Invoice_Totals.Discount*Invoice_Totals.Total_Tax1)/100,Adjustment1,Adjustment2, Invoice_Totals.Cashier_ID,Invoice_Totals.Pro_Desc;"
    
    Set crMixmatch = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crMixmatch
        .Database.AddADOCommand cnData, cmd
        .txtInvoiceNum.SetUnboundFieldSource "{ado.Invoice_Num}"
        .txtTotals.SetUnboundFieldSource "{ado.Total_Price}"
        .txtDiscountAmt.SetUnboundFieldSource "{ado.Amt_Disc}"
        .txtAdj1Amt.SetUnboundFieldSource "{ado.Adjustment1}"
        .txtAdj2Amt.SetUnboundFieldSource "{ado.Adjustment2}"
        .txtCashierID.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtTableID.SetUnboundFieldSource "{ado.Orig_OnHoldID}"
        .txtPromotion.SetUnboundFieldSource "{ado.Pro_Desc}"
        .txtCust.SetUnboundFieldSource "{ado.CustNum}"
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
    End With
    Set iReport = crMixmatch
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmMixmatch_Report"
End Sub


Private Sub cmdOrderDetail80_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL, sql1 As String
    Dim CRReport As CRAXDDRT.Report
    
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    Select Case cboKit_Printer.ListIndex
        Case 0
            SQL = "SELECT Kitchen_Order_Master.Invoice_Number, Kitchen_Order_Master.Station_ID," & _
            " Kitchen_Order_Master.Store_ID, Kitchen_Order_Master.Cashier_ID, " & _
            " Kitchen_Order_Master.Table_ID,  Kitchen_Order_Items.Send_KP_Date," & _
            " Kitchen_Order_Items.Send_KP_Time, Kitchen_Order_Items.ItemName, Kitchen_Order_Items.Quantity," & _
            " Kitchen_Order_Items.Price,Kitchen_Order_Items.Price*Kitchen_Order_Items.Quantity as Amt, Kitchen_Order_Items.LineNum, Kitchen_Order_Items.Kit_Desc " & _
            " FROM Kitchen_Order_Master INNER JOIN Kitchen_Order_Items ON Kitchen_Order_Master.Invoice_Number = Kitchen_Order_Items.Invoice_Number" & _
            " where Send_KP_Date>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "' and Send_KP_Date<='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'"
        Case Else
        sql1 = "SELECT Kitchen_Order_Master.Invoice_Number, Kitchen_Order_Master.Station_ID," & _
            " Kitchen_Order_Master.Store_ID, Kitchen_Order_Master.Cashier_ID, " & _
            " Kitchen_Order_Master.Table_ID,  Kitchen_Order_Items.Send_KP_Date," & _
            " Kitchen_Order_Items.Send_KP_Time, Kitchen_Order_Items.ItemName, Kitchen_Order_Items.Quantity," & _
            " Kitchen_Order_Items.Price,Kitchen_Order_Items.Price*Kitchen_Order_Items.Quantity as Amt, Kitchen_Order_Items.LineNum, Kitchen_Order_Items.Kit_Desc " & _
            " FROM Kitchen_Order_Master INNER JOIN Kitchen_Order_Items ON Kitchen_Order_Master.Invoice_Number = Kitchen_Order_Items.Invoice_Number" & _
            " where Kitchen_Order_Items.Printer_ID='" & Trim(Printer_ID) & "' and Send_KP_Date>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "' and Send_KP_Date<='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'"
       
    End Select
    
    Set crSend_KP_Detail80 = Nothing
    Set crSend_KP_Detail58 = Nothing
        cmd.ActiveConnection = cnData
        Select Case cboKit_Printer.ListIndex
        Case 0
            cmd.CommandText = SQL
        Case Else
            cmd.CommandText = sql1
        End Select
        cmd.Execute
    If ReceiptType = 80 Then
        Set CRReport = crSend_KP_Detail80
    Else
        Set CRReport = crSend_KP_Detail58
    End If
    
    With CRReport
        .Database.AddADOCommand cnData, cmd
        
        .txtInvoiceNum.SetUnboundFieldSource "{ado.Invoice_Number}"
        .txtTableID.SetUnboundFieldSource "{ado.Table_ID}"
        .txtCashierID.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtPluName.SetUnboundFieldSource "{ado.ItemName}"
        .txtCost.SetUnboundFieldSource "{ado.Price}"
        .txtQty.SetUnboundFieldSource "{ado.Quantity}"
        .txtAmt.SetUnboundFieldSource "{ado.Amt}"
        .txtTime.SetUnboundFieldSource "{ado.Send_KP_Time}"
'        .txtStoreID.SetUnboundFieldSource "{ado.store_ID}"
        .txtKitDesc.SetUnboundFieldSource "{ado.Kit_Desc}"
        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign
        If cboKit_Filter.ListIndex = 1 Then
            .txtStoreID.SetUnboundFieldSource "{ado.Store_ID}"
            .Section1.Suppress = False
            .Section2.Suppress = False
            .Section3.Suppress = True
            .Section4.Suppress = True
        ElseIf cboKit_Filter.ListIndex = 2 Then
            .txtStationID.SetUnboundFieldSource "{ado.Station_ID}"
            .Section3.Suppress = False
            .Section4.Suppress = False
            .Section1.Suppress = True
            .Section2.Suppress = True
        Else
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section3.Suppress = True
            .Section4.Suppress = True
        End If
        
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtCost
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtStore
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtSumStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
    End With
    Set iReport = CRReport
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = iReport
        .Show vbModal
    End With
Exit Sub
Handle:
Exit Sub
MsgBox Err.Number & Err.Description & Me.name & " cmdBaocaochitiet_Click"
End Sub

Private Sub cmdProfit_Click()
    frmLost_Profit_State.Show vbModal
    
End Sub

Private Sub cmdReport80_Click()
    fraBaocaobanhang.Visible = False
    fraKitchen.Visible = False
    fra80.Visible = True

End Sub

Private Sub cmdStockReport_Click()
On Error GoTo Handle
        fraBaocaobanhang.Visible = False
        fraKitchen.Visible = False
        fra80.Visible = False
        
        If MsgBox("B¹n ®· tÝnh tån kho ch­a?", vbYesNo) = vbNo Then
            frmCal_TonTemp.Show vbModal
        End If
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdSaleReport_Click"
End Sub


Private Sub cmdStockList_Click()
    frmStockType.Show vbModal
End Sub

Private Sub cmdSystemFlag_Click()
    With frmPassword
        .FormActionKey = "SystemFlag"
        .Show vbModal
    End With
End Sub

Private Sub cmdAdjustment_Click()
    frmAdjustment.Show vbModal
End Sub

Private Sub cmdAdministrative_Click()
On Error GoTo Handle
        fraCashier.Visible = False
        fraSys.Visible = False
        fraList.Visible = True
        fraReport.Visible = False
        
        fraEditName.Visible = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdCashier_Click"
End Sub

Private Sub cmdBaocaochitiet_Click()
On Error GoTo Handle
'    With frmShowCashierReport
'        .Let_FromDate = Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00")
'        .Let_ToDate = Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00")
'        .Show vbModal
'
'    End With
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL, SQL2, SQL3, SQL4, SQL5, SQL6, SQLSort As String
    Dim RptID As Integer
    Select Case cboSalesort.ListIndex
            Case 0: SQLSort = " Order by Invoice_Itemized.ItemNum  ASC"
            Case 1: SQLSort = " Order by Invoice_Itemized.DiffItemName  ASC"
            Case 2: SQLSort = " Order by sum(Invoice_Itemized.Quantity)  DESC"
            Case Else: SQLSort = " Order by Invoice_Itemized.ItemNum  ASC"
    End Select
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    'Khong danh cho karaoke
   SQL = "SELECT Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime" & _
          " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number " & _
          " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
          " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'" & _
          " GROUP BY Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer" & SQLSort

    SQL2 = "SELECT Invoice_Itemized.ItemNum, Invoice_Totals.Store_ID, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime" & _
           " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number" & _
           " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
           " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'" & _
           " GROUP BY Invoice_Itemized.ItemNum, Invoice_Totals.Store_ID, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer" & SQLSort

    SQL3 = "SELECT Invoice_Itemized.ItemNum, Invoice_Totals.Station_ID, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime" & _
            " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number" & _
        " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
         " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'" & _
        " GROUP BY Invoice_Itemized.ItemNum, Invoice_Totals.Station_ID, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer" & SQLSort

    SQL4 = "SELECT Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime, Inventory.Dept_ID, Departments.Description" & _
            " FROM Departments INNER JOIN (Inventory INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Departments.Dept_ID = Inventory.Dept_ID" & _
            " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
             " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'" & _
            " GROUP BY Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer, Inventory.Dept_ID, Departments.Description" & SQLSort
          
    SQL5 = "SELECT Invoice_Totals.Cashier_ID,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime, Inventory.Dept_ID, Departments.Description" & _
            " FROM Departments INNER JOIN (Inventory INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Departments.Dept_ID = Inventory.Dept_ID" & _
            " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
             " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'" & _
            " GROUP BY Invoice_Totals.Cashier_ID,Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer, Inventory.Dept_ID, Departments.Description" & SQLSort
            
    SQL6 = "SELECT Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Sum(Invoice_Itemized.Quantity) AS Qty, Avg(Invoice_Itemized.PricePer) AS Price, Count(Invoice_Totals.Invoice_Number) AS Count_Invoice_Number, Max(Invoice_Totals.DateTime) AS MaxOfDateTime, MainGroup.GroupName" & _
                " FROM MainGroup INNER JOIN (Invoice_Totals INNER JOIN (Departments INNER JOIN (Inventory INNER JOIN Invoice_Itemized ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Departments.Dept_ID = Inventory.Dept_ID) ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number) ON MainGroup.GroupNo = Departments.MainGroup" & _
            " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
             " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'" & _
            " GROUP BY Invoice_Itemized.ItemNum, Invoice_Itemized.DiffItemName, Invoice_Itemized.PricePer, MainGroup.GroupName" & SQLSort
    Set crSaleReport = Nothing
        cmd.ActiveConnection = cnData
        Select Case cboSaleFilter.ListIndex
            Case 0
                cmd.CommandText = SQL
            Case 1
                cmd.CommandText = SQL2
            Case 2
                cmd.CommandText = SQL3
            Case 3:
                cmd.CommandText = SQL4
            Case 4:
                cmd.CommandText = SQL5
            Case 5:
                cmd.CommandText = SQL6
        End Select
        cmd.Execute
    With crSaleReport
        .Database.AddADOCommand cnData, cmd
        .txtPluCode.SetUnboundFieldSource "{ado.ItemNum}"
        .txtPluName.SetUnboundFieldSource "{ado.DiffItemName}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtCost.SetUnboundFieldSource "{ado.Price}"

        ''''''''''''''''''''''''''''''''''''''''
'        .lblStt.SetText DescArrReport(43)
'        .txtTitle.SetText DescArrReport(42)
'        .lblItemcode.SetText DescArrReport(44)
'        .lblItemName.SetText DescArrReport(45)
'        .lblUnit.SetText DescArrReport(48)
'        .lblQty.SetText DescArrReport(46)
'        .lblPrice.SetText DescArrReport(47)
'        .lblAmount.SetText DescArrReport(49)
'        .lblInword.SetText DescArrReport(56)
'        .lblCashier.SetText DescArrReport(57)
'        .lblChief.SetText DescArrReport(58)
'        .lblDirector.SetText DescArrReport(59)
'        .lblSign1.SetText DescArrReport(60)
'        .lblSign2.SetText DescArrReport(60)
'        .lblsign3.SetText DescArrReport(60)
'        .lblFromdate.SetText DescArrReport(40)
'        .lblTodate.SetText DescArrReport(41)

        If cboSaleFilter.ListIndex = 1 Then
            .txtStoreID.SetUnboundFieldSource "{ado.Store_ID}"
            .Section1.Suppress = False
            .Section2.Suppress = False
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section5.Suppress = True
            .Section12.Suppress = True
            .Section14.Suppress = True
        ElseIf cboSaleFilter.ListIndex = 2 Then
            .txtStationID.SetUnboundFieldSource "{ado.Station_ID}"
            .Section3.Suppress = False
            .Section4.Suppress = False
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section5.Suppress = True
            .Section12.Suppress = True
            .Section14.Suppress = True
        ElseIf cboSaleFilter.ListIndex = 3 Then
            .txtGroup.SetUnboundFieldSource "{ado.Description}"
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section5.Suppress = False
            .Section12.Suppress = True
            .Section14.Suppress = True
        ElseIf cboSaleFilter.ListIndex = 4 Then
            .txtGroup.SetUnboundFieldSource "{ado.Description}"
            .txtCashier.SetUnboundFieldSource "{ado.Cashier_ID}"
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section5.Suppress = False
            .Section12.Suppress = False
            .Section14.Suppress = True
        ElseIf cboSaleFilter.ListIndex = 5 Then
            .txtMaingroup.SetUnboundFieldSource "{ado.GroupName}"
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section5.Suppress = True
            .Section12.Suppress = True
            .Section14.Suppress = False
        Else
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section5.Suppress = True
            .Section12.Suppress = True
            .Section14.Suppress = True
        End If

        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtCost
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtStore
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtSumStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    Set iReport = crSaleReport
    With frmShowReport
        .Report_Number = 2
        .Get_fDate = Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00")
        .Get_tDate = Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00")
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdBaocaochitiet_Click"
End Sub

Private Sub cmdBaocaotonghop_Click()
On Error GoTo Handle
'Goi du lieu ban hang vao trong bao cao tong hop
    Call Get_Data_Report_General(gfCONVERT_DATE_TO_STRING(dtpFromDate), gfCONVERT_DATE_TO_STRING(dtpToDate))
    '''''''''''''''''''''''''''''''''
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
    
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop
    
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = "SELECT * from RP_General"
    
    Set cGeneralReport = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With cGeneralReport
        .Database.AddADOCommand cnData, cmd
        .txtDescription.SetUnboundFieldSource "{ado.Description}"
        .txtQty.SetUnboundFieldSource "{ado.Qty}"
        .txtPrice.SetUnboundFieldSource "{ado.AVG_Price}"
        .txtAmt.SetUnboundFieldSource "{ado.Amount}"
        .txtCountTrans.SetUnboundFieldSource "{ado.CountTrans}"
        .txtCountDis.SetUnboundFieldSource "{ado.CountDist}"
        .txtDisAmt.SetUnboundFieldSource "{ado.AmountDist}"
        .txtCountDel.SetUnboundFieldSource "{ado.CountDeleteOrdered}"
        .txtDeleteAmt.SetUnboundFieldSource "{ado.AmountDeleteOrdered}"
        .txtCountDelNot.SetUnboundFieldSource "{ado.CountDelete}"
        .txtDelNotAmt.SetUnboundFieldSource "{ado.AmountCountDelete}"
        .txtCountReceipt.SetUnboundFieldSource "{ado.CountReceipt}"
        .txtReceiptAmt.SetUnboundFieldSource "{ado.AmountReceipt}"
        .txtCountExpense.SetUnboundFieldSource "{ado.CountPayouts}"
        .txtExpenseAmt.SetUnboundFieldSource "{ado.AmountPayouts}"
        ''''
        'Giam % mon
        .txtCountLineDisc.SetUnboundFieldSource "{ado.KarDiscountCount}"
        .txtAmountLineDisc.SetUnboundFieldSource "{ado.KarDiscountAmount}"
        ''''
        .txtCountCA.SetUnboundFieldSource "{ado.CountCA}"
        .txtAmountCA.SetUnboundFieldSource "{ado.AmountCA}"
        
        .txtCountOA.SetUnboundFieldSource "{ado.CountOA}"
        .txtAmountOA.SetUnboundFieldSource "{ado.AmountOA}"
        
        .txtCountCredit.SetUnboundFieldSource "{ado.CountCredit}"
        .txtAmountCredit.SetUnboundFieldSource "{ado.AmountCredit}"
        
        .txtCountCheck.SetUnboundFieldSource "{ado.CountCheck}"
        .txtAmountCheck.SetUnboundFieldSource "{ado.AmountCheck}"
        
        .txtCountGiftCard.SetUnboundFieldSource "{ado.CountGC}"
        .txtAmountGiftCard.SetUnboundFieldSource "{ado.AmountGC}"
        
        .txtCountROA.SetUnboundFieldSource "{ado.CountROA}"
        .txtAmountROA.SetUnboundFieldSource "{ado.AmountROA}"
        
        .txtCountOpen.SetUnboundFieldSource "{ado.CountOpen}"
        .txtAmountOpen.SetUnboundFieldSource "{ado.AmountOpen}"
        
        .txtCountSer.SetUnboundFieldSource "{ado.Service_Charge_Count}"
        .txtServAmt.SetUnboundFieldSource "{ado.Service_Charge_Amt}"
        
        .txtVATCount.SetUnboundFieldSource "{ado.VAT_Count}"
        .txtVATAmount.SetUnboundFieldSource "{ado.VAT_Amt}"
        
        .txtsokhach.SetUnboundFieldSource "{ado.Personal}"
        
        .adjAmt1.SetUnboundFieldSource "{ado.Adjustment1}"
        .CountAdj1.SetUnboundFieldSource "{ado.CountAdj1}"
        
        .AdjAmt2.SetUnboundFieldSource "{ado.Adjustment2}"
        .CountAdj2.SetUnboundFieldSource "{ado.CountAdj2}"
        
        .AdjAmt3.SetUnboundFieldSource "{ado.Adjustment3}"
        .CountAdj3.SetUnboundFieldSource "{ado.CountAdj3}"
        
        .AdjAmt4.SetUnboundFieldSource "{ado.Adjustment4}"
        .CountAdj4.SetUnboundFieldSource "{ado.CountAdj4}"
        
        .AdjAmt5.SetUnboundFieldSource "{ado.Adjustment5}"
        .CountAdj5.SetUnboundFieldSource "{ado.CountAdj5}"
        
        .AdjAmt6.SetUnboundFieldSource "{ado.Adjustment6}"
        .CountAdj6.SetUnboundFieldSource "{ado.CountAdj6}"
        ''''
        .txtCountReceiveMoney.SetUnboundFieldSource "{ado.CountReceive}"
        .txtAmountReceiveMoney.SetUnboundFieldSource "{ado.AmountReceive}"
        
        ''''''''''''''''
        .txtCountReserve.SetUnboundFieldSource "{ado.CountReserve}"
        .txtAmountReserve.SetUnboundFieldSource "{ado.AmountReserve}"
        
        '''''
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        ''Gan cac nhan report
        .lblStt.SetText DescArrReport(43)
        .txtTitle.SetText DescArrReport(1)
        .lblGroupName.SetText DescArrReport(2)
        .lblFromdate.SetText DescArrReport(40)
        .lblToDate.SetText DescArrReport(41)
        .lblQty.SetText DescArrReport(46)
        .lblPrice.SetText DescArrReport(47)
        .lblAmt.SetText DescArrReport(49)
        .lblTotalGroup.SetText DescArrReport(50)
        .lblTransaction.SetText DescArrReport(3)
        .lblServiceCharge.SetText DescArrReport(20)
        .lblDiscount.SetText DescArrReport(6)
        .lblService.SetText DescArrReport(52)
        .lblAdj1.SetText DescArrReport(7)
        .lblAdj2.SetText DescArrReport(8)
        .lblCash.SetText DescArrReport(12)
        .lblBalance.SetText DescArrReport(9)
        .lblCheck.SetText DescArrReport(10)
        .lblCredit.SetText DescArrReport(11)
        .lblNotPay.SetText DescArrReport(14)
        .lblCorrection.SetText DescArrReport(4)
        .lblVoid.SetText DescArrReport(5)
        .lblReceipt.SetText DescArrReport(54)
        .lblPay.SetText DescArrReport(55)
        .lblIndraw.SetText DescArrReport(21)
        .lblCashindrawer.SetText DescArrReport(13)
        .lblInword.SetText DescArrReport(56)
        .lblCashier.SetText DescArrReport(57)
        .lblChief.SetText DescArrReport(58)
        .lblDirector.SetText DescArrReport(59)
        .lblSign1.SetText DescArrReport(60)
        .lblSign2.SetText DescArrReport(60)
        .lblSign3.SetText DescArrReport(60)
        .lblVAT.SetText DescArrReport(61)
        'ket thuc

        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtPrice
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtsumAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtDisAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtDeleteAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtDelNotAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtReceiptAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtExpenseAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .TxtTotal
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAmountCA
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAmountOA
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAmountCheck
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAmountCredit
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAmountOpen
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtCashInDrawer
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtAmount
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtServAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .AdjAmt2
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .adjAmt1
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    
    Set iReport = cGeneralReport
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdBaocaotonghop_Click"
End Sub

Private Sub cmdCal_Instruction_Click()
On Error GoTo Handle
        fraCashier.Visible = False
        fraSys.Visible = False
        fraList.Visible = False
        fraReport.Visible = False
        fraEditName.Visible = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdCashier_Click"
End Sub

Private Sub cmdCashier_Click()
    On Error GoTo Handle
        fraCashier.Visible = True
        fraSys.Visible = False
        fraList.Visible = False
        fraReport.Visible = False
        fraEditName.Visible = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdCashier_Click"
End Sub

Private Sub cmdCashierReport_Click()
       frmLost_Profit_State.Show vbModal
End Sub

Private Sub cmdCashierRepot_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
    
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    If UserLevel <> 1 Then
        SQL = "SELECT Count(Invoice_Totals.Invoice_Number) AS [Transaction]," & _
        " Count(Invoice_Totals.Discount) AS CountDis, " & _
        " Sum(Invoice_Totals.Total_Price) AS SumTP,Sum(Invoice_Totals.Adjustment1) AS Adj1,Sum(Invoice_Totals.Adjustment2) AS Adj2, " & _
        " Sum(Invoice_Totals.Grand_Total) AS sumGT," & _
        " Sum(Invoice_Totals.Adjustment3) AS Adj3,Sum(Invoice_Totals.Adjustment4) AS Adj4,Sum(Invoice_Totals.Adjustment5) AS Adj5,Sum(Invoice_Totals.Adjustment6) AS Adj6," & _
        " Invoice_Totals.Cashier_ID," & _
        " sum(Invoice_Totals.Discount*Invoice_Totals.Total_Price/100) AS AmtDis," & _
        " sum(Invoice_Totals.Service_Charge*Invoice_Totals.Total_Price/100) AS AmtSer," & _
        " sum(Invoice_Totals.VATFee*Invoice_Totals.Total_Tax1/100) AS AmtVAT," & _
        " sum(Invoice_Totals.AddMoney) AS AmtReceive,Left([DateTime],8)as Datesale " & _
        " From Invoice_Totals" & _
        " where Left([DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And Left([DateTime],8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "' and Invoice_Totals.Cashier_ID='" & UserID & "' and Invoice_Totals.status='C'" & _
        " GROUP BY Invoice_Totals.Cashier_ID, Left([DateTime],8)"
    Else
        SQL = "SELECT Count(Invoice_Totals.Invoice_Number) AS [Transaction]," & _
        " Count(Invoice_Totals.Discount) AS CountDis, " & _
        " Sum(Invoice_Totals.Total_Price) AS SumTP,Sum(Invoice_Totals.Adjustment1) AS Adj1,Sum(Invoice_Totals.Adjustment2) AS Adj2, Sum(Invoice_Totals.Grand_Total)" & _
        " AS sumGT, Invoice_Totals.Cashier_ID," & _
         " Sum(Invoice_Totals.Adjustment3) AS Adj3,Sum(Invoice_Totals.Adjustment4) AS Adj4,Sum(Invoice_Totals.Adjustment5) AS Adj5,Sum(Invoice_Totals.Adjustment6) AS Adj6," & _
        " sum(Invoice_Totals.Discount*Invoice_Totals.Total_Price/100) AS AmtDis," & _
        " sum(Invoice_Totals.Service_Charge*Invoice_Totals.Total_Price/100) AS AmtSer," & _
        " sum(Invoice_Totals.VATFee*Invoice_Totals.Total_Tax1/100) AS AmtVAT," & _
        " sum(Invoice_Totals.AddMoney) AS AmtReceive,Left([DateTime],8)as Datesale " & _
        " From Invoice_Totals" & _
        " where Left([DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And Left([DateTime],8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "' and Invoice_Totals.status='C'" & _
        " GROUP BY Invoice_Totals.Cashier_ID, Left([DateTime],8)"
    End If
    Set crCashierReport = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crCashierReport
        .Database.AddADOCommand cnData, cmd
        .Transaction.SetUnboundFieldSource "{ado.Transaction}"
        .Total.SetUnboundFieldSource "{ado.SumTP}"
        .CountDis.SetUnboundFieldSource "{ado.CountDis}"
        .txtMoneyReceive.SetUnboundFieldSource "{ado.AmtReceive}"
        .DisAmt.SetUnboundFieldSource "{ado.AmtDis}"
        .Datesale.SetUnboundFieldSource "{ado.Datesale}"
        .CashierID.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtVAT.SetUnboundFieldSource "{ado.AMTVAT}"
        .Adjustment1.SetUnboundFieldSource "{ado.Adj1}"
        .Adjustment2.SetUnboundFieldSource "{ado.Adj2}"
        .Adjustment3.SetUnboundFieldSource "{ado.Adj3}"
        .Adjustment4.SetUnboundFieldSource "{ado.Adj4}"
        .Adjustment5.SetUnboundFieldSource "{ado.Adj5}"
        .Adjustment6.SetUnboundFieldSource "{ado.Adj6}"
        .Sercharge.SetUnboundFieldSource "{ado.AmtSer}"
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        With .Total
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .DisAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtMoneyReceive
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .AmtTotal
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
        With .txtVAT
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    Set iReport = crCashierReport
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
    Unload frmShowCashierReport
    prbBanhang.Value = 0
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdCashierRepot_Click"
End Sub

Private Sub cmdChangeID_Click()
    frmChangeID.Show vbModal
End Sub

Private Sub cmdChangpass_Click()
    frmChangePassword.Show vbModal
End Sub


Private Sub cmdCustomer_Click()
    frmCustomer.Show vbModal
End Sub

Private Sub cmdCustSetup_Click()
    frmIncoming_Outgoing.Show vbModal
End Sub

Private Sub cmdDeletedItems_Click()
On Error GoTo Handle

    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
    
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = "SELECT DISTINCT Items_Deleted.Sec_ID, Items_Deleted.PrintCount,  items_Deleted.Invoice_Num as Invoice_No, Items_Deleted.Table_ID, Items_Deleted.Cashier_ID, Items_Deleted.PluNo, Items_Deleted.Quantity, Items_Deleted.Price, Items_Deleted.Quantity*Items_Deleted.Price AS Amount, Left([DateTime],8) AS DateInvoice, Items_Deleted.Ordered, Items_Deleted.Reason, Right([DateTime],8) AS TimeInvoice, Inventory.ItemName, Inventory.Unit" & _
          " FROM Inventory INNER JOIN Items_Deleted ON Inventory.ItemNum = Items_Deleted.PluNo" & _
          " Where Left([DateTime],8)>='" & gfCONVERT_DATE_TO_STRING(dtpFromDate.Value) & "' and Left([DateTime],8)<='" & gfCONVERT_DATE_TO_STRING(dtpToDate.Value) & "'"

    Set crDeleteItems = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crDeleteItems
        .Database.AddADOCommand cnData, cmd
        .txtserver.SetUnboundFieldSource "{ado.Sec_ID}"
        .txtTable.SetUnboundFieldSource "{ado.Table_ID}"
        .txtBill.SetUnboundFieldSource "{ado.Invoice_No}"
        .txtCashier.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtPluCode.SetUnboundFieldSource "{ado.PluNo}"
        .txtQty.SetUnboundFieldSource "{ado.Quantity}"
        .txtItemName.SetUnboundFieldSource "{ado.ItemName}"
        .txtUnit.SetUnboundFieldSource "{ado.Unit}"
        .txtCost.SetUnboundFieldSource "{ado.Price}"
        .txtAmt.SetUnboundFieldSource "{ado.Amount}"
        .txtReason.SetUnboundFieldSource "{ado.Reason}"
        .txtDate.SetUnboundFieldSource "{ado.DateInvoice}"
        .txtTime.SetUnboundFieldSource "{ado.TimeInvoice}"
        .printcount.SetUnboundFieldSource "{ado.PrintCount}"
        .blOrder.SetUnboundFieldSource "{ado.Ordered}"
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtCost
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Field11
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
    End With
    Set iReport = crDeleteItems
    With frmShowReport_DeleteItems
        .Report = iReport
        .Let_Fromdate = gfCONVERT_DATE_TO_STRING(dtpFromDate.Value)
        .Let_Todate = gfCONVERT_DATE_TO_STRING(dtpToDate.Value)
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdTongbill_Click"
End Sub

Private Sub cmdDept_Click()

    frmDept.Show vbModal
End Sub

Private Sub cmdDone_Click()
    Unload Me
    Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
End Sub

Private Sub cmdEmployee_Click()
    
    frmemployee.Show vbModal
End Sub

Private Sub cmdExpensive_Click()
    frmPhieuchi.Show vbModal
End Sub

Private Sub cmdExpensiveList_Click()
    frmExpenses.Show vbModal
End Sub

Private Sub cmdGiftCard_Click()
    frmGiftCard.Show vbModal
End Sub

Private Sub cmdGroup_Click()
    frmDepartement.Show vbModal
End Sub

Private Sub cmdHourly_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
    
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = "SELECT Invoice_Totals.Invoice_Number, Substring([DateTime],9,2) AS Hourly," & _
    " Sum(Invoice_Totals.Grand_Total) AS TotalAmt, Left([DateTime],8) AS DateInvoice," & _
    " Invoice_Totals.Total_Price, Invoice_Totals.Discount, Invoice_Totals.Station_ID" & _
    " From Invoice_Totals" & _
    " WHERE (((Left([DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & " ')) and Invoice_Totals.Status='C'" & _
    " GROUP BY Invoice_Totals.DateTime, Invoice_Totals.Invoice_Number,Invoice_Totals.Total_Price, Invoice_Totals.Discount, Invoice_Totals.Station_ID,Invoice_Totals.Status;"
    Set crHourly = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crHourly
        .Database.AddADOCommand cnData, cmd
        .BillNO.SetUnboundFieldSource "{ado.Invoice_Number}"
        .fromTime.SetUnboundFieldSource "{ado.Hourly}"
        .totals.SetUnboundFieldSource "{ado.Total_Price}"
        .discount.SetUnboundFieldSource "{ado.Discount}"
        .Amount.SetUnboundFieldSource "{ado.TotalAmt}"
        .server.SetUnboundFieldSource "{ado.Station_ID}"
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        With .Amount
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .totals
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .discount
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Field7
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Field12
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Field13
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Field14
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Field15
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Field16
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Field17
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Field6
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Field4
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Field8
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
    End With
    Set iReport = crHourly
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdHourly_Click"
End Sub

Private Sub cmdImExList_Click()
    frmInOutType.Show vbModal
End Sub

Private Sub cmdInstock_Click()
'   With frmSelectStock
'      .Let_state = "IN"
'      .Show vbModal
'   End With
    frmInstockB.Show vbModal
End Sub

Private Sub cmdInstruction_Click()
    frmGeneralData.Show vbModal
End Sub

Private Sub cmdInvoiceList_Click()
    frmPreviewBill.Show vbModal
End Sub

Private Sub cmdInvoiceSetup_Click()
    frmGeneralBill.Show vbModal
End Sub

Private Sub cmdJobCode_Click()
    frmJobcode.Show vbModal
End Sub

Private Sub cmdKaraoke_Click()
    
End Sub

Private Sub cmdLocationName_Click()
    frmLocationName.Show vbModal
End Sub

Private Sub cmdMaingroup_Click()
    frmMainGroup.Show vbModal
End Sub

Private Sub cmdMaterial_Click()
    frmSetMPLU.Show vbModal
End Sub

Private Sub cmdMedia_Click()
On Error GoTo Handle
    frmMedia.Show vbModal
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "  cmdMedia_Click"
End Sub

Private Sub cmdNotPaymented_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
    
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop
'
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
        SQL = "SELECT Invoice_Totals_Notes.Invoice_Number, Invoice_Totals.Orig_OnHoldID," & _
    "  Left([DateTime],8) AS DateOpen, Invoice_Totals.Grand_Total," & _
    " Invoice_Totals.Station_ID, Invoice_Totals.Status, Invoice_Totals.Cashier_ID" & _
    " FROM Invoice_Totals INNER JOIN Invoice_Totals_Notes ON (Invoice_Totals.Store_ID = Invoice_Totals_Notes.Store_ID)" & _
    " AND (Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number)" & _
    " where Invoice_Totals.Status = 'O' or Invoice_Totals.Status= 'P' and " & _
    " Left(Invoice_Totals.[DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' and Left(Invoice_Totals.[DateTime],8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'" & _
    " order by Invoice_Totals.Invoice_Number"
    Set crNotPaymented = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crNotPaymented
        .Database.AddADOCommand cnData, cmd
        .txtTableNo.SetUnboundFieldSource "{ado.Orig_OnHoldID}"
        .txtBillNo.SetUnboundFieldSource "{ado.Invoice_Number}"
        .txtserver.SetUnboundFieldSource "{ado.Station_ID}"
        .txtCashier.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtAmt.SetUnboundFieldSource "{ado.Grand_Total}"
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        With .txtAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtsumAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
    End With
    Set iReport = crNotPaymented
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdTongbill_Click"
End Sub

Private Sub cmdOutStock_Click()
'   With frmSelectStock
'        .Let_state = "OUT"
'        .Show vbModal
'   End With
    frmOustockB.Show vbModal
End Sub

Private Sub cmdPriceRate_Click()
    frmSetupPrice.Show vbModal
End Sub

Private Sub cmdPrinterName_Click()
On Error GoTo Handle
    
        frmPrintName.Show vbModal
    
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdPrinterName_Click"
End Sub

Private Sub cmdReceipt_Click()
    On Error GoTo Handle
        frmPhieuthu.Show vbModal
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdReceipt_Click "
End Sub

Private Sub cmdReceiptList_Click()
    frmReceipt.Show vbModal
End Sub

Private Sub cmdReport_Click()
On Error GoTo Handle
        fraCashier.Visible = False
        fraSys.Visible = False
        fraList.Visible = False
        fraReport.Visible = True
        fraEditName.Visible = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdCashier_Click"
End Sub

Private Sub cmdReportGroup_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL, strSQL1, strSQL2, strSQL3 As String
    Dim RptID As Integer
    
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = "SELECT  Invoice_Totals.Store_ID, Sum(Invoice_Itemized.Quantity) AS qty,  sum(Invoice_Itemized.PricePer*Invoice_Itemized.Quantity) AS Amount  ," & _
    " Inventory.Dept_ID, Departments.Description" & _
    " FROM Departments INNER JOIN (Inventory INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized" & _
    " ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number)" & _
    " ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Departments.Dept_ID" & _
    " = Inventory.Dept_ID " & _
    " where Left([DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And Left([DateTime],8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "' " & _
    " GROUP BY Invoice_Totals.Store_ID,Inventory.Dept_ID, Departments.Description"
    
    strSQL1 = "SELECT  Sum(Invoice_Itemized.Quantity) AS qty,  sum(Invoice_Itemized.PricePer*Invoice_Itemized.Quantity) AS Amount   ," & _
    " Inventory.Dept_ID, Departments.Description" & _
    " FROM Departments INNER JOIN (Inventory INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized" & _
    " ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number)" & _
    " ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Departments.Dept_ID" & _
    " = Inventory.Dept_ID " & _
    " where Left([DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And Left([DateTime],8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "' " & _
    " GROUP BY Inventory.Dept_ID, Departments.Description"
    
    strSQL2 = "SELECT  Invoice_Totals.Station_ID,Sum(Invoice_Itemized.Quantity) AS qty,  sum(Invoice_Itemized.PricePer*Invoice_Itemized.Quantity) AS Amount    ," & _
    " Inventory.Dept_ID, Departments.Description" & _
    " FROM Departments INNER JOIN (Inventory INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized" & _
    " ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number)" & _
    " ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Departments.Dept_ID" & _
    " = Inventory.Dept_ID " & _
    " where Left([DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And Left([DateTime],8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "' " & _
    " GROUP BY Invoice_Totals.Station_ID,Inventory.Dept_ID, Departments.Description"
    
    strSQL3 = "SELECT MainGroup.GroupNo,Departments.Description,Inventory.Dept_ID, Sum([Quantity]) AS qty,sum(Invoice_Itemized.PricePer*Invoice_Itemized.Quantity) AS Amount, MainGroup.GroupName, Left([DateTime],8) AS Expr3" & _
            " FROM Invoice_Totals INNER JOIN ((Inventory INNER JOIN (MainGroup INNER JOIN Departments ON MainGroup.GroupNo = Departments.MainGroup) ON Inventory.Dept_ID = Departments.Dept_ID) INNER JOIN Invoice_Itemized ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number" & _
            " where Left([DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And Left([DateTime],8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "' " & _
            " group by MainGroup.GroupNo,MainGroup.GroupName, Departments.Description, Left([DateTime],8),Inventory.Dept_ID"
    
    Set crGroup = Nothing
        cmd.ActiveConnection = cnData
        Select Case cboSaleFilter.ListIndex
        Case 0
            cmd.CommandText = strSQL1
        Case 1
            cmd.CommandText = SQL
        Case 2
            cmd.CommandText = strSQL2
        Case 5
            cmd.CommandText = strSQL3
        Case Else
            cmd.CommandText = strSQL1
        End Select
        cmd.Execute
    With crGroup
        .Database.AddADOCommand cnData, cmd
        .txtGroupName.SetUnboundFieldSource "{ado.Description}"
        .txtQty.SetUnboundFieldSource "{ado.qty}"
        .txtAmt.SetUnboundFieldSource "{ado.Amount}"
        If cboSaleFilter.ListIndex = 1 Then
            .txtStoreID.SetUnboundFieldSource "{ado.Store_ID}"
            .Section1.Suppress = False
            .Section2.Suppress = False
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section5.Suppress = True
            .Section11.Suppress = True
        ElseIf cboSaleFilter.ListIndex = 2 Then
            .txtStationID.SetUnboundFieldSource "{ado.Station_ID}"
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section3.Suppress = False
            .Section4.Suppress = False
            .Section5.Suppress = True
            .Section11.Suppress = True
        ElseIf cboSaleFilter.ListIndex = 5 Then
            .txtMaingroup.SetUnboundFieldSource "{ado.GroupName}"
            .txtGroupNo.SetUnboundFieldSource "{ado.GroupNo}"
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section5.Suppress = False
            .Section11.Suppress = False
        Else
            .Section1.Suppress = True
            .Section2.Suppress = True
            .Section3.Suppress = True
            .Section4.Suppress = True
        End If
        
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        With .txtQty
            .DecimalPlaces = DecimalQtyNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtGroupAmount
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtSumStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtAvgPrice
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtsumAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        
    End With
    Set iReport = crGroup
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdBaocaochitiet_Click"
End Sub

Private Sub cmdRightSelection_Click()
    If UserLevel = 1 Or UserID = "881507" Then frmRightSelection.Show vbModal
End Sub

Private Sub cmdSalary_Click()
frmSalary.Show vbModal
End Sub

Private Sub cmdSaleReport_Click()
On Error GoTo Handle
        fraBaocaobanhang.Visible = True
        fraKitchen.Visible = False
        fra80.Visible = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdSaleReport_Click"
End Sub

Private Sub cmdSetMLink_Click()
    frmSetMenuLink.Show vbModal
End Sub

Private Sub cmdSetup_Click()
On Error GoTo Handle
        fraCashier.Visible = False
        fraList.Visible = False
        fraReport.Visible = False
        fraSys.Visible = True
        fraEditName.Visible = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdCashier_Click"
End Sub

Private Sub cmdSetupPrint_Click()
    If ArrayFlag(SF(6), 5) = 1 Then
        frmPrint_Location.Show vbModal
    Else
        frmPrintDefault.Show vbModal
    End If
End Sub

Private Sub cmdShift_Click()
    frmWork_Shift.Show vbModal
End Sub

Private Sub cmdSKU_Click()
    frmItems.Show vbModal
End Sub

Private Sub cmdTable_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
    
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop
'
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = "SELECT Invoice_Totals_Notes.Invoice_Number  as BillNo,Invoice_Totals.Store_ID, Invoice_Totals.Orig_OnHoldID, right(Invoice_Totals_Notes.OpenTime,8) as Opentime," & _
    " substring(Invoice_Totals_Notes.ClosingTime,9,8) as ClosingTime, Left([DateTime],8) AS DateOpen, Invoice_Totals.Grand_Total," & _
    " Invoice_Totals.Station_ID, Invoice_Totals.Status" & _
    " FROM Invoice_Totals INNER JOIN Invoice_Totals_Notes ON (Invoice_Totals.Store_ID = Invoice_Totals_Notes.Store_ID)" & _
    " AND (Invoice_Totals.Invoice_Number = Invoice_Totals_Notes.Invoice_Number)" & _
    " where Left([DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' and Left([DateTime],8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'" & _
    " order by Invoice_Totals.Invoice_Number"
    Set crTableTotal = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crTableTotal
        .Database.AddADOCommand cnData, cmd
        .txtTableNo.SetUnboundFieldSource "{ado.Orig_OnHoldID}"
        .txtBillNo.SetUnboundFieldSource "{ado.BillNo}"
        .txtOpentime.SetUnboundFieldSource "{ado.OpenTime}"
        .txtClosingTime.SetUnboundFieldSource "{ado.ClosingTime}"
        .txtAmount.SetUnboundFieldSource "{ado.Grand_Total}"
        .txtStatus.SetUnboundFieldSource "{ado.Status}"
        
        If cboSaleFilter.ListIndex = 1 Then
            .txtStoreID.SetUnboundFieldSource "{ado.Store_ID}"
            .Section5.Suppress = False
            .Section11.Suppress = False
            .Section12.Suppress = True
            .Section13.Suppress = True
        ElseIf cboSaleFilter.ListIndex = 2 Then
            .txtserver.SetUnboundFieldSource "{ado.Station_ID}"
            .Section12.Suppress = False
            .Section13.Suppress = False
            .Section5.Suppress = True
            .Section11.Suppress = True
        Else
            .Section5.Suppress = True
            .Section11.Suppress = True
            .Section12.Suppress = True
            .Section13.Suppress = True
        End If
        
        
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        With .txtAmount
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtSumStore
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtSumStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtsumAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With

    End With
    Set iReport = crTableTotal
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdTongbill_Click"
End Sub

Private Sub cmdTaxRate_Click()
    frmTaxRate.Show vbModal
End Sub


Private Sub cmdThuchi_Click()
On Error GoTo Handle
        fraBaocaobanhang.Visible = False
        fraKitchen.Visible = False
        fra80.Visible = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdThuchi_Click"
End Sub

Private Sub cmdThuchiSys_Click()
On Error GoTo Handle
        fraCashier.Visible = False
        fraSys.Visible = False
        fraList.Visible = False
        fraReport.Visible = False
        fraEditName.Visible = False
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  cmdCashier_Click"
End Sub

Private Sub cmdTichluy_Click()
    Set cnData = Get_Connection(BK_ServerName, BK_DataBaseName, BK_UserLog, BK_DB_Password)
End Sub

Private Sub cmdTimerSetup_Click()
    With frmPassword
        .FormActionKey = "SetColor"
        .Show vbModal
    End With
    
End Sub

Private Sub cmdTon80_Click()
On Error GoTo Handle
If Month(dtpFromDate.Value) = Month(dtpToDate.Value) Then
    With frmShowStockB
        .Report_ID = 5
        .Get_FromDate = gfCONVERT_DATE_TO_STRING(dtpFromDate.Value)
        .Get_ToDate = gfCONVERT_DATE_TO_STRING(dtpToDate.Value)
        .Show vbModal
    End With
    
Else
    MsgBox "B¸o c¸o kho ®­îc tÝnh trong 1 th¸ng"
End If
Exit Sub
Handle:
    MsgBox Err.Description & Me.name & " cmdTon80_Click"
End Sub

Private Sub cmdTongbill_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim rs As New Recordset
    Dim SQL As String
    Dim RptID As Integer
    
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    
'    SQL = "SELECT Invoice_Totals.Invoice_Number,Invoice_Totals.Store_ID,Invoice_Totals.Station_ID, Invoice_Totals.CustNum, Left([DateTime],8) AS DateInvoice," & _
'    " Invoice_Totals.Total_Cost, Invoice_Totals.Total_Price, Invoice_Totals.Discount, Invoice_Totals.Grand_Total," & _
'    " Invoice_Totals.AddMoney, Invoice_Totals.Cashier_ID" & _
'    " FROM Invoice_Totals" & _
'    " WHERE (((Invoice_Totals.Status)='C')) and Left([DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And Left([DateTime], 8) <= '" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'" & _
'    " order by Invoice_Totals.Invoice_Number"
'
SQL = "SELECT right('0000' & Invoice_Totals.Invoice_Number,4)as billNo,Invoice_Totals.Orig_OnHoldID as TableNo," & _
        "  Invoice_Totals.Payment_Method,Invoice_Totals.InvType,Invoice_Totals.Store_ID, Invoice_Totals.Station_ID," & _
        " Invoice_Totals.CustNum, Left([DateTime],8) AS DateInvoice, Invoice_Totals.Total_Cost, Invoice_Totals.Total_Price," & _
        " Invoice_Totals.Adjustment1,Invoice_Totals.Adjustment2,(Invoice_Totals.Total_Tax1*Invoice_Totals.VATFee/100)as VAT, " & _
        " Invoice_Totals.Discount, Invoice_Totals.Grand_Total, Invoice_Totals.AddMoney, Invoice_Totals.Cashier_ID," & _
        " (Invoice_Totals.Total_Price*Invoice_Totals.Service_Charge/100)as Service, Invoice_Totals.Personals" & _
      " from Invoice_Totals " & _
      " WHERE Left([DateTime],8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And Left([DateTime], 8) <= '" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'" & _
      " and Invoice_Totals.Status<>'CO' and  Invoice_Totals.Status<>'O' and left(Invoice_Totals.Status,1)<>'T' " & _
      " GROUP BY Invoice_Totals.Invoice_Number, Invoice_Totals.Orig_OnHoldID,Invoice_Totals.Store_ID,Invoice_Totals.InvType,Invoice_Totals.Payment_Method, Invoice_Totals.Station_ID, Invoice_Totals.CustNum, Left([DateTime],8), Invoice_Totals.Total_Cost, Invoice_Totals.Total_Price," & _
      " Invoice_Totals.Discount, Invoice_Totals.Grand_Total, Invoice_Totals.AddMoney, Invoice_Totals.Cashier_ID," & _
      " Invoice_Totals.Personals,(Invoice_Totals.Total_Tax1*Invoice_Totals.VATFee/100),(Invoice_Totals.Total_Price*Invoice_Totals.Service_Charge/100),Adjustment2,Adjustment1"

    
    Set crBillTotal = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crBillTotal
        .Database.AddADOCommand cnData, cmd
        .txtBillDate.SetUnboundFieldSource "{ado.DateInvoice}"
        .txtBillNo.SetUnboundFieldSource "{ado.billNo}"
        .txtTable.SetUnboundFieldSource "{ado.TableNo}"
        .txtCustNo.SetUnboundFieldSource "{ado.CustNum}"
        .TxtTotal.SetUnboundFieldSource "{ado.Total_Price}"
        .txtDiscount.SetUnboundFieldSource "{ado.Discount}"
        .txtTotalAmt.SetUnboundFieldSource "{ado.Grand_Total}"
        .txtAdj1.SetUnboundFieldSource "{ado.Adjustment1}"
        .txtAdj2.SetUnboundFieldSource "{ado.Adjustment2}"
        .txtService.SetUnboundFieldSource "{ado.Service}"
        .txtsokhach.SetUnboundFieldSource "{ado.Personals}"
        .txtReceiveMoney.SetUnboundFieldSource "{ado.AddMoney}"
        .txtPaymentMethod.SetUnboundFieldSource "{ado.Payment_Method}"
        .txtCashierID.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtVAT.SetUnboundFieldSource "{ado.VAT}"
        .txtPrintCount.SetUnboundFieldSource "{ado.InvType}"
        If cboSaleFilter.ListIndex = 1 Then
            .txtStoreID.SetUnboundFieldSource "{ado.Store_ID}"
            .Section3.Suppress = False
            .Section4.Suppress = False
            .Section11.Suppress = True
            .Section5.Suppress = True
        ElseIf cboSaleFilter.ListIndex = 2 Then
            .txtStationID.SetUnboundFieldSource "{ado.Station_ID}"
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section11.Suppress = False
            .Section5.Suppress = False
        Else
            .Section3.Suppress = True
            .Section4.Suppress = True
            .Section5.Suppress = True
            .Section11.Suppress = True
        End If
        
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        
        With .txtTotalStore
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtDisAmountStore
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtDisAmountStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalAmtStore
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalAmtStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtSokhachStore
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtSokhachStation
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtAmtDis
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .TxtTotal
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtTotalAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtDiscount
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Field1
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .Field3
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        With .txtSumTotalAmt
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
    End With
    Set iReport = crBillTotal
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " cmdTongbill_Click"
End Sub


Private Sub cmdTranfer_Click()
On Error GoTo Handle
    MsgBox "B¹n nªn tÝnh tån kho tr­íc khi chuyÓn kho, ®Ó sè l­îng trong kho chÝnh ®­îc chÝnh x¸c", vbInformation
    If MsgBox("B¹n cã tÝnh tån kho kh«ng?", vbYesNo) = vbYes Then
        With frmCal_TonTemp
            .Show vbModal
        End With
    End If
    frmTranstock.Show vbModal
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdTranfer_Click"
End Sub

Private Sub cmdVAT_Click()
    frmVAT.Show vbModal
End Sub

Private Sub cmdVendor_Click()
    frmSupplier.Show vbModal
End Sub


Private Sub cmdXNT80_Click()
On Error GoTo Handle
    With frmShowStockB
        .Report_ID = 7
        .Get_FromDate = gfCONVERT_DATE_TO_STRING(dtpFromDate.Value)
        .Get_ToDate = gfCONVERT_DATE_TO_STRING(dtpToDate.Value)
        .Show vbModal
    End With
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " - cmdXNT80_Click"
End Sub

Private Sub cmdXuat80_Click()
On Error GoTo Handle
    With frmShowStockB
        .Report_ID = 6
        .Get_FromDate = gfCONVERT_DATE_TO_STRING(dtpFromDate.Value)
        .Get_ToDate = gfCONVERT_DATE_TO_STRING(dtpToDate.Value)
        .Show vbModal
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & "- cmdXuat80_Click"
End Sub

Private Sub cmdReport_Location_Click()
On Error GoTo errHdl
    Dim CRReport As New CRAXDDRT.Report
    Dim SQL As String
    Dim cmd As New ADODB.Command
    Dim FromDate, ToDate As String
    
    FromDate = Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00")
    ToDate = Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00")
    If UserLevel = 1 Then
        SQL = "SELECT sum(Invoice_Totals.Grand_Total) as total, Table_Diagram_Sections.Section_ID, Invoice_Totals.Station_ID" & _
            " FROM Table_Diagram_Sections INNER JOIN Invoice_Totals ON Table_Diagram_Sections.Location_ID = Invoice_Totals.Station_ID" & _
            " where left(Invoice_Totals.DateTime,8)>= '" & FromDate & "' and left(Invoice_Totals.DateTime,8)<='" & ToDate & "'" & _
            " and Invoice_Totals.Status<>'CO' and  Invoice_Totals.Status<>'O' and left(Invoice_Totals.Status,1)<>'T' " & _
            " Group by Table_Diagram_Sections.Section_ID, Invoice_Totals.Station_ID"
    Else
        SQL = "SELECT  sum(Invoice_Totals.Grand_Total) as total, Table_Diagram_Sections.Section_ID, Invoice_Totals.Station_ID" & _
            " FROM Table_Diagram_Sections INNER JOIN Invoice_Totals ON Table_Diagram_Sections.Location_ID = Invoice_Totals.Station_ID" & _
            " where left(Invoice_Totals.DateTime,8)>= '" & FromDate & "' and left(Invoice_Totals.DateTime,8)<='" & ToDate & "' and Invoice_Totals.Cashier_ID='" & UserID & "'" & _
            " and Invoice_Totals.Status<>'CO' and  Invoice_Totals.Status<>'O' and left(Invoice_Totals.Status,1)<>'T' " & _
            " Group by Table_Diagram_Sections.Section_ID, Invoice_Totals.Station_ID"
    End If
    Set crLocation = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crLocation
        .Database.AddADOCommand cnData, cmd
        .TxtTotal.SetUnboundFieldSource "{ado.Total}"
        .Location.SetUnboundFieldSource "{ado.Section_ID}"
'        .Cashier.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtLocation.SetUnboundFieldSource "{ado.Station_ID}"
        .txtFromDate.SetText gfCONVERT_STRING_TO_DATE(FromDate)
        .txtToDate.SetText gfCONVERT_STRING_TO_DATE(ToDate)
        With .TxtTotal
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark
        End With
    End With
    
    Set iReport = crLocation
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = iReport
        .Show vbModal
    End With
Exit Sub
errHdl:
Exit Sub
    MsgBox Err.Number & " - cmdBillList_Click - " & Err.Description
End Sub

Private Sub cmdEmp_Totals_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL As String
    
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop

'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    'Khong danh cho karaoke
    SQL = "SELECT Employee.Cashier_ID, Employee.EmpName, Sum(Invoice_Totals.Grand_Total) AS total" & _
                " FROM Employee INNER JOIN Invoice_Totals ON Employee.Cashier_ID = Invoice_Totals.OrderMan" & _
               " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
              " and [Invoice_Totals].[Status]<>'CO' and left(invoice_totals.status,1)<>'T'" & _
              " GROUP BY Employee.Cashier_ID, Employee.EmpName"

    Set crEmp_Totals = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crEmp_Totals
        .Database.AddADOCommand cnData, cmd
        .txtEmpID.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtEmpName.SetUnboundFieldSource "{ado.EmpName}"
        .txtTotals.SetUnboundFieldSource "{ado.total}"
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
        With .txtTotals
            .DecimalPlaces = DecimalAmtNumber
            .DecimalSymbol = DecimalMark
            .ThousandsSeparators = True
            .ThousandSymbol = DigitGroupMark

        End With
        
    End With
    Set iReport = crEmp_Totals
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
Exit Sub
MsgBox Err.Number & Err.Description & Me.name & " cmdEmp_Totals_Click "
End Sub

Private Sub CommandButton1_Click()

End Sub

Private Sub CommandButton2_Click()
    frmReport_Emp.Show vbModal
End Sub

Private Sub CommandButton3_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL As String
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = " SELECT Customer.CustNum, Customer.CustName, Invoice_Totals.Cashier_ID, Invoice_Totals.Discount, Count(Invoice_Totals.CustNum) AS Num, Invoice_Totals.DateTime" & _
        " FROM Customer INNER JOIN Invoice_Totals ON Customer.CustNum = Invoice_Totals.CustNum" & _
        " GROUP BY Customer.CustNum, Customer.CustName, Invoice_Totals.Cashier_ID, Invoice_Totals.Discount, Invoice_Totals.DateTime" & _
        " HAVING (((Customer.CustNum)<>'101') AND ((Invoice_Totals.Discount)>0)" & _
        "AND (left(Invoice_Totals.DateTime,8)>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And Left(Invoice_Totals.DateTime,8)<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))"


    Set crCustSaleReport = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    With crCustSaleReport
        .Database.AddADOCommand cnData, cmd
        .txtCustID.SetUnboundFieldSource "{ado.CustNum}"
        .txtCustName.SetUnboundFieldSource "{ado.CustName}"
        .txtDiscount.SetUnboundFieldSource "{ado.Discount}"
        .txtTotals.SetUnboundFieldSource "{ado.Num}"
        .txtCashierID.SetUnboundFieldSource "{ado.Cashier_ID}"
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
    End With
    Set iReport = crCustSaleReport
    With frmShowReport
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & " cmdBaocaochitiet_Click"
End Sub

Private Sub cmdMainGroup80_Click()
On Error GoTo Handle
    Dim cmd As New ADODB.Command
    Dim SQL As String
    Dim CRReport As CRAXDDRT.Report
    Do While prbBanhang.Value < prbBanhang.Max
        prbBanhang.Value = prbBanhang.Value + 1
    Loop
'    If cnData.State = 0 Then
'        Set cnData = Get_Connection(WorkingFolder & "\Database.mdb", "100881administrator")
'    End If
    SQL = "SELECT MainGroup.GroupName, sum(Invoice_Itemized.Quantity) as QTY," & _
          " sum(Invoice_Itemized.Quantity*Invoice_Itemized.PricePer) as Amount" & _
          " FROM ((MainGroup INNER JOIN Departments ON MainGroup.GroupNo = " & _
          " Departments.MainGroup) INNER JOIN Inventory ON Departments.Dept_ID =" & _
          " Inventory.Dept_ID) INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized" & _
          " ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number)" & _
          " ON Inventory.ItemNum = Invoice_Itemized.ItemNum" & _
          " WHERE (((Left([Invoice_Totals].[DateTime],8))>='" & Format(Year(dtpFromDate.Value), "0000") & Format(Month(dtpFromDate.Value), "00") & Format(Day(dtpFromDate.Value), "00") & "' And (Left([Invoice_Totals].[DateTime],8))<='" & Format(Year(dtpToDate.Value), "0000") & Format(Month(dtpToDate.Value), "00") & Format(Day(dtpToDate.Value), "00") & "'))" & _
          " GROUP BY  MainGroup.GroupName"
    Set crMainGroup = Nothing
    Set crMainGroup58 = Nothing
        cmd.ActiveConnection = cnData
        cmd.CommandText = SQL
        cmd.Execute
    If ReceiptType = 80 Then
        Set CRReport = crMainGroup
    Else
        Set CRReport = crMainGroup58
    End If
    
    With CRReport
        .Database.AddADOCommand cnData, cmd
        .txtGroupName.SetUnboundFieldSource "{ado.GroupName}"
        .txtQty.SetUnboundFieldSource "{ado.QTY}"
        .txtAmount.SetUnboundFieldSource "{ado.Amount}"
        .txtFromDate.SetText dtpFromDate.Value
        .txtToDate.SetText dtpToDate.Value
'canh le

        .TopMargin = TopAlign
        .BottomMargin = BottomAlign
        .LeftMargin = LeftAlign
        .RightMargin = RightAlign
       
    End With
    Set iReport = CRReport
    With frmShow_Report_80
        .Let_Printer = GetSettingStr("Report", "Report_DeviceName", True, myIniFile)
        .Report = iReport
        .Show vbModal
    End With
    prbBanhang.Value = 0
Exit Sub
Handle:
    Exit Sub
    MsgBox Err.Number & Err.Description & Me.name & " cmdBaocaochitiet_Click"
End Sub

Private Sub dtpFromDate_Change()
'On Error GoTo Handle
'    If Format(Year(dtpFromDate.Value), "0000") <= Format(Year(Date), "0000") Then
'        If Format(Month(dtpFromDate.Value), "00") <= Format(Month(Date), "00") Then
'            strDateTime = Format(Month(Date), "00") & Format(Year(Date), "0000")
'            Set cnData = Get_Connection(ReportFolder & strDateTime & "\database.mdb", "100881administrator")
'        Else
'            MsgBox "Th¸ng " & Month(dtpFromDate.Value) & " ch­a cã d÷ liÖu"
'
'        End If
'    Else
'        MsgBox "N¨m" & Year(dtpFromDate.Value) & " ch­a cã d÷ liÖu"
'    End If
'Exit Sub
'Handle:
'MsgBox Err.Number & Err.Description & Me.Name
End Sub

Private Sub dtpToDate_Change()
'On Error GoTo Handle
'    If Format(Year(dtpToDate.Value), "0000") <= Format(Year(Date), "0000") Then
'        If dtpToDate.Value < dtpFromDate.Value Then
'            MsgBox "N¨m ®Õn kg thÓ nhá h¬n n¨m b¾t ®Çu !!!"
'        Else
'            If Format(Month(dtpToDate.Value), "00") <= Format(Month(Date), "00") Then
'                Set cnData = Get_Connection(ReportFolder & strDateTime & "\database.mdb", "100881administrator")
'            Else
'                MsgBox "Th¸ng" & Month(dtpToDate.Value) & " ch­a cã d÷ liÖu"
'
'            End If
'        End If
'    Else
'        MsgBox "N¨m" & Year(dtpToDate.Value) & " ch­a cã d÷ liÖu"
'    End If
'Exit Sub
'Handle:
'MsgBox Err.Number & Err.Description & Me.Name
End Sub

Private Sub Form_Activate()
On Error GoTo Handle
    Dim ctrl As Control
    'If cmdAddCashier.Font.Name <> CurFont Then Call Set_Language(Me, CurFont)
    DescArr = LoadLanguage(LngFile, "#02:009:")
    DescArrReport = LoadLanguage(LngFile, "#05:001:")
    For Each ctrl In Me
        DoEvents
        If Left(ctrl.Tag, 1) = "L" Then ctrl.Caption = DescArr(Mid(ctrl.Tag, 2))
    Next ctrl
    If UserLevel <> 1 Then
        Call CheckRight
    Else
        cmdTichluy.Visible = True
        cmdVAT.Visible = True
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " Form_Activate"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    End If
    If KeyCode = vbKeyF5 Then
        Set cnData = Get_Connection(BK_ServerName, BK_DataBaseName, BK_UserLog, BK_DB_Password)
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo Handle
        DescArr = LoadLanguage(LngFile, "#02:009:")
        fraCashier.Visible = False
        fraReport.Visible = True
        fraList.Visible = False
        fraSys.Visible = False
        fraEditName.Visible = False
        dtpFromDate.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
        dtpToDate.Value = gfCONVERT_STRING_TO_DATE(DateDefault)
        strDateTime = Mid(DateDefault, 5, 2) & Left(DateDefault, 4)
        Call AddSort
        Call AddFilter
        Call AddSort80
        Call AddFilter80
        Call AddKit_Filter
        Call AddKit_Sort
        Call addPrinter
        Call cmdReport_Click
    Exit Sub
Handle:
    MsgBox Err.Number & Err.Description & Me.name & "  Form_Load"
End Sub

Private Sub MyButton1_Click()
    fraBaocaobanhang.Visible = False
    fraKitchen.Visible = False
    fra80.Visible = False

End Sub
Public Sub Get_Data_Report_General(strFromDate As String, strToDate As String)
On Error GoTo Handle
Dim SQL, strDeleteOrder, strDeleenNotOrder, strPayout, strReceipt, strDiscount As String
Dim strCA, strOA, strCT, strROA, strGC, strCC, strOpening As String
Dim strSer, strAdj1, strAdj2, strAdj3, strAdj4, strAdj5, strAdj6, strVAT, strPersonal, strLineDiscount As String
Dim rsDeleteOrder As New ADODB.Recordset
Dim rsDeleteNotOrder As New ADODB.Recordset
Dim rsGeneral As New ADODB.Recordset
Dim rsdiscount As New ADODB.Recordset
Dim rsReceipt As New ADODB.Recordset
Dim rsPayout As New ADODB.Recordset
Dim rsGroup As New ADODB.Recordset
Dim rsTrans As New ADODB.Recordset
Dim rsCA As New ADODB.Recordset
Dim rsOA As New ADODB.Recordset
Dim rsCT As New ADODB.Recordset
Dim rsCC As New ADODB.Recordset
Dim rsGC As New ADODB.Recordset
Dim rsROA As New ADODB.Recordset
Dim rsVAT As New ADODB.Recordset
Dim rsPersonal As New ADODB.Recordset
Dim rsLineDiscount As New ADODB.Recordset
Dim rsOpening As New ADODB.Recordset
Dim rsService_Charge As New ADODB.Recordset
Dim rsAdjustment1 As New ADODB.Recordset
Dim rsAdjustment2 As New ADODB.Recordset
Dim rsAdjustment3 As New ADODB.Recordset
Dim rsAdjustment4 As New ADODB.Recordset
Dim rsAdjustment5 As New ADODB.Recordset
Dim rsAdjustment6 As New ADODB.Recordset

Dim strReceive As String
Dim rsReceive As New ADODB.Recordset
Dim strReserve As String
Dim rsReserve As New ADODB.Recordset
'Xoa toan bo du lieu trong bao cao tong hop
cnData.Execute "Delete  from RP_General"

' Loc lay tat ca cac nhom hang ban duoc
    SQL = "SELECT   Sum(Invoice_Itemized.Quantity) AS qty,  sum(Invoice_Itemized.PricePer*Invoice_Itemized.Quantity) AS Amount    ," & _
    " Inventory.Dept_ID, Departments.Description" & _
    " FROM Departments INNER JOIN (Inventory INNER JOIN (Invoice_Totals INNER JOIN Invoice_Itemized" & _
    " ON Invoice_Totals.Invoice_Number = Invoice_Itemized.Invoice_Number)" & _
    " ON Inventory.ItemNum = Invoice_Itemized.ItemNum) ON Departments.Dept_ID" & _
    " = Inventory.Dept_ID " & _
    " where Left([DateTime],8)>='" & strFromDate & "' And Left([DateTime],8)<='" & strToDate & "' and invoice_totals.status<>'CO' and left(invoice_totals.status,1)<>'T' " & _
    " GROUP BY Inventory.Dept_ID, Departments.Description"
    
' Lay tat ca cac mat hang bi xoa
    strDeleteOrder = "SELECT Count(Items_Deleted.PluNo) AS CountDelete," & _
    " Sum(Items_Deleted.Quantity * Items_Deleted.Price) AS SumAmount" & _
    " From Items_Deleted " & _
    " WHERE (((Left(Items_Deleted.DateTime,8))>='" & strFromDate & "' And (Left(Items_Deleted.DateTime,8))<='" & strToDate & "')) and  Items_Deleted.Ordered=true"
    
'Lay tat ca cac records thu
    strReceipt = "SELECT count( Income.ID) AS CountReceipt, sum(Income.Amount) AS Amount" & _
                " From Income" & _
                " WHERE Income.DateTime>='" & strFromDate & "' and Income.DateTime<='" & strToDate & "'"
'Lay Giam % Mon
    strLineDiscount = "SELECT count(Invoice_Totals.Invoice_Number)as CountLineDisc," & _
                     " sum([Quantity]*[PricePer]*[LineDisc]/100) AS AmtLineDis" & _
                     " FROM Invoice_Totals INNER JOIN Invoice_Itemized ON" & _
                     " Invoice_Totals.Invoice_Number=Invoice_Itemized.Invoice_Number" & _
                     " WHERE (((Left(Invoice_Totals.DateTime,8))>='" & strFromDate & "' And" & _
                     " (Left(Invoice_Totals.DateTime,8))<='" & strToDate & "')) and" & _
                     " Invoice_Itemized.LineDisc>0 and Invoice_totals.Status<>'CO'"
' Lay tat ca cac records chi
    strPayout = "SELECT count(Payouts.ID) as CountPayOut, sum(Payouts.Amount) as Amount" & _
                " From Payouts" & _
                " Where  Payouts.DateTime >='" & strFromDate & "' and Payouts.DateTime<= '" & strToDate & "'"

' Lay tat ca so lam discount va so tien
    strDiscount = "SELECT Count(Invoice_Totals.Invoice_Number) as CountDis," & _
                  " sum(Invoice_Totals.Discount* Invoice_Totals.Total_Price/100) as Amount" & _
                  " From Invoice_Totals " & _
                  " Where left(Invoice_Totals.DateTime,8)>='" & strFromDate & "' and  left(Invoice_Totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Discount>0"
'Lay tat ca ca phi thu tien mat
    strReceive = "SELECT Count(Invoice_Totals.Invoice_Number) as CountReceive," & _
                  " sum(Invoice_Totals.AddMoney) as AmountReceive" & _
                  " From Invoice_Totals " & _
                  " Where left(Invoice_Totals.DateTime,8)>='" & strFromDate & "' and  left(Invoice_Totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.AddMoney>0"

'Lay tat ca ca mon bi xoa khi chua Order
strDeleenNotOrder = "SELECT Count(Items_Deleted.PluNo) AS CountDelete," & _
    " Sum(Items_Deleted.Quantity*Items_Deleted.Price) AS SumAmount" & _
    " From Items_Deleted " & _
    " WHERE (((Left(Items_Deleted.DateTime,8))>='" & strFromDate & "' And (Left(Items_Deleted.DateTime,8))<='" & strToDate & "')) and  Items_Deleted.Ordered=false"
strCA = " Select count(Invoice_Totals.Invoice_Number) AS CountCA, sum( Invoice_Totals.CA_Amount) AS AmountCA from Invoice_Totals" & _
    " where left(Invoice_totals.DateTime,8)>='" & strFromDate & "' and left(Invoice_totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Status='C' and Invoice_Totals.CA_Amount>=0"
strOA = "Select count(Invoice_Totals.Invoice_Number) AS CountOA, sum( Invoice_Totals.OA_Amount) AS AmountOA from Invoice_Totals" & _
    " where left(Invoice_totals.DateTime,8)>='" & strFromDate & "' and left(Invoice_totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Status='OA' and Invoice_Totals.OA_Amount>0 "
strCT = "Select count(Invoice_Totals.Invoice_Number) AS CountCheck, sum( Invoice_Totals.CT_Amount) AS AmountCheck from Invoice_Totals" & _
    " where left(Invoice_totals.DateTime,8)>='" & strFromDate & "' and left(Invoice_totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Status='CT' and Invoice_Totals.CT_Amount>0"
strCC = "Select count(Invoice_Totals.Invoice_Number) AS CountCredit, sum( Invoice_Totals.CC_Amount) AS AmountCredit from Invoice_Totals" & _
    " where left(Invoice_totals.DateTime,8)>='" & strFromDate & "' and left(Invoice_totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Status='CC' and Invoice_Totals.CC_Amount>0"
strGC = "Select count(Invoice_Totals.Invoice_Number) AS CountGC, sum( Invoice_Totals.GC_Amount) AS AmountGC from Invoice_Totals" & _
    " where left(Invoice_totals.DateTime,8)>='" & strFromDate & "' and left(Invoice_totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Status='GC' and Invoice_Totals.GC_Amount>0"
strROA = "Select count(Invoice_Totals.Invoice_Number) AS CountROA, sum( Invoice_Totals.ROA_Amount) AS AmountROA from Invoice_Totals" & _
    " where left(Invoice_totals.DateTime,8)>='" & strFromDate & "' and left(Invoice_totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Status='ROA' and Invoice_Totals.ROA_Amount>0"
strOpening = "Select count(Invoice_Totals.Invoice_Number) AS CountOpen, sum( Invoice_Totals.Grand_Total) AS AmountOpen from Invoice_Totals" & _
    " where  Invoice_Totals.Status='O' or Invoice_Totals.Status='P' and left(Invoice_totals.DateTime,8)>='" & strFromDate & "' and left(Invoice_Totals.DateTime,8)<='" & strToDate & "' "

strSer = "SELECT Count(Invoice_Totals.Invoice_Number) as CountSer," & _
                  " sum(Invoice_Totals.Service_Charge* Invoice_Totals.Total_Price/100) as Amount" & _
                  " From Invoice_Totals " & _
                  " Where left(Invoice_Totals.DateTime,8)>='" & strFromDate & "' and  left(Invoice_Totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Service_Charge>0"
strVAT = "SELECT Count(Invoice_Totals.Invoice_Number) as VATCount," & _
                  " sum(Invoice_Totals.VATFee* Invoice_Totals.Total_Tax1/100) as VATAmount" & _
                  " From Invoice_Totals " & _
                  " Where left(Invoice_Totals.DateTime,8)>='" & strFromDate & "' and  left(Invoice_Totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.VATFee>0"
strPersonal = "SELECT sum(Invoice_Totals.Personals) as Personal" & _
                  " From Invoice_Totals " & _
                  " Where left(Invoice_Totals.DateTime,8)>='" & strFromDate & "' and  left(Invoice_Totals.DateTime,8)<='" & strToDate & "'"
                  
strAdj1 = "SELECT sum(Invoice_Totals.Adjustment1) as Amount, Count(Invoice_Totals.Adjustment1) AS CountAdj1" & _
                  " From Invoice_Totals " & _
                  " Where left(Invoice_Totals.DateTime,8)>='" & strFromDate & "' and  left(Invoice_Totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Adjustment1<>0 and Invoice_Totals.Status<>'CO'"
strAdj2 = "SELECT sum(Invoice_Totals.Adjustment2) as Amount, Count(Invoice_Totals.Adjustment2) AS CountAdj2" & _
                  " From Invoice_Totals " & _
                  " Where left(Invoice_Totals.DateTime,8)>='" & strFromDate & "' and  left(Invoice_Totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Adjustment2<>0 and Invoice_Totals.Status<>'CO'"
         
strAdj3 = "SELECT sum(Invoice_Totals.Adjustment3) as Amount, Count(Invoice_Totals.Adjustment3) AS CountAdj3" & _
                  " From Invoice_Totals " & _
                  " Where left(Invoice_Totals.DateTime,8)>='" & strFromDate & "' and  left(Invoice_Totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Adjustment3<>0 and Invoice_Totals.Status<>'CO'"
         
strAdj4 = "SELECT sum(Invoice_Totals.Adjustment4) as Amount, Count(Invoice_Totals.Adjustment4) AS CountAdj4" & _
                  " From Invoice_Totals " & _
                  " Where left(Invoice_Totals.DateTime,8)>='" & strFromDate & "' and  left(Invoice_Totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Adjustment4<>0 and Invoice_Totals.Status<>'CO'"
         
strAdj5 = "SELECT sum(Invoice_Totals.Adjustment5) as Amount, Count(Invoice_Totals.Adjustment5) AS CountAdj5" & _
                  " From Invoice_Totals " & _
                  " Where left(Invoice_Totals.DateTime,8)>='" & strFromDate & "' and  left(Invoice_Totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Adjustment5<>0 and Invoice_Totals.Status<>'CO'"
         
strAdj6 = "SELECT sum(Invoice_Totals.Adjustment6) as Amount, Count(Invoice_Totals.Adjustment6) AS CountAdj6" & _
                  " From Invoice_Totals " & _
                  " Where left(Invoice_Totals.DateTime,8)>='" & strFromDate & "' and  left(Invoice_Totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Adjustment6<>0 and Invoice_Totals.Status<>'CO'"
         
strReserve = "SELECT Count(Invoice_Totals.Invoice_Number) as CountReserve," & _
                  " sum(Invoice_Totals.Reserve) as AmountReserve" & _
                  " From Invoice_Totals " & _
                  " Where left(Invoice_Totals.DateTime,8)>='" & strFromDate & "' and  left(Invoice_Totals.DateTime,8)<='" & strToDate & "' and Invoice_Totals.Reserve>0 and Invoice_Totals.Status<>'CO'"
                  
' Mo tat ca cac bang de Insert du lieu
Set rsGeneral = OpenCriticalTable("select * from RP_General", cnData)
Set rsdiscount = OpenCriticalTable(strDiscount, cnData)
Set rsDeleteNotOrder = OpenCriticalTable(strDeleenNotOrder, cnData)
Set rsDeleteOrder = OpenCriticalTable(strDeleteOrder, cnData)
Set rsReceipt = OpenCriticalTable(strReceipt, cnData)
Set rsPayout = OpenCriticalTable(strPayout, cnData)
Set rsGroup = OpenCriticalTable(SQL, cnData)
Set rsTrans = OpenCriticalTable("Select count(Invoice_Totals.Invoice_Number) as CountTrans from Invoice_Totals where left(Invoice_totals.DateTime,8)>='" & strFromDate & "' and left(Invoice_totals.DateTime,8)<='" & strToDate & "'", cnData)
Set rsCA = OpenCriticalTable(strCA, cnData)
Set rsOA = OpenCriticalTable(strOA, cnData)
Set rsCT = OpenCriticalTable(strCT, cnData)
Set rsGC = OpenCriticalTable(strGC, cnData)
Set rsROA = OpenCriticalTable(strROA, cnData)
Set rsCC = OpenCriticalTable(strCC, cnData)
Set rsOpening = OpenCriticalTable(strOpening, cnData)
Set rsService_Charge = OpenCriticalTable(strSer, cnData)
Set rsVAT = OpenCriticalTable(strVAT, cnData)
Set rsPersonal = OpenCriticalTable(strPersonal, cnData)
Set rsAdjustment1 = OpenCriticalTable(strAdj1, cnData)
Set rsAdjustment2 = OpenCriticalTable(strAdj2, cnData)
Set rsAdjustment3 = OpenCriticalTable(strAdj3, cnData)

Set rsAdjustment4 = OpenCriticalTable(strAdj4, cnData)
Set rsAdjustment5 = OpenCriticalTable(strAdj5, cnData)
Set rsAdjustment6 = OpenCriticalTable(strAdj6, cnData)

Set rsReceive = OpenCriticalTable(strReceive, cnData)
Set rsLineDiscount = OpenCriticalTable(strLineDiscount, cnData)
Set rsReserve = OpenCriticalTable(strReserve, cnData)


With rsGeneral
    
    Do While Not rsGroup.EOF
        .addNew
        .Fields("Dept_ID") = rsGroup.Fields("Dept_ID")
        .Fields("Description") = rsGroup.Fields("Description")
        .Fields("Qty") = rsGroup.Fields("Qty")
        .Fields("AVG_Price") = rsGroup.Fields("Amount") / rsGroup.Fields("Qty")
        .Fields("Amount") = rsGroup.Fields("Amount")
        .Fields("CountDeleteOrdered") = rsDeleteOrder.Fields("CountDelete")
        .Fields("AmountDeleteOrdered") = rsDeleteOrder.Fields("SumAmount")
        .Fields("CountDelete") = rsDeleteNotOrder.Fields("CountDelete")
        .Fields("AmountCountDelete") = rsDeleteNotOrder.Fields("SumAmount")
        .Fields("CountDist") = rsdiscount.Fields("CountDis")
        .Fields("AmountDist") = rsdiscount.Fields("Amount")
        .Fields("CountReceipt") = CDbl("0" & rsReceipt.Fields("CountReceipt"))
        .Fields("AmountReceipt") = CDbl("0" & rsReceipt.Fields("Amount"))
        .Fields("CountPayouts") = CDbl("0" & rsPayout.Fields("CountPayOut"))
        .Fields("AmountPayouts") = CDbl("0" & rsPayout.Fields("Amount"))
        .Fields("CountTrans") = rsTrans.Fields("CountTrans")       '
        .Fields("CountCA") = rsCA.Fields("CountCA")
        .Fields("AmountCA") = rsCA.Fields("AmountCA")
        .Fields("CountOA") = rsOA.Fields("CountOA")
        .Fields("AmountOA") = rsOA.Fields("AmountOA")
        
        .Fields("CountCheck") = rsCT.Fields("CountCheck")
        .Fields("AmountCheck") = rsCT.Fields("AmountCheck")
        'Giam % Mon
        .Fields("KarDiscountCount") = rsLineDiscount.Fields("CountLineDisc")
        .Fields("KarDiscountAmount") = rsLineDiscount.Fields("AmtLineDis")
        
        .Fields("CountCredit") = rsCC.Fields("CountCredit")
        .Fields("AmountCredit") = rsCC.Fields("AmountCredit")
        
        .Fields("CountGC") = rsGC.Fields("CountGC")
        .Fields("AmountGC") = rsGC.Fields("AmountGC")
        
        .Fields("CountROA") = rsROA.Fields("CountROA")
        .Fields("AmountROA") = rsROA.Fields("AmountROA")
        
        .Fields("CountOpen") = rsOpening.Fields("CountOpen")
        .Fields("AmountOpen") = rsOpening.Fields("AmountOpen")
        
        .Fields("Service_Charge_Count") = rsService_Charge.Fields("CountSer")
        .Fields("Service_Charge_Amt") = rsService_Charge.Fields("Amount")
        
        .Fields("VAT_Count") = rsVAT.Fields("VATCount")
        .Fields("VAT_Amt") = rsVAT.Fields("VATAmount")
        .Fields("Personal") = rsPersonal.Fields("Personal")
        
        .Fields("Adjustment1") = rsAdjustment1.Fields("Amount")
        .Fields("CountAdj1") = rsAdjustment1.Fields("CountAdj1")
        
        .Fields("Adjustment2") = rsAdjustment2.Fields("Amount")
        .Fields("CountAdj2") = rsAdjustment2.Fields("CountAdj2")
        
        .Fields("Adjustment3") = rsAdjustment3.Fields("Amount")
        .Fields("CountAdj3") = rsAdjustment3.Fields("CountAdj3")
        
        .Fields("Adjustment4") = rsAdjustment4.Fields("Amount")
        .Fields("CountAdj4") = rsAdjustment4.Fields("CountAdj4")
        
        .Fields("Adjustment5") = rsAdjustment5.Fields("Amount")
        .Fields("CountAdj5") = rsAdjustment5.Fields("CountAdj5")
        
        .Fields("Adjustment6") = rsAdjustment6.Fields("Amount")
        .Fields("CountAdj6") = rsAdjustment6.Fields("CountAdj6")
        
        .Fields("CountReceive") = CDbl("0" & rsReceive.Fields("CountReceive"))
        .Fields("AmountReceive") = CDbl("0" & rsReceive.Fields("AmountReceive"))
        
        .Fields("CountReserve") = CDbl("0" & rsReserve.Fields("CountReserve"))
        .Fields("AmountReserve") = CDbl("0" & rsReserve.Fields("AmountReserve"))
        
        .Update
        rsGroup.MoveNext
    Loop
End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & ""
End Sub

Public Sub AddFilter()
On Error GoTo Handle
Dim i As Integer

    With cboSaleFilter
        .Clear
        For i = 1 To 3
            .AddItem DescArr(60 + i)
        Next i
        .AddItem DescArr(86)
        .AddItem DescArr(107)
        .AddItem DescArr(108)
        .ListIndex = 0
    End With

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " AddFilter"
End Sub
Public Sub AddFilter80()
On Error GoTo Handle
Dim i As Integer

    With cbo80Filter
        .Clear
        For i = 1 To 3
            .AddItem DescArr(60 + i)
        Next i
        .AddItem DescArr(86)
        .AddItem DescArr(107)
        .AddItem DescArr(108)
        .ListIndex = 0
    End With

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " AddFilter"
End Sub
''''''''''''''''''''''''''''''''
Public Sub AddSort()
On Error GoTo Handle
Dim i As Integer
    With cboSalesort
        .Clear
        For i = 1 To 4
            .AddItem DescArr(63 + i)
        Next i
        .ListIndex = 0
    End With

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " AddSort"
End Sub
Public Sub AddSort80()
On Error GoTo Handle
Dim i As Integer
    With cbo80Sort
        .Clear
        For i = 1 To 4
            .AddItem DescArr(63 + i)
        Next i
        .ListIndex = 0
    End With

Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " AddSort"
End Sub


Public Sub CheckRight()
On Error GoTo Handle
    'Dim rsus As New ADODB.Recordset
        
        With MyRight
            rsuser.MoveFirst
            Do While Not rsuser.EOF
                If StrComp(rsuser.Fields("UserName"), userName, 1) = 0 Then
                    .FullRight = rsuser.Fields("UserRight")
                    .Sodoban = RightDeCode(Left(.FullRight, 64))
                    .Banhang = RightDeCode(Mid(.FullRight, 65, 64))
                    .Danhmuc = RightDeCode(Mid(.FullRight, 129, 64))
                    .Nhanvien = RightDeCode(Mid(.FullRight, 193, 64))
                    .Caidathethong = RightDeCode(Mid(.FullRight, 257, 64))
                    .Caidatdanhmuc = RightDeCode(Mid(.FullRight, 321, 64))
                    .Baocao = RightDeCode(Mid(.FullRight, 385, 64))
                    .kho = RightDeCode(Mid(.FullRight, 449, 64))
                    .Thuchi = RightDeCode(Mid(.FullRight, 513, 64))
                    .Suaten = RightDeCode(Mid(.FullRight, 577, 16))
                    Exit Do
                End If
                rsuser.MoveNext
            Loop
            If Mid(.Nhanvien, 1, 1) = 0 Then
                  cmdCashier.Enabled = False
            Else: cmdCashier.Enabled = True
            End If
            If Mid(.Nhanvien, 2, 1) = 0 Then
                  cmdAddCashier.Enabled = False
            Else: cmdAddCashier.Enabled = True
            End If
            If Mid(.Nhanvien, 3, 1) = 0 Then
                  cmdChangpass.Enabled = False
            Else: cmdChangpass.Enabled = True
            End If
            If Mid(.Nhanvien, 4, 1) = 0 Then
                  cmdChangeID.Enabled = False
            Else: cmdChangeID.Enabled = True
            End If
            If Mid(.Nhanvien, 5, 1) = 0 Then
                  cmdRightSelection.Enabled = False
            Else: cmdRightSelection.Enabled = True
            End If
            If Mid(.Nhanvien, 6, 1) = 0 Then
                  cmddept.Enabled = False
            Else: cmddept.Enabled = True
            End If
            If Mid(.Nhanvien, 7, 1) = 0 Then
                  cmdEmployee.Enabled = False
            Else: cmdEmployee.Enabled = True
            End If
            If Mid(.Nhanvien, 8, 1) = 0 Then
                  cmdShift.Enabled = False
            Else: cmdShift.Enabled = True
            End If
            
            If Mid(.Nhanvien, 9, 1) = 0 Then
                  cmdJobCode.Enabled = False
            Else: cmdJobCode.Enabled = True
            End If
            If Mid(.Nhanvien, 10, 1) = 0 Then
                  cmdSalary.Enabled = False
            Else: cmdSalary.Enabled = True
            End If
            
            If Mid(.Nhanvien, 12, 1) = 0 Then
                  cmdInvoiceList.Enabled = False
            Else: cmdInvoiceList.Enabled = True
            End If
            If Mid(.Nhanvien, 13, 1) = 0 Then
                  cmdInvoiceHoldList.Enabled = False
            Else: cmdInvoiceHoldList.Enabled = True
            End If
            
            If Mid(.Caidathethong, 1, 1) = 0 Then
                  cmdSetup.Enabled = False
            Else: cmdSetup.Enabled = True
            End If
            If Mid(.Caidathethong, 2, 1) = 0 Then
                  cmdSystemFlag.Enabled = False
            Else: cmdSystemFlag.Enabled = True
            End If
            If Mid(.Caidathethong, 3, 1) = 0 Then
                  cmdInvoiceSetup.Enabled = False
            Else: cmdInvoiceSetup.Enabled = True
            End If
             If Mid(.Caidathethong, 4, 1) = 0 Then
                  cmdTimerSetup.Enabled = False
            Else: cmdTimerSetup.Enabled = True
            End If
            If Mid(.Caidathethong, 5, 1) = 0 Then
                  cmdSetupPrint.Enabled = False
            Else: cmdSetupPrint.Enabled = True
            End If
            If Mid(.Caidathethong, 6, 1) = 0 Then
                  cmdAdjustment.Enabled = False
            Else: cmdAdjustment.Enabled = True
            End If
            If Mid(.Caidathethong, 7, 1) = 0 Then
                  cmdPriceRate.Enabled = False
            Else: cmdPriceRate.Enabled = True
            End If
            If Mid(.Caidathethong, 8, 1) = 0 Then
                  cmdLayout.Enabled = False
            Else: cmdLayout.Enabled = True
            End If
            If Mid(.Caidathethong, 9, 1) = 0 Then
                  cmdTaxRate.Enabled = False
            Else: cmdTaxRate.Enabled = True
            End If
            If Mid(.Caidathethong, 10, 1) = 0 Then
                  cmdConnect.Enabled = False
            Else: cmdConnect.Enabled = True
            End If
            If Mid(.Caidathethong, 11, 1) = 0 Then
                  cmdVAT.Visible = False
            Else: cmdVAT.Visible = True
            End If
            If Mid(.Caidatdanhmuc, 1, 1) = 0 Then
                  cmdAdministrative.Enabled = False
            Else: cmdAdministrative.Enabled = True
            End If
            If Mid(.Caidatdanhmuc, 2, 1) = 0 Then
                  cmdMaingroup.Enabled = False
            Else: cmdMaingroup.Enabled = True
            End If
            If Mid(.Caidatdanhmuc, 3, 1) = 0 Then
                  cmdGroup.Enabled = False
            Else: cmdGroup.Enabled = True
            End If
            If Mid(.Caidatdanhmuc, 4, 1) = 0 Then
                  cmdSKU.Enabled = False
            Else: cmdSKU.Enabled = True
            End If
           
            If Mid(.Caidatdanhmuc, 6, 1) = 0 Then
                  cmdCustomer.Enabled = False
            Else: cmdCustomer.Enabled = True
            End If
            If Mid(.Caidatdanhmuc, 7, 1) = 0 Then
                  cmdVendor.Enabled = False
            Else: cmdVendor.Enabled = True
            End If
           
            
            If Mid(.Caidatdanhmuc, 12, 1) = 0 Then
                  cmdMedia.Enabled = False
            Else: cmdMedia.Enabled = True
            End If
            If Mid(.Caidatdanhmuc, 13, 1) = 0 Then
                  cmdGiftCard.Enabled = False
            Else: cmdGiftCard.Enabled = True
            End If
            If Mid(.Caidatdanhmuc, 14, 1) = 0 Then
                  cmdInstruction.Enabled = False
            Else: cmdInstruction.Enabled = True
            End If
            If Mid(.Baocao, 1, 1) = 0 Then
                  cmdReport.Enabled = False
            Else: cmdReport.Enabled = True
            End If
           
'            If Mid(.Baocao, 8, 1) = 0 Then
'                  cmdBangke.Enabled = False
'            Else: cmdBangke.Enabled = True
'            End If
'            If Mid(.Baocao, 9, 1) = 0 Then
'                  cmdBangke.Enabled = False
'            Else: cmdBangke.Enabled = True
'            End If
'            If Mid(.Baocao, 10, 1) = 0 Then
'                  cmdBangke.Enabled = False
'            Else: cmdBangke.Enabled = True
'            End If
'            If Mid(.Baocao, 11, 1) = 0 Then
'                  cmdBangke.Enabled = False
'            Else: cmdBangke.Enabled = True
'            End If
            
            If Mid(.Baocao, 11, 1) = 0 Then
                  cmdSaleReport.Enabled = False
                  cmdReport80.Enabled = False
            Else
                cmdSaleReport.Enabled = True
                cmdReport80.Enabled = True
            End If
            
            If Mid(.Baocao, 12, 1) = 0 Then
                  cmdBaocaotonghop.Enabled = False
                  cmdGeneral80.Enabled = False
            Else
                cmdBaocaotonghop.Enabled = True
                cmdGeneral80.Enabled = True
            End If
            If Mid(.Baocao, 13, 1) = 0 Then
                  cmdReportGroup.Enabled = False
            Else: cmdReportGroup.Enabled = True
            End If
            If Mid(.Baocao, 14, 1) = 0 Then
                  cmdBaocaochitiet.Enabled = False
                  cmdDetail80.Enabled = False
            Else
                cmdBaocaochitiet.Enabled = True
                cmdDetail80.Enabled = True
            End If
            If Mid(.Baocao, 15, 1) = 0 Then
                  cmdTongbill.Enabled = False
                  cmdBill80.Enabled = False
            Else
                cmdTongbill.Enabled = True
                cmdBill80.Enabled = True
            End If
            If Mid(.Baocao, 16, 1) = 0 Then
                  cmdTable.Enabled = False
                  
            Else
                cmdTable.Enabled = True
                
            End If
            If Mid(.Baocao, 17, 1) = 0 Then
                  cmdCashierRepot.Enabled = False
                  cmdCashier80.Enabled = False
            Else
                cmdCashierRepot.Enabled = True
                cmdCashier80.Enabled = True
            End If
            If Mid(.Baocao, 18, 1) = 0 Then
               
                cmdBanchuathu80.Enabled = False
                cmdNotPaymented.Enabled = False
            Else
                cmdNotPaymented.Enabled = True
                cmdBanchuathu80.Enabled = True
            End If
            If Mid(.Baocao, 19, 1) = 0 Then
                  cmdDeletedItems.Enabled = False
            Else: cmdDeletedItems.Enabled = True
            End If

            If Mid(.Baocao, 20, 1) = 0 Then
                  cmdHourly.Enabled = False
            Else: cmdHourly.Enabled = True
            End If
            
            If Mid(.Baocao, 21, 1) = 0 Then
                  cmdKitchen_List.Enabled = False
            Else: cmdKitchen_List.Enabled = True
            End If
            
            If Mid(.Baocao, 22, 1) = 0 Then
                  cmdKit_General.Enabled = False
            Else: cmdKit_General.Enabled = True
            End If
            
            
            
'            If Mid(.Baocao, 28, 1) = 0 Then
'                  cmdProfit.Enabled = False
'            Else: cmdProfit.Enabled = True
'            End If
            
'            If Mid(.Baocao, 29, 1) = 0 Then
'                  cmdLevelMaterial.Enabled = False
'            Else: cmdLevelMaterial.Enabled = True
'            End If
            
'            If Mid(.Baocao, 30, 1) = 0 Then
'                  cmdAnalyse.Enabled = False
'            Else: cmdAnalyse.Enabled = True
'            End If
            
            
            
            If Mid(.Suaten, 1, 1) = 0 Then
                  cmdEditName.Enabled = False
            Else: cmdEditName.Enabled = True
            End If
            
            If Mid(.Suaten, 2, 1) = 0 Then
                  cmdLocationName.Enabled = False
            Else: cmdLocationName.Enabled = True
            End If
            If Mid(.Suaten, 3, 1) = 0 Then
                  cmdPrinterName.Enabled = False
            Else: cmdPrinterName.Enabled = True
            End If

        End With

   ' CloseRecordset res
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " CheckRight"
End Sub


Public Sub addPrinter()
On Error GoTo Handle
    Dim rsPrint As New ADODB.Recordset
    Set rsPrint = Open_Table(cnData, "Friendly_Printers")
    cboKit_Printer.Clear
    With rsPrint
        cboKit_Printer.AddItem "TÊt c¶"
        Do While Not .EOF
            cboKit_Printer.AddItem .Fields("PrinterName")
            .MoveNext
        Loop
        cboKit_Printer.ListIndex = 0
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " addPrinter"
End Sub

Public Sub AddKit_Filter()
On Error GoTo Handle
    Dim i As Integer
    With cboKit_Filter
        .Clear
            .AddItem "Taát caû"
        For i = 1 To 2
            .AddItem DescArr(81 + i)
        Next i
        .ListIndex = 0
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " AddKit_Filter"
End Sub
Public Sub AddKit_Sort()
On Error GoTo Handle
    Dim i As Integer
    With cboKit_Sort
        .Clear
        For i = 1 To 3
            .AddItem DescArr(83 + i)
        Next i
        .ListIndex = 0
    End With
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " AddKit_Sort"
End Sub


Public Sub AddTable_Nhap_Xuat_Ton_Temp()
On Error GoTo Handle
    Dim cat As New ADOX.Catalog
    Dim i As Integer
    Dim bln As Boolean
    bln = False
    cat.ActiveConnection = myProvider
        For i = 1 To cat.Tables.count - 1
            If cat.Tables(i).name = "Inventory_Calcu_Temp" Then
                bln = True
            End If
        Next
        If bln = False Then
            cat.Tables.Append CreateTable_Temp
        End If
       
    Exit Sub
Handle:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & "AddTable_Nhap_Xuat_Ton_Table"
End Sub


