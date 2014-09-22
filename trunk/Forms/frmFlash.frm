VERSION 5.00
Begin VB.Form frmFlash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4530
   ClientLeft      =   1545
   ClientTop       =   1050
   ClientWidth     =   6285
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   ".VnArial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFlash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "MDIMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4425
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   6090
      Begin VB.Timer TimerCtrl 
         Left            =   5760
         Top             =   270
      End
      Begin VB.Label lblCustomer 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   2850
         Width           =   5025
      End
      Begin VB.Label lblCompanyProduct 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "VNI-Algerian"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   660
         TabIndex        =   9
         Top             =   480
         Width           =   5055
      End
      Begin VB.Label lblLastUser 
         Caption         =   "T¸c gi¶:"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   3090
         Width           =   5355
      End
      Begin VB.Line Line2 
         X1              =   60
         X2              =   5910
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   720
         X2              =   5880
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label lblAddress 
         Caption         =   "Address"
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
         Height          =   615
         Left            =   510
         TabIndex        =   7
         Top             =   1230
         Width           =   5475
      End
      Begin VB.Label lblCompanyInit 
         Alignment       =   2  'Center
         Caption         =   "Company Init"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   390
         TabIndex        =   6
         Top             =   165
         Width           =   5115
      End
      Begin VB.Label lblWarningMes 
         Caption         =   "Warning Message"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   3840
         Width           =   5085
      End
      Begin VB.Label lblUserIDKey 
         Caption         =   "Vò Kh¾c CËn"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   3450
         Width           =   4065
      End
      Begin VB.Label lblWarning 
         Alignment       =   1  'Right Justify
         Caption         =   "Warning"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   3840
         Width           =   705
      End
      Begin VB.Label lblProductName 
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   480
         TabIndex        =   3
         Top             =   1890
         Width           =   5580
      End
      Begin VB.Label lblLicenseTo 
         Caption         =   "LicenseTo"
         BeginProperty Font 
            Name            =   ".VnArial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   2520
         Width           =   5355
      End
   End
End
Attribute VB_Name = "frmFlash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTime As Double
Dim LngFileExist As Boolean
Dim Code_Used As String
Dim isActive As Boolean

Private Sub Form_Activate()
On Error GoTo errHdl
If isActive Then Exit Sub
'Detect previous threading
    If App.PrevInstance Then
        Unload Me
        MsgBox "Ch­¬ng tr×nh ®· ch¹y råi!" & vbCrLf & _
            vbCrLf & "Program was running!", vbExclamation
        End
    End If
'--------------------------------------------------------
    If LngFileExist = False Then
        Unload Me
        End
    End If
'First Run
    If GetSettingStr("SYSTEM", "Start Working", "", myIniFile) = "" Then
        SaveSettingStr "SYSTEM", "Start Working", Year(Date), myIniFile
    End If
    
    DigitGroupMark = GetSettingStr("NUMBER", "Digit Group Symbol", ".", myIniFile)
    DecimalMark = GetSettingStr("NUMBER", "Decimal Symbol", ",", myIniFile)
    DigitsGroup = CInt(GetSettingStr("NUMBER", "Digit Group", "3", myIniFile))
    DecimalQtyNumber = CInt(GetSettingStr("NUMBER", "Quantity Decimal", "2", myIniFile))
    DecimalAmtNumber = CInt(GetSettingStr("NUMBER", "Amount Decimal", "2", myIniFile))
    formatNum = GetSettingStr("NUMBER", "Formatnum", formatNum, myIniFile)
    CurrencySymbol = GetSettingStr("NUMBER", "CurrencyBymbol", CurrencySymbol, myIniFile)
    Store_ID = GetSettingStr("SYSTEM", "Station", "01", myIniFile)
    
    ServerName = GetSettingStr("SYSTEM", "ServerName", "", myIniFile)
    DataBaseName = GetSettingStr("SYSTEM", "DatabaseName", "", myIniFile)
    UserLog = GetSettingStr("SYSTEM", "UserLogin", "sa", myIniFile)
    DB_Password = GetSettingStr("SYSTEM", "Password", "", myIniFile)
    DB_Password = En_Decryption.MalgoDecrypt(DB_Password, 10)
    
    'load infor backup
    
    BK_ServerName = GetSettingStr("SYSTEM", "Backup_ServerName", "", myIniFile)
    BK_DataBaseName = GetSettingStr("SYSTEM", "Backup_DatabaseName", "", myIniFile)
    BK_UserLog = GetSettingStr("SYSTEM", "Backup_UserLogin", "sa", myIniFile)
    BK_DB_Password = GetSettingStr("SYSTEM", "Backup_Password", "", myIniFile)
    BK_DB_Password = En_Decryption.MalgoDecrypt(BK_DB_Password, 10)
    
    'Canh  le
    TopAlign = GetSettingStr("ALIGN", "Top", "0", myIniFile)
    BottomAlign = GetSettingStr("ALIGN", "Bottom", "0", myIniFile)
    LeftAlign = GetSettingStr("ALIGN", "Left", "0", myIniFile)
    RightAlign = GetSettingStr("ALIGN", "Right", "0", myIniFile)
    
    'LÊy lo¹i m¸y in
     
     CurFont = GetSettingStr("SYSTEM", "Font", "", myIniFile)
     ColorFont = GetSettingStr("SYSTEM", "FontColor", "", myIniFile)
     ShapeColor = GetSettingStr("SYSTEM", "ShapeColor", "", myIniFile)
     bkColor = GetSettingStr("SYSTEM", "bkColor", "", myIniFile)
     
     ReceiptType = GetSettingStr("PRINTER", "Receipt_Type", "80", myIniFile)
    OrderType = GetSettingStr("PRINTER", "Order_Type", "80", myIniFile)
    
    Sort_By = GetSettingStr("SORT", "Sort_by", "Dept_ID,ItemNum", myIniFile)
    
    'check_Date_Lock
    Call check_Date_Lock
    ''''''End
    
    Date_Open = En_Decryption.MalgoDecrypt(GetSettingStr("SYSTEM", "DateOpen", "", myIniFile), 5)
    
    If Date_Open = "" Then
        SaveSettingStr "SYSTEM", "DateOpen", En_Decryption.MalgoEncrypt(gfCONVERT_DATE_TO_STRING(Date), 5), myIniFile
        Date_Open = Format(Date, "dd/MM/yyyy")
    Else
        Date_Open = gfCONVERT_STRING_TO_DATE(Date_Open)
    End If
    lblUserIDKey.Visible = True
    TimerCtrl.Interval = 500

    If ServerName = "" Then
        frmConnect_Data.Show vbModal
    End If
    isActive = True
   
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Form_Activate"
End Sub

Private Sub Form_Load()
On Error GoTo errHdl

    Dim vFlag As Boolean
    Dim tmpLngFile As String
    Const STARTANTIDEBUGGER As Long = 0
    CommandStr = Command
    
     'Kiem tra du lieu luu ngay nao
     If Hour(Format(Now, "HH:mm:ss")) < 5 Then
        frmDate.Show vbModal
    Else
        DateDefault = Format(Year(Date), "0000") & Format(Month(Date), "00") & Format(Day(Date), "00")
    End If
'-----------------------------------------------------------
'Initial File
    myIniFile = App.Path & "\" & App.EXEName & ".ini"
    tmpLngFile = GetSettingStr("SYSTEM", "Language", "", myIniFile)
'    WorkingFolder = GetSettingStr("Default Site", "Default Site", True, myIniFile)
    BackupFolder = GetSettingStr("SYSTEM", "Backup Site", True, myIniFile)
    'WorkingFolder = ReportFolder & Format(Month(Date), "00") & Format(Year(Date), "00")
'LanguageLoad:
    If tmpLngFile = "" Or Dir(App.Path & tmpLngFile) = "" Then
        frmLanguageSelection.Show vbModal, Me
        If LngFile = "" Then
            LngFileExist = False
            Exit Sub
        Else
            LngFileExist = True
        End If
        SaveSettingStr "SYSTEM", "Language", Replace(LngFile, App.Path, ""), myIniFile
    Else
        LngFileExist = True
        LngFile = App.Path & tmpLngFile
    End If
    LngFolder = RemoveExtFile(LngFile)
    Call LoadFont(CurFont, LngFile)
    If InStr(LngFile, "Vietnamese") Then
        vFlag = True
    Else
        vFlag = False
    End If
    sTime = Timer
    
    With lblCompanyInit.Font
        .name = ".VnArialH"
        .Size = 8
        .Bold = True
        .Italic = False
        .Underline = False
    End With

lblCompanyProduct.ForeColor = vbRed
    With lblProductName.Font
        .name = ".VnArialH"
        .Size = 20
        .Bold = True
        .Italic = True
    End With
    With lblCompanyProduct.Font
        .name = "VNI-Algerian"
        .Size = 16
        .Bold = True
    End With
    lblProductName.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
    If vFlag Then
        lblCompanyInit.Caption = "C«ng ty TNHH Th­¬ng m¹i - DÞch vô"
        lblCompanyProduct.Caption = "phuùc thaïnh vinh"
        lblLicenseTo.Caption = "TÊt c¶ mäi quyÒn"
        lblCustomer.Caption = "            ®­îc b¶o l­u"
        lblAddress.Caption = "§Þa chØ: 565/6 B×nh Thíi, P.10, Q.11,Tp.HCM" & _
        Chr(13) & "§iÖn tho¹i: 08-3.8867.869   -  0918.655.887"
        lblWarning.Caption = " "
        lblLastUser = "T¸c gi¶:"
        lblUserIDKey.Caption = "Vò Kh¾c CËn"
        lblWarningMes.Caption = App.LegalCopyright
    Else
        lblCompanyInit.Caption = "Service and Trading Co.,Ltd"
        lblCompanyProduct.Caption = "phuc thanh vinh"
        lblLicenseTo.Caption = ""
        lblCustomer.Caption = " All Right Reserved"
        lblAddress.Caption = "Address : 565/6 Binh Thoi Street,ward 10,Dist 11,HCM City" & _
        Chr(13) & "Telephone: 08-8867.869   -   0918.655.887"
        lblWarning.Caption = " "
        lblLastUser = "Author:"
        lblUserIDKey.Caption = "Vu Khac Can"
        lblWarningMes.Caption = App.LegalCopyright
    End If
    
   
    
'    frmLogin.Show vbModal
     ProcessID = MachineID.Get_ProcessID
    Mac_ID = Right("3C7D9F5JG0S1" & GetMACs_IfTable, 12)
    Code_Used = UCase(ProcessID) ' & Mac_ID)
    
    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - Form_Load"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If BK_ServerName = "" Then
        SaveSettingStr "SYSTEM", "Backup_ServerName", ServerName, myIniFile
        SaveSettingStr "SYSTEM", "Backup_DatabaseName", DataBaseName, myIniFile
        SaveSettingStr "SYSTEM", "Backup_Password", En_Decryption.MalgoEncrypt(DB_Password, 10), myIniFile
        SaveSettingStr "SYSTEM", "Backup_UserLogin", UserLog, myIniFile
    End If
End Sub

Private Sub TimerCtrl_Timer()
On Error GoTo errHdl
Dim date_Expired As String
Dim hFile As Double
Dim tmpStr, LogFile As String
If Len(Date_Open) = 8 Then Date_Open = gfCONVERT_STRING_TO_DATE(Date_Open)
LogFile = "C:\Windows\System32\KernelSys.sys"
    If Timer - sTime > 2 Then
        Unload Me
        If Dir(LogFile, vbDirectory) <> "" Then
            hFile = FreeFile
            Open LogFile For Input As #hFile
            Do While Not EOF(hFile)
                DoEvents
                Line Input #hFile, tmpStr
            Loop
            Close #hFile
            If UCase(En_Decryption.MalgoDecrypt(tmpStr, 15)) = UCase(Code_Used) Then
                Unload Me
                frmLogin.Show vbModal
            Else
                Kill LogFile
                frmDemo.Show vbModal
            End If
        ElseIf Dir("C:\Windows\System32\sysTrial.dll", vbDirectory) <> "" Then
        Dim Date_issue As Integer
        Dim date_lock As String
            hFile = FreeFile
                Open "C:\Windows\System32\systrial.dll" For Input As #hFile
                Do While Not EOF(hFile)
                    DoEvents
                    Line Input #hFile, tmpStr
                Loop
                Close #hFile
                date_lock = En_Decryption.MalgoDecrypt(tmpStr, 5)
                Date_issue = (Val(Left(date_lock, 4)) - Year(Date)) * 365 + (Val(Mid(date_lock, 5, 2)) - Val(Month(Date))) * 30 + Val(Right(date_lock, 2)) - Day(Date)
                
                If gfCONVERT_DATE_TO_STRING(Date) < gfCONVERT_DATE_TO_STRING(Date_Open) Then
                    MsgBox "B¹n kh«ng thÓ lïi ngµy hÖ thèng ®Ó ®­îc dïng b¶n trial !", vbCritical
                Else
                        If Date_issue <= 30 Then
                            If Date_issue <= 0 Then
                                MsgBox "B¹n ®· hÕt h¹n sö dông phÇn mÒm!", vbInformation
                                Exit Sub
                            Else
                                SaveSettingStr "SYSTEM", "DateOpen", En_Decryption.MalgoEncrypt(gfCONVERT_DATE_TO_STRING(Date), 5), myIniFile
                                MsgBox "B¹n cßn " & Date_issue & " ngµy sö dông, vui lßng liªn hÖ 0918.655.887 ®Ó ®­îc gia h¹n"
                                Unload Me
                                 frmLogin.Show vbModal

                            End If
                        Else
                                SaveSettingStr "SYSTEM", "DateOpen", En_Decryption.MalgoEncrypt(gfCONVERT_DATE_TO_STRING(Date), 5), myIniFile
                                 frmLogin.Show vbModal
                        End If
            End If
        Else

            If gfCONVERT_DATE_TO_STRING(Date) < gfCONVERT_DATE_TO_STRING(Date_Open) Then
                MsgBox "B¹n kh«ng thÓ lïi ngµy hÖ thèng ®Ó ®­îc dïng b¶n trial !", vbInformation
            Else
                date_Expired = get_Trial_Date
                If date_Expired = "" Then
                    frmDemo.Show vbModal
                Else
                    If date_Expired >= gfCONVERT_DATE_TO_STRING(Date) Then
                        MsgBox "B¹n cßn " & (Val(Mid(date_Expired, 5, 2)) - Val(Month(Date))) * 30 + Val(Right(date_Expired, 2)) - Day(Date) + 1 & " ngµy dïng thö"
                         SaveSettingStr "SYSTEM", "DateOpen", En_Decryption.MalgoEncrypt(gfCONVERT_DATE_TO_STRING(Date), 5), myIniFile
                         Unload Me
                         frmLogin.Show vbModal
                    Else
                        MsgBox "B¶n dïng thö cña b¹n ®· hÕt thêi gian miÔn phÝ, vui lßng liªn hÖ: 0918.655.887 ®Ó ®­îc cÊp m· sö dông"
                        frmLicense.Show vbModal
                    End If
                End If
            End If
        End If

    End If

    Exit Sub
errHdl:
    MsgBox Err.Number & " : " & Err.Description & vbCrLf _
    & Me.name & " - TimerCtrl_Timer"
End Sub


Public Sub check_Date_Lock()
On Error GoTo Handle
Dim str As String
 If cnData.State = adStateClosed Then Set cnData = Get_Connection(ServerName, DataBaseName, UserLog, DB_Password)
    If Not Check_Table_exist("DateLock") Then
        str = "CREATE TABLE [dbo].[DateLock]([LifeTime] [nvarchar](30)  NULL)"
        cnData.Execute str
    End If
Exit Sub
Handle:
MsgBox Err.Number & Err.Description & Me.name & " check_Date_Lock"
End Sub
