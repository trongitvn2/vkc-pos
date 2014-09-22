Attribute VB_Name = "mdlVBKey"
Option Explicit

Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function OpenClipboard Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32.dll" () As Long
Private Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32.dll" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32.dll" () As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32.dll" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal hMem As Long) As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)


Private Enum KEYTYPES
    BACK_CHAR = 8
    ESCAPE_KEY = 27
    
    DOUBLE_CHAR = 1
    D_CHAR = 2
    BREVE_MARK = 3
    TONE_MARK = 4
    UN_MARK = 5
                 
End Enum


Private Enum VOWELS
    NONE_ = 1
    BREVE_ = 2
    TONE_ = 3
    TONE_BREVE_ = 4
End Enum

Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const CF_UNICODETEXT = 13
Private Const VK_BACK = &H8
Private Const VK_INSERT = &H2D
Private Const VK_SHIFT = &H10

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Const BEFORE_D As String = "B,C,D,G,H,K,L,M,N,P,Q,S,T,V,X"
Private Const CONSONANT As String = "C,E,I,M,N,P,T,U,Y"
Private Const MAX_WORD_LENGTH = 6
Private Const MAX_VOWEL_LENGTH = 3

Public uBuf As String
Public Buf As String
Public Bks As Integer
Public Keys As Integer
Public Vkey As Boolean
Private TOff As Boolean
Private LOff As Integer
Private Lw As Boolean
Public OStyle As Boolean
Private TArr() As Variant
Private UArr() As Variant
Private STRING_RESET As String
Private VK_BACK_SCAN As Long
Private VK_SHIFT_SCAN As Long
Private VK_INSERT_SCAN As Long
Private hMouseHook As Long
Private hKeyHook As Long
Private ProcessK() As Variant
Private Initialized As Boolean
Private IsEnd As Boolean
Public UCaseFirst As Boolean


'============================ BEGIN PROCESS ==============================

Public Function IsProcessKey(Ch As Long) As Boolean
    If Not Initialized Then InitApp
    Dim I As Long
    For I = LBound(ProcessK) To UBound(ProcessK)
        If ProcessK(I) = Ch Then
            IsProcessKey = True
            Exit Function
        End If
    Next I
    IsProcessKey = False
End Function

Public Sub InitApp()
    Initialized = True
    ProcessK = Array(16, 32, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 96, 97, 98, 99, 100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 186, 187, 188, 189, 190, 191, 219, 220, 221, 222)
    STRING_RESET = "`~!@#$%^&*()-_=+;:'\|,<.>/?1234567890 "
    TArr = Array("af", "as", "ar", "ax", "aj", "aa", "aaf", "aas", "aar", "aax", "aaj", "aw", "awf", "aws", "awr", "awx", "awj", "dd", "ef", "es", "er", "ex", "ej", "ee", "eef", "ees", "eer", "eex", "eej", "if", "is", "ir", "ix", "ij", "of", "os", "or", "ox", "oj", "oo", "oof", "oos", "oor", "oox", "ooj", "ow", "owf", "ows", "owr", "owx", "owj", "uf", "us", "ur", "ux", "uj", "uw", "uwf", "uws", "uwr", "uwx", "uwj", "yf", "ys", "yr", "yx", "yj", "AF", "AS", "AR", "AX", "AJ", "AA", "AAF", "AAS", "AAR", "AAX", "AAJ", "AW", "AWF", "AWS", "AWR", "AWX", "AWJ", "DD", "EF", "ES", "ER", "EX", "EJ", "EE", "EEF", "EES", "EER", "EEX", "EEJ", "IF", "IS", "IR", "IX", "IJ", "OF", "OS", "OR", "OX", "OJ", "OO", "OOF", "OOS", "OOR", "OOX", "OOJ", "OW", "OWF", "OWS", "OWR", "OWX", "OWJ", "UF", "US", "UR", "UX", "UJ", "UW", "UWF", "UWS", "UWR", "UWX", "UWJ", "YF", "YS", "YR", "YX", "YJ")
        
    UArr = Array( _
        ChrW$(&HE0), ChrW$(&HE1), ChrW$(&H1EA3), ChrW$(&HE3), ChrW$(&H1EA1), ChrW$(&HE2), ChrW$(&H1EA7), ChrW$(&H1EA5), _
        ChrW$(&H1EA9), ChrW$(&H1EAB), ChrW$(&H1EAD), ChrW$(&H103), ChrW$(&H1EB1), ChrW$(&H1EAF), ChrW$(&H1EB3), _
        ChrW$(&H1EB5), ChrW$(&H1EB7), ChrW$(&H111), ChrW$(&HE8), ChrW$(&HE9), ChrW$(&H1EBB), ChrW$(&H1EBD), _
        ChrW$(&H1EB9), ChrW$(&HEA), ChrW$(&H1EC1), ChrW$(&H1EBF), ChrW$(&H1EC3), ChrW$(&H1EC5), ChrW$(&H1EC7), _
        ChrW$(&HEC), ChrW$(&HED), ChrW$(&H1EC9), ChrW$(&H129), ChrW$(&H1ECB), ChrW$(&HF2), ChrW$(&HF3), ChrW$(&H1ECF), _
        ChrW$(&HF5), ChrW$(&H1ECD), ChrW$(&HF4), ChrW$(&H1ED3), ChrW$(&H1ED1), ChrW$(&H1ED5), ChrW$(&H1ED7), _
        ChrW$(&H1ED9), ChrW$(&H1A1), ChrW$(&H1EDD), ChrW$(&H1EDB), ChrW$(&H1EDF), ChrW$(&H1EE1), ChrW$(&H1EE3), _
        ChrW$(&HF9), ChrW$(&HFA), ChrW$(&H1EE7), ChrW$(&H169), ChrW$(&H1EE5), ChrW$(&H1B0), ChrW$(&H1EEB), _
        ChrW$(&H1EE9), ChrW$(&H1EED), ChrW$(&H1EEF), ChrW$(&H1EF1), ChrW$(&H1EF3), ChrW$(&HFD), ChrW$(&H1EF7), _
        ChrW$(&H1EF9), ChrW$(&H1EF5), ChrW$(&HC0), ChrW$(&HC1), ChrW$(&H1EA2), ChrW$(&HC3), ChrW$(&H1EA0), _
        ChrW$(&HC2), ChrW$(&H1EA6), ChrW$(&H1EA4), ChrW$(&H1EA8), ChrW$(&H1EAA), ChrW$(&H1EAC), ChrW$(&H102), _
        ChrW$(&H1EB0), ChrW$(&H1EAE), ChrW$(&H1EB2), ChrW$(&H1EB4), ChrW$(&H1EB6), ChrW$(&H110), ChrW$(&HC8), _
        ChrW$(&HC9), ChrW$(&H1EBA), ChrW$(&H1EBC), ChrW$(&H1EB8), ChrW$(&HCA), ChrW$(&H1EC0), ChrW$(&H1EBE), _
        ChrW$(&H1EC2), ChrW$(&H1EC4), ChrW$(&H1EC6), ChrW$(&HCC), ChrW$(&HCD), ChrW$(&H1EC8), ChrW$(&H128), _
        ChrW$(&H1ECA), ChrW$(&HD2), ChrW$(&HD3), ChrW$(&H1ECE), ChrW$(&HD5), ChrW$(&H1ECC), ChrW$(&HD4), _
        ChrW$(&H1ED2), ChrW$(&H1ED0), ChrW$(&H1ED4), ChrW$(&H1ED6), ChrW$(&H1ED8), ChrW$(&H1A0), ChrW$(&H1EDC), _
        ChrW$(&H1EDA), ChrW$(&H1EDE), ChrW$(&H1EE0), ChrW$(&H1EE2), ChrW$(&HD9), ChrW$(&HDA), ChrW$(&H1EE6), _
        ChrW$(&H168), ChrW$(&H1EE4), ChrW$(&H1AF), ChrW$(&H1EEA), ChrW$(&H1EE8), ChrW$(&H1EEC), ChrW$(&H1EEE), _
        ChrW$(&H1EF0), ChrW$(&H1EF2), ChrW$(&HDD), ChrW$(&H1EF6), ChrW$(&H1EF8), ChrW$(&H1EF4))
    VK_BACK_SCAN = MapVirtualKey(VK_BACK, 0)
    VK_INSERT_SCAN = MapVirtualKey(VK_INSERT, 0)
    VK_SHIFT_SCAN = MapVirtualKey(VK_SHIFT, 0)
    Vkey = True
End Sub



Private Sub PasteCmd()
    Dim Sh As Integer
    Sh = GetKeyState(VK_SHIFT) And &H80
    
    If Sh Then
        keybd_event VK_SHIFT, VK_SHIFT_SCAN, KEYEVENTF_KEYUP, 0
        
        keybd_event VK_SHIFT, VK_SHIFT_SCAN, 0, 0
        keybd_event VK_INSERT, VK_INSERT_SCAN, KEYEVENTF_EXTENDEDKEY, 0
        keybd_event VK_INSERT, VK_INSERT_SCAN, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        keybd_event VK_SHIFT, VK_SHIFT_SCAN, KEYEVENTF_KEYUP, 0
        
        keybd_event VK_SHIFT, VK_SHIFT_SCAN, 0, 0
    Else
        keybd_event VK_SHIFT, VK_SHIFT_SCAN, 0, 0
        keybd_event VK_INSERT, VK_INSERT_SCAN, KEYEVENTF_EXTENDEDKEY, 0
        keybd_event VK_INSERT, VK_INSERT_SCAN, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        keybd_event VK_SHIFT, VK_SHIFT_SCAN, KEYEVENTF_KEYUP, 0
    End If
End Sub

Public Sub pushBuf(sBuf As String)
    SetData sBuf
    PasteCmd
End Sub


Public Sub clrBuffer()
    Buf = ""
    Keys = 0
    uBuf = ""
    Bks = 0
    LOff = 0
    TOff = False
End Sub



Public Sub PutC(Ch As String)

    Keys = Keys + 1
    
    If Keys >= 3 Then
        If Right$(Buf, 2) = ". " And Mid$(Buf, Keys - 2) <> "." Then
            IsEnd = True
        Else
            IsEnd = False
        End If
    ElseIf Keys = 2 Then
        If Right$(Buf, 2) = ". " Then
            IsEnd = True
        Else
            IsEnd = False
        End If
    End If
    
    
    If IsEnd And UCaseFirst Then
        uBuf = Right$(Buf, 1)
        Bks = Len(uBuf)
        Buf = Buf & UCase$(Ch)
        uBuf = Right$(Buf, 2)
        Exit Sub
    End If
        
    Buf = Buf & Ch

    If Vkey Then
        If Keys > 1 Then
            Dim S1 As String, S2 As String
            S1 = Mid$(Buf, Keys - 1, 1)
            S2 = Mid$(Buf, Keys, 1)
            If (UCase$(Left$(FromUni(S1), 1)) = "O" And UCase$(Left$(FromUni(S2), 1)) = "A") Or (UCase$(Left$(FromUni(S1), 1)) = "O" And UCase$(Left$(FromUni(S2), 1)) = "E") Or (UCase$(Left$(FromUni(S1), 1)) = "U" And UCase$(Left$(FromUni(S2), 1)) = "Y") Then
                If (VowelOf(S1) = TONE_) And (VowelOf(S2) = NONE_) Then
                    If OStyle Then Exit Sub
                    If Not OStyle Then
                        Mid$(Buf, Keys - 1, 1) = Left$(FromUni(S1), 1)
                        Mid$(Buf, Keys, 1) = ToUni(S2 & Right$(FromUni(S1), 1))
                        uBuf = Mid$(Buf, Keys - 1, Keys - (Keys - 1) + 1)
                        Bks = 1
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    If Vkey Then
        If InStr(1, STRING_RESET, Ch, vbTextCompare) > 0 Then TOff = False
    End If
End Sub

Private Function KeyOf(kCode As Long) As KEYTYPES
    
    If kCode = VK_BACK Then
        KeyOf = BACK_CHAR
        Exit Function
    ElseIf kCode = 27 Then
        KeyOf = ESCAPE_KEY
        Exit Function

    ElseIf ((kCode = 65) Or (kCode = 97) Or (kCode = 69) Or (kCode = 101) Or (kCode = 111) Or (kCode = 79)) Then
        KeyOf = DOUBLE_CHAR
        Exit Function
    ElseIf ((kCode = 100) Or (kCode = 68)) Then
        KeyOf = D_CHAR
        Exit Function
    ElseIf ((kCode = 119) Or (kCode = 87)) Then
        KeyOf = BREVE_MARK
        Exit Function
    ElseIf ((kCode = 70) Or (kCode = 83) Or (kCode = 82) Or (kCode = 88) Or (kCode = 74) Or (kCode = 102) Or (kCode = 115) Or (kCode = 114) Or (kCode = 120) Or (kCode = 106)) Then
        KeyOf = TONE_MARK
        Exit Function
    ElseIf ((kCode = 90) Or (kCode = 122)) Then
        KeyOf = UN_MARK
        Exit Function
    End If
    KeyOf = 0
End Function

Public Sub PrK(Ch As Long)
    If Not Initialized Then InitApp
    If KeyOf(Ch) = BACK_CHAR Then
        PBack Ch
    ElseIf KeyOf(Ch) = ESCAPE_KEY Then
        P_E_Key Ch
    ElseIf KeyOf(Ch) = DOUBLE_CHAR Then
        PDblChar Chr$(Ch)
    ElseIf KeyOf(Ch) = D_CHAR Then
        PDChar Chr$(Ch)
    ElseIf KeyOf(Ch) = BREVE_MARK Then
        PBMark Chr$(Ch)
    ElseIf KeyOf(Ch) = TONE_MARK Then
        PTMark Chr$(Ch)
    ElseIf KeyOf(Ch) = UN_MARK Then
        PUnMark Chr$(Ch)
    ElseIf InStr(1, CONSONANT, Chr$(Ch), vbTextCompare) > 0 Then
        PLast Chr$(Ch)
    Else
        PutC Chr$(Ch)
    End If

End Sub

Private Sub P_E_Key(Ch As Long)
    'Chua su dung
    clrBuffer
End Sub


Private Function GetLastWord(S As String) As Long
    Dim Counts As Byte, I As Long
    If S = "" Then GetLastWord = 0
    
    For I = Len(S) To 1 Step -1
        Counts = Counts + 1
        If ((Counts > MAX_WORD_LENGTH) Or (InStr(1, STRING_RESET, Mid$(S, I, 1), vbTextCompare) > 0)) Then
            Exit For
        End If
    Next I
    GetLastWord = IIf(I > 0, I, 1)
End Function


Private Sub PBack(Ch As Long)
    If Ch <> VK_BACK Then Exit Sub
    If Keys > 0 Then
    
        Keys = Keys - 1
        
        
        If Keys < LOff Then
            LOff = 0
            TOff = False
        End If
        
        Buf = Left$(Buf, Keys)
        If Vkey Then
            If Keys > 1 Then
                Dim S1 As String, S2 As String
                S1 = Mid$(Buf, Keys - 1, 1)
                S2 = Mid$(Buf, Keys, 1)
                If (UCase$(Left$(FromUni(S1), 1)) = "O" And UCase$(Left$(FromUni(S2), 1)) = "A") Or (UCase$(Left$(FromUni(S1), 1)) = "O" And UCase$(Left$(FromUni(S2), 1)) = "E") Then
                    If (VowelOf(S1) = NONE_) And (VowelOf(S2) = TONE_) Then
                        If OStyle Then
                            uBuf = Mid$(Buf, Keys - 1, Keys - (Keys - 1) + 1)
                            Bks = Len(uBuf)
                            Mid$(Buf, Keys, 1) = Left$(FromUni(S2), 1)
                            Mid$(Buf, Keys - 1, 1) = ToUni(S1 & Right$(FromUni(S2), 1))
                            uBuf = Mid$(Buf, Keys - 1, Keys - (Keys - 1) + 1)
                            Bks = Bks + 1
                            Exit Sub
                        End If
                    End If
                ElseIf (UCase$(Left$(FromUni(S1), 1)) = "U" And UCase$(Left$(FromUni(S2), 1)) = "Y") Then
                    If (VowelOf(S1) = NONE_) And (VowelOf(S2) = TONE_) Then
                        If OStyle Then
                            uBuf = Mid$(Buf, Keys - 1, Keys - (Keys - 1) + 1)
                            Bks = Len(uBuf)
                            Mid$(Buf, Keys, 1) = Left$(FromUni(S2), 1)
                            Mid$(Buf, Keys - 1, 1) = ToUni(S1 & Right$(FromUni(S2), 1))
                            uBuf = Mid$(Buf, Keys - 1, Keys - (Keys - 1) + 1)
                            Bks = Bks + 1
                            Exit Sub
                        End If
                    End If

                End If
            End If
        End If
    Else
        Keys = 0
        Buf = ""
        uBuf = ""
        clrBuffer
    End If
End Sub




Private Sub PUnMark(Ch As String)
    If ((TOff = True) Or (Keys <= 0)) Then
        PutC Ch
        Exit Sub
    End If
    
    
    Dim F As Integer, L As Integer, R As Integer, Fn As Boolean
    
    F = GetLastWord(Buf)
    If LOff > F Then F = LOff
    L = Keys
    
    Fn = False
    Do While F <= L
        If VowelOf(Mid$(Buf, F, 1)) <> 0 Then
            Fn = True
            Exit Do
        End If
        F = F + 1
    Loop
    
    If Not Fn Then
        PutC Ch
        Exit Sub
    End If
    
    Fn = False
    Do While L >= F
        If VowelOf(Mid$(Buf, L, 1)) = TONE_BREVE_ Then
            Fn = True
            Exit Do
        End If
        L = L - 1
    Loop
        
    If Not Fn Then
        L = Keys
        Fn = False
        Do While L >= F
            If VowelOf(Mid$(Buf, L, 1)) = TONE_ Then
                Fn = True
                Exit Do
            End If
            L = L - 1
        Loop
    End If
    
    If Not Fn Then
        L = Keys
        Fn = False
        Do While L >= F
            If VowelOf(Mid$(Buf, L, 1)) = BREVE_ Then
                Fn = True
                Exit Do
            End If
            L = L - 1
        Loop
    End If
    
    If Not Fn Then
        L = Keys
        Fn = False
        Do While L >= F
            If VowelOf(Mid$(Buf, L, 1)) = TONE_BREVE_ Then
                Fn = True
                Exit Do
            End If
            L = L - 1
        Loop
    End If
    
    If Not Fn Then
        PutC Ch
        Exit Sub
    End If
    
    If L < F Then L = F
    
    R = L
    uBuf = Mid$(Buf, R, Keys - R + 1)
    Bks = Len(uBuf)
    Dim sAnsi As String
    sAnsi = FromUni(Mid$(Buf, R, 1))
    If VowelOf(Mid$(Buf, R, 1)) = TONE_BREVE_ Then
        Mid$(Buf, R, 1) = ToUni(Left$(sAnsi, 2))
        uBuf = Mid$(Buf, R, Keys - R + 1)
        Exit Sub
    ElseIf VowelOf(Mid$(Buf, R, 1)) = TONE_ Then
        Mid$(Buf, R, 1) = Left$(sAnsi, 1)
        uBuf = Mid$(Buf, R, Keys - R + 1)
        Exit Sub
    ElseIf VowelOf(Mid$(Buf, R, 1)) = BREVE_ Then
        Mid$(Buf, R, 1) = Left$(sAnsi, 1)
        uBuf = Mid$(Buf, R, Keys - R + 1)
        Exit Sub
    End If
End Sub



Private Sub PLast(Ch As String)
    If (Not Vkey) Or (TOff) Or (Keys <= 0) Then
        PutC Ch
        Exit Sub
    End If
    
    Dim P As Integer, Fn As Boolean
    
    Fn = False
    P = GetLastWord(Buf)
    If P < LOff Then P = LOff
    
    Do While (P <= Keys)
        If VowelOf(Mid$(Buf, P, 1)) = TONE_ Then
            Fn = True
            Exit Do
        End If
        P = P + 1
    Loop
    
    If Not Fn Then
        PutC Ch
        Exit Sub
    End If
    If P > Keys Then P = Keys

    Dim I As Integer
    I = Keys
    
    Do While I >= P
        If VowelOf(Mid$(Buf, I, 1)) <> 0 Then
            Exit Do
        End If
        I = I - 1
    Loop
    
    If I < P Then I = P
    
    Dim S As String
    S = Mid$(Buf, P, I - P + 1)

    uBuf = Mid$(Buf, P, Keys - P + 1)
    Bks = Len(uBuf)
    PutC Ch
    uBuf = Mid$(Buf, P, Keys - P + 1)
    
    Select Case Len(S)
        Case 2:
            Dim S1 As String
            S1 = FromUni(Mid$(Buf, P, 1))
            
            If P < Keys And (UCase$(Left$(S1, 1)) = "O") Then
                If VowelOf(Mid$(Buf, P + 1, 1)) = NONE_ And (UCase$(Right$(S, 1)) = "A" Or UCase$(Right$(S, 1)) = "E") Then
                    Mid$(Buf, P + 1, 1) = ToUni(Mid$(Buf, P + 1, 1) & Right$(S1, 1))
                    Mid$(Buf, P, 1) = Left$(S1, 1)
                    uBuf = Mid$(Buf, P, Keys - P + 1)
                    Exit Sub
                End If
            ElseIf P < Keys And (UCase$(Left$(S1, 1)) = "U") Then
                If VowelOf(Mid$(Buf, P + 1, 1)) = NONE_ And UCase$(Right$(S, 1)) = "Y" Then
                    Mid$(Buf, P + 1, 1) = ToUni(Mid$(Buf, P + 1, 1) & Right$(S1, 1))
                    Mid$(Buf, P, 1) = Left$(S1, 1)
                    uBuf = Mid$(Buf, P, Keys - P + 1)
                    Exit Sub
                End If
            End If
        Case 3:
            'Chua xu ly
    End Select
    
End Sub

Private Function VowelOf(Ch As String) As VOWELS
    If InStr(1, "aeiouyAEIOUY", Ch, vbBinaryCompare) > 0 Then
        VowelOf = NONE_
        Exit Function
        
    ElseIf InStr(1, ChrW$(226) & ChrW$(259) & ChrW$(234) & ChrW$(244) & ChrW$(417) & ChrW$(432) & ChrW$(194) & ChrW$(258) & ChrW$(202) & ChrW$(212) & ChrW$(416) & ChrW$(431), Ch, vbTextCompare) > 0 Then
        VowelOf = BREVE_
        Exit Function
        
    ElseIf InStr(1, ChrW$(224) & ChrW$(225) & ChrW$(7843) & ChrW$(227) & _
                    ChrW$(7841) & ChrW$(232) & ChrW$(233) & ChrW$(7867) & _
                    ChrW$(7869) & ChrW$(7865) & ChrW$(236) & ChrW$(237) & _
                    ChrW$(7881) & ChrW$(297) & ChrW$(7883) & ChrW$(242) & _
                    ChrW$(243) & ChrW$(7887) & ChrW$(245) & ChrW$(7885) & _
                    ChrW$(249) & ChrW$(250) & ChrW$(7911) & ChrW$(361) & _
                    ChrW$(7909) & ChrW$(7923) & ChrW$(253) & ChrW$(7927) & _
                    ChrW$(7929) & ChrW$(7925) & ChrW$(192) & ChrW$(193) & _
                    ChrW$(7842) & ChrW$(195) & ChrW$(7840) & ChrW$(200) & _
                    ChrW$(201) & ChrW$(7866) & ChrW$(7868) & ChrW$(7864) & _
                    ChrW$(204) & ChrW$(205) & ChrW$(7880) & ChrW$(296) & _
                    ChrW$(7882) & ChrW$(210) & ChrW$(211) & ChrW$(7886) & _
                    ChrW$(213) & ChrW$(7884) & ChrW$(217) & ChrW$(218) & _
                    ChrW$(7910) & ChrW$(360) & ChrW$(7908) & ChrW$(7922) & _
                    ChrW$(221) & ChrW$(7926) & ChrW$(7928) & ChrW$(7924), Ch, vbTextCompare) > 0 Then
        VowelOf = TONE_
        Exit Function
        
    ElseIf InStr(1, ChrW$(7847) & ChrW$(7845) & ChrW$(7849) & ChrW$(7851) & _
                    ChrW$(7853) & ChrW$(7857) & ChrW$(7855) & ChrW$(7859) & _
                    ChrW$(7861) & ChrW$(7863) & ChrW$(7873) & ChrW$(7871) & _
                    ChrW$(7875) & ChrW$(7877) & ChrW$(7879) & ChrW$(7891) & _
                    ChrW$(7889) & ChrW$(7893) & ChrW$(7895) & ChrW$(7897) & _
                    ChrW$(7901) & ChrW$(7899) & ChrW$(7903) & ChrW$(7905) & _
                    ChrW$(7907) & ChrW$(7915) & ChrW$(7913) & ChrW$(7917) & _
                    ChrW$(7919) & ChrW$(7921) & ChrW$(7846) & ChrW$(7844) & _
                    ChrW$(7848) & ChrW$(7850) & ChrW$(7852) & ChrW$(7856) & _
                    ChrW$(7854) & ChrW$(7858) & ChrW$(7860) & ChrW$(7862) & _
                    ChrW$(7872) & ChrW$(7870) & ChrW$(7874) & ChrW$(7876) & _
                    ChrW$(7878) & ChrW$(7890) & ChrW$(7888) & ChrW$(7892) & _
                    ChrW$(7894) & ChrW$(7896) & ChrW$(7900) & ChrW$(7898) & _
                    ChrW$(7902) & ChrW$(7904) & ChrW$(7906) & ChrW$(7914) & _
                    ChrW$(7912) & ChrW$(7916) & ChrW$(7918) & ChrW$(7920), Ch, vbTextCompare) > 0 Then
        VowelOf = TONE_BREVE_
        Exit Function
    End If
    
    VowelOf = 0
End Function


Private Function ToUni(S As String) As String
    Dim I As Long, J As Long, sResult As String
    sResult = S
    For I = UBound(UArr) To LBound(UArr) Step -1
        For J = 1 To Len(S)
            If (LCase$(TArr(I)) = LCase$(Mid$(S, J, 3))) And (Mid$(S, J, 1) = Left$(TArr(I), 1)) Then sResult = Replace$(sResult, Mid$(S, J, 3), UArr(I))
        Next J
        
        For J = 1 To Len(S)
            If (LCase$(TArr(I)) = LCase$(Mid$(S, J, 2))) And (Mid$(S, J, 1) = Left$(TArr(I), 1)) Then sResult = Replace$(sResult, Mid$(S, J, 2), UArr(I))
        Next J
    Next I
    
    ToUni = sResult
    
End Function


Private Function FromUni(S As String) As String
    Dim I As Long, J As Long, sResult As String
    sResult = S
        
    For I = UBound(UArr) To LBound(UArr) Step -1
        For J = 1 To Len(S)
            If UArr(I) = Mid$(S, J, 1) Then sResult = Replace$(sResult, Mid$(S, J, 1), TArr(I))
        Next J
    Next I
    
    FromUni = sResult
    
End Function


Private Sub PDChar(Ch As String)
    If LOff > 0 Then
        If UCase$(Ch) <> UCase$(Mid$(Buf, LOff, 1)) Then TOff = False
    End If

    If (TOff = True) Or (Keys <= 0) Or (Not Vkey) Then
        PutC Ch
        Exit Sub
    End If
                                                                            
    Dim F As Integer, L As Integer, R As Integer
        
    F = GetLastWord(Buf)
    If F < LOff Then F = LOff
    L = Keys
    
    Dim Fn As Boolean
    
    Fn = False
    Do While F <= L And L - F <= MAX_WORD_LENGTH
        If (Mid$(Buf, F, 1) = "d" Or Mid$(Buf, F, 1) = "D" Or Mid$(Buf, F, 1) = ChrW$(272) Or Mid$(Buf, F, 1) = ChrW$(273)) Then
            Fn = True
            Exit Do
        End If
        F = F + 1
    Loop
    
    If Not Fn Then
        PutC Ch
        Exit Sub
    End If
        
    Fn = False
    Do While L >= F
        If (Mid$(Buf, L, 1) = ChrW$(272) Or Mid$(Buf, L, 1) = ChrW$(273)) Then
            Fn = True
            Exit Do
        End If
        L = L - 1
    Loop
    
    If Not Fn Then
        Fn = False
        Do While L >= F
            If (Mid$(Buf, L, 1) = "D" Or Mid$(Buf, L, 1) = "d") Then
                Fn = True
                Exit Do
            End If
            L = L - 1
        Loop
    End If
    
    If Not Fn Then L = F
    
    R = L
    
    If R > 1 Then
        If InStr(1, STRING_RESET & BEFORE_D, Mid$(Buf, R - 1, 1), vbTextCompare) <= 0 Then
            PutC Ch
            Exit Sub
        End If
    End If
    
    uBuf = Mid$(Buf, R, Keys - R + 1)
    Bks = Len(uBuf)
    
    If Mid$(Buf, R, 1) = "d" Then
        Mid$(Buf, R, 1) = ChrW$(273)
        uBuf = Mid$(Buf, R, Keys - R + 1)
        Exit Sub
    ElseIf Mid$(Buf, R, 1) = "D" Then
        Mid$(Buf, R, 1) = ChrW$(272)
        uBuf = Mid$(Buf, R, Keys - R + 1)
        Exit Sub
    Else
        TOff = True
        Mid$(Buf, R, 1) = Left$(FromUni(Mid$(Buf, R, 1)), 1)
        PutC Ch
        uBuf = Mid$(Buf, R, Keys - R + 1)
        Exit Sub
    End If
End Sub



Private Sub PDblChar(Ch As String)
    If LOff > 0 Then
        If UCase$(Ch) <> UCase$(Mid$(Buf, LOff, 1)) Then TOff = False
    End If

    If (TOff = True) Or (Keys <= 0) Or (Not Vkey) Then
        PutC Ch
        Exit Sub
    End If

    Dim F As Integer, L As Integer, P As Integer, sw As String, Fnd As Boolean
    
    F = GetLastWord(Buf)
    If F < LOff Then F = LOff
    L = Keys
    Fnd = False
    
    Do While F <= L
        If VowelOf(Mid$(Buf, F, 1)) <> 0 Then
            Fnd = True
            Exit Do
        End If
        F = F + 1
    Loop
    
    If Not Fnd Then
        PutC Ch
        Exit Sub
    End If
    
    Fnd = False
    Do While L >= F
        If UCase$(Left$(FromUni(Mid$(Buf, L, 1)), 1)) = UCase$(Ch) Then
            Fnd = True
            Exit Do
        End If
        L = L - 1
    Loop
    If L < F Then L = F
    
    If Not Fnd Then
        PutC Ch
        Exit Sub
    End If
    P = L
    
    Do While P >= F And L - P <= MAX_VOWEL_LENGTH
        If VowelOf(Mid$(Buf, P, 1)) = 0 Then
            Exit Do
        End If
        P = P - 1
    Loop
    
    If P < F Then P = F
    
    sw = Mid$(Buf, P, L - P + 1)


    F = P
    
    Select Case Len(sw)
        Case 1:
            If L < Keys Then
                If (InStr(1, "c,p,t", Mid$(Buf, L + 1, 1), vbTextCompare) > 0) And InStr(1, "f,r,x", Right$(FromUni(Left$(sw, 1)), 1), vbTextCompare) > 0 Then
                    PutC Ch
                    Exit Sub
                End If
            End If
            P = L
        Case 2:
            If VowelOf(Right$(sw, 1)) = NONE_ Then
                If UCase$(Right$(sw, 1)) = "A" Then
                    If L < Keys Then
                        If (VowelOf(Mid$(Buf, L + 1)) <> 0) Or (InStr(1, "c,m,n,p,t", Mid$(Buf, L + 1, 1), vbTextCompare) <= 0) Then
                            PutC Ch
                            Exit Sub
                        End If
                    End If
                    If (UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "U") Then
                        If (VowelOf(Left$(sw, 1)) = BREVE_) Then
                            Mid$(Buf, F, 1) = Left$(FromUni(Left$(sw, 1)), 1)
                            Mid$(Buf, L, 1) = ToUni(Mid$(Buf, L, 1) & Ch)
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                            Exit Sub
                        ElseIf (VowelOf(Left$(sw, 1)) = TONE_BREVE_) Then
                            Mid$(Buf, F, 1) = Left$(FromUni(Left$(sw, 1)), 1)
                            Mid$(Buf, L, 1) = ToUni(Mid$(Buf, L, 1) & Ch & Right$(FromUni(Left$(sw, 1)), 1))
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                            Exit Sub
                        ElseIf (VowelOf(Left$(sw, 1)) = TONE_) Then
                            If L < Keys Then
                                If (InStr(1, "c,p,t", Mid$(Buf, L + 1, 1), vbTextCompare) > 0) And InStr(1, "f,r,x", Right$(FromUni(Left$(sw, 1)), 1), vbTextCompare) > 0 Then
                                    PutC Ch
                                    Exit Sub
                                End If
                            End If
                            Mid$(Buf, F, 1) = Left$(FromUni(Left$(sw, 1)), 1)
                            Mid$(Buf, L, 1) = ToUni(Mid$(Buf, L, 1) & Ch & Right$(FromUni(Left$(sw, 1)), 1))
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                            Exit Sub
                        ElseIf (VowelOf(Left$(sw, 1)) = NONE_) Then
                            P = L
                        End If
                    Else
                        PutC Ch
                        Exit Sub
                    End If
                ElseIf UCase$(Right$(sw, 1)) = "E" Then
                    If L < Keys Then
                        If (InStr(1, "c,m,n,p,t,u", Mid$(Buf, L + 1, 1), vbTextCompare) <= 0) Then
                            PutC Ch
                            Exit Sub
                        End If
                        If (InStr(1, "c,p,t", Mid$(Buf, L + 1, 1), vbTextCompare) > 0) And InStr(1, "f,r,x", Right$(FromUni(Left$(sw, 1)), 1), vbTextCompare) > 0 Then
                            PutC Ch
                            Exit Sub
                        End If
                    End If
                
                    If (UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "I") Or (UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "Y") Then
                        If VowelOf(Left$(sw, 1)) = TONE_ Then

                            Mid$(Buf, F, 1) = Left$(FromUni(Left$(sw, 1)), 1)
                            Mid$(Buf, L, 1) = ToUni(Mid$(Buf, L, 1) & Ch & Right$(FromUni(Left$(sw, 1)), 1))
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                            Exit Sub
                        ElseIf VowelOf(Left$(sw, 1)) = NONE_ Then
                            P = L
                        End If
                    ElseIf (UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "U") Then
                        If VowelOf(Left$(sw, 1)) = TONE_BREVE_ Then
                            PutC Ch
                            Exit Sub
                        ElseIf VowelOf(Left$(sw, 1)) = TONE_ Then
                            If L < Keys Then
                                If (InStr(1, "c,p,t", Mid$(Buf, L + 1, 1), vbTextCompare) > 0) And InStr(1, "f,r,x", Right$(FromUni(Left$(sw, 1)), 1), vbTextCompare) > 0 Then
                                    PutC Ch
                                    Exit Sub
                                End If
                            End If
                            
                            Mid$(Buf, F, 1) = Left$(FromUni(Left$(sw, 1)), 1)
                            Mid$(Buf, L, 1) = ToUni(Mid$(Buf, L, 1) & Ch & Right$(FromUni(Left$(sw, 1)), 1))
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                        ElseIf VowelOf(Left$(sw, 1)) = BREVE_ Then
                            Mid$(Buf, F, 1) = Left$(FromUni(Left$(sw, 1)), 1)
                            Mid$(Buf, L, 1) = ToUni(Mid$(Buf, L, 1) & Ch)
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                        ElseIf VowelOf(Left$(sw, 1)) = NONE_ Then
                            P = L
                        End If
                    Else
                        PutC Ch
                        Exit Sub
                    End If
                ElseIf UCase$(Right$(sw, 1)) = "O" Then
                    If L < Keys Then
                        If (InStr(1, "c,m,n,p,t,i", Mid$(Buf, L + 1, 1), vbTextCompare) <= 0) Then
                            PutC Ch
                            Exit Sub
                        End If
                        
                        If (InStr(1, "c,p,t", Mid$(Buf, L + 1, 1), vbTextCompare) > 0) And InStr(1, "f,r,x", Right$(FromUni(Left$(sw, 1)), 1), vbTextCompare) > 0 Then
                            PutC Ch
                            Exit Sub
                        End If
                        
                    End If
                    If UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "U" Then
                        If VowelOf(Left$(sw, 1)) = TONE_BREVE_ Then
                            PutC Ch
                            Exit Sub
                        ElseIf VowelOf(Left$(sw, 1)) = TONE_ Then
                            If L < Keys Then
                                If (InStr(1, "c,p,t", Mid$(Buf, L + 1, 1), vbTextCompare) > 0) And InStr(1, "f,r,x", Right$(FromUni(Left$(sw, 1)), 1), vbTextCompare) > 0 Then
                                    PutC Ch
                                    Exit Sub
                                End If
                            End If
                            
                            Mid$(Buf, F, 1) = Left$(FromUni(Left$(sw, 1)), 1)
                            Mid$(Buf, L, 1) = ToUni(Mid$(Buf, L, 1) & Ch & Right$(FromUni(Left$(sw, 1)), 1))
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                        ElseIf VowelOf(Left$(sw, 1)) = BREVE_ Then
                            Mid$(Buf, F, 1) = Left$(FromUni(Left$(sw, 1)), 1)
                            Mid$(Buf, L, 1) = ToUni(Mid$(Buf, L, 1) & Ch)
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                        ElseIf VowelOf(Left$(sw, 1)) = NONE_ Then
                            P = L
                        End If
                    Else
                        PutC Ch
                        Exit Sub
                    End If
                End If
            ElseIf VowelOf(Right$(sw, 1)) = BREVE_ Then
                If UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "O" Then
                    If UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "U" Then
                        If VowelOf(Left$(sw, 1)) = NONE_ Then
                            P = L
                        ElseIf VowelOf(Left$(sw, 1)) = BREVE_ Then
                            Mid$(Buf, F, 1) = Left$(FromUni(Mid$(Buf, F, 1)), 1)
                            Mid$(Buf, L, 1) = ToUni(Left$(FromUni(Mid$(Buf, L, 1)), 1) & Ch)
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                            Exit Sub
                        Else
                            P = L
                        End If
                    Else
                        PutC Ch
                        Exit Sub
                    End If
                ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "A" Then
                    P = L
                ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "E" Then
                    P = L
                End If
            ElseIf VowelOf(Right$(sw, 1)) = TONE_ Then
                If UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "A" Then
                    If UCase$(Left$(sw, 1)) <> "U" Then
                        PutC Ch
                        Exit Sub
                    Else
                        P = L
                    End If
                ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "E" Then
                ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "O" Then
                End If
            ElseIf VowelOf(Right$(sw, 1)) = TONE_BREVE_ Then
                If UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "A" Then
                    PutC Ch
                    Exit Sub
                ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "E" Then
                    If UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "U" Then
                        P = L
                    ElseIf UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "I" Then
                        P = L
                    Else
                        PutC Ch
                        Exit Sub
                    End If
                ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "O" Then
                    If UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "U" Then
                        If VowelOf(Left$(sw, 1)) = NONE_ Then
                            Mid$(Buf, L, 1) = Left$(FromUni(Mid$(Buf, L, 1)), 1)
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                            Exit Sub
                        ElseIf VowelOf(Left$(sw, 1)) = BREVE_ Then
                            Mid$(Buf, F, 1) = Left$(FromUni(Mid$(Buf, F, 1)), 1)
                            Mid$(Buf, L, 1) = ToUni(Left$(FromUni(Mid$(Buf, L, 1)), 1) & Ch & Right$(FromUni(Mid$(Buf, L, 1)), 1))
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                            Exit Sub
                        'ElseIf VowelOf(Left$(sW, 1)) = TONE_ Then
                        'ElseIf VowelOf(Left$(sW, 1)) = TONE_BREVE_ Then
                        End If
                    Else
                        PutC Ch
                        Exit Sub
                    End If
                End If
            Else
                P = L
            End If
        Case 3
            If VowelOf(Right$(sw, 1)) > NONE_ Then
                P = L
            ElseIf VowelOf(Mid$(sw, 2, 1)) > NONE_ Then
                P = L - 1
            ElseIf VowelOf(Left$(sw, 1)) > NONE_ Then
                P = L - 2
            Else
                P = L
            End If
            If UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "E" And UCase$(Left$(FromUni(Mid$(sw, 2, 1)), 1)) = "Y" And UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "U" Then
                PutC Ch
                Exit Sub
            Else
                If VowelOf(Left$(sw, 1)) = TONE_ Then
                    Mid$(Buf, F, 1) = Left$(FromUni(Left$(sw, 1)), 1)
                    Mid$(Buf, L, 1) = ToUni(Mid$(Buf, L, 1) & Ch & Right$(FromUni(Left$(sw, 1)), 1))
                    P = F
                    uBuf = Mid$(Buf, P, Keys - P + 1)
                    Bks = Len(uBuf)
                    Exit Sub
                ElseIf VowelOf(Mid$(sw, 2, 1)) = TONE_ Then
                    Mid$(Buf, L - 1, 1) = Left$(FromUni(Mid$(sw, 2, 1)), 1)
                    Mid$(Buf, L, 1) = ToUni(Mid$(Buf, L, 1) & Ch & Right$(FromUni(Mid$(sw, 2, 1)), 1))
                    P = L - 1
                    uBuf = Mid$(Buf, P, Keys - P + 1)
                    Bks = Len(uBuf)
                    Exit Sub
                End If
            End If
        Case Else
            'Chua tim ra tu co 4 nguyen am
    End Select
    
    uBuf = Mid$(Buf, P, Keys - P + 1)
    Bks = Len(uBuf)
    
    Dim sAnsi As String
    sAnsi = FromUni(Mid$(Buf, P, 1))
    
    If VowelOf(Mid$(Buf, P, 1)) = NONE_ Then
        Mid$(Buf, P, 1) = ToUni(Mid$(Buf, P, 1) & Ch)
        uBuf = Mid$(Buf, P, Keys - P + 1)
        Exit Sub
    ElseIf VowelOf(Mid$(Buf, P, 1)) = BREVE_ Then
        If UCase$(Right$(sAnsi, 1)) = UCase$(Ch) Then
            Mid$(Buf, P, 1) = Left$(sAnsi, 1)
            PutC Ch
            LOff = Keys
            TOff = True
            uBuf = Mid$(Buf, P, Keys - P + 1)
            Exit Sub
        Else
            Mid$(Buf, P, 1) = ToUni(Left$(sAnsi, 1) & Ch)
            uBuf = Mid$(Buf, P, Keys - P + 1)
            Exit Sub
        End If
    ElseIf VowelOf(Mid$(Buf, P, 1)) = TONE_ Then
        Mid$(Buf, P, 1) = ToUni(Left$(sAnsi, 1) & Ch & Right$(sAnsi, 1))
        uBuf = Mid$(Buf, P, Keys - P + 1)
        Exit Sub
    ElseIf VowelOf(Mid$(Buf, P, 1)) = TONE_BREVE_ Then
        If UCase$(Mid$(sAnsi, 2, 1)) = UCase$(Ch) Then
            Mid$(Buf, P, 1) = ToUni(Left$(sAnsi, 1) & Right$(sAnsi, 1))
            PutC Ch
            LOff = Keys
            TOff = True
            uBuf = Mid$(Buf, P, Keys - P + 1)
            Exit Sub
        Else
            Mid$(Buf, P, 1) = ToUni(Left$(sAnsi, 1) & Ch & Right$(sAnsi, 1))
            uBuf = Mid$(Buf, P, Keys - P + 1)
            Exit Sub
        End If
    End If
End Sub



Private Sub PTMark(Ch As String)
    
    If (TOff = True) Or (Keys <= 0) Or (Not Vkey) Then
        PutC Ch
        Exit Sub
    End If
    
    Dim F As Integer, L As Integer, P As Integer, sw As String, Fnd As Boolean
    
    F = GetLastWord(Buf)
    If F < LOff Then F = LOff
    L = Keys
    Fnd = False
    
    Do While F <= L
        If VowelOf(Mid$(Buf, F, 1)) <> 0 Then
            Fnd = True
            Exit Do
        End If
        F = F + 1
    Loop
    
    If Not Fnd Then
        PutC Ch
        Exit Sub
    End If
    
    Do While L >= F
        If VowelOf(Mid$(Buf, L, 1)) <> 0 Then
            Fnd = True
            Exit Do
        End If
        L = L - 1
    Loop
    If L < F Then L = F
    
    P = F
    F = L
    Do While F >= P
        If (VowelOf(Mid$(Buf, F, 1)) = 0) Or (L - F > MAX_VOWEL_LENGTH) Then
            Exit Do
        End If
        F = F - 1
    Loop
    If F < P Then F = P
    
    sw = Mid$(Buf, F, L - F + 1)

    If L < Keys Then
        If (InStr(1, "f,r,x", Ch, vbTextCompare) > 0) And InStr(1, "c,t,p", Mid$(Buf, L + 1, 1), vbTextCompare) > 0 Then
            PutC Ch
            Exit Sub
        End If
    End If
    
    Select Case Len(sw)
    
        Case 1:
            If L > 1 Then
                If (UCase$(Mid$(Buf, L - 1, 1)) = "Q" And UCase$(sw) = "U") Then
                    PutC Ch
                    Exit Sub
                End If
            End If
            P = L
        Case 2:
            
            If F > 1 Then
                If (UCase$(Mid$(Buf, F - 1, 1)) = "Q" And UCase$(Mid$(Buf, F, 1)) = "U") Or (UCase$(Mid$(Buf, F - 1, 1)) = "G" And UCase$(Mid$(Buf, F, 1)) = "I") Then
                    P = L
                End If
            End If
            If UCase$(Right$(sw, 1)) = "O" And UCase$(Left$(FromUni(Left$(sw, 1)), 1)) <> "A" And UCase$(Left$(FromUni(Left$(sw, 1)), 1)) <> "E" Then
                PutC Ch
                Exit Sub
            End If
            If ((UCase$(Left$(sw, 1)) = "A" And UCase$(Right$(sw, 1)) = "Y") Or (UCase$(Left$(sw, 1)) = "A" And UCase$(Right$(sw, 1)) = "I") Or (UCase$(Left$(sw, 1)) = "O" And UCase$(Right$(sw, 1)) = "I") Or (UCase$(Left$(sw, 1)) = "U" And UCase$(Right$(sw, 1)) = "I") Or (UCase$(Left$(sw, 1)) = "U" And UCase$(Right$(sw, 1)) = "A") Or (UCase$(Left$(sw, 1)) = "O" And UCase$(Right$(sw, 1)) = "E")) Then
                P = L - 1
            ElseIf VowelOf(Right$(sw, 1)) > NONE_ Then
                P = L
            ElseIf VowelOf(Left$(sw, 1)) > NONE_ Then
                P = L - 1
            Else
                If ((UCase$(Left$(sw, 1)) = "O" And UCase$(Right$(sw, 1)) = "A") Or (UCase$(Left$(sw, 1)) = "O" And UCase$(Right$(sw, 1)) = "E") Or (UCase$(Left$(sw, 1)) = "U" And UCase$(Right$(sw, 1)) = "A") Or (UCase$(Left$(sw, 1)) = "U" And UCase$(Right$(sw, 1)) = "Y")) Then
                    If OStyle Then
                        If L < Keys Then
                            P = L
                        Else
                            P = L - 1
                        End If
                    Else
                        P = L
                    End If
                End If
            End If
        Case 3:
            If F > 1 Then
                If (UCase$(Mid$(Buf, F - 1, 1)) = "Q" And UCase$(Mid$(Buf, F, 1)) = "U") Then
                    P = L - 1
                End If
            End If
        

            Dim II As Integer
            Fnd = False
            For II = Len(sw) To 1 Step -1
                If VowelOf(Mid$(sw, II, 1)) = TONE_BREVE_ Then
                    Fnd = True
                    Exit For
                End If
            Next II
                
                
            If Not Fnd Then
                Fnd = False
                For II = Len(sw) To 1 Step -1
                    If VowelOf(Mid$(sw, II, 1)) = TONE_ Then
                        Fnd = True
                        Exit For
                    End If
                Next II
            Else
                If II = 3 Then
                    P = L
                ElseIf II = 2 Then
                    P = L - 1
                ElseIf II = 1 Then
                    P = F
                End If
            End If
                
            If Not Fnd Then
                Fnd = False
                For II = Len(sw) To 1 Step -1
                    If VowelOf(Mid$(sw, II, 1)) = BREVE_ Then
                        Fnd = True
                        Exit For
                    End If
                Next II
            Else
                If II = 3 Then
                    P = L
                ElseIf II = 2 Then
                    P = L - 1
                ElseIf II = 1 Then
                    P = L - 2
                End If
            End If
            
            If Not Fnd Then
                P = L - 1
            Else
                If II = 3 Then
                    P = L
                ElseIf II = 2 Then
                    P = L - 1
                ElseIf II = 1 Then
                    P = L - 2
                End If
            End If
                
    End Select
    
    uBuf = Mid$(Buf, P, Keys - P + 1)
    Bks = Len(uBuf)
    
    Dim sAnsi As String
    sAnsi = FromUni(Mid$(Buf, P, 1))
    If VowelOf(Mid$(Buf, P, 1)) = NONE_ Then
        Mid$(Buf, P, 1) = ToUni(Mid$(Buf, P, 1) & Ch)
        uBuf = Mid$(Buf, P, Keys - P + 1)
        Exit Sub
    ElseIf VowelOf(Mid$(Buf, P, 1)) = BREVE_ Then
        Mid$(Buf, P, 1) = ToUni(sAnsi & Ch)
        uBuf = Mid$(Buf, P, Keys - P + 1)
        Exit Sub
    ElseIf VowelOf(Mid$(Buf, P, 1)) = TONE_ Then
        If UCase$(Right$(sAnsi, 1)) = UCase$(Ch) Then
            Mid$(Buf, P, 1) = Left$(sAnsi, 1)
            PutC Ch
            LOff = Keys
            TOff = True
            uBuf = Mid$(Buf, P, Keys - P + 1)
            Exit Sub
        Else
            Mid$(Buf, P, 1) = ToUni(Left$(sAnsi, 1) & Ch)
            uBuf = Mid$(Buf, P, Keys - P + 1)
            Exit Sub
        End If
    ElseIf VowelOf(Mid$(Buf, P, 1)) = TONE_BREVE_ Then
        If UCase$(Right$(sAnsi, 1)) = UCase$(Ch) Then
            Mid$(Buf, P, 1) = ToUni(Left$(sAnsi, 2))
            PutC Ch
            LOff = Keys
            TOff = True
            uBuf = Mid$(Buf, P, Keys - P + 1)
            Exit Sub
        Else
            Mid$(Buf, P, 1) = ToUni(Left$(sAnsi, 2) & Ch)
            uBuf = Mid$(Buf, P, Keys - P + 1)
            Exit Sub
        End If
    End If
    
    
End Sub



Private Sub PBMark(Ch As String)
    If LOff > 0 Then
        If UCase$(Ch) <> UCase$(Mid$(Buf, LOff, 1)) Then TOff = False
    End If

    If ((TOff = True) Or (Not Vkey) Or Keys <= 0) Then
        PutC Ch
        Exit Sub
    End If
    
    Dim F As Integer, L As Integer, P As Integer, sw As String, Fnd As Boolean
    
    F = GetLastWord(Buf)
    If F < LOff Then F = LOff
    
    L = Keys
    Fnd = False
    
    Do While L >= F
        If (UCase$(Left$(FromUni(Mid$(Buf, L, 1)), 1)) = "A") Or ((UCase$(Left$(FromUni(Mid$(Buf, L, 1)), 1))) = "O") Or ((UCase$(Left$(FromUni(Mid$(Buf, L, 1)), 1))) = "U") Then
            Fnd = True
            Exit Do
        End If
        L = L - 1
    Loop
    
    If Not Fnd Then
        If InStr(1, "b,c,d,g,h,l,m,n,r,s,t,v,x" & STRING_RESET, Mid$(Buf, Keys, 1), vbTextCompare) > 0 Then
            uBuf = Right$(Buf, 1)
            Bks = Len(uBuf)

            PutC IIf(Ch = UCase$(Ch), ToUni("UW"), ToUni("uw"))
            Lw = True
            uBuf = Right$(Buf, 2)
        Else
            PutC Ch
        End If
        Exit Sub
    End If
    
    
    If L < F Then L = F
    P = L
    Fnd = False
    Do While (P >= F) And (L - P <= MAX_VOWEL_LENGTH)
        If VowelOf(Mid$(Buf, P, 1)) = 0 Then
            Exit Do
        End If
        P = P - 1
    Loop
    
    If P < F Then P = F
    
    Do While F <= L
        If VowelOf(Mid$(Buf, F, 1)) <> 0 Then
            Exit Do
        End If
        F = F + 1
    Loop
JUMP:
    sw = Mid$(Buf, F, L - F + 1)


    P = L

    Select Case Len(sw)
        Case 1:
            If UCase$(FromUni(sw)) = "A" Then
                If L < Keys Then
                    If VowelOf(Mid$(Buf, L + 1, 1)) > 0 Then
                        PutC Ch
                        Exit Sub
                    End If
                    If InStr(1, "c,m,n,p,t", Mid$(Buf, L + 1, 1), vbTextCompare) <= 0 Then
                        PutC Ch
                        Exit Sub
                    End If
                End If
                P = L
            End If
            If UCase$(sw) = "U" Then
                If L > 1 Then
                    If UCase$(Mid$(Buf, L - 1, 1)) = "Q" Then
                        PutC Ch
                        Exit Sub
                    Else
                        P = L
                    End If
                End If
                If L < Keys Then
                    If VowelOf(Mid$(Buf, L + 1, 1)) <> 0 And UCase$(Left$(FromUni(Mid$(Buf, L + 1, 1)), 1)) <> "O" And UCase$(Left$(FromUni(Mid$(Buf, L + 1, 1)), 1)) <> "A" Then
                        PutC Ch
                        Exit Sub
                    End If
                End If
            End If
            P = L
        Case 2:
            If VowelOf(Right$(sw, 1)) = NONE_ Then
                If UCase$(Right$(sw, 1)) = "A" Then
                
                    If L < Keys Then
                        If VowelOf(Mid$(Buf, L + 1, 1)) > 0 Then
                            PutC Ch
                            Exit Sub
                        End If
                        If InStr(1, "c,m,n,p,t", Mid$(Buf, L + 1, 1), vbTextCompare) <= 0 Then
                            PutC Ch
                            Exit Sub
                        End If
                    End If
                
                    If UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "U" Then
                        If F > 1 Then
                            If UCase$(Mid$(Buf, F - 1, 1)) = "Q" Then
                                P = L
                            Else
                                If L < Keys Then
                                    If (VowelOf(Left$(sw, 1)) = TONE_BREVE_) Or (VowelOf(Left$(sw, 1)) = TONE_) Or (VowelOf(Left$(sw, 1)) = NONE_) Then
                                        PutC Ch
                                        Exit Sub
                                    Else
                                        P = L - 1
                                    End If
                                Else
                                    P = L - 1
                                End If
                            End If
                        Else
                            P = L - 1
                        End If
                    ElseIf UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "O" Then
                        If VowelOf(Left$(sw, 1)) = NONE_ Then
                            P = L
                        ElseIf VowelOf(Left$(sw, 1)) = BREVE_ Then
                            Mid$(Buf, F, 1) = Left$(FromUni(Mid$(Buf, F, 1)), 1)
                            Mid$(Buf, L, 1) = ToUni(Mid$(Buf, L, 1) & Ch)
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                            Exit Sub
                        ElseIf VowelOf(Left$(sw, 1)) = TONE_ Then
                            If L < Keys Then
                                If InStr(1, "c,m,n,p,t", Mid$(Buf, L + 1, 1), vbTextCompare) <= 0 Then
                                    PutC Ch
                                    Exit Sub
                                End If
                                
                                If (InStr(1, "C,P,T", Mid$(Buf, L + 1, 1), vbTextCompare) > 0) And (InStr(1, "F,R,X", Right$(FromUni(Left$(sw, 1)), 1), vbTextCompare) > 0) Then
                                    PutC Ch
                                    Exit Sub
                                End If
                            End If
                            Mid$(Buf, F, 1) = Left$(FromUni(Mid$(Buf, F, 1)), 1)
                            Mid$(Buf, L, 1) = ToUni(Mid$(Buf, L, 1) & Ch & Right$(FromUni(Left$(sw, 1)), 1))
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                            Exit Sub
                        ElseIf VowelOf(Left$(sw, 1)) = TONE_BREVE_ Then
                            PutC Ch
                            Exit Sub
                        End If
                    Else
                        PutC Ch
                        Exit Sub
                    End If
                ElseIf UCase$(Right$(sw, 1)) = "O" Then
                    If L < Keys Then
                        If InStr(1, "i,c,m,n,p,t,u", Mid$(Buf, L + 1, 1), vbTextCompare) <= 0 Then
                            PutC Ch
                            Exit Sub
                        End If
                    End If
                
                    If UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "U" Then
                        If VowelOf(Left$(sw, 1)) = NONE_ Then
                            If F > 1 Then
                                If UCase$(Mid$(Buf, F - 1, 1)) = "Q" Then
                                    P = L
                                Else
                                    Mid$(Buf, L, 1) = ToUni(Mid$(Buf, L, 1) & Ch)
                                    Mid$(Buf, F, 1) = ToUni(Mid$(Buf, F, 1) & Ch)
                                    P = L - 1
                                    uBuf = Mid$(Buf, P, Keys - P + 1)
                                    Bks = Len(uBuf)
                                    Exit Sub
                                End If
                            Else
                                P = L - 1
                            End If
                        ElseIf VowelOf(Left$(sw, 1)) = BREVE_ Then
                            P = L
                        Else
                            Mid$(Buf, F, 1) = ToUni(Left$(FromUni(Mid$(Buf, F, 1)), 1) & Ch)
                            Mid$(Buf, L, 1) = ToUni(Mid$(Buf, L, 1) & Ch & Right$(FromUni(Left$(sw, 1)), 1))
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                            Exit Sub
                        End If
                    Else
                        P = L
                    End If
                ElseIf UCase$(Right$(sw, 1)) = "U" Then
                    If L < Keys Then
                        If (VowelOf(Mid$(Buf, L + 1, 1)) <> 0) Or (InStr(1, "c,m,n,p,t", Mid$(Buf, L + 1, 1), vbTextCompare) <= 0) Then
                            PutC Ch
                            Exit Sub
                        End If
                    End If
                    If UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "U" Then
                        If VowelOf(Left$(sw, 1)) = NONE_ Then
                            If L >= Keys Then
                                P = L - 1
                            Else
                                PutC Ch
                                Exit Sub
                            End If
                        ElseIf VowelOf(Left$(sw, 1)) = BREVE_ Then
                            P = L - 1
                        ElseIf VowelOf(Left$(sw, 1)) = TONE_ Then
                            P = L - 1
                        ElseIf VowelOf(Left$(sw, 1)) = TONE_BREVE_ Then
                            P = L - 1
                        End If
                    Else
                        P = L
                    End If
                End If
            ElseIf VowelOf(Right$(sw, 1)) = BREVE_ Then
                If VowelOf(Left$(sw, 1)) = NONE_ Then
                    If UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "A" Then
                        If UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "O" Then
                            P = L
                        Else
                            PutC Ch
                            Exit Sub
                        End If
                    ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "O" Then
                        Mid$(Buf, F, 1) = ToUni(Mid$(Buf, F, 1) & Ch)
                        Mid$(Buf, L, 1) = ToUni(Left$(FromUni(Mid$(Buf, L, 1)), 1) & Ch)
                        P = L - 1
                        uBuf = Mid$(Buf, P, Keys - P + 1)
                        Bks = Len(uBuf)
                        Exit Sub
                    ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "U" Then
                        PutC Ch
                        Exit Sub
                    End If
                ElseIf VowelOf(Left$(sw, 1)) = BREVE_ Then
                    If UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "A" Then
                        If UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "U" Then
                            P = L - 1
                        Else
                            PutC Ch
                            Exit Sub
                        End If
                    ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "O" Then
                        If VowelOf(Right$(sw, 1)) = NONE_ Then
                            If VowelOf(Left$(sw, 1)) = NONE_ Then
                                Mid$(Buf, F, 1) = ToUni(Mid$(Buf, F, 1) & Ch)
                                Mid$(Buf, L, 1) = ToUni(Left$(FromUni(Mid$(Buf, L, 1)), 1) & Ch)
                                P = L - 1
                                uBuf = Mid$(Buf, P, Keys - P + 1)
                                Bks = Len(uBuf)
                                Exit Sub
                            Else
                                P = L
                            End If
                        ElseIf VowelOf(Right$(sw, 1)) = BREVE_ Then
                            Mid$(Buf, L, 1) = Left$(FromUni(Right$(sw, 1)), 1)
                            P = L - 1
                        Else
                            P = L
                        End If
                    ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "U" Then
                        PutC Ch
                        Exit Sub
                    End If
                ElseIf VowelOf(Left$(sw, 1)) = TONE_ Then
                    If UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "A" Then
                        PutC Ch
                        Exit Sub
                    ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "O" Then
                        Mid$(Buf, F, 1) = ToUni(Mid$(Buf, F, 1) & Ch)
                        Mid$(Buf, L, 1) = ToUni(Left$(FromUni(Mid$(Buf, L, 1)), 1) & Ch)
                        P = L - 1
                        uBuf = Mid$(Buf, P, Keys - P + 1)
                        Bks = Len(uBuf)
                        Exit Sub
                    ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "U" Then
                        PutC Ch
                        Exit Sub
                    End If
                ElseIf VowelOf(Left$(sw, 1)) = TONE_BREVE_ Then
                 If UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "A" Then
                        PutC Ch
                        Exit Sub
                    ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "O" Then
                        Mid$(Buf, F, 1) = ToUni(Mid$(Buf, F, 1) & Ch)
                        Mid$(Buf, L, 1) = ToUni(Left$(FromUni(Mid$(Buf, L, 1)), 1) & Ch)
                        P = L - 1
                        uBuf = Mid$(Buf, P, Keys - P + 1)
                        Bks = Len(uBuf)
                        Exit Sub
                    ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "U" Then
                        PutC Ch
                        Exit Sub
                    End If
                End If
            ElseIf VowelOf(Right$(sw, 1)) = TONE_ Then
                If (UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "A") Then
                    If L < Keys Then
                        If VowelOf(Mid$(Buf, L + 1, 1)) > 0 Then
                            PutC Ch
                            Exit Sub
                        End If
                        If InStr(1, "c,m,n,p,t", Mid$(Buf, L + 1, 1), vbTextCompare) <= 0 Then
                            PutC Ch
                            Exit Sub
                        End If
                    End If
                    P = L
                ElseIf (UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "O") Then
                    If UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "U" Then
                        If VowelOf(Left$(sw, 1)) = NONE_ Then
                            Mid$(Buf, F, 1) = ToUni(Mid$(Buf, F, 1) & Ch)
                            Mid$(Buf, L, 1) = ToUni(Left$(FromUni(Mid$(Buf, L, 1)), 1) & Ch & Right$(FromUni(Mid$(Buf, L, 1)), 1))
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                            Exit Sub
                        ElseIf VowelOf(Left$(sw, 1)) = BREVE_ Then
                            Mid$(Buf, F, 1) = Left$(FromUni(Mid$(Buf, F, 1)), 1)
                            Mid$(Buf, L, 1) = ToUni(Left$(FromUni(Mid$(Buf, L, 1)), 1) & Ch & Right$(FromUni(Mid$(Buf, L, 1)), 1))
                            P = L - 1
                            uBuf = Mid$(Buf, P, Keys - P + 1)
                            Bks = Len(uBuf)
                            Exit Sub
                        Else
                            P = L
                        End If
                    Else
                        PutC Ch
                        Exit Sub
                    End If
                ElseIf (UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "U") Then
                    P = L
                End If
            ElseIf VowelOf(Right$(sw, 1)) = BREVE_ Then
                If UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "A" Then
                ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "O" Then
                ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "U" Then
                End If
            ElseIf VowelOf(Right$(sw, 1)) = TONE_BREVE_ Then
                If UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "A" Then
                    If UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "O" Then
                        P = L
                    Else
                        PutC Ch
                        Exit Sub
                    End If
                ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "O" Then
                    If UCase$(Mid$(FromUni(Right$(sw, 1)), 2, 1)) = UCase$(Ch) Then
                        If UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "U" Then
                            If VowelOf(Left$(sw, 1)) = NONE_ Then
                                P = L
                            ElseIf VowelOf(Left$(sw, 1)) = BREVE_ Then
                                Mid$(Buf, L, 1) = Left$(FromUni(Mid$(Buf, L, 1)), 1)
                                P = L - 1
                            Else
                                P = L
                            End If
                        Else
                            PutC Ch
                            Exit Sub
                        End If
                    Else
                        If UCase$(Left$(FromUni(Left$(sw, 1)), 1)) = "U" Then
                            If VowelOf(Left$(sw, 1)) = NONE_ Then
                                Mid$(Buf, F, 1) = ToUni(Mid$(Buf, F, 1) & Ch)
                                Mid$(Buf, L, 1) = ToUni(Left$(FromUni(Mid$(Buf, L, 1)), 1) & Ch & Right$(FromUni(Mid$(Buf, L, 1)), 1))
                                P = L - 1
                                uBuf = Mid$(Buf, P, Keys - P + 1)
                                Bks = Len(uBuf)
                                Exit Sub
                            ElseIf VowelOf(Left$(sw, 1)) = BREVE_ Then
                                Mid$(Buf, F, 1) = Left$(FromUni(Mid$(Buf, F, 1)), 1)
                                Mid$(Buf, L, 1) = ToUni(Left$(FromUni(Mid$(Buf, L, 1)), 1) & Ch & Right$(FromUni(Mid$(Buf, L, 1)), 1))
                                P = L - 1
                                uBuf = Mid$(Buf, P, Keys - P + 1)
                                Bks = Len(uBuf)
                                Exit Sub
                            Else
                                P = L
                            End If
                        Else
                            PutC Ch
                            Exit Sub
                        End If
                    End If
                ElseIf UCase$(Left$(FromUni(Right$(sw, 1)), 1)) = "U" Then
                    P = L
                End If
                
            End If
        Case 3:
            If (UCase$(Left$(FromUni(Left$(sw, 1)), 10))) = "U" And (UCase$(Left$(FromUni(Mid$(sw, 2, 1)), 10))) = "O" Then
                If InStr(1, "c,i,m,n,p,t,u", Right$(sw, 1), vbTextCompare) <= 0 Then
                    PutC Ch
                    Exit Sub
                End If
                L = L - 1
                GoTo JUMP
            Else
                PutC Ch
                Exit Sub
            End If
        Case Else
            P = L
    End Select
    
    uBuf = Mid$(Buf, P, Keys - P + 1)
    Bks = Len(uBuf)
    
    Dim sAnsi As String
    sAnsi = FromUni(Mid$(Buf, P, 1))
    
    If VowelOf(Mid$(Buf, P, 1)) = NONE_ Then
        If UCase$(Mid$(Buf, P, 1)) = "U" Then Lw = False
        Mid$(Buf, P, 1) = ToUni(Mid$(Buf, P, 1) & Ch)
        uBuf = Mid$(Buf, P, Keys - P + 1)
        Exit Sub
    ElseIf VowelOf(Mid$(Buf, P, 1)) = BREVE_ Then
        If UCase$(Right$(sAnsi, 1)) = UCase$(Ch) Then
            If (UCase$(Left$(sAnsi, 1)) = "U") And Lw Then
                Mid$(Buf, P, 1) = Right$(sAnsi, 1)
            Else
                Mid$(Buf, P, 1) = Left$(sAnsi, 1)
                PutC Ch
            End If
            LOff = Keys
            TOff = True
            uBuf = Mid$(Buf, P, Keys - P + 1)
            Exit Sub
        Else
            Mid$(Buf, P, 1) = ToUni(Left$(sAnsi, 1) & Ch)
            uBuf = Mid$(Buf, P, Keys - P + 1)
            Exit Sub
        End If
    ElseIf VowelOf(Mid$(Buf, P, 1)) = TONE_ Then
        Mid$(Buf, P, 1) = ToUni(Left$(sAnsi, 1) & Ch & Right$(sAnsi, 1))
        uBuf = Mid$(Buf, P, Keys - P + 1)
        Exit Sub
    ElseIf VowelOf(Mid$(Buf, P, 1)) = TONE_BREVE_ Then
        If UCase$(Mid$(sAnsi, 2, 1)) = UCase(Ch) Then
            Mid$(Buf, P, 1) = ToUni(Left$(sAnsi, 1) & Right$(sAnsi, 1))
            PutC Ch
            LOff = Keys
            TOff = True
            uBuf = Mid$(Buf, P, Keys - P + 1)
            Exit Sub
        Else
            Mid$(Buf, P, 1) = ToUni(Left$(sAnsi, 1) & Ch & Right$(sAnsi, 1))
            uBuf = Mid$(Buf, P, Keys - P + 1)
            Exit Sub
        End If
    End If
    
End Sub



Private Sub SetData(S As String)
    Dim sPtr As Long, iLen As Long, iLock As Long
    OpenClipboard 0&
    EmptyClipboard
    iLen = LenB(S) + 2
    sPtr = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, iLen)
    iLock = GlobalLock(sPtr)
    lstrcpy iLock, StrPtr(S)
    GlobalUnlock sPtr
    SetClipboardData CF_UNICODETEXT, sPtr
    CloseClipboard
End Sub
