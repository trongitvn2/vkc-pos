Attribute VB_Name = "CapInternals"


Option Explicit

Type OPENFILENAME
   lStructSize As Long
   hwndOwner As Long
   hInstance As Long
   lpstrFilter As String
   lpstrCustomFilter As String
   nMaxCustFilter As Long
   nFilterIndex As Long
   lpstrFile As String
   nMaxFile As Long
   lpstrFileTitle As String
   nMaxFileTitle As Long
   lpstrInitialDir As String
   lpstrTitle As String
   Flags As Long
   nFileOffset As Integer
   nFileExtension As Integer
   lpstrDefExt As String
   lCustData As Long
   lpfnHook As Long
   lpTemplateName As String
End Type

Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

'------------------------------------------------------------------
' DirectShow filters property type
'-------------------------------------------------------------------
Global Const PROP_VIDEO_DEVICE = &H1
Global Const PROP_VIDEO_CAPTURE_STREAM = &H2
Global Const PROP_VIDEO_PREVIEW_STREAM = &H4
Global Const PROP_AUDIO_DEVICE = &H8
Global Const PROP_AUDIO_CAPTURE_STREAM = &H10
Global Const PROP_TV_TUNER = &H20
Global Const PROP_TV_AUDIO = &H40
Global Const PROP_CROSSBAR = &H80
Global Const PROP_VIDEO_COMPRESSOR = &H100
Global Const PROP_SECOND_CROSSBAR = &H200
Global Const PROP_AUDIO_RENDERER = &H400
Global Const PROP_VIDEO_PROC_1 = &H1000
Global Const PROP_VIDEO_PROC_2 = &H2000
Global Const PROP_VIDEO_PROC_3 = &H4000
Global Const PROP_VIDEO_PROC_4 = &H8000
Global Const PROP_VIDEO_PROC_5 = &H10000
Global Const PROP_AUDIO_PROC_1 = &H100000
Global Const PROP_AUDIO_PROC_2 = &H200000
Global Const PROP_AUDIO_PROC_3 = &H400000

Global Const FCATEGORY_VIDEO_DEVICE = 0
Global Const FCATEGORY_VIDEO_COMPRESSOR = 1
Global Const FCATEGORY_AUDIO_DEVICE = 2
Global Const FCATEGORY_AUDIO_COMPRESSOR = 3
Global Const FCATEGORY_VIDEO_EFFECTS1 = 4
Global Const FCATEGORY_VIDEO_EFFECTS2 = 5
Global Const FCATEGORY_VIDEO_FILTERS_ALL = 6
Global Const FCATEGORY_VIDEO_FILTERS_RGB = 7
Global Const FCATEGORY_VIDEO_FILTERS_YUV = 8
Global Const FCATEGORY_AUDIO_FILTERS_ALL = 9

'------------------------------------------------------------------
' DirectShow IDs for status and error event
'------------------------------------------------------------------

Global Const IDS_CAP_BEGIN = 300               '/* "Capture Start" */
Global Const IDS_CAP_END = 301                 '/* "Capture End" */

Global Const IDS_CAP_INFO = 401                '/* "%s" */
Global Const IDS_CAP_OUTOFMEM = 402            '/* "Out of memory" */
Global Const IDS_CAP_FILEEXISTS = 403          '/* "File '%s' exists -- overwrite it?" */
Global Const IDS_CAP_ERRORPALOPEN = 404        '/* "Error opening palette '%s'" */
Global Const IDS_CAP_ERRORPALSAVE = 405        '/* "Error saving palette '%s'" */
Global Const IDS_CAP_ERRORDIBSAVE = 406        '/* "Error saving frame '%s'" */
Global Const IDS_CAP_DEFAVIEXT = 407           '/* "avi" */
Global Const IDS_CAP_DEFPALEXT = 408           '/* "pal" */
Global Const IDS_CAP_CANTOPEN = 409            '/* "Cannot open '%s'" */
Global Const IDS_CAP_SEQ_MSGSTART = 410        '/* "Select OK to start capture\nof video sequence\nto %s." */
Global Const IDS_CAP_SEQ_MSGSTOP = 411         '/* "Hit ESCAPE or click to end capture" */

Global Const IDS_CAP_VIDEDITERR = 412          '/* "An error occurred while trying to run VidEdit." */
Global Const IDS_CAP_READONLYFILE = 413        '/* "The file '%s' is a read-only file." */
Global Const IDS_CAP_WRITEERROR = 414          '/* "Unable to write to file '%s'.\nDisk may be full." */
Global Const IDS_CAP_NODISKSPACE = 415         '/* "There is no space to create a capture file on the specified device." */
Global Const IDS_CAP_SETFILESIZE = 416         '/* "Set File Size" */
Global Const IDS_CAP_SAVEASPERCENT = 417       '/* "SaveAs: %2ld%%  Hit Escape to abort." */

Global Const IDS_CAP_DRIVER_ERROR = 418        '/* Driver specific error message */

Global Const IDS_CAP_WAVE_OPEN_ERROR = 419     '/* "Error: Cannot open the wave input device.\nCheck sample size, frequency, and channels." */
Global Const IDS_CAP_WAVE_ALLOC_ERROR = 420    '/* "Error: Out of memory for wave buffers." */
Global Const IDS_CAP_WAVE_PREPARE_ERROR = 421  '/* "Error: Cannot prepare wave buffers." */
Global Const IDS_CAP_WAVE_ADD_ERROR = 422      '/* "Error: Cannot add wave buffers." */
Global Const IDS_CAP_WAVE_SIZE_ERROR = 423     '/* "Error: Bad wave size." */

Global Const IDS_CAP_VIDEO_OPEN_ERROR = 424    '/* "Error: Cannot open the video input device." */
Global Const IDS_CAP_VIDEO_ALLOC_ERROR = 425   '/* "Error: Out of memory for video buffers." */
Global Const IDS_CAP_VIDEO_PREPARE_ERROR = 426 '/* "Error: Cannot prepare video buffers." */
Global Const IDS_CAP_VIDEO_ADD_ERROR = 427     '/* "Error: Cannot add video buffers." */
Global Const IDS_CAP_VIDEO_SIZE_ERROR = 428    '/* "Error: Bad video size." */

Global Const IDS_CAP_FILE_OPEN_ERROR = 429     '/* "Error: Cannot open capture file." */
Global Const IDS_CAP_FILE_WRITE_ERROR = 430    '/* "Error: Cannot write to capture file.  Disk may be full." */
Global Const IDS_CAP_RECORDING_ERROR = 431     '/* "Error: Cannot write to capture file.  Data rate too high or disk full." */
Global Const IDS_CAP_RECORDING_ERROR2 = 432    '/* "Error while recording" */
Global Const IDS_CAP_AVI_INIT_ERROR = 433      '/* "Error: Unable to initialize for capture." */
Global Const IDS_CAP_NO_FRAME_CAP_ERROR = 434  '/* "Warning: No frames captured.\nConfirm that vertical sync interrupts\nare configured and enabled." */
Global Const IDS_CAP_NO_PALETTE_WARN = 435     '/* "Warning: Using default palette." */
Global Const IDS_CAP_MCI_CONTROL_ERROR = 436   '/* "Error: Unable to access MCI device." */
Global Const IDS_CAP_MCI_CANT_STEP_ERROR = 437 '/* "Error: Unable to step MCI device." */
Global Const IDS_CAP_NO_AUDIO_CAP_ERROR = 438  '/* "Error: No audio data captured.\nCheck audio card settings." */
Global Const IDS_CAP_AVI_DRAWDIB_ERROR = 439   '/* "Error: Unable to draw this data format." */
Global Const IDS_CAP_COMPRESSOR_ERROR = 440    '/* "Error: Unable to initialize compressor." */
Global Const IDS_CAP_AUDIO_DROP_ERROR = 441    '/* "Error: Audio data was lost during capture, reduce capture rate." */

Global Const IDS_CAP_STAT_LIVE_MODE = 500      '/* "Live window" */
Global Const IDS_CAP_STAT_OVERLAY_MODE = 501   '/* "Overlay window" */
Global Const IDS_CAP_STAT_CAP_INIT = 502       '/* "Setting up for capture - Please wait" */
Global Const IDS_CAP_STAT_CAP_FINI = 503       '/* "Finished capture, now writing frame %ld" */
Global Const IDS_CAP_STAT_PALETTE_BUILD = 504  '/* "Building palette map" */
Global Const IDS_CAP_STAT_OPTPAL_BUILD = 505   '/* "Computing optimal palette" */
Global Const IDS_CAP_STAT_I_FRAMES = 506       '/* "%d frames" */
Global Const IDS_CAP_STAT_L_FRAMES = 507       '/* "%ld frames" */
Global Const IDS_CAP_STAT_CAP_L_FRAMES = 508   '/* "Captured %ld frames" */
Global Const IDS_CAP_STAT_CAP_AUDIO = 509      '/* "Capturing audio" */
Global Const IDS_CAP_STAT_VIDEOCURRENT = 510   '/* "Captured %ld frames (%ld dropped) %d.%03d sec." */
Global Const IDS_CAP_STAT_VIDEOAUDIO = 511     '/* "Captured %d.%03d sec.  %ld frames (%ld dropped) (%d.%03d fps).  %ld audio bytes (%d,%03d sps)" */
Global Const IDS_CAP_STAT_VIDEOONLY = 512      '/* "Captured %d.%03d sec.  %ld frames (%ld dropped) (%d.%03d fps)" */
Global Const IDS_CAP_STAT_FRAMESDROPPED = 513  '/* "Dropped %ld of %ld frames (%d.%02d%%) during capture." */

Function GetSaveName(Initial As String, FileType As Integer) As String
Dim of As OPENFILENAME
Dim fn As String * 260
Dim rc As Long, Index As Long, Pos As Long

fn = Chr(0)
If (Len(Initial) > 0) Then
   Index = 0
   Do
      Pos = Index
      Index = InStr(Pos + 1, Initial, "\", vbTextCompare)
   Loop While (Index > 0)
   of.lpstrFile = Right(Initial, Len(Initial) - Pos)
   of.lpstrInitialDir = Left$(Initial, Pos - 1)
End If
of.lStructSize = Len(of)
of.hwndOwner = MainForm.hWnd
Select Case FileType
   Case 0
      of.lpstrFilter = "All Supported Types" & Chr(0) & "*.bmp;*.jpg;*.pic" & Chr(0) & "BMP Files" & Chr(0) & "*.bmp" & Chr(0) & "JPEG Files" & Chr(0) & "*.jpg" & Chr(0) & "ePIC Files" & Chr(0) & "*.pic" & Chr(0)
      of.lpstrDefExt = "jpg"
   Case 1
      of.lpstrFilter = "AVI File" & Chr(0) & "*.avi" & Chr(0)
      of.lpstrDefExt = "avi"
    Case 2
      of.lpstrFilter = "CapturePro 3.0 profile (.cpf)" & Chr(0) & "*.cpf" & Chr(0)
      of.lpstrDefExt = "cpf"
End Select
of.lpstrFile = fn
of.nMaxFile = 260

rc = GetSaveFileName(of)

If (rc > 0) Then
   GetSaveName = of.lpstrFile
Else
   GetSaveName = Initial
End If

End Function

Function GetLoadName(Initial As String, FileType As Integer) As String
' Call the GetOpenFileName function directly so we don't
' have to use the Common Dialog custom control

Dim of As OPENFILENAME
Dim fn As String * 260
Dim rc As Long, Index As Long, Pos As Long

fn = Chr(0)
If (Len(Initial) > 0) Then
   Index = 0
   Do
      Pos = Index
      Index = InStr(Pos + 1, Initial, "\", vbTextCompare)
   Loop While (Index > 0)
   of.lpstrFile = Right(Initial, Len(Initial) - Pos)
   of.lpstrInitialDir = Left$(Initial, Pos - 1)
End If
of.lStructSize = Len(of)
of.hwndOwner = MainForm.hWnd
Select Case FileType
   Case 0
      of.lpstrFilter = "All Supported Types" & Chr(0) & "*.bmp;*.jpg;*.pic" & Chr(0) & "BMP Files" & Chr(0) & "*.bmp" & Chr(0) & "JPEG Files" & Chr(0) & "*.jpg" & Chr(0) & "ePIC Files" & Chr(0) & "*.pic" & Chr(0)
      of.lpstrDefExt = "jpg"
   Case 1
      of.lpstrFilter = "AVI File" & Chr(0) & "*.avi" & Chr(0)
      of.lpstrDefExt = "avi"
    Case 2
      of.lpstrFilter = "CapturePro 3.0 profile (.cpf)" & Chr(0) & "*.cpf" & Chr(0)
      of.lpstrDefExt = "cpf"
End Select
of.lpstrFile = fn
of.nMaxFile = 260

rc = GetOpenFileName(of)

If (rc > 0) Then
   GetLoadName = of.lpstrFile
Else
   GetLoadName = Initial
End If

End Function

