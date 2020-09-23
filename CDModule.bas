Attribute VB_Name = "CDModule"
Option Explicit
Public Declare Function SHShutDownDialog Lib "shell32" Alias "#60" (ByVal YourGuess As Long) As Long
Private Type OPENFILENAME
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
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Type PrintDlg
     lStructSize As Long
     hwndOwner As Long
     hDevMode As Long
     hDevNames As Long
     hdc As Long
     flags As Long
     nFromPage As Integer
     nToPage As Integer
     nMinPage As Integer
     nMaxPage As Integer
     nCopies As Integer
     hInstance As Long
     lCustData As Long
     lpfnPrintHook As Long
     lpfnSetupHook As Long
     lpPrintTemplateName As String
     lpSetupTemplateName As String
     hPrintTemplate As Long
     hSetupTemplate As Long
End Type
Const PD_NOSELECTION = &H4
Const PD_PRINTTOFILE = &H20
Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PrintDlg) As Long
Public Const GMEM_FIXED = &H0
Public Const GMEM_ZEROINIT = &H40
Public Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Public Const LF_FACESIZE = 32
Public Const CF_ANSIONLY = &H400&
Public Const CF_APPLY = &H200&
Public Const CF_BITMAP = 2
Public Const CF_DIB = 8
Public Const CF_DIF = 5
Public Const CF_DSPBITMAP = &H82
Public Const CF_DSPENHMETAFILE = &H8E
Public Const CF_DSPMETAFILEPICT = &H83
Public Const CF_DSPTEXT = &H81
Public Const CF_EFFECTS = &H100&
Public Const CF_ENABLEHOOK = &H8&
Public Const CF_ENABLETEMPLATE = &H10&
Public Const CF_ENABLETEMPLATEHANDLE = &H20&
Public Const CF_ENHMETAFILE = 14
Public Const CF_FIXEDPITCHONLY = &H4000&
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_GDIOBJFIRST = &H300
Public Const CF_GDIOBJLAST = &H3FF
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_LIMITSIZE = &H2000&
Public Const CF_METAFILEPICT = 3
Public Const CF_NOFACESEL = &H80000
Public Const CF_NOVECTORFONTS = &H800&
Public Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Public Const CF_NOSCRIPTSEL = &H800000
Public Const CF_NOSIMULATIONS = &H1000&
Public Const CF_NOSIZESEL = &H200000
Public Const CF_NOSTYLESEL = &H100000
Public Const CF_NOVERTFONTS = &H1000000
Public Const CF_OEMTEXT = 7
Public Const CF_OWNERDISPLAY = &H80
Public Const CF_PALETTE = 9
Public Const CF_PRINTERFONTS = &H2
Public Const CF_PENDATA = 10
Public Const CF_PRIVATEFIRST = &H200
Public Const CF_PRIVATELAST = &H2FF
Public Const CF_RIFF = 11
Public Const CF_SCREENFONTS = &H1
Public Const CF_SELECTSCRIPT = &H400000
Public Const CF_SCRIPTSONLY = CF_ANSIONLY
Public Const CF_SHOWHELP = &H4&
Public Const CF_SYLK = 4
Public Const CF_TIFF = 6
Public Const CF_TEXT = 1
Public Const CF_TTONLY = &H40000
Public Const CF_UNICODETEXT = 13
Public Const CF_USESTYLE = &H80&
Public Const CF_WAVE = 12
Public Const CF_WYSIWYG = &H8000
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
  
Public Const MAX_PATH = 260
  

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long

Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Type LOGFONT
        lfHeight As Long
        lfWidth As Long
        lfEscapement As Long
        lfOrientation As Long
        lfWeight As Long
        lfItalic As Byte
        lfUnderline As Byte
        lfStrikeOut As Byte
        lfCharSet As Byte
        lfOutPrecision As Byte
        lfClipPrecision As Byte
        lfQuality As Byte
        lfPitchAndFamily As Byte
        lfFaceName(LF_FACESIZE) As Byte
End Type

Type ChooseFont
        lStructSize As Long
        hwndOwner As Long
        hdc As Long
        lpLogFont As Long
        iPointSize As Long
        flags As Long
        rgbColors As Long
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
        hInstance As Long
                                      
        lpszStyle As String
        nSizeMax As Long
        nFontType As Integer
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long
End Type

Private Const OFN_HIDEREADONLY = &H4

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As ChooseFont) As Long
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Type ChooseColor
     lStructSize As Long
     hwndOwner As Long
     hInstance As Long
     rgbResult As Long
     lpCustColors As String
     flags As Long
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type
Const CC_RGBINIT = &H1
Const CC_FULLOPEN = &H2
Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As ChooseColor) As Long

Public Function ShowColorDlg(hWndForm As Long, vColor As Variant, Optional FullOpen As Boolean = False) As Long
Dim udtCC As ChooseColor
Dim iReturn As Long
    With udtCC
        .rgbResult = vColor
        .lStructSize = Len(udtCC)
        .hwndOwner = hWndForm
        .hInstance = App.hInstance
    End With
    
    If FullOpen = True Then
       udtCC.flags = CC_RGBINIT Or CC_FULLOPEN
    Else
       udtCC.flags = CC_RGBINIT
    End If
        
    udtCC.lpCustColors = String$(16 * 4, 0)
        
    
    iReturn = ChooseColor(udtCC)
    If (iReturn) Then
           ShowColorDlg = (udtCC.rgbResult)
    Else
           ShowColorDlg = -1
    End If
End Function

Public Function ChooseFontDialog(hWnd As Long, obj As Object)
Dim CFont As ChooseFont
Dim hMem As Long
Dim LFont As LOGFONT
Dim sFontName As String

    hMem = GlobalAlloc(GPTR, Len(LFont))
    
    CFont.hInstance = App.hInstance
    CFont.hwndOwner = hWnd
    CFont.lpLogFont = hMem
    CFont.lStructSize = Len(CFont)
    CFont.flags = CF_BOTH
    If ChooseFont(CFont) Then
        CopyMemory LFont, ByVal hMem, Len(LFont)
        sFontName = Space(LF_FACESIZE)
        CopyMemory ByVal sFontName, LFont.lfFaceName(0), LF_FACESIZE
        With obj.Font
            .Name = sFontName
            .Size = CFont.iPointSize / 10
            .Bold = LFont.lfWeight
            .Italic = LFont.lfItalic
            .Underline = LFont.lfUnderline
            .Strikethrough = LFont.lfStrikeOut
            .Charset = LFont.lfCharSet
        End With
    End If
    Call GlobalFree(hMem)
End Function
Public Function BrowseForFolder(hwndOwner As Long, sPrompt As String) As String
   Dim udtBrowseInfo As BrowseInfo
   Dim iNull As Integer
   Dim lpIDList As Long
   Dim lResult As Long
   Dim sPath As String
     With udtBrowseInfo
        .hwndOwner = hwndOwner
        .lpszTitle = lstrcat(sPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
     End With
     lpIDList = SHBrowseForFolder(udtBrowseInfo)
     If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        lResult = SHGetPathFromIDList(lpIDList, sPath)
  
        Call CoTaskMemFree(lpIDList)
        iNull = InStr(sPath, vbNullChar)
        If iNull Then sPath = Left$(sPath, iNull - 1)
     End If
    BrowseForFolder = sPath
End Function
Public Function ShowOpenDlg(hWndForm As Long, Optional sFilter As String = "", Optional sTitle As String = "Open") As String

    Dim ofn As OPENFILENAME
    Dim iReturn As Long
    
    If sFilter = "" Then
        sFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    End If
    
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = hWndForm
        .hInstance = App.hInstance
        .lpstrFilter = sFilter
        .lpstrFile = Space$(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = CurDir
        .lpstrTitle = sTitle
        .flags = OFN_HIDEREADONLY
    End With
       
        iReturn = GetOpenFileName(ofn)

        If (iReturn) Then
                ShowOpenDlg = Trim$(ofn.lpstrFile)
        Else
                ShowOpenDlg = "Cancel"
        End If
End Function
Public Function ShowPrint(hWndForm As Long, hdcForm As Long, Optional PrintToFile As Boolean = False, Optional iFormPage As Integer = 0, Optional iToPage As Integer = 0, Optional iNumOfPage As Integer = 0, Optional iNumOfCopies As Integer = 1) As Boolean
Dim lFromPage As Integer, lMin As Integer, lMax As Integer, lCopies As Integer
Dim iReturn As Long
Dim tPrintDlg As PrintDlg

    With tPrintDlg
         .lStructSize = Len(tPrintDlg)
         .hwndOwner = hWndForm
         .hdc = hdcForm
        
        If PrintToFile = True Then
            .flags = PD_NOSELECTION Or PD_PRINTTOFILE
        Else
            .flags = PD_NOSELECTION
        End If
        
        .nFromPage = iFormPage
        .nToPage = iToPage
        .nMinPage = 0
        .nMaxPage = iNumOfPage
        .nCopies = iNumOfCopies
        .hInstance = App.hInstance
        .lpPrintTemplateName = "Print"""
    End With
    iReturn = PrintDlg(tPrintDlg)
    If iReturn Then
            lFromPage = tPrintDlg.nFromPage
            lFromPage = tPrintDlg.nToPage
            lMin = tPrintDlg.nMinPage
            lMax = tPrintDlg.nMaxPage
            lCopies = tPrintDlg.nCopies
            If tPrintDlg.flags = (PD_NOSELECTION Or PD_PRINTTOFILE) Then Debug.Print "File"
            ShowPrint = True
    Else
            ShowPrint = False
    End If
End Function
Public Function ShowSavedlg(hWndForm As Long, Optional sFilter As String = "", Optional sTitle As String = "Save As") As String
Dim ofn As OPENFILENAME
Dim iReturn As Long
    If sFilter = "" Then
        sFilter = "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    End If
    With ofn
        .lStructSize = Len(ofn)
        .hwndOwner = hWndForm
        .hInstance = App.hInstance
        .lpstrFilter = sFilter
        .lpstrFile = Space$(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space$(254)
        .nMaxFileTitle = 255
        .lpstrInitialDir = CurDir
        .lpstrTitle = sTitle
        .flags = OFN_HIDEREADONLY
    End With
    iReturn = GetSaveFileName(ofn)
    If (iReturn) Then
       ShowSavedlg = Trim$(ofn.lpstrFile)
    Else
        ShowSavedlg = "Cancel"
    End If
End Function
