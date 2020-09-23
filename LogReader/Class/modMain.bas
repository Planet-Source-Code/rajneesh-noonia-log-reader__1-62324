Attribute VB_Name = "modMain"
'***********************************************************************************
'FILENAME       :   cGlobal.cls

'AUTHOR         :   Rajneesh Noonia

'VERSION HISTORY:
'                   Date        Who             Version     Description
'                   12/04/05    Rajneesh Noonia    1.0      Implementation

'DESCRIPTION    :   This only one creatable class in this project which provides
'                   Instances of other classes to client.Some general functions
'                   are also defined in this class.This class also proves common
'                   dialog box functionalities
'
'***********************************************************************************

Option Explicit

Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long '  only used if FOF_SIMPLEPROGRESS
End Type
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Const FO_COPY = &H2
Private Const FO_DELETE = &H3
Private Const FO_MOVE = &H1
Private Const FO_RENAME = &H4
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10  ' No Confirmation
Private Const FOF_NOCONFIRMMKDIR = &H200
Private Const FOF_SIMPLEPROGRESS = &H100



Private Type GUIDs
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Enum SysParameter
    Domain = 0
    Computer = 1
    User = 2
End Enum

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long
Public Enum EShellShowConstants
     essSW_HIDE = 0
     essSW_MAXIMIZE = 3
     essSW_MINIMIZE = 6
     essSW_SHOWMAXIMIZED = 3
     essSW_SHOWMINIMIZED = 2
     essSW_SHOWNORMAL = 1
     essSW_SHOWNOACTIVATE = 4
     essSW_SHOWNA = 8
     essSW_SHOWMINNOACTIVE = 7
     essSW_SHOWDEFAULT = 10
     essSW_RESTORE = 9
     essSW_SHOW = 5
End Enum
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5            '  access denied
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2                     '  file not found
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_PNF = 3                     '  path not found
Private Const SE_ERR_OOM = 8                     '  out of memory
Private Const SE_ERR_SHARE = 26

Public Enum EErrorCommonDialog
    eeBaseCommonDialog = 13450  ' CommonDialog
End Enum

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalCompact Lib "kernel32" (ByVal dwMinFree As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
    
    
Private Const SM_CXVSCROLL As Long = 2

Private Const SC_CLOSE As Long = &HF060&
Private Const MIIM_STATE As Long = &H1&
Private Const MIIM_ID As Long = &H2&
Private Const MFS_GRAYED As Long = &H3&
Private Const WM_NCACTIVATE As Long = &H86

Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MINIMIZEBOX As Long = &H20000
Private Const WS_MAXIMIZEBOX As Long = &H10000
Private Const GWL_STYLE As Long = (-16&)
Private Const SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4

Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd&, ByVal nIndex&) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd&, ByVal hWndInsertAfter&, ByVal x&, ByVal y&, ByVal cx&, ByVal cy&, ByVal wFlags&) As Long




Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Declare Function GetSystemMenu Lib "user32" ( _
    ByVal hWnd As Long, ByVal bRevert As Long) As Long

Private Declare Function GetMenuItemInfo Lib "user32" Alias _
    "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SetMenuItemInfo Lib "user32" Alias _
    "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, _
    ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long

Private Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

Private Declare Function IsWindow Lib "user32" _
    (ByVal hWnd As Long) As Long
    

Private Const MAX_PATH = 260
Private Const MAX_FILE = 260

Private Type OPENFILENAME
    lStructSize As Long          ' Filled with UDT size
    hWndOwner As Long            ' Tied to Owner
    hInstance As Long            ' Ignored (used only by templates)
    lpstrFilter As String        ' Tied to Filter
    lpstrCustomFilter As String  ' Ignored (exercise for reader)
    nMaxCustFilter As Long       ' Ignored (exercise for reader)
    nFilterIndex As Long         ' Tied to FilterIndex
    lpstrFile As String          ' Tied to FileName
    nMaxFile As Long             ' Handled internally
    lpstrFileTitle As String     ' Tied to FileTitle
    nMaxFileTitle As Long        ' Handled internally
    lpstrInitialDir As String    ' Tied to InitDir
    lpstrTitle As String         ' Tied to DlgTitle
    flags As Long                ' Tied to Flags
    nFileOffset As Integer       ' Ignored (exercise for reader)
    nFileExtension As Integer    ' Ignored (exercise for reader)
    lpstrDefExt As String        ' Tied to DefaultExt
    lCustData As Long            ' Ignored (needed for hooks)
    lpfnHook As Long             ' Ignored (good luck with hooks)
    lpTemplateName As Long       ' Ignored (good luck with templates)
End Type

Private Declare Function GetOpenFileName Lib "COMDLG32" _
    Alias "GetOpenFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" _
    Alias "GetSaveFileNameA" (file As OPENFILENAME) As Long
Private Declare Function GetFileTitle Lib "COMDLG32" _
    Alias "GetFileTitleA" (ByVal szFile As String, _
    ByVal szTitle As String, ByVal cbBuf As Long) As Long


Public Type BROWSEINFO
  hOwner As Long
  pidlRoot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lParam As Long
  iImage As Long
End Type
'BROWSEINFO.ulFlags values:
Public Enum BrowserFlags
    BIF_RETURNONLYFSDIRS = &H1
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_STATUSTEXT = &H4
    BIF_RETURNFSANCESTORS = &H8
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
End Enum

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
Alias "SHGetPathFromIDListA" (ByVal pidl As Long, _
ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Public Enum EOpenFile
    OFN_READONLY = &H1
    OFN_OVERWRITEPROMPT = &H2
    OFN_HIDEREADONLY = &H4
    OFN_NOCHANGEDIR = &H8
    OFN_SHOWHELP = &H10
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_NOVALIDATE = &H100
    OFN_ALLOWMULTISELECT = &H200
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_PATHMUSTEXIST = &H800
    OFN_FILEMUSTEXIST = &H1000
    OFN_CREATEPROMPT = &H2000
    OFN_SHAREAWARE = &H4000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOLONGNAMES = &H40000
    OFN_EXPLORER = &H80000
    OFN_NODEREFERENCELINKS = &H100000
    OFN_LONGNAMES = &H200000
End Enum

Private Type TCHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As Long
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As Long
End Type

Private Declare Function ChooseColor Lib "COMDLG32.DLL" _
    Alias "ChooseColorA" (Color As TCHOOSECOLOR) As Long

Public Enum EChooseColor
    CC_RGBInit = &H1
    CC_FullOpen = &H2
    CC_PreventFullOpen = &H4
    CC_ColorShowHelp = &H8
' Win95 only
    CC_SolidColor = &H80
    CC_AnyColor = &H100
' End Win95 only
    CC_ENABLEHOOK = &H10
    CC_ENABLETEMPLATE = &H20
    CC_EnableTemplateHandle = &H40
End Enum
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Type TCHOOSEFONT
    lStructSize As Long         ' Filled with UDT size
    hWndOwner As Long           ' Caller's window handle
    hdc As Long                 ' Printer DC/IC or NULL
    lpLogFont As Long           ' Pointer to LOGFONT
    iPointSize As Long          ' 10 * size in points of font
    flags As Long               ' Type flags
    rgbColors As Long           ' Returned text color
    lCustData As Long           ' Data passed to hook function
    lpfnHook As Long            ' Pointer to hook function
    lpTemplateName As Long      ' Custom template name
    hInstance As Long           ' Instance handle for template
    lpszStyle As String         ' Return style field
    nFontType As Integer        ' Font type bits
    iAlign As Integer           ' Filler
    nSizeMin As Long            ' Minimum point size allowed
    nSizeMax As Long            ' Maximum point size allowed
End Type
Private Declare Function ChooseFont Lib "COMDLG32" _
    Alias "ChooseFontA" (chfont As TCHOOSEFONT) As Long

Private Const LF_FACESIZE = 32
Private Type LOGFONT
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

Public Enum EChooseFont
    CF_ScreenFonts = &H1
    CF_PrinterFonts = &H2
    CF_BOTH = &H3
    CF_FontShowHelp = &H4
    CF_UseStyle = &H80
    CF_EFFECTS = &H100
    CF_AnsiOnly = &H400
    CF_NoVectorFonts = &H800
    CF_NoOemFonts = CF_NoVectorFonts
    CF_NoSimulations = &H1000
    CF_LimitSize = &H2000
    CF_FixedPitchOnly = &H4000
    CF_WYSIWYG = &H8000  ' Must also have ScreenFonts And PrinterFonts
    CF_ForceFontExist = &H10000
    CF_ScalableOnly = &H20000
    CF_TTOnly = &H40000
    CF_NoFaceSel = &H80000
    CF_NoStyleSel = &H100000
    CF_NoSizeSel = &H200000
    ' Win95 only
    CF_SelectScript = &H400000
    CF_NoScriptSel = &H800000
    CF_NoVertFonts = &H1000000

    CF_InitToLogFontStruct = &H40
    CF_Apply = &H200
    CF_EnableHook = &H8
    CF_EnableTemplate = &H10
    CF_EnableTemplateHandle = &H20
    CF_FontNotSupported = &H238
End Enum

' These are extra nFontType bits that are added to what is returned to the
' EnumFonts callback routine

Public Enum EFontType
    Simulated_FontType = &H8000
    Printer_FontType = &H4000
    Screen_FontType = &H2000
    Bold_FontType = &H100
    Italic_FontType = &H200
    Regular_FontType = &H400
End Enum

Private Type TPRINTDLG
    lStructSize As Long
    hWndOwner As Long
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
    lpPrintTemplateName As Long
    lpSetupTemplateName As Long
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type

'  DEVMODE collation selections
Private Const DMCOLLATE_FALSE = 0
Private Const DMCOLLATE_TRUE = 1

Private Declare Function PrintDlg Lib "COMDLG32.DLL" _
    Alias "PrintDlgA" (prtdlg As TPRINTDLG) As Integer

Public Enum EPrintDialog
    PD_ALLPAGES = &H0
    PD_SELECTION = &H1
    PD_PAGENUMS = &H2
    PD_NOSELECTION = &H4
    PD_NOPAGENUMS = &H8
    PD_COLLATE = &H10
    PD_PRINTTOFILE = &H20
    PD_PRINTSETUP = &H40
    PD_NOWARNING = &H80
    PD_RETURNDC = &H100
    PD_RETURNIC = &H200
    PD_RETURNDEFAULT = &H400
    PD_SHOWHELP = &H800
    PD_ENABLEPRINTHOOK = &H1000
    PD_ENABLESETUPHOOK = &H2000
    PD_ENABLEPRINTTEMPLATE = &H4000
    PD_ENABLESETUPTEMPLATE = &H8000
    PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
    PD_ENABLESETUPTEMPLATEHANDLE = &H20000
    PD_USEDEVMODECOPIES = &H40000
    PD_USEDEVMODECOPIESANDCOLLATE = &H40000
    PD_DISABLEPRINTTOFILE = &H80000
    PD_HIDEPRINTTOFILE = &H100000
    PD_NONETWORKBUTTON = &H200000
End Enum

Private Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type

Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32
Private Type DevMode
    dmDeviceName As String * CCHDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCHFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

' New Win95 Page Setup dialogs are up to you
Private Type POINTL
    x As Long
    y As Long
End Type
Private Type RECT
    Left As Long
    TOp As Long
    Right As Long
    Bottom As Long
End Type


Private Type TPAGESETUPDLG
    lStructSize                 As Long
    hWndOwner                   As Long
    hDevMode                    As Long
    hDevNames                   As Long
    flags                       As Long
    ptPaperSize                 As POINTL
    rtMinMargin                 As RECT
    rtMargin                    As RECT
    hInstance                   As Long
    lCustData                   As Long
    lpfnPageSetupHook           As Long
    lpfnPagePaintHook           As Long
    lpPageSetupTemplateName     As Long
    hPageSetupTemplate          As Long
End Type

' EPaperSize constants same as vbPRPS constants
Public Enum EPaperSize
    epsLetter = 1          ' Letter, 8 1/2 x 11 in.
    epsLetterSmall         ' Letter Small, 8 1/2 x 11 in.
    epsTabloid             ' Tabloid, 11 x 17 in.
    epsLedger              ' Ledger, 17 x 11 in.
    epsLegal               ' Legal, 8 1/2 x 14 in.
    epsStatement           ' Statement, 5 1/2 x 8 1/2 in.
    epsExecutive           ' Executive, 7 1/2 x 10 1/2 in.
    epsA3                  ' A3, 297 x 420 mm
    epsA4                  ' A4, 210 x 297 mm
    epsA4Small             ' A4 Small, 210 x 297 mm
    epsA5                  ' A5, 148 x 210 mm
    epsB4                  ' B4, 250 x 354 mm
    epsB5                  ' B5, 182 x 257 mm
    epsFolio               ' Folio, 8 1/2 x 13 in.
    epsQuarto              ' Quarto, 215 x 275 mm
    eps10x14               ' 10 x 14 in.
    eps11x17               ' 11 x 17 in.
    epsNote                ' Note, 8 1/2 x 11 in.
    epsEnv9                ' Envelope #9, 3 7/8 x 8 7/8 in.
    epsEnv10               ' Envelope #10, 4 1/8 x 9 1/2 in.
    epsEnv11               ' Envelope #11, 4 1/2 x 10 3/8 in.
    epsEnv12               ' Envelope #12, 4 1/2 x 11 in.
    epsEnv14               ' Envelope #14, 5 x 11 1/2 in.
    epsCSheet              ' C size sheet
    epsDSheet              ' D size sheet
    epsESheet              ' E size sheet
    epsEnvDL               ' Envelope DL, 110 x 220 mm
    epsEnvC3               ' Envelope C3, 324 x 458 mm
    epsEnvC4               ' Envelope C4, 229 x 324 mm
    epsEnvC5               ' Envelope C5, 162 x 229 mm
    epsEnvC6               ' Envelope C6, 114 x 162 mm
    epsEnvC65              ' Envelope C65, 114 x 229 mm
    epsEnvB4               ' Envelope B4, 250 x 353 mm
    epsEnvB5               ' Envelope B5, 176 x 250 mm
    epsEnvB6               ' Envelope B6, 176 x 125 mm
    epsEnvItaly            ' Envelope, 110 x 230 mm
    epsenvmonarch          ' Envelope Monarch, 3 7/8 x 7 1/2 in.
    epsEnvPersonal         ' Envelope, 3 5/8 x 6 1/2 in.
    epsFanfoldUS           ' U.S. Standard Fanfold, 14 7/8 x 11 in.
    epsFanfoldStdGerman    ' German Standard Fanfold, 8 1/2 x 12 in.
    epsFanfoldLglGerman    ' German Legal Fanfold, 8 1/2 x 13 in.
    epsUser = 256          ' User-defined
End Enum

' EPrintQuality constants same as vbPRPQ constants
Public Enum EPrintQuality
    epqDraft = -1
    epqLow = -2
    epqMedium = -3
    epqHigh = -4
End Enum

Public Enum EOrientation
    eoPortrait = 1
    eoLandscape
End Enum

Private Declare Function PageSetupDlg Lib "COMDLG32" _
    Alias "PageSetupDlgA" (lppage As TPAGESETUPDLG) As Boolean

Public Enum EPageSetup
    PSD_Defaultminmargins = &H0 ' Default (printer's)
    PSD_InWinIniIntlMeasure = &H0
    PSD_MINMARGINS = &H1
    PSD_MARGINS = &H2
    PSD_INTHOUSANDTHSOFINCHES = &H4
    PSD_INHUNDREDTHSOFMILLIMETERS = &H8
    PSD_DISABLEMARGINS = &H10
    PSD_DISABLEPRINTER = &H20
    PSD_NoWarning = &H80
    PSD_DISABLEORIENTATION = &H100
    PSD_ReturnDefault = &H400
    PSD_DISABLEPAPER = &H200
    PSD_ShowHelp = &H800
    PSD_EnablePageSetupHook = &H2000
    PSD_EnablePageSetupTemplate = &H8000
    PSD_EnablePageSetupTemplateHandle = &H20000
    PSD_EnablePagePaintHook = &H40000
    PSD_DisablePagePainting = &H80000
End Enum

Public Enum EPageSetupUnits
    epsuInches
    epsuMillimeters
End Enum

' Common dialog errors

Private Declare Function CommDlgExtendedError Lib "COMDLG32" () As Long

Public Enum EDialogError
    CDERR_DIALOGFAILURE = &HFFFF

    CDERR_GENERALCODES = &H0
    CDERR_STRUCTSIZE = &H1
    CDERR_INITIALIZATION = &H2
    CDERR_NOTEMPLATE = &H3
    CDERR_NOHINSTANCE = &H4
    CDERR_LOADSTRFAILURE = &H5
    CDERR_FINDRESFAILURE = &H6
    CDERR_LOADRESFAILURE = &H7
    CDERR_LOCKRESFAILURE = &H8
    CDERR_MEMALLOCFAILURE = &H9
    CDERR_MEMLOCKFAILURE = &HA
    CDERR_NOHOOK = &HB
    CDERR_REGISTERMSGFAIL = &HC

    PDERR_PRINTERCODES = &H1000
    PDERR_SETUPFAILURE = &H1001
    PDERR_PARSEFAILURE = &H1002
    PDERR_RETDEFFAILURE = &H1003
    PDERR_LOADDRVFAILURE = &H1004
    PDERR_GETDEVMODEFAIL = &H1005
    PDERR_INITFAILURE = &H1006
    PDERR_NODEVICES = &H1007
    PDERR_NODEFAULTPRN = &H1008
    PDERR_DNDMMISMATCH = &H1009
    PDERR_CREATEICFAILURE = &H100A
    PDERR_PRINTERNOTFOUND = &H100B
    PDERR_DEFAULTDIFFERENT = &H100C

    CFERR_CHOOSEFONTCODES = &H2000
    CFERR_NOFONTS = &H2001
    CFERR_MAXLESSTHANMIN = &H2002

    FNERR_FILENAMECODES = &H3000
    FNERR_SUBCLASSFAILURE = &H3001
    FNERR_INVALIDFILENAME = &H3002
    FNERR_BUFFERTOOSMALL = &H3003

    CCERR_CHOOSECOLORCODES = &H5000
End Enum

' Array of custom colors lasts for life of app
Private alCustom(0 To 15) As Long, fNotFirst As Boolean

Public Enum EPrintRange
    eprAll
    eprPageNumbers
    eprSelection
End Enum
Private m_lApiReturn As Long
Private m_lExtendedError As Long
Private m_dvmode As DevMode


Private Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long


Private Const CB_ERR = -1, CB_SELECTSTRING = &H14D, CB_SHOWDROPDOWN = &H14F, CBN_SELENDOK = 9

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function


'call this function in KeyPress event method
Public Function AutoMatchCBBox(ByRef cbBox As Object, ByVal KeyAscii As Integer) As Integer
    
        
    Dim strFindThis As String, bContinueSearch As Boolean
    Dim lResult As Long, lStart As Long, lLength As Long
    AutoMatchCBBox = 0 ' block cbBox since we handle everything
    bContinueSearch = True
    lStart = cbBox.SelStart
    lLength = cbBox.SelLength

    On Error GoTo ErrHandle
        
    If KeyAscii < 32 Then 'control char
        bContinueSearch = False
        cbBox.SelLength = 0 'select nothing since we will delete/enter
        If KeyAscii = Asc(vbBack) Then 'take care BackSpace and Delete first
            If lLength = 0 Then 'delete last char
                If Len(cbBox) > 0 Then ' in case user delete empty cbBox
                    cbBox.Text = Left(cbBox.Text, Len(cbBox) - 1)
                End If
            Else 'leave unselected char(s) and delete rest of text
                cbBox.Text = Left(cbBox.Text, lStart)
            End If
            cbBox.SelStart = Len(cbBox) 'set insertion position @ the end of string
        ElseIf KeyAscii = vbKeyReturn Then  'user select this string
            cbBox.SelStart = Len(cbBox)
            lResult = SendMessage(cbBox.hWnd, CBN_SELENDOK, 0, 0)
            AutoMatchCBBox = KeyAscii 'let caller a chance to handle "Enter"
        End If
    Else 'generate searching string
        If lLength = 0 Then
            strFindThis = cbBox.Text & Chr(KeyAscii) 'No selection, append it
        Else
            strFindThis = Left(cbBox.Text, lStart) & Chr(KeyAscii)
        End If
    End If
    
    If bContinueSearch Then 'need to search
        Call VBComBoBoxDroppedDown(cbBox)  'open dropdown list
        lResult = SendMessage(cbBox.hWnd, CB_SELECTSTRING, -1, ByVal strFindThis)
        If lResult = CB_ERR Then 'not found
            cbBox.Text = strFindThis 'set cbBox as whatever it is
            cbBox.SelLength = 0 'no selected char(s) since not found
            cbBox.SelStart = Len(cbBox) 'set insertion position @ the end of string
        Else
            'found string, highlight rest of string for user
            cbBox.SelStart = Len(strFindThis)
            cbBox.SelLength = Len(cbBox) - cbBox.SelStart
        End If
    End If
    On Error GoTo 0
    Exit Function
    
ErrHandle:
    'got problem, simply return whatever pass in
    Debug.Print "Failed: AutoCompleteComboBox due to : " & Err.Description
    Debug.Assert False
    AutoMatchCBBox = KeyAscii
    On Error GoTo 0
End Function

'open dorpdown list
Private Sub VBComBoBoxDroppedDown(ByRef cbBox As ComboBox)
    Call SendMessage(cbBox.hWnd, CB_SHOWDROPDOWN, Abs(True), 0)
End Sub


    
Public Function GetName(ByVal pObjSys As SysParameter) As String
    Dim pStrName As String
    Dim plngLength As Long
    pStrName = String(255, " ")
    plngLength = Len(pStrName)
    If (pObjSys = Computer) Then
         pStrName = Environ$("COMPUTERNAME")
         If (pStrName = "") Then
            Call GetComputerName(pStrName, plngLength)
         End If
    ElseIf (pObjSys = User) Then
        'This Will return Thread Specific User Name
        'Environ$("USERNAME")
        Call GetUserName(pStrName, plngLength)
    Else
        pStrName = Environ$("USERDOMAIN")
    End If
    pStrName = Left(pStrName, plngLength)
    GetName = Replace(pStrName, Chr(0), "")
End Function


'------------------------Common Dialog API---------------------
Public Property Get APIReturn() As Long
    'return object's APIReturn property
    APIReturn = m_lApiReturn
End Property
Public Property Get ExtendedError() As Long
    'return object's ExtendedError property
    ExtendedError = m_lExtendedError
End Property



Public Function VBGetOpenFileName(Filename As String, _
                           Optional FileTitle As String, _
                           Optional FileMustExist As Boolean = True, _
                           Optional MultiSelect As Boolean = False, _
                           Optional ReadOnly As Boolean = False, _
                           Optional HideReadOnly As Boolean = False, _
                           Optional Filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional flags As Long = 0) As Boolean

    Dim opfile As OPENFILENAME, s As String, afFlags As Long
    
    m_lApiReturn = 0
    m_lExtendedError = 0

With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    
    .flags = (-FileMustExist * OFN_FILEMUSTEXIST) Or _
            (-MultiSelect * OFN_ALLOWMULTISELECT) Or _
             (-ReadOnly * OFN_READONLY) Or _
             (-HideReadOnly * OFN_HIDEREADONLY) Or _
             (flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hWndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    ' To make Windows-style filter, replace | and : with nulls
    Dim ch As String, i As Integer
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    s = Filename & String$(MAX_PATH - Len(Filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    ' All other fields set to zero
    
    m_lApiReturn = GetOpenFileName(opfile)
    Select Case m_lApiReturn
    Case 1
        ' Success
        VBGetOpenFileName = True
        Filename = StrZToStr(.lpstrFile)
        FileTitle = StrZToStr(.lpstrFileTitle)
        flags = .flags
        ' Return the filter index
        FilterIndex = .nFilterIndex
        ' Look up the filter the user selected and return that
        Filter = FilterLookup(.lpstrFilter, FilterIndex)
        If (.flags And OFN_READONLY) Then ReadOnly = True
    Case 0
        ' Cancelled
        VBGetOpenFileName = False
        Filename = ""
        FileTitle = ""
        flags = 0
        FilterIndex = -1
        Filter = ""
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBGetOpenFileName = False
        Filename = ""
        FileTitle = ""
        flags = 0
        FilterIndex = -1
        Filter = ""
    End Select
End With
End Function
Private Function StrZToStr(s As String) As String
    StrZToStr = Left$(s, lstrlen(s))
End Function

Function VBGetSaveFileName(Filename As String, _
                           Optional FileTitle As String, _
                           Optional OverWritePrompt As Boolean = True, _
                           Optional Filter As String = "All (*.*)| *.*", _
                           Optional FilterIndex As Long = 1, _
                           Optional InitDir As String, _
                           Optional DlgTitle As String, _
                           Optional DefaultExt As String, _
                           Optional Owner As Long = -1, _
                           Optional flags As Long) As Boolean
            
    Dim opfile As OPENFILENAME, s As String

    m_lApiReturn = 0
    m_lExtendedError = 0

With opfile
    .lStructSize = Len(opfile)
    
    ' Add in specific flags and strip out non-VB flags
    .flags = (-OverWritePrompt * OFN_OVERWRITEPROMPT) Or _
             OFN_HIDEREADONLY Or _
             (flags And CLng(Not (OFN_ENABLEHOOK Or _
                                  OFN_ENABLETEMPLATE)))
    ' Owner can take handle of owning window
    If Owner <> -1 Then .hWndOwner = Owner
    ' InitDir can take initial directory string
    .lpstrInitialDir = InitDir
    ' DefaultExt can take default extension
    .lpstrDefExt = DefaultExt
    ' DlgTitle can take dialog box title
    .lpstrTitle = DlgTitle
    
    ' Make new filter with bars (|) replacing nulls and double null at end
    Dim ch As String, i As Integer
    For i = 1 To Len(Filter)
        ch = Mid$(Filter, i, 1)
        If ch = "|" Or ch = ":" Then
            s = s & vbNullChar
        Else
            s = s & ch
        End If
    Next
    ' Put double null at end
    s = s & vbNullChar & vbNullChar
    .lpstrFilter = s
    .nFilterIndex = FilterIndex

    ' Pad file and file title buffers to maximum path
    s = Filename & String$(MAX_PATH - Len(Filename), 0)
    .lpstrFile = s
    .nMaxFile = MAX_PATH
    s = FileTitle & String$(MAX_FILE - Len(FileTitle), 0)
    .lpstrFileTitle = s
    .nMaxFileTitle = MAX_FILE
    ' All other fields zero
    
    m_lApiReturn = GetSaveFileName(opfile)
    Select Case m_lApiReturn
    Case 1
        VBGetSaveFileName = True
        Filename = StrZToStr(.lpstrFile)
        FileTitle = StrZToStr(.lpstrFileTitle)
        flags = .flags
        ' Return the filter index
        FilterIndex = .nFilterIndex
        ' Look up the filter the user selected and return that
        Filter = FilterLookup(.lpstrFilter, FilterIndex)
    Case 0
        ' Cancelled:
        VBGetSaveFileName = False
        Filename = ""
        FileTitle = ""
        flags = 0
        FilterIndex = 0
        Filter = ""
    Case Else
        ' Extended error:
        VBGetSaveFileName = False
        m_lExtendedError = CommDlgExtendedError()
        Filename = ""
        FileTitle = ""
        flags = 0
        FilterIndex = 0
        Filter = ""
    End Select
End With
End Function

Private Function FilterLookup(ByVal sFilters As String, ByVal iCur As Long) As String
    Dim iStart As Long, iEnd As Long, s As String
    iStart = 1
    If sFilters = "" Then Exit Function
    Do
        ' Cut out both parts marked by null character
        iEnd = InStr(iStart, sFilters, vbNullChar)
        If iEnd = 0 Then Exit Function
        iEnd = InStr(iEnd + 1, sFilters, vbNullChar)
        If iEnd Then
            s = Mid$(sFilters, iStart, iEnd - iStart)
        Else
            s = Mid$(sFilters, iStart)
        End If
        iStart = iEnd + 1
        If iCur = 1 Then
            FilterLookup = s
            Exit Function
        End If
        iCur = iCur - 1
    Loop While iCur
End Function

Function VBGetFileTitle(sFile As String) As String
    Dim sFileTitle As String, cFileTitle As Integer

    cFileTitle = MAX_PATH
    sFileTitle = String$(MAX_PATH, 0)
    cFileTitle = GetFileTitle(sFile, sFileTitle, MAX_PATH)
    If cFileTitle Then
        VBGetFileTitle = ""
    Else
        VBGetFileTitle = Left$(sFileTitle, InStr(sFileTitle, vbNullChar) - 1)
    End If

End Function

' ChooseColor wrapper
Function VBChooseColor(Color As Long, _
                       Optional AnyColor As Boolean = True, _
                       Optional FullOpen As Boolean = False, _
                       Optional DisableFullOpen As Boolean = False, _
                       Optional Owner As Long = -1, _
                       Optional flags As Long) As Boolean

    Dim chclr As TCHOOSECOLOR
    
    chclr.lStructSize = Len(chclr)
    
    ' Color must get reference variable to receive result
    ' Flags can get reference variable or constant with bit flags
    ' Owner can take handle of owning window
    If Owner <> -1 Then chclr.hWndOwner = Owner

    ' Assign color (default uninitialized value of zero is good default)
    chclr.rgbResult = Color

    ' Mask out unwanted bits
    Dim afMask As Long
    afMask = CLng(Not (CC_ENABLEHOOK Or _
                       CC_ENABLETEMPLATE))
    ' Pass in flags
    chclr.flags = afMask And (CC_RGBInit Or _
                  IIf(AnyColor, CC_AnyColor, CC_SolidColor) Or _
                  (-FullOpen * CC_FullOpen) Or _
                  (-DisableFullOpen * CC_PreventFullOpen))

    ' If first time, initialize to white
    If fNotFirst = False Then InitColors

    chclr.lpCustColors = VarPtr(alCustom(0))
    ' All other fields zero
    
    m_lApiReturn = ChooseColor(chclr)
    Select Case m_lApiReturn
    Case 1
        ' Success
        VBChooseColor = True
        Color = chclr.rgbResult
    Case 0
        ' Cancelled
        VBChooseColor = False
        Color = -1
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBChooseColor = False
        Color = -1
    End Select

End Function

Private Sub InitColors()
    Dim i As Integer
    ' Initialize with first 16 system interface colors
    For i = 0 To 15
        alCustom(i) = GetSysColor(i)
    Next
    fNotFirst = True
End Sub

' Property to read or modify custom colors (use to save colors in registry)
Public Property Get CustomColor(i As Integer) As Long
    ' If first time, initialize to white
    If fNotFirst = False Then InitColors
    If i >= 0 And i <= 15 Then
        CustomColor = alCustom(i)
    Else
        CustomColor = -1
    End If
End Property

Public Property Let CustomColor(i As Integer, iValue As Long)
    ' If first time, initialize to system colors
    If fNotFirst = False Then InitColors
    If i >= 0 And i <= 15 Then
        alCustom(i) = iValue
    End If
End Property

' ChooseFont wrapper
Function VBChooseFont(CurFont As Font, _
                      Optional PrinterDC As Long = -1, _
                      Optional Owner As Long = -1, _
                      Optional Color As Long = vbBlack, _
                      Optional MinSize As Long = 0, _
                      Optional MaxSize As Long = 0, _
                      Optional flags As Long = 0) As Boolean

    m_lApiReturn = 0
    m_lExtendedError = 0

    ' Unwanted Flags bits
    Const CF_FontNotSupported = CF_Apply Or CF_EnableHook Or CF_EnableTemplate
    
    ' Flags can get reference variable or constant with bit flags
    ' PrinterDC can take printer DC
    If PrinterDC = -1 Then
        PrinterDC = 0
        If flags And CF_PrinterFonts Then PrinterDC = Printer.hdc
    Else
        flags = flags Or CF_PrinterFonts
    End If
    ' Must have some fonts
    If (flags And CF_PrinterFonts) = 0 Then flags = flags Or CF_ScreenFonts
    ' Color can take initial color, receive chosen color
    If Color <> vbBlack Then flags = flags Or CF_EFFECTS
    ' MinSize can be minimum size accepted
    If MinSize Then flags = flags Or CF_LimitSize
    ' MaxSize can be maximum size accepted
    If MaxSize Then flags = flags Or CF_LimitSize

    ' Put in required internal flags and remove unsupported
    flags = (flags Or CF_InitToLogFontStruct) And Not CF_FontNotSupported
    
    ' Initialize LOGFONT variable
    Dim fnt As LOGFONT
    Const PointsPerTwip = 1440 / 72
    fnt.lfHeight = -(CurFont.Size * (PointsPerTwip / Screen.TwipsPerPixelY))
    fnt.lfWeight = CurFont.Weight
    fnt.lfItalic = CurFont.Italic
    fnt.lfUnderline = CurFont.Underline
    fnt.lfStrikeOut = CurFont.Strikethrough
    ' Other fields zero
    StrToBytes fnt.lfFaceName, CurFont.Name

    ' Initialize TCHOOSEFONT variable
    Dim cf As TCHOOSEFONT
    cf.lStructSize = Len(cf)
    If Owner <> -1 Then cf.hWndOwner = Owner
    cf.hdc = PrinterDC
    cf.lpLogFont = VarPtr(fnt)
    cf.iPointSize = CurFont.Size * 10
    cf.flags = flags
    cf.rgbColors = Color
    cf.nSizeMin = MinSize
    cf.nSizeMax = MaxSize
    
    ' All other fields zero
    m_lApiReturn = ChooseFont(cf)
    Select Case m_lApiReturn
    Case 1
        ' Success
        VBChooseFont = True
        flags = cf.flags
        Color = cf.rgbColors
        CurFont.Bold = cf.nFontType And Bold_FontType
        'CurFont.Italic = cf.nFontType And Italic_FontType
        CurFont.Italic = fnt.lfItalic
        CurFont.Strikethrough = fnt.lfStrikeOut
        CurFont.Underline = fnt.lfUnderline
        CurFont.Weight = fnt.lfWeight
        CurFont.Size = cf.iPointSize / 10
        CurFont.Name = BytesToStr(fnt.lfFaceName)
    Case 0
        ' Cancelled
        VBChooseFont = False
    Case Else
        ' Extended error
        m_lExtendedError = CommDlgExtendedError()
        VBChooseFont = False
    End Select
        
End Function

' PrintDlg wrapper
Function VBPrintDlg(hdc As Long, _
                    Optional PrintRange As EPrintRange = eprAll, _
                    Optional DisablePageNumbers As Boolean, _
                    Optional FromPage As Long = 1, _
                    Optional ToPage As Long = &HFFFF, _
                    Optional DisableSelection As Boolean, _
                    Optional Copies As Integer, _
                    Optional ShowPrintToFile As Boolean, _
                    Optional DisablePrintToFile As Boolean = True, _
                    Optional PrintToFile As Boolean, _
                    Optional Collate As Boolean, _
                    Optional PreventWarning As Boolean, _
                    Optional Owner As Long, _
                    Optional Printer As Object, _
                    Optional flags As Long) As Boolean
    Dim afFlags As Long, afMask As Long
    
    m_lApiReturn = 0
    m_lExtendedError = 0
    
    ' Set PRINTDLG flags
    afFlags = (-DisablePageNumbers * PD_NOPAGENUMS) Or _
              (-DisablePrintToFile * PD_DISABLEPRINTTOFILE) Or _
              (-DisableSelection * PD_NOSELECTION) Or _
              (-PrintToFile * PD_PRINTTOFILE) Or _
              (-(Not ShowPrintToFile) * PD_HIDEPRINTTOFILE) Or _
              (-PreventWarning * PD_NOWARNING) Or _
              (-Collate * PD_COLLATE) Or _
              PD_USEDEVMODECOPIESANDCOLLATE Or _
              PD_RETURNDC
    If PrintRange = eprPageNumbers Then
        afFlags = afFlags Or PD_PAGENUMS
    ElseIf PrintRange = eprSelection Then
        afFlags = afFlags Or PD_SELECTION
    End If
    ' Mask out unwanted bits
    afMask = CLng(Not (PD_ENABLEPRINTHOOK Or _
                       PD_ENABLEPRINTTEMPLATE))
    afMask = afMask And _
             CLng(Not (PD_ENABLESETUPHOOK Or _
                       PD_ENABLESETUPTEMPLATE))
    
    ' Fill in PRINTDLG structure
    Dim pd As TPRINTDLG
    pd.lStructSize = Len(pd)
    pd.hWndOwner = Owner
    pd.flags = afFlags And afMask
    pd.nFromPage = FromPage
    pd.nToPage = ToPage
    pd.nMinPage = 1
    pd.nMaxPage = &HFFFF
    
    ' Show Print dialog
    m_lApiReturn = PrintDlg(pd)
    Select Case m_lApiReturn
    Case 1
        VBPrintDlg = True
        ' Return dialog values in parameters
        hdc = pd.hdc
        If (pd.flags And PD_PAGENUMS) Then
            PrintRange = eprPageNumbers
        ElseIf (pd.flags And PD_SELECTION) Then
            PrintRange = eprSelection
        Else
            PrintRange = eprAll
        End If
        FromPage = pd.nFromPage
        ToPage = pd.nToPage
        PrintToFile = (pd.flags And PD_PRINTTOFILE)
        ' Get DEVMODE structure from PRINTDLG
        Dim pDevMode As Long
        pDevMode = GlobalLock(pd.hDevMode)
        CopyMemory m_dvmode, ByVal pDevMode, Len(m_dvmode)
        Call GlobalUnlock(pd.hDevMode)
        ' Get Copies and Collate settings from DEVMODE structure
        Copies = m_dvmode.dmCopies
        Collate = (m_dvmode.dmCollate = DMCOLLATE_TRUE)
                
        ' Set default printer properties
        On Error Resume Next
        If Not (Printer Is Nothing) Then
            Printer.Copies = Copies
            Printer.Orientation = m_dvmode.dmOrientation
            Printer.PaperSize = m_dvmode.dmPaperSize
            Printer.PrintQuality = m_dvmode.dmPrintQuality
        End If
        On Error GoTo 0
    Case 0
        ' Cancelled
        VBPrintDlg = False
    Case Else
        ' Extended error:
        m_lExtendedError = CommDlgExtendedError()
        VBPrintDlg = False
    End Select
    
End Function
Private Property Get DevMode() As DevMode
    DevMode = m_dvmode
End Property

' PageSetupDlg wrapper
Function VBPageSetupDlg(Optional Owner As Long, _
                        Optional DisableMargins As Boolean, _
                        Optional DisableOrientation As Boolean, _
                        Optional DisablePaper As Boolean, _
                        Optional DisablePrinter As Boolean, _
                        Optional LeftMargin As Long, _
                        Optional MinLeftMargin As Long, _
                        Optional RightMargin As Long, _
                        Optional MinRightMargin As Long, _
                        Optional TopMargin As Long, _
                        Optional MinTopMargin As Long, _
                        Optional BottomMargin As Long, _
                        Optional MinBottomMargin As Long, _
                        Optional PaperSize As EPaperSize = epsLetter, _
                        Optional Orientation As EOrientation = eoPortrait, _
                        Optional PrintQuality As EPrintQuality = epqDraft, _
                        Optional Units As EPageSetupUnits = epsuInches, _
                        Optional Printer As Object, _
                        Optional flags As Long) As Boolean
    Dim afFlags As Long, afMask As Long
    
    m_lApiReturn = 0
    m_lExtendedError = 0
    ' Mask out unwanted bits
    afMask = Not (PSD_EnablePagePaintHook Or _
                  PSD_EnablePageSetupHook Or _
                  PSD_EnablePageSetupTemplate)
    ' Set TPAGESETUPDLG flags
    afFlags = (-DisableMargins * PSD_DISABLEMARGINS) Or _
              (-DisableOrientation * PSD_DISABLEORIENTATION) Or _
              (-DisablePaper * PSD_DISABLEPAPER) Or _
              (-DisablePrinter * PSD_DISABLEPRINTER) Or _
              PSD_MARGINS Or PSD_MINMARGINS And afMask
    Dim lUnits As Long
    If Units = epsuInches Then
        afFlags = afFlags Or PSD_INTHOUSANDTHSOFINCHES
        lUnits = 1000
    Else
        afFlags = afFlags Or PSD_INHUNDREDTHSOFMILLIMETERS
        lUnits = 100
    End If
    
    Dim psd As TPAGESETUPDLG
    ' Fill in PRINTDLG structure
    psd.lStructSize = Len(psd)
    psd.hWndOwner = Owner
    psd.rtMargin.TOp = TopMargin * lUnits
    psd.rtMargin.Left = LeftMargin * lUnits
    psd.rtMargin.Bottom = BottomMargin * lUnits
    psd.rtMargin.Right = RightMargin * lUnits
    psd.rtMinMargin.TOp = MinTopMargin * lUnits
    psd.rtMinMargin.Left = MinLeftMargin * lUnits
    psd.rtMinMargin.Bottom = MinBottomMargin * lUnits
    psd.rtMinMargin.Right = MinRightMargin * lUnits
    psd.flags = afFlags
    
    ' Show Print dialog
    If PageSetupDlg(psd) Then
        VBPageSetupDlg = True
        ' Return dialog values in parameters
        TopMargin = psd.rtMargin.TOp / lUnits
        LeftMargin = psd.rtMargin.Left / lUnits
        BottomMargin = psd.rtMargin.Bottom / lUnits
        RightMargin = psd.rtMargin.Right / lUnits
        MinTopMargin = psd.rtMinMargin.TOp / lUnits
        MinLeftMargin = psd.rtMinMargin.Left / lUnits
        MinBottomMargin = psd.rtMinMargin.Bottom / lUnits
        MinRightMargin = psd.rtMinMargin.Right / lUnits
        
        ' Get DEVMODE structure from PRINTDLG
        Dim dvmode As DevMode, pDevMode As Long
        pDevMode = GlobalLock(psd.hDevMode)
        CopyMemory dvmode, ByVal pDevMode, Len(dvmode)
        Call GlobalUnlock(psd.hDevMode)
        PaperSize = dvmode.dmPaperSize
        Orientation = dvmode.dmOrientation
        PrintQuality = dvmode.dmPrintQuality
        ' Set default printer properties
        On Error Resume Next
        If Not (Printer Is Nothing) Then
            Printer.Copies = dvmode.dmCopies
            Printer.Orientation = dvmode.dmOrientation
            Printer.PaperSize = dvmode.dmPaperSize
            Printer.PrintQuality = dvmode.dmPrintQuality
        End If
        On Error GoTo 0
    End If

End Function

#If fComponent = 0 Then
Private Sub ErrRaise(e As Long)
    Dim sText As String, sSource As String
    If e > 1000 Then
        sSource = App.EXEName & ".CommonDialog"
        Err.Raise COMError(e), sSource, sText
    Else
        ' Raise standard Visual Basic error
        sSource = App.EXEName & ".VBError"
        Err.Raise e, sSource
    End If
End Sub
#End If


Private Sub StrToBytes(ab() As Byte, s As String)
    If IsArrayEmpty(ab) Then
        ' Assign to empty array
        ab = StrConv(s, vbFromUnicode)
    Else
        Dim cab As Long
        ' Copy to existing array, padding or truncating if necessary
        cab = UBound(ab) - LBound(ab) + 1
        If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
        'If UnicodeTypeLib Then
        '    Dim st As String
        '    st = StrConv(s, vbFromUnicode)
        '    CopyMemoryStr ab(LBound(ab)), st, cab
        'Else
            CopyMemoryStr ab(LBound(ab)), s, cab
        'End If
    End If
End Sub


Private Function BytesToStr(ab() As Byte) As String
    BytesToStr = StrConv(ab, vbUnicode)
End Function

Private Function COMError(e As Long) As Long
    COMError = e Or vbObjectError
End Function
'
Private Function IsArrayEmpty(va As Variant) As Boolean
    Dim v As Variant
    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)
End Function


Public Function GetFolder(ByVal hWndOwner As Long, ByVal sTitle As String) As String
    Dim bInf As BROWSEINFO
    Dim retval As Long
    Dim PathID As Long
    Dim RetPath As String
    Dim Offset As Integer
    
    'Set the properties of the folder dialog
    bInf.hOwner = hWndOwner
    bInf.lpszTitle = sTitle
    bInf.ulFlags = BIF_RETURNONLYFSDIRS
    
    'Show the Browse For Folder dialog
    PathID = SHBrowseForFolder(bInf)
    RetPath = Space$(512)
    retval = SHGetPathFromIDList(ByVal PathID, ByVal RetPath)
    If retval Then
      'Trim off the null chars ending the path
      'and display the returned folder
      Offset = InStr(RetPath, Chr$(0))
      GetFolder = Left$(RetPath, Offset - 1)
    End If

End Function


Public Function ShellEx( _
        ByVal sFile As String, _
        Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
        Optional ByVal sParameters As String = "", _
        Optional ByVal sDefaultDir As String = "", _
        Optional sOperation As String = "open", _
        Optional Owner As Long = 0 _
    ) As Boolean
Dim lR As Long
Dim lErr As Long, sErr As Long

    If (InStr(UCase$(sFile), ".EXE") <> 0) Then
        eShowCmd = 0
    End If
    On Error Resume Next
    If (sParameters = "") And (sDefaultDir = "") Then
        lR = ShellExecuteForExplore(Owner, sOperation, sFile, 0, 0, essSW_SHOWNORMAL)
    Else
        lR = ShellExecute(Owner, sOperation, sFile, sParameters, sDefaultDir, eShowCmd)
    End If
    If (lR < 0) Or (lR > 32) Then
        ShellEx = True
    Else
        ' raise an appropriate error:
        lErr = vbObjectError + 1048 + lR
        Select Case lR
        Case 0
            lErr = 7: sErr = "Out of memory"
        Case ERROR_FILE_NOT_FOUND
            lErr = 53: sErr = "File not found"
        Case ERROR_PATH_NOT_FOUND
            lErr = 76: sErr = "Path not found"
        Case ERROR_BAD_FORMAT
            sErr = "The executable file is invalid or corrupt"
        Case SE_ERR_ACCESSDENIED
            lErr = 75: sErr = "Path/file access error"
        Case SE_ERR_ASSOCINCOMPLETE
            sErr = "This file type does not have a valid file association."
        Case SE_ERR_DDEBUSY
            lErr = 285: sErr = "The file could not be opened because the target application is busy.  Please try again in a moment."
        Case SE_ERR_DDEFAIL
            lErr = 285: sErr = "The file could not be opened because the DDE transaction failed.  Please try again in a moment."
        Case SE_ERR_DDETIMEOUT
            lErr = 286: sErr = "The file could not be opened due to time out.  Please try again in a moment."
        Case SE_ERR_DLLNOTFOUND
            lErr = 48: sErr = "The specified dynamic-link library was not found."
        Case SE_ERR_FNF
            lErr = 53: sErr = "File not found"
        Case SE_ERR_NOASSOC
            sErr = "No application is associated with this file type."
        Case SE_ERR_OOM
            lErr = 7: sErr = "Out of memory"
        Case SE_ERR_PNF
            lErr = 76: sErr = "Path not found"
        Case SE_ERR_SHARE
            lErr = 75: sErr = "A sharing violation occurred."
        Case Else
            sErr = "An error occurred occurred whilst trying to open or print the selected file."
        End Select
                
        Err.Raise lErr, , App.EXEName & ".GShell", sErr
        ShellEx = False
    End If

End Function



'*******************************************************************************
' Enables / Disables the close button on the titlebar and in the system menu
' of the form window passed.
'-------------------------------------------------------------------------------
' Return Values:
'
'    0  Close button state changed succesfully / nothing to do.
'   -1  Invalid Window Handle (hWnd argument) Passed to the function
'   -2  Failed to switch command ID of Close menu item in system menu
'   -3  Failed to switch enabled state of Close menu item in system menu
'
'-------------------------------------------------------------------------------
' Parameters:
'
'   hWnd    The window handle of the form whose close button is to be enabled/
'           disabled / greyed out.
'
'   Enable  True if the close button is to be enabled, or False if it is to
'           be disabled / greyed out.
'

Public Function EnableCloseButton(ByVal hWnd As Long, Enable As Boolean) _
                                                                As Integer
    Const xSC_CLOSE As Long = -10
    
    ' Check that the window handle passed is valid
    
    EnableCloseButton = -1
    If IsWindow(hWnd) = 0 Then Exit Function
    
    ' Retrieve a handle to the window's system menu
    
    Dim hMenu As Long
    hMenu = GetSystemMenu(hWnd, 0)
    
    ' Retrieve the menu item information for the close menu item/button
    
    Dim MII As MENUITEMINFO
    MII.cbSize = Len(MII)
    MII.dwTypeData = String(80, 0)
    MII.cch = Len(MII.dwTypeData)
    MII.fMask = MIIM_STATE
    
    If Enable Then
        MII.wID = xSC_CLOSE
    Else
        MII.wID = SC_CLOSE
    End If
    
    EnableCloseButton = -0
    If GetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function
    
    ' Switch the ID of the menu item so that VB can not undo the action itself
    
    Dim lngMenuID As Long
    lngMenuID = MII.wID
    
    If Enable Then
        MII.wID = SC_CLOSE
    Else
        MII.wID = xSC_CLOSE
    End If
    
    MII.fMask = MIIM_ID
    EnableCloseButton = -2
    If SetMenuItemInfo(hMenu, lngMenuID, False, MII) = 0 Then Exit Function
    
    ' Set the enabled / disabled state of the menu item
    
    If Enable Then
        MII.fState = (MII.fState Or MFS_GRAYED)
        MII.fState = MII.fState - MFS_GRAYED
    Else
        MII.fState = (MII.fState Or MFS_GRAYED)
    End If
    
    MII.fMask = MIIM_STATE
    EnableCloseButton = -3
    If SetMenuItemInfo(hMenu, MII.wID, False, MII) = 0 Then Exit Function
    
    ' Activate the non-client area of the window to update the titlebar, and
    ' draw the close button in its new state.
    
    SendMessage hWnd, WM_NCACTIVATE, True, 0
    
    EnableCloseButton = 0
    
End Function


'Get File Name From Path
Public Function GetFileName(ByVal pPath As String) As String
    Dim pName As String
    Dim pLoc As Long
    
    pLoc = InStrRev(pPath, "/")
    If (pLoc = 0) Then pLoc = InStrRev(pPath, "\")
    If (pLoc <> 0) Then
        GetFileName = Trim(Mid(pPath, pLoc + 1))
    Else
        GetFileName = Trim(pPath)
    End If
    
End Function

'Get Directory From  file Path
Public Function GetFileDir(ByVal pPath As String) As String
    On Error GoTo ErrorTrap
    Dim pName As String
    Dim pLoc As Long
    
    pLoc = InStrRev(pPath, "/")
    If (pLoc = 0) Then pLoc = InStrRev(pPath, "\")
    If (pLoc <> 0) Then
        GetFileDir = Trim(Mid(pPath, 1, pLoc - 1))
        If (Len(GetFileDir) = 2) Then
            GetFileDir = GetFileDir & "\"
        End If
    Else
        GetFileDir = Trim(pPath)
    End If
    Exit Function
ErrorTrap:
    GetFileDir = ""
End Function

Public Function ScrollBarWidth() As Long
    ScrollBarWidth = GetSystemMetrics(SM_CXVSCROLL)
End Function



Public Sub Change_Style(ByVal hWnd As Long)
    
    Call SetWindowLong(hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) Xor _
                              (WS_THICKFRAME) Or (WS_MAXIMIZEBOX))
    ' Tell the nonclient area of the form to redraw itself
    Call SetWindowPos(hWnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
End Sub




Public Sub DeleteFolder(gFile As String, bAllowUndo As Boolean)
    Dim op As SHFILEOPSTRUCT
    With op
        .wFunc = FO_DELETE
        .pFrom = gFile
        If bAllowUndo = True Then
            .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
        End If
    End With
    SHFileOperation op
End Sub


Public Sub DeleteFile(gFile As String, bAllowUndo As Boolean)
    Dim op As SHFILEOPSTRUCT
    With op
        .wFunc = FO_DELETE
        .pFrom = gFile
        If bAllowUndo = True Then
            .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
        End If
    End With
    SHFileOperation op
End Sub
'
'not used in this example
'
Public Sub CopyFile(sFileSource As String, NewLoc As String)
    Dim op As SHFILEOPSTRUCT
    If sFileSource = "" Or NewLoc = "" Then Exit Sub
    With op
        .wFunc = FO_COPY
        .pTo = NewLoc
        .pFrom = sFileSource
        .fFlags = FOF_ALLOWUNDO '+ FOF_SIMPLEPROGRESS
    End With
    SHFileOperation op
End Sub

Public Sub MoveFile(sFileSource As String, NewLoc As String)
    Dim op As SHFILEOPSTRUCT
    If sFileSource = "" Or NewLoc = "" Then Exit Sub
    With op
        .wFunc = FO_MOVE
        .pTo = NewLoc
        .pFrom = sFileSource
        .fFlags = FOF_ALLOWUNDO
    End With
    SHFileOperation op
End Sub






