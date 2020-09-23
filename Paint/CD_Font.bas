Attribute VB_Name = "CD_Font"
Option Explicit

' private internal buffer
Dim iAction As Integer
Dim lAPIReturn As Long
Dim bCancelError As Boolean
Dim lColor As Long
Dim sDialogTitle As String
Dim lExtendedError As Long
Dim lFlags As Long
Dim bFontBold As Boolean
Dim bFontItalic As Boolean
Dim sFontName As String
Dim lFontSize As Long
Dim bFontStrikethru As Boolean
Dim bFontUnderline As Boolean
Dim lHelpCommand As Long
Dim sHelpContext As String
Dim sHelpFile As String
Dim sHelpKey As String
Dim objObject As Object

Dim lhWndOwner As Long

Public Enum DlgFontFlags
   CF_SCREENFONTS = &H1
   CF_PRINTERFONTS = &H2
   CF_SHOWHELP = &H4&
   CF_EFFECTS = &H100&
   CF_ANSIONLY = &H400&
   CF_NOVECTORFONTS = &H800&
   CF_FIXEDPITCHONLY = &H4000&
   CF_FORCEFONTEXIST = &H10000
   CF_SCALABLEONLY = &H20000
   CF_TTONLY = &H40000
   CF_NOFACESEL = &H80000
   CF_NOSTYLESEL = &H100000
   CF_NOSIZESEL = &H200000
   CF_SELECTSCRIPT = &H400000
   CF_NOSCRIPTSEL = &H800000
   CF_NOVERTFONTS = &H1000000
   CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
   CF_NOOEMFONTS = CF_NOVECTORFONTS
   CF_SCRIPTSONLY = CF_ANSIONLY
End Enum

'API
Private Const CLSCD_NOACTION = 0
Private Const CLSCD_SHOWFONT = 4

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT)

Private Const CLSCD_USERCANCELED = 0
Private Const CLSCD_USERSELECTED = 1

Private Const CLSCD_ERRNUMUSRCANCEL = 32755
Private Const CLSCD_ERRDESUSRCANCEL = "Cancel was selected."
Private Const CLSCD_ERRNUMUSRBUFFER = 32756
Private Const CLSCD_ERRDESUSRBUFFER = "Buffer to small"
Private Const FW_BOLD = 700

Private Const FNERR_BUFFERTOOSMALL = &H3003
Private Const FNERR_SUBCLASSFAILURE = &H3001

Private Const LF_FACESIZE = 32

Private Type tLOGFONT
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

Private Type tChooseFont
   lStructSize As Long
   hWndOwner As Long          '  caller's window handle
   hdc As Long                '  printer DC/IC or NULL
   lpLogFont As Long 'tLOGFONT      '  ptr. to a LOGFONT struct
   iPointSize As Long         '  10 * size in points of selected font
   Flags As DlgFontFlags         '  enum. type flags
   rgbColors As Long          '  returned text color
   lCustData As Long          '  data passed to hook fn.
   lpfnHook As Long           '  ptr. to hook function
   lpTemplateName As String   '  custom template name
   hInstance As Long          '  instance handle of.EXE that
                              '  contains cust. dlg. template
   lpszStyle As String        '  return the style field here
                              '  must be LF_FACESIZE or bigger
   nFontType As Integer       '  same value reported to the EnumFonts
                              '  call back with the extra FONTTYPE_
                              '  bits added
   MISSING_ALIGNMENT As Integer
   nSizeMin As Long           '  minimum pt size allowed &
   nSizeMax As Long           '  max pt size allowed if
                              '    CF_LIMITSIZE is used
End Type

Private Declare Function ChooseFontA Lib "comdlg32.dll" (pChoosefont As tChooseFont) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
'API memory functions
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CpyMemValAdrFromRefAdr Lib "kernel32" Alias "RtlMoveMemory" (ByVal hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Sub CpyMemRefAdrFromValAdr Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Any, ByVal cbCopy As Long)

' Read Only
Public Property Get Action() As Integer
   Action = iAction
End Property

' Read Only
Public Property Get APIReturn() As Long
   APIReturn = lAPIReturn
End Property

' Read/Write
Public Property Get CancelError() As Boolean
   CancelError = bCancelError
End Property
Public Property Let CancelError(vNewValue As Boolean)
   bCancelError = vNewValue
End Property

' Read/Write
Public Property Get Color() As Long
   Color = lColor
End Property
Public Property Let Color(vNewValue As Long)
   lColor = vNewValue
End Property


' Read/Write
Public Property Get DialogTitle() As String
   DialogTitle = sDialogTitle
End Property
Public Property Let DialogTitle(vNewValue As String)
   sDialogTitle = vNewValue
End Property

' Read Only
Public Property Get ExtendedError() As Long
   ExtendedError = lExtendedError
End Property


' Read/Write
Public Property Get Flags() As Long
   Flags = lFlags
End Property
Public Property Let Flags(vNewValue As Long)
   lFlags = vNewValue
End Property

' Read/Write
Public Property Get FontBold() As Boolean
   FontBold = bFontBold
End Property
Public Property Let FontBold(vNewValue As Boolean)
   bFontBold = vNewValue
End Property

' Read/Write
Public Property Get FontItalic() As Boolean
   FontItalic = bFontItalic
End Property
Public Property Let FontItalic(vNewValue As Boolean)
   bFontItalic = vNewValue
End Property

' Read/Write
Public Property Get FontName() As String
   FontName = sFontName
End Property
Public Property Let FontName(vNewValue As String)
   sFontName = vNewValue
End Property

' Read/Write
Public Property Get FontSize() As Long
   FontSize = lFontSize
End Property
Public Property Let FontSize(vNewValue As Long)
   lFontSize = vNewValue
End Property

' Read/Write
Public Property Get FontStrikethru() As Boolean
   FontStrikethru = bFontStrikethru
End Property
Public Property Let FontStrikethru(vNewValue As Boolean)
   bFontStrikethru = vNewValue
End Property

' Read/Write
Public Property Get FontUnderline() As Boolean
   FontUnderline = bFontUnderline
End Property
Public Property Let FontUnderline(vNewValue As Boolean)
   bFontUnderline = vNewValue
End Property


' Read/Write
Public Property Get hWndOwner() As Long
   hWndOwner = lhWndOwner
End Property
Public Property Let hWndOwner(vNewValue As Long)
   lhWndOwner = vNewValue
End Property

' Read/Write
Public Property Get HelpCommand() As Long
   HelpCommand = lHelpCommand
End Property
Public Property Let HelpCommand(vNewValue As Long)
   lHelpCommand = vNewValue
End Property

' Read/Write
Public Property Get HelpContext() As String
   HelpContext = sHelpContext
End Property
Public Property Let HelpContext(vNewValue As String)
   sHelpContext = vNewValue
End Property

' Read/Write
Public Property Get HelpFile() As String
   HelpFile = sHelpFile
End Property
Public Property Let HelpFile(vNewValue As String)
   sHelpFile = vNewValue
End Property

' Read/Write
Public Property Get HelpKey() As String
   HelpKey = sHelpKey
End Property
Public Property Let HelpKey(vNewValue As String)
   sHelpKey = vNewValue
End Property


'  Read Only
Public Property Get Object() As Object
   Object = objObject
End Property
Private Sub StringToByteArray(ByVal txt As String, b)
   Dim I As Long
   
   For I = 0 To UBound(b) - 1
      Select Case I
      Case Is = Len(txt): b(I) = 0
      Case Is > Len(txt): Exit For ' b(I) = 0
      Case Is < Len(txt): b(I) = Asc(Mid(txt, I + 1, 1))
      End Select
'      Debug.Print b(I);
   Next I
'   Debug.Print

End Sub

'Provide the ShowFont method and interface with the Win32 ChooseFont function.
Public Sub ShowFont()
   Dim vLogFont As tLOGFONT
   Dim vChooseFont As tChooseFont
   Dim lLogFontSize As Long
   Dim lLogFontAddress As Long
   Dim lMemHandle As Long
   Dim lReturn As Long
   Dim sFont As String
   Dim lBytePoint As Long
   Dim I As Long
   
   'On Error GoTo ShowFontError
   
   iAction = CLSCD_SHOWFONT    'Action property
   lAPIReturn = 0  'APIReturn property
   lExtendedError = 0  'ExtendedError property

   StringToByteArray sFontName, vLogFont.lfFaceName
   If bFontBold = True Then vLogFont.lfWeight = FW_BOLD
   If bFontItalic = True Then vLogFont.lfItalic = 1
   If bFontUnderline = True Then vLogFont.lfUnderline = 1
   If bFontStrikethru = True Then vLogFont.lfStrikeOut = 1
   vLogFont.lfHeight = lFontSize * 1.333
   
   lLogFontSize = Len(vLogFont)
   lMemHandle = GlobalAlloc(GHND, lLogFontSize)
   If lMemHandle Then
      lLogFontAddress = GlobalLock(lMemHandle)
      If lLogFontAddress Then
         CpyMemValAdrFromRefAdr lLogFontAddress, vLogFont, lLogFontSize
         vChooseFont.lpLogFont = lLogFontAddress
         vChooseFont.iPointSize = lFontSize * 10
         vChooseFont.Flags = lFlags Or &H40& 'CF_INITTOLOGFONTSTRUCT
         vChooseFont.rgbColors = lColor
         vChooseFont.hWndOwner = lhWndOwner
         vChooseFont.lStructSize = Len(vChooseFont)
         
         lAPIReturn = ChooseFontA(vChooseFont)    'store to APIReturn property
         Select Case lAPIReturn
            Case CLSCD_USERCANCELED
               If bCancelError = True Then
                  lReturn = GlobalUnlock(lMemHandle)
                  lReturn = GlobalFree(lMemHandle)
                  On Error GoTo 0
                  Err.Raise Number:=CLSCD_ERRNUMUSRCANCEL, _
                      Description:=CLSCD_ERRDESUSRCANCEL
                  Exit Sub
                  End If
            Case CLSCD_USERSELECTED
               lReturn = GlobalUnlock(lMemHandle)
               lReturn = GlobalFree(lMemHandle)
               CpyMemRefAdrFromValAdr vLogFont, lLogFontAddress, lLogFontSize
               bFontBold = IIf(vLogFont.lfWeight >= FW_BOLD, True, False)
               bFontItalic = IIf(vLogFont.lfItalic <> 0, True, False)
               bFontUnderline = IIf(vLogFont.lfUnderline = 1, True, False)
               bFontStrikethru = IIf(vLogFont.lfStrikeOut = 1, True, False)
               sFontName = StrConv(vLogFont.lfFaceName, vbUnicode)
               'sFontName = sByteArrayToString(vLogFont.lfFaceName())
               lFontSize = CLng(vChooseFont.iPointSize / 10)
               lColor = vChooseFont.rgbColors
            Case Else   'An error occurred.
               lReturn = GlobalUnlock(lMemHandle)
               lReturn = GlobalFree(lMemHandle)
               lExtendedError = CommDlgExtendedError
            End Select
         Else
         lReturn = GlobalFree(lMemHandle)
         Exit Sub
         End If
      End If
   Exit Sub
   
ShowFontError:
   lReturn = GlobalUnlock(lMemHandle)
   lReturn = GlobalFree(lMemHandle)
   Exit Sub

End Sub

