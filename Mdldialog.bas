Attribute VB_Name = "Mdldialog"
Option Explicit
Public PathOfDialog

Public Const FW_NORMAL = 400
Public Const DEFAULT_CHARSET = 1
Public Const OUT_DEFAULT_PRECIS = 0
Public Const CLIP_DEFAULT_PRECIS = 0
Public Const DEFAULT_QUALITY = 0
Public Const DEFAULT_PITCH = 0
Public Const FF_ROMAN = 16
Public Const CF_PRINTERFONTS = &H2
Public Const CF_SCREENFONTS = &H1
Public Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Public Const CF_EFFECTS = &H100&
Public Const CF_FORCEFONTEXIST = &H10000
Public Const CF_INITTOLOGFONTSTRUCT = &H40&
Public Const CF_LIMITSIZE = &H2000&
Public Const REGULAR_FONTTYPE = &H400
Public Const LF_FACESIZE = 32
Public Const CCHDEVICENAME = 32
Public Const CCHFORMNAME = 32
Public Const GMEM_MOVEABLE = &H2
Public Const GMEM_ZEROINIT = &H40
Public Const DM_DUPLEX = &H1000&
Public Const DM_ORIENTATION = &H1&
Public Const PD_PRINTSETUP = &H40
Public Const PD_DISABLEPRINTTOFILE = &H80000

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
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

Public Type CHOOSECOLOR
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Type LOGFONT
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
        lfFaceName As String * 31
End Type
Public Type CHOOSEFONT
        lStructSize As Long
        hWndOwner As Long          '  caller's window handle
        hDC As Long                '  printer DC/IC or NULL
        lpLogFont As Long          '  ptr. to a LOGFONT struct
        iPointSize As Long         '  10 * size in points of selected font
        Flags As Long              '  enum. type flags
        rgbColors As Long          '  returned text color
        lCustData As Long          '  data passed to hook fn.
        lpfnHook As Long           '  ptr. to hook function
        lpTemplateName As String     '  custom template name
        hInstance As Long          '  instance handle of.EXE that
                                       '    contains cust. dlg. template
        lpszStyle As String          '  return the style field here
                                       '  must be LF_FACESIZE or bigger
        nFontType As Integer          '  same value reported to the EnumFonts
                                       '    call back with the extra FONTTYPE_
                                       '    bits added
        MISSING_ALIGNMENT As Integer
        nSizeMin As Long           '  minimum pt size allowed &
        nSizeMax As Long           '  max pt size allowed if
                                       '    CF_LIMITSIZE is used
End Type
Private Type PRINTDLG_TYPE
    lStructSize As Long
    hWndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    Flags As Long
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
Private Type DEVNAMES_TYPE
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
    extra As String * 100
End Type
Private Type DEVMODE_TYPE
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
Public Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function PrintDialog Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLG_TYPE) As Long

Public Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONT) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Dim OFName As OPENFILENAME
Dim CustomColors() As Byte
Public lfont As LOGFONT

Function ShowColor(FormaHandle As Form) As Long
    Dim cc As CHOOSECOLOR
    Dim Custcolor(16) As Long

    Dim lReturn As Long

    'set the structure size
    cc.lStructSize = Len(cc)
    'Set the owner
    cc.hWndOwner = FormaHandle.hwnd
    'set the application's instance
    cc.hInstance = App.hInstance
    'set the custom colors (converted to Unicode)
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    'no extra flags
    cc.Flags = 0

    'Show the 'Select Color'-dialog
    If CHOOSECOLOR(cc) <> 0 Then
        ShowColor = cc.rgbResult
        CustomColors = StrConv(cc.lpCustColors, vbFromUnicode)
    Else
        ShowColor = -1
    End If
End Function
Function ShowOpen(FormaHandle As Form) As String
    'Set the structure size
    OFName.lStructSize = Len(OFName)
    'Set the owner window
    OFName.hWndOwner = FormaHandle.hwnd
    'Set the application's instance
    OFName.hInstance = App.hInstance
    'Set the filet  FST
    OFName.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Text Files (*.INI)" + Chr$(0) + "*.INI" + Chr$(0) + "Text Files (*.log)" + Chr$(0) + "*.log" + Chr$(0) + "Text Files (*.X)" + Chr$(0) + "*.X" + Chr$(0) + "Text Files (*.Inf)" + Chr$(0) + "*.Inf" + Chr$(0) + "Functions Of Script Text Files (*.FST)" + Chr$(0) + "*.FST" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'Create a buffer
    OFName.lpstrFile = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFile = 255
    'Create a buffer
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum number of chars
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    OFName.lpstrInitialDir = PathOfDialog
    'Set the dialog title
    OFName.lpstrTitle = "Open File"
    'no extra flags
    OFName.Flags = 0

    'Show the 'Open File'-dialog
    If GetOpenFileName(OFName) Then
        ShowOpen = Trim$(OFName.lpstrFile)
        PathOfDialog = Trim$(OFName.lpstrFile)
    Else
        ShowOpen = ""
    End If
End Function
 Function ShowFont(FormaHandle As Form) As String
    Dim cf As CHOOSEFONT, hMem As Long, pMem As Long
    Dim fontname As String, retval As Long
    lfont.lfHeight = 0  ' determine default height
    lfont.lfWidth = 0  ' determine default width
    lfont.lfEscapement = 0  ' angle between baseline and escapement vector
    lfont.lfOrientation = 0  ' angle between baseline and orientation vector
    lfont.lfWeight = FW_NORMAL  ' normal weight i.e. not bold
    lfont.lfCharSet = DEFAULT_CHARSET  ' use default character set
    lfont.lfOutPrecision = OUT_DEFAULT_PRECIS  ' default precision mapping
    lfont.lfClipPrecision = CLIP_DEFAULT_PRECIS  ' default clipping precision
    lfont.lfQuality = DEFAULT_QUALITY  ' default quality setting
    lfont.lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN  ' default pitch, proportional with serifs
    lfont.lfFaceName = "Times New Roman" & vbNullChar  ' string must be null-terminated
    ' Create the memory block which will act as the LOGFONT structure buffer.
    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
    pMem = GlobalLock(hMem)  ' lock and get pointer
    CopyMemory ByVal pMem, lfont, Len(lfont)  ' copy structure's contents into block
    ' Initialize dialog box: Screen and printer fonts, point size between 10 and 72.
    cf.lStructSize = Len(cf)  ' size of structure
    cf.hWndOwner = FormaHandle.hwnd  ' window Form1 is opening this dialog box
    cf.hDC = Printer.hDC  ' device context of default printer (using VB's mechanism)
    cf.lpLogFont = pMem   ' pointer to LOGFONT memory block buffer
    cf.iPointSize = 120  ' 12 point font (in units of 1/10 point)
    cf.Flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
    cf.rgbColors = RGB(0, 0, 0)  ' black
    cf.nFontType = REGULAR_FONTTYPE  ' regular font type i.e. not bold or anything
    cf.nSizeMin = 10  ' minimum point size
    cf.nSizeMax = 72  ' maximum point size
    ' Now, call the function.  If successful, copy the LOGFONT structure back into the structure
    ' and then print out the attributes we mentioned earlier that the user selected.
    retval = CHOOSEFONT(cf)  ' open the dialog box
    If retval <> 0 Then  ' success
        CopyMemory lfont, ByVal pMem, Len(lfont)  ' copy memory back
        ' Now make the fixed-length string holding the font name into a "normal" string.
        ShowFont = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
        Debug.Print  ' end the line
    End If
    ' Deallocate the memory block we created earlier.  Note that this must
    ' be done whether the function succeeded or not.
    retval = GlobalUnlock(hMem)  ' destroy pointer, unlock block
    retval = GlobalFree(hMem)  ' free the allocated memory
End Function
 Function ShowSave(FormaHandle As Form) As String
    OFName.lStructSize = Len(OFName)
    OFName.hWndOwner = FormaHandle.hwnd
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Text Files (*.INI)" + Chr$(0) + "*.INI" + Chr$(0) + "Text Files (*.log)" + Chr$(0) + "*.log" + Chr$(0) + "Text Files (*.X)" + Chr$(0) + "*.X" + Chr$(0) + "Text Files (*.Inf)" + Chr$(0) + "*.Inf" + Chr$(0) + "Functions Of Script Text Files (*.FST)" + Chr$(0) + "*.FST" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    OFName.lpstrFile = Space$(254)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = Space$(254)
    OFName.nMaxFileTitle = 255
    OFName.lpstrInitialDir = PathOfDialog
    OFName.lpstrTitle = "Save File"
    OFName.Flags = 0
    
    Dim Thedot  As String
If GetSaveFileName(OFName) Then

ShowSave = OFName.lpstrFile

Else
ShowSave = ""
End If


End Function
Public Sub ShowPrinter(frmOwner As Form, Optional PrintFlags As Long)
    '-> Code by Donald Grover
    Dim PrintDlg As PRINTDLG_TYPE
    Dim DEVMODE As DEVMODE_TYPE
    Dim DevName As DEVNAMES_TYPE

    Dim lpDevMode As Long, lpDevName As Long
    Dim bReturn As Integer
    Dim objPrinter As Printer, NewPrinterName As String

    ' Use PrintDialog to get the handle to a memory
    ' block with a DevMode and DevName structures

    PrintDlg.lStructSize = Len(PrintDlg)
    PrintDlg.hWndOwner = frmOwner.hwnd

    PrintDlg.Flags = PrintFlags
    On Error Resume Next
    'Set the current orientation and duplex setting
    DEVMODE.dmDeviceName = Printer.DeviceName
    DEVMODE.dmSize = Len(DEVMODE)
    DEVMODE.dmFields = DM_ORIENTATION Or DM_DUPLEX
    DEVMODE.dmPaperWidth = Printer.Width
    DEVMODE.dmOrientation = Printer.Orientation
    DEVMODE.dmPaperSize = Printer.PaperSize
    DEVMODE.dmDuplex = Printer.Duplex
    On Error GoTo 0

    'Allocate memory for the initialization hDevMode structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DEVMODE))
    lpDevMode = GlobalLock(PrintDlg.hDevMode)
    If lpDevMode > 0 Then
        CopyMemory ByVal lpDevMode, DEVMODE, Len(DEVMODE)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
    End If

    'Set the current driver, device, and port name strings
    With DevName
        .wDriverOffset = 8
        .wDeviceOffset = .wDriverOffset + 1 + Len(Printer.DriverName)
        .wOutputOffset = .wDeviceOffset + 1 + Len(Printer.Port)
        .wDefault = 0
    End With

    With Printer
        DevName.extra = .DriverName & Chr(0) & .DeviceName & Chr(0) & .Port & Chr(0)
    End With

    'Allocate memory for the initial hDevName structure
    'and copy the settings gathered above into this memory
    PrintDlg.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(DevName))
    lpDevName = GlobalLock(PrintDlg.hDevNames)
    If lpDevName > 0 Then
        CopyMemory ByVal lpDevName, DevName, Len(DevName)
        bReturn = GlobalUnlock(lpDevName)
    End If

    'Call the print dialog up and let the user make changes
    If PrintDialog(PrintDlg) <> 0 Then

        'First get the DevName structure.
        lpDevName = GlobalLock(PrintDlg.hDevNames)
        CopyMemory DevName, ByVal lpDevName, 45
        bReturn = GlobalUnlock(lpDevName)
        GlobalFree PrintDlg.hDevNames

        'Next get the DevMode structure and set the printer
        'properties appropriately
        lpDevMode = GlobalLock(PrintDlg.hDevMode)
        CopyMemory DEVMODE, ByVal lpDevMode, Len(DEVMODE)
        bReturn = GlobalUnlock(PrintDlg.hDevMode)
        GlobalFree PrintDlg.hDevMode
        NewPrinterName = UCase$(Left(DEVMODE.dmDeviceName, InStr(DEVMODE.dmDeviceName, Chr$(0)) - 1))
        If Printer.DeviceName <> NewPrinterName Then
            For Each objPrinter In Printers
                If UCase$(objPrinter.DeviceName) = NewPrinterName Then
                    Set Printer = objPrinter
                    'set printer toolbar name at this point
                End If
            Next
        End If

        On Error Resume Next
        'Set printer object properties according to selections made
        'by user
        Printer.Copies = DEVMODE.dmCopies
        Printer.Duplex = DEVMODE.dmDuplex
        Printer.Orientation = DEVMODE.dmOrientation
        Printer.PaperSize = DEVMODE.dmPaperSize
        Printer.PrintQuality = DEVMODE.dmPrintQuality
        Printer.ColorMode = DEVMODE.dmColor
        Printer.PaperBin = DEVMODE.dmDefaultSource
        On Error GoTo 0
    End If
End Sub



