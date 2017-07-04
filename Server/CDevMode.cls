VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDevMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *************************************************************************
'  Copyright ©2001 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' Win32 API declares
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' Values used to define DEVMODE structure
Private Const CCHDEVICENAME As Long = 32
Private Const CCHFORMNAME As Long = 32

' Structure used to cache values
Private Type DevMode
   dmDeviceName As String * CCHDEVICENAME
   dmSpecVersion As Integer
   dmDriverVersion As Integer
   dmSize As Integer               ' not exposed
   dmDriverExtra As Integer        ' not exposed
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
   dmLogPixels As Integer
   dmBitsPerPel As Long
   dmPelsWidth As Long
   dmPelsHeight As Long
   dmNup As Long            ' union with dmDisplayFlags As Long
   dmDisplayFrequency As Long
End Type

' Bin selections
Private Const DMBIN_UPPER = 1
Private Const DMBIN_ONLYONE = 1
Private Const DMBIN_LOWER = 2
Private Const DMBIN_MIDDLE = 3
Private Const DMBIN_MANUAL = 4
Private Const DMBIN_ENVELOPE = 5
Private Const DMBIN_ENVMANUAL = 6
Private Const DMBIN_AUTO = 7
Private Const DMBIN_TRACTOR = 8
Private Const DMBIN_SMALLFMT = 9
Private Const DMBIN_LARGEFMT = 10
Private Const DMBIN_LARGECAPACITY = 11
Private Const DMBIN_CASSETTE = 14
Private Const DMBIN_FIRST = DMBIN_UPPER
Private Const DMBIN_LAST = DMBIN_CASSETTE
Private Const DMBIN_USER = 256               '  device specific bins start here

' Orientation selections
Private Const DMORIENT_PORTRAIT = 1
Private Const DMORIENT_LANDSCAPE = 2

' Paper selections
Private Const DMPAPER_LETTER = 1                 '  Letter 8 1/2 x 11 in
Private Const DMPAPER_LETTERSMALL = 2            '  Letter Small 8 1/2 x 11 in
Private Const DMPAPER_TABLOID = 3                '  Tabloid 11 x 17 in
Private Const DMPAPER_LEDGER = 4                 '  Ledger 17 x 11 in
Private Const DMPAPER_LEGAL = 5                  '  Legal 8 1/2 x 14 in
Private Const DMPAPER_STATEMENT = 6              '  Statement 5 1/2 x 8 1/2 in
Private Const DMPAPER_EXECUTIVE = 7              '  Executive 7 1/4 x 10 1/2 in
Private Const DMPAPER_A3 = 8                     '  A3 297 x 420 mm
Private Const DMPAPER_A4 = 9                     '  A4 210 x 297 mm
Private Const DMPAPER_A4SMALL = 10               '  A4 Small 210 x 297 mm
Private Const DMPAPER_A5 = 11                    '  A5 148 x 210 mm
Private Const DMPAPER_B4 = 12                    '  B4 250 x 354
Private Const DMPAPER_B5 = 13                    '  B5 182 x 257 mm
Private Const DMPAPER_FOLIO = 14                 '  Folio 8 1/2 x 13 in
Private Const DMPAPER_QUARTO = 15                '  Quarto 215 x 275 mm
Private Const DMPAPER_10X14 = 16                 '  10x14 in
Private Const DMPAPER_11X17 = 17                 '  11x17 in
Private Const DMPAPER_NOTE = 18                  '  Note 8 1/2 x 11 in
Private Const DMPAPER_ENV_9 = 19                 '  Envelope #9 3 7/8 x 8 7/8
Private Const DMPAPER_ENV_10 = 20                '  Envelope #10 4 1/8 x 9 1/2
Private Const DMPAPER_ENV_11 = 21                '  Envelope #11 4 1/2 x 10 3/8
Private Const DMPAPER_ENV_12 = 22                '  Envelope #12 4 \276 x 11
Private Const DMPAPER_ENV_14 = 23                '  Envelope #14 5 x 11 1/2
Private Const DMPAPER_CSHEET = 24                '  C size sheet
Private Const DMPAPER_DSHEET = 25                '  D size sheet
Private Const DMPAPER_ESHEET = 26                '  E size sheet
Private Const DMPAPER_ENV_DL = 27                '  Envelope DL 110 x 220mm
Private Const DMPAPER_ENV_C5 = 28                '  Envelope C5 162 x 229 mm
Private Const DMPAPER_ENV_C3 = 29                '  Envelope C3  324 x 458 mm
Private Const DMPAPER_ENV_C4 = 30                '  Envelope C4  229 x 324 mm
Private Const DMPAPER_ENV_C6 = 31                '  Envelope C6  114 x 162 mm
Private Const DMPAPER_ENV_C65 = 32               '  Envelope C65 114 x 229 mm
Private Const DMPAPER_ENV_B4 = 33                '  Envelope B4  250 x 353 mm
Private Const DMPAPER_ENV_B5 = 34                '  Envelope B5  176 x 250 mm
Private Const DMPAPER_ENV_B6 = 35                '  Envelope B6  176 x 125 mm
Private Const DMPAPER_ENV_ITALY = 36             '  Envelope 110 x 230 mm
Private Const DMPAPER_ENV_MONARCH = 37           '  Envelope Monarch 3.875 x 7.5 in
Private Const DMPAPER_ENV_PERSONAL = 38          '  6 3/4 Envelope 3 5/8 x 6 1/2 in
Private Const DMPAPER_FANFOLD_US = 39            '  US Std Fanfold 14 7/8 x 11 in
Private Const DMPAPER_FANFOLD_STD_GERMAN = 40    '  German Std Fanfold 8 1/2 x 12 in
Private Const DMPAPER_FANFOLD_LGL_GERMAN = 41    '  German Legal Fanfold 8 1/2 x 13 in
Private Const DMPAPER_FIRST = DMPAPER_LETTER
Private Const DMPAPER_LAST = DMPAPER_FANFOLD_LGL_GERMAN

' Print qualities
Private Const DMRES_DRAFT = (-1)
Private Const DMRES_LOW = (-2)
Private Const DMRES_MEDIUM = (-3)
Private Const DMRES_HIGH = (-4)

' Color enable/disable for color printers
Private Const DMCOLOR_MONOCHROME = 1
Private Const DMCOLOR_COLOR = 2

' Duplex enable
Private Const DMDUP_SIMPLEX = 1
Private Const DMDUP_VERTICAL = 2
Private Const DMDUP_HORIZONTAL = 3

' TrueType options
Private Const DMTT_BITMAP = 1            '  print TT fonts as graphics
Private Const DMTT_DOWNLOAD = 2          '  download TT fonts as soft fonts
Private Const DMTT_SUBDEV = 3            '  substitute device fonts for TT fonts
Private Const DMTT_DOWNLOAD_OUTLINE = 4  '  download TT fonts as outline soft fonts

' Collation selections
Private Const DMCOLLATE_FALSE = 0
Private Const DMCOLLATE_TRUE = 1

' dmNup, multiple logical page per physical page options
Private Const DMNUP_SYSTEM = 1
Private Const DMNUP_ONEUP = 2

' Member variables
Private m_dm As DevMode

' *********************************************
'  Initialize/Terminate
' *********************************************
Private Sub Class_Initialize()
   '
End Sub

Private Sub Class_Terminate()
   '
End Sub

' *********************************************
'  Public Properties (Read-Only)
' *********************************************
Public Property Get BitsPerPel() As Long
   ' Specifies the color resolution, in bits per pixel, of
   ' the display device (for example: 4 bits for 16 colors,
   ' 8 bits for 256 colors, or 16 bits for 65,536 colors).
   ' Display drivers use this member, for example, in the
   ' ChangeDisplaySettings function. Printer drivers do not
   ' use this member
   BitsPerPel = m_dm.dmBitsPerPel
End Property

Public Property Get Collate() As Integer
   ' Specifies whether collation should be used when printing
   ' multiple copies. (This member is ignored unless the printer
   ' driver indicates support for collation by setting the
   ' dmFields member to DM_COLLATE.) This member can be be one
   ' of the DMCOLLATE_* values.
   Collate = m_dm.dmCollate
End Property

Public Property Get Color() As Integer
   ' Switches between color and monochrome on color
   ' printers, using DMCOLOR_* values.
   Color = m_dm.dmColor
End Property

Public Property Get Copies() As Integer
   ' Selects the number of copies printed if the device
   ' supports multiple-page copies.
   Copies = m_dm.dmCopies
End Property
   
Public Property Get DefaultSource() As Integer
   ' Specifies the paper source. To retrieve a list of the
   ' available paper sources for a printer, use the
   ' DeviceCapabilities function with the DC_BINS flag.
   ' This member can be one of the DMBIN_* values, or it
   ' can be a device-specific value greater than or equal
   ' to DMBIN_USER.
   DefaultSource = m_dm.dmDefaultSource
End Property

Public Property Get DeviceName() As String
   ' Specifies the the "friendly" name of the printer;
   ' for example, "PCL/HP LaserJet" in the case of
   ' PCL/HP LaserJet®. This string is unique among device
   ' drivers. Note that this name may be truncated to fit
   ' in the dmDeviceName array.
   DeviceName = TrimNull(m_dm.dmDeviceName)
End Property

Public Property Get DisplayFrequency() As Long
   ' Specifies the frequency, in hertz (cycles per second),
   ' of the display device in a particular mode. This value
   ' is also known as the display device's vertical refresh
   ' rate. Display drivers use this member. It is used, for
   ' example, in the ChangeDisplaySettings function. Printer
   ' drivers do not use this member.
   DisplayFrequency = m_dm.dmDisplayFrequency
End Property

Public Property Get DriverVersion() As Integer
   ' Specifies the printer driver version number assigned
   ' by the printer driver developer.
   DriverVersion = m_dm.dmDriverVersion
End Property

Public Property Get Duplex() As Integer
   ' Selects duplex or double-sided printing for printers
   ' capable of duplex printing.
   Duplex = m_dm.dmDuplex
End Property

Public Property Get Fields() As Long
   ' Specifies whether certain members of the DEVMODE
   ' structure have been initialized. If a member is
   ' initialized, its corresponding bit is set, otherwise
   ' the bit is clear. A printer driver supports only
   ' those DEVMODE members that are appropriate for the
   ' printer technology.
   Fields = m_dm.dmFields
End Property

Public Property Get FormName() As String
   ' Windows NT/Windows 2000: Specifies the name of the form
   ' to use; for example, "Letter" or "Legal". A complete set
   ' of names can be retrieved by using the EnumForms function.
   ' Windows 95: Printer drivers do not use this member.
   FormName = TrimNull(m_dm.dmFormName)
End Property

Public Property Get LogPixels() As Integer
   ' Specifies the number of pixels per logical inch. Printer
   ' drivers do not use this member.
   LogPixels = m_dm.dmLogPixels
End Property

Public Property Get NUP() As Long
   ' Specifies where the NUP is done. It can be one of the
   ' DMNUP_* values.
   NUP = m_dm.dmNup
End Property

Public Property Get Orientation() As Integer
   ' For printer devices only, selects the orientation of
   ' the paper. This member can be either:
   '   DMORIENT_PORTRAIT (1) or
   '   DMORIENT_LANDSCAPE (2).
   Orientation = m_dm.dmOrientation
End Property

Public Property Get PaperLength() As Integer
   ' For printer devices only, overrides the length of the paper
   ' specified by the dmPaperSize member, either for custom paper
   ' sizes or for devices such as dot-matrix printers that can
   ' print on a page of arbitrary length. These values, along with
   ' all other values in this structure that specify a physical
   ' length, are in tenths of a millimeter.
   PaperLength = m_dm.dmPaperLength
End Property

Public Property Get PaperSize() As Integer
   ' For printer devices only, selects the size of the paper
   ' to print on. This member can be set to zero if the length
   ' and width of the paper are both set by the dmPaperLength
   ' and dmPaperWidth members. Otherwise, the dmPaperSize member
   ' can be set to one of the DMPAPER_* predefined values.
   PaperSize = m_dm.dmPaperSize
End Property

Public Property Get PaperWidth() As Integer
   ' For printer devices only, overrides the width of the paper
   ' specified by the dmPaperSize member.
   PaperWidth = m_dm.dmPaperWidth
End Property

Public Property Get PelsHeight() As Long
   ' Specifies the height, in pixels, of the visible device
   ' surface. Display drivers use this member, for example,
   ' in the ChangeDisplaySettings function. Printer drivers
   ' do not use this member.
   PelsHeight = m_dm.dmPelsHeight
End Property

Public Property Get PelsWidth() As Long
   ' Specifies the width, in pixels, of the visible device
   ' surface. Display drivers use this member, for example,
   ' in the ChangeDisplaySettings function. Printer drivers
   ' do not use this member.
   PelsWidth = m_dm.dmPelsWidth
End Property

Public Property Get PrintQuality() As Integer
   ' Specifies the printer resolution. There are four predefined
   ' device-independent (DMRES_*) values. If a positive value is
   ' specified, it specifies the number of dots per inch (DPI) and
   ' is therefore device dependent.
   PrintQuality = m_dm.dmPrintQuality
End Property

Public Property Get ScaleFactor() As Integer
   ' Specifies the factor by which the printed output is to be
   ' scaled. The apparent page size is scaled from the physical
   ' page size by a factor of dmScale/100. For example, a letter-
   ' sized page with a dmScale value of 50 would contain as much
   ' data as a page of 17- by 22-inches because the output text
   ' and graphics would be half their original height and width.
   ScaleFactor = m_dm.dmScale
End Property

Public Property Get SpecVersion() As Integer
   ' Specifies the version number of the initialization data
   ' specification on which the structure is based. To ensure
   ' the correct version is used for any operating system,
   ' use DM_SPECVERSION.
   SpecVersion = m_dm.dmSpecVersion
End Property

Public Property Get TTOption() As Integer
   ' Specifies how TrueType® fonts should be printed. This
   ' member can be one of the DMTT_* values.
   TTOption = m_dm.dmTTOption
End Property

Public Property Get YResolution() As Integer
   ' Specifies the y-resolution, in dots per inch, of the printer.
   ' If the printer initializes this member, the dmPrintQuality
   ' member specifies the x-resolution, in dots per inch, of the
   ' printer.
   YResolution = m_dm.dmYResolution
End Property

' *********************************************
'  Friend Methods
' *********************************************
Friend Sub Initialize(ByVal lpDevMode As Long)
   ' Just a simple copy to fill cache.
   If lpDevMode Then
      Call CopyMemory(m_dm, ByVal lpDevMode, Len(m_dm))
   End If
End Sub

' *********************************************
'  Private Methods
' *********************************************
Private Function TrimNull(ByVal StrIn As String) As String
   Dim nul As Long
   
   ' Truncate input string at first null.
   ' If no nulls, perform ordinary Trim.
   nul = InStr(StrIn, vbNullChar)
   Select Case nul
      Case Is > 1
         TrimNull = Left$(StrIn, nul - 1)
      Case 1
         TrimNull = ""
      Case 0
         TrimNull = Trim$(StrIn)
   End Select
End Function

