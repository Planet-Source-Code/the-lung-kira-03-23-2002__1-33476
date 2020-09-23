Attribute VB_Name = "wingdi"
Option Explicit


Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hdc As Long, ByVal nIndex As Long) As Long


Public Const DRIVERVERSION As Long = 0
Public Const TECHNOLOGY As Long = 2
Public Const HORZSIZE As Long = 4
Public Const VERTSIZE As Long = 6
Public Const HORZRES As Long = 8
Public Const VERTRES As Long = 10
Public Const BITSPIXEL As Long = 12
Public Const PLANES As Long = 14
Public Const NUMBRUSHES As Long = 16
Public Const NUMPENS As Long = 18
Public Const NUMMARKERS As Long = 20
Public Const NUMFONTS As Long = 22
Public Const NUMCOLORS As Long = 24
Public Const PDEVICESIZE As Long = 26
Public Const CURVECAPS As Long = 28
Public Const LINECAPS As Long = 30
Public Const POLYGONALCAPS As Long = 32
Public Const TEXTCAPS As Long = 34
Public Const CLIPCAPS As Long = 36
Public Const RASTERCAPS As Long = 38
Public Const ASPECTX As Long = 40
Public Const ASPECTY As Long = 42
Public Const ASPECTXY As Long = 44
Public Const LOGPIXELSX As Long = 88
Public Const LOGPIXELSY As Long = 90
Public Const SIZEPALETTE As Long = 104
Public Const NUMRESERVED As Long = 106
Public Const COLORRES As Long = 108
Public Const PHYSICALWIDTH As Long = 110
Public Const PHYSICALHEIGHT As Long = 111
Public Const PHYSICALOFFSETX As Long = 112
Public Const PHYSICALOFFSETY As Long = 113
Public Const SCALINGFACTORX As Long = 114
Public Const SCALINGFACTORY As Long = 115
Public Const VREFRESH As Long = 116
Public Const DESKTOPVERTRES As Long = 117
Public Const DESKTOPHORZRES As Long = 118
Public Const BLTALIGNMENT As Long = 119
Public Const SHADEBLENDCAPS As Long = 120
Public Const COLORMGMTCAPS As Long = 121

Public Const CCHDEVICENAME As Long = 32
Public Const CCHFORMNAME As Long = 32

Public Const DISPLAY_DEVICE_ATTACHED_TO_DESKTOP As Long = &H1
Public Const DISPLAY_DEVICE_MULTI_DRIVER As Long = &H2
Public Const DISPLAY_DEVICE_PRIMARY_DEVICE As Long = &H4
Public Const DISPLAY_DEVICE_MIRRORING_DRIVER As Long = &H8
Public Const DISPLAY_DEVICE_VGA_COMPATIBLE As Long = &H10
Public Const DISPLAY_DEVICE_REMOVABLE As Long = &H20
Public Const DISPLAY_DEVICE_MODESPRUNED As Long = &H8000000
Public Const DISPLAY_DEVICE_REMOTE As Long = &H4000000
Public Const DISPLAY_DEVICE_DISCONNECT As Long = &H2000000

Public Const DM_ORIENTATION As Long = &H1
Public Const DM_PAPERSIZE As Long = &H2
Public Const DM_PAPERLENGTH As Long = &H4
Public Const DM_PAPERWIDTH As Long = &H8
Public Const DM_SCALE As Long = &H10
Public Const DM_POSITION As Long = &H20
Public Const DM_NUP As Long = &H40
Public Const DM_COPIES As Long = &H100
Public Const DM_DEFAULTSOURCE As Long = &H200
Public Const DM_PRINTQUALITY As Long = &H400
Public Const DM_COLOR As Long = &H800
Public Const DM_DUPLEX As Long = &H1000
Public Const DM_YRESOLUTION As Long = &H2000
Public Const DM_TTOPTION As Long = &H4000
Public Const DM_COLLATE As Long = &H8000
Public Const DM_FORMNAME As Long = &H10000
Public Const DM_LOGPIXELS As Long = &H20000
Public Const DM_BITSPERPEL As Long = &H40000
Public Const DM_PELSWIDTH As Long = &H80000
Public Const DM_PELSHEIGHT As Long = &H100000
Public Const DM_DISPLAYFLAGS As Long = &H200000
Public Const DM_DISPLAYFREQUENCY As Long = &H400000
Public Const DM_ICMMETHOD As Long = &H800000
Public Const DM_ICMINTENT As Long = &H1000000
Public Const DM_MEDIATYPE As Long = &H2000000
Public Const DM_DITHERTYPE As Long = &H4000000
Public Const DM_PANNINGWIDTH As Long = &H8000000
Public Const DM_PANNINGHEIGHT As Long = &H10000000

Public Const LF_FACESIZE As Long = 32
Public Const LF_FULLFACESIZE As Long = 64


Public Type DEVMODE
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
    dmBitsPerPel As Long
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long

    dmPanningWidth As Long
    dmPanningHeight As Long
End Type

Public Type DISPLAY_DEVICE
    cb As Long
    DeviceName As String * 32
    DeviceString As String * 128
    StateFlags As Long
    DeviceID As String * 128
    DeviceKey As String * 128
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
    lfFaceName(LF_FACESIZE - 1) As Byte
End Type
