Attribute VB_Name = "GflLib"
' Graphics File Library Extended
'
' GFL library Copyright (c) 1991-2007 Pierre-e Gougelet
' All rights reserved
' Commercial use is not authorized without agreement
'
' Interface for Visual Basic : Jérôme Quintard (contact@jeromequintard.com)

Option Explicit

Public Const GFL_VERSION = "3.05"
Public Const GFL_FALSE = 0
Public Const GFL_TRUE = 1

'Erreurs
Public Enum GFL_ERROR
    GFL_NO_ERROR = 0
    GFL_ERROR_FILE_OPEN = 1
    GFL_ERROR_FILE_READ = 2
    GFL_ERROR_FILE_CREATE = 3
    GFL_ERROR_FILE_WRITE = 4
    GFL_ERROR_NO_MEMORY = 5
    GFL_ERROR_UNKNOWN_FORMAT = 6
    GFL_ERROR_BAD_BITMAP = 7
    GFL_ERROR_BAD_FORMAT_INDEX = 10
    GFL_ERROR_BAD_PARAMETERS = 50
    GFL_UNKNOWN_ERROR = 255
End Enum

'Origines
Public Enum GFL_ORIGIN
    GFL_LEFT = &H0
    GFL_RIGHT = &H1
    GFL_TOP = &H0
    GFL_BOTTOM = &H10
    GFL_TOP_LEFT = (GFL_TOP Or GFL_LEFT)
    GFL_BOTTOM_LEFT = (GFL_BOTTOM Or GFL_LEFT)
    GFL_TOP_RIGHT = (GFL_TOP Or GFL_RIGHT)
    GFL_BOTTOM_RIGHT = (GFL_BOTTOM Or GFL_RIGHT)
End Enum

'Type de Compressions
Public Enum GFL_COMPRESSION
    GFL_NO_COMPRESSION = 0
    GFL_RLE = 1
    GFL_LZW = 2
    GFL_JPEG = 3
    GFL_ZIP = 4
    GFL_SGI_RLE = 5
    GFL_CCITT_RLE = 6
    GFL_CCITT_FAX3 = 7
    GFL_CCITT_FAX3_2D = 8
    GFL_CCITT_FAX4 = 9
    GFL_WAVELET = 10
    GFL_LZW_PREDICTOR = 11
    GFL_UNKNOWN_COMPRESSION = 255
End Enum

'Type de Bitmap
Public Enum GFL_BITMAP_TYPE
    GFL_BINARY = &H1
    GFL_GREY = &H2
    GFL_COLORS = &H4
    GFL_RGB = &H10
    GFL_RGBA = &H20
    GFL_BGR = &H40
    GFL_ABGR = &H80
    GFL_BGRA = &H100
    GFL_ARGB = &H200
    GFL_CMYK = &H400
End Enum

'Structure BITMAP
Public Type GFL_COLORMAP
    Red(255) As Byte
    Green(255) As Byte
    Blue(255) As Byte
End Type

Public Type GFL_COLOR
    Red As Integer
    Green As Integer
    Blue As Integer
    Alpha As Integer
End Type

'Structure BITMAP
Public Type GFL_BITMAP
    Type As Integer
    Origin As Integer
    Width As Long
    Height As Long
    BytesPerLine As Long
    LinePadding As Integer
    BitsPerComponent As Integer
    ComponentsPerPixel As Integer
    BytesPerPixel As Integer
    Xdpi As Integer
    Ydpi As Integer
    TransparentIndex As Integer
    Reserved As Integer
    ColorUsed As Long
    ColorMap As Long
    data As Long
    Comment As Long
    MetaData As Long
    XOffset As Long
    YOffset As Long
    Name As Long
End Type

'Ordre des cannaux
Public Enum GFL_CORDER
    GFL_CORDER_INTERLEAVED = 0
    GFL_CORDER_SEQUENTIAL = 1
    GFL_CORDER_SEPARATE = 2
End Enum

'Type de cannaux
Public Enum GFL_CTYPE
    GFL_CTYPE_GREYSCALE = 0
    GFL_CTYPE_RGB = 1
    GFL_CTYPE_BGR = 2
    GFL_CTYPE_RGBA = 3
    GFL_CTYPE_ABGR = 4
    GFL_CTYPE_CMY = 5
    GFL_CTYPE_CMYK = 6
End Enum

'Load_Params Flags
Public Const GFL_LOAD_SKIP_ALPHA = &H1
Public Const GFL_LOAD_IGNORE_READ_ERROR = &H2
Public Const GFL_LOAD_BY_EXTENSION_ONLY = &H4
Public Const GFL_LOAD_READ_ALL_COMMENT = &H8
Public Const GFL_LOAD_FORCE_COLOR_MODEL = &H10
Public Const GFL_LOAD_PREVIEW_NO_CANVAS_RESIZE = &H20
Public Const GFL_LOAD_BINARY_AS_GREY = &H40
Public Const GFL_LOAD_ORIGINAL_COLORMODEL = &H40
Public Const GFL_LOAD_ONLY_FIRST_FRAME = &H100
Public Const GFL_LOAD_ORIGINAL_DEPTH = &H200
Public Const GFL_LOAD_METADATA = &H400
Public Const GFL_LOAD_COMMENT = &H800
Public Const GFL_LOAD_HIGH_QUALITY_THUMBNAIL = &H1000

'Inutilisé dans cette version donc seulement un type Long pour respecter
'la taille de la structure
Public Type LOAD_CALLBACK_STRUCT
    Read As Long
    Tell As Long
    Seek As Long

        AllocateBitmap As Long
        AllocateBitmapParams As Long
        Progress As Long
        ProgressParams As Long
        WantCancel As Long
        WantCancelParams As Long
End Type

Public Type SAVE_CALLBACK_STRUCT
    Write As Long
    Tell As Long
    Seek As Long
    GetLine As Long
    GetLineParams As Long
End Type

'Load Params Struct
Public Type GFL_LOAD_PARAMS
        flags As Long
        FormatIndex As Long ' -1 pour détection automatique
        ImageWanted As Long
        Origin As Integer
        ColorModel As Integer
        LinePadding As Long
        DefaultAlpha As Byte
        PsdNoAlphaForNonLayer As Byte
        PngComposeWithAlpha As Byte
        WMFHighResolution As Byte
        Width As Long
        Height As Long
        Offset As Long
        ChannelOrder As Integer
        ChannelType As Integer
        PcdBase As Integer
        EpsDpi As Integer
        EpsWidth As Long
        EpsHeight As Long
                LutType As Integer
                CompressRation As Integer
                MaxFileSize As Long
                LutData As Long
                LutFilename As Long

                CameraRawUseAutomaticBalance As Byte
                CameraRawUseCameraBalance As Byte
                Reserved4 As Integer
                CameraRawGamma As Long
                CameraRawBrightness As Long
                CameraRawRedScaling As Long
                CameraRawBlueScaling As Long
                CameraRawFilterDomain As Long
                CameraRawFilterRange As Long
                
                Callbacks As LOAD_CALLBACK_STRUCT
                UserParams As Long
End Type

'Save Params Struct
Public Const GFL_SAVE_REPLACE_EXTENSION = &H1
Public Const GFL_SAVE_WANT_FILENAME = &H2
Public Const GFL_SAVE_ANYWAY = &H4

Public Type GFL_SAVE_PARAMS
    flags As Long
    FormatIndex As Long
    Compression As Integer
    Quality As Integer
    CompressionLevel As Integer
    Interlaced As Byte
    Progressive As Byte
    OptimizeHuffmanTable As Byte
    InAscii As Byte
        LutType As Integer
        Reserved As Integer
    MaxFileSize As Long
        LutData As Long
        LutFilename As Long
    Offset As Long
    ChannelOrder As Integer
    ChannelType As Integer
    Callbacks As SAVE_CALLBACK_STRUCT
    UserParams As Long
End Type

Public Enum GFL_COLORMODEL
    GFL_CM_RGB = 0
    GFL_CM_GREY = 1
    GFL_CM_CMY = 2
    GFL_CM_CMYK = 3
    GFL_CM_YCBCR = 4
    GFL_CM_YUV16 = 5
    GFL_CM_LAB = 6
    GFL_CM_LOGLUV = 7
    GFL_CM_LOGL = 8
End Enum

'File Information Struct
Public Type GFL_FILE_INFORMATION
    Type As Integer
    Origin As Integer
    Width As Long
    Height As Long
    FormatIndex As Long
    FormatName As String * 8
    Description As String * 64
    Xdpi As Integer
    Ydpi As Integer
    BitsPerComponent As Integer
    ComponentsPerPixel As Integer
    NumberOfImages As Long
    FileSize As Long
    ColorModel As Integer
    Compression As Integer
    CompressionDescription As String * 64
        XOffset As Long
        YOffset As Long
        ExtraInfos As Long
End Type

Public Const GFL_READ = &H1
Public Const GFL_WRITE = &H2

'Format Information Scruct
Public Type GFL_FORMAT_INFORMATION
    Index As Long
    Name As String * 8
    Description As String * 64
    Status As Long
    NumberOfExtension As Long
    Extension(16) As String * 8
End Type


Public Declare Function gflMemoryAlloc Lib "libgfl340.dll" (ByVal Size As Long) As Long
Public Declare Function gflMemoryRealloc Lib "libgfl340.dll" (ByVal Ptr As Long, ByVal Size As Long) As Long
Public Declare Sub gflMemoryFree Lib "libgfl340.dll" (ByVal Ptr As Long)

Public Declare Function gflGetVersion Lib "libgfl340.dll" () As Long
Public Declare Function gflGetVersionOfLibformat Lib "libgfl340.dll" () As Long

Public Declare Function gflLibraryInit Lib "libgfl340.dll" () As Integer
Public Declare Sub gflLibraryExit Lib "libgfl340.dll" ()
Public Declare Sub gflEnableLZW Lib "libgfl340.dll" (ByVal Enable As Byte)
Public Declare Sub gflSetPluginsPathname Lib "libgfl340.dll" (ByVal Path As String)

'Bitmap = pointeur vers GFL_BITMAP. Pas nécéssaire de l'initialiser avant : il peut contenir 0

Public Declare Function gflGetNumberOfFormat Lib "libgfl340.dll" () As Long
Public Declare Function gflGetFormatIndexByName Lib "libgfl340.dll" (ByVal Name As String) As Long
Public Declare Function gflGetFormatNameByIndex Lib "libgfl340.dll" (ByVal Index As Long) As Long
Public Declare Function gflFormatIsSupported Lib "libgfl340.dll" (ByVal Name As String) As Byte
Public Declare Function gflFormatIsWritableByIndex Lib "libgfl340.dll" (ByVal Index As Long) As Byte
Public Declare Function gflFormatIsWritableByName Lib "libgfl340.dll" (ByVal Name As String) As Byte
Public Declare Function gflFormatIsReadableByIndex Lib "libgfl340.dll" (ByVal Index As Long) As Byte
Public Declare Function gflFormatIsReadableByName Lib "libgfl340.dll" (ByVal Name As String) As Byte
Public Declare Function gflGetDefaultFormatSuffixByIndex Lib "libgfl340.dll" (ByVal Index As Long) As Long
Public Declare Function gflGetDefaultFormatSuffixByName Lib "libgfl340.dll" (ByVal Name As String) As Long
Public Declare Function gflGetFormatDescriptionByIndex Lib "libgfl340.dll" (ByVal Index As Long) As Long
Public Declare Function gflGetFormatDescriptionByName Lib "libgfl340.dll" (ByVal Name As String) As Long
Public Declare Function gflGetFormatInformationByName Lib "libgfl340.dll" (ByVal Name As String, ByRef info As GFL_FORMAT_INFORMATION) As Integer
Public Declare Function gflGetFormatInformationByIndex Lib "libgfl340.dll" (ByVal Index As Long, ByRef info As GFL_FORMAT_INFORMATION) As Integer

Public Declare Function gflGetErrorString Lib "libgfl340.dll" (ByVal Error As Integer) As Long
Public Declare Function gflGetLabelForColorModel Lib "libgfl340.dll" (ByVal color_model As Integer) As Long
Public Declare Function gflGetFileInformation Lib "libgfl340.dll" (ByVal FileName As String, ByVal Index As Long, ByRef info As GFL_FILE_INFORMATION) As Integer
Public Declare Sub gflFreeFileInformation Lib "libgfl340.dll" (ByRef info As GFL_FILE_INFORMATION)
Public Declare Sub gflGetDefaultLoadParams Lib "libgfl340.dll" (ByRef params As GFL_LOAD_PARAMS)
Public Declare Function gflLoadBitmap Lib "libgfl340.dll" (ByVal FileName As String, ByRef BITMAP As Long, ByRef params As GFL_LOAD_PARAMS, ByRef info As GFL_FILE_INFORMATION) As Integer
Public Declare Sub gflGetDefaultPreviewParams Lib "libgfl340.dll" (ByRef params As GFL_LOAD_PARAMS)
Public Declare Function gflLoadPreview Lib "libgfl340.dll" (ByVal FileName As String, ByVal Width As Long, ByVal Height As Long, ByRef BITMAP As Long, ByRef params As GFL_LOAD_PARAMS, ByRef info As GFL_FILE_INFORMATION) As Integer
Public Declare Sub gflGetDefaultSaveParams Lib "libgfl340.dll" (ByRef params As GFL_SAVE_PARAMS)
Public Declare Function gflSaveBitmap Lib "libgfl340.dll" (ByVal FileName As String, ByRef BITMAP As GFL_BITMAP, ByRef params As GFL_SAVE_PARAMS) As Integer

Public Declare Function gflSaveBitmapIntoHandle Lib "libgfl340.dll" (ByVal Handle As Long, ByRef BITMAP As GFL_BITMAP, ByRef params As GFL_SAVE_PARAMS) As Integer
Public Declare Function gflLoadPreviewFromHandle Lib "libgfl340.dll" (ByVal Handle As Long, ByVal Width As Long, ByVal Height As Long, ByRef BITMAP As Long, ByRef params As GFL_LOAD_PARAMS, ByRef info As GFL_FILE_INFORMATION) As Integer
Public Declare Function gflLoadBitmapFromHandle Lib "libgfl340.dll" (ByVal Handle As Long, ByRef BITMAP As Long, ByRef info As GFL_LOAD_PARAMS, ByRef info As GFL_FILE_INFORMATION) As Integer

Public Declare Function gflFileCreate Lib "libgfl340.dll" (ByRef Handle As Long, ByVal FileName As String, ByVal image_count As Long, ByRef params As GFL_SAVE_PARAMS) As Integer
Public Declare Function gflFileAddPicture Lib "libgfl340.dll" (ByVal Handle As Long, ByRef BITMAP As GFL_BITMAP) As Integer
Public Declare Sub gflFileClose Lib "libgfl340.dll" (ByVal Handle As Long)

Public Declare Function gflCloneBitmap Lib "libgfl340.dll" (ByRef BITMAP As GFL_BITMAP) As Long
Public Declare Function gflAllockBitmap Lib "libgfl340.dll" (ByVal bmptype As Integer, ByVal Width As Long, ByVal Height As Long, ByVal Line_padding As Long, ByRef Color As GFL_COLOR) As Long
Public Declare Sub gflFreeBitmap Lib "libgfl340.dll" (ByRef BITMAP As GFL_BITMAP)
Public Declare Sub gflFreeBitmapLong Lib "libgfl340.dll" Alias "gflFreeBitmap" (ByVal BITMAP As Long)
Public Declare Sub gflFreeBitmapData Lib "libgfl340.dll" (ByRef BITMAP As GFL_BITMAP)
Public Declare Sub gflFreeBitmapDataLong Lib "libgfl340.dll" Alias "gflFreeBitmapData" (ByVal BITMAP As Long)

Public Declare Function gflLoadBitmapFromMemory Lib "libgfl340.dll" (ByRef data As Byte, ByVal dat_length As Long, ByRef BITMAP As Long, ByRef params As GFL_LOAD_PARAMS, ByRef info As GFL_FILE_INFORMATION) As Integer
Public Declare Function gflLoadEXIF Lib "libgfl340.dll" (ByVal FileName As String, ByVal flags As Long) As Long

Public Const GFL_RESIZE_QUICK = 0
Public Const GFL_RESIZE_BILINEAR = 1
Public Const GFL_RESIZE_HERMITE = 2
Public Const GFL_RESIZE_GAUSSIAN = 3
Public Const GFL_RESIZE_BELL = 4
Public Const GFL_RESIZE_BSPLINE = 5
Public Const GFL_RESIZE_MITSHELL = 6
Public Const GFL_RESIZE_LANCZOS = 7

Public Enum GFL_MODE
    GFL_MODE_TO_BINARY = 1
    GFL_MODE_TO_4GREY = 2
    GFL_MODE_TO_8GREY = 3
    GFL_MODE_TO_16GREY = 4
    GFL_MODE_TO_32GREY = 5
    GFL_MODE_TO_64GREY = 6
    GFL_MODE_TO_128GREY = 7
    GFL_MODE_TO_216GREY = 8
    GFL_MODE_TO_256GREY = 9
    GFL_MODE_TO_8COLORS = 12
    GFL_MODE_TO_16COLORS = 13
    GFL_MODE_TO_32COLORS = 14
    GFL_MODE_TO_64COLORS = 15
    GFL_MODE_TO_128COLORS = 16
    GFL_MODE_TO_216COLORS = 17
    GFL_MODE_TO_256COLORS = 18
    GFL_MODE_TO_RGB = 19
    GFL_MODE_TO_RGBA = 20
    GFL_MODE_TO_BGR = 21
    GFL_MODE_TO_ABGR = 22
    GFL_MODE_TO_BGRA = 23
    GFL_MODE_TO_ARGB = 24
End Enum

Public Enum GFL_MODE_PARAMS
    GFL_MODE_NO_DITHER = 0
    GFL_MODE_PATTERN_DITHER = 1
    GFL_MODE_HALTONE45_DITHER = 2  ' Only with GFL_MODE_TO_BINARY
    GFL_MODE_HALTONE90_DITHER = 3  ' Only with GFL_MODE_TO_BINARY
    GFL_MODE_ADAPTIVE = 4
    GFL_MODE_FLOYD_STEINBERG = 5   ' Only with GFL_MODE_TO_BINARY
End Enum

Type GFL_RECT
    x As Long
    y As Long
    w As Long
    h As Long
End Type

Type GFL_POINT
    x As Long
    y As Long
End Type

Public Declare Function gflSetColorAt Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByVal x As Long, ByVal y As Long, ByRef Color As GFL_COLOR) As Integer
Public Declare Function gflGetColorAt Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByVal x As Long, ByVal y As Long, ByRef Color As GFL_COLOR) As Integer
Public Declare Function gflChangeColorDepth Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByRef dst As Long, ByVal Mode As GFL_MODE, ByVal params As GFL_MODE_PARAMS) As Integer
Public Declare Function gflResize Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByRef dst As Long, ByVal Width As Long, ByVal Height As Long, ByVal method As Long, ByVal flags As Long) As Integer
Public Declare Function gflReplaceColor Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByRef dst As Long, ByRef Color As GFL_COLOR, ByRef new_color As GFL_COLOR, ByVal Tolerance As Long) As Integer

'dst = pointeur vers GFL_BITMAP
Public Declare Function gflFlipVertical Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByRef dst As Long) As Integer
Public Declare Function gflFlipHorizontal Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByRef dst As Long) As Integer
Public Declare Function gflCrop Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByRef dst As Long, ByRef Rect As GFL_RECT) As Integer

Public Enum GFL_CANVASRESIZE
    GFL_CANVASRESIZE_CENTER = 0
    GFL_CANVASRESIZE_TOPLEFT = 1
    GFL_CANVASRESIZE_TOPRIGHT = 2
    GFL_CANVASRESIZE_BOTTOMLEFT = 3
    GFL_CANVASRESIZE_BOTTOMRIGHT = 4
    GFL_CANVASRESIZE_TOP = 5
    GFL_CANVASRESIZE_BOTTOM = 6
    GFL_CANVASRESIZE_LEFT = 7
    GFL_CANVASRESIZE_RIGHT = 8
End Enum

Public Declare Function gflResizeCanvas Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByRef dst As Long, ByVal Width As Long, ByVal Height As Long, ByVal Mode As GFL_CANVASRESIZE, ByRef Color As GFL_COLOR) As Integer
Public Declare Function gflRotate Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByRef dst As Long, ByVal angle As Long, ByRef Color As GFL_COLOR) As Integer
Public Declare Function gflRotateFine Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByRef dst As Long, ByVal angle As Double, ByRef Color As GFL_COLOR) As Integer
Public Declare Function gflBitblt Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByRef Rect As GFL_RECT, ByRef dst As GFL_BITMAP, ByVal x_dest As Long, ByVal y_dest As Long) As Integer
Public Declare Function gflBitbltEx Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByRef Rect As GFL_RECT, ByRef dst As GFL_BITMAP, ByVal x_dest As Long, ByVal y_dest As Long) As Integer
Public Declare Function gflAutoCrop Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByRef dst As Long, ByRef Color As GFL_COLOR, ByVal Tolerance As Long) As Integer
Public Declare Function gflMerge Lib "libgfl340.dll" (ByRef src As Long, ByRef Origin As GFL_POINT, ByRef Opacity As Long, ByVal num_bitmap As Long, ByRef dst As Long) As Integer

Public Const GFL_IPTC_BYLINE = &H50
Public Const GFL_IPTC_BYLINETITLE = &H55
Public Const GFL_IPTC_CREDITS = &H6E
Public Const GFL_IPTC_SOURCE = &H73
Public Const GFL_IPTC_CAPTIONWRITER = &H7A
Public Const GFL_IPTC_CAPTION = &H78
Public Const GFL_IPTC_HEADLINE = &H69
Public Const GFL_IPTC_SPECIALINSTRUCTIONS = &H28
Public Const GFL_IPTC_OBJECTNAME = &H5
Public Const GFL_IPTC_DATECREATED = &H37
Public Const GFL_IPTC_RELEASEDATE = &H1E
Public Const GFL_IPTC_TIMECREATED = &H3C
Public Const GFL_IPTC_RELEASETIME = &H23
Public Const GFL_IPTC_CITY = &H5A
Public Const GFL_IPTC_STATE = &H5F
Public Const GFL_IPTC_COUNTRY = &H65
Public Const GFL_IPTC_COUNTRYCODE = &H64
Public Const GFL_IPTC_SUBLOCATION = &H5C
Public Const GFL_IPTC_ORIGINALTRREF = &H67
Public Const GFL_IPTC_CATEGORY = &HF
Public Const GFL_IPTC_COPYRIGHT = &H74
Public Const GFL_IPTC_EDITSTATUS = &H7
Public Const GFL_IPTC_PRIORITY = &HA
Public Const GFL_IPTC_OBJECTCYCLE = &H4B
Public Const GFL_IPTC_JOBID = &H16
Public Const GFL_IPTC_PROGRAM = &H41
Public Const GFL_IPTC_KEYWORDS = &H19
Public Const GFL_IPTC_SUPCATEGORIES = &H14

Public Declare Function gflBitmapGetIPTCValue Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByVal id As Long, ByVal Name As String, ByVal value_length As Long) As Integer

Public Const GFL_EXIF_MAKER = "&H010F"
Public Const GFL_EXIF_MODEL = "&H0110"
Public Const GFL_EXIF_ORIENTATION = "&H0112"
Public Const GFL_EXIF_EXPOSURETIME = "&H829A"
Public Const GFL_EXIF_FNUMBER = "&H829D"
Public Const GFL_EXIF_DATETIME_ORIGINAL = "&H9003"
Public Const GFL_EXIF_SHUTTERSPEED = "&H9201"
Public Const GFL_EXIF_APERTURE = "&H9202"
Public Const GFL_EXIF_MAXAPERTURE = "&H9205"
Public Const GFL_EXIF_FOCALLENGTH = "&H920A"

Public Declare Function gflBitmapGetEXIFValue Lib "libgfl340.dll" (ByRef src As GFL_BITMAP, ByVal tag As Long, ByVal Name As String, ByVal value_length As Long) As Integer

