Attribute VB_Name = "GflLibe"
' Graphics File Library Extended
'
' GFL library Copyright (c) 1991-2007 Pierre-e Gougelet
' All rights reserved
' Commercial use is not authorized without agreement
'
' Interface for Visual Basic : Jérôme Quintard (contact@jeromequintard.com)

Option Explicit

Public Declare Function gflGetNumberOfColorsUsed Lib "libgfle340.dll" (ByRef src As GFL_BITMAP) As Long

Public Declare Function gflNegative Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long) As Integer
Public Declare Function gflBrightness Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal Brightness As Long) As Integer
Public Declare Function gflContrast Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal Contrast As Long) As Integer
Public Declare Function gflGamma Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal Gamma As Double) As Integer
Public Declare Function gflLogCorrection Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long) As Integer
Public Declare Function gflNormalize Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long) As Integer
Public Declare Function gflEqualize Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long) As Integer
Public Declare Function gflEqualizeOnLuminance Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long) As Integer
Public Declare Function gflBalance Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByRef Color As GFL_COLOR) As Integer
Public Declare Function gflAdjust Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal Brightness As Long, ByVal Contrast As Long, ByVal Gamma As Double) As Integer
Public Declare Function gflAdjustHLS Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal h_increment As Long, ByVal l_increment As Long, ByVal s_increment As Long) As Integer
Public Declare Function gflSepia Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal Percent As Long) As Integer
Public Declare Function gflSepiaExt Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal Percent As Long, ByRef Color As GFL_COLOR) As Integer

Public Declare Function gflAverage Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal filter_size As Long) As Integer
Public Declare Function gflSoften Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal percentage As Long) As Integer
Public Declare Function gflBlur Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal percentage As Long) As Integer
Public Declare Function gflGaussianBlur Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal filter_size As Long) As Integer
Public Declare Function gflMaximum Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal filter_size As Long) As Integer
Public Declare Function gflMinimum Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal filter_size As Long) As Integer
Public Declare Function gflMedianBox Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal filter_size As Long) As Integer
Public Declare Function gflMedianCross Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal filter_size As Long) As Integer
Public Declare Function gflSharpen Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal percentage As Long) As Integer

Public Declare Function gflEnhanceDetail Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long) As Integer
Public Declare Function gflEnhanceFocus Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long) As Integer
Public Declare Function gflFocusRestoration Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long) As Integer
Public Declare Function gflEdgeDetectLight Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long) As Integer
Public Declare Function gflEdgeDetectMedium Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long) As Integer
Public Declare Function gflEdgeDetectHeavy Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long) As Integer
Public Declare Function gflEmboss Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long) As Integer
Public Declare Function gflEmbossMore Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long) As Integer

Public Type GFL_FILTER
    Size As Integer
    Matrix(48) As Integer
    Divisor As Integer
    Bias As Integer
End Type

Public Declare Function gflConvolve Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByRef filter As GFL_FILTER) As Integer

Public Enum GFL_SWAPCOLORS_MODE
    GFL_SWAPCOLORS_RBG = 0
    GFL_SWAPCOLORS_BGR = 1
    GFL_SWAPCOLORS_BRG = 2
    GFL_SWAPCOLORS_GRB = 3
    GFL_SWAPCOLORS_GBR = 4
End Enum

Public Declare Function gflSwapColors Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Dst As Long, ByVal Mode As GFL_SWAPCOLORS_MODE) As Integer

Public Enum GFL_LOSSLESS_TRANSFORM
    GFL_LOSSLESS_TRANSFORM_NONE = 0
    GFL_LOSSLESS_TRANSFORM_ROTATE90 = 1
    GFL_LOSSLESS_TRANSFORM_ROTATE180 = 2
    GFL_LOSSLESS_TRANSFORM_ROTATE270 = 3
    GFL_LOSSLESS_TRANSFORM_VERTICAL_FLIP = 4
    GFL_LOSSLESS_TRANSFORM_HORIZONTAL_FLIP = 5
End Enum

Public Declare Function gflJpegLosslessTransform Lib "libgfle340.dll" (ByVal filename As String, ByVal transform As GFL_LOSSLESS_TRANSFORM) As Integer

Public Declare Function gflConvertBitmapIntoDIB Lib "libgfle340.dll" (ByRef bitmap As GFL_BITMAP, ByRef hDib As Long) As Integer
Public Declare Function gflConvertBitmapIntoDDB Lib "libgfle340.dll" (ByRef bitmap As GFL_BITMAP, ByRef hBitmap As Long) As Integer
Public Declare Function gflConvertDIBIntoBitmap Lib "libgfle340.dll" (ByVal hDib As Long, ByRef bitmap As Long) As Integer
Public Declare Function gflConvertDDBIntoBitmap Lib "libgfle340.dll" (ByVal hBitmap As Long, ByRef bitmap As Long) As Integer

Public Declare Function gflLoadBitmapIntoDIB Lib "libgfle340.dll" (ByVal filename As String, ByRef hDib As Long, ByRef params As GFL_LOAD_PARAMS, ByRef info As GFL_FILE_INFORMATION) As Integer
Public Declare Function gflLoadBitmapIntoDDB Lib "libgfle340.dll" (ByVal filename As String, ByRef hBitmap As Long, ByRef params As GFL_LOAD_PARAMS, ByRef info As GFL_FILE_INFORMATION) As Integer

Public Declare Function gflAddText Lib "libgfle340.dll" (ByRef bitmap As GFL_BITMAP, ByVal text As String, ByVal font_name As String, ByVal x As Long, ByVal y As Long, ByVal font_size As Long, ByVal orientation As Long, ByVal italic As Byte, ByVal bold As Byte, ByVal strike_out As Byte, ByVal underline As Byte, ByVal antialias As Byte, ByRef Color As GFL_COLOR) As Integer

Public Declare Function gflImportFromClipboard Lib "libgfle340.dll" (ByRef bitmap As Long) As Integer
Public Declare Function gflExportIntoClipboard Lib "libgfle340.dll" (ByRef bitmap As GFL_BITMAP) As Integer
Public Declare Function gflImportFromHWND Lib "libgfle340.dll" (ByVal hwnd As Long, ByRef Rect As GFL_RECT, ByRef bitmap As Long) As Integer

Public Enum GFL_LINE_STYLE
    GFL_LINE_STYLE_SOLID = 0
    GFL_LINE_STYLE_DASH = 1
    GFL_LINE_STYLE_DOT = 2
    GFL_LINE_STYLE_DASHDOT = 3
    GFL_LINE_STYLE_DASHDOTDOT = 4
End Enum

Public Declare Function gflDrawPointColor Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByVal x As Long, ByVal y As Long, ByVal line_width As Long, ByRef line_color As GFL_COLOR, ByRef Dst As Long) As Integer
Public Declare Function gflDrawLineColor Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByVal X0 As Long, ByVal Y0 As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal line_width As Long, ByRef line_color As GFL_COLOR, ByVal line_style As Long, ByRef Dst As Long) As Integer
Public Declare Function gflDrawPolylineColor Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Points As GFL_POINT, ByVal num_points As Long, ByVal line_width As Long, ByRef line_color As GFL_COLOR, ByVal line_style As Long, ByRef Dst As Long) As Integer
Public Declare Function gflDrawRectangleColor Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByVal X0 As Long, ByVal Y0 As Long, ByVal Width As Long, ByVal Height As Long, ByRef fill_color As GFL_COLOR, ByVal line_width As Long, ByRef line_color As GFL_COLOR, ByVal line_style As Long, ByRef Dst As Long) As Integer
Public Declare Function gflDrawPolygonColor Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByRef Points As GFL_POINT, ByVal num_points As Long, ByRef fill_color As GFL_COLOR, ByVal line_width As Long, ByVal line_color As Long, ByVal line_style As Long, ByRef Dst As Long) As Integer
Public Declare Function gflDrawCircleColor Lib "libgfle340.dll" (ByRef src As GFL_BITMAP, ByVal x As Long, ByVal y As Long, ByVal Radius As Long, ByVal fill_color As Long, ByVal line_width As Long, ByVal line_color As Long, ByVal line_style As Long, ByRef Dst As Long) As Integer
