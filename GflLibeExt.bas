Attribute VB_Name = "GflLibeExt"
' Graphics File Library Extended
'
' GFL library Copyright (c) 1991-2002 Pierre-e Gougelet
' All rights reserved
' Commercial use is not authorized without agreement
'
' Interface for Visual Basic : Jérôme Quintard (contact@jeromequintard.com)

Option Explicit

'Sert à copier la mémoire entre 2 pointeurs
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDest As Long, ByVal lpvSource As Long, ByVal cbCopy As Long)

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long

Private Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(0 To 255) As RGBQUAD
End Type

Private Const BI_RGB = 0&
Private Const CBM_INIT = &H4
Private Const DIB_RGB_COLORS = 0

Private Const SRCCOPY = &HCC0020
Private Const SRCAND = &H8800C6
Private Const SRCPAINT = &HEE0086

Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByVal lpBits As Long, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long


Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long '
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Type bitmap
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type

'Fonction pour retourner une DIB d'un pointeur
Public Sub extGetDIBfromPtr(ByVal src As Long, ByRef dest As BITMAPINFO)
    CopyMemory VarPtr(dest), src, Len(dest)
End Sub

'Affiche un bitmap sur un DC en passant par ConvertBitmapInfoDIB
Public Function extShowBitmapOnDCEx(ByRef bitmap As GFL_BITMAP, ByVal hdc As Long)
Dim DibInfo As BITMAPINFO
Dim hndDib As Long
Dim ptrDib As Long
Dim Data As Long
Dim Ret As Integer
Ret = gflConvertBitmapIntoDIB(bitmap, hndDib) 'On converti le bitmap en DIB

If Ret = GFL_NO_ERROR Then 'Si aucune erreur
    ptrDib = GlobalLock(hndDib) 'On récupère le pointeur sur l'handle DIB renvoyé par gflConvertBitmapIntoDIB
        extGetDIBfromPtr ptrDib, DibInfo 'On récupère le DIB depuis le pointeur
        If bitmap.Type <> GFL_BINARY And bitmap.Type <> GFL_GREY And bitmap.Type <> GFL_COLORS Then
            Data = ptrDib + Len(DibInfo.bmiHeader) 'Si c'est un RGB on indique où trouver ses données
        Else 'Sinon
            Data = ptrDib + (DibInfo.bmiHeader.biClrUsed * 4) + Len(DibInfo.bmiHeader) 'On indique où trouver ses données
        End If
        GlobalUnlock ptrDib 'On libère le pointeur puis on dessine le bitmap sur le contexte
    StretchDIBits hdc, 0, 0, bitmap.Width, bitmap.Height, 0, 0, bitmap.Width, bitmap.Height, Data, DibInfo, DIB_RGB_COLORS, SRCCOPY
End If
End Function

Public Sub extShowTransparencyBitmapOnDCEx(ByRef bitmap As GFL_BITMAP, ByVal hdc As Long, ByRef Color As OLE_COLOR)
Dim hBitmap As Long, hBmpMask As Long

    gflConvertBitmapIntoDDB bitmap, hBitmap
    hBmpMask = PrepareMask(hBitmap, GetColor(bitmap, Color))
    DrawTransparentBitmap hdc, 0, 0, bitmap.Width, bitmap.Height, hBitmap, hBmpMask, 0, 0
    DeleteObject hBmpMask
    DeleteObject hBitmap
End Sub

Private Function PrepareMask(hBmpSource As Long, clrpTransColor As OLE_COLOR) As Long
Dim bm As bitmap
Dim hdcSrc As Long, hdcDst As Long
Dim hbmSrcT As Long, hbmDstT As Long, hBmpMask As Long
Dim clrTrans As OLE_COLOR, clrSaveBk As OLE_COLOR, clrSaveDstText As OLE_COLOR

   GetObject hBmpSource, Len(bm), bm
   hBmpMask = CreateBitmap(bm.bmWidth, bm.bmHeight, 1, 1, &H0)

   hdcSrc = CreateCompatibleDC(&H0)
   hdcDst = CreateCompatibleDC(&H0)

   hbmSrcT = SelectObject(hdcSrc, hBmpSource)
   hbmDstT = SelectObject(hdcDst, hBmpMask)

   clrTrans = clrpTransColor

   clrSaveBk = SetBkColor(hdcSrc, clrTrans)

   BitBlt hdcDst, 0, 0, bm.bmWidth, bm.bmHeight, hdcSrc, 0, 0, SRCCOPY

   clrSaveDstText = SetTextColor(hdcSrc, RGB(255, 255, 255))
   SetBkColor hdcSrc, RGB(0, 0, 0)

   BitBlt hdcSrc, 0, 0, bm.bmWidth, bm.bmHeight, hdcDst, 0, 0, SRCAND

   SetTextColor hdcDst, clrSaveDstText

   SetBkColor hdcSrc, clrSaveBk
   SelectObject hdcSrc, hbmSrcT
   SelectObject hdcDst, hbmDstT

   DeleteDC hdcSrc
   DeleteDC hdcDst

PrepareMask = hBmpMask
End Function

Private Sub DrawTransparentBitmap(hdc As Long, xstart As Long, ystart As Long, wwidth As Long, wheight As Long, hBitmap As Long, hBmpMask As Long, xsource As Long, ysource As Long)
Dim hMemoryDC As Long, holdDC As Long
Dim hBmpMemoryDC As Long, hBmpOldDc As Long

    hBmpMemoryDC = CreateCompatibleDC(hdc)
    hBmpOldDc = SelectObject(hBmpMemoryDC, hBitmap)

    hMemoryDC = CreateCompatibleDC(&H0)
    holdDC = SelectObject(hMemoryDC, hBmpMask)

    BitBlt hdc, xstart, ystart, wwidth, wheight, hMemoryDC, xsource, ysource, SRCAND
    BitBlt hdc, xstart, ystart, wwidth, wheight, hBmpMemoryDC, xsource, ysource, SRCPAINT

    SelectObject hMemoryDC, holdDC
    DeleteDC hMemoryDC

    SelectObject hBmpMemoryDC, hBmpOldDc
    DeleteDC hBmpMemoryDC
End Sub

Private Function GetColor(ByRef bitmap As GFL_BITMAP, ByVal Color As OLE_COLOR) As OLE_COLOR
Dim GflColor As GFL_COLOR
Dim GflColorMap As GFL_COLORMAP
Dim Ind As Integer
Dim ColorIndex As Byte
Dim Min_D2 As Double
Dim D2 As Long
    
    
GflColor.Red = Color And &HFF&                  'On la récupère valeur du rouge
GflColor.Green = (Color And &HFF00&) \ &H100&   'On la récupère valeur du vert
GflColor.Blue = (Color And &HFF0000) \ &H10000  'On la récupère valeur du bleu
    
Min_D2 = &H7FFFFFF 'Le masque pour le minimum de distance au carré
    
For Ind = 0 To bitmap.ColorUsed - 1 'Pour chaque index
    If bitmap.ColorMap <> &H0 Then 'Si il y a une color map
        extGetGflColorMapFromPtr bitmap.ColorMap, GflColorMap 'On l'extrait
        'Pour chaque index on cherche la distance au carré entre les différentes valeurs
        D2 = (CLng(GflColorMap.Red(Ind)) - CLng(GflColor.Red)) * (CLng(GflColorMap.Red(Ind)) - CLng(GflColor.Red)) _
           + (CLng(GflColorMap.Green(Ind)) - CLng(GflColor.Green)) * (CLng(GflColorMap.Green(Ind)) - CLng(GflColor.Green)) _
           + (CLng(GflColorMap.Blue(Ind)) - CLng(GflColor.Blue)) * (CLng(GflColorMap.Blue(Ind)) - CLng(GflColor.Blue))
    Else
        'Pour niveau de gris on cherche la distance au carré entre les différentes valeurs
        D2 = (Ind - GflColor.Red) * (Ind - GflColor.Red) _
           + (Ind - GflColor.Green) * (Ind - GflColor.Green) _
           + (Ind - GflColor.Blue) * (Ind - GflColor.Blue)
    End If
    If D2 < Min_D2 Then Min_D2 = D2: ColorIndex = Ind 'On indique la plus petite distance par rapport à notre couleur
Next
    
GetColor = RGB(GflColorMap.Red(ColorIndex), GflColorMap.Green(ColorIndex), GflColorMap.Blue(ColorIndex))
End Function


