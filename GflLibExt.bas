Attribute VB_Name = "GflLibExt"
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
'Sert à convertir un pointeur en string
Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

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

Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByVal lpBits As Long, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

'Fonction pour transformer le pointeur vers GFL_BITMAP en un autre GFL_BITMAP
'src  = pointeur gflBitmap
'dest = la structure GFL_BITMAP qui reçoit les données
Public Sub extGetGflBitmapFromPtr(ByVal src As Long, ByRef dest As GFL_BITMAP)
    gflFreeBitmapData dest 'On supprime les data de notre image
    CopyMemory VarPtr(dest), src, Len(dest) 'On copie les datas du bitmap temporaire dans notre bitmap
    gflMemoryFree src 'On supprime le bitmap temporaire de la mémoire
End Sub

'Fonction pour transformer le pointeur vers GFL_COLORMAP en un autre GFL_COLORMAP
'src  = pointeur gflColorMap
'dest = la structure GFL_COLORMAP qui reçoit les données
Public Sub extGetGflColorMapFromPtr(ByVal src As Long, ByRef dest As GFL_COLORMAP)
    CopyMemory VarPtr(dest), src, Len(dest)
End Sub

'Fonction pour retourner les commentaires d'un fichier
'src  = pointeur vers le tableau
'dest = tableau dynamique STRING qui reçoit les données
'nbr  = nombre d'éléments à récupérer
Public Sub extGetGflComments(ByVal src As Long, ByRef dst() As String, ByVal nbr As Integer)
Dim PtrComments As Long
Dim NbrComments As Integer
Dim PtrComment  As Long
Dim PosComment  As Integer

PtrComments = src
NbrComments = nbr

ReDim dst(0)
For PosComment = 0 To nbr - 1
  CopyMemory VarPtr(PtrComment), PtrComments, 4
  PtrComments = PtrComments + 4
  ReDim Preserve dst(PosComment)
  dst(PosComment) = extGetStr(PtrComment)
Next PosComment
End Sub

'Fonction pour retourner un string depuis un pointeur
'src = pointeur
Public Function extGetStr(ByVal src As Long)
    Dim dest As String
    Dim lendest As Integer
    lendest = lstrlen(src)
    dest = Space(lendest)
    lstrcpy dest, src
    extGetStr = dest
End Function

'Fonction pour supprimer les caractères NULL et TRIMMER un string
'src = string
Public Function extRTN(src As String)
    extRTN = Trim(Replace(src, Chr(0), ""))
End Function

'Fonction pour récupérer l'adresse d'un procédure, renvoyée par l'instruction AddressOf
Public Function extFarProc(pfn As Long) As Long
  extFarProc = pfn
End Function

Public Function extShowBitmapOnDc(ByRef BITMAP As GFL_BITMAP, ByVal hdc As Long)
Dim DibInfo As BITMAPINFO
    DibInfo = extGetDIBFromBitmap(BITMAP)
    StretchDIBits hdc, 0, 0, BITMAP.Width, BITMAP.Height, 0, 0, BITMAP.Width, BITMAP.Height, BITMAP.Data, DibInfo, DIB_RGB_COLORS, SRCCOPY
End Function

Public Function scanalign(pwidth&) As Long
	scanalign = (pwidth& + 3) And &HFFFFFFFC
End Function

Public Function byteperscanline(ByVal pwidth&, ByVal bitcount&) As Long
	Select Case bitcount&
	Case 1
	byteperscanline = scanalign((pwidth& + 7) \ 8)
	Case 4
	byteperscanline = scanalign((pwidth& + 1) \ 2)
	Case 8
	byteperscanline = scanalign(pwidth&)
	Case 24
	byteperscanline = scanalign(pwidth& * 3)
	Case 32
	'byteperscanline
	End Select
End Function

Public Function extGetDIBFromBitmap(ByRef BITMAP As GFL_BITMAP) As BITMAPINFO
Dim DibInfo As BITMAPINFO
Dim ColorMap As GFL_COLORMAP
Dim Ind As Integer

With DibInfo.bmiHeader
    .biSize = Len(DibInfo.bmiHeader) 'La taille de la structure
    .biWidth = BITMAP.Width 'Largeur du bitmap
    .biHeight = BITMAP.Height 'Hauteur du bitmap
    .biPlanes = 1 'Nombre de plans
    .biCompression = BI_RGB 'Toujours BI_RGB
    .biClrImportant = 0 'Toujours 0

 If BITMAP.Xdpi = BITMAP.Ydpi Then
    If BITMAP.Xdpi <> 0 Then

    .biXPelsPerMeter = CLng(BITMAP.Xdpi / 25.4 * 1000)
    .biYPelsPerMeter = CLng(BITMAP.Ydpi / 25.4 * 1000)

    Else
    .biXPelsPerMeter = 100
    .biYPelsPerMeter = 100

    End If
 Else
 .biXPelsPerMeter = CLng(BITMAP.Xdpi / 25.4 * 1000)
 .biYPelsPerMeter = CLng(BITMAP.Ydpi / 25.4 * 1000)

 End If

End With

With DibInfo
	Select Case BITMAP.Type

	Case 1
		.bmiHeader.biBitCount = 1
		.bmiHeader.biSizeImage = byteperscanline(BITMAP.Width, 1) * BITMAP.Height 'Bitmap.BytesPerLine * Bitmap.Height
		.bmiHeader.biClrUsed = 2
	Case 2, 4

		.bmiHeader.biBitCount = 8
		.bmiHeader.biSizeImage = byteperscanline(BITMAP.Width, 8) * BITMAP.Height 'GflBitmap.BytesPerLine * GflBitmap.Height
		.bmiHeader.biClrUsed = 0
	Case 64, 10
		.bmiHeader.biBitCount = 24
		.bmiHeader.biSizeImage = byteperscanline(BITMAP.Width, 24) * BITMAP.Height
		.bmiHeader.biClrUsed = 0
	End Select

    Select Case BITMAP.Type 'Suivant le type de bitmap
        Case GFL_COLORS 'Si GFL_COLORS (une palette de couleurs indexé de 0 à 255)
           extGetGflColorMapFromPtr BITMAP.ColorMap, ColorMap 'On récupère les index du bitmap dans une structure GflColorMap
            For Ind = 0 To 255 'Pour chaque index on indique à la structure Dib_Info
                .bmiColors(Ind).rgbBlue = ColorMap.Blue(Ind) 'La valeur du bleu
                .bmiColors(Ind).rgbGreen = ColorMap.Green(Ind) 'La valeur du vert
                .bmiColors(Ind).rgbRed = ColorMap.Red(Ind) 'La valeur du rouge
            Next Ind
        Case GFL_BINARY 'Si GFL_BINARY (seulement noir et blanc donc 2 couleurs indexés) on indique à Dib_Info qu'il y a deux index
            .bmiColors(0).rgbBlue = 0: .bmiColors(0).rgbGreen = 0: .bmiColors(0).rgbRed = 0 'Le premier vaut Red=0,Green=0,Blue=0 donc du noir (#000000)
            .bmiColors(1).rgbBlue = 255: .bmiColors(1).rgbGreen = 255: .bmiColors(1).rgbRed = 255 'Le second vaut Red=255,Green=255,Blue=255 donc blanc (#FFFFFF)
        Case GFL_GREY 'Si GFL_GREY (On travaille sur 256 nuances)
            For Ind = 0 To 255 'Pour chaque index de Dib_Info on indique une nuance de gris
                .bmiColors(Ind).rgbBlue = Ind  ' Il y a 256 index et 256 gris on donne à chaque index
                .bmiColors(Ind).rgbGreen = Ind ' et pour chaque canaux la valeur de la nuance
                .bmiColors(Ind).rgbRed = Ind
            Next Ind
    End Select
End With
extGetDIBFromBitmap = DibInfo
End Function
