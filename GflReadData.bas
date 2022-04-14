Attribute VB_Name = "GflReadData"
'1.76   FIX     Correction de READ_SeekFunction

Public Type READ_DATA
    Data() As Byte
    Index As Long
    Length As Long
End Type

Public RData As READ_DATA

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDest As Long, ByVal lpvSource As Long, ByVal cbCopy As Long)

'Lit un fichier dans une structure READ_DATA
Public Function extLoadFile(File As String, Data As READ_DATA)
Dim ff As Integer

ff = FreeFile
Open File For Binary As ff
    ReDim Data.Data(0 To LOF(ff) - 1) As Byte
    Data.Length = LOF(ff)
    Data.Index = 0
    Get #ff, , Data.Data
Close #ff
End Function

'Fonction pour définir des Data sur un pointeur
Public Sub extSetDataToPtr(ByVal src As Long, ByVal dest As Long, ByVal Size As Long)
     CopyMemory src, dest, Size
End Sub

Public Function READ_ReadFunction(ByRef RData As READ_DATA, ByVal Buffer As Long, ByVal Size As Long) As Long
    If (RData.Index + Size) >= RData.Length Then Size = RData.Length - RData.Index
    extSetDataToPtr Buffer, VarPtr(RData.Data(0)) + RData.Index, Size
    RData.Index = RData.Index + Size
    READ_ReadFunction = Size
End Function
    
Public Function READ_TellFunction(ByRef RData As READ_DATA) As Long
   READ_TellFunction = RData.Index
End Function

Public Function READ_SeekFunction(ByRef RData As READ_DATA, ByVal Offset As Long, ByVal Origin As Long) As Long
    If Origin = 1 Then Offset = Offset + RData.Index
    If Origin = 2 Then Offset = RData.Length - RData.Index
    If Offset > RData.Length Then READ_SeekFunction = -1
    RData.Index = Offset
    READ_SeekFunction = RData.Index
End Function
