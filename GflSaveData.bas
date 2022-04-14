Attribute VB_Name = "GflSaveData"
Public Type SAVE_DATA
    CurrentPosition As Long
    CurrentSize As Long
    CurrentAllocatedSize As Long
    Data As Long
End Type

Public SData As SAVE_DATA 'La structure qui recevra les données en transition

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDest As Long, ByVal lpvSource As Long, ByVal cbCopy As Long)

'Ecrit un fichier depuis une structure SAVE_DATA
Public Function extSaveFile(File As String, Data As SAVE_DATA)
Dim Buffer() As Byte
Dim ff As Integer

ff = FreeFile
Open File For Binary As ff
    ReDim Buffer(0 To Data.CurrentSize) As Byte
    CopyMemory VarPtr(Buffer(0)), Data.Data, Data.CurrentSize
    Put ff, , Buffer
Close
End Function

'Ecrit les données en mémoire d'une structure SAVE_DATA dans un tableau de byte
Public Function extWriteArray(Data As SAVE_DATA) As Byte()
Dim Buffer() As Byte
ReDim Buffer(0 To Data.CurrentSize) As Byte
CopyMemory VarPtr(Buffer(0)), Data.Data, Data.CurrentSize
extWriteArray = Buffer
End Function

'Fonction pour retourner des Data depuis un pointeur
Public Sub extGetDataFromPtr(ByVal dest As Long, ByVal Buffer As Long, ByVal Size As Long)
     CopyMemory dest, Buffer, Size
End Sub

'Renvoie la plus grande valeur des deux
Public Function extMax(bsize, Size) As Long
If bsize >= Size Then extMax = bsize Else extMax = Size
End Function

Public Function WRITE_WriteFunction(ByRef SData As SAVE_DATA, ByVal Buffer As Long, ByVal Size As Long) As Long
With SData
    If .CurrentPosition + Size >= .CurrentAllocatedSize Then
        .CurrentAllocatedSize = .CurrentAllocatedSize + extMax(16384, Size)
        If .Data = 0 Then
            .Data = gflMemoryAlloc(.CurrentAllocatedSize)
        Else
            .Data = gflMemoryRealloc(.Data, .CurrentAllocatedSize)
        End If
        If .Data = 0 Then WRITE_WriteFunction = 0
    End If
    extGetDataFromPtr .Data + .CurrentPosition, Buffer, Size
    .CurrentPosition = .CurrentPosition + Size
    .CurrentSize = .CurrentSize + Size
    WRITE_WriteFunction = Size
End With
End Function

Public Function WRITE_TellFunction(ByRef SData As SAVE_DATA) As Long
    WRITE_TellFunction = SData.CurrentPosition
End Function

Public Function WRITE_SeekFunction(ByRef SData As SAVE_DATA, ByVal Offset As Long, ByVal Origin As Long) As Long
    If Offset >= SData.CurrentSize Then WRITE_SeekFunction = -1
    SData.CurrentPosition = Offset
    WRITE_SeekFunction = 0
End Function


