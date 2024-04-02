Attribute VB_Name = "modAPI_vbaSquash"
                                                                                                                                                                                            ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||              vbaSquash (v1)           ||||||||||||||||||||||||||||||||||                                                                                     ' _
    ||||||||||||||||||||||||||                                       ||||||||||||||||||||||||||||||||||                                                                                     ' _
    |||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
                                                                                                                                                                                            ' _
    AUTHOR:   Kallun Willock
                                                                                                                                                                                            ' _
    URL:      MS - https://docs.microsoft.com/en-us/windows/win32/api/_cmpapi/
                                                                                                                                                                                            ' _
    NOTES:    Uses Win32 APIs in the cabinet.dll library to compress and decompress bytes.                                                                                                  ' _
              Only available on Windows 8+, and will not work on OS up to (and including) Win7.
                                                                                                                                                                                            ' _
    LICENSE:  MIT                                                                                                                                                                           ' _
                                                                                                                                                                                            ' _
    VERSION:  1.0        02/04/2024

Option Explicit

Public Enum COMPRESS_ALGORITHM_ENUM
    MSZIP = 2
    XPRESS = 3
    XPRESS_HUFF = 4
    LZMS = 5
End Enum

#If VBA7 Then
    Private Declare PtrSafe Function CreateCompressor Lib "cabinet.dll" (ByVal CompressionAlgorithm As COMPRESS_ALGORITHM_ENUM, ByVal AllocationRoutines As Long, ByRef hCompressor As LongPtr) As Long
    Private Declare PtrSafe Function Compress Lib "cabinet.dll" (ByVal hCompressor As LongPtr, ByVal UncompressedData As LongPtr, ByVal UncompressedDataSize As Long, ByVal CompressedBuffer As LongPtr, ByVal CompressedBufferSize As Long, ByRef CompressedBufferSize As Long) As Long
    Private Declare PtrSafe Function CloseCompressor Lib "cabinet.dll" (ByVal hCompressor As LongPtr) As Long
    
    Private Declare PtrSafe Function CreateDecompressor Lib "cabinet.dll" (ByVal CompressionAlgorithm As COMPRESS_ALGORITHM_ENUM, ByVal AllocationRoutines As Long, ByRef hDecompressor As LongPtr) As Long
    Private Declare PtrSafe Function Decompress Lib "cabinet.dll" (ByVal hCompressor As LongPtr, ByVal CompressedData As LongPtr, ByVal CompressedDataSize As Long, ByVal UncompressedBuffer As LongPtr, ByVal UncompressedBufferSize As Long, ByRef UncompressedDataSize As Long) As Long
    Private Declare PtrSafe Function CloseDecompressor Lib "cabinet.dll" (ByVal hDecompressor As LongPtr) As Long
#Else
    Private Enum LongPtr
    [_]
    End Enum
    Private Declare Function CreateCompressor Lib "cabinet.dll" (ByVal CompressionAlgorithm As COMPRESS_ALGORITHM_ENUM, ByVal AllocationRoutines As Long, ByRef hCompressor As LongPtr) As Long
    Private Declare Function Compress Lib "cabinet.dll" (ByVal hCompressor As LongPtr, ByVal UncompressedData As LongPtr, ByVal UncompressedDataSize As Long, ByVal CompressedBuffer As LongPtr, ByVal CompressedBufferSize As Long, ByRef CompressedBufferSize As Long) As Long
    Private Declare Function CloseCompressor Lib "cabinet.dll" (ByVal hCompressor As LongPtr) As Long
    
    Private Declare Function CreateDecompressor Lib "cabinet.dll" (ByVal CompressionAlgorithm As COMPRESS_ALGORITHM_ENUM, ByVal AllocationRoutines As Long, ByRef hDecompressor As LongPtr) As Long
    Private Declare Function Decompress Lib "cabinet.dll" (ByVal hCompressor As LongPtr, ByVal CompressedData As LongPtr, ByVal CompressedDataSize As Long, ByVal UncompressedBuffer As LongPtr, ByVal UncompressedBufferSize As Long, ByRef UncompressedDataSize As Long) As Long
    Private Declare Function CloseDecompressor Lib "cabinet.dll" (ByVal hDecompressor As LongPtr) As Long
#End If

Public Function CompressBytes(ByRef Source() As Byte, _
                              Optional ByVal Algorithm As COMPRESS_ALGORITHM_ENUM = LZMS) _
                              As Byte()

    If VBA.LenB(Source(0)) Then
        Dim Result As Long, hCompressor As LongPtr
        Result = CreateCompressor(Algorithm, 0, hCompressor)
        On Error GoTo ErrHandler
        If Result <> 0 Then
            Dim ByteLength As Long, Data As Long
            ByteLength = VBA.LenB(Source)
            ReDim Buffer(ByteLength - 1) As Byte
            If Compress(hCompressor, VarPtr(Source(0)), ByteLength, VarPtr(Buffer(0)), ByteLength, Data) Then
                If Data Then CompressBytes = VBA.LeftB(Buffer, Data)
            End If
ErrHandler:
            CloseCompressor hCompressor
        End If
        Erase Buffer
    End If
    
End Function

Public Function DecompressBytes(ByRef Source() As Byte, _
                                Optional ByVal Algorithm As COMPRESS_ALGORITHM_ENUM = LZMS) _
                                As Byte()
    
    If VBA.LenB(Source(0)) Then
        Dim Result As Long, hCompressor As LongPtr
        Result = CreateDecompressor(Algorithm, 0, hCompressor)
        On Error GoTo ErrHandler
        If Result <> 0 Then
            Dim ByteLength As Long, Data As Long
            ReDim Buffer(0) As Byte
            ByteLength = VBA.LenB(Source)
            If Decompress(hCompressor, VarPtr(Source(0)), ByteLength, VarPtr(Buffer(0)), 0, Data) = 0 Then
                ReDim Buffer(Data - 1)
                If Decompress(hCompressor, VarPtr(Source(0)), ByteLength, VarPtr(Buffer(0)), Data, Data) Then
                    If Data Then DecompressBytes = VBA.LeftB(Buffer, Data)
                End If
            End If
ErrHandler:
            CloseDecompressor hCompressor
        End If
        Erase Buffer
    End If
    
End Function

Public Function CompressString(ByVal Target As String, _
                               Optional ByVal Algorithm As COMPRESS_ALGORITHM_ENUM = LZMS) _
                               As String
    
    If VBA.LenB(Target) Then
        Dim TempBytes() As Byte
        TempBytes = Target
        CompressString = CompressBytes(TempBytes, Algorithm)
    End If
    
End Function

Public Function DecompressString(ByVal Target As String, _
                                 Optional ByVal Algorithm As COMPRESS_ALGORITHM_ENUM = LZMS) _
                                 As String
    
    If VBA.LenB(Target) Then
        Dim TempBytes() As Byte
        TempBytes = Target
        DecompressString = DecompressBytes(TempBytes, Algorithm)
    End If
    
End Function

Public Function CompressFile(ByVal TargetFilename As String, _
                             Optional ByVal CreateNewFile As Boolean = True, _
                             Optional ByVal Algorithm As COMPRESS_ALGORITHM_ENUM = LZMS) _
                             As String
    
    If VBA.LenB(Dir(TargetFilename)) Then
        Dim FileBytes() As Byte, CompressedData() As Byte
        FileBytes = ReadFile(TargetFilename)
        CompressedData = CompressBytes(FileBytes, Algorithm)
        If CreateNewFile Then TargetFilename = TargetFilename & "_COMPRESSED"
        Call WriteFile(TargetFilename, CompressedData, Not (CreateNewFile))
    End If
    
End Function

Public Function DecompressFile(ByVal TargetFilename As String, _
                               Optional ByVal CreateNewFile As Boolean = True, _
                               Optional ByVal Algorithm As COMPRESS_ALGORITHM_ENUM = LZMS) _
                               As String

    If VBA.LenB(Dir(TargetFilename)) Then
        Dim FileBytes() As Byte, CompressedData() As Byte
        FileBytes = ReadFile(TargetFilename)
        CompressedData = DecompressBytes(FileBytes, Algorithm)
        If CreateNewFile Then TargetFilename = TargetFilename & "_DECOMPRESSED"
        If WriteFile(TargetFilename, CompressedData, Not (CreateNewFile)) = 0 Then
            MsgBox "Compression Fail"
        End If
    End If
    
End Function

Public Function ReadFile(ByVal TargetFilename As String) As Byte()

    If VBA.LenB(Dir(TargetFilename)) Then
        Dim FileNum As Long, TempData() As Byte
        FileNum = FreeFile
        Open TargetFilename For Binary Access Read As #FileNum
        ReDim TempData(0 To LOF(FileNum) - 1&) As Byte
        Get #FileNum, , TempData
        Close #FileNum
        ReadFile = TempData
        Erase TempData
    End If
    
End Function

Public Function WriteFile(ByVal TargetFilename As String, _
                          ByRef FileData() As Byte, _
                          Optional ByVal DeleteExisting As Boolean = True) _
                          As Long
    
    If VBA.LenB(Dir(TargetFilename)) Then
        If DeleteExisting Then VBA.Kill TargetFilename Else Exit Function
    End If
    
    Dim FileNum As Long
    FileNum = FreeFile
    Open TargetFilename For Binary Access Write As #FileNum
    Put #FileNum, , FileData
    Close #FileNum
    
End Function


