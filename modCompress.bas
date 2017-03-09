Attribute VB_Name = "modCompress"
Option Explicit

Public Enum ZCompressLevels
    Z_NO_COMPRESSION = 0
    Z_BEST_SPEED = 1
    Z_BEST_COMPRESSION = 9
    Z_DEFAULT_COMPRESSION = (-1)
End Enum

Private Declare Function Compress Lib "zlib.dll" Alias "compress2" (ByRef DestinationArray As Byte, ByRef destLen As Long, ByRef SourceArray As Byte, ByVal SourceLen As Long, ByVal CompressionLevel As Long) As Long
Private Declare Function ZCompress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function ZUncompress Lib "zlib.dll" Alias "uncompress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function Uncompress Lib "zlib.dll" Alias "uncompress" (ByRef DestinationArray As Byte, ByRef destLen As Long, ByRef SourceArray As Byte, ByVal SourceLen As Long) As Long
Public Function ZCompressByteArray(ByRef ArrayToCompress() As Byte, _
                                   ByRef Return_Array() As Byte, _
                                   Optional ByVal CompressionLevel As ZCompressLevels = Z_BEST_COMPRESSION, _
                                   Optional ByRef Return_ErrorCode As Long, _
                                   Optional ByVal TagOriginalSize As Boolean = True) As Boolean
    On Error GoTo ErrorTrap

    Dim OrigSize As String
    Dim ArrayLenS As Long
    Dim ArrayLenD As Long
    Dim CharCount As Long
    Dim MyCounter As Long

    Return_ErrorCode = 0

    ' Get the size of the source array
    ArrayLenS = UBound(ArrayToCompress) + 1
    If ArrayLenS = 0 Then
        ZCompressByteArray = True
        Exit Function
    End If

    ' Calculate the size of the desitnation buffer - (SourceLen * 0.001) + 12)
    ArrayLenD = ArrayLenS + ((ArrayLenS * 0.001) + 15)    ' Extra 3 bytes added on for some extra padding (avoids errors)

    ' Clear the return array
    Erase Return_Array
    ReDim Return_Array(ArrayLenD) As Byte

    ' Call the API to compress the byte array
    Return_ErrorCode = Compress(Return_Array(0), ArrayLenD, ArrayToCompress(0), ArrayLenS, CompressionLevel)
    If Return_ErrorCode <> 0 Then
        ZCompressByteArray = False
    Else
        ZCompressByteArray = True
    End If

    ' Redimention the resulting array to fit it's content
    If TagOriginalSize = False Then
        ReDim Preserve Return_Array(ArrayLenD - 1) As Byte

        ' Append the original size of the byte array to then end of the byte array.
        ' This is used in the "ZDecompressByteArray" function to automatically get the
        ' original size of the array (MAX = 2.1GB : 2,147,483,647 bytes).
    Else
        If ArrayLenS > 2147483647 Then
            ReDim Preserve Return_Array(ArrayLenD - 1) As Byte
            Exit Function
        End If

        ' Get the tag to append to the end of the byte array
        OrigSize = CStr(ArrayLenS)
        OrigSize = OrigSize & String(11 - Len(OrigSize), vbNullChar)
        OrigSize = String(5, vbNullChar) & OrigSize

        ' Redimention the size of the return array to it's compressed size, plus
        ' 16 bytes which contains the original size of the byte array.
        ' TAG Format = <5 x NULL> <ORIG SIZE> <(10 - Len(<ORIG SIZE>)) x NULL> <1 x NULL TERMINATOR>
        ReDim Preserve Return_Array(ArrayLenD + 16) As Byte

        ' Add the original size to the end
        For MyCounter = ArrayLenD To ArrayLenD + 16
            CharCount = CharCount + 1
            Return_Array(MyCounter) = Asc(Right(Left(OrigSize, CharCount), 1))
        Next
    End If

    Exit Function

ErrorTrap:

    If Err.number = 0 Then      ' No Error
        Resume Next
    ElseIf Err.number = 20 Then    ' Resume Without Error
        Resume Next
    Else                        ' Unknown Error
        MsgBox Err.Source & " caused the following error :" & Chr$(13) & Chr$(13) & "Error Number = " & CStr(Err.number) & Chr$(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
        Return_ErrorCode = Err.number
        Err.Clear
    End If

End Function
Public Function ZDecompressByteArray(ByRef ArrayToDecompress() As Byte, _
                                     ByRef Return_Array() As Byte, _
                                     Optional ByRef Return_ErrorCode As Long, _
                                     Optional ByVal OriginalSize As Long = -1) As Boolean
    On Error GoTo ErrorTrap

    Dim TestTag As String
    Dim OrigSize As String
    Dim ArrayLenS As Long
    Dim ArrayLenD As Long
    Dim MyCounter As Long

    Return_ErrorCode = 0

    ' Get the size of the source array
    ArrayLenS = UBound(ArrayToDecompress) + 1
    If ArrayLenS = 0 Then
        ZDecompressByteArray = True
        Exit Function
    End If

    ' Get the original array size from the value the user specified
    If OriginalSize <> -1 Then
        ArrayLenD = OriginalSize

        ' Get the original array size from the TAG value appended to the
        ' array by the "ZCompressByteArray" function
    Else
        For MyCounter = (ArrayLenS - 17) To ArrayLenS - 1
            TestTag = TestTag & Chr$(ArrayToDecompress(MyCounter))
        Next
        If Left(TestTag, 5) <> String(5, vbNullChar) Then
            Return_ErrorCode = -1
            Exit Function
        Else
            ' Get the original size from the tag value
            OrigSize = Right(TestTag, Len(TestTag) - 5)
            OrigSize = Left(OrigSize, InStr(OrigSize, vbNullChar) - 1)
            ArrayLenD = CLng(OrigSize)

            ' Redimention the array to cut off the tag
            ReDim Preserve ArrayToDecompress(ArrayLenS - 18) As Byte
            ArrayLenS = ArrayLenS - 17
        End If
    End If

    ' Clear the return array
    Erase Return_Array
    ReDim Return_Array(ArrayLenD) As Byte

    ' Decompress the byte array
    Return_ErrorCode = Uncompress(Return_Array(0), ArrayLenD, ArrayToDecompress(0), ArrayLenS)
    If Return_ErrorCode <> 0 Then
        ZDecompressByteArray = False
    Else
        ZDecompressByteArray = True
    End If

    Exit Function

ErrorTrap:

    If Err.number = 0 Then      ' No Error
        Resume Next
    ElseIf Err.number = 20 Then    ' Resume Without Error
        Resume Next
    Else                        ' Unknown Error
        MsgBox Err.Source & " caused the following error :" & Chr$(13) & Chr$(13) & "Error Number = " & CStr(Err.number) & Chr$(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
        Return_ErrorCode = Err.number
        Err.Clear
    End If

End Function
Public Function ZCompressFile(ByVal FileToCompress As String, _
                              ByVal OutputFile As String, _
                              Optional ByVal CompressionLevel As ZCompressLevels = Z_BEST_COMPRESSION, Optional ByRef Return_ErrorCode As Long, _
                              Optional ByVal OverwriteFile As Boolean = True) As Boolean
    On Error Resume Next

    Dim CompressedBuffer() As Byte
    Dim FileBuffer() As Byte
    Dim FileNum As Long

    Return_ErrorCode = 0

    ' Make sure the parameters passed are valid
    If FileToCompress = "" Or OutputFile = "" Then
        Exit Function
    ElseIf Dir(FileToCompress) = "" Then
        Exit Function
    ElseIf Dir(OutputFile) <> "" And OverwriteFile = False Then
        Exit Function
    End If

    ' Delete the file in case it already exists
    Kill OutputFile
    On Error GoTo ErrorTrap

    ' Create a buffer to recieve the file to compress
    ReDim FileBuffer(FileLen(FileToCompress) - 1)

    ' Read in the file
    FileNum = FreeFile
    Open FileToCompress For Binary Access Read As #FileNum
    Get #FileNum, , FileBuffer()
    Close #FileNum

    ' Compress the bytes that make up the file
    If ZCompressByteArray(FileBuffer, CompressedBuffer, CompressionLevel, Return_ErrorCode, True) = True Then
        ' Write out the compressed file
        FileNum = FreeFile
        Open OutputFile For Binary Access Write As #FileNum
        Put #FileNum, , CompressedBuffer()
        Close #FileNum
        ZCompressFile = True
    End If

CleanUp:

    ' Clean up memory that was used
    Erase CompressedBuffer
    Erase FileBuffer
    Close #FileNum

    Exit Function

ErrorTrap:

    If Err.number = 0 Then      ' No Error
        Resume Next
    ElseIf Err.number = 20 Then    ' Resume Without Error
        Resume Next
    Else                        ' Unknown Error
        MsgBox Err.Source & " caused the following error :" & Chr$(13) & Chr$(13) & "Error Number = " & CStr(Err.number) & Chr$(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
        Return_ErrorCode = Err.number
        Err.Clear
        Resume CleanUp
    End If

End Function
Public Function ZDecompressFile(ByVal FileToDecompress As String, _
                                ByVal OutputFile As String, _
                                Optional ByRef Return_ErrorCode As Long, _
                                Optional ByVal OverwriteFile As Boolean = True) As Boolean
    On Error Resume Next

    Dim DecompressedBuffer() As Byte
    Dim FileBuffer() As Byte
    Dim FileNum As Long

    Return_ErrorCode = 0

    ' Make sure the parameters passed are valid
    If FileToDecompress = "" Or OutputFile = "" Then
        Exit Function
    ElseIf Dir(FileToDecompress) = "" Then
        Exit Function
    ElseIf Dir(OutputFile) <> "" And OverwriteFile = False Then
        Exit Function
    End If

    ' Delete the file in case it already exists
    Kill OutputFile
    On Error GoTo ErrorTrap

    ' Create a buffer to recieve the file to compress
    ReDim FileBuffer(FileLen(FileToDecompress) - 1)

    ' Read in the file
    FileNum = FreeFile
    Open FileToDecompress For Binary Access Read As #FileNum
    Get #FileNum, , FileBuffer()
    Close #FileNum

    ' Compress the bytes that make up the file
    If ZDecompressByteArray(FileBuffer, DecompressedBuffer, Return_ErrorCode) = True Then
        ' Write out the compressed file
        FileNum = FreeFile
        Open OutputFile For Binary Access Write As #FileNum
        Put #FileNum, , DecompressedBuffer()
        Close #FileNum
        ZDecompressFile = True
    End If

CleanUp:

    ' Clean up memory that was used
    Erase DecompressedBuffer
    Erase FileBuffer
    Close #FileNum

    Exit Function

ErrorTrap:

    If Err.number = 0 Then      ' No Error
        Resume Next
    ElseIf Err.number = 20 Then    ' Resume Without Error
        Resume Next
    Else                        ' Unknown Error
        MsgBox Err.Source & " caused the following error :" & Chr$(13) & Chr$(13) & "Error Number = " & CStr(Err.number) & Chr$(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
        Return_ErrorCode = Err.number
        Err.Clear
        Resume CleanUp
    End If

End Function

Public Function CompressString(Data, Optional Key)
    Dim lKey As Long  'original size
    Dim sTmp As String  'string buffer
    Dim bData() As Byte  'data buffer
    Dim bRet() As Byte  'output buffer
    Dim lCSz As Long  'compressed size

    If TypeName(Data) = "Byte()" Then    'if given byte array data
        bData = Data  'copy to data buffer
    ElseIf TypeName(Data) = "String" Then    'if given string data
        If Len(Data) > 0 Then    'if there is data
            sTmp = Data    'copy to string buffer
            ReDim bData(Len(sTmp) - 1)    'allocate data buffer
            CopyMemory bData(0), ByVal sTmp, Len(sTmp)    'copy to data buffer
            sTmp = vbNullString    'deallocate string buffer
        End If
    End If
    If StrPtr(bData) <> 0 Then    'if data buffer contains data
        lKey = UBound(bData) + 1    'get data size
        lCSz = lKey + (lKey * 0.01) + 12    'estimate compressed size
        ReDim bRet(lCSz - 1)    'allocate output buffer
        Call ZCompress(bRet(0), lCSz, bData(0), lKey)    'compress data (lCSz returns actual size)
        ReDim Preserve bRet(lCSz - 1)    'resize output buffer to actual size
        Erase bData    'deallocate data buffer
        If IsMissing(Key) Then    'if Key variable not supplied
            ReDim bData(lCSz + 3)    'allocate data buffer
            CopyMemory bData(0), lKey, 4    'copy key to buffer
            CopyMemory bData(4), bRet(0), lCSz    'copy data to buffer
            Erase bRet    'deallocate output buffer
            bRet = bData    'copy to output buffer
            Erase bData    'deallocate data buffer
        Else    'Key variable is supplied
            Key = lKey    'set Key variable
        End If
        If TypeName(Data) = "Byte()" Then    'if given byte array data
            CompressString = bRet    'return output buffer
        ElseIf TypeName(Data) = "String" Then    'if given string data
            sTmp = Space(UBound(bRet) + 1)    'allocate string buffer
            CopyMemory ByVal sTmp, bRet(0), UBound(bRet) + 1    'copy to string buffer
            CompressString = sTmp    'return string buffer
            sTmp = vbNullString    'deallocate string buffer
        End If
        Erase bRet    'deallocate output buffer
    End If
End Function

Public Function UncompressString(Data, Optional ByVal Key)
    Dim lKey As Long  'original size
    Dim sTmp As String  'string buffer
    Dim bData() As Byte  'data buffer
    Dim bRet() As Byte  'output buffer
    Dim lCSz As Long  'compressed size

    If TypeName(Data) = "Byte()" Then    'if given byte array data
        bData = Data    'copy to data buffer
    ElseIf TypeName(Data) = "String" Then    'if given string data
        If Len(Data) > 0 Then    'if there is data
            sTmp = Data    'copy to string buffer
            ReDim bData(Len(sTmp) - 1)    'allocate data buffer
            CopyMemory bData(0), ByVal sTmp, Len(sTmp)    'copy to data buffer
            sTmp = vbNullString    'deallocate string buffer
        End If
    End If
    If StrPtr(bData) <> 0 Then    'if there is data
        If IsMissing(Key) Then    'if Key variable not supplied
            lCSz = UBound(bData) - 3    'get actual data size
            CopyMemory lKey, bData(0), 4    'copy key value to key
            ReDim bRet(lCSz - 1)    'allocate output buffer
            CopyMemory bRet(0), bData(4), lCSz    'copy data to output buffer
            Erase bData    'deallocate data buffer
            bData = bRet    'copy to data buffer
            Erase bRet    'deallocate output buffer
        Else    'Key variable is supplied
            lCSz = UBound(bData) + 1    'get data size
            lKey = Key    'get Key
        End If
        ReDim bRet(lKey - 1)    'allocate output buffer
        Call ZUncompress(bRet(0), lKey, bData(0), lCSz)    'decompress to output buffer
        Erase bData    'deallocate data buffer
        If TypeName(Data) = "Byte()" Then    'if given byte array data
            UncompressString = bRet    'return output buffer
        ElseIf TypeName(Data) = "String" Then    'if given string data
            sTmp = Space(lKey)    'allocate string buffer
            CopyMemory ByVal sTmp, bRet(0), lKey    'copy to string buffer
            UncompressString = sTmp    'return string buffer
            sTmp = vbNullString    'deallocate string buffer
        End If
        Erase bRet    'deallocate return buffer
    End If
End Function

