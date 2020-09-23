Attribute VB_Name = "mdlFile"
Option Explicit


Const sLocation As String = "mdlFile"


Public Function File_Exist(ByVal sFileName As String) As Boolean
On Error GoTo VB_Error

    Dim hHandle As Long
    
    hHandle = CreateFile(sFileName, 0&, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0&, 0&)
    If hHandle = INVALID_HANDLE_VALUE Then
        Call Error_API(Err.LastDllError, sLocation & "\File_Exist", "CreateFile")
        File_Exist = False
    Else
        If CloseHandle(hHandle) = False Then Call Error_API(Err.LastDllError, sLocation & "\File_Exist", "CloseHandle")
        File_Exist = True
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\File_Exist")
Resume Next
End Function

Public Function File_Pointer_Handle(ByVal hFile As Long) As Double
On Error GoTo VB_Error
    
    If GetFileType(hFile) = FILE_TYPE_DISK Then
        Dim lo As Long
        Dim hi As Long
        
        lo = SetFilePointer(hFile, 0, hi, FILE_CURRENT): If lo = INVALID_SET_FILE_POINTER Then Call Error_API(Err.LastDllError, sLocation & "\File_Pointer_Handle", "SetFilePointer")
        
        File_Pointer_Handle = int32x32_int64(lo, hi)
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\File_Pointer_Handle")
Resume Next
End Function

Public Function File_Read_Name(ByVal sFileName As String, ByVal lLength As Long, ByVal lStart As Long) As String
On Error GoTo VB_Error

    If sFileName = vbNullString Then Exit Function
    If lLength = 0 Then Exit Function
    Select Case File_Size_Name(sFileName)
        Case 0: Exit Function
        Case lStart: Exit Function
    End Select
    
    
    Dim hFile As Long
    Dim sBuffer As String
    Dim lRead As Long
    
    hFile = CreateFile(sFileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0&, 0&): If hFile = INVALID_HANDLE_VALUE Then Call Error_API(Err.LastDllError, sLocation & "\File_Size_Name", "CreateFile")
    
    If File_Pointer_Handle(hFile) <> lStart Then
        If SetFilePointer(hFile, lStart, 0&, FILE_BEGIN) = INVALID_SET_FILE_POINTER Then Call Error_API(Err.LastDllError, sLocation & "\File_Size_Name", "SetFilePointer")
    End If
    
    sBuffer = String$(lLength, 0)
    
    If ReadFile(hFile, ByVal sBuffer, lLength, lRead, ByVal 0&) = False Then
        Call Error_API(Err.LastDllError, sLocation & "\File_Size_Name", "ReadFile")
    Else
        If lRead = 0 Then Call Error_API(Err.LastDllError, sLocation & "\File_Size_Name", "ReadFile")
    End If
    
    If SetFilePointer(hFile, 0&, 0&, FILE_BEGIN) = INVALID_SET_FILE_POINTER Then Call Error_API(Err.LastDllError, sLocation & "\File_Size_Name", "SetFilePointer")
    If CloseHandle(hFile) = False Then Call Error_API(Err.LastDllError, sLocation & "\File_Size_Name", "CloseHandle")
    
    
    If lRead < lLength Then sBuffer = Left(sBuffer, lRead)
    
    File_Read_Name = sBuffer
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\File_Size_Name")
Resume Next
End Function

Public Function File_Size_Name(ByVal sFileName As String) As Double
On Error GoTo VB_Error

    Dim hi As Long
    Dim lo As Long
    
    Dim hHandle As Long
    
    hHandle = CreateFile(sFileName, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0&, 0&): If hHandle = INVALID_HANDLE_VALUE Then Call Error_API(Err.LastDllError, sLocation & "\File_Size_Name", "CreateFile")
    
    If Function_Exist("kernel32.dll", "GetFileSizeEx") = True Then
        Dim LARGE_INTEGER As LARGE_INTEGER
        If GetFileSizeEx(hHandle, LARGE_INTEGER) = False Then Call Error_API(Err.LastDllError, sLocation & "\File_Size_Name", "GetFileSizeEx")
        
        File_Size_Name = int32x32_int64(LARGE_INTEGER.LowPart, LARGE_INTEGER.HighPart)
    Else
        lo = GetFileSize(hHandle, hi): If lo = INVALID_FILE_SIZE Then Call Error_API(Err.LastDllError, sLocation & "\File_Size_Name", "GetFileSize")
        
        File_Size_Name = int32x32_int64(lo, hi)
    End If
    
    If CloseHandle(hHandle) = False Then Call Error_API(Err.LastDllError, sLocation & "\File_Size_Name", "CloseHandle")
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\File_Size_Name")
Resume Next
End Function

Public Sub File_Write_Name(ByVal sFileName As String, ByVal sData As String, ByVal lStart As Long, ByVal lFlags As Long)
On Error GoTo VB_Error
    
    If sFileName = vbNullString Then Exit Sub
    If sData = vbNullString Then Exit Sub
    
    
    Dim hFile As Long
    Dim lWrite As Long
    
    hFile = CreateFile(sFileName, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, lFlags, 0&, 0&): If hFile = INVALID_HANDLE_VALUE Then Call Error_API(Err.LastDllError, sLocation & "\File_Write_Name", "CreateFile")
    
    If File_Pointer_Handle(hFile) <> lStart Then
        If SetFilePointer(hFile, lStart, 0&, FILE_BEGIN) = INVALID_SET_FILE_POINTER Then Call Error_API(Err.LastDllError, sLocation & "\File_Write_Name", "SetFilePointer")
    End If
    
    If WriteFile(hFile, ByVal sData, Len(sData), lWrite, ByVal 0&) = False Then Call Error_API(Err.LastDllError, sLocation & "\File_Write_Name", "WriteFile")
    If SetFilePointer(hFile, 0&, 0&, FILE_BEGIN) = INVALID_SET_FILE_POINTER Then Call Error_API(Err.LastDllError, sLocation & "\File_Write_Name", "SetFilePointer")
    If CloseHandle(hFile) = False Then Call Error_API(Err.LastDllError, sLocation & "\File_Write_Name", "CloseHandle")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\File_Size_Name")
Resume Next
End Sub

