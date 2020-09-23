Attribute VB_Name = "mdlRegistry"
Option Explicit


Const sLocation As String = "mdlRegistry"


Public Sub Reg_DeleteValue(ByVal hKey As Long, ByVal sPath As String, ByVal sValueName As String)
On Error GoTo VB_Error
    
    Dim hCurKey As Long
    
    lErrors = RegOpenKeyEx(hKey, sPath, 0&, KEY_SET_VALUE, hCurKey): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_DeleteValue", "RegOpenKeyEx")
    lErrors = RegDeleteValue(hCurKey, sValueName): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_DeleteValue", "RegDeleteValue")
    lErrors = RegCloseKey(hCurKey): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_DeleteValue", "RegCloseKey")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Reg_EnumValue")
Resume Next
End Sub

Public Function Reg_EnumValue(ByVal hKey As Long, ByVal sPath As String, ByRef sValueName() As String, ByRef sData() As String, ByRef lDataType() As Long) As Long
On Error GoTo VB_Error

    Dim lIncrement As Long
    
    Dim hCurKey As Long
    Dim lValues As Long
    Dim lMaxValueNameLength As Long
    Dim lMaxValueLength As Long
    
    Dim lValueType As Long
    Dim lValue As Long
    Dim lDataLength As Long
    
    
    lErrors = RegOpenKeyEx(hKey, sPath, 0&, KEY_QUERY_VALUE, hCurKey): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_EnumValue", "RegOpenKeyEx")
    lErrors = RegQueryInfoKey(hCurKey, 0&, 0&, 0&, ByVal 0&, 0&, 0&, lValues, lMaxValueNameLength, lMaxValueLength, 0&, ByVal 0&): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_EnumValue", "RegQueryValueEx")
    
    For lIncrement = 0 To lValues - 1
        ReDim Preserve sValueName(lIncrement)
        sValueName(lIncrement) = String$(lMaxValueNameLength + 1, 0)
        lValue = Len(sValueName(lIncrement))
        
        ReDim Preserve sData(lIncrement)
        sData(lIncrement) = String$(lMaxValueLength, 0)
        ReDim Preserve lDataType(lIncrement)
        lDataLength = lMaxValueLength
        
        lErrors = RegEnumValue(hCurKey, lIncrement, sValueName(lIncrement), lValue, 0&, lDataType(lIncrement), ByVal sData(lIncrement), lDataLength): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_EnumValue", "RegEnumValue")
        
        If lValue > 0 Then
            sValueName(lIncrement) = Str_NullTerm_Fix(sValueName(lIncrement))
        End If
        If lDataLength > 0 Then
            sData(lIncrement) = Str_NullTerm_Fix(Left$(sData(lIncrement), lDataLength))
        End If
    Next lIncrement
    
    lErrors = RegCloseKey(hCurKey): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_EnumValue", "RegCloseKey")
    
    Reg_EnumValue = lValues
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Reg_EnumValue")
Resume Next
End Function

Public Function Reg_Read(ByVal hKey As Long, ByVal sPath As String, ByVal sValue As String, Optional ByRef bFail As Byte) As Variant
On Error GoTo VB_Error

    Dim hCurKey As Long
    Dim lDataBufferSize As Long
    Dim lValueType As Long
    
    bFail = 0
    
    lErrors = RegOpenKeyEx(hKey, sPath, 0&, KEY_QUERY_VALUE, hCurKey)
    If lErrors <> ERROR_SUCCESS Then
        Call Error_API(lErrors, sLocation & "\Reg_Read", "RegOpenKeyEx")
        bFail = 1
    End If
    
    lErrors = RegQueryValueEx(hCurKey, sValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
    If lErrors <> ERROR_SUCCESS Then
        Call Error_API(lErrors, sLocation & "\Reg_Read", "RegQueryValueEx")
        bFail = 2
    End If
    
    Select Case lValueType
        Case REG_BINARY
            Dim sBuffer As String
            sBuffer = String$(lDataBufferSize, 0)
            
            lErrors = RegQueryValueEx(hCurKey, sValue, 0&, lValueType, ByVal sBuffer, lDataBufferSize): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_Read", "RegQueryValueEx")
            
            Reg_Read = sBuffer
        Case REG_DWORD
            Dim lBuffer As Long
            lDataBufferSize = 4
            
            lErrors = RegQueryValueEx(hCurKey, sValue, 0&, lValueType, lBuffer, lDataBufferSize): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_Read", "RegQueryValueEx")
            Reg_Read = lBuffer
        Case REG_SZ
            sBuffer = String$(lDataBufferSize, 0)
            
            lErrors = RegQueryValueEx(hCurKey, sValue, 0&, lValueType, ByVal sBuffer, lDataBufferSize): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_Read", "RegQueryValueEx")
            
            Reg_Read = Str_NullTerm_Fix(sBuffer)
    End Select
    
    lErrors = RegCloseKey(hCurKey): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_Read", "RegCloseKey")
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Reg_Read")
Resume Next
End Function

Public Sub Reg_Write(ByVal hKey As Long, ByVal sPath As String, ByVal sValue As String, ByVal vData As Variant, ByVal lType As Long, Optional ByRef lDisposition As Long)
On Error GoTo VB_Error

    Dim lRetKey As Long
    
    lErrors = RegCreateKeyEx(hKey, sPath, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_SET_VALUE, ByVal 0&, lRetKey, lDisposition): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_Write", "RegCreateKeyEx")
    
    Select Case lType
        Case REG_BINARY
            lErrors = RegSetValueEx(lRetKey, sValue, 0&, lType, ByVal CStr(vData), Len(vData)): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_Write", "RegSetValueEx")
        Case REG_DWORD
            lErrors = RegSetValueEx(lRetKey, sValue, 0&, lType, uint32_int32(vData), 4): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_Write", "RegSetValueEx")
        Case REG_SZ
            lErrors = RegSetValueEx(lRetKey, sValue, 0&, lType, ByVal CStr(vData), Len(vData)): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_Write", "RegSetValueEx")
    End Select
    
    lErrors = RegCloseKey(lRetKey): If lErrors <> ERROR_SUCCESS Then Call Error_API(lErrors, sLocation & "\Reg_Write", "RegCloseKey")
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\Reg_Write")
Resume Next
End Sub
