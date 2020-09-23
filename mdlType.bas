Attribute VB_Name = "mdlType"
Option Explicit


Const sLocation As String = "mdlType"


Public Function int32x32_int64(ByVal lLo As Long, ByVal lHi As Long) As Double
On Error GoTo VB_Error
    
    Dim dLo As Double
    Dim dHi As Double
    
    If lLo < 0 Then
        dLo = (2 ^ 32) + lLo
    Else
        dLo = lLo
    End If
    If lHi < 0 Then
        dHi = (2 ^ 32) + lHi
    Else
        dHi = lHi
    End If
    
    int32x32_int64 = (dLo + (dHi * (2 ^ 32)))
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\int32x32_int64")
Resume Next
End Function

Public Function int32_uint32(ByVal lValue As Long) As Double
On Error GoTo VB_Error

    If lValue < 0 Then
        int32_uint32 = lValue + 4294967296#
    Else
        int32_uint32 = lValue
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\int32_uint32")
Resume Next
End Function

Public Function uint32_int32(ByVal lValue As Double) As Long
On Error GoTo VB_Error

    If lValue > 2147483647 Then
        uint32_int32 = lValue - 4294967296#
    Else
        uint32_int32 = lValue
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\uint32_int32")
Resume Next
End Function

Public Function int16_uint16(ByVal iValue As Integer) As Long
On Error GoTo VB_Error
    
    If iValue < 0 Then
        int16_uint16 = iValue + 65536
    Else
        int16_uint16 = iValue
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\int16_uint16")
Resume Next
End Function

Public Function uint16_int16(ByVal iValue As Long) As Integer
On Error GoTo VB_Error
    
    If iValue > 32767 Then
        uint16_int16 = iValue - 65536
    Else
        uint16_int16 = iValue
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\uint16_int16")
Resume Next
End Function

Public Function int32_int16(ByVal lValue As Long) As Integer
On Error GoTo VB_Error
    
    Dim iValue As Integer
    Call MoveMemory(iValue, lValue, 2)
    int32_int16 = iValue
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\int32_int16")
Resume Next
End Function

Public Function int16_int8(ByVal iValue As Long) As Byte
On Error GoTo VB_Error
    
    Dim bValue As Integer
    Call MoveMemory(bValue, iValue, 1)
    int16_int8 = bValue
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\int16_int8")
Resume Next
End Function

Public Function LOBYTE(ByVal iValue As Integer) As Byte
On Error GoTo VB_Error

    LOBYTE = int16_int8(iValue) And &HFFFF
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\LOBYTE")
Resume Next
End Function

Public Function HIBYTE(ByVal iValue As Integer) As Byte
On Error GoTo VB_Error
    
    HIBYTE = int16_int8(RShift_int16(iValue, 8))
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\HIBYTE")
Resume Next
End Function

Public Function LOWORD(ByVal lValue As Long) As Integer
On Error GoTo VB_Error

    LOWORD = int32_int16(lValue) And &HFFFF
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\LOWORD")
Resume Next
End Function

Public Function HIWORD(ByVal lValue As Long) As Integer
On Error GoTo VB_Error
    
    HIWORD = int32_int16(RShift_int32(lValue, 16))
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\HIWORD")
Resume Next
End Function

Public Function MAKELONG(ByVal wLow As Long, ByVal wHigh As Long) As Long
On Error GoTo VB_Error
    
    MAKELONG = (wLow And &HFFFF&) Or LShift_int32(uint32_int32(wHigh And &HFFFF&), 16)
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\MAKELONG")
Resume Next
End Function

Public Function MAKEWORD(ByVal bLow As Byte, ByVal bHigh As Byte) As Integer
On Error GoTo VB_Error
    
    MAKEWORD = (bLow And &HFF) Or LShift_int16(bHigh And &HFF, 8)
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\MAKEWORD")
Resume Next
End Function

Public Function RGB(ByVal byRed As Byte, ByVal byGreen As Byte, ByVal byBlue As Byte) As Long
On Error GoTo VB_Error
    
    RGB = byRed Or LShift_int8(byGreen, 8) Or LShift_int16(byBlue, 16)
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\RGB")
Resume Next
End Function
