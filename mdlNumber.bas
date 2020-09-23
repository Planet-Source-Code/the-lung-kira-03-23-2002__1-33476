Attribute VB_Name = "mdlNumber"
Option Explicit


Const sLocation As String = "mdlNumber"

Public Type TIME_LENGTH
    lMilliseconds As Long
    lSeconds As Long
    lMinutes As Long
    lHours As Long
    lDays As Long
End Type


Public Function Percentage(ByVal dValue As Double, ByVal dTotal As Double, ByVal lRound As Long) As Double
On Error GoTo VB_Error

    If dValue <> 0 Then
        If dTotal <> 0 Then
            Percentage = Round((dValue / dTotal) * 100, lRound)
        End If
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Percentage")
Resume Next
End Function

Public Function LShift_int32(ByVal lValue As Long, ByVal lPlaces As Long) As Long
On Error GoTo VB_Error
    
    LShift_int32 = uint32_int32(lValue * (2 ^ lPlaces))
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\LShift_int32")
Resume Next
End Function

Public Function RShift_int32(ByVal lValue As Long, ByVal lPlaces As Long) As Long
On Error GoTo VB_Error
    
    Dim dValue As Double
    dValue = int32_uint32(lValue)
    RShift_int32 = uint32_int32(Int(dValue / (2 ^ lPlaces)))
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\RShift_int32")
Resume Next
End Function

Public Function LShift_int16(ByVal iValue As Integer, ByVal lPlaces As Long) As Integer
On Error GoTo VB_Error
    
    LShift_int16 = uint16_int16(iValue * (2 ^ lPlaces))
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\LShift_int16")
Resume Next
End Function

Public Function RShift_int16(ByVal iValue As Integer, ByVal lPlaces As Long) As Integer
On Error GoTo VB_Error
    
    Dim lValue As Double
    lValue = int16_uint16(iValue)
    RShift_int16 = uint16_int16(lValue \ (2 ^ lPlaces))
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\RShift_int16")
Resume Next
End Function

Public Function LShift_int8(ByVal bValue As Byte, ByVal lPlaces As Long) As Byte
On Error GoTo VB_Error
    
    LShift_int8 = bValue * (2 ^ lPlaces)
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\LShift_int8")
Resume Next
End Function

Public Function RShift_int8(ByVal bValue As Byte, ByVal lPlaces As Long) As Byte
On Error GoTo VB_Error
    
    RShift_int8 = bValue \ (2 ^ lPlaces)
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\RShift_int8")
Resume Next
End Function

Public Sub Number_TimeLength(ByVal dNumber As Double, ByRef TIME_LENGTH As TIME_LENGTH)
On Error GoTo VB_Error
    
    With TIME_LENGTH
        dNumber = Fix(dNumber / 10000)
        .lMilliseconds = (dNumber Mod 1000)
        
        dNumber = Fix(dNumber / 1000)
        .lSeconds = (dNumber Mod 60)
        
        dNumber = Fix(dNumber / 60)
        .lMinutes = (dNumber Mod 60)
        
        dNumber = Fix(dNumber / 60)
        .lHours = (dNumber Mod 24)
        
        dNumber = Fix(dNumber / 24)
        .lDays = dNumber
    End With
    
Exit Sub
VB_Error:
Call Error_VB(Err, sLocation & "\RShift_int8")
Resume Next
End Sub
