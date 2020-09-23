Attribute VB_Name = "guiddef"
Option Explicit


Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
