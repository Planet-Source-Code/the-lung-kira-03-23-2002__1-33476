Attribute VB_Name = "mdlVB"
Option Explicit


Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef lpObject() As Any) As Long
Public Declare Function VarPtr Lib "msvbvm60.dll" (ByRef lpObject As Any) As Long


Const sLocation As String = "mdlVB"


Public Function MinMax(ByVal dValue As Double, ByVal dMin As Double, ByVal dMax As Double) As Double
On Error GoTo VB_Error
    
    If dValue > dMax Then
        MinMax = dMax
    Else
        If dValue < dMin Then
            MinMax = dMin
        Else
            MinMax = dValue
        End If
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Adjust_Token_Priv")
Resume Next
End Function
