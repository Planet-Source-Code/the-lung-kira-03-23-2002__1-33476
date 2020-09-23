Attribute VB_Name = "mdlString"
Option Explicit


Const sLocation As String = "mdlString"


Public Function ByteArray_String(ByRef abVar() As Byte) As String
On Error GoTo VB_Error

    ByteArray_String = String$(UBound(abVar()) + 1, 0)
    Call MoveMemory(ByVal ByteArray_String, abVar(0), UBound(abVar()) + 1)
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\ByteArray_String")
Resume Next
End Function

Public Function GUID_String(GUID As GUID) As String
On Error GoTo VB_Error

    Dim sReturn As String
    
    sReturn = "{" & Right$("00000000" & ltoa_(GUID.Data1, 16), 8) & "-"
    sReturn = sReturn & Right$("0000" & ltoa_(GUID.Data2, 8), 4) & "-"
    sReturn = sReturn & Right$("0000" & ltoa_(GUID.Data3, 8), 4) & "-"
    
    Dim lIncrement As Long
    For lIncrement = 0 To 7
        sReturn = sReturn & Right$("00" & ltoa_(GUID.Data4(lIncrement), 16), 2)
    Next lIncrement
    
    sReturn = sReturn & "}"
    GUID_String = sReturn
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\GUID_String")
Resume Next
End Function

Public Function Str_BckSlhTerm_Fix(ByVal sData As String) As String
On Error GoTo VB_Error

    If Len(sData) < 2 Then Exit Function
    
    If Right$(sData, 1) = "\" Then
        Str_BckSlhTerm_Fix = Left$(sData, Len(sData) - 1)
    Else
        Str_BckSlhTerm_Fix = sData
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Str_BckSlhTerm_Fix")
Resume Next
End Function

Public Function Str_CrLfTerm_Fix(ByVal sData As String) As String
On Error GoTo VB_Error

    Dim lPos As Long
    lPos = InStr(1, sData, vbCrLf)
    
    If lPos > 0 Then
        Str_CrLfTerm_Fix = Left$(sData, lPos - 2)
    Else
        Str_CrLfTerm_Fix = sData
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Str_CrLfTerm_Fix")
Resume Next
End Function

Public Function Str_NullTerm_Fix(ByVal sData As String) As String
On Error GoTo VB_Error

    Dim lPos As Long
    lPos = InStr(1, sData, vbNullChar)
    
    If lPos > 0 Then
        Str_NullTerm_Fix = Left$(sData, lPos - 1)
    Else
        Str_NullTerm_Fix = sData
    End If
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Str_NullTerm_Fix")
Resume Next
End Function

Public Function Unicode_Ascii(ByVal sUnicode As String, ByVal lFlags As Long) As String
On Error GoTo VB_Error

    Dim sAscii As String
    
    sAscii = String$(Len(sUnicode), 0)
    lErrors = WideCharToMultiByte(CP_ACP, lFlags, sUnicode, Len(sUnicode), sAscii, Len(sAscii), vbNullString, False): If lErrors = 0 Then Call Error_API(Err.LastDllError, sLocation & "\Unicode_Ascii", "WideCharToMultiByte")
    
    Unicode_Ascii = Left$(sAscii, lErrors)
    
Exit Function
VB_Error:
Call Error_VB(Err, sLocation & "\Unicode_Ascii")
Resume Next
End Function

