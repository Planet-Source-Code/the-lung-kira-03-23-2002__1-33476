Attribute VB_Name = "winnetwk"
Option Explicit


Declare Function WNetEnumCachedPasswords Lib "mpr.dll" (ByVal s As String, ByVal i As Integer, ByVal b As Byte, ByVal CallBackProc As Long, ByVal l As Long) As Long


Public Type PASSWORD_CACHE_ENTRY
    cbEntry As Integer              'Size of this returned structure in bytes
    cbResource As Integer           'Size of the resource string, in bytes
    cbPassword As Integer           'Size of the password string, in bytes
    iEntry As Byte                  'Entry position In PWL file
    nType As Byte                   'Type of entry
    abResource(1 To 1024) As Byte   'Buffer to hold resource string, followed by password string
End Type
