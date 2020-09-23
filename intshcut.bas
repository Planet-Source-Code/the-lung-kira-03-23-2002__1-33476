Attribute VB_Name = "intshcut"
Option Explicit


Public Declare Function InetIsOffline Lib "url.dll" (ByVal dwFlags As Long) As Boolean
