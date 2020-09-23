Attribute VB_Name = "powrprof"
Option Explicit


Public Declare Function CallNtPowerInformation Lib "powrprof.dll" (ByVal InformationLevel As POWER_INFORMATION_LEVEL, ByRef lpInputBuffer As Any, ByVal nInputBufferSize As Long, ByRef lpOutputBuffer As Any, ByVal nOutputBufferSize As Long) As Long
