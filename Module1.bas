Attribute VB_Name = "Module1"
'Copyright 2002 Bobby Carter
Type Memorystatus
dwLength As Long

dwMemoryLoad As Long
dwTotalPhys As Long
dwAvailPhys As Long

dwTotalPageFile As Long
dwAvailPageFile As Long

dwTotalVirtual As Long
dwAvailVirtual As Long

End Type
Declare Sub GlobalMemoryStatus Lib "Kernel32" (lpBuffer As Memorystatus)



