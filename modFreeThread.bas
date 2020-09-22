Attribute VB_Name = "modFreeThread"
Public Const CTF_COINIT = &H8
Public Const CTF_INSIST = &H1
Public Const CTF_PROCESS_REF = &H4
Public Const CTF_THREAD_REF = &H2

Public Declare Function SHCreateThread Lib "shlwapi.dll" (ByVal pfnThreadProc As Long, pData As Any, ByVal dwFlags As Long, ByVal pfnCallback As Long) As Long

Sub Thisishow()
'SHCreateThread AddressOf myNewThread, ByVal 0&, CTF_INSIST, ByVal 0&


End Sub


