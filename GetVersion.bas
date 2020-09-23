Attribute VB_Name = "GetVersionMod"
Option Explicit

Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer


Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
'Used to determine what OS Version
Public Const WINNT As Integer = 2
Public Const WIN98 As Integer = 1

Public Function GetVersion() As Integer
  Dim udtOSInfo As OSVERSIONINFO
  Dim intRetVal As Integer
         
  'Initialize the type's buffer sizes
    With udtOSInfo
        .dwOSVersionInfoSize = 148
        .szCSDVersion = Space$(128)
    End With
    
  'Make an API Call to Retrieve the OSVersion info
    intRetVal = GetVersionExA(udtOSInfo)
  
  'Set the return value
    GetVersion = udtOSInfo.dwPlatformId
End Function

