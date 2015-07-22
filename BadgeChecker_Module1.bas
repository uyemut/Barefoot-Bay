Attribute VB_Name = "Module1"
Option Explicit
' This is used by GetUserName() to find the current user's
' name from the API

Declare Function Get_User_Name Lib "advapi32.dll" Alias _
"GetUserNameA" (ByVal lpBuffer As String, _
nSize As Long) As Long

Private Declare Function GetComputerName Lib "kernel32" _
        Alias "GetComputerNameA" ( _
        ByVal lpBuffer As String, _
        ByRef nSize As Long) As Long

Function GetUserName() As String
  Dim lpBuff As String * 25
  Get_User_Name lpBuff, 25
  GetUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
End Function

Public Property Get ComputerName() As String

  Dim stBuff As String * 255, lAPIResult As Long
  Dim lBuffLen As Long
  
  lBuffLen = 255
  
  lAPIResult = GetComputerName(stBuff, lBuffLen)
  
  If lBuffLen > 0 Then ComputerName = Left(stBuff, lBuffLen)

End Property
Public Function GetMachineName() As String
  GetMachineName = ComputerName
End Function
