Attribute VB_Name = "Module1"
Option Explicit
' This is used by GetUserName() to find the current user's
' name from the API

Declare Function Get_User_Name Lib "advapi32.dll" Alias _
"GetUserNameA" (ByVal lpBuffer As String, _
nSize As Long) As Long

'Private Declare PtrSafe Function GetComputerName Lib "kernel32" _
'        Alias "GetComputerNameA" ( _
'        ByVal lpBuffer As String, _
'        ByRef nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" _
        Alias "GetComputerNameA" ( _
        ByVal lpBuffer As String, _
        ByRef nSize As Long) As Long
Public glGetUserName As String
Public glMachineName As String
Public glComputerName As String

Function GetUserName() As String
  Dim lpBuff As String * 25
  Static glGetUserName As String
  Get_User_Name lpBuff, 25
  If glGetUserName = "" Then
     glGetUserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)
  End If
  GetUserName = glGetUserName
End Function

Public Property Get ComputerName() As String
  Dim stBuff As String * 255, lAPIResult As Long
  Dim lBuffLen As Long
  Static glComputerName As String
  
  lBuffLen = 255
  lAPIResult = GetComputerName(stBuff, lBuffLen)
   
  If lBuffLen > 0 And glComputerName = "" Then
     glComputerName = Left(stBuff, lBuffLen)
  End If
  ComputerName = glComputerName
End Property
Public Function GetMachineName() As String
  Static glGetMachineName As String
  If glGetMachineName = "" Then
     glGetMachineName = ComputerName
  End If
  GetMachineName = glGetMachineName
End Function

Public Function TomsDebugger() As Integer
    TomsDebugger = 4
' Level 0 means no debugging at all
End Function

Sub Mail_small_Text_Outlook(olTransactionCode As Integer)
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Working in Office 2000-2013
    Dim OutApp  As Object
    Dim OutMail As Object
    Dim StrBody As String

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    StrBody = "Hi there" & vbNewLine & vbNewLine & _
              "This is line 1" & vbNewLine & _
              "This is line 2" & vbNewLine & _
              "This is line 3" & vbNewLine & _
              "This is line 4"

    On Error Resume Next
    With OutMail
        .To = "CalendarCord@bbrd.org"
        .CC = ""
        .BCC = ""
        If olTransactionCode = 1 Then
           .Subject = "[RosieTheRobot]RV Storage New Transaction"
        ElseIf olTransactionCode = 2 Then
           .Subject = "[RosieTheRobot]RV Storage Cancel Transaction"
        ElseIf olTransactionCode = 3 Then
           .Subject = "[RosieTheRobot]RV Storage Change Transaction"
        Else
           .Subject = "[RosieTheRobot]InValid RV Storage Transaction Type sent"
        End If
        .Body = StrBody
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
'        .Send   'or use .Display
         .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Private Sub Workbook_Open()
   Application.Visible = False
   CRUD.Show
End Sub
Sub Button1_Click()
   Application.Visible = False
   CRUD.Show
End Sub

