VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SpecialFunc 
   Caption         =   "BAREFOOT BAY RECREATION DISTRICT                                       .. RV STORAGE  SUPPLEMENT SCREEN"
   ClientHeight    =   5616
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12630
   OleObjectBlob   =   "SpecialFunc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SpecialFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wkAcct As String
Dim wkSpace As String
Dim wkTranType As Integer
Dim wkval As Long
Dim wkDebugg As Boolean
Dim wkPreventInfiniteLoop As Integer
Dim wkInfiniteLoopLimit   As Integer

Private Sub vbMovingSpaces_Click()
' *****************************************************
' **  Begining of vbMovingSpaces_Click() Subroutine ***
' *****************************************************
Dim wkComment As String
If TomsDebugger > 5 Then
   wkDebugg = True
Else
   wkDebugg = False
End If

'  MsgBox " This Function is not YET implemented, Return to Screen "
If Trim(Me.vb2Space2.Value) = "" Then
    Me.vb2Space2.SetFocus
    MsgBox " Please Enter An Second Space Location "
    Exit Sub
End If

wkAcct = Me.vb2Acct
wkSpace = Me.vb2Space2
wkTranType = 11
wkComment = "Moved From Space " & Me.vb2Space1 & " to " & Me.vb2Space2 & " effective " & _
                CRUD.vbEffectiveDate & _
               ". Recorded on " & Date & " by " & GetUserName()

wkval = CRUD.oneSpaceAdd(wkAcct, wkSpace, wkComment, wkTranType)

If wkval > 0 Then
   Exit Sub
End If
    
wkSpace = Me.vb2Space1
wkTranType = 11

wkComment = "Space #" & Me.vb2Space1 & " is cancelled effective " & CRUD.vbEffectiveDate & _
               "Moved to Space " & Me.vb2Space2 & _
               ". Recorded on " & Date & " by " & GetUserName() & _
               " from " & GetMachineName() & " on " & GetFirstNonLocalIPAddress()

wkval = CRUD.oneSpaceDelete(wkAcct, wkSpace, wkComment, wkTranType)

If wkval > 0 Then
   Exit Sub
End If

If wkDebugg = True Then
   MsgBox "Space Moved .", vbOKOnly + vbInformation, wkAcct & _
          " Moved to New Location " & Me.vb2Space2
End If

Me.vb2InfoBar = "Location " & Me.vb2Space1 & " moved to " & _
            Me.vb2Space2 & " for Account:" & wkAcct & _
            " has been successfully completed."

'clear the data for the Next Screen
' ClearForm

Me.vb2Space2.SetFocus
' *****************************************************
' **  Ending of vbMovingSpaces_Click() Subroutine   ***
' *****************************************************
    
End Sub

Private Sub vbReplacementCard_Click()
'  MsgBox " This Function is not YET implemented, Return to Screen "
' *****************************************************
' **  Begining of butUpdate_Click() Subroutine      ***
' *****************************************************
Dim iRow As Long
Dim iRowStart As Long
Dim ws As Worksheet
Dim ws2 As Worksheet
Dim wkAcct As String
Dim wkSpace As String
Dim wkMima As String
Dim Rng As Range
Dim wkAccess1Change As Boolean
Dim wkAccess2Change As Boolean
Dim wkHardSearch As Boolean
Dim wkDebugg As Boolean

If TomsDebugger > 5 Then
   wkDebugg = True
Else
   wkDebugg = False
End If

wkHardSearch = True
wkAccess2Change = False
iRowStart = 2

If Me.vb2ReasonList1 <> "" Then
   wkAccess1Change = True
End If
If Me.vb2ReasonList2 <> "" Then
   wkAccess2Change = True
End If

Set ws = Worksheets("Space #")

'check for a Account number
If Trim(CRUD.vbAcct.Value) = "" Then
    CRUD.vbAcct.SetFocus
    MsgBox " Please Enter An Account "
    Exit Sub
End If

'check for an Existing Account Number

On Error Resume Next
Set Rng = ws.Columns("A:A").Find(What:=CRUD.vbAcct, After:=ws.Range("A" & iRowStart - 1), _
       SearchOrder:=xlByRows, _
       SearchDirection:=xlNext, LookIn:=xlFormulas)
       
If Err.Number <> 0 Then
   MsgBox "Account Not found! Try Again."
   Exit Sub
End If

iRow = Rng.Row

If CRUD.vbAcct.Value <> ws.Cells(iRow, 1).Value Then
   MsgBox "Account Not found! Try Again."
   Exit Sub
End If

'  Account found , continue processing

If CRUD.FormEdits(CRUD.vbAcct, CRUD.vbSpace, CRUD) > 0 Then
   Exit Sub
End If

'Check to verify what is on the screen is where the correct row is pointed to !!

If CRUD.vbAcct.Value = ws.Cells(iRow, 1) And _
   CRUD.vbSpace = ws.Cells(iRow, 2) Then
' all is ok
ElseIf CRUD.vbRow <> iRow Then
   MsgBox " Are we on the correct row ?? , Try Again " & _
            CRUD.vbRow & " vs. " & iRow
   Exit Sub
Else
   MsgBox " There is a serious problem here, call Tech Support"
   CRUD.vbInfoBar = "Acct: " & CRUD.vbAcct & " in Space " & CRUD.vbSpace & _
         " vs. Acct: " & ws.Cells(iRow, 1) & " in Space " & _
         ws.Cells(iRow, 2)
   Exit Sub
End If

wkAcct = ws.Cells(iRow, 1)
wkSpace = ws.Cells(iRow, 2)
ws.Range("A" & iRow).Select

' Now Update the Historical tracking Database

Set ws2 = Worksheets("Historical tracking")
ws2.Activate
iRow = 2
iRowStart = 2

'find the correct sticker number
FetchCorrectStickerNumber:
wkPreventInfiniteLoop = wkPreventInfiniteLoop + 1

'find the correct (First instance ! ! ) sticker number
On Error Resume Next
Set Rng = ws2.Columns("E:E").Find(What:=CRUD.vbSticker, After:=ws2.Range("E" & iRowStart), _
       SearchOrder:=xlByRows, _
       SearchDirection:=xlNext, LookIn:=xlFormulas)
       
If Err.Number <> 0 Then
   MsgBox "Sticker Number found! Fatal Error"
   Exit Sub
End If

iRow = Rng.Row

If ws2.Cells(iRow, 1) <> CRUD.vbAcct Then
    MsgBox "Major Error , why is the account number different ? " & _
    ws2.Cells(iRow, 1) & " vs. " & CRUD.vbAcct & vbCrLf & _
    "This needs some investigation.  Check the Excel spreadsheet."
    Exit Sub
End If

If iRow < iRowStart Then
   MsgBox "The Correct Sticker Number " & CRUD.vbSticker & _
             " Not found! Fatal Error" & vbCrLf & _
             " Please investigate. "
   Exit Sub
End If

'find the correct Space Number AND sticker number
On Error Resume Next
Set Rng = ws2.Columns("B:B").Find(What:=Int(Right(Me.vb2Space1, 3)), _
       After:=ws2.Range("B" & iRow - 1), _
       SearchOrder:=xlByRows, _
       SearchDirection:=xlNext, LookIn:=xlFormulas)
       
If Err.Number <> 0 Then
   MsgBox "The Proper Space Number " & Me.vb2Space1 & _
             " Not found! Fatal Error" & vbCrLf & _
             " Please investigate. "
   Exit Sub
End If

iRow = Rng.Row
If ws2.Cells(iRow, 5).Value <> CRUD.vbSticker Then
   MsgBox "Can not locate the Sticker Number " & CRUD.vbSticker & _
             " Again. Fatal Error" & vbCrLf & _
             " Please investigate. "
   Exit Sub
End If

'find the correct Access Code Number Plus all other fields must match !
On Error Resume Next
Set Rng = ws2.Columns("F:F").Find(What:=Int(Me.vb2Access1), _
       After:=ws2.Range("F" & iRow - 1), _
       SearchOrder:=xlByRows, _
       SearchDirection:=xlNext, LookIn:=xlFormulas)
       
If Err.Number <> 0 Then
   MsgBox "Access Code " & Me.vb2Access1 & _
             " Not found! Fatal Error" & vbCrLf & _
             " Please investigate. "
   Exit Sub
End If

iRow = Rng.Row
ws2.Range("A" & iRow).Select








If Int(ws2.Cells(iRow, 6).Value) <> Int(Me.vb2Access1.Value) Then
    MsgBox "Major Error , why is the account number different ? " & _
    ws2.Cells(iRow, 1) & " vs. " & CRUD.vbAcct & vbCrLf & _
    "This needs some investigation.  Check the Excel spreadsheet."
End If









If wkAccess1Change = True Or wkAccess2Change = True Then
   ws2.Cells(iRow, 1).EntireRow.Resize(Int(1)).Insert
   ws2.Cells(iRow, 1).Value = CRUD.vbAcct
   ws2.Cells(iRow, 4).Value = CRUD.vbName
   ws2.Cells(iRow, 5) = CRUD.vbSticker
   ws2.Cells(iRow, 6) = Me.vb2Access2
   ws2.Cells(iRow, 11) = "New Secure - ID card issued" & _
            " # " & Me.vb2Access2 & " effective " & CRUD.vbEffectiveDate & _
            ". Recorded on " & Date & _
            " by " & GetUserName() & " from " & GetMachineName() & _
            " on " & GetFirstNonLocalIPAddress()
   ws2.Cells(iRow + 1, 11) = ws2.Cells(iRow + 1, 11) & vbCrLf & Me.vb2ReasonList1 & _
            " card # " & Me.vb2Access1 & " effective " & _
            CRUD.vbEffectiveDate & ". Recorded on " & Date & _
            " by " & GetUserName() & " from " & GetMachineName() & _
            " on " & GetFirstNonLocalIPAddress()
   
   ws2.Cells(iRow + 1, 12).Value = Date
   ws2.Cells(iRow + 1, 13).Value = Time
   ws2.Cells(iRow + 1, 14) = GetUserName()
   ws2.Cells(iRow + 1, 15) = GetMachineName()
   ws2.Cells(iRow, 16) = GetFirstNonLocalIPAddress()
   ws2.Cells(iRow + 1, 14) = GetUserName()
   ws2.Cells(iRow + 1, 15) = GetMachineName()
   ws2.Cells(iRow + 1, 16) = GetFirstNonLocalIPAddress()
'  We'll make this range of cells have a Yellow Background
   ws2.Range("A" & iRow + 1).Select
   ws2.Range("A" & iRow + 1 & ":K" & iRow + 1).Select
       With Selection.Interior
           .Pattern = xlSolid
           .PatternColorIndex = xlAutomatic
           .Color = 65535
           .TintAndShade = 0
           .PatternTintAndShade = 0
       End With
End If


ws2.Cells(iRow, 2).Value = Right(Me.vb2Space1.Value, 3)
If Left(Me.vb2Space1, 1) = "M" Then
   ws2.Cells(iRow, 3) = "Micco"
ElseIf Left(Me.vb2Space1, 1) = "W" Then
   ws2.Cells(iRow, 3) = "West"
Else
   ws2.Cells(iRow, 3) = "*** ERROR ***"
End If

If ws2.Cells(iRow, 11) > "" Then
   ws2.Cells(iRow, 11) = ws2.Cells(iRow, 11) & vbCrLf & Me.vb2ReasonList1 & Date & _
                   " by " & GetUserName() & " from " & GetMachineName() & _
                    " on " & GetFirstNonLocalIPAddress()
   With ws2.Rows(iRow)
        .RowHeight = .RowHeight * 2
   End With
Else
   ws2.Cells(iRow, 11) = Me.vb2ReasonList1 & Date & _
                   " by " & GetUserName() & " from " & GetMachineName() & _
                    " on " & GetFirstNonLocalIPAddress()
End If
ws2.Cells(iRow, 12).Value = Date
ws2.Cells(iRow, 13).Value = Time


If ws2.Cells(iRow, 1) = ws2.Cells(iRow + 1, 1) And _
   ws2.Cells(iRow, 2) = ws2.Cells(iRow + 1, 2) And _
   ws2.Cells(iRow, 3) = ws2.Cells(iRow + 1, 3) And _
   ws2.Cells(iRow, 5) = ws2.Cells(iRow + 1, 5) And _
   ws2.Cells(iRow, 6) = ws2.Cells(iRow + 1, 6) And _
   ws2.Cells(iRow, 7) = ws2.Cells(iRow + 1, 7) Then
' We have a duplicate column underneath, Lets update those items Also !
    ws2.Cells(iRow + 1, 2).Value = Right(Me.vb2Space1.Value, 3)
    If Left(Me.vb2Space1, 1) = "M" Then
       ws2.Cells(iRow + 1, 3) = "Micco"
    ElseIf Left(Me.vb2Space1, 1) = "W" Then
       ws2.Cells(iRow + 1, 3) = "West"
    Else
       ws2.Cells(iRow + 1, 3) = "*** ERROR ***"
    End If
'    ws2.Cells(iRow + 1, 4) = Me.vbName
    ws2.Cells(iRow + 1, 6) = Me.vb2Access1.Value
    ws2.Cells(iRow + 1, 7) = Me.vb2Access2.Value

    ws2.Cells(iRow + 1, 11) = ws2.Cells(iRow + 1, 11) & Me.vb2ReasonList1 & Date & _
                   " by " & GetUserName() & " from " & GetMachineName() & _
                    " on " & GetFirstNonLocalIPAddress() & vbCrLf
    ws2.Cells(iRow + 1, 12).Value = Date
    ws2.Cells(iRow + 1, 13).Value = Time
End If

If wkDebugg = True Then
   MsgBox "Space " & Me.vb2Space1 & " Updated Successfully.", _
          vbOKOnly + vbInformation, "Acct# " & wkAcct & " Updated Successfully."
End If

Me.vb2InfoBar = "Account " & CRUD.vbAcct & " in Space " & Me.vb2Space1 & _
 " Secure Id card has been successfully updated. "

'clear the data for the Next Screen

CRUD.vbAcct.SetFocus

' *****************************************************
' **  End of vbReplacementCard_Click Subroutine     ***
' *****************************************************
End Sub
Private Sub ClearForm3()
    Me.vb2Access1 = ""
    Me.vb2Access2 = ""
    Me.vb2Access2 = ""
    Me.vb2Access3 = ""
    Me.vb2InfoBar = ""
    Me.vb2ReasonList1 = ""
    Me.vb2ReasonList2 = ""
    Me.vb2Space1 = ""
    Me.vb2Space2 = ""
    Me.vb2iRow = ""
    Me.vb2Acct = ""
    
    wkInfiniteLoopLimit = 2000
End Sub
Private Sub butExit_Click()
    ClearForm3
    SpecialFunc.Hide
'    SpecialFunc.Visible = False
    CRUD.Show
End Sub


