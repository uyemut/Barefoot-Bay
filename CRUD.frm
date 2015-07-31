VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CRUD 
   Caption         =   "BAREFOOT BAY RECREATION DISTRICT  -  ALL Storage Changes and Updates          .    .   .  version 1.00 (July 17th, 2015)"
   ClientHeight    =   10530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11700
   OleObjectBlob   =   "CRUD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CRUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim wkPreventInfiniteLoop As Integer
Dim wkInfiniteLoopLimit   As Integer

Private Sub butAdd_Click()
' *****************************************************
' **  Start of butAdd_Click() Subroutine            ***
' *****************************************************
Dim wkAcct As String
Dim wkSpace As String
Dim wkItemStr As String
Dim wkSpaceStr As String
Dim wkTranType As Integer
Dim wkval As Long
Dim wkDebugg As Boolean
If TomsDebugger > 5 Then
   wkDebugg = True
Else
   wkDebugg = False
End If

wkAcct = Me.vbAcct
wkSpace = Me.vbSpace
wkTranType = 1
wkItemStr = "Item "
wkSpaceStr = wkSpace

wkval = oneSpaceAdd(wkAcct, wkSpace, "", wkTranType)

If wkval > 0 Then
   Exit Sub
End If

If Me.vbSpace2 > "" Then
    wkSpace = Me.vbSpace2
    wkTranType = 2
    wkItemStr = "Items "
    wkSpaceStr = wkSpaceStr & " and " & wkSpace

    wkval = oneSpaceAdd(wkAcct, wkSpace, "", wkTranType)

    If wkval > 0 Then
        Exit Sub
    End If
End If

If wkDebugg = True Then
   MsgBox wkItemStr & "Added ", vbOKOnly + vbInformation, _
          wkSpaceStr & " Successfully Added."
End If
Me.vbInfoBar = "Account " & Me.vbAcct & " in " & wkItemStr & wkSpaceStr & _
               " Successfully Added."

'clear the data for the Next Screen
UserForm_Initialize
ClearForm

Me.vbAcct.SetFocus
' *****************************************************
' **   End of butAdd_Click() Subroutine             ***
' *****************************************************
End Sub

Private Sub butDelete_Click()
' *****************************************************
' **  Begining of butDelete_Click() Subroutine      ***
' *****************************************************
Dim wkAcct As String
Dim wkSpace As String
Dim wkTranType As Integer
Dim wkval As Long
Dim wkDebugg As Boolean
Dim wkHardSearch As Boolean

wkHardSearch = True
If TomsDebugger > 5 Then
   wkDebugg = True
Else
   wkDebugg = False
End If

wkAcct = Me.vbAcct
wkSpace = Me.vbSpace
wkTranType = 1

wkval = oneSpaceDelete(wkAcct, wkSpace, "", wkTranType)

If wkval > 0 Then
   Exit Sub
End If

If Me.vbSpace2 > "" Then
    wkSpace = Me.vbSpace2
    wkTranType = 2

    wkval = oneSpaceDelete(wkAcct, wkSpace, "", wkTranType)

    If wkval > 0 Then
       Exit Sub
    End If
End If

If wkDebugg = True Then
   MsgBox "Item Deleted ", vbOKOnly + vbInformation, Me.vbAcct.Value & " Item Deleted"
End If

Me.vbInfoBar = "Account " & Me.vbAcct & " in Space " & Me.vbSpace & _
 " has been successfully purged. "

'clear the data for the Next Screen
UserForm_Initialize
ClearForm

Me.vbAcct.SetFocus
' *****************************************************
' **       End of butDelete_Click() Subroutine      ***
' *****************************************************
End Sub

Public Function oneSpaceAdd(wkAcct As String, _
                             wkSpace As String, _
                             wkCommentOverRide As String, _
                             wkTranType As Integer)
' *****************************************************
' **     Start of oneSpaceAdd() Subroutine          ***
' *****************************************************
Dim iRow As Long
Dim iRowStart As Long
Dim iSpace As Long
Dim iEditErrorCount As Integer
Dim ws As Worksheet
Dim ws2 As Worksheet
Dim Rng As Range
Dim Rng2 As Range
Dim wkMima As String
Dim wkDebugg As Boolean
Dim wkAltSticker As String
Set ws = Worksheets("Space #")
iRow = 1
iRowStart = 1
iEditErrorCount = 0

If TomsDebugger > 5 Then
   wkDebugg = True
Else
   wkDebugg = False
End If

' Let's put some more edits
iEditErrorCount = iEditErrorCount + FormEdits(wkAcct, wkSpace, Me)

' Verify that Space does not exist
On Error Resume Next
iSpace = ws.Cells.Find(What:=wkSpace, _
       SearchOrder:=xlRows, _
       SearchDirection:=xlPrevious, LookIn:=xlValues).Row

If Err.Number = 0 And ws.Cells(iSpace, 1).Value > 0 Then
   MsgBox "Space " & wkSpace & " Is Occuppied! Choose Another Space."
   iEditErrorCount = iEditErrorCount + 1
ElseIf ws.Cells(iSpace, 1) = "" Then
' All is Ok, the reserved spot in the spreadsheet is there
Else
   MsgBox "Space " & wkSpace & " Is Missing! See Tech Support."
End If

'End of Screen Edits
If iEditErrorCount > 0 Then
    oneSpaceAdd = iEditErrorCount
    iEditErrorCount = 0
    Exit Function
End If

If oneSpaceAdd = FunctionalityEdits(1) > 0 Then
   ClearForm
   Exit Function
End If

'Unprotect the WorkSheet
wkMima = ws.Range("AG1").Value
ws.Unprotect Password:=wkMima

ws.Range("A" & iSpace & ":AA" & iSpace).Interior.ColorIndex = 0
ws.Activate
If wkTranType = 2 Then
    If ws.Cells(iSpace, 1) = ws.Cells(iSpace - 1, 1) Then
      Rng2 = ws.Range("A" & iSpace & ":A" & iSpace - 1).Select
    Else
      Rng2 = ws.Range("A" & iSpace & ":A" & iSpace + 1).Select
    End If
    Rng2.Font.ColorIndex = 27
End If
ws.Range(iSpace, 1).Select
ws.Range("A" & iSpace & ":AA" & iSpace).ClearComments

ws.Range("A" & iSpace & ":AA" & iSpace).NumberFormat = "@"
ws.Range("G" & iSpace).NumberFormat = "MM/DD/YYYY"
ws.Range("I" & iSpace).NumberFormat = "#"
ws.Range("K" & iSpace).NumberFormat = "MM/DD/YYYY"
ws.Range("T" & iSpace).NumberFormat = "#"
ws.Range("Y" & iSpace).NumberFormat = "MM/DD/YYYY"

'copy the data to the database
ws.Cells(iSpace, 1) = wkAcct
' ws.Cells(iSpace, 2) = wkSpace
ws.Cells(iSpace, 3) = Me.vbName.Value
ws.Cells(iSpace, 4) = Me.vbAddress1.Value
ws.Cells(iSpace, 5) = Me.vbAddress2.Value
ws.Cells(iSpace, 6) = "New"
ws.Cells(iSpace, 7).Value = Me.vbContractDate.Value
ws.Cells(iSpace, 8) = Me.vbSticker
'ws.Cells(iSpace, 9).Value = Me.vbAccessCode1.Value
ws.Cells(iSpace, 10) = Me.vbLicense1
ws.Cells(iSpace, 11) = Me.vbExpireDate.Value
ws.Cells(iSpace, 12) = Me.vbRegistrationOnfile
ws.Cells(iSpace, 13) = Me.vbVechileDescr
ws.Cells(iSpace, 14) = Me.vbTelephone1.Value
ws.Cells(iSpace, 15) = Me.vbTelephone2.Value
ws.Cells(iSpace, 16) = Me.vbNote
' ws.Cells(iSpace, 19) = "=IF($A3>"",VLOOKUP($H3,'Historical tracking'!$E$2:$K$444,7,FALSE),"")"
ws.Cells(iSpace, 21) = ""
ws.Cells(iSpace, 25) = Me.vbNoticeDate
ws.Cells(iSpace, 26) = Me.vbEmail

'Protect the WorkSheet Again
ws.Protect Password:=wkMima, DrawingObjects:=True, Contents:=True, Scenarios:=True _
    , AllowFormattingRows:=True, AllowInsertingRows:=True, _
    AllowDeletingRows:=True

' Now Copy identical Data to the 2nd Log Database

Set ws2 = Worksheets("Historical tracking")
iRow = 2

'find first empty row in Historical tracking sheet
 iRow = ws2.Cells.Find(What:="*", SearchOrder:=xlRows, _
  SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1
  
'find the correct sticker number
On Error Resume Next
Set Rng = ws2.Columns("E:E").Find(What:=Me.vbSticker, After:=ws2.Range("E" & iRowStart), _
       SearchOrder:=xlByRows, _
       SearchDirection:=xlNext, LookIn:=xlFormulas)
       
If Err.Number <> 0 Then ' If we can't find the pre-allocated sticker number - put it at the end of the list
    MsgBox "Sticker Number " & Me.vbSticker & _
           " Not found! Entry will be put at the bottom of the list "
ElseIf Me.vbSticker = ws2.Cells(Rng.Row, 5) And _
    wkAcct = ws2.Cells(Rng.Row, 1) Then
    If Int(Right(wkSpace, 3)) > ws2.Cells(Rng.Row, 2) Then
        iRow = Rng.Row + 1
    Else
        iRow = Rng.Row
    End If
    ws2.Cells(iRow, 1).EntireRow.Resize(Int(1)).Insert
    ws2.Range("A" & iRow & ":AA" & iRow).NumberFormat = "@"
    ws2.Range("B" & iRow).NumberFormat = "0"
    ws2.Range("F" & iRow).NumberFormat = "0"
    ws2.Range("G" & iRow).NumberFormat = "0"
    ws2.Range("L" & iRow).NumberFormat = "MM/DD/YYYY"
ElseIf ws2.Cells(Rng.Row, 1) <> "" Or ws2.Cells(Rng.Row, 3) <> "" Then
    MsgBox "Sticker #" & Me.vbSticker & _
            " has been used, will put entry at the bottome of the list."
Else
    iRow = Rng.Row
End If

ws2.Cells(iRow, 1) = wkAcct
ws2.Cells(iRow, 2) = Right(wkSpace, 3)

If Left(wkSpace, 1) = "M" Then
    ws2.Cells(iRow, 3) = "Micco"
Else
    ws2.Cells(iRow, 3) = "West"
End If

ws2.Cells(iRow, 4) = Me.vbName
ws2.Cells(iRow, 5) = Me.vbSticker
ws2.Cells(iRow, 6) = Me.vbAccessCode1.Value
ws2.Cells(iRow, 7) = Me.vbAccessCode2.Value
ws2.Cells(iRow, 8) = Me.vbLicense1
ws2.Cells(iRow, 9) = Me.vbExpireDate.Value
ws2.Cells(iRow, 10) = "New Contract"
ws2.Cells(iRow, 11) = wkCommentOverRide
ws2.Cells(iRow, 12).Value = Date
ws2.Cells(iRow, 13).Value = Time
ws2.Cells(iRow, 14) = ""
ws2.Cells(iRow, 15) = ""
ws2.Cells(iRow, 16) = ""
ws2.Cells(iRow, 17) = GetUserName()
ws2.Cells(iRow, 18) = GetMachineName()
ws2.Cells(iRow, 19) = GetFirstNonLocalIPAddress()

Me.vbRow = iRow

' *****************************************************
' **  End of butAdd_Click() Subroutine              ***
' *****************************************************
End Function

Public Function oneSpaceDelete(wkAcct As String, _
                                wkSpace As String, _
                                wkCommentOverRide As String, _
                                wkTranType As Integer)
' *****************************************************
' **  Begining of oneSpaceDelete() Subroutine      ***
' *****************************************************
Dim iRow As Long
Dim iRowStart As Long
Dim iRowHeight As Long
Dim ws As Worksheet
Dim ws2 As Worksheet
Dim wkMima As String
Dim Rng As Range
iRowHeight = 19.5

'Lets put this in for cases where we have Multiple Accounts that have multiple RV Lot Spaces
If Me.vbRow.Value = "" Then
   iRowStart = 1
Else
   iRowStart = Me.vbRow.Value
End If

Set ws = Worksheets("Space #")

'check for a Account number
If Trim(wkAcct) = "" Then
    Me.vbAcct.SetFocus
    MsgBox " Please Enter An Account "
    Exit Function
End If

'check for a Space Number

On Error Resume Next
Set Rng = ws.Columns("A:A").Find(What:=wkAcct, After:=ws.Range("A" & iRowStart - 1), _
       SearchOrder:=xlByRows, _
       SearchDirection:=xlNext, LookIn:=xlFormulas)
       
If Err.Number <> 0 Then
   MsgBox "Account Not found! Try Again."
   oneSpaceDelete = 1
   ClearForm
   Exit Function
End If

iRow = Rng.Row

If wkAcct <> ws.Cells(iRow, 1).Value Then
   MsgBox "Account Not found! Try Again."
   oneSpaceDelete = 2
    ClearForm
   Exit Function
End If

If oneSpaceDelete = FunctionalityEdits(2) > 0 Then
   ClearForm
   Exit Function
End If

'Check to verify what is on the screen is where the correct row is pointed to !!

If wkAcct = ws.Cells(iRow, 1) And _
   wkSpace = ws.Cells(iRow, 2) Then
' all is ok
ElseIf Me.vbRow <> iRow Then
   MsgBox " Are we on the correct row ?? , Try Again " & _
            Me.vbRow & " vs. " & iRow
   oneSpaceDelete = 2002
   Exit Function
Else
   MsgBox " There is a serious problem here, call Tech Support"
   Me.vbInfoBar = "Acct: " & wkAcct & " in Space " & wkSpace & _
         " vs. Acct: " & ws.Cells(iRow, 1) & " in Space " & _
         ws.Cells(iRow, 2)
   oneSpaceDelete = 2004
   Exit Function
End If

'Unprotect the SPACE WorkSheet
wkMima = ws.Range("AG1").Value
ws.Unprotect Password:=wkMima

ws.Activate
ws.Range("A" & iRow).Select
'clear the data from the this row
ws.Cells(iRow, 1) = ""
' ws.Cells(iRow, 2) = wkSpace
ws.Cells(iRow, 3) = ""
ws.Cells(iRow, 4) = ""
ws.Cells(iRow, 5) = ""
ws.Cells(iRow, 6) = ""
ws.Cells(iRow, 7) = ""
ws.Cells(iRow, 8) = ""
'ws.Cells(iRow, 9) = Me.vbAccessCode1.Value
ws.Cells(iRow, 10) = ""
ws.Cells(iRow, 11) = ""
ws.Cells(iRow, 12) = ""
ws.Cells(iRow, 13) = ""
ws.Cells(iRow, 14) = ""
ws.Cells(iRow, 15) = ""
ws.Cells(iRow, 16) = ""
' ws.Cells(iRow, 19).Value = "=IF($A2>"",VLOOKUP($H2,'Historical tracking'!$E$2:$K$837,2,FALSE),"")"
ws.Cells(iRow, 21) = "DELETED on " & Str(Date) & " at " & Str(Time)
ws.Cells(iRow, 26) = ""

' Let's reset the hight to the default, Clear any comments and remove any coloring.
Me.vbRow = iRow
ws.Rows(iRow).RowHeight = iRowHeight
ws.Range("A" & iRow & ":AA" & iRow).Interior.ColorIndex = 0
ws.Range("A" & iRow & ":AA" & iRow).ClearComments

'Protect the SPACE WorkSheet Again
ws.Protect Password:=wkMima, DrawingObjects:=True, Contents:=True, Scenarios:=True _
    , AllowFormattingRows:=True, AllowInsertingRows:=True, _
    AllowDeletingRows:=True

' Now Copy identical Data to the Historical Log Database

Set ws2 = Worksheets("Historical tracking")
ws2.Activate
iRow = 2
iRowStart = 2

'find the correct sticker number
FetchCorrectStickerNumber:
wkPreventInfiniteLoop = wkPreventInfiniteLoop + 1

On Error Resume Next
Set Rng = ws2.Columns("E:E").Find(What:=Me.vbSticker, _
       After:=ws2.Range("E" & iRowStart), _
       SearchOrder:=xlByRows, _
       SearchDirection:=xlNext, LookIn:=xlFormulas)
       
If Err.Number <> 0 Then
   MsgBox "Sticker Number " & Me.vbSticker & _
             " Not found! Fatal Error" & vbCrLf & _
             " Please investigate. "
   oneSpaceDelete = 2006
   ClearForm
   Exit Function
End If

iRow = Rng.Row

If iRow < iRowStart Then
   MsgBox "The Correct Sticker Number " & Me.vbSticker & _
             " Not found! Fatal Error" & vbCrLf & _
             " Please investigate. "
   oneSpaceDelete = 2008
   ClearForm
   Exit Function
End If

'find the correct Space Number AND sticker number
On Error Resume Next
Set Rng = ws2.Columns("B:B").Find(What:=Int(Right(wkSpace, 3)), _
       After:=ws2.Range("B" & iRow - 1), _
       SearchOrder:=xlByRows, _
       SearchDirection:=xlNext, LookIn:=xlFormulas)
       
If Err.Number <> 0 Then
   MsgBox "The Proper Space Number " & Me.vbSpace & _
             " Not found! Fatal Error" & vbCrLf & _
             " Please investigate. "
   oneSpaceDelete = 2012
   ClearForm
   Exit Function
End If

iRow = Rng.Row
If ws2.Cells(iRow, 5).Value <> Me.vbSticker Then
   MsgBox "Can not locate the Sticker Number " & Me.vbSticker & _
             " Again. Fatal Error" & vbCrLf & _
             " Please investigate. "
   oneSpaceDelete = 2014
   ClearForm
   Exit Function
End If

If ws2.Cells(iRow, 1).Value <> wkAcct Then
    If wkPreventInfiniteLoop < 5 Then
        MsgBox "Major Error , why is the account number different ? " & _
             ws2.Cells(iRow, 1).Value & " vs. " & wkAcct
    End If
    iRowStart = iRow
    If wkPreventInfiniteLoop < wkInfiniteLoopLimit Then GoTo FetchCorrectStickerNumber
End If

If CStr(Format(ws2.Cells(iRow, 2), "000")) <> Right(wkSpace, 3) Then
    If wkPreventInfiniteLoop < 5 Then
        MsgBox "Major Error , why is the space number different from the form ? " & _
            ws2.Cells(iRow, 2).Value & " vs. " & Right(wkSpace, 3)
    End If
    iRowStart = iRow
    If wkPreventInfiniteLoop < wkInfiniteLoopLimit Then GoTo FetchCorrectStickerNumber
End If

'  We'll make this range of cells have a Yellow Background
ws2.Range("A" & iRow).Select
ws2.Range("A" & iRow & ":K" & iRow).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
If ws2.Cells(iRow, 11) = "" Then
   If wkCommentOverRide > "" Then
      ws2.Cells(iRow, 11) = wkCommentOverRide
   Else
      ws2.Cells(iRow, 11) = "Space " & wkSpace & " cancelled effective " & Me.vbEffectiveDate & _
               ". Recorded on " & Date & " by " & GetUserName() & _
               " from " & GetMachineName() & " on " & GetFirstNonLocalIPAddress()
   End If
Else
    If wkCommentOverRide > "" Then
       ws2.Cells(iRow, 11) = ws2.Cells(iRow, 11) & vbCrLf & wkCommentOverRide
    Else
       ws2.Cells(iRow, 11) = ws2.Cells(iRow, 11) & vbCrLf & "Space " & wkSpace & _
                " cancelled effective " & Me.vbEffectiveDate & ". Recorded on " & Date & _
                " by " & GetUserName() & " from " & GetMachineName() & _
                " on " & GetFirstNonLocalIPAddress()
        With ws2.Rows(iRow)
            .RowHeight = .RowHeight * 2
        End With
    End If
End If

ws2.Cells(iRow, 12).Value = Date
ws2.Cells(iRow, 13).Value = Time
ws2.Cells(iRow, 14) = GetUserName()
ws2.Cells(iRow, 15) = GetMachineName()
ws2.Cells(iRow, 16) = GetFirstNonLocalIPAddress()

If ws2.Cells(iRow, 1) = ws2.Cells(iRow + 1, 1) And _
   ws2.Cells(iRow, 2) = ws2.Cells(iRow + 1, 2) And _
   ws2.Cells(iRow, 3) = ws2.Cells(iRow + 1, 3) And _
   ws2.Cells(iRow, 5) = ws2.Cells(iRow + 1, 5) And _
   ws2.Cells(iRow, 6) = ws2.Cells(iRow + 1, 6) And _
   ws2.Cells(iRow, 7) = ws2.Cells(iRow + 1, 7) Then
' We have a duplicate column underneath, Lets highlight it as deleted Also !
'    and put time, location and date stamps on it
   ws2.Range("A" & iRow + 1 & ":K" & iRow + 1).Select
       With Selection.Interior
           .Pattern = xlSolid
           .PatternColorIndex = xlAutomatic
           .Color = 65535
           .TintAndShade = 0
           .PatternTintAndShade = 0
    End With
    If ws2.Cells(iRow + 1, 11) <> "" Then
       ws2.Cells(iRow + 1, 11) = ws2.Cells(iRow + 1, 1) & vbCrLf
    End If
    ws2.Cells(iRow + 1, 11) = ws2.Cells(iRow, 11) & "Cancelled effective " & _
                    Me.vbEffectiveDate & ". Recorded on " & Date & _
                   " by " & GetUserName() & " from " & GetMachineName() & _
                    " on " & GetFirstNonLocalIPAddress()
    ws2.Cells(iRow + 1, 12).Value = Date
    ws2.Cells(iRow + 1, 13).Value = Time
    ws2.Cells(iRow + 1, 14) = GetUserName()
    ws2.Cells(iRow + 1, 15) = GetMachineName()
    ws2.Cells(iRow + 1, 16) = GetFirstNonLocalIPAddress()
End If

' *****************************************************
' **       End of oneSpaceDelete() Subroutine      ***
' *****************************************************
End Function

Private Sub butFind_Click()
' *****************************************************
' **  Begining of butFind_Click() Subroutine        ***
' *****************************************************
Dim iRow As Long
Dim iRowStart As Long
Dim ws As Worksheet
Dim Rng As Range
Dim wkHardSearch As Boolean
Dim wkDebugg As Boolean

wkHardSearch = True
' wkDebugg = True
If TomsDebugger > 8 Then
   wkDebugg = True
Else
   wkDebugg = False
End If
'Lets put this in for cases where we have Multiple Accounts that have multiple RV Lot Spaces
If Me.vbRow.Value = "" Then
   iRowStart = 1
Else
   iRowStart = Me.vbRow.Value
End If

Set ws = Worksheets("Space #")
'iRow = 0

'check for a Account number
If Trim(Me.vbAcct.Value) = "" Then
    Me.vbAcct.SetFocus
    MsgBox " Please Enter An Account "
    Exit Sub
End If

'check for a Space Number

On Error Resume Next
Set Rng = ws.Columns("A:A").Find(What:=Me.vbAcct, After:=ws.Range("A" & iRowStart), _
       SearchOrder:=xlByRows, _
       SearchDirection:=xlNext, LookIn:=xlFormulas)
       
If Err.Number <> 0 Then
   MsgBox "Account Not found! Try Again."
   ClearForm
   Exit Sub
End If

iRow = Rng.Row

If Me.vbAcct.Value <> ws.Cells(iRow, 1).Value Then
   MsgBox "Account Not found! Try Again."
   ClearForm
   Exit Sub
End If

If FunctionalityEdits(0) > 0 Then
   ClearForm
   Exit Sub
End If

'copy the data to the Form

Me.vbAcct.Value = ws.Cells(iRow, 1).Value
Me.vbAcctPrev = ws.Cells(iRow, 1).Value
Me.vbSpace.Value = ws.Cells(iRow, 2).Value
Me.vbSpacePrev.Value = ws.Cells(iRow, 2).Value
' ----- ******************************* -----
Me.vbName.Value = ws.Cells(iRow, 3).Value
Me.vbAddress1.Value = ws.Cells(iRow, 4).Value
Me.vbAddress2.Value = ws.Cells(iRow, 5).Value
Me.vbNew.Value = ws.Cells(iRow, 6)
Me.vbContractDate.Value = ws.Cells(iRow, 7).Value
Me.vbSticker.Value = ws.Cells(iRow, 8).Value
Me.vbStickerPrev.Value = ws.Cells(iRow, 8).Value
If ws.Cells(iRow, 9) > 0 Then
   Me.vbAccessCode1.Value = ws.Cells(iRow, 9).Value
Else
   Me.vbAccessCode1 = ""
End If
Me.vbLicense1.Value = ws.Cells(iRow, 10).Value
Me.vbExpireDate.Value = ws.Cells(iRow, 11).Value
Me.vbRegistrationOnfile.Value = ws.Cells(iRow, 12).Value
Me.vbVechileDescr.Value = ws.Cells(iRow, 13).Value
Me.vbTelephone1.Value = ws.Cells(iRow, 14).Value
Me.vbTelephone2.Value = ws.Cells(iRow, 15).Value
Me.vbNote.Value = ws.Cells(iRow, 16).Value
If ws.Cells(iRow, 20) > 0 Then
   Me.vbAccessCode2.Value = ws.Cells(iRow, 20).Value
Else
   Me.vbAccessCode2 = ""
End If
Me.vbNoticeDate = ws.Cells(iRow, 25).Value
Me.vbEmail.Value = ws.Cells(iRow, 26).Value
Me.vbInfoBar = ""

' **************************************************************
' *** Need to look up the Historical Tracking Sheet for the
' ***     Second Access Code, or perhaps add another Vlookup column
' **************************************************************

Me.vbRow = iRow

If wkDebugg = True Then
   MsgBox "Item Found ", vbOKOnly + vbInformation, Me.vbAcct.Value & " Item Found"
End If

Me.butSpecial.Visible = True

Me.vbAcct.SetFocus
' *****************************************************
' **  End of butFind_Click() Subroutine             ***
' *****************************************************
End Sub

Private Sub butUpdate_Click()
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

'Lets put this in for cases where we have Multiple Accounts that have multiple RV Lot Spaces
If Me.vbRow.Value = "" Then
   iRowStart = 1
Else
   iRowStart = Me.vbRow.Value
End If

Set ws = Worksheets("Space #")
'iRow = 0

'check for a Account number
If Trim(Me.vbAcct.Value) = "" Then
    Me.vbAcct.SetFocus
    MsgBox " Please Enter An Account "
    Exit Sub
End If

'check for an Existing Account Number

On Error Resume Next
Set Rng = ws.Columns("A:A").Find(What:=Me.vbAcct, After:=ws.Range("A" & iRowStart - 1), _
       SearchOrder:=xlByRows, _
       SearchDirection:=xlNext, LookIn:=xlFormulas)
       
If Err.Number <> 0 Then
   MsgBox "Account Not found! Try Again."
   ClearForm
   Exit Sub
End If

iRow = Rng.Row

If Me.vbAcct.Value <> ws.Cells(iRow, 1).Value Then
   MsgBox "Account Not found! Try Again."
   ClearForm
   Exit Sub
End If

'  Account found , continue processing

If FormEdits(Me.vbAcct, Me.vbSpace, Me) > 0 Then
   ClearForm
   Exit Sub
End If

If FunctionalityEdits(3) > 0 Then
   ClearForm
   Exit Sub
End If

'Check to verify what is on the screen is where the correct row is pointed to !!

If Me.vbAcct.Value = ws.Cells(iRow, 1) And _
   Me.vbSpace = ws.Cells(iRow, 2) Then
' all is ok
ElseIf Me.vbRow <> iRow Then
   MsgBox " Are we on the correct row ?? , Try Again " & _
            Me.vbRow & " vs. " & iRow
   Exit Sub
Else
   MsgBox " There is a serious problem here, call Tech Support"
   Me.vbInfoBar = "Acct: " & Me.vbAcct & " in Space " & Me.vbSpace & _
         " vs. Acct: " & ws.Cells(iRow, 1) & " in Space " & _
         ws.Cells(iRow, 2)
   Exit Sub
End If

wkAcct = ws.Cells(iRow, 1)
wkSpace = ws.Cells(iRow, 2)

'**************************************************
'Unprotect the Space No. WorkSheet
'**************************************************
wkMima = ws.Range("AG1").Value
ws.Unprotect Password:=wkMima

'copy the updated Form data to the Spreadsheet

'Me.vbAcct.Value = ws.Cells(iRow, 1).Value
ws.Cells(iRow, 1) = Me.vbAcct
'  We CAN NOT change the Space Number at this Moment, It is the MAJOR KEY !!!

'ws.Cells(iRow, 2)"" = Me.vbSpace.Value
ws.Cells(iRow, 3) = Me.vbName
ws.Cells(iRow, 4) = Me.vbAddress1
ws.Cells(iRow, 5) = Me.vbAddress2
ws.Cells(iRow, 6) = "Change"
ws.Cells(iRow, 7).Value = Me.vbContractDate.Value
ws.Cells(iRow, 8) = Me.vbSticker
'ws.Cells(iRow, 9) = Me.vbAccessCode1

ws.Cells(iRow, 10) = Me.vbLicense1
ws.Cells(iRow, 11).Value = Me.vbExpireDate.Value
ws.Cells(iRow, 12) = Me.vbRegistrationOnfile
ws.Cells(iRow, 13) = Me.vbVechileDescr
ws.Cells(iRow, 14) = Me.vbTelephone1
ws.Cells(iRow, 15) = Me.vbTelephone2
ws.Cells(iRow, 16) = Me.vbNote
' ws.Cells(iRow, 19) = "=IF($A3>"",VLOOKUP($H3,'Historical tracking'!$E$2:$K$444,7,FALSE),"")"
ws.Cells(iRow, 25) = Me.vbNoticeDate
ws.Cells(iRow, 26) = Me.vbEmail
Me.vbRow = iRow

'**************************************************
'Protect the Space No. WorkSheet Again
'**************************************************
ws.Protect Password:=wkMima, DrawingObjects:=True, Contents:=True, Scenarios:=True _
    , AllowFormattingRows:=True, AllowInsertingRows:=True, _
    AllowDeletingRows:=True

' Now Copy identical Data to the Historical tracking Database

Set ws2 = Worksheets("Historical tracking")
iRow = 2
iRowStart = 2

'find the correct (First instance ! ! ) sticker number
On Error Resume Next
Set Rng = ws2.Columns("E:E").Find(What:=Me.vbSticker, After:=ws2.Range("E" & iRowStart), _
       SearchOrder:=xlByRows, _
       SearchDirection:=xlNext, LookIn:=xlFormulas)
       
If Err.Number <> 0 Then
   MsgBox "Sticker Number found! Fatal Error"
   ClearForm
   Exit Sub
End If

iRow = Rng.Row

If ws2.Cells(iRow, 1) <> Me.vbAcct Then
    MsgBox "Major Error , why is the account number different ? " & _
    ws2.Cells(iRow, 1) & " vs. " & Me.vbAcct & vbCrLf & _
    "This needs some investigation.  Check the Excel spreadsheet."
End If

If Int(ws2.Cells(iRow, 6).Value) = Int(Me.vbAccessCode1.Value) Then
   wkAccess1Change = False
Else
   wkAccess1Change = True
End If

If Int(ws2.Cells(iRow, 7).Value) = Int(Me.vbAccessCode2.Value) Then
   wkAccess2Change = False
Else
   wkAccess2Change = True
End If

If wkAccess1Change = True Or wkAccess2Change = True Then
   ws2.Cells(iRow, 1).EntireRow.Resize(Int(1)).Insert
   ws2.Cells(iRow, 1).Value = Me.vbAcct
   ws2.Cells(iRow, 5) = Me.vbSticker
   ws2.Cells(iRow, 11) = ws2.Cells(iRow + 1, 11)
   If ws2.Cells(iRow + 1, 11) > "" Then
      ws2.Cells(iRow + 1, 11) = ws2.Cells(iRow + 1, 11) & vbCrLf & "Returned Access Card " & _
                   ws2.Cells(iRow + 1, 6) & " on " & Date & _
                   " by " & GetUserName() & " from " & GetMachineName() & _
                    " on " & GetFirstNonLocalIPAddress()
    Else
      ws2.Cells(iRow + 1, 11) = "Returned Access Card " & _
                   ws2.Cells(iRow + 1, 6) & " on " & Date & _
                   " by " & GetUserName() & " from " & GetMachineName() & _
                    " on " & GetFirstNonLocalIPAddress()
        End If
   ws2.Cells(iRow + 1, 12).Value = Date
   ws2.Cells(iRow + 1, 13).Value = Time
   ws2.Cells(iRow + 1, 14) = GetUserName()
   ws2.Cells(iRow + 1, 15) = GetMachineName()
   ws2.Cells(iRow + 1, 16) = GetFirstNonLocalIPAddress()
End If

ws2.Cells(iRow, 2).Value = Right(Me.vbSpace.Value, 3)
If Left(Me.vbSpace, 1) = "M" Then
   ws2.Cells(iRow, 3) = "Micco"
ElseIf Left(Me.vbSpace, 1) = "W" Then
   ws2.Cells(iRow, 3) = "West"
Else
   ws2.Cells(iRow, 3) = "*** ERROR ***"
End If
' We CAN NOT CHANGE the LOCATION CODE
'    NOR CAN WE CHANGE THE STICKER NUMBER - BOTH ARE KEYS
ws2.Cells(iRow, 4) = Me.vbName
ws2.Cells(iRow, 6) = Me.vbAccessCode1.Value
ws2.Cells(iRow, 7) = Me.vbAccessCode2.Value

If ws2.Cells(iRow, 11) > "" Then
   ws2.Cells(iRow, 11) = ws2.Cells(iRow, 11) & vbCrLf & "Updated on " & Date & _
                   " by " & GetUserName() & " from " & GetMachineName() & _
                    " on " & GetFirstNonLocalIPAddress()
   With ws2.Rows(iRow)
        .RowHeight = .RowHeight * 2
   End With
Else
   ws2.Cells(iRow, 11) = "Updated on " & Date & _
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
    ws2.Cells(iRow + 1, 2).Value = Right(Me.vbSpace.Value, 3)
    If Left(Me.vbSpace, 1) = "M" Then
       ws2.Cells(iRow + 1, 3) = "Micco"
    ElseIf Left(Me.vbSpace, 1) = "W" Then
       ws2.Cells(iRow + 1, 3) = "West"
    Else
       ws2.Cells(iRow + 1, 3) = "*** ERROR ***"
    End If
    ws2.Cells(iRow + 1, 4) = Me.vbName
    ws2.Cells(iRow + 1, 6) = Me.vbAccessCode1.Value
    ws2.Cells(iRow + 1, 7) = Me.vbAccessCode2.Value

    ws2.Cells(iRow + 1, 11) = ws2.Cells(iRow + 1, 11) & "Updated on " & Date & _
                   " by " & GetUserName() & " from " & GetMachineName() & _
                    " on " & GetFirstNonLocalIPAddress() & vbCrLf
    ws2.Cells(iRow + 1, 12).Value = Date
    ws2.Cells(iRow + 1, 13).Value = Time
End If

If wkDebugg = True Then
   MsgBox "Space " & wkSpace & " Updated Successfully.", _
          vbOKOnly + vbInformation, "Acct# " & wkAcct & " Updated Successfully."
End If

Me.vbInfoBar = "Account " & Me.vbAcct & " in Space " & Me.vbSpace & _
 " has been successfully updated. "

'clear the data for the Next Screen
UserForm_Initialize
ClearForm

Me.vbAcct.SetFocus

' *****************************************************
' **  End of butUpdate Subroutine                   ***
' *****************************************************
End Sub

Private Sub UserForm_Initialize()
' *****************************************************
' **  Begining of UserForm_Initialize() Subroutine  ***
' *****************************************************

Dim iRow As Long
Dim iRowStart As Long
Dim iLastRow As Long
Dim iBreakInLocationRow As Integer
Dim iSpace As Long
Dim iEditErrorCount As Integer
Dim iListBoxMiccoMaxItem As Integer
Dim iListBoxWestMaxItem As Integer
Dim iListBoxPos As Integer
Dim ws_iRowMaxLimit
Dim ws As Worksheet
Dim Rng As Range
Dim wkLocation As String
Dim wkDebugg As Boolean
Set ws = Worksheets("Space #")
iRow = 1
iRowStart = 2
iLastRow = 1
iEditErrorCount = 0

ws_iRowMaxLimit = 500
If TomsDebugger > 5 Then
   wkDebugg = True
Else
   wkDebugg = False
End If

'find first empty row in database
iLastRow = ws.Cells.Find(What:="*", SearchOrder:=xlRows, _
           SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1
  
'check for the First Empty Account Number where account number is nulls

On Error Resume Next
Set Rng = ws.Columns("A:A").Find(What:="", After:=ws.Range("A" & iRowStart), _
       SearchOrder:=xlByRows, _
       SearchDirection:=xlNext, LookIn:=xlFormulas)
       
If Err.Number <> 0 Then
   MsgBox "Very Very Strange, No Null Account Numbers found! Please Investigate."
   MsgBox " This is a very fatal error, abort mission ! "
   iRow = 1
   Exit Sub
End If

iRow = Rng.Row

If ws.Cells(iRow, 1) <> "" And iLastRow >= ws_iRowMaxLimit Then
    If wkDebugg = True Then
       MsgBox "Empty Spaces Not found! The Lots Must Be FULL."
    End If
    With vbAvailSpaceMiccoList
         .AddItem "LOTS ARE FULL"
    End With
    Exit Sub
ElseIf ws.Cells(iRow, 1) = "" And ws.Cells(iRow, 2) = "" Then
    If wkDebugg = True Then
       MsgBox "Empty Lot Not found! Both Lots are FULL."
    End If
    With vbAvailSpaceMiccoList
         .AddItem "BOTH FULL"
    End With
    Exit Sub
ElseIf Left(ws.Cells(iRow, 2), 1) = "M" Then
'   this is an acceptable piece of data , "M" = Micco Lot
ElseIf Left(ws.Cells(iRow, 2), 1) = "W" Then
'   This is an acceptable data , "W" = West Lot
Else
    If wkDebugg = True Then
       MsgBox "Errow, Blank space found in Spreadsheet! Please remove it please."
    End If
    With vbAvailSpaceMiccoList
         .AddItem "Blank Space"
    End With
    Exit Sub
End If
  
'Let save the location we're dealing with ( should be "M" in all cases! )
wkLocation = Left(ws.Cells(iRow, 2), 1)
'Clear out the two lot list boxes on the screen
vbAvailSpaceMiccoList.Clear
vbAvailSpaceWestList.Clear
  
'Add Micco Available Lots to ListBox Micco
With vbAvailSpaceMiccoList
    Do While ws.Cells(iRow, 2) <> "" And Left(ws.Cells(iRow, 2), 1) = wkLocation
        .AddItem ws.Cells(iRow, 2).Value
        iRowStart = iRow
'       iRowStart = iRow + 1
        On Error Resume Next
        Set Rng = ws.Columns("A:A").Find(What:="", After:=ws.Range("A" & iRowStart), _
                  SearchOrder:=xlByRows, _
                  SearchDirection:=xlNext, LookIn:=xlFormulas)
        If Err.Number <> 0 Then
           MsgBox "Very Very Strange, No Null Account Numbers found after Micco " & _
                  " Lot! Investigation Needed."
           MsgBox " This is a very fatal error, abort mission ! "
           iRow = 1
           Exit Sub
        End If
        iRow = Rng.Row
        ' Lets put in a Range Check, so we dont go for an infinite loop !!!
        If iRow > ws_iRowMaxLimit Then
            Exit Sub
        End If
    Loop
End With
    
    
iListBoxMiccoMaxItem = vbAvailSpaceMiccoList.ListCount
wkLocation = Left(ws.Cells(iRow, 2), 1)

'Add Items to West Lots's ListBox
With vbAvailSpaceWestList
    Do While ws.Cells(iRow, 2) <> "" And Left(ws.Cells(iRow, 2), 1) = wkLocation
        .AddItem ws.Cells(iRow, 2).Value
        iRowStart = iRow
        On Error Resume Next
        Set Rng = ws.Columns("A:A").Find(What:="", After:=ws.Range("A" & iRowStart), _
                  SearchOrder:=xlByRows, _
                  SearchDirection:=xlNext, LookIn:=xlFormulas)
        If Err.Number <> 0 Then
           MsgBox "Very Very Strange, No Null Account Numbers after West Lot" & _
                   " Lot! Please Investigate this."
           MsgBox " This is a very fatal error, abort mission ! "
           iRow = 1
           Exit Sub
        End If
        iRow = Rng.Row
        ' Lets put in a Range Check, so we dont go for an infinite loop !!!
        If iRow > ws_iRowMaxLimit Then
            Exit Sub
        End If
    Loop
End With

iListBoxWestMaxItem = vbAvailSpaceWestList.ListCount
iRowStart = iRow
iBreakInLocationRow = iRow + 1

On Error Resume Next
Set Rng = ws.Columns("B:B").Find(What:="*", After:=ws.Range("B" & iRowStart), _
            SearchOrder:=xlByRows, _
            SearchDirection:=xlNext, LookIn:=xlFormulas)
If Err.Number <> 0 Then
' I will assume that this is the end of the list
    iRow = 999
Else
    iRow = Rng.Row
End If

'  Tom, do we have some Un - Sorted Location data at the bottom of the spreadsheet ??
wkLocation = Left(ws.Cells(iRow, 2), 1)

'REMOVE Items to West Lots's ListBox
With vbAvailSpaceWestList
    Do While ws.Cells(iRow, 1) <> "" And (ws.Cells(iRow, 2) <> "") And _
                 ((iRow < iLastRow) And (iRow >= iBreakInLocationRow - 5))
        'Check if this Item is on the List box, if it is remove it.
        iListBoxPos = iListBoxWestMaxItem
        Do While iListBoxPos > 0
            If .List(iListBoxPos) = ws.Cells(iRow, 2) Then
                .RemoveItem (iListBoxPos)
            End If
            iListBoxPos = iListBoxPos - 1
        Loop
        'Fetch the next spreadsheet item that has an Location number
        iRowStart = iRow
        On Error Resume Next
        Set Rng = ws.Columns("B:B").Find(What:="*", After:=ws.Range("B" & iRowStart), _
                  SearchOrder:=xlByRows, _
                  SearchDirection:=xlNext, LookIn:=xlFormulas)
        If Err.Number <> 0 Then
           ' I will assume that this is the end of the list
           iRow = 999
        Else
           iRow = Rng.Row
         End If
        ' Lets put in a Range Check, so we dont go for an infinite loop !!!
        If iRow > ws_iRowMaxLimit Then
            Exit Sub
        End If
    Loop
End With

iRow = iBreakInLocationRow

'REMOVE Items from the MICCO Lots's ListBox
With vbAvailSpaceMiccoList
    Do While ws.Cells(iRow, 1) <> "" And (ws.Cells(iRow, 2) <> "") And _
                 ((iRow < iLastRow) And (iRow >= iBreakInLocationRow))
        'Check if this Item is on the List box, if it is remove it.
        iListBoxPos = iListBoxMiccoMaxItem
        For iListBoxPos = 0 To iListBoxMiccoMaxItem
            If .List(iListBoxPos) = ws.Cells(iRow, 2) Then
                .RemoveItem (iListBoxPos)
            End If
        Next
        'Fetch the next spreadsheet item that has an Location number
        iRowStart = iRow
        On Error Resume Next
        Set Rng = ws.Columns("B:B").Find(What:="*", After:=ws.Range("B" & iRowStart), _
                  SearchOrder:=xlByRows, _
                  SearchDirection:=xlNext, LookIn:=xlFormulas)
        If Err.Number <> 0 Then
           ' I will assume that this is the end of the list
           iRow = 999
        Else
           iRow = Rng.Row
         End If
        ' Lets put in a Range Check, so we dont go for an infinite loop !!!
        If iRow > ws_iRowMaxLimit Then
            Exit Sub
        End If
    Loop
End With

iRow = iBreakInLocationRow

' *****************************************************
' **  End of UserForm_Initialize() Subroutine       ***
' *****************************************************
End Sub

Private Sub UserForm_Terminate()
    Application.Visible = True
'    MsgBox " we are closing the window by the yellow X "
'    MsgBox " Don't forget to save your changes. "
End Sub

Private Sub ClearForm()
' *****************************************************
' **  Begining of ClearForm() Subroutine            ***
' *****************************************************
   Me.vbAcctPrev = ""
   Me.vbSpace = ""
   Me.vbSpace2 = ""
   Me.vbSpacePrev = ""
   Me.vbName = ""
   Me.vbAddress1 = ""
   Me.vbAddress2 = ""
   Me.vbNew = "New"
   Me.vbContractDate.Value = Date
   Me.vbEffectiveDate.Value = Date - 2
   Me.vbNoticeDate = ""
   Me.vbSticker = ""
   Me.vbStickerPrev = ""
   Me.vbAccessCode1 = ""
   Me.vbAccessCode2 = ""
   Me.vbLicense1 = ""
   Me.vbLicense2 = ""
   Me.vbExpireDate.Value = ""
   Me.vbRegistrationOnfile = "N"
   Me.vbVechileDescr = ""
   Me.vbTelephone1 = ""
   Me.vbTelephone2 = ""
   Me.vbNote = ""
   Me.vbEmail = ""
'   Me.vbInfoBar = ""
   Me.vbRow = 1
   
   Me.butSpecial.Visible = False
   wkPreventInfiniteLoop = 0
   wkInfiniteLoopLimit = 2000

' *****************************************************
' **  End of ClearForm() Subroutine                 ***
' *****************************************************
End Sub
Private Sub butSpecial_Click()
    CRUD.Hide
    SpecialFunc.vb2Acct = Me.vbAcct
    SpecialFunc.vb2Access1 = Me.vbAccessCode1
    SpecialFunc.vb2Access3 = Me.vbAccessCode2
    SpecialFunc.vb2Space1 = Me.vbSpace
    
    With SpecialFunc.vb2ReasonList1
      .AddItem "Replacing Lost Card"
      .AddItem "Surrendering Defective Card"
      .AddItem "Extra Security Card"
      .AddItem ""
      .AddItem "Other"
    End With
   
    With SpecialFunc.vb2ReasonList2
      .AddItem "Replacing Lost Card"
      .AddItem "Surrendering Defective Card"
      .AddItem "Extra Security Card"
      .AddItem ""
      .AddItem "Other"
    End With
    
    SpecialFunc.Show
End Sub

Public Function FormEdits(wkAcct As String, _
                           wkSpace As String, _
                           wkForm As CRUD) As Integer
' *****************************************************
' **  Begining of FormEdits() Function              ***
' *****************************************************
Dim iEditErrorCount As Integer
Dim wk2Space As Integer

iEditErrorCount = 0

'check for a Account number
If Trim(wkAcct) = "" Then
    wkForm.vbAcct.SetFocus
    MsgBox "Account Number is required, Please complete the form"
    iEditErrorCount = 0
    FormEdits = 1
ElseIf Int(wkAcct) > 6999 Then
    wkForm.vbAcct.SetFocus
    MsgBox "Account Number is out of Range, " & _
        "The Account Number can not be greater than 6999."
    iEditErrorCount = 0
    FormEdits = 2
End If

'Check for a Space Number
If Trim(wkSpace) = "" Then
    wkForm.vbSpace.SetFocus
    MsgBox "Space Number is required, Please complete the form"
    iEditErrorCount = 0
    FormEdits = 3
End If

wkSpace = StrConv(wkSpace, vbUpperCase)

' More Space number edits
wk2Space = Right(wkSpace, 3)

If wkSpace = "M205A" Then
ElseIf wk2Space = 0 Then
    wkForm.vbSpace.SetFocus
    MsgBox wkForm.vbSpace & " is not a valid space location , no space labled Zero. " & _
        " Please change."
    iEditErrorCount = iEditErrorCount + 1
ElseIf wk2Space < 274 And Left(wkSpace, 1) = "M" Then
ElseIf wk2Space < 97 And Left(wkSpace, 1) = "W" Then
ElseIf Left(wkSpace, 1) = "M" Then
    wkForm.vbSpace.SetFocus
    MsgBox wkSpace & " is not a valid space location , max. space is 273 " & _
        "for the Micco Lot. Please change."
    iEditErrorCount = iEditErrorCount + 1
ElseIf Left(wkSpace, 1) = "W" Then
    wkForm.vbSpace.SetFocus
    MsgBox wkSpace & " is not a valid space location , max. space is 96 " & _
        "for the West Lot. Please change."
    iEditErrorCount = iEditErrorCount + 1
Else
    wkForm.vbSpace.SetFocus
    MsgBox wkSpace & " is not a valid Lot Location , Lots can only " & _
        "be M for Micco and W for West. Please update first Character."
    iEditErrorCount = iEditErrorCount + 1
End If

'Check for a Name
If Trim(wkForm.vbName) = "" Then
    wkForm.vbName.SetFocus
    MsgBox "Account name is required, Please complete the form"
    iEditErrorCount = iEditErrorCount + 1
End If

wkForm.vbName = StrConv(wkForm.vbName, vbProperCase)

'Check for an Address
If Trim(wkForm.vbAddress1) = "" Then
    wkForm.vbAddress1.SetFocus
    MsgBox "Address is required, Please complete the form"
    iEditErrorCount = iEditErrorCount + 1
End If

'Check for a Contract Date
If Trim(wkForm.vbContractDate) = "" Then
    wkForm.vbContractDate.Value = Date
    MsgBox "Contract Date is required, Today's Date will be put as a default."
    iEditErrorCount = iEditErrorCount + 1
End If

'Check for a Veichle Sticker Number
If Trim(wkForm.vbSticker) = "" Then
    wkForm.vbSticker.SetFocus
    MsgBox "Sticker Number is required, Please complete the form"
    iEditErrorCount = iEditErrorCount + 1
ElseIf wkForm.vbSticker.Value > 2399 And wkForm.vbSticker < 823 Then
    wkForm.vbSticker.SetFocus
    MsgBox "Sticker Number is out of range, verify Sticker Number again please."
    iEditErrorCount = iEditErrorCount + 1
End If

'Check for a Vechile Description
If Trim(wkForm.vbVechileDescr) = "" Then
    wkForm.vbVechileDescr.SetFocus
    MsgBox "Vechile description is required, Please complete the form"
    iEditErrorCount = iEditErrorCount + 1
End If

'Check for at least One Telephone Number
If Trim(wkForm.vbTelephone1) = "" Then
    wkForm.vbTelephone1.SetFocus
    MsgBox "At least one telephone number is required, Please complete the form"
    iEditErrorCount = iEditErrorCount + 1
End If

If Len(wkSpace) = 4 Then
Else
    wkForm.vbSpace.SetFocus
    MsgBox wkSpace & " is not a valid space location , max. space is 273. Please change."
    iEditErrorCount = iEditErrorCount + 1
End If

' Let's put some more edits

If Len(wkAcct) > 4 Then
    wkForm.vbAcct.SetFocus
    MsgBox wkAcct & " is not a valid account number, too many digits. Please edit."
    iEditErrorCount = iEditErrorCount + 1
End If

If Len(wkForm.vbSticker) = 4 Then
Else
    wkForm.vbSticker.SetFocus
    MsgBox wkForm.vbSticker & " is not a valid sticker number, all stickers have 4 digits. Try again."
    iEditErrorCount = iEditErrorCount + 1
End If

If iEditErrorCount > 0 Then
   FormEdits = iEditErrorCount + 100
Else
   FormEdits = 0
End If

' *****************************************************
' **  End of Function FormEdits() Function          ***
' *****************************************************
End Function

Function FunctionalityEdits(olTransactionCode As Integer) As Integer
' *****************************************************
' **  Beginning of Function FunctionalityEdits()    ***
' *****************************************************
Dim wkDebugg As Boolean
Dim wkDebuggLvl As Integer

If wkDebuggLvl = TomsDebugger > 2 Then
   wkDebugg = True
Else
   wkDebugg = False
End If

If olTransactionCode = 3 Then
   If (Me.vbAcct <> Me.vbAcctPrev) Or _
      (Me.vbSpace <> Me.vbSpacePrev) Or _
      (Me.vbSticker <> Me.vbStickerPrev) Then
      MsgBox "The Account, Space or Sticker fields can NOT be " & _
             " updated at this time, please do this manually " & _
             " on the excel spreadsheet.  "
      FunctionalityEdits = 101
      Exit Function
    End If
End If

If Me.vbPrintForm.Value = True Then
   MsgBox "Print Form Option - Not Yet Implemented."
End If

If Me.vbSendFaxToAdmin.Value = True Then
'   MsgBox "Fax Option - Not Yet Implemented."
'  Mail_small_Text_Outlook
   Mail_small_Text_Outlook (1)
End If

If Me.vbEmailSecureIdCardInfo.Value = True Then
   'MsgBox "Email Option - Not Yet Implemented."
   Mail_Text_From_Txtfile_Outlook (olTransactionCode)
End If

If wkDebugg = True Then
    MsgBox "Contract Sucessfully added", vbOKOnly + vbInformation, _
       "New Contract Successfully Added"
End If

FunctionalityEdits = 0
' *****************************************************
' **  End of Function FunctionalityEdits()          ***
' *****************************************************
End Function

Sub Mail_Text_From_Txtfile_Outlook(olTransactionCode As Integer)
' **********************************************************
' **  Begining of Mail_From_Txtfile_Outlook() Subroutine ***
' **********************************************************
'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
'Working in Office 2000-2013
    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
'        .To = "CalendarCord@bbrd.org"
        .To = "SecureDoor@ADSSecurity.com;"
        .CC = "suecuddie@bbrd.org;calendarcord@bbrd.org;"
        .BCC = ""
        If olTransactionCode = 1 Then
            .Subject = "[YELLOWSTONE]Activation Updates of Secure-Id Access Cards - Barefoot Bay Recreation District"
            If Left(Me.vbSpace.Value, 1) = "M" Then
              .Body = "To Whom It may Concern:" & vbCrLf & vbCrLf & _
                      "Please change the following for the MICCO lot ONLY :" & _
                       vbCrLf & vbCrLf & "CHANGE " & Me.vbAccessCode1 & "     " & "    to " & _
                       Me.vbName & "     Account: " & Me.vbAcct
            ElseIf Left(Me.vbSpace.Value, 1) = "W" Then
              .Body = "To Whom It may Concern:" & vbCrLf & vbCrLf & _
                      "Please change the following for the  WEST lot ONLY :" & _
                       vbCrLf & vbCrLf & "CHANGE " & Me.vbAccessCode1 & "     " & "    to " & _
                       Me.vbName & "     Account: " & Me.vbAcct
            Else
              .Body = "INVALID Lot CODE Entered "
            End If
        ElseIf olTransactionCode = 2 Then
           .Subject = "[YELLOWSTONE]De-Activation of Secure-Id Access Cards - Barefoot Bay Recreation District"
           .Body = "To Whom It may Concern:" & vbCrLf & vbCrLf & _
                   "Please De-Activate the following Secure-Id Card(s):" & _
                   vbCrLf & vbCrLf & Me.vbAccessCode1 & vbCrLf & Me.vbAccessCode2
        ElseIf olTransactionCode = 3 Then
            .Subject = "[YELLOWSTONE]Activation Updates of Secure-Id Access Cards - Barefoot Bay Recreation District"
            If Left(Me.vbSpace.Value, 1) = "M" Then
              .Body = "To Whom It may Concern:" & vbCrLf & vbCrLf & _
                      "Please change the following for the MICCO lot ONLY :" & _
                       vbCrLf & vbCrLf & "CHANGE " & Me.vbAccessCode1 & "     " & "    to " & _
                       Me.vbName & "     Account: " & Me.vbAcct
            ElseIf Left(Me.vbSpace.Value, 1) = "W" Then
              .Body = "To Whom It may Concern:" & vbCrLf & vbCrLf & _
                      "Please change the following for the  WEST lot ONLY :" & _
                       vbCrLf & vbCrLf & "CHANGE " & Me.vbAccessCode1 & "     " & "    to " & _
                       Me.vbName & "     Account: " & Me.vbAcct
            Else
              .Body = "INVALID Lot CODE Entered "
            End If
        Else
           .Subject = "[YELLOWSTONE]InValid RV Storage Transaction Type sent"
        End If
        '.Body = GetBoiler("C:\test.txt")
        'You can add a file like this
        '.Attachments.Add ("C:\test.txt")
'        .Send   'or use .Display
'        .Send
        .Display
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
' **********************************************************
' **    End of Mail_From_Txtfile_Outlook() Subroutine    ***
' **********************************************************
End Sub

Private Sub butQuit_Click()
    Application.Visible = True
    Unload Me
End Sub

Function GetBoiler(ByVal sFile As String) As String
'Dick Kusleika
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(sFile).OpenAsTextStream(1, -2)
    GetBoiler = ts.readall
    ts.Close
End Function

