VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RETRIEVE 
   Caption         =   "BAREFOOT BAY RECREATION DISTRICT  -  EZ BADGES  BFB BADGE CHECKER  .    .   .  version 0.33 (July 31th, 2015)"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11700
   OleObjectBlob   =   "RETRIEVE.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RETRIEVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TomsDebugg, _
    wkTopWindowPos, _
    wkLeftWindowPos, _
    wkExcelVersionUsed, _
    wkNoteDays, _
    wkAnnoyLimit, _
    wkAnnoyCnt, _
    wkNagFlag, _
    wkRestrictedListDays  As Integer
    
Dim retVal As Long
Dim wkCloseFile  As Boolean

Dim ws1 As Worksheet, _
    ws2 As Worksheet, _
    ws3 As Worksheet, _
    ws4 As Worksheet, _
    ws5 As Worksheet, _
    ws9 As Worksheet

Dim wb1 As Workbook, _
    wb2 As Workbook, _
    wb3 As Workbook, _
    wb4 As Workbook, _
    wb5 As Workbook
    
Dim wkAcct, _
    wsMasterURL, _
    wsMasterURLSuffix, _
    wsNoteFileHund, _
    wsNoteFileOneThou, _
    wsNoteFileTwoThou, _
    wsNoteFileThreeThou, _
    wsNoteFileFourThou, _
    wsNoteFileFiveThou, _
    wsNoteFileSixThou As String

Dim vFile1, _
    vFile2, _
    vFile3, _
    vFile4, _
    vFile5, _
    vFile21, _
    vFile22, _
    vFile23, _
    vFile24, _
    vFile25, _
    vFile26, _
    vFile27, _
    vFile28, _
    vFile29 As Variant

Private Sub butFind_Click()
' *****************************************************
' **  Begining of butFind_Click() Subroutine        ***
' *****************************************************
Dim iRow As Long, _
    iRowStart As Long, _
    iRowLastItem As Long, _
    jRow, _
    wkDayDiff   As Long
    
Dim Rng1 As Range, _
    Rng2 As Range, _
    Rng3 As Range, _
    Rng4 As Range

Dim wkHardSearch As Boolean
Dim wkDebugg As Boolean
Dim wkFileDate As Date

wkHardSearch = True
If TomsDebugger > 8 Then
   wkDebugg = True
Else
   wkDebugg = False
End If

'Let Intialize our fields
ClearForm3
ClearForm2

'check for an Account number
If Trim(Me.vbAcct.Value) = Empty Then
    Me.vbAcct.SetFocus
    MsgBox " Please Enter An Account "
    Exit Sub
End If
If FormEdits() > 0 Then
    Me.vbAcct.SetFocus
    MsgBox " Errors Encountered, please correct "
    Exit Sub
End If

iRowStart = 0
iRow = 0
jRow = 2
Range("V" & jRow) = Date
Range("W" & jRow) = Time()
Range("X" & jRow) = GetUserName()
Range("Y" & jRow) = GetMachineName()
Range("Z" & jRow) = GetFirstNonLocalIPAddress()
Application.StatusBar = "Fetching Customer Notes Spreadsheet for Account " & _
                            Me.vbAcct
'*********************************************************************
'Open the FIRST target workbook
'*********************************************************************
If Me.vbAcctNoteFileName = Empty Then
    vFile2 = Application.GetOpenFilename("Excel-files,*.xls", _
            1, "Select Correct Note File To Open", , False)
    'if the user didn't select a file, exit sub
    If TypeName(vFile2) = "Boolean" Then Exit Sub
    Workbooks.Open Filename:=vFile2, ReadOnly:=True
    vbAcctNoteFileName = vFile2
    'Set targetworkbook
    Set wb2 = ActiveWorkbook
ElseIf Me.vbAcctNoteFileName = getCorrectNotesFile(Me.vbAcct.Value) Then
    Workbooks.Open Filename:=Me.vbAcctNoteFileName, ReadOnly:=True
    Set wb2 = ActiveWorkbook
Else
    Me.vbAcctNoteFileName = getCorrectNotesFile(Me.vbAcct.Value)
    Workbooks.Open Filename:=Me.vbAcctNoteFileName, ReadOnly:=True
    Set wb2 = ActiveWorkbook
    vFile2 = Me.vbAcctNoteFileName
End If

Application.StatusBar = "Processing Customer Notes Spreadsheet for Account " & _
                            Me.vbAcct

Set ws2 = wb2.Worksheets("Sheet1")
Range("A1").Select

Set Rng1 = Sheets("Sheet1").Columns("A").Find(What:=Me.vbAcct, LookIn:=xlValues, lookat:=xlWhole)

If Err.Number <> 0 Then
   MsgBox "Account Not found on the Notes List! Try Again."
   Me.vbInfoBar = Err.Description
   ClearForm
   retVal = buildCheckList(2, False)
   If wkCloseFile Then wb2.Close SaveChanges:=False
   Exit Sub
End If

If Rng1 Is Nothing Then
    MsgBox " Account Not found on Note Spreadsheet! Try Again."
    ClearForm
    retVal = buildCheckList(2, False)
    If wkCloseFile Then wb2.Close SaveChanges:=False
    Exit Sub
End If
Range("D" & Rng1.Row).Select

'
'    MsgBox " The modification date for this note file is " & FileDateTime(Me.vbAcctNoteFileName) & _
'                  "."
'
wkFileDate = FileDateTime(Me.vbAcctNoteFileName)

If wkNagFlag = 1 Then
    wkDayDiff = DateDiff("d", wkFileDate, Date)
    If wkDayDiff > wkNoteDays Then
        If wkAnnoyCnt < wkAnnoyLimit Then
            Call ShowNagMsgForDate("The Customer Notes file has not been updated in ", _
                       " days. Please refresh using the 'Get Latest Notes' Button", _
                                wkFileDate, True)
        End If
        wkAnnoyCnt = wkAnnoyCnt + 1
    End If
End If

Application.StatusBar = "Building Checklist item for Customer Notes for Account " & _
                            Me.vbAcct
retVal = buildCheckList(2, True)
iRow = Rng1.Row

Me.vbAddress2.Value = Cells(iRow, 3).Value
Me.vbNote.Value = Cells(iRow, 4).Value
Me.vbRowNote.Value = iRow
ws2.Range("A" & iRow).Select
ws1.Activate
'  Range("I" & jRow & ":O" & jRow).Value = ws2.Range("A" & iRow & ":G" & iRow).Value
'  We need special Code for excel 2003 falling over if cell has too much information.
'   This is specific the the Customer Notes Column.  When we isolate the copy to Just THAT Column , it Works!
Range("I" & jRow & ":K" & jRow).Value = ws2.Range("A" & iRow & ":C" & iRow).Value
Range("L" & jRow).Value = ws2.Range("D" & iRow).Value
Range("M" & jRow & ":O" & jRow).Value = ws2.Range("E" & iRow & ":G" & iRow).Value

If wkCloseFile Then wb2.Close SaveChanges:=False
Application.StatusBar = "Fetching Resident Address Spreadsheet for Account " & _
                            Me.vbAcct

'******************************************************
'Open the SECOND target workbook
'******************************************************
If Me.vbAddrListFileName = "" Then
    vFile3 = Application.GetOpenFilename("Excel-files,*.xls", _
            1, "Select Current Resident Address File", , False)
    'if the user didn't select a file, exit sub
    If TypeName(vFile3) = "Boolean" Then Exit Sub
    Workbooks.Open Filename:=vFile3, ReadOnly:=True
    vbAddrListFileName = vFile3
    'Set targetworkbook
    Set wb3 = ActiveWorkbook
Else
    Workbooks.Open Filename:=Me.vbAddrListFileName, ReadOnly:=True
    Set wb3 = ActiveWorkbook
End If
'
'    MsgBox " The modification date for Resident file is " & FileDateTime(Me.vbAddrListFileName) & _
'                  "."
'
wkFileDate = FileDateTime(Me.vbAddrListFileName)

If wkNagFlag = 1 Then
    wkDayDiff = DateDiff("d", wkFileDate, Date)
    If wkDayDiff > wkRestrictedListDays Then
        If wkAnnoyCnt < wkAnnoyLimit Then
            Call ShowNagMsgForDate(" The Resident Address file has not been updated in ", _
                 " days. Please inform mangagement and refresh using the 'Get Latest Notes' Button", _
                                wkFileDate, True)
        End If
        wkAnnoyCnt = wkAnnoyCnt + 1
    End If
End If

Application.StatusBar = "Processing Resident Address Spreadsheet for Account " & _
                            Me.vbAcct
Range("A1").Select
Set ws3 = wb3.Worksheets("Sheet1")

'search for a Account Number on the BFB Address List

Set Rng2 = Sheets("Sheet1").Columns("A").Find(What:=Me.vbAcct, LookIn:=xlValues, lookat:=xlWhole)

Me.vbInfoBar = Err.Description
If Err.Number <> 0 Then
   MsgBox "Problem searching on Address List! Call Systems."
   ClearForm
   retVal = buildCheckList(3, False)
   If wkCloseFile Then wb3.Close SaveChanges:=False
   Exit Sub
End If

If Rng2 Is Nothing Then
    MsgBox " Account Not found on Address Spreadsheet! Try Again."
    ClearForm
    retVal = buildCheckList(3, False)
    If wkCloseFile Then wb3.Close SaveChanges:=False
    Exit Sub
End If

iRow = Rng2.Row
ws3.Range("A" & iRow).Select

Me.vbAcct.Value = Cells(iRow, 1).Value
Me.vbAcctPrev = Cells(iRow, 1).Value
' ----- ******************************* -----
Me.vbName.Value = Cells(iRow, 2).Value
Me.vbTaxIdCounty = Cells(iRow, 3).Value
Me.vbTaxIdCountyPrev = Cells(iRow, 3).Value
Me.vbAddress1.Value = Cells(iRow, 4).Value & " " & Cells(iRow, 5).Value
Me.vbRow = iRow
Application.StatusBar = "Building Checklist for the Resident Address for Account " & _
                            Me.vbAcct
retVal = buildCheckList(3, True)

ws1.Activate
Range("D" & jRow & ":H" & jRow).Value = ws3.Range("A" & iRow & ":E" & iRow).Value
If wkCloseFile Then wb3.Close SaveChanges:=False

Application.StatusBar = "Fetching Restricted List Spreadsheet for Account " & _
                            Me.vbAcct
'*****************************************************
'Open the Restricted List workbook
'*****************************************************
If Me.vbRestrictedListFileName = Empty Then
    vFile4 = Application.GetOpenFilename("Excel-files,*.xls", _
            1, "Select Restricted List File To Open", , False)
    'if the user didn't select a file, exit sub
    If TypeName(vFile4) = "Boolean" Then Exit Sub
    Workbooks.Open Filename:=vFile4, ReadOnly:=True
    vbRestrictedListFileName = vFile4
    Set wb4 = ActiveWorkbook
Else
    Workbooks.Open Filename:=vbRestrictedListFileName, ReadOnly:=True
    Set wb4 = ActiveWorkbook
End If

'
'    MsgBox " The modification date for Restricted List file is " & FileDateTime(Me.vbRestrictedListFileName) & _
'                  "."
'
wkFileDate = FileDateTime(Me.vbRestrictedListFileName)

If wkNagFlag = 1 Then
    wkDayDiff = DateDiff("d", wkFileDate, Date)
    If wkDayDiff > wkRestrictedListDays Then
        If wkAnnoyCnt < wkAnnoyLimit Then
            Call ShowNagMsgForDate(" The Resident Address file has not been updated in ", _
                 " days. Please inform mangagement and refresh using the 'Get Latest Notes' Button", _
                                wkFileDate, True)
        End If
        wkAnnoyCnt = wkAnnoyCnt + 1
    End If
End If

Application.StatusBar = "Processing Restricted List Spreadsheet for Account " & _
                            Me.vbAcct
Set ws4 = wb4.Worksheets("Sheet1")
Range("A1").Select

Set Rng4 = Sheets("Sheet1").Columns("A").Find(What:=Me.vbAcct, LookIn:=xlValues, lookat:=xlWhole)

If Err.Number <> 0 Then
   MsgBox "Problem Processing the Restricted List File ! Call Systems."
   Me.vbInfoBar = Err.Description
   ClearForm
   retVal = buildCheckList(4, False)
   If wkCloseFile Then wb4.Close SaveChanges:=False
   Exit Sub
End If

Application.StatusBar = "Building Checklist for Restricted List for Account " & _
                            Me.vbAcct
If Rng4 Is Nothing Then
'    MsgBox " Account Not found on Restricted Listed Spreadsheet! Good, we may continue processing."
'    ws1.Activate
    ws1.Range("P" & jRow) = "Account " & _
          Me.vbAcct & " Passed Restricted List Test"
    ws1.Range("Q" & jRow & ":S" & jRow).Value = ""
    retVal = buildCheckList(4, True)
    iRow = 1
Else
    iRow = Rng4.Row
    retVal = buildCheckList(4, False)
    ws4.Range("A" & iRow).Select
'    ws1.Activate
    ws1.Range("P" & jRow & ":S" & jRow).Value = ws4.Range("A" & iRow & ":D" & iRow).Value
End If

Me.vbRowRestrictedList.Value = iRow
ws4.Range("A" & iRow).Select
If wkCloseFile Then wb4.Close SaveChanges:=False

'*****************************************************
' Connect to the Brevard County Tax Website and
'   Pull in the data
'*****************************************************

Application.StatusBar = "Fetching Resident Tax Data for Account " & _
                            Me.vbAcct
retVal = FetchTaxId(Me.vbTaxIdCounty, wb1, ws1)
Application.StatusBar = "Processing Options ... "
If retVal = 0 Then
   Me.vbTaxIdCountyPrev = Me.vbTaxIdCounty
ElseIf retVal = 9010 Then
   ClearForm
   MsgBox "Error - Tax Id is a required Field ! Try Again."
   Me.vbTaxIdCounty.SetFocus
   retVal = buildCheckList(5, False)
   Exit Sub
ElseIf retVal = 9001 Then
   ClearForm2
   MsgBox "Error with Connection to CountyTaxRecord Object ! Try Again."
   Me.vbTaxIdCounty.SetFocus
   retVal = buildCheckList(5, False)
   Exit Sub
ElseIf retVal = 9003 Then
   ClearForm2
   MsgBox "Problem Accessing Tax Account using Website ! Try Again."
   Me.vbTaxIdCounty.SetFocus
   retVal = buildCheckList(5, False)
   Exit Sub
Else
   ClearForm2
   MsgBox "Error - Invalid return code on funtion ! Call Systems."
   Me.vbTaxIdCounty.SetFocus
   retVal = buildCheckList(5, False)
   Exit Sub
End If
retVal = buildCheckList(5, True)

Range("A" & jRow).Select

Me.vbWebInfo1 = Cells(1, 1).Value
If Cells(16, 1).Value = "VERY GOOD" Then
    Me.vbWebInfo2 = Cells(12, 1).Value
    Me.vbWebInfo2.BackColor = RGB(0, 255, 0)
    Me.vbWebInfo3 = Cells(13, 1).Value
    Me.vbWebInfo3.BackColor = RGB(0, 255, 0)
    retVal = buildCheckList(6, True)
Else
    Me.vbWebInfo2 = Cells(12, 1).Value
'    Me.vbWebInfo2.BackColor = &HCE00E0 PURPLE
    Me.vbWebInfo2.BackColor = &HFF
    Me.vbWebInfo3 = ws1.Range("A13").Value
'    Me.vbWebInfo3.BackColor = &HCE00E0 PURPLE
    Me.vbWebInfo3.BackColor = &HFF
    retVal = buildCheckList(6, False)
End If
'*****************************************************
'  Write to consolidated log file
'
'*****************************************************
Application.StatusBar = "Writing to the Log File ... "
retVal = WriteConsolitdatedLog(wb1, ws1)
Application.StatusBar = "Processing Options ... "

If Me.vbTaxIdCounty <> "" Then
    FunctionalityEdits
End If

If wkDebugg = True Then
   MsgBox "Item Found ", vbOKOnly + vbInformation, Me.vbAcct.Value & " Item Found"
End If

If wkDebugg = True Then
   MsgBox "Item Found ", vbOKOnly + vbInformation, Me.vbAcct.Value & " Item Found"
End If
Me.vbInfoBar = "Processing Complete for Account " & Me.vbAcct & "."
Application.StatusBar = Me.vbInfoBar

Me.vbAcct.SetFocus
retVal = buildCheckList(7, True)
' *****************************************************
' **  End of butFind_Click() Subroutine             ***
' *****************************************************
End Sub

Private Sub butNoteUpdate_Click()
' *****************************************************
' **  Begining of butNoteUpdate_Click() Subroutine  ***
' *****************************************************
Dim fso As Object

Me.vbInfoBar = "Initializing the copying of Production Files for ALL Accounts."
Application.StatusBar = Me.vbInfoBar

Set fso = VBA.CreateObject("Scripting.FileSystemObject")

If fso.FileExists(vFile21) = False Then
   MsgBox vFile21 & " does not exist , Serious problem, Call Systems "
   Exit Sub
End If
fso.CopyFile Source:=vFile21, Destination:=vFile3
If fso.FileExists(vFile22) = False Then
   MsgBox vFile21 & " does not exist , Serious problem, Call Systems "
   Exit Sub
End If
fso.CopyFile Source:=vFile22, Destination:=vFile4

If fso.FileExists(vFile23) = False Then
   MsgBox vFile23 & " does not exist , Serious problem, Call Systems "
   Exit Sub
End If
fso.CopyFile Source:=vFile23, Destination:=wsNoteFileHund
If fso.FileExists(vFile24) = False Then
   MsgBox vFile24 & " does not exist , Serious problem, Call Systems "
   Exit Sub
End If
fso.CopyFile Source:=vFile24, Destination:=wsNoteFileOneThou
If fso.FileExists(vFile25) = False Then
   MsgBox vFile25 & " does not exist , Serious problem, Call Systems "
   Exit Sub
End If
fso.CopyFile Source:=vFile25, Destination:=wsNoteFileTwoThou
If fso.FileExists(vFile26) = False Then
   MsgBox vFile26 & " does not exist , Serious problem, Call Systems "
   Exit Sub
End If
fso.CopyFile Source:=vFile26, Destination:=wsNoteFileThreeThou
If fso.FileExists(vFile27) = False Then
   MsgBox vFile27 & " does not exist , Serious problem, Call Systems "
   Exit Sub
End If
fso.CopyFile Source:=vFile27, Destination:=wsNoteFileFourThou
If fso.FileExists(vFile28) = False Then
   MsgBox vFile28 & " does not exist , Serious problem, Call Systems "
   Exit Sub
End If
fso.CopyFile Source:=vFile28, Destination:=wsNoteFileFiveThou
If fso.FileExists(vFile29) = False Then
   MsgBox vFile29 & " does not exist , Serious problem, Call Systems "
   Exit Sub
End If
fso.CopyFile Source:=vFile29, Destination:=wsNoteFileSixThou
'MsgBox "Feature Not Implemented Yet. "

Me.vbInfoBar = "Finished the replication of Production Files to Database."
Application.StatusBar = Me.vbInfoBar

' *****************************************************
' **  Ending of butNoteUpdate_Click() Subroutine    ***
' *****************************************************
End Sub

Private Sub butQuit_Click()
    Application.Visible = True
    Unload Me
    Application.Quit
End Sub
Private Sub butClear_Click()
    UserForm_Initialize
End Sub

Private Sub ClearForm()
' *****************************************************
' **  Begining of ClearForm() Subroutine            ***
' *****************************************************
'   Me.vbAcctPrev = ""
'   Me.vbTaxIdCountyPrev = ""
   Me.vbAcct = ""
   Me.vbTaxIdCounty = ""
   Me.vbName = ""
   Me.vbAddress1 = ""
   Me.vbAddress2 = ""
   Me.vbNote = ""
'   Me.vbInfoBar = ""
   Me.vbRow = 1
   Me.vbRowNote = 1
   Me.vbRowRestrictedList = 1
   ClearForm3
   ClearForm2
' *****************************************************
' **  End of ClearForm() Subroutine                 ***
' *****************************************************
End Sub

Private Sub ClearForm3()
' *****************************************************
' **   Begining of ClearForm3() Subroutine          ***
' *****************************************************
   Me.vbCheckList1 = ""
   Me.vbCheckList2 = ""
   Me.vbCheckList3 = ""
   Me.vbCheckList4 = ""
   Me.vbCheckList5 = ""
   Me.vbCheckList6 = ""
   Me.vbCheckList7 = ""
   
   Me.vbCheckList1.Caption = "PASSED"
   Me.vbCheckList2.Caption = "PASSED"
   Me.vbCheckList3.Caption = "PASSED"
   Me.vbCheckList4.Caption = "PASSED"
   Me.vbCheckList5.Caption = "PASSED"
   Me.vbCheckList6.Caption = "PASSED"
   Me.vbCheckList7.Caption = "PASSED"
   
   Me.vbCheckList1.BackColor = RGB(0, 255, 0)
   Me.vbCheckList2.BackColor = RGB(0, 255, 0)
   Me.vbCheckList3.BackColor = RGB(0, 255, 0)
   Me.vbCheckList4.BackColor = RGB(0, 255, 0)
   Me.vbCheckList5.BackColor = RGB(0, 255, 0)
   Me.vbCheckList6.BackColor = RGB(0, 255, 0)
   Me.vbCheckList7.BackColor = RGB(0, 255, 0)
   
   Me.vbCheckList1.Visible = False
   Me.vbCheckList2.Visible = False
   Me.vbCheckList3.Visible = False
   Me.vbCheckList4.Visible = False
   Me.vbCheckList5.Visible = False
   Me.vbCheckList6.Visible = False
   Me.vbCheckList7.Visible = False

   Me.vlCheckList1.Visible = False
   Me.vlCheckList2.Visible = False
   Me.vlCheckList3.Visible = False
   Me.vlCheckList4.Visible = False
   Me.vlCheckList5.Visible = False
   Me.vlCheckList6.Visible = False
   Me.vlCheckList7.Visible = False

' *****************************************************
' **  End of ClearForm3() Subroutine                ***
' *****************************************************
End Sub
Private Sub ClearForm2()
    Me.vbWebInfo1 = ""
    Me.vbWebInfo2 = ""
    Me.vbWebInfo2.BackColor = RGB(255, 255, 255)
    Me.vbWebInfo3 = ""
    Me.vbWebInfo3.BackColor = RGB(255, 255, 255)
End Sub
Private Sub ClearFileNames()
' *****************************************************
' **  Begining of ClearFileNames() Subroutine       ***
' *****************************************************
   Me.vbAcctNoteFileName = ""
   Me.vbAddrListFileName = ""
   Me.vbRestrictedListFileName = ""
' *****************************************************
' **  End of ClearFileNames() Subroutine            ***
' *****************************************************
End Sub

Private Function FormEdits() As Integer
' *****************************************************
' **  Begining of FormEdits() Function              ***
' *****************************************************
Dim iEditErrorCount As Integer, _
    wkSpace As Integer

iEditErrorCount = 0
FormEdits = 0

'check for a Valid Account number
If Trim(Me.vbAcct) = "" Then
    Me.vbAcct.SetFocus
    MsgBox "Account Number is required, Please complete the form"
    iEditErrorCount = 0
    FormEdits = 1
ElseIf Me.vbAcct.Value > 6135 Then
    Me.vbAcct.SetFocus
    MsgBox "Account Number is out of Range, " & _
        "The Account Number can not be greater than 6133."
    iEditErrorCount = 0
    FormEdits = 2
End If

' Let's put some more edits

If Len(Me.vbAcct) > 4 Then
    Me.vbAcct.SetFocus
    MsgBox Me.vbAcct & " is not a valid account number, too many digits. Please edit."
    iEditErrorCount = iEditErrorCount + 1
End If

If Len(Me.vbAcct) < 4 Then
    Me.vbAcct.SetFocus
    MsgBox Me.vbAcct & " is not a valid account number, too few digits. Leading Zeros are probably necessary."
    iEditErrorCount = iEditErrorCount + 1
End If

If iEditErrorCount > 0 Then
   FormEdits = FormEdits + (iEditErrorCount * 100)
End If

' *****************************************************
' **  End of Function FormEdits() Function          ***
' *****************************************************
End Function
Function FunctionalityEdits() As Integer
' *****************************************************
' **  Beginning of Function FunctionalityEdits()    ***
' *****************************************************
Dim wkDebugg As Boolean

If TomsDebugger = TomsDebugger > 2 Then
   wkDebugg = True
Else
   wkDebugg = False
End If

If Me.vbPrintForm = True Then
    vbPrintForm_ShowIt (Me.vbTaxIdCounty)
End If
If Me.vbSendFaxToAdmin = True Then
    vbSendFaxToAdmin_ShowIt (Me.vbTaxIdCounty)
End If

If wkDebugg = True Then
    MsgBox "Webpage Successfully Launched", vbOKOnly + vbInformation, _
       "New Webpage For Account Successfully Launched"
End If

FunctionalityEdits = 0
' *****************************************************
' **  End of Function FunctionalityEdits()          ***
' *****************************************************

End Function

Private Sub butTaxIDFetch_Click()
' *****************************************************
' **  Start of butTaxIDFetch_Click() Subroutine     ***
' *****************************************************
Dim jRow As Long
Dim wsCurrTaxId As String
Dim wkHardSearch As Boolean, _
    wkDebugg As Boolean

wkHardSearch = True
jRow = 2

If TomsDebugger > 8 Then
   wkDebugg = True
Else
   wkDebugg = False
End If

'Set source workbook
Windows(vFile1).Activate

'Clear Unrelated fields
Me.vbName.Value = ""
Me.vbAddress1.Value = ""
Me.vbAddress2.Value = ""
Me.vbWebInfo1.Value = ""
Me.vbWebInfo2.Value = ""
Me.vbWebInfo3.Value = ""
Me.vbNote.Value = ""


retVal = FetchTaxId(Me.vbTaxIdCounty, wb1, ws1)
Me.vbTaxIdCounty.SetFocus
Select Case retVal
        Case 0
        Me.vbTaxIdCountyPrev = Me.vbTaxIdCounty
    Case 9010
        ClearForm
        MsgBox "Error - Tax Id is a required Field ! Try Again."
        Exit Sub
        Case 9001
        ClearForm2
        MsgBox "Error with Connection to CountyTaxRecord Object ! Try Again."
   Exit Sub
        Case 9003
        ClearForm2
        MsgBox "Problem Accessing Tax Account using Website ! Try Again."
        Exit Sub
    Case Else
        ClearForm2
        MsgBox "Error - Invalid return code on funtion ! Call Systems."
        Exit Sub
End Select

Range("A" & jRow).Select

Me.vbWebInfo1 = Cells(1, 1).Value
If Cells(16, 1).Value = "VERY GOOD" Then
    Me.vbWebInfo2 = Cells(12, 1).Value
    Me.vbWebInfo2.BackColor = RGB(0, 255, 0)
    Me.vbWebInfo3 = Cells(13, 1).Value
    Me.vbWebInfo3.BackColor = RGB(0, 255, 0)
    retVal = buildCheckList(6, True)
Else
    Me.vbWebInfo2 = ws1.Range("A12").Value
'    Me.vbWebInfo2.BackColor = &HCE00E0 PURPLE
    Me.vbWebInfo2.BackColor = &HFF
    Me.vbWebInfo3 = ws1.Range("A13").Value
'    Me.vbWebInfo3.BackColor = &HCE00E0 PURPLE
    Me.vbWebInfo3.BackColor = &HFF
    retVal = buildCheckList(6, False)
End If

If Me.vbTaxIdCounty <> "" Then
    FunctionalityEdits
End If

retVal = buildCheckList(7, True)
Me.vbTaxIdCounty.SetFocus
' *****************************************************
' **  End of butTaxIDFetch_Click() Subroutine       ***
' *****************************************************

End Sub

Private Function FetchTaxId(TaxIdCounty As String, _
                            wb As Workbook, _
                            ws As Worksheet) As Long
' *****************************************************
' **         Start of FetchTaxId() Subroutine       ***
' *****************************************************
'check for an Account number
ws.Activate
ws.Range("D26").Select

If Trim(TaxIdCounty) = "" Then
    FetchTaxId = 90010
    Exit Function
End If
' *****************************************************
' **  FETCH DATA FROM COUNTY WEBSITE                ***
' *****************************************************
' *****************************************************
' **  SPECIAL CODE FOR EXCEL 2003 CODE BELOW        ***
' **    NOT EXECUTED IF , NOT NEEDED FOR EXCEL 2003 !!!!
' **   WILL NOT WORK !!!  EXECUTE  FOR 2007 +       ***
' *****************************************************
Select Case wkExcelVersionUsed
Case Is > 2003
    On Error Resume Next
        With wb.Connections("CountyTaxRecord")
            .Name = "CountyTaxRecord"
            .Description = "Link to Brevard County tax assessors office"
        End With
    If Err.Number <> 0 Then
        FetchTaxId = 9001
        Exit Function
    End If
    ws.Activate
    Range("D26").Select
Case Else
    ws.Activate
    Range("D26:D146").Select
End Select

' *****************************************************
' **  SPECIAL CODE FOR EXCEL 2003 END               ***
' *****************************************************
On Error Resume Next
With Selection.QueryTable
        .Connection = "URL;" & wsMasterURL & TaxIdCounty & wsMasterURLSuffix
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
End With
If Err.Number <> 0 Then
   FetchTaxId = 9003
   Exit Function
End If
FetchTaxId = 0
' *****************************************************
' ***         End  of FetchTaxId() Function         ***
' *****************************************************
End Function
Private Function WriteConsolitdatedLog(wb As Workbook, _
                                       ws As Worksheet) As Long
' *****************************************************
' **  Begining of WriteConsolitdatedLog() Function  ***
' *****************************************************
Dim iRow, _
    iRowLastItem, _
    jRow As Long
    
iRowLastItem = 63999
jRow = 2

If vFile5 = Empty Then
    vFile5 = Application.GetOpenFilename("Excel-files,*.xls", _
            1, "Select Correct Log File To Open", , False)
    'if the user didn't select a file, exit sub
    If TypeName(vFile5) = "Boolean" Then
       WriteConsolitdatedLog = 8001
       Exit Function
    End If
    Workbooks.Open Filename:=vFile5, Editable:=True
    'Set targetworkbook
    Set wb5 = ActiveWorkbook
Else
    Workbooks.Open Filename:=vFile5, Editable:=True
    Set wb5 = ActiveWorkbook
End If

Set ws5 = wb5.Worksheets("Sheet1")
'find first empty row in database
 iRowLastItem = ws5.Cells.Find(What:="*", SearchOrder:=xlRows, _
  SearchDirection:=xlPrevious, LookIn:=xlValues).Row + 1

ws5.Range("A" & iRowLastItem & ":H" & iRowLastItem).Value = ws1.Range("D" & jRow & ":K" & jRow).Value
ws5.Range("I" & iRowLastItem).Value = ws1.Range("L" & jRow).Value
ws5.Range("J" & iRowLastItem & ":W" & iRowLastItem).Value = ws1.Range("M" & jRow & ":Z" & jRow).Value
ws5.Range("X" & iRowLastItem).Value = Me.vbWebInfo1
ws5.Range("Y" & iRowLastItem).Value = Me.vbWebInfo2
ws5.Range("Z" & iRowLastItem).Value = Me.vbWebInfo3

wb5.Close SaveChanges:=True

WriteConsolitdatedLog = 0
' *****************************************************
' **     End of WriteConsolitdatedLog() Function    ***
' *****************************************************
End Function

Private Function getCorrectNotesFile(AcctNo As String) As String
' *****************************************************
' **  Start of getCorrectNotesFile() Function       ***
' *****************************************************
Dim convertedAcctNo As Integer

convertedAcctNo = Int(Trim(AcctNo))

Select Case convertedAcctNo
    Case Is < 1000
        getCorrectNotesFile = wsNoteFileHund
    Case Is < 2000
        getCorrectNotesFile = wsNoteFileOneThou
    Case Is < 3000
        getCorrectNotesFile = wsNoteFileTwoThou
    Case Is < 4000
        getCorrectNotesFile = wsNoteFileThreeThou
    Case Is < 5000
        getCorrectNotesFile = wsNoteFileFourThou
    Case Is < 6000
        getCorrectNotesFile = wsNoteFileFiveThou
    Case Is < 7000
        getCorrectNotesFile = wsNoteFileSixThou
    Case Is < 10000
        getCorrectNotesFile = "Bad RANGE in Account NUmbers"
Case Else
        getCorrectNotesFile = "HELP**HELP***HLP**"
End Select
    
' *****************************************************
' **  Start of getCorrectNotesFile() Function       ***
' *****************************************************
End Function

Private Function buildCheckList(CheckListNo As Integer, _
                 PassOrFail As Boolean) As Integer
' *****************************************************
' **  Start of buildCheckList() Function            ***
' *****************************************************

If CheckListNo = 7 Then
   If Me.vbCheckList1.Caption = "FAILED" Or _
      Me.vbCheckList2.Caption = "FAILED" Or _
      Me.vbCheckList3.Caption = "FAILED" Or _
      Me.vbCheckList4.Caption = "FAILED" Or _
      Me.vbCheckList5.Caption = "FAILED" Or _
      Me.vbCheckList6.Caption = "FAILED" Then
      ' Then overWRITE the second parameter to FALSE !!!
      PassOrFail = False
    End If
End If

If PassOrFail = False Then
    Me.vbCheckList7.BackColor = &HE0
    Me.vbCheckList7.Caption = "FAILED"
'    Me.vlCheckList7.BackColor = &HE0
Else
    Me.vbCheckList7.BackColor = RGB(0, 255, 0)
    Me.vbCheckList7.Caption = "PASSED"
'    Me.vlCheckList7.BackColor = RGB(0, 255, 0)
End If

Select Case CheckListNo

    Case 1
        Me.vbCheckList1 = True
        Me.vbCheckList1.BackColor = Me.vbCheckList7.BackColor
        Me.vbCheckList1.Visible = True
        Me.vlCheckList1.Visible = True
        Me.vbCheckList1.Caption = Me.vbCheckList7.Caption
'        Me.vlCheckList1.BackColor = Me.vlCheckList7.BackColor
   
    Case 2
        Me.vbCheckList2 = True
        Me.vbCheckList2.BackColor = Me.vbCheckList7.BackColor
        Me.vbCheckList2.Visible = True
        Me.vlCheckList2.Visible = True
        Me.vbCheckList2.Caption = Me.vbCheckList7.Caption
'        Me.vlCheckList2.BackColor = Me.vlCheckList7.BackColor
   
    Case 3
        Me.vbCheckList3 = True
        Me.vbCheckList3.BackColor = Me.vbCheckList7.BackColor
        Me.vbCheckList3.Visible = True
        Me.vlCheckList3.Visible = True
        Me.vbCheckList3.Caption = Me.vbCheckList7.Caption
'        Me.vlCheckList3.BackColor = Me.vlCheckList7.BackColor
   
    Case 4
        Me.vbCheckList4 = True
        Me.vbCheckList4.BackColor = Me.vbCheckList7.BackColor
        Me.vbCheckList4.Visible = True
        Me.vlCheckList4.Visible = True
        Me.vbCheckList4.Caption = Me.vbCheckList7.Caption
'        Me.vlCheckList4.BackColor = Me.vlCheckList7.BackColor
   
    Case 5
        Me.vbCheckList5 = True
        Me.vbCheckList5.BackColor = Me.vbCheckList7.BackColor
        Me.vbCheckList5.Visible = True
        Me.vlCheckList5.Visible = True
        Me.vbCheckList5.Caption = Me.vbCheckList7.Caption
'        Me.vlCheckList5.BackColor = Me.vlCheckList7.BackColor
        
    Case 6
        Me.vbCheckList6 = True
        Me.vbCheckList6.BackColor = Me.vbCheckList7.BackColor
        Me.vbCheckList6.Visible = True
        Me.vlCheckList6.Visible = True
        Me.vbCheckList6.Caption = Me.vbCheckList7.Caption
'        Me.vlCheckList6.BackColor = Me.vlCheckList7.BackColor
        
    Case 7
        Me.vbCheckList7 = True
        Me.vbCheckList7.BackColor = Me.vbCheckList7.BackColor
        Me.vbCheckList7.Visible = True
        Me.vlCheckList7.Visible = True
        Me.vbCheckList7.Caption = Me.vbCheckList7.Caption
'        Me.vlCheckList7.BackColor = Me.vlCheckList7.BackColor
        
    Case Else:
        buildCheckList = 1
        Exit Function
End Select

buildCheckList = 0
' *****************************************************
' **  End  of buildCheckList() Function            ***
' *****************************************************
End Function

Private Sub FetchURL(TaxIdCounty As String)

' NOTE NOTE ****
'  Just remember to open the Excel Tools, reference and add "Microsoft Internet Controls" DLL references !!!
'      or this script will NOT work !!!
'    ***********************************************************

Dim IE As Object
Dim shellWins As New ShellWindows
Dim IE_TabURL As String
Dim intRowPosition As Integer

intRowPosition = 1

If TaxIdCounty = "" Then
   TaxIdCounty = ""
End If

Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = True

IE.Navigate ws9.Range("A" & intRowPosition) & TaxIdCounty

While IE.Busy
    DoEvents
Wend

intRowPosition = intRowPosition + 1

While ws9.Range("A" & intRowPosition) <> vbNullString
    IE.Navigate ws9.Range("A" & intRowPosition) & TaxIdCounty, CLng(2048)

    While IE.Busy
        DoEvents
    Wend

    intRowPosition = intRowPosition + 1
Wend

Set IE = Nothing
End Sub
Private Sub IE_Automation(TaxIdCounty As String)
    'Dim i As Long
    Dim IE As Object
    'Dim objElement As Object
    'Dim objCollection As Object
 
    ' Create InternetExplorer Object
    Set IE = CreateObject("InternetExplorer.Application")
    
    If TaxIdCounty = "" Then
       TaxIdCounty = ""
    End If
 
    ' You can uncoment Next line To see form results
'   IE.Visible = False
 
    ' Send the form data To URL As POST binary request
    IE.Navigate wsMasterURL & TaxIdCounty & wsMasterURLSuffix
 
    ' Statusbar
    Application.StatusBar = "www.brevard.county-taxes.com is loading for Tax ID Number " & _
                            TaxIdCounty & ". Please wait..."
 
    ' Wait while IE loading...
    '  REMOVED because it is not necessary.
    'Do While IE.Busy
    '    Application.Wait DateAdd("s", 1, Now)
    'Loop
    '
    ' Show IE
    IE.Visible = True
 
    ' Clean up
    Set IE = Nothing
    'Set objElement = Nothing
    'Set objCollection = Nothing
 
    Application.StatusBar = ""
End Sub

Private Sub vbPrintForm_ShowIt(TaxIdCounty As String)
    FetchURL (TaxIdCounty)
End Sub

Private Sub vbSendFaxToAdmin_ShowIt(TaxIdCounty As String)
    IE_Automation (TaxIdCounty)
End Sub

Private Sub UserForm_Activate()
    Me.StartUpPosition = 0
    Me.Top = Application.Top + wkTopWindowPos
    Me.Left = Application.Left + Application.Width - Me.Width - wkLeftWindowPos
End Sub

Private Sub UserForm_Terminate()
If Application.Visible = False Then
    Application.Visible = True
End If
'    Application.Visible = True
'    MsgBox " we are closing the window by the yellow X "
'    MsgBox " Don't forget to save your changes. "
End Sub

Private Sub UserForm_Initialize()
' *****************************************************
' **  Begining of UserForm_Initialize() Subroutine  ***
' *****************************************************
Dim wkGetUserName As String
Me.vbAcctPrev = ""
Me.vbTaxIdCountyPrev = ""
wkGetUserName = GetUserName()
ClearForm
ClearForm3
ClearForm2

Set wb1 = ActiveWorkbook
Set ws1 = wb1.Worksheets("EZBadges")
Set ws9 = wb1.Worksheets("Configuration")

Me.vbInfoBar = ""

vFile1 = ws9.Range("B19").Value
wkCloseFile = ws9.Range("B21")
Windows(vFile1).Activate

' Get nagging parameter defaults for Note files & Restricted List & Resident Address List
'   These variables will produce a popup window after x amount of days if they do not
'   update the note, resident address and restricted list lookup excel files.
wkNoteDays = ws9.Range("B22")
wkRestrictedListDays = ws9.Range("C22")
wkAnnoyLimit = ws9.Range("D22")
wkNagFlag = ws9.Range("E22")

wsMasterURL = Trim(ws9.Range("A1").Value)
wsMasterURLSuffix = Trim(ws9.Range("B1"))

wkTopWindowPos = Int(ws9.Range("B2").Value)
wkLeftWindowPos = Int(ws9.Range("C2").Value)
wkExcelVersionUsed = Int(ws9.Range("B4").Value)

If ws9.Range("C5").Value = "Y" Then
   Me.vbPrintForm = True
ElseIf ws9.Range("C5").Value = "N" Then
   Me.vbPrintForm = False
Else
   Me.vbPrintForm = False
End If

If ws9.Range("C6").Value = "Y" Then
   Me.vbSendFaxToAdmin = True
ElseIf ws9.Range("C6").Value = "N" Then
   Me.vbSendFaxToAdmin = False
Else
   Me.vbSendFaxToAdmin = False
End If

' ONLY display this part of the Screen if it is Me or Sue Cuddie
If wkGetUserName = "Sue Cuddie" Or _
   wkGetUserName = "CalendarCord" Or _
   wkGetUserName = "Calendar" Or _
   wkGetUserName = "Kimi Cheng" Then
   Me.vbDontAskForFiles.Visible = True
Else
   Me.vbDontAskForFiles.Visible = False
End If

If ws9.Range("C7").Value = "Y" Then
   Me.vbDontAskForFiles = True
ElseIf ws9.Range("C7").Value = "N" Then
   Me.vbDontAskForFiles = False
Else
   Me.vbDontAskForFiles = False
End If

If Me.vbDontAskForFiles.Value = True Then
   vFile2 = Trim(ws9.Range("B8").Value)
   vFile3 = Trim(ws9.Range("B9").Value)
   vFile4 = Trim(ws9.Range("B10").Value)
   vFile5 = Trim(ws9.Range("B3").Value)
   Me.vbAcctNoteFileName = vFile2
   Me.vbAddrListFileName = vFile3
   Me.vbRestrictedListFileName = vFile4
   
   wsNoteFileHund = Trim(ws9.Range("B12").Value)
   wsNoteFileOneThou = Trim(ws9.Range("B13").Value)
   wsNoteFileTwoThou = Trim(ws9.Range("B14").Value)
   wsNoteFileThreeThou = Trim(ws9.Range("B15").Value)
   wsNoteFileFourThou = Trim(ws9.Range("B16").Value)
   wsNoteFileFiveThou = Trim(ws9.Range("B17").Value)
   wsNoteFileSixThou = Trim(ws9.Range("B18").Value)

End If

vFile21 = Trim(ws9.Range("B25").Value)
vFile22 = Trim(ws9.Range("B26").Value)
vFile23 = Trim(ws9.Range("B27").Value)
vFile24 = Trim(ws9.Range("B28").Value)
vFile25 = Trim(ws9.Range("B29").Value)
vFile26 = Trim(ws9.Range("B30").Value)
vFile27 = Trim(ws9.Range("B31").Value)
vFile28 = Trim(ws9.Range("B32").Value)
vFile29 = Trim(ws9.Range("B33").Value)

TomsDebugg = Int(ws9.Range("B20").Value)
retVal = buildCheckList(1, True)
' *****************************************************
' **  End of UserForm_Initialize() Subroutine       ***
' *****************************************************
End Sub
Public Function TomsDebugger() As Integer
    TomsDebugger = TomsDebugg
' Level 0 means no debugging at all
End Function

Public Sub ShowNagMsgForDate(ByVal inMsg As Variant, _
                             ByVal inMsg2 As Variant, _
                    Optional ByVal inDate As Date = #1/1/1970#, _
                    Optional showModDateForFileInSeperateMsgBox As Boolean = False)
        MsgBox inMsg & _
                     DateDiff("d", inDate, Date) & _
                     inMsg2 & _
                    vbCrLf & vbCrLf & "Thank You. "
'
        If showModDateForFileInSeperateMsgBox Then
            MsgBox " The modification date for this file is " & _
                     inDate & "."
        End If
'
End Sub


