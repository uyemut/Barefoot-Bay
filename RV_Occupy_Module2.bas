Attribute VB_Name = "Module2"
Option Explicit

' VBA MODULE: Get all IP Addresses of your machine
' (c) 2005 Wayne Phillips (http://www.everythingaccess.com)
' Written 18/05/2005
'
' REQUIREMENTS: Windows 98 or above, Access 97 and above
'
' Please read the full tutorial here:
' http://www.everythingaccess.com/tutorials.asp?ID=Get-all-IP-Addresses-of-your-machine
'
' Please leave the copyright notices in place.
' Thank you.
'
'Option Compare Database

'A couple of API functions we need in order to query the IP addresses in this machine
' Deactivating for compatiblity with Excel 2003 & 2007
'Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'Public Declare PtrSafe Function GetIpAddrTable Lib "Iphlpapi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

' Deactivating for compatiblity with Excel 2003 & 2007
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetIpAddrTable Lib "Iphlpapi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long
Public glGetIPAddress As String

'The structures returned by the API call GetIpAddrTable...
Type IPINFO
    dwAddr As Long         ' IP address
    dwIndex As Long         ' interface index
    dwMask As Long         ' subnet mask
    dwBCastAddr As Long     ' broadcast address
    dwReasmSize As Long    ' assembly size
    Reserved1 As Integer
    Reserved2 As Integer
End Type

Public Function ConvertIPAddressToString(longAddr As Long) As String

    Dim IPBytes(3) As Byte
    Dim lngCount As Long

    'Converts a long IP Address to a string formatted 255.255.255.255
    'Note: Could use inet_ntoa instead

    CopyMemory IPBytes(0), longAddr, 4 ' IP Address is stored in four bytes (255.255.255.255)

    'Convert the 4 byte values to a formatted string
    While lngCount < 4

        ConvertIPAddressToString = ConvertIPAddressToString + _
                                    CStr(IPBytes(lngCount)) + _
                                    IIf(lngCount < 3, ".", "")

        lngCount = lngCount + 1

    Wend

End Function

Public Function GetFirstNonLocalIPAddress()

    Dim Ret As Long, Tel As Long
    Dim bytBuffer() As Byte
    Dim IPTableRow As IPINFO
    Dim lngCount As Long
    Dim lngBufferRequired As Long
    Dim lngStructSize As Long
    Dim lngNumIPAddresses As Long
    Dim strIPAddress As String
    Static glGetIPAddress As String

On Error GoTo ErrorHandler:

    Call GetIpAddrTable(ByVal 0&, lngBufferRequired, 1)

    If lngBufferRequired > 0 Then

        ReDim bytBuffer(0 To lngBufferRequired - 1) As Byte

        If GetIpAddrTable(bytBuffer(0), lngBufferRequired, 1) = 0 Then

            'We've successfully obtained the IP Address details...

            'How big is each structure row?...
            lngStructSize = LenB(IPTableRow)

            'First 4 bytes is a long indicating the number of entries in the table
            CopyMemory lngNumIPAddresses, bytBuffer(0), 4

            While lngCount < lngNumIPAddresses

                'bytBuffer contains the IPINFO structures (after initial 4 byte long)
                CopyMemory IPTableRow, _
                            bytBuffer(4 + (lngCount * lngStructSize)), _
                            lngStructSize

                strIPAddress = ConvertIPAddressToString(IPTableRow.dwAddr)

                If Not ((strIPAddress = "127.0.0.1")) Then

                    If glGetIPAddress = "" Then
                        glGetIPAddress = strIPAddress
                    End If
                    GetFirstNonLocalIPAddress = glGetIPAddress
                    Exit Function

                End If

                lngCount = lngCount + 1

            Wend

        End If

    End If

Exit Function

ErrorHandler:
    MsgBox "An error has occured in GetIPAddresses():" & vbCrLf & vbCrLf & _
            Err.Description & " (" & CStr(Err.Number) & ")"
End Function


Function RDB_Create_PDF(Myvar As Object, FixedFilePathName As String, _
                 OverwriteIfFileExist As Boolean, OpenPDFAfterPublish As Boolean) As String
'*******************************************************************************************
'**                                                                                        *
'**  Code is from Microsoft's MSDN Library Website
'**       http://msdn.microsoft.com/en-us/library/ee834871%28v=office.11%29.aspx
'**                                                                                        *
'*******************************************************************************************
    Dim FileFormatstr As String
    Dim Fname As Variant

    'Test to see if the Microsoft Create/Send add-in is installed.
    If Dir(Environ("commonprogramfiles") & "\Microsoft Shared\OFFICE" _
         & Format(Val(Application.Version), "00") & "\EXP_PDF.DLL") <> "" Then

        If FixedFilePathName = "" Then
            'Open the GetSaveAsFilename dialog to enter a file name for the PDF file.
            FileFormatstr = "PDF Files (*.pdf), *.pdf"
            Fname = Application.GetSaveAsFilename("", filefilter:=FileFormatstr, _
                  Title:="Create PDF")

            'If you cancel this dialog, exit the function.
            If Fname = False Then Exit Function
        Else
            Fname = FixedFilePathName
        End If

        'If OverwriteIfFileExist = False then test to see if the PDF
        'already exists in the folder and exit the function if it does.
        If OverwriteIfFileExist = False Then
            If Dir(Fname) <> "" Then Exit Function
        End If

        'Now export the PDF file.
        On Error Resume Next
        Myvar.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                FileName:=Fname, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False, _
                OpenAfterPublish:=OpenPDFAfterPublish
        On Error GoTo 0

        'If the export is successful, return the file name.
        If Dir(Fname) <> "" Then RDB_Create_PDF = Fname
    End If
End Function

Function RDB_Mail_PDF_Outlook(FileNamePDF As String, StrTo As String, _
                              StrSubject As String, StrBody As String, Send As Boolean)
'*******************************************************************************************
'**                                                                                        *
'**  Code is from Microsoft's MSDN Library Website
'**       http://msdn.microsoft.com/en-us/library/ee834871%28v=office.11%29.aspx
'**
'**  Example of use :RDB_Mail_PDF_Outlook Filename, "ron@debruin.nl", "This is the subject", _
'**         "See the attached PDF file with the last figures" _
'**         & vbNewLine & vbNewLine & "Regards Ron de bruin", False
'** Where:
'**    The first argument is the file name (do not change this).
'**    To whom you want to send the e-mail.
'**    The subject line of the e-mail.
'**    What do you want in the body of the e-mail.
'**    Do you want to display the e-mail (set to False) or send it directly (set to True).
'**                                                                                        *
'*******************************************************************************************
    Dim OutApp As Object
    Dim OutMail As Object

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = StrTo
        .CC = ""
        .BCC = ""
        .Subject = StrSubject
        .Body = StrBody
        .Attachments.Add FileNamePDF
        If Send = True Then
            .Send
        Else
            .Display
        End If
    End With
    On Error GoTo 0

    Set OutMail = Nothing
    Set OutApp = Nothing
End Function

Sub RDB_Workbook_To_PDF()
'*******************************************************************************************
'**                                                                                        *
'**  Code is from Microsoft's MSDN Library Website
'**       http://msdn.microsoft.com/en-us/library/ee834871%28v=office.11%29.aspx
'**                                                                                        *
'*******************************************************************************************
    Dim FileName As String

    'Call the function with the correct arguments.
    FileName = RDB_Create_PDF(ActiveWorkbook, "", True, True)

    'For a fixed file name and to overwrite the file each time you run the macro, use the following statement.
    'RDB_Create_PDF(ActiveWorkbook, "C:\Users\Ron\Test\YourPdfFile.pdf", True, True)

    If FileName <> "" Then
    'Uncomment the following statement if you want to send the PDF by mail.
        'RDB_Mail_PDF_Outlook FileName, "ron@debruin.nl", "This is the subject", _
           "See the attached PDF file with the last figures" _
          & vbNewLine & vbNewLine & "Regards Ron de bruin", False
    Else
        MsgBox "It is not possible to create the PDF; possible reasons:" & vbNewLine & _
               "Microsoft Add-in is not installed" & vbNewLine & _
               "You canceled the GetSaveAsFilename dialog" & vbNewLine & _
               "The path to save the file in arg 2 is not correct" & vbNewLine & _
               "You didn't want to overwrite the existing PDF if it exists."
    End If
End Sub


Sub RDB_Worksheet_Or_Worksheets_To_PDF()
'*******************************************************************************************
'**                                                                                        *
'**  Code is from Microsoft's MSDN Library Website
'**       http://msdn.microsoft.com/en-us/library/ee834871%28v=office.11%29.aspx
'**                                                                                        *
'*******************************************************************************************
    Dim FileName As String

    If ActiveWindow.SelectedSheets.Count > 1 Then
        MsgBox "There is more than one sheet selected," & vbNewLine & _
               "and every selected sheet will be published."
    End If

    'Call the function with the correct arguments.
    'You can also use Sheets("Sheet3") instead of ActiveSheet in the code(the sheet does not need to be active then).
    FileName = RDB_Create_PDF(ActiveSheet, "", True, True)

    'For a fixed file name and to overwrite it each time you run the macro, use the following statement.
    'RDB_Create_PDF(ActiveSheet, "C:\Users\Ron\Test\YourPdfFile.pdf", True, True)

    If FileName <> "" Then
        'Uncomment the following statement if you want to send the PDF by e-mail.
        'RDB_Mail_PDF_Outlook FileName, "ron@debruin.nl", "This is the subject", _
           "See the attached PDF file with the last figures" _
          & vbNewLine & vbNewLine & "Regards Ron de bruin", False
    Else
        MsgBox "It is not possible to create the PDF; possible reasons:" & vbNewLine & _
               "Add-in is not installed" & vbNewLine & _
               "You canceled the GetSaveAsFilename dialog" & vbNewLine & _
               "The path to save the file is not correct" & vbNewLine & _
               "PDF file exists and you canceled overwriting it."
    End If
End Sub

Sub RDB_Selection_Range_To_PDF()
'*******************************************************************************************
'**                                                                                        *
'**  Code is from Microsoft's MSDN Library Website
'**       http://msdn.microsoft.com/en-us/library/ee834871%28v=office.11%29.aspx
'**                                                                                        *
'*******************************************************************************************
    Dim FileName As String

    If ActiveWindow.SelectedSheets.Count > 1 Then
        MsgBox "There is more than one sheet selected," & vbNewLine & _
               "unselect the sheets and try the macro again."
    Else
        'Call the function with the correct arguments.

        'For a fixed range use this line.
        FileName = RDB_Create_PDF(Range("A10:I15"), "", True, True)

        'For the selection use this line.
        'FileName = RDB_Create_PDF(Selection, "", True, True)

        'For a fixed file name and to overwrite it each time you run the macro, use the following statement.
        'RDB_Create_PDF(Selection, "C:\Users\Ron\Test\YourPdfFile.pdf", True, True)

        If FileName <> "" Then
            'Uncomment the following statement if you want to send the PDF by mail.
           'RDB_Mail_PDF_Outlook FileName, "ron@debruin.nl", "This is the subject", _
              "See the attached PDF file with the last figures" _
             & vbNewLine & vbNewLine & "Regards Ron de bruin", False
        Else
            MsgBox "It is not possible to create the PDF;, possible reasons:" & vbNewLine & _
                   "Microsoft Add-in is not installed" & vbNewLine & _
                   "You canceled the GetSaveAsFilename dialog" & vbNewLine & _
                   "The path to save the file in arg 2 is not correct" & vbNewLine & _
                   "You didn't want to overwrite the existing PDF if it exists."
        End If
    End If
End Sub

Sub Mail_Every_Worksheet_With_Address_In_A1_PDF()
'This example works in Excel 2007 and Excel 2010.
'*******************************************************************************************
'**                                                                                        *
'**  Code is from Microsoft's MSDN Library Website
'**       http://msdn.microsoft.com/en-us/library/ee834871%28v=office.11%29.aspx
'**                                                                                        *
'*******************************************************************************************
    Dim sh As Worksheet
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim FileName As String

    'Set a temporary path to save the PDF files.
    'You can also use another folder similar to
    'TempFilePath = "C:\Users\Ron\MyFolder\"
    TempFilePath = Environ$("temp") & "\"

    'Loop through each worksheet.
    For Each sh In ThisWorkbook.Worksheets
        FileName = ""

        'Test A1 for an e-mail address.
        If sh.Range("A1").Value Like "?*@?*.?*" Then

            'If there is an e-mail address in A1, create the file name and the PDF.
            TempFileName = TempFilePath & "Sheet " & sh.Name & " of " _
                         & ThisWorkbook.Name & " " _
                         & Format(Now, "dd-mmm-yy h-mm-ss") & ".pdf"

            FileName = RDB_Create_PDF(sh, TempFileName, True, False)


            'If publishing is set, create the mail.
            If FileName <> "" Then
                RDB_Mail_PDF_Outlook FileName, sh.Range("A1").Value, "This is the subject", _
                   "See the attached PDF file with the last figures" _
                   & vbNewLine & vbNewLine & "Regards Ron de bruin", False

                'After the e-mail is created, delete the PDF file in TempFilePath.
                If Dir(TempFileName) <> "" Then Kill TempFileName

                Else
                   MsgBox "Not possible to create the PDF, possible reasons:" & vbNewLine & _
                   "Microsoft Add-in is not installed" & vbNewLine & _
                       "The path to save the file in arg 2 is not correct" & vbNewLine & _
                       "You didn't want to overwrite the existing PDF if it exist"
                End If

            End If
    Next sh
End Sub


