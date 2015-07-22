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

                    GetFirstNonLocalIPAddress = strIPAddress
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
