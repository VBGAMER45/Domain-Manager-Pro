Attribute VB_Name = "modEncode"
' ***********************************************************************
'
' CLASS : clsCoder.cls
'
' PURPOSE : Provide access to the URL Coding / Decoding routines
'
' WRITTEN BY : Alon Hirsch
'
' COMPANY : Debtpack (Pty) Ltd. - Development
'
' DATE : 11 February 2002
'
' ***********************************************************************
Option Explicit
DefInt A-Z

' characters allowed in a URL without needing to be encoded
Private Const URLValid = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

Public Function sURLEncode(ByVal sWork As String) As String
    ' This function will URLEncode sWork and return it as the value of the function
    Dim iLoop As Integer
    Dim iLen As Integer
    Dim sRet As String
    Dim sTemp As String
    
    ' prepare the result string
    sRet = ""

    ' check if we have a string to work with
    If Len(sWork) > 0 Then
        ' we do - determine the length of the string
        iLen = Len(sWork)
        ' check all the characters (one by one)
        For iLoop = 1 To iLen
            ' check each character in turn
            ' get the next character
            sTemp = Mid$(sWork, iLoop, 1)
            ' is the character a valid one or not
            If InStr(1, URLValid, sTemp, vbBinaryCompare) = 0 Then
                If sTemp = Chr$(32) Then
                    ' convert space to +
                    sTemp = "+"
                Else
                    ' not valid - use HEX representation of it
                    sTemp = "%" & Right$("0" & Hex(Asc(sTemp)), 2)
                End If
            End If
            ' add this to the returned string
            sRet = sRet & sTemp
        Next iLoop
        ' return the final result
        sURLEncode = sRet
    End If
End Function
Public Function sURLDecodeB(ByVal sWork As String) As String
    ' This function will scan through the entire sWork and replace all valid
    ' URL Encoded character with their ASCII character value and then return the
    Dim sTemp As String
    Dim sChar As String
    Dim sNewString As String
    Dim lPos1 As Long
    Dim lLen As Long
    Dim lChar As Long
    
    ' prepare the result string
    sNewString = ""
    
    ' determine the lengh of the data to process
    lLen = Len(sWork)
    
    ' loop through each character (NOT BYTE)
    For lChar = 1 To lLen
        ' retrieve the character
        sChar = Mid$(sWork, lChar, 1)
        ' now examine the character
        If sChar = "%" Then
            ' encoded character - decode the next 2 characters
            sTemp = Mid$(sWork, lChar + 1, 2)
            sNewString = sNewString & ChrB$("&H" & sTemp)
            ' increment counter to skip the encoded value
            lChar = lChar + 2
        ElseIf sChar = "+" Then
            ' is a space - decode it
            sNewString = sNewString & ChrB$(32)
        Else
            ' not decoded - use it as is
            sNewString = sNewString & ChrB$(AscB(sChar))
        End If
    Next lChar
    
    ' return the new string to the calling process
    sURLDecodeB = sNewString
End Function
Public Function sURLDecode(ByVal sWork As String) As String
    ' This function will scan through the entire sWork and replace all valid
    ' URL Encoded character with their ASCII character value
    Dim sTemp As String
    Dim sChar As String
    Dim lPos1 As Long
    Dim lPos2 As Long
    Dim lChar As Long
    Dim bFirst As Boolean
    
    ' start with an empty string
    sTemp = ""
    lPos2 = 1
    bFirst = True
    
    ' start by replacing all + with spaces
    sWork = Replace(sWork, "+", Chr$(32))
    
    ' *** now handle the actuall encoded stuff
    ' find the first occurrence
    lPos1 = InStr(1, sWork, "%", vbTextCompare)
    If lPos1 = 0 Then
        ' none found - return the entire string
        sTemp = sWork
    Else
        ' check as long as there are still encoeded characters.
        Do While lPos1 <> 0
            ' find the first %
            ' check if we found one or not
            If lPos1 <> 0 Then
                ' we found 1 - decode it and add it to the result
                If bFirst Then
                    ' this is the first time in - stemp is all data up to the first %
                    sTemp = Left$(sWork, lPos1 - 1)
                    bFirst = False
                Else
                    ' add all the data from the last position to the current position
                    sTemp = sTemp & Mid$(sWork, lPos2 + 2, (lPos1 - lPos2 - 2))
                End If
                sChar = Mid$(sWork, lPos1 + 1, 2)
                lChar = CLng("&H" & sChar)
                sTemp = sTemp & Chr$(lChar)
                ' start at the next position
                lPos2 = lPos1 + 1
            End If
            
            ' check for the next one
            lPos1 = InStr(lPos2, sWork, "%", vbTextCompare)
            If lPos1 = 0 Then
                ' no more - add the rest of the string to be checked
                sTemp = sTemp & Mid$(sWork, lPos2 + 2)
            End If
        Loop
    End If
    ' return the string we have decoded
    sURLDecode = sTemp
End Function
Public Function sURLEncodeB(ByVal sWork As String) As String
    ' This function will URLEncodeB sWork and return it as the value of the function
    ' This performs a BYTE-WISE encoding
    Dim iLoop As Integer
    Dim iLen As Integer
    Dim iTemp As Integer
    Dim sRet As String
    Dim sTemp As String
    Dim bTemp As Byte
    
    ' prepare the result string
    sRet = ""

    ' check if we have a string to work with
    If LenB(sWork) > 0 Then
        ' we do - determine the length of the string
        iLen = LenB(sWork)
        ' check all the characters (one by one)
        For iLoop = 1 To iLen
            ' check each character in turn
            ' get the next character
            iTemp = AscB(MidB$(sWork, iLoop, 1))
            ' is the character a valid one or not
            If (iTemp < 65 Or iTemp > 90) And (iTemp < 97 Or iTemp > 122) Then
                'If sTemp = Chr$(32) Then
                If iTemp = 32 Then
                    ' convert space to +
                    sTemp = "+"
                Else
                    ' not valid - use HEX representation of it
                    sTemp = "%" & Right$("0" & Hex(iTemp), 2)
                End If
            Else
                sTemp = Chr$(iTemp)
            End If
            ' add this to the returned string
            sRet = sRet & sTemp
        Next iLoop
        ' return the final result
        sURLEncodeB = sRet
    End If
End Function

