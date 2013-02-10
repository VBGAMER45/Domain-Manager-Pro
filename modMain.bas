Attribute VB_Name = "modMain"
Option Explicit

Global strWhoisDotCom As String
Global strWhoisDotNet As String
Global strWhoisDotOrg As String
Global strWhoisDotBiz As String
Global strWhoisDotInfo As String
Global ExpireDays As Long 'How many days to check if domain is expired
Global blnGetIPAddress As Boolean

Global blnStartUp As Boolean
Global blnRawWhois As Boolean

Global MruList As New clsMRU

Global IsInTray As Boolean
Public MySysTray As New CSystrayIcon

Public Const DotCom As String = "whois.internic.net"
Public Const DotNet As String = "whois.internic.net"
Public Const DotOrg As String = "whois.opensrs.net"
Public Const DotBiz As String = "whois.internic.net"
Public Const DotInfo As String = "whois.internic.net"
'Custom Menu
Private Type MenuType
    MenuTitle As String
    MenuLink As String
End Type
Global CustomMenu() As MenuType
Private Type WhoisServersType
    Extension As String
    Server As String
End Type
Global WhoisServers() As WhoisServersType
'Windows XP Controls
Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200
'For Reg
Const REG_SZ = 1 ' Unicode nul terminated string
Const REG_BINARY = 3 ' Free form binary
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Sub SaveString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim ret
    'Create a new key
    RegCreateKey hKey, strPath, ret
    'Save a string to the key
    RegSetValueEx ret, strValue, 0, REG_SZ, ByVal strData, Len(strData)
    'close the key
    RegCloseKey ret
End Sub
Sub DelSetting(hKey As Long, strPath As String, strValue As String)
    Dim ret
    'Create a new key
    RegCreateKey hKey, strPath, ret
    'Delete the key's value
    RegDeleteValue ret, strValue
    'close the key
    RegCloseKey ret
End Sub

Public Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function
Sub Main()
    Call InitCommonControlsVB
    frmMain.Show
End Sub
Public Sub RegRun(Path As String, KeyName As String)
   ' Dim Reg As Object
   ' Set Reg = CreateObject("wscript.shell")
   ' Reg.RegWrite "HKEY_LOCAL_MACHINE\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN\" & Keyname, Path
    Call SaveString(HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", KeyName, Path)
    
    
End Sub
Public Sub RemoveRegRun(KeyName As String)
    Call DelSetting(HKEY_LOCAL_MACHINE, "SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\RUN", KeyName)
End Sub
Public Function FileExists(ByVal Path As String) As Boolean
'*****************************
'Purpose: Checks wether a FileExists or not
'*****************************
  If Len(Path) = 0 Then Exit Function
  If Dir(Path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> vbNullString Then FileExists = True
End Function
Public Sub LoadWhoisData()
    On Error GoTo nofile
    Dim f As Long, strData As String
    
    ReDim WhoisServers(0)
    
    f = FreeFile
    Open App.Path & "\whois.ini" For Input As #f
    Do While Not EOF(f)
        Line Input #f, strData
        
        If Left$(LCase(strData), 4) = "tld=" Then
            ReDim Preserve WhoisServers(UBound(WhoisServers) + 1)
            
            WhoisServers(UBound(WhoisServers)).Extension = Right(strData, Len(strData) - 4)
        End If
        If Left$(LCase(strData), 12) = "whoisserver=" Then
           
            WhoisServers(UBound(WhoisServers)).Server = Right(strData, Len(strData) - 12)
        End If
        
    Loop
    Close #f
    Exit Sub
nofile:
    MsgBox "Erorr loading whois data: " & Err.Description

End Sub

Public Function ParseCSVLine(strSource As String, Optional strDelimiter As String = ",", Optional strQuote As String = """") As Variant()
    'Function parses 1 line of text (from a CSV file), returning a variant array
    'containing the column values stored in the CSV.
    'Optionally may pass the delimiter that separates fields (default is comma),
    'and/or the text string quote character (default is ").
    'Note: Strings in the CSV do not have to be contained within quotes as long
    '      as the delimiter character is not part of the string.
    
    'All fields (including strings in quotes, and numbers) will be stored as
    'variant.  Fields will range from array(0) to array(UBound(array)).
    'To assign explicit values to your var's, use the following method:
    '   A = CInt(array(0)) 'Integer
    '   B = CStr(array(1)) 'String
    '  etc.
        
    Dim intTest As Integer
    Dim intCount As Integer, intEnd As Integer
    Dim parseText As String, chunk As String
    Dim varHold() As Variant
    
    'initialize for new array
    intCount = 0: intEnd = 1
    parseText = strSource
    ReDim varHold(0)

    'process fields until no more delimiters found
    Do While intEnd > 0
        
        If Len(parseText) > 0 And Left(LTrim(parseText), 1) = strQuote Then
            '----------------------------
            'Process quoted fields here!
            '----------------------------
            parseText = LTrim(parseText)
            If Len(parseText) > 1 Then
                'Find ending quote
                intEnd = InStr(2, parseText, strQuote)
                If intEnd = 0 Then intEnd = Len(parseText) + 1 '<-last field
                'Extract field value
                If intEnd = 2 Then chunk = "" Else chunk = Mid(parseText, 2, intEnd - 2)
                If intEnd < Len(parseText) Then
                    'Find next delimiter
                    intEnd = InStr(intEnd + 1, parseText, strDelimiter)
                Else
                    'If no delimiter, then end parsing
                    intEnd = 0
                End If
            Else
                'if opening quote is last character, set last field blank & end parsing
                chunk = "": intEnd = 0
            End If
        Else
            '------------------------------
            'Process non-quoted fields here!
            '------------------------------
            If Len(parseText) > 0 Then
                'Find next delimiter
                intEnd = InStr(1, parseText, strDelimiter)
                If intEnd = 0 Then intEnd = Len(parseText) + 1
                'Extract field value
                If intEnd = 1 Then chunk = "" Else chunk = Left(parseText, intEnd - 1)
                If intEnd > Len(parseText) Then intEnd = 0  'detect end of string
            Else
                'If last field is blank, set it and end parsing
                chunk = "": intEnd = 0
            End If
        End If
        
        'Remove current field from parsing string
        If intEnd = Len(parseText) Or intEnd = 0 Then
            parseText = ""
        Else
            parseText = Right(parseText, Len(parseText) - intEnd)
        End If
            
        'increase the array and store new field
        If intCount > 0 Then ReDim Preserve varHold(UBound(varHold) + 1)
        varHold(UBound(varHold)) = CVar(chunk)
        
        intCount = intCount + 1 'increment record count
    Loop
    
    'Assign temp array to function value:
    ParseCSVLine = varHold
    
    'for debugging to the immediate window
    'For intTest = LBound(varHold) To UBound(varHold)
    '    Debug.Print "#" & intTest & ": " & varHold(intTest)
    'Next
End Function


