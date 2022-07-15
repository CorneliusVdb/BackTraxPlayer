Attribute VB_Name = "modReadINI"
Option Explicit
'============================================================================================================
'                                                  C H A N G E S                                            '
'============================================================================================================
'  Please make comments HERE to indicate the changes that you did.                                          '
'  Format :                                                                                                 '
'-----------------------------------------------------------------------------------------------------------'
'  User     :  Cornelius                                                                                    '
'  Date     :  17/04/2001                                                                                   '
'  Procedure:  LoadBankCDVInfo                                                                              '
'  Descrip  :  Short description of change                                                                  '
'-----------------------------------------------------------------------------------------------------------'
'  User     :                                                                                               '
'  Date     :                                                                                               '
'  Procedure:                                                                                               '
'  Descrip  :                                                                                               '
'-----------------------------------------------------------------------------------------------------------'
'  User     :                                                                                               '
'  Date     :                                                                                               '
'  Procedure:                                                                                               '
'  Descrip  :                                                                                               '
'-----------------------------------------------------------------------------------------------------------'
'  User     :                                                                                               '
'  Date     :                                                                                               '
'  Procedure:                                                                                               '
'  Descrip  :                                                                                               '
'-----------------------------------------------------------------------------------------------------------'

'Declarations
Public X                      As Double
Public GVersion               As String
'=====================================================================
'  These values can be found in the VERSION.INI file on the server   '
'  To customize for each project, check the VERSION.INI file and     '
'  make the appropriate entries there-in.                            '
'=====================================================================
Public strAppPath             As String
Public strDistrib             As String
Public strSetupPath           As String
Public strVersionPath         As String
'The path to the VERSION.INI
Public strVersionFile         As String ' = "\\ASSUPRIMARY\AsiaInst\Version.ini"

'Declare the functions to read ini files
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long


Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

' Registry manipulation API's (32-bit)
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003
Private Const ERROR_SUCCESS = 0&
Private Const ERROR_NO_MORE_ITEMS = 259&

Private Const REG_SZ = 1
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const HKeyVersion = "SOFTWARE\ISS\VersionCount"
'==============================================================================
' Add the following line in the startup procedure (ie. Form_load or Sub Main) '
'==============================================================================
'  CheckVersion                                                               '
'==============================================================================

Public Sub CheckVersion()
On Error GoTo err1
Dim sTMP As String
'===============================================================
'               Routine to do autamated setup                  '
'===============================================================
'                                                              '
'===============================================================
'  Get all the default paths, version and names                '
'  from the Version.ini                                        '
'===============================================================
strVersionFile = App.FileDescription
GVersion = ReadIni(UCase(App.EXEName), "Version")
strAppPath = App.Path
strDistrib = ReadIni(UCase(App.EXEName), "Distribute")
strSetupPath = ReadIni(UCase(App.EXEName), "SetupPath")
strVersionPath = ReadIni(UCase(App.EXEName), "VersionPath")

'===============================================================
'                  Check versions are different                '
'===============================================================
'  If not the correct VERSION, start the DISTRIBUTION routine  '
'===============================================================
If GVersion <> App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "0000") Then
   MsgBox "Version " & GVersion & " will now be implemented !!!       ", vbExclamation, "New Version"
   '===============================================================
   '                       Start the copy program                 '
   '===============================================================
   '  This routine will copy the distribute program from the      '
   '  server. After this filecopy, it will SHELL (START) the      '
   '  distribute program to do the filecopy of the latest version '
   '  of the EXE (also from the server) to my App.path.           '
   '  It will end this program.                                   '
   '===============================================================
''''   Dim Ret
''''   Dim sString
''''   Ret = GetString(HKEY_LOCAL_MACHINE, HKeyVersion, "Version")
''''   If Ret = "" Then
''''        GoTo LoadDistibute
''''   Else
''''        If Ret <> sTMP Then
''''            GoTo LoadDistibute
''''        Else
''''          Exit Sub
''''        End If
''''    End If
''''Exit Sub
''''LoadDistibute:
   FileCopy strVersionPath & strDistrib, strAppPath & "\" & strDistrib
   '===============================================================
   '        EXCECUTE the copy-program on my LOCAL machine         '
   '  This process is a-synchronic, and will continue with the    '
   '  code AFTER the Shell-function. That means we can terminate  '
   '  this program now.                                           '
   '===============================================================
   X = Shell(strAppPath & "\" & strDistrib, vbNormalFocus)
   '===============================================================
   '                    Ends the current program                  '
   '===============================================================
   End
End If
Exit Sub
err1:
MsgBox "Error : " & Err.Number & Chr(13) & Chr(13) & "Description : " & Err.Description & " ...                " & Chr(13) & "Source : " & Err.Source & Chr(13) & Chr(13) & "                                    ", vbCritical, "Program ERROR"
End
End Sub

Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim ret
    RegOpenKey hKey, strPath, ret
    GetString = RegQueryStringValue(ret, strValue)
    RegCloseKey ret
End Function

Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
    Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
    lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, Chr$(0))
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strData As Integer
            lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
            If lResult = 0 Then
                RegQueryStringValue = strData
            End If
        End If
    End If
End Function

Private Sub SaveStringLong(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim ret As Long
    Dim stringbuffer As String
    
    RegCreateKey hKey, strPath, ret
    stringbuffer = strData & vbNullChar
    RegSetValueEx ret, strValue, 0, REG_SZ, ByVal stringbuffer, Len(stringbuffer)
    RegCloseKey ret
    
End Sub

Public Function ReadIni(sApp As String, sKey As String, Optional SFile As Variant) As String
'=====================================================================
'  This routine reads the INI file caled VERSION.INI                 '
'  the first command (GetPrivateProfileInt) gets the key             '
'  to loop thru (in this example only)                               '
'  the second command (GetPrivateProfileString) gets the             '
'  actual antries from the INI after the '='                         '
'=====================================================================
'SYNTAX and MEANING:                                                 '
'=====================================================================
'GetPrivateProfileString(                                            '
'                        AppName,      :  Entry identifier in INI    '
'                        Key,          :  Key to look for in INI     '
'                        "",           :  Default ??                 '
'                        Return Value  :  Result found in INI        '
'                        100,          :  Lenght of string to read   '
'                        FileName      :  Path and Name of INI       '
'                        )                                           '
'=====================================================================
Dim sRetStr As String * 255
Dim iRetLen As Integer
Dim iPos1 As Long

On Error Resume Next
If Not IsMissing(SFile) Then
   strVersionFile = SFile
Else
   strVersionFile = App.Path & "\Info.ini"
End If

'Read the appropriate ini entry
iRetLen = GetPrivateProfileString(sApp, sKey, "", sRetStr, 255, strVersionFile)
If iRetLen <> 0 Then
   iPos1 = InStr(1, sRetStr, Chr(0))         'Kry posisie van einde van text
   ReadIni = Mid(sRetStr, 1, iPos1 - 1)      'Set Function = to return value
End If

End Function

