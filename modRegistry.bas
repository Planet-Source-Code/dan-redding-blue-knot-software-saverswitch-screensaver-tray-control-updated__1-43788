Attribute VB_Name = "modRegistry"
Option Explicit
 
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByRef lpData As Long, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Long, ByVal cbData As Long) As Long
Declare Function RegSetValueExB Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Byte, ByVal cbData As Long) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Const ERROR_SUCCESS = 0&
Const ERROR_BADDB = 1009&
Const ERROR_BADKEY = 1010&
Const ERROR_CANTOPEN = 1011&
Const ERROR_CANTREAD = 1012&
Const ERROR_CANTWRITE = 1013&
Const ERROR_OUTOFMEMORY = 14&
Const ERROR_INVALID_PARAMETER = 87&
Const ERROR_ACCESS_DENIED = 5&
Const ERROR_NO_MORE_ITEMS = 259&
Const ERROR_MORE_DATA = 234&

Const REG_NONE = 0&
Const REG_SZ = 1&
Const REG_EXPAND_SZ = 2&
Const REG_BINARY = 3&
Const REG_DWORD = 4&
Const REG_DWORD_LITTLE_ENDIAN = 4&
Const REG_DWORD_BIG_ENDIAN = 5&
Const REG_LINK = 6&
Const REG_MULTI_SZ = 7&
Const REG_RESOURCE_LIST = 8&
Const REG_FULL_RESOURCE_DESCRIPTOR = 9&
Const REG_RESOURCE_REQUIREMENTS_LIST = 10&

Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ

Function SetLongRegValue(lKey As Long, SubKey As String, Entry As String, Value As Long) As Boolean
Dim lReturn As Long, hKey As Long
    lReturn = RegOpenKeyEx(lKey, SubKey, 0, KEY_WRITE, hKey) 'open the key
    If lReturn = ERROR_SUCCESS Then
        lReturn = RegSetValueExA(hKey, Entry, 0, REG_DWORD, Value, 4)
        SetLongRegValue = (lReturn = ERROR_SUCCESS)
        lReturn = RegCloseKey(hKey) 'close the key
    Else 'if there was an error opening the key
        SetLongRegValue = False
    End If
End Function

Function GetLongRegValue(lKey As Long, SubKey As String, Entry As String, Optional lDefault As Long = 0) As Long
Dim lReturn As Long, hKey As Long, lBuffer As Long
    lReturn = RegOpenKeyEx(lKey, SubKey, 0, KEY_READ, hKey) 'open the key
    If lReturn = ERROR_SUCCESS Then 'if the key could be opened then
        lReturn = RegQueryValueExA(hKey, Entry, 0, REG_DWORD, lBuffer, 4) 'get the value from the registry
        If lReturn = ERROR_SUCCESS Then 'if the value could be retreived then
            lReturn = RegCloseKey(hKey) 'close the key
            GetLongRegValue = lBuffer  'return the value
        Else                        'otherwise, if the value couldnt be retreived
            GetLongRegValue = lDefault 'return Error to the user
        End If
    Else 'otherwise, if the key couldnt be opened
        GetLongRegValue = lDefault 'return Error to the user
    End If
End Function

Function SetBinaryRegValue(lKey As Long, SubKey As String, Entry As String, Value As String) As Boolean
Dim lReturn As Long, ByteArray() As Byte, lDataSize As Long, iLoop As Integer, hKey As Long
   lReturn = RegOpenKeyEx(lKey, SubKey, 0, KEY_WRITE, hKey) 'open the key
   If lReturn = ERROR_SUCCESS Then 'if the key was open successfully then
      lDataSize = Len(Value)
      ReDim ByteArray(lDataSize)
      For iLoop = 1 To lDataSize
          ByteArray(iLoop) = Asc(Mid$(Value, iLoop, 1))
      Next iLoop
      lReturn = RegSetValueExB(hKey, Entry, 0, REG_BINARY, ByteArray(1), lDataSize) 'write the value
      SetBinaryRegValue = (lReturn = ERROR_SUCCESS)
      lReturn = RegCloseKey(hKey) 'close the key
   Else 'if there was an error opening the key
      SetBinaryRegValue = False
   End If
End Function

Function GetBinaryRegValue(lKey As Long, SubKey As String, Entry As String, Optional sDefault As String) As String
Dim lReturn As Long, lBufferSize As Long, sBuffer As String, hKey As Long
    lReturn = RegOpenKeyEx(lKey, SubKey, 0, KEY_READ, hKey) 'open the key
    If lReturn = ERROR_SUCCESS Then 'if the key could be opened
        lBufferSize = 1
        lReturn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, 0, lBufferSize)  'get the value from the registry
        sBuffer = Space(lBufferSize)
        lReturn = RegQueryValueEx(hKey, Entry, 0, REG_BINARY, sBuffer, lBufferSize) 'get the value from the registry
        If lReturn = ERROR_SUCCESS Then 'if the value could be retreived then
        lReturn = RegCloseKey(hKey)  'close the key
            GetBinaryRegValue = sBuffer 'return the value to the user
        Else                        'otherwise, if the value couldnt be retreived
            GetBinaryRegValue = sDefault 'return Error to the user
        End If
    Else 'otherwise, if the key couldnt be opened
        GetBinaryRegValue = sDefault 'return Error to the user
    End If
End Function

Function DeleteRegKey(lKey As Long, KeyName As String)
Dim lReturn As Long, hKey As Long
    lReturn = RegOpenKeyEx(lKey, KeyName, 0, KEY_WRITE, hKey) 'open the key
    If lReturn = ERROR_SUCCESS Then 'if the key could be opened then
        lReturn = RegDeleteKey(hKey, KeyName) 'delete the key
        lReturn = RegCloseKey(hKey)  'close the key
    End If
End Function

Function DeleteRegValue(lKey As Long, SubKey As String, KeyName As String)
Dim lReturn As Long, hKey As Long
    lReturn = RegOpenKeyEx(lKey, SubKey, 0, KEY_WRITE, hKey) 'open the key
    If lReturn = ERROR_SUCCESS Then 'if the key could be opened then
        lReturn = RegDeleteValue(hKey, KeyName) 'delete the key
        lReturn = RegCloseKey(hKey)  'close the key
    End If
End Function

'Function ErrorMsg(lErrorCode As Long) As String
'    Dim GetErrorMsg
''If an error does accurr, and the user wants error messages displayed, then
''display one of the following error messages
'
'Select Case lErrorCode
'       Case 1009, 1015
'            GetErrorMsg = "The Registry Database is corrupt!"
'       Case 2, 1010
'            GetErrorMsg = "Bad Key Name"
'       Case 1011
'            GetErrorMsg = "Can't Open Key"
'       Case 4, 1012
'            GetErrorMsg = "Can't Read Key"
'       Case 5
'            GetErrorMsg = "Access to this key is denied"
'       Case 1013
'            GetErrorMsg = "Can't Write Key"
'       Case 8, 14
'            GetErrorMsg = "Out of memory"
'       Case 87
'            GetErrorMsg = "Invalid Parameter"
'       Case 234
'            GetErrorMsg = "There is more data than the buffer has been allocated to hold."
'       Case Else
'            GetErrorMsg = "Undefined Error Code:  " & Str$(lErrorCode)
'End Select
'
'End Function

Function GetStringRegValue(lKey As Long, SubKey As String, Entry As String, Optional strDefault As String = "") As String
Dim lReturn As Long, hKey As Long, strBuffer As String, lBufferSize As Long
    lReturn = RegOpenKeyEx(lKey, SubKey, 0, KEY_READ, hKey) 'open the key
    If lReturn = ERROR_SUCCESS Then 'if the key could be opened then
        strBuffer = Space$(512)     'make a buffer
        lBufferSize = Len(strBuffer)
        lReturn = RegQueryValueEx(hKey, Entry, 0, REG_SZ, strBuffer, lBufferSize) 'get the value from the registry
        If lReturn = ERROR_SUCCESS Then 'if the value could be retreived then
            lReturn = RegCloseKey(hKey)  'close the key
            strBuffer = Trim$(strBuffer)
            If lBufferSize > 1 Then
                GetStringRegValue = Left$(strBuffer, lBufferSize - 1) 'return the value to the user
            Else
                GetStringRegValue = ""
            End If
        Else 'otherwise, if the value couldnt be retreived
            GetStringRegValue = strDefault
        End If
    Else 'otherwise, if the key couldnt be opened
        GetStringRegValue = strDefault
    End If
End Function

Function CreateRegKey(lKey As Long, SubKey As String)
Dim lReturn As Long, hKey As Long
    lReturn = RegCreateKey(lKey, SubKey, hKey) 'create the key
    If lReturn = ERROR_SUCCESS Then 'if the key was created then
        lReturn = RegCloseKey(hKey) 'close the key
    End If
End Function

Function SetStringRegValue(lKey As Long, SubKey As String, Entry As String, Value As String) As Boolean
Dim lReturn As Long, hKey As Long
    lReturn = RegOpenKeyEx(lKey, SubKey, 0, KEY_WRITE, hKey) 'open the key
    If lReturn = ERROR_SUCCESS Then 'if the key was open successfully then
        lReturn = RegSetValueEx(hKey, Entry, 0, REG_SZ, ByVal Value, Len(Value)) 'write the value
        If lReturn = ERROR_SUCCESS Then   'if there was an error writting the value
            SetStringRegValue = True
        Else
            SetStringRegValue = False
        End If
        lReturn = RegCloseKey(hKey) 'close the key
    Else 'if there was an error opening the key
        SetStringRegValue = False
    End If
End Function
