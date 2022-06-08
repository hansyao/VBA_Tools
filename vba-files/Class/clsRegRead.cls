VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRegRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ERROR_SUCCESS As Long = 0&
Private Const ERROR_MORE_DATA As Long = 234

Private Const KEY_READ As Long = &H20019
Private Const KEY_WOW64_64KEY As Long = &H100&
Private Const KEY_WOW64_32KEY As Long = &H200&

Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_DWORD_BIG_ENDIAN = 5
Private Const REG_MULTI_SZ = 7
Private Const REG_SZ = 1

Public Enum RegReadErrorConstants
    rrClosedError = vbObjectError Or &H9902&
    rrSystemError = vbObjectError Or &H9900&
    rrValTypeError = vbObjectError Or &H9904&
End Enum
#If False Then 'Preserve case of constants.
Private rrClosedError, rrSystemError, rrValTypeError
#End If

Public Enum RegReadKeyConstants
    HKEY_CLASSES_ROOT = &H80000000
    HKCR = HKEY_CLASSES_ROOT
    HKEY_CURRENT_USER = &H80000001
    HKCU = HKEY_CURRENT_USER
    HKEY_LOCAL_MACHINE = &H80000002
    HKLM = HKEY_LOCAL_MACHINE
    HKEY_USERS = &H80000003
    HKU = HKEY_USERS
End Enum
#If False Then 'Preserve case of constants.
Private HKEY_CLASSES_ROOT, HKCR, HKEY_CURRENT_USER
Private HKCU, HKEY_LOCAL_MACHINE, HKLM, HKEY_USERS, HKU
#End If

Public Enum RegReadWOW64Constants
    KEY64 = KEY_WOW64_64KEY
    KEY32 = KEY_WOW64_32KEY
End Enum
#If False Then 'Preserve case of constants.
Private KEY64, KEY32
#End If

#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
    ByVal Destination As LongPtr, _
    ByVal Source As LongPtr, _
    ByVal Length As Long)

Private Declare PtrSafe Function RegCloseKey Lib "advapi32" (ByVal hKey As LongPtr) As Long

Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32" _
    Alias "RegOpenKeyExW" ( _
    ByVal hKey As LongPtr, _
    ByVal lpSubKey As LongPtr, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    ByRef phkResult As LongPtr) As Long

Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32" _
    Alias "RegQueryValueExW" ( _
    ByVal hKey As LongPtr, _
    ByVal lpValueName As LongPtr, _
    ByVal lpReserved As LongPtr, _
    ByRef lpType As Long, _
    ByVal lpData As LongPtr, _
    ByRef lpcbData As Long) As Long

#Else
Private Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" ( _
    ByVal Destination As Long, _
    ByVal Source As Long, _
    ByVal Length As Long)

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32" _
    Alias "RegOpenKeyExW" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As Long, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    ByRef phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32" _
    Alias "RegQueryValueExW" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As Long, _
    ByVal lpReserved As Long, _
    ByRef lpType As Long, _
    ByVal lpData As Long, _
    ByRef lpcbData As Long) As Long

#End If

#If VBA7 Then
Private mhOpenKey As LongPtr
#Else
Private mhOpenKey As Long
#End If
Private mlngType As Long
Private mlngResult As Long


Public Property Get SystemError() As Long
    SystemError = mlngResult
End Property

Public Property Get ValueType() As Long
    ValueType = mlngType
End Property

Public Sub CloseKey()
    If mhOpenKey <> 0 Then
        mlngResult = RegCloseKey(mhOpenKey)
        If mlngResult <> ERROR_SUCCESS Then
            err.Raise rrSystemError, _
                      "RegRead.CloseKey", _
                      "System error " & CStr(mlngResult)
        Else
            mhOpenKey = 0
        End If
    End If
End Sub

Public Sub OpenKey( _
    ByVal hKey As RegReadKeyConstants, _
    ByVal SubKey As String, _
    Optional ByVal WOW64 As RegReadWOW64Constants = KEY32)

    mlngResult = RegOpenKeyEx(hKey, _
                             StrPtr(SubKey), _
                             0, _
                             KEY_READ Or WOW64, _
                             mhOpenKey)
    If mlngResult <> ERROR_SUCCESS Then
        err.Raise rrSystemError, _
                  "RegRead.OpenKey", _
                  "System error " & CStr(mlngResult) & vbNewLine & SubKey
    End If
End Sub

Public Function QueryValue(ByVal ValueName As String) As Variant
    Dim lngDataSize As Long
    Dim bytBuffer() As Byte
    Dim strBuffer As String
    Dim bytTemp As Byte
    Dim lngDWORD As Long

    If mhOpenKey = 0 Then
        err.Raise rrClosedError, _
                  "RegRead.QueryValue", _
                  "Must open registry key before querying values"
    End If
    mlngResult = RegQueryValueEx(mhOpenKey, _
                                StrPtr(ValueName), _
                                0, _
                                mlngType, _
                                ByVal 0&, _
                                lngDataSize)
    If mlngResult <> ERROR_SUCCESS Then
        err.Raise rrSystemError, _
                  "RegRead.QueryValue", _
                  "System error " & CStr(mlngResult)
    End If
    ReDim bytBuffer(lngDataSize - 1)
    mlngResult = RegQueryValueEx(mhOpenKey, _
                                StrPtr(ValueName), _
                                0, _
                                mlngType, _
                                VarPtr(bytBuffer(0)), _
                                lngDataSize)
    If mlngResult <> ERROR_SUCCESS Then
        err.Raise rrSystemError, _
                  "RegRead.QueryValue", _
                  "System error " & CStr(mlngResult)
    End If
    Select Case mlngType
        Case REG_BINARY
            QueryValue = bytBuffer

        Case REG_MULTI_SZ, REG_SZ
            strBuffer = Left$(bytBuffer, lngDataSize \ 2 - 1)
            If mlngType = REG_SZ Then
                QueryValue = strBuffer
            Else
                QueryValue = Split(strBuffer, vbNullChar)
            End If

        Case REG_DWORD, REG_DWORD_BIG_ENDIAN
            If mlngType = REG_DWORD_BIG_ENDIAN Then
                bytTemp = bytBuffer(0)
                bytBuffer(0) = bytBuffer(3)
                bytBuffer(3) = bytTemp
                bytTemp = bytBuffer(1)
                bytBuffer(1) = bytBuffer(2)
                bytBuffer(2) = bytTemp
            End If
            CopyMemory VarPtr(lngDWORD), VarPtr(bytBuffer(0)), LenB(lngDWORD)
            QueryValue = lngDWORD

        Case Else
            err.Raise rrValTypeError, _
                      "RegRead.QueryValue", _
                      "Unsupported registry value type " & CStr(mlngType)
    End Select
End Function

Private Sub Class_Terminate()
    CloseKey
End Sub
