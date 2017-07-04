Attribute VB_Name = "modReg"
Option Explicit

Public Sub CreateRegLong(ByVal EnmHive As RegistryHives, ByVal StrSubKey As String, ByVal strValueName As String, ByVal LngData As Long, Optional ByVal EnmType As RegistryLongTypes = REG_DWORD_LITTLE_ENDIAN)
    Dim hKey As Long
    Call CreateSubKey(EnmHive, StrSubKey)
    hKey = GetSubKeyHandle(EnmHive, StrSubKey, KEY_ALL_ACCESS)
    RegSetValueEx hKey, strValueName, 0, EnmType, LngData, 4
    RegCloseKey hKey
End Sub

Public Sub CreateSubKey(ByVal EnmHive As RegistryHives, ByVal StrSubKey As String)
    Dim hKey As Long
    RegCreateKey EnmHive, StrSubKey & Chr(0), hKey
    RegCloseKey hKey
End Sub

Private Function GetSubKeyHandle(ByVal EnmHive As RegistryHives, ByVal StrSubKey As String, Optional ByVal EnmAccess As RegistryKeyAccess = KEY_READ) As Long
    Dim hKey As Long
    Dim retVal As Long
    retVal = RegOpenKeyEx(EnmHive, StrSubKey, 0, EnmAccess, hKey)
    If retVal <> ERROR_SUCCESS Then
        hKey = 0
    End If
    GetSubKeyHandle = hKey
End Function

Public Function SetKeyValue(lPredefinedKey As Long, sKeyName As String, sValueName As String, vValueSetting As Variant, lValueType As Long)
    Dim lRetVal As Long
    Dim hKey As Long
    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
    RegCloseKey (hKey)
End Function

Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    Dim lValue As Long
    Dim sValue As String
    Select Case lType
        Case REG_SZ
            sValue = vValue
            SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, sValue, Len(sValue))
        Case REG_DWORD
            lValue = vValue
            SetValueEx = RegSetValueExLong(hKey, sValueName, 0&, lType, lValue, 4)
    End Select
End Function

Public Function Query_Value(lPredefinedKey As Long, sKeyName As String, sValueName As String)
    Dim lRetVal As Long
    Dim hKey As Long
    Dim vValue As Variant
    lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
    lRetVal = Query_ValueEx(hKey, sValueName, vValue)
    Query_Value = vValue
    RegCloseKey (hKey)
End Function

Private Function Query_ValueEx(ByVal lhKey As Long, ByVal szValueName As String, vValue As Variant) As Long
    Dim cch As Long
    Dim lrc As Long
    Dim lType As Long
    Dim lValue As Long
    Dim sValue As String
    On Error GoTo QueryValueExError
    lrc = RegQueryValueExNULL(lhKey, szValueName, 0&, lType, 0&, cch)
    If lrc <> ERROR_NONE Then Error 5
    Select Case lType
        Case REG_SZ:
        sValue = String(cch, 0)
        lrc = RegQueryValueExString(lhKey, szValueName, 0&, lType, sValue, cch)
        If lrc = ERROR_NONE Then
            vValue = Left$(sValue, cch)
        Else
            vValue = Empty
        End If
    Case REG_DWORD:
        lrc = RegQueryValueExLong(lhKey, szValueName, 0&, lType, lValue, cch)
        If lrc = ERROR_NONE Then vValue = lValue
    Case Else
        lrc = -1
    End Select
QueryValueExExit:
    Query_ValueEx = lrc
    Exit Function
QueryValueExError:
    Resume QueryValueExExit
End Function

