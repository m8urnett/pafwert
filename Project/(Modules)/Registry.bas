Attribute VB_Name = "modRegistry"
Option Explicit

Private Const REG_SZ                    As Long = 1
Private Const HKEY_LOCAL_MACHINE        As Long = &H80000002
Private Const BASE_KEY                  As String = "SOFTWARE"
Private Const ERROR_NONE                As Long = 0
Private Const ERROR_KEY_DOES_NOT_EXIST  As Long = 2
Private Const READ_CONTROL              As Long = &H20000
Private Const STANDARD_RIGHTS_READ      As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_ALL       As Long = &H1F0000
Private Const KEY_QUERY_VALUE           As Long = &H1
Private Const KEY_SET_VALUE             As Long = &H2
Private Const KEY_CREATE_SUB_KEY        As Long = &H4
Private Const KEY_ENUMERATE_SUB_KEYS    As Long = &H8
Private Const KEY_NOTIFY                As Long = &H10
Private Const KEY_CREATE_LINK           As Long = &H20
Private Const SYNCHRONIZE               As Long = &H100000
Private Const KEY_ALL_ACCESS            As Long = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Private Const KEY_READ                  As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

Private Declare Function RegCloseKey _
                Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegCreateKeyEx _
                Lib "advapi32.dll" _
                Alias "RegCreateKeyExA" (ByVal hKey As Long, _
                                         ByVal lpSubKey As String, _
                                         ByVal Reserved As Long, _
                                         ByVal lpClass As String, _
                                         ByVal dwOptions As Long, _
                                         ByVal samDesired As Long, _
                                         ByVal lpSecurityAttributes As Long, _
                                         phkResult As Long, _
                                         lpdwDisposition As Long) As Long

Private Declare Function RegOpenKeyEx _
                Lib "advapi32.dll" _
                Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       phkResult As Long) As Long

Private Declare Function RegQueryValueExString _
                Lib "advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          lpType As Long, _
                                          ByVal lpData As String, _
                                          lpcbData As Long) As Long

Private Declare Function RegQueryValueExNULL _
                Lib "advapi32.dll" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          lpType As Long, _
                                          ByVal lpData As Long, _
                                          lpcbData As Long) As Long

Private Declare Function RegSetValueExString _
                Lib "advapi32.dll" _
                Alias "RegSetValueExA" (ByVal hKey As Long, _
                                        ByVal lpValueName As String, _
                                        ByVal Reserved As Long, _
                                        ByVal dwType As Long, _
                                        ByVal lpData As String, _
                                        ByVal cbData As Long) As Long

'*******************************************************************************
' SaveStringSetting (SUB)
'
' DESCRIPTION:
' Own version of VB's SaveSetting to store strings under
' HKEY_LOCAL_MACHINE\SOFTWARE instead of
' HKEY_CURRENT_USER\Software\VB and VBA Program Settings
'
' PARAMETERS:
' (In) - sAppName - String - The first level
' (In) - sSection - String - The second level
' (In) - sKey     - String - The key in the second level
' (In) - sSetting - String - The new value for the key
'*******************************************************************************
Public Sub SaveStringSetting(ByVal sAppName As String, _
                             ByVal sSection As String, _
                             ByVal sKey As String, _
                             ByVal sSetting As String)
   Dim lRetVal              As Long
   Dim sNewKey              As String
   Dim lDisposition         As Long
   Dim lHandle              As Long
   Dim lErrNumber           As Long
   Dim sErrDescription      As String
   Dim sErrSource           As String
   On Error GoTo ERROR_HANDLER

   If Trim(sAppName) = "" Then
      Err.Raise vbObjectError + 1000, , "AppName may not be empty"
   End If

   If Trim(sSection) = "" Then
      Err.Raise vbObjectError + 1001, , "Section may not be empty"
   End If

   If Trim(sKey) = "" Then
      Err.Raise vbObjectError + 1002, , "Key may not be empty"
   End If

   sNewKey = BASE_KEY & "\" & Trim(sAppName) & "\" & Trim(sSection)
   ' Create the key or open it if it already exists
   lRetVal = RegCreateKeyEx(HKEY_LOCAL_MACHINE, sNewKey, 0, vbNullString, 0, KEY_ALL_ACCESS, 0, lHandle, lDisposition)

   If lRetVal <> ERROR_NONE Then
      Err.Raise vbObjectError + 2000 + lRetVal, , "Could not open/create registry section"
   End If

   ' Set the key value
   lRetVal = RegSetValueExString(lHandle, sKey, 0, REG_SZ, sSetting, Len(sSetting))

   If lRetVal <> ERROR_NONE Then
      Err.Raise vbObjectError + 2000 + lRetVal, , "Could not set key value"
   End If

TIDY_UP:
   On Error Resume Next
   RegCloseKey lHandle

   If lErrNumber <> 0 Then
      On Error GoTo 0
      Err.Raise lErrNumber, sErrSource, sErrDescription
   End If

   Exit Sub

ERROR_HANDLER:
   lErrNumber = Err.Number
   sErrSource = Err.Source
   sErrDescription = Err.Description
   Resume TIDY_UP
End Sub

'*******************************************************************************
' GetStringSetting (FUNCTION)
'
' DESCRIPTION:
' Own version of VB's GetSetting to retrieve strings under
' HKEY_LOCAL_MACHINE\SOFTWARE instead of
' HKEY_CURRENT_USER\Software\VB and VBA Program Settings
'
' PARAMETERS:
' (In) - sAppName - String - The first level
' (In) - sSection - String - The second level
' (In) - sKey     - String - The key in the second level
' (In) - sDefault - String -
'
' RETURN VALUE:
' String - The value stored in the key or sDefault if not found
'*******************************************************************************
Public Function GetStringSetting(ByVal sAppName As String, _
                                 ByVal sSection As String, _
                                 ByVal sKey As String, _
                                 Optional ByVal sDefault As String) As String
   Dim lRetVal              As Long
   Dim sFullKey             As String
   Dim lHandle              As Long
   Dim lType                As Long
   Dim lLength              As Long
   Dim sValue               As String
   Dim lErrNumber           As Long
   Dim sErrDescription      As String
   Dim sErrSource           As String
   On Error GoTo ERROR_HANDLER

   If Trim(sAppName) = "" Then
      Err.Raise vbObjectError + 1000, , "AppName may not be empty"
   End If

   If Trim(sSection) = "" Then
      Err.Raise vbObjectError + 1001, , "Section may not be empty"
   End If

   If Trim(sKey) = "" Then
      Err.Raise vbObjectError + 1002, , "Key may not be empty"
   End If

   sFullKey = BASE_KEY & "\" & Trim(sAppName) & "\" & Trim(sSection)
   ' Open up the key
   lRetVal = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sFullKey, 0, KEY_READ, lHandle)

   If lRetVal <> ERROR_NONE Then
      If lRetVal = ERROR_KEY_DOES_NOT_EXIST Then
         GetStringSetting = sDefault
         Exit Function

      Else
         Err.Raise vbObjectError + 2000 + lRetVal, , "Could not open registry section"
      End If
   End If

   ' Get size and type
   lRetVal = RegQueryValueExNULL(lHandle, sKey, 0, lType, 0, lLength)

   If lRetVal <> ERROR_NONE Then
      GetStringSetting = sDefault
      Exit Function

   End If

   ' Is it stored as a string in the registry?
   If lType = REG_SZ Then
      sValue = String(lLength, 0)

      If lLength = 0 Then
         GetStringSetting = ""
      Else
         lRetVal = RegQueryValueExString(lHandle, sKey, 0, lType, sValue, lLength)

         If lRetVal = ERROR_NONE Then
            GetStringSetting = Left(sValue, lLength - 1)
         Else
            GetStringSetting = sDefault
         End If
      End If

   Else
      Err.Raise vbObjectError + 2000 + lType, , "Registry data not a string"
   End If

TIDY_UP:
   On Error Resume Next
   RegCloseKey lHandle

   If lErrNumber <> 0 Then
      On Error GoTo 0
      Err.Raise lErrNumber, sErrSource, sErrDescription
   End If

   Exit Function

ERROR_HANDLER:
   lErrNumber = Err.Number
   sErrSource = Err.Source
   sErrDescription = Err.Description
   Resume TIDY_UP
End Function
