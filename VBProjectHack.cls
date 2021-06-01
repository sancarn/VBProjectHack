VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBProjectHack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal reserved As Long, ByVal lpType As Long, lpData As Any, ByVal lpcbData As Long) As Long
Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const KEY_ALL_ACCESS As Long = &H3F

Private pIsHacked As Boolean
Private bKeys(10 To 20) As Boolean

'Sub Workbook_Open()
'  With new VBProjectHack
'    Dim wb as workbook: set wb = workbooks.open("...")
'    ThisWorkbook.VBProject.References.AddFromFile(wb.FullMame)
'  End With
'End Sub
'Sub Workbook_BeforeClose()
'  With new VBProjectHack
'    'Remove Reference
'    wb.close
'  End With
'End Sub


Private Sub Class_Initialize()
  Dim bIsHacked As Boolean
  If Not isVBEExtinsible() Then
    Dim i As Long: For i = 10 To 20
      Dim sKey As String: sKey = "Software\Microsoft\Office\" & i & ".0\Excel\Security"
      If RegistryValueExists(sKey) Then
        If Not RegistryGetAccessVBOM(sKey) Then
          Call RegistrySetAccessVBOM(sKey, 1)
          bKeys(i) = True
          bIsHacked = True
        End If
      End If
    Next
    Call EnsureExtended
    If Not isVBEExtinsible() Then
      MsgBox "Unable to automatically set VBE Extensibility. Please set manually and try again", vbCritical
      End
    End If
    pIsHacked = bIsHacked
  End If
End Sub

Private Sub Class_Terminate()
  'If hacked, remove hack
  If pIsHacked Then
    Dim i As Long: For i = 10 To 20
      If bKeys(i) Then
        Dim sKey As String: sKey = "Software\Microsoft\Office\" & i & ".0\Excel\Security"
        Call RegistrySetAccessVBOM(sKey, 0)
      End If
    Next
    Call EnsureExtended
  End If
End Sub




'Open and close macro security to force Excel to check/update VBE security
Private Sub EnsureExtended()
  Call AppActivate(Windows(1).Caption)
  Call SendKeys("%tms{enter}")
  Call Wait(1)
End Sub

'Wait an amount of seconds
Private Sub Wait(ByVal nSeconds As Long)
  Dim start As Date: start = Now()
  While Now() < start + nSeconds / 24 / 60 / 60
    DoEvents
  Wend
End Sub

'Check if VBE is extensible
'@returns {Boolean} True if VBE Is Extensible (i.e. checked within the IDE, rather than the registry)
Private Function isVBEExtinsible() As Boolean
  On Error Resume Next
  isVBEExtinsible = Len(ThisWorkbook.VBProject.Name) > 0
End Function

'Check whether a registry key value exists
'@param {String} obtain value from this key
'@returns {Boolean} True if key exists, False otherwise
Private Function RegistryValueExists(ByVal sKey As String) As Boolean
  Const HKEY_CURRENT_USER As Long = &H80000001
  Const KEY_ALL_ACCESS As Long = &H3F
  Dim hKey As Long
  Dim res As Long: res = RegOpenKeyEx(HKEY_CURRENT_USER, sKey, 0&, KEY_ALL_ACCESS, hKey)
  If res = 0 Then
    res = RegQueryValueEx(hKey, "AccessVBOM", 0&, 0&, 0&, 0&)
    RegistryValueExists = res = 0 Or res = 234
  End If
  Call RegCloseKey(hKey)
End Function

'Set registry key AccessVBOM to a DWORD value (1 or 0)
'@param {String} obtain value from this key
'@returns {Long} HResult of RegSetValueEx()
Private Function RegistrySetAccessVBOM(ByVal sKey As String, ByVal lValue As Long) As Long
  Const HKEY_CURRENT_USER As Long = &H80000001
  Const KEY_ALL_ACCESS As Long = &H3F
  Const REG_DWORD As Long = 4&
  Dim hKey As Long
  Dim res As Long: res = RegOpenKeyEx(HKEY_CURRENT_USER, sKey, 0&, KEY_ALL_ACCESS, hKey)
  If res = 0 Then RegistrySetAccessVBOM = RegSetValueEx(hKey, "AccessVBOM", 0&, REG_DWORD, lValue, 4&) = 0
  Call RegCloseKey(hKey)
End Function

'Obtain the value of AccessVBOM registry keey
'@param {String} obtain value from this key
'@returns {Boolean} True if access permitted, False if access restricted
Private Function RegistryGetAccessVBOM(ByVal sKey As String) As Boolean
  Const HKEY_CURRENT_USER As Long = &H80000001
  Const KEY_ALL_ACCESS As Long = &H3F
  Const REG_DWORD As Long = 4&
  Dim iBool As Long
  Dim res As Long: res = RegOpenKeyEx(HKEY_CURRENT_USER, sKey, 0&, KEY_ALL_ACCESS, hKey)
  If res = 0 Then res = RegQueryValueEx(hKey, "AccessVBOM", 0&, REG_DWORD, iBool, 4&)
  RegistryGetAccessVBOM = iBool = 1
  Call RegCloseKey(hKey)
End Function
