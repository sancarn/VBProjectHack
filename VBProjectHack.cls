VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBProjectHack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Usage:
'  Sub Workbook_Open()
'    With new VBProjectHack
'      Dim wb as workbook: set wb = workbooks.open("...")
'      ThisWorkbook.VBProject.References.AddFromFile(wb.FullMame)
'    End With
'  End Sub
'
'  Sub Workbook_BeforeClose()
'    With new VBProjectHack
'      wb.close
'    End With
'  End Sub

#if VBA7 then
  Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
  Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
  Private Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal reserved As Long, ByVal lpType As Long, lpData As Any, ByVal lpcbData As Long) As Long
  Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
#else
  Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
  Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
  Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal reserved As Long, ByVal lpType As Long, lpData As Any, ByVal lpcbData As Long) As Long
  Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
#end if

Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const KEY_ALL_ACCESS As Long = &H3F
Private Enum HackType
    None
    SoftwareMicrosoft
    SoftwarePoliciesMicrosoft
End Enum
Private pAccessAvailable As Boolean
Private pErrors As New Collection

Private pIsHacked As Boolean
Private pHackType As HackType


Public Property Get AccessAvailable() As Boolean
  AccessAvailable = pAccessAvailable
End Property
Public Property Get IsHacked() As Boolean
  IsHacked = pIsHacked
End Property
Public Property Get Errors() As Collection
  Set Errors = pErrors
End Property

Private Sub Class_Initialize()
  Dim bIsHacked As Boolean
  If Not isVBEExtinsible() Then
    If CheckSetKey("Software\Microsoft\Office\" & Application.Version & "\Excel\Security", True) Then
        bIsHacked = True
        pHackType = SoftwareMicrosoft
    ElseIf CheckSetKey("Software\Policies\Microsoft\Office\" & Application.Version & "\Excel\Security", True) Then
        bIsHacked = True
        pHackType = SoftwarePoliciesMicrosoft
    Else
        bIsHacked = False
    End If
    
    Call EnsureExtended
    
    If isVBEExtinsible() Then
      pAccessAvailable = True
    Else
      pAccessAvailable = False
      pErrors.Add "Unable to automatically set VBE Extensibility. Please set manually and try again"
    End If
    
    pIsHacked = bIsHacked
  Else
    pAccessAvailable = True
  End If
End Sub

Private Sub Class_Terminate()
  'If hacked, remove hack
  If pIsHacked Then
    Select Case pHackType
      Case HackType.SoftwareMicrosoft
        Call RegistrySetAccessVBOM("Software\Microsoft\Office\" & Application.Version & "\Excel\Security", 0)
      Case HackType.SoftwarePoliciesMicrosoft
        Call RegistrySetAccessVBOM("Software\Policies\Microsoft\Office\" & Application.Version & "\Excel\Security", 0)
    End Select
    Call EnsureExtended
  End If
End Sub


Private Function CheckSetKey(sKey As String, toValue As Boolean) As Boolean
    CheckSetKey = False
    If RegistryKeyExists(sKey) Then
      If RegistryValueExists(sKey) Then
        If Not RegistryGetAccessVBOM(sKey) Then
          Call RegistrySetAccessVBOM(sKey, IIf(toValue, 1, 0))
          CheckSetKey = True
        End If
      End If
    End If
End Function

'Open and close macro security to force Excel to check/update VBE security
Private Sub EnsureExtended()
  Call AppActivate(Windows(1).Caption)
  Call SendKeys("{enter}")                                  'Async function, will execute after ExecuteMso is called
  Call Application.CommandBars.ExecuteMso("MacroSecurity")  'Sync function, will execute and wait for enter/user to clear menu.
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
  isVBEExtinsible = Len(ThisWorkbook.VBProject.name) > 0
End Function

'Check whether a registry key exists
'@param {String} obtain value from this key
'@returns {Boolean} True if key exists, False otherwise
Private Function RegistryKeyExists(ByVal sKey As String) As Boolean
  Const HKEY_CURRENT_USER As Long = &H80000001
  Const KEY_READ As Long = &H20019
  Dim hKey As Long
  Dim res As Long: res = RegOpenKeyEx(HKEY_CURRENT_USER, sKey, 0&, KEY_READ, hKey)
  RegistryKeyExists = res = 0
End Function

'Check whether a registry key value exists
'@param {String} obtain value from this key
'@returns {Boolean} True if key value exists, False otherwise
Private Function RegistryValueExists(ByVal sKey As String) As Boolean
  Const HKEY_CURRENT_USER As Long = &H80000001
  Const KEY_READ As Long = &H20019
  Dim hKey As Long
  Dim res As Long: res = RegOpenKeyEx(HKEY_CURRENT_USER, sKey, 0&, KEY_READ, hKey)
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
  Const KEY_SET_VALUE As Long = &H2
  Const REG_DWORD As Long = 4&
  Dim hKey As Long
  Dim res As Long: res = RegOpenKeyEx(HKEY_CURRENT_USER, sKey, 0&, KEY_SET_VALUE, hKey)
  If res = 0 Then RegistrySetAccessVBOM = RegSetValueEx(hKey, "AccessVBOM", 0&, REG_DWORD, lValue, 4&) = 0
  Call RegCloseKey(hKey)
End Function

'Obtain the value of AccessVBOM registry keey
'@param {String} obtain value from this key
'@returns {Boolean} True if access permitted, False if access restricted
Private Function RegistryGetAccessVBOM(ByVal sKey As String) As Boolean
  Const HKEY_CURRENT_USER As Long = &H80000001
  Const KEY_READ As Long = &H20019
  Const REG_DWORD As Long = 4&
  Dim iBool As Long
  Dim res As Long: res = RegOpenKeyEx(HKEY_CURRENT_USER, sKey, 0&, KEY_READ, hKey)
  If res = 0 Then res = RegQueryValueEx(hKey, "AccessVBOM", 0&, REG_DWORD, iBool, 4&)
  RegistryGetAccessVBOM = iBool = 1
  Call RegCloseKey(hKey)
End Function
