Attribute VB_Name = "inifile_handler"
'INI Setting Helping-Handler
'-------------------------------
'  example for an INI file:
'
'       [Colors]  <---------- called 'section'
'       formColor = Blue
'        ^            ^
'        ^            ^
'   called 'key'    called value
'
'
' ...yes just like registry-files (which are actually INI files...)
'

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpbuffurnedString As String, ByVal nBuffSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Sub INIwrite(sSection As String, sKeyName As String, sValue As String, sinifile As String)
    On Error Resume Next
    Call WritePrivateProfileString(sSection, sKeyName, sValue, sinifile)
    DoEvents
End Sub

Public Function INIread(sSection As String, sKeyName As String, sValue As String, sinifile As String) As String
    INIread = vbNullString
    
    On Error Resume Next
    Dim dwSize As Long
    Dim nBuffSize As Long
    Dim buff As String
    
    buff = Space$(2048)
    nBuffSize = Len(buff)
    dwSize = GetPrivateProfileString(sSection, sKeyName, sDefValue, buff, nBuffSize, sinifile)
    If (dwSize > 0) Then INIread = Trim$(Left$(buff, dwSize))
End Function

