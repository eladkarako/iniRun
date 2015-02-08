Attribute VB_Name = "MainModule"
Option Explicit
Private Const lmaxPath As Long = 260&
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Private Const SW_SHOWNORMAL As Long = &H1&
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Const csSection1 As String = "Information"
Private Const csSection1Key1 As String = "Parent_Folder"
Private Const csSection1Key2 As String = "Arguments"
Private Const csSection1Key3 As String = "Full_Path"




Public Sub Main()
    On Error GoTo endsub
    
    'The Name Of Associated INI file.
    Dim sinifile As String
    sinifile = App.Path & "\" & App.EXEName & ".ini"
    If (Dir$(sinifile) = vbNullString) Then
        Open sinifile For Append As #1
        DoEvents
        Close #1

        inifile_handler.INIwrite csSection1, csSection1Key1, "Fill " & csSection1Key1 & " Here...", sinifile
        inifile_handler.INIwrite csSection1, csSection1Key2, "Fill " & csSection1Key2 & " Here...", sinifile
        inifile_handler.INIwrite csSection1, csSection1Key3, "Fill " & csSection1Key3 & " Here...", sinifile
        MsgBox "Fill Information Inside " & sinifile & " .", vbInformation Or vbOKOnly, "Creating New INI File, and Quiting..."
        GoTo endsub
    End If
    sinifile = getShortPath(sinifile)
    
    
    'preread values from ini file, value 1
    Dim stmp_parentpath As String
    stmp_parentpath = Trim$(inifile_handler.INIread(csSection1, csSection1Key1, vbNullString, sinifile))
    If (Dir$(stmp_parentpath) <> vbNullString) Then stmp_parentpath = getShortPath(stmp_parentpath)
    
    
    'preread values from ini file, value 2
    Dim stmp_arg As String
    stmp_arg = Trim$(inifile_handler.INIread(csSection1, csSection1Key2, vbNullString, sinifile))
    
    
    'preread values from ini file, value 3
    Dim stmp_fullpath As String
    stmp_fullpath = Trim$(inifile_handler.INIread(csSection1, csSection1Key3, vbNullString, sinifile))
    If (stmp_fullpath = vbNullString) Then GoTo endsub
    
    
    stmp_fullpath = getShortPath(stmp_fullpath)
    
    
    
    'execute acording to readed data.
    ShellExecute ByVal 0&, vbNullString, _
                 stmp_fullpath, stmp_arg, stmp_parentpath, SW_SHOWNORMAL

endsub:
End Sub


Private Function getShortPath(ByRef sFileOrFolder_Path As String) As String
    On Error Resume Next
    
    Dim sBuffer As String
    sBuffer = String$(lmaxPath, 0)
    
    Dim Length As Long
    Length = GetShortPathName(sFileOrFolder_Path, sBuffer, lmaxPath)
    getShortPath = Trim$(UCase$(Left$(sBuffer, Length)))
    
    If (getShortPath = vbNullString) Then getShortPath = Trim$(sFileOrFolder_Path)
End Function

