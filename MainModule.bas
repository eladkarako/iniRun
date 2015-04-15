Attribute VB_Name = "MainModule"
Option Explicit
Private Const lmaxPath As Long = 260&
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal lBuffer As Long) As Long

Private Const SW_FORCEMINIMIZE As Long = &HB&    'Minimizes a window, even if the thread that owns the window is not responding. This flag should only be used when minimizing windows from a different thread.
Private Const SW_HIDE As Long = &H0&             'Hides the window and activates another window.
Private Const SW_MAXIMIZE As Long = &H3&         'Maximizes the specified window.
Private Const SW_MINIMIZE As Long = &H6&         'Minimizes the specified window and activates the next top-level window in the Z order.
Private Const SW_RESTORE As Long = &H9&          'Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
Private Const SW_SHOW As Long = &H5&             'Activates the window and displays it in its current size and position.
Private Const SW_SHOWDEFAULT As Long = &HA&      'Sets the show state based on the SW_ value specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application.
Private Const SW_SHOWMAXIMIZED As Long = &H3&    'Activates the window and displays it as a maximized window.
Private Const SW_SHOWMINIMIZED As Long = &H2&    'Activates the window and displays it as a minimized window.
Private Const SW_SHOWMINNOACTIVE As Long = &H7&  'Displays the window as a minimized window. This value is similar to SW_SHOWMINIMIZED, except the window is not activated.
Private Const SW_SHOWNA As Long = &H8&           'Displays the window in its current size and position. This value is similar to SW_SHOW, except that the window is not activated.
Private Const SW_SHOWNOACTIVATE As Long = &H4&   'Displays a window in its most recent size and position. This value is similar to SW_SHOWNORMAL, except that the window is not activated.
Private Const SW_SHOWNORMAL As Long = &H1&       'Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time.

'Private Const SW_SHOWNORMAL As Long = &H1&

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Const csSection1 As String = "Information"
Private Const csSection1Key1 As String = "Parent_Folder"
Private Const csSection1Key2 As String = "Arguments"
Private Const csSection1Key3 As String = "Full_Path"
Private Const csSection1Key4 As String = "Run_Mode"




Public Sub Main()
    On Error GoTo endsub
    
    'The Name Of Associated INI file.
    Dim sinifile As String
    sinifile = App.Path & "\" & App.EXEName & ".ini"
    If (Dir$(sinifile) = vbNullString) Then
        Open sinifile For Append As #1
        DoEvents
        Close #1

        inifile_handler.INIwrite csSection1, csSection1Key1, "Fill " & csSection1Key1 & " Here..." & vbNewLine & _
"; for example: Parent_Folder=C:\WINDOWS\system32\" & vbNewLine, sinifile

        inifile_handler.INIwrite csSection1, csSection1Key2, "Fill " & csSection1Key2 & " Here..." & vbNewLine & _
"; (no need for inverted commas between arguments, even if they are long..) for example: Arguments=-arg1 -arg2" & vbNewLine, sinifile
        
        inifile_handler.INIwrite csSection1, csSection1Key3, "Fill " & csSection1Key3 & " Here..." & vbNewLine & _
"; for example: Full_Path=C:\WINDOWS\system32\calc.exe" & vbNewLine, sinifile

        inifile_handler.INIwrite csSection1, csSection1Key4, "Fill " & csSection1Key4 & " Here..." & vbNewLine & _
 _
"; Run_Mode Fill Options (default/missing/not of the following- SW_SHOWNORMAL will be used instead.)" & vbNewLine & _
"; ------------------------------------------------------------------------------------------------------" & vbNewLine & _
";   SW_FORCEMINIMIZE    Minimizes a window, even if the thread that owns the window is not responding. This flag should only be used when minimizing windows from a different thread." & vbNewLine & _
";   SW_HIDE             Hides the window and activates another window." & vbNewLine & _
";   SW_MAXIMIZE         Maximizes the specified window." & vbNewLine & _
";   SW_MINIMIZE         Minimizes the specified window and activates the next top-level window in the Z order." & vbNewLine & _
";   SW_RESTORE          Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window." & vbNewLine & _
";   SW_SHOW             Activates the window and displays it in its current size and position." & vbNewLine & _
";   SW_SHOWDEFAULT      Sets the show state based on the SW_ value specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application." & vbNewLine & _
";   SW_SHOWMAXIMIZED    Activates the window and displays it as a maximized window." & vbNewLine & _
";   SW_SHOWMINIMIZED    Activates the window and displays it as a minimized window." & vbNewLine & _
";   SW_SHOWMINNOACTIVE  Displays the window as a minimized window. This value is similar to SW_SHOWMINIMIZED, except the window is not activated." & vbNewLine & _
";   SW_SHOWNA           Displays the window in its current size and position. This value is similar to SW_SHOW, except that the window is not activated." & vbNewLine & _
";   SW_SHOWNOACTIVATE   Displays a window in its most recent size and position. This value is similar to SW_SHOWNORMAL, except that the window is not activated." & vbNewLine & _
";   SW_SHOWNORMAL       Activates and displays a window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when displaying the window for the first time." & vbNewLine _
       , sinifile
        
        
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
    
    
    'preread values from ini file, value 4
    Dim stmp_runmode As String
    stmp_runmode = LCase$(Trim$(inifile_handler.INIread(csSection1, csSection1Key4, vbNullString, sinifile)))
    If (stmp_runmode = vbNullString) Then stmp_runmode = SW_SHOWNORMAL
    Select Case stmp_runmode
    
      Case "SW_FORCEMINIMIZE"
          stmp_runmode = SW_FORCEMINIMIZE
      Case "SW_HIDE"
          stmp_runmode = SW_HIDE
      Case "SW_MAXIMIZE"
          stmp_runmode = SW_MAXIMIZE
      Case "SW_MINIMIZE"
          stmp_runmode = SW_MINIMIZE
      Case "SW_RESTORE"
          stmp_runmode = SW_RESTORE
      Case "SW_SHOW"
          stmp_runmode = SW_SHOW
      Case "SW_SHOWDEFAULT"
          stmp_runmode = SW_SHOWDEFAULT
      Case "SW_SHOWMAXIMIZED"
          stmp_runmode = SW_SHOWMAXIMIZED
      Case "SW_SHOWMINIMIZED"
          stmp_runmode = SW_SHOWMINIMIZED
      Case "SW_SHOWMINNOACTIVE"
          stmp_runmode = SW_SHOWMINNOACTIVE
      Case "SW_SHOWNA"
          stmp_runmode = SW_SHOWNA
      Case "SW_SHOWNOACTIVATE"
          stmp_runmode = SW_SHOWNOACTIVATE
      Case "SW_SHOWNORMAL"
          stmp_runmode = SW_SHOWNORMAL
      Case Else
          stmp_runmode = SW_SHOWNORMAL
      End Select
   
    'execute acording to readed data.
    ShellExecute ByVal 0&, vbNullString, _
                 stmp_fullpath, stmp_arg, stmp_parentpath, stmp_runmode

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

