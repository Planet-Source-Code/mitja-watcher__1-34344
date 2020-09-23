Attribute VB_Name = "GetExeModule"
Option Explicit

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Private Declare Function CreateToolhelp32Snapshot Lib "KERNEL32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "KERNEL32" (ByVal hSnapshot As Long, lppe As Any) As Long
Private Declare Function Process32Next Lib "KERNEL32" (ByVal hSnapshot As Long, lppe As Any) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Any, ByVal lParam As Long) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Sub MoveMemory Lib "KERNEL32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Function lstrcpy Lib "KERNEL32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long

Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
    
Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long           '1 = Windows 95.
                                  '2 = Windows NT
   szCSDVersion As String * 128
End Type

Public Declare Function GetVersionExA Lib "KERNEL32" (lpVersionInformation As OSVERSIONINFO) As Integer

Global Win32 As Boolean

Public Type FileInfo
    CompanyName As String
    FileDescription As String
    FileVersion As String
    InternalName As String
    LegalCopyright As String
    OriginalFileName As String
    ProductName As String
    ProductVersion As String
End Type

Public Enum VerisonReturnValue
    eOK = 1
    eNoVersion = 2
End Enum


Private Search_hWnd As Long
Private Const MAX_PATH = 260

Private Const TH32CS_SNAPHEAPLIST = &H1
Private Const TH32CS_SNAPPROCESS = &H2
Private Const TH32CS_SNAPTHREAD = &H4
Private Const TH32CS_SNAPMODULE = &H8
Private Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST + TH32CS_SNAPPROCESS + TH32CS_SNAPTHREAD + TH32CS_SNAPMODULE)
Private Const TH32CS_INHERIT = &H80000000

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

Public Const PROCESS_QUERY_INFORMATION = 1024
Public Const PROCESS_VM_READ = 16

Private Function EnumWindowsCallback(ByVal x_hWnd As Long, ByVal lpData As Long) As Long
    Dim X As String, result As Long, hWndProcess As Long
    
    result = GetWindowThreadProcessId(x_hWnd, hWndProcess)
    If hWndProcess = lpData Then '
        If (GetParent(x_hWnd) = 0) And (IsWindowEnabled(x_hWnd) > 0) And (IsWindowVisible(x_hWnd) > 0) Then
            Search_hWnd = x_hWnd
            EnumWindowsCallback = 0
        Else
            EnumWindowsCallback = 1
        End If
    Else
        EnumWindowsCallback = 1
    End If
End Function

Public Function GetExe(wnd As Long)
    Dim hSnapshot As Long, lRet As Long, p As PROCESSENTRY32
    Dim proc As Long, ExeFile As String
    Dim tFile As FileInfo, procid As Long, n As Long
    Dim ErrVal As VerisonReturnValue
    Dim PName As String
    SendDebug "Executing GetExe Function.", 2
    On Error Resume Next
    SendDebug "   Searching for exe file of window " & wnd, 2
    If Win32 Then 'Win9x, Win2k
        p.dwSize = Len(p)
        hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPALL, ByVal 0)
        If Err > 0 Then SendDebug Error, 4
        If hSnapshot Then
            lRet = Process32First(hSnapshot, p)
            n = GetWindowThreadProcessId(wnd, procid)
            SendDebug "   Window thread process id of window " & wnd & " is " & procid, 2
            If Err > 0 Then SendDebug Error, 4
            Do While lRet
                proc = p.th32ProcessID
                If procid = proc Then
                    ExeFile = Left$(p.szExeFile, InStr(p.szExeFile, Chr$(0)) - 1)
                    SendDebug "   Process EXE file path is: " & ExeFile, 2
                    ErrVal = GetFileVersionInformation(ExeFile, tFile)
                    If Err > 0 Then SendDebug Error, 4
                    If Trim(tFile.ProductName) = "" Then
                      If Trim(tFile.FileDescription) = "" Then
                          PName = Resolve(ExeFile)
                          'If Trim(PName) = "" Then PName = ExeFile
                          If Err > 0 Then SendDebug Error, 4
                          SendDebug "   Coudn't find product name whit standard mode. Using filepath.", 2
                      Else
                          PName = tFile.FileDescription
                          SendDebug "   Coudn't find product name whit standard mode. Using File Description.", 2
                          If Err > 0 Then SendDebug Error, 4
                      End If
                    Else
                      PName = tFile.ProductName
                      SendDebug "   Product name is found.", 2
                      If Err > 0 Then SendDebug Error, 4
                    End If
                    SendDebug "   Result is: " & PName, 2
                    GetExe = PName
                    If Err > 0 Then SendDebug Error, 4
                    Exit Do
                End If
                lRet = Process32Next(hSnapshot, p)
                If Err > 0 Then SendDebug Error, 4
            Loop
            lRet = CloseHandle(hSnapshot)
        End If
    Else 'WinNT
        
       Dim cb As Long
       Dim cbNeeded As Long
       Dim NumElements As Long
       Dim ProcessIDs() As Long
       Dim cbNeeded2 As Long
       Dim NumElements2 As Long
       Dim Modules(1 To 200) As Long
       Dim ModuleName As String
       Dim nSize As Long
       Dim hProcess As Long
       Dim i As Long
       'Get the array containing the process id's for each process object
       cb = 8
       cbNeeded = 96
       Do While cb <= cbNeeded
          cb = cb * 2
          ReDim ProcessIDs(cb / 4) As Long
          lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
          If Err > 0 Then SendDebug Error, 4
       Loop
       If Err > 0 Then SendDebug Error, 4
       NumElements = cbNeeded / 4
       n = GetWindowThreadProcessId(wnd, procid)
       If Err > 0 Then SendDebug Error, 4
       'For i = 1 To NumElements
          'Get a handle to the Process
          SendDebug "   Opening process Nb. " & procid, 2
          hProcess = OpenProcess(PROCESS_QUERY_INFORMATION _
             Or PROCESS_VM_READ, 0, procid) 'ProcessIDs(i))
          If Err > 0 Then SendDebug Error, 4
          'Got a Process handle
          If hProcess <> 0 Then
              'Get an array of the module handles for the specified
              'process
              SendDebug "   Searching Process Modules", 2
              lRet = EnumProcessModules(hProcess, Modules(1), 200, _
                                           cbNeeded2)
              'If the Module Array is retrieved, Get the ModuleFileName
              If Err > 0 Then SendDebug Error, 4
              If lRet <> 0 Then
                 SendDebug "   Search form module complete", 2
                 ModuleName = Space(MAX_PATH)
                 nSize = 500
                 lRet = GetModuleFileNameExA(hProcess, Modules(1), _
                        ModuleName, nSize)
                 If Err > 0 Then SendDebug Error, 4
                 
                 ExeFile = Left(ModuleName, lRet)
                 SendDebug "   Module Exe Filename is " & ExeFile, 2
                 ErrVal = GetFileVersionInformation(ExeFile, tFile)
                If Err > 0 Then SendDebug Error, 4
                If Trim(tFile.ProductName) = "" Then
                  If Trim(tFile.FileDescription) = "" Then
                      PName = Resolve(ExeFile)
                      'If Trim(PName) = "" Then PName = ExeFile
                      If Err > 0 Then SendDebug Error, 4
                      SendDebug "   Coudn't find product name whit standard mode. Using filepath.", 2
                  Else
                      PName = tFile.FileDescription
                      SendDebug "   Coudn't find product name whit standard mode. Using File Description.", 2
                      If Err > 0 Then SendDebug Error, 4
                  End If
                Else
                  PName = tFile.ProductName
                  SendDebug "   Product name is found.", 2
                  If Err > 0 Then SendDebug Error, 4
                End If
                SendDebug "   Result is: " & PName, 2
                GetExe = PName
              End If
          End If
       
       lRet = CloseHandle(hProcess)
       If Err > 0 Then SendDebug Error, 4
       
    End If
    SendDebug "GetExe Function Complete", 2
End Function

Private Function HwndFromShell(ProcessX As Long) As Long
    Dim result As Long
    Search_hWnd = 0
    result = EnumWindows(AddressOf EnumWindowsCallback, ProcessX) '
    HwndFromShell = Search_hWnd
End Function

Private Function Resolve(strFile As String) As String
    Dim tempfile As String, pos As Long, EndPos As Long
    Dim startpos As Long, posProductName As String
    Dim orgProductName As String
    Resolve = strFile
    Exit Function
    Open strFile For Binary As #1 'Opens it for binary
        tempfile = Space(LOF(1)) 'Pads to length of string
        Get #1, , tempfile 'Dumps contents of file to string
    Close #1
    
    pos = InStr(1, tempfile, nullPadd("StringFileInfo"))
    startpos = pos
    If pos = 0 Then pos = 1
    startpos = InStr(pos, tempfile, nullPadd("ProductName"))
    If startpos > 0 Then
        posProductName = startpos + 28
        EndPos = InStr(posProductName, tempfile, String(3, Chr(0)))
        orgProductName = Mid(tempfile, posProductName, EndPos - (posProductName))
        Resolve = Replace(orgProductName, Chr(0), "")
    End If
    
End Function
    
Private Function nullPadd(strData As String) As String
    If strData = "" Then
        Exit Function 'If nothing their exit function
    End If
    
    'Dont declare variables until checked for data
    Dim lenData As Long
    Dim tmpString As String
    
    For lenData = 1 To Len(strData) 'Cycle through the data
        tmpString = tmpString & Chr(0) & Mid(strData, lenData, 1)  'Add nulls between every character
    Next lenData
    
    nullPadd = Chr(1) & tmpString 'Sends it back out w/ identifier
End Function

Function Replace(txt As String, kaj As String, scim As String)
    Dim X As Integer, txt2 As String
    If InStr(1, txt, kaj) = 0 Then
        txt2 = txt
        Exit Function
    End If
    For X = 1 To Len(txt)
        If Mid(txt, X, 1) = kaj Then
            txt2 = txt2 & scim
        Else
            txt2 = txt2 & Mid(txt, X, 1)
        End If
    Next X
    Replace = txt2
End Function

Private Function GetFileVersionInformation(ByRef pstrFieName As String, ByRef tFileInfo As FileInfo) As VerisonReturnValue

    Dim lBufferLen As Long, lDummy As Long
    Dim sBuffer() As Byte
    Dim lVerPointer As Long
    Dim lRet As Long
    Dim Lang_Charset_String As String
    Dim HexNumber As Long
    Dim i As Integer
    Dim strTemp As String
    
    'Clear the Buffer tFileInfo
    tFileInfo.CompanyName = ""
    tFileInfo.FileDescription = ""
    tFileInfo.FileVersion = ""
    tFileInfo.InternalName = ""
    tFileInfo.LegalCopyright = ""
    tFileInfo.OriginalFileName = ""
    tFileInfo.ProductName = ""
    tFileInfo.ProductVersion = ""
    lBufferLen = GetFileVersionInfoSize(pstrFieName, lDummy)

    If lBufferLen < 1 Then
        GetFileVersionInformation = eNoVersion
        'Exit Function
    End If

    ReDim sBuffer(lBufferLen)
    lRet = GetFileVersionInfo(pstrFieName, 0&, lBufferLen, sBuffer(0))

    If lRet = 0 Then
        GetFileVersionInformation = eNoVersion
        Exit Function
    End If

    lRet = VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lVerPointer, lBufferLen)

    If lRet = 0 Then
        GetFileVersionInformation = eNoVersion
        Exit Function
    End If

    Dim bytebuffer(255) As Byte
    MoveMemory bytebuffer(0), lVerPointer, lBufferLen
    HexNumber = bytebuffer(2) + bytebuffer(3) * &H100 + bytebuffer(0) * &H10000 + bytebuffer(1) * &H1000000
    Lang_Charset_String = Hex(HexNumber)
    'Pull it all apart:
    '04------= SUBLANG_ENGLISH_USA
    '--09----= LANG_ENGLISH
    '----04E4 = 1252 = Codepage for Windows:Multilingual

    Do While Len(Lang_Charset_String) < 8
        Lang_Charset_String = "0" & Lang_Charset_String
    Loop
    'If Mid(Lang_Charset_String, 1, 4) = "0000" Then Lang_Charset_String = "040904E4"
   
    Dim strVersionInfo(7) As String
    strVersionInfo(0) = "CompanyName"
    strVersionInfo(1) = "FileDescription"
    strVersionInfo(2) = "FileVersion"
    strVersionInfo(3) = "InternalName"
    strVersionInfo(4) = "LegalCopyright"
    strVersionInfo(5) = "OriginalFileName"
    strVersionInfo(6) = "ProductName"
    strVersionInfo(7) = "ProductVersion"
    Dim buffer As String

    For i = 0 To 7
        buffer = String(255, 0)
        strTemp = "\StringFileInfo\" & Lang_Charset_String _
        & "\" & strVersionInfo(i)
        lRet = VerQueryValue(sBuffer(0), strTemp, lVerPointer, lBufferLen)

        If lRet = 0 Then
            GetFileVersionInformation = eNoVersion
            If Mid(Lang_Charset_String, 1, 4) = "0000" Then Lang_Charset_String = "040904E4"
            'Exit Function
        End If

        lstrcpy buffer, lVerPointer
        buffer = Mid$(buffer, 1, InStr(buffer, vbNullChar) - 1)

        Select Case i
            Case 0
            tFileInfo.CompanyName = buffer
            Case 1
            tFileInfo.FileDescription = buffer
            Case 2
            tFileInfo.FileVersion = buffer
            Case 3
            tFileInfo.InternalName = buffer
            Case 4
            tFileInfo.LegalCopyright = buffer
            Case 5
            tFileInfo.OriginalFileName = buffer
            Case 6
            tFileInfo.ProductName = buffer
            Case 7
            tFileInfo.ProductVersion = buffer
        End Select
Next i
GetFileVersionInformation = eOK
End Function

Public Function getVersion() As Long
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer
   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)
   getVersion = osinfo.dwPlatformId
End Function


