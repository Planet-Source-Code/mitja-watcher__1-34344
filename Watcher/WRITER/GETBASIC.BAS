Attribute VB_Name = "basAPI"
Option Explicit

' General API functions.

Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOREPOSITION = &H200
Private Const SWP_NOSIZE = &H1

Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "KERNEL32" (ByVal lpBuffer As String, nSize As Long) As Long
Private TaskBarhWnd As Long
Public Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

Public Function GetUserName() As String
   Dim UserName As String * 255

   Call GetUserNameA(UserName, 255)
   GetUserName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

Public Function GetComputerName() As String
   Dim UserName As String * 255

   Call GetComputerNameA(UserName, 255)
   GetComputerName = Left$(UserName, InStr(UserName, Chr$(0)) - 1)
End Function

Public Function GetActiveWindowTitle(ByVal ReturnParent As Boolean) As String
   Dim i As Long
   Dim j As Long
   'SendDebug "Executing GetActiveWindowTitle Function", 2
   i = GetForegroundWindow
      
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
   
      i = j
   End If
   'SendDebug "   Result is " & GetWindowTitle(i), 2
   GetActiveWindowTitle = GetWindowTitle(i)
   'SendDebug "GetActiveWindowTitle Complete", 2
End Function
'
' Returns the handle of the active window.
'    if GetParent = true then the parent window is
'                   returned.
'
Public Function GetActiveWindow(ByVal ReturnParent As Boolean) As Long
   Dim i As Long
   Dim j As Long
   
   i = GetForegroundWindow
      
   If ReturnParent Then
      Do While i <> 0
         j = i
         i = GetParent(i)
      Loop
   
      i = j
   End If
   
   GetActiveWindow = i
End Function


Public Function GetWindowTitle(ByVal hwnd As Long) As String
   Dim l As Long
   Dim s As String
   
   l = GetWindowTextLength(hwnd)
   s = Space(l + 1)
   
   GetWindowText hwnd, s, l + 1
   
   GetWindowTitle = Left$(s, l)
End Function

Public Sub Pause(ByVal seconds As Single)
   Call Sleep(Int(seconds * 1000#))
End Sub

