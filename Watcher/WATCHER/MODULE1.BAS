Attribute VB_Name = "Module1"
Private Declare Function GetUserNameA Lib "advapi32.dll" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerNameA Lib "KERNEL32" (ByVal lpBuffer As String, nSize As Long) As Long

Global sexes As Integer
Global Baza As Database
Global WrkSpace As Workspace
Global DebugLevel As Integer

Type Keys
   Date As Date
   Duration As Integer
   State As Boolean
End Type

Type Prog
   ime As String
   Title As String
   Date As Date
   Duration As Integer
End Type

Public Prog() As Prog
Public Tipke() As Keys
Global User As Boolean
Global Computer As Boolean
Global LocalIP As String
Global Licences As Integer

Sub Main()
   ReDim Prog(0)
   ReDim Tipke(0)
   Dim Txt As String
   Dim tim As Long
   Dim Datum As Date
   If Command = "/SetUp" Then
      Form6.Visible = True
      Form6.ZOrder 0
   Else
      Form8.Visible = True
      DoEvents
      On Error Resume Next
      Open App.Path & "\Info.dat" For Input As #1
      If Err > 0 Then
         If Err = 53 Then
            Form9.Visible = True
            Form9.ZOrder 0
            Exit Sub
         Else
            End
         End If
      End If
      Do Until Trim(Txt) = "-Begin-"
         Line Input #1, Txt
      Loop
      Do Until Trim(Txt) = "-End-" Or EOF(1)
         Line Input #1, Txt
         If Txt = "-End-" Then Exit Do
         Txt = DeCrypt(Txt)
         cmd = Mid(Txt, 1, InStr(1, Txt, "=") - 1)
         con = Mid(Txt, InStr(1, Txt, "=") + 1)
         Select Case cmd
            Case "UserName"
               If con <> GetUserName Then
                  End
               Else
                  User = True
               End If

            Case "ComputerName"
               If con <> GetComputerName Then
                  End
               Else
                  Computer = True
               End If

            Case "LocalIP"
               LocalIP = con
               'If con <> Form1.Winsock1.LocalIP Then End

            Case "NumberOfLicences"
               Licences = con
            
            Case "Datum"
               Datum = con
               
         End Select
      Loop
      If User = False Then End
      If Computer = False Then End
'      If DateDiff("d", Datum, Date) > 30 Then
'         MsgBox "Your 30 Day free trial has expired. Register today and resume use.", vbExclamation, "Watcher"
'         End
'      End If
      'Licences = 20
      Set WrkSpace = DBEngine.CreateWorkspace("Watcher", "admin", "")
      Set Baza = WrkSpace.OpenDatabase(App.Path & "\Data.mdb", False, Flase, ";PWD=enigma")
      tim = Timer + 5
      Do Until tim < Timer
         DoEvents
      Loop
      Form8.Visible = False
      MDIForm1.Visible = True
   End If
End Sub

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

Function Crypt(Txt As String) As String
   Do Until Len(Txt) / 3 = Fix(Len(Txt) / 3)
      Txt = Txt & " "
   Loop
   For x = 1 To Len(Txt) Step 3
      w1 = w1 & Chr(Asc(Mid(Txt, x, 1)) + 1)
      w2 = w2 & Chr(Asc(Mid(Txt, x + 1, 1)) + 1)
      w3 = w3 & Chr(Asc(Mid(Txt, x + 2, 1)) + 1)
   Next x
   txt2 = Chr(64 + Len(Txt))
   txt2 = txt2 & w1 & w2 & w3
   For x = Len(txt2) To 49
      txt2 = txt2 & Chr(Rnd * 25 + 97)
   Next x
   For x = 1 To Len(txt2)
      n = n + Asc(Mid(txt2, x, 1))
   Next x
   For m = 1 To Len(CStr(n))
      txt2 = txt2 & Chr(Mid(n, m, 1) + 100)
   Next m
   Crypt = txt2
End Function

Function DeCrypt(Txt As String) As String
   l = Asc(Mid(Txt, 1, 1)) - 64
   For x = 1 To 50
      n = n + Asc(Mid(Txt, x, 1))
   Next x
   For m = 51 To Len(Txt)
      txt2 = txt2 & Asc(Mid(Txt, m, 1)) - 100
   Next m
   If n <> CInt(txt2) Then End
   Txt = Mid(Txt, 2, l)
   For n = 1 To l / 3
      For x = n To Len(Txt) Step l / 3
         w1 = w1 & Chr(Asc(Mid(Txt, x, 1)) - 1)
      Next x
   Next n
   DeCrypt = Trim(w1)
End Function
