VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{A459B50D-BEA5-11D3-BB52-DC244387C843}#2.0#0"; "DBAR.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7905
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1020
      Top             =   4395
   End
   Begin Project2.MyBar MyBar1 
      Height          =   3345
      Left            =   30
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   135
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   5900
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   555
      Top             =   4395
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3480
      Left            =   45
      TabIndex        =   8
      Top             =   105
      Visible         =   0   'False
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   6138
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "<Date                    |<Product Name                 |<Title                          |<Duration     "
   End
   Begin VB.Frame Frame1 
      Height          =   4140
      Left            =   5625
      TabIndex        =   0
      Top             =   15
      Width           =   2220
      Begin ComctlLib.TreeView TreeView1 
         Height          =   3825
         Left            =   60
         TabIndex        =   2
         Top             =   210
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   6747
         _Version        =   327682
         Style           =   5
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame3 
      Height          =   900
      Left            =   5625
      TabIndex        =   3
      Top             =   4035
      Width           =   2220
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   510
         Width           =   1140
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   165
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date To:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   510
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Date From:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   165
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Height          =   900
      Left            =   0
      TabIndex        =   1
      Top             =   4035
      Width           =   5655
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   105
         ScaleHeight     =   195
         ScaleWidth      =   5370
         TabIndex        =   9
         Top             =   510
         Width           =   5430
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   195
         Index           =   2
         Left            =   2685
         TabIndex        =   13
         Top             =   255
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   195
         Index           =   1
         Left            =   5385
         TabIndex        =   12
         Top             =   225
         Width           =   90
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "?"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   11
         Top             =   225
         Width           =   90
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4575
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":0000
            Key             =   "uporabnik"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":031A
            Key             =   "skupina"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Poskusi As Integer

Sub Form_Load()
    Dim rec As Recordset, nodx As Node, n As Integer
    TreeView1.Nodes.Clear
    Set rec = Baza.OpenRecordset("SELECT * FROM Groups")
    Do Until rec.EOF
        TreeView1.Nodes.Add , , "g" & rec!id, rec!Name, "skupina"
        rec.MoveNext
    Loop
    
    Set rec = Baza.OpenRecordset("SELECT * FROM users")
    Do Until rec.EOF
        If IsNull(rec!inGroup) Then
            Set nodx = TreeView1.Nodes.Add(, , "u" & rec!id, rec!Name, "uporabnik")
            nodx.Tag = rec!IP
        Else
            Set nodx = TreeView1.Nodes.Add("g" & rec!inGroup, 4, "u" & rec!id, rec!Name, "uporabnik")
            nodx.Tag = rec!IP
        End If
        rec.MoveNext
    Loop
   MSFlexGrid1.ColWidth(0) = GetSetting("Watcher", "MsFlexGrid", "Row1", 1400)
   MSFlexGrid1.ColWidth(1) = GetSetting("Watcher", "MsFlexGrid", "Row2", 1875)
   MSFlexGrid1.ColWidth(2) = GetSetting("Watcher", "MsFlexGrid", "Row3", 1560)
   MSFlexGrid1.ColWidth(3) = GetSetting("Watcher", "MsFlexGrid", "Row4", 915)
   
   On Error Resume Next
   Form1.Winsock1.Protocol = sckUDPProtocol
   Form1.Winsock1.RemotePort = 30010
   Form1.Winsock1.LocalPort = 30000
   Form1.Winsock1.SendData ""

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    MyBar1.Visible = False
    MSFlexGrid1.Width = Me.Width - (Frame2.Left + Frame1.Width) 'Me.Width - Frame1.Width - (3 * MSFlexGrid1.Left)
    MSFlexGrid1.Height = Me.Height - Frame2.Height - (3 * MSFlexGrid1.Top) + 250
    
    Frame1.Left = MSFlexGrid1.Left + MSFlexGrid1.Width
    Frame3.Left = Frame1.Left
    
    Frame1.Height = Me.Height - Frame3.Height + 50
    Frame3.Top = Frame1.Height + Frame1.Top - 120
    TreeView1.Height = Frame1.Height - TreeView1.Top * 2
    Frame2.Left = MSFlexGrid1.Left
    Frame2.Width = Me.Width - (Frame2.Left + Frame3.Width - 30)
    Frame2.Top = Frame3.Top
    Frame2.Height = Frame3.Height

    MyBar1.Left = MSFlexGrid1.Left
    MyBar1.Top = MSFlexGrid1.Top
    MyBar1.Width = MSFlexGrid1.Width
    MyBar1.Height = MSFlexGrid1.Height
    
    Picture1.Width = Frame2.Width - Picture1.Left * 2
    Label3(0).Left = Picture1.Left
    Label3(1).Left = Picture1.Left + Picture1.Width - Label3(1).Width
    Label3(2).Left = Picture1.Left + Picture1.Width / 2 + Label3(2).Width / 2
    MyBar1.Visible = True
End Sub

Private Sub MSFlexGrid1_LostFocus()
   SaveSetting "Watcher", "MsFlexGrid", "Row1", MSFlexGrid1.ColWidth(0)
   SaveSetting "Watcher", "MsFlexGrid", "Row2", MSFlexGrid1.ColWidth(1)
   SaveSetting "Watcher", "MsFlexGrid", "Row3", MSFlexGrid1.ColWidth(2)
   SaveSetting "Watcher", "MsFlexGrid", "Row4", MSFlexGrid1.ColWidth(3)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    Picture1.ToolTipText = MakeTime(x / Picture1.Tag)
    On Error GoTo 0
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsDate(Text1) = False And Text1 <> "" Then
         Text1 = ""
         MsgBox "Please Enter Date.", vbExclamation
      Else
         Timer1.Enabled = True
      End If
   End If
End Sub

Private Sub Text1_LostFocus()
   If IsDate(Text1) = False And Text1 <> "" Then
      Text1 = ""
      MsgBox "Please Enter Date.", vbExclamation
   End If
End Sub

Private Sub Timer1_Timer()
   On Error GoTo fuckyou
   If Timer1.Tag <> "" Then
      Winsock1.RemoteHost = Timer1.Tag
      'MDIForm1.StatusBar1.Panels(1).Text = "Connecting....."
      Send "Hello"
      'Timer1.Enabled = False
      Timer1.Interval = 2000
      Retrys = Retrys + 1
      MDIForm1.StatusBar1.Panels(1).Text = "Connecting.....(" & Retrys & ")"
      If Poskusi > 5 Then
         Timer1.Enabled = False
         MDIForm1.StatusBar1.Panels(1).Text = "Can't connect to computer."
      End If
   End If
   Exit Sub
fuckyou:
'   MsgBox Error, , "Writer: Error"
End Sub

Private Sub TreeView1_BeforeLabelEdit(Cancel As Integer)
   Cancel = True
End Sub

Private Sub TreeView1_Click()
   If TreeView1.SelectedItem.Tag <> "" Then
      Timer1.Enabled = False
      Poskusi = 0
      Timer1.Tag = TreeView1.SelectedItem.Tag
      Timer1.Enabled = True
      MDIForm1.Caption = "DEMO Watcher 1.0 - " & TreeView1.SelectedItem.Text
   End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
   Dim Txt As String, x As Integer, n As Integer
   Dim Datum As Date, ime As String
   Dim LineFrom As Double, LineTo As Double
   Dim w As Double, d As Long, s As Integer
   
   On Error GoTo Napaka
   If Winsock1.RemoteHost <> Timer1.Tag Then Exit Sub
   
   Select Case MDIForm1.StatusBar1.Tag
      Case 1
         MDIForm1.StatusBar1.Panels(1).Text = "Transfering data... /"
         MDIForm1.StatusBar1.Tag = 2
      Case 2
         MDIForm1.StatusBar1.Panels(1).Text = "Transfering data... --"
         MDIForm1.StatusBar1.Tag = 3
      Case 3
         MDIForm1.StatusBar1.Panels(1).Text = "Transfering data... \"
         MDIForm1.StatusBar1.Tag = 4
      Case 4
         MDIForm1.StatusBar1.Panels(1).Text = "Transfering data... |"
         MDIForm1.StatusBar1.Tag = 1
      Case Else
         MDIForm1.StatusBar1.Panels(1).Text = "Transfering data... /"
         MDIForm1.StatusBar1.Tag = 2
   End Select
   Winsock1.GetData Txt
   If DebugLevel >= 1 Then
      Form5.List1.AddItem ":--> " & Txt
      Form5.List1.ListIndex = Form5.List1.ListCount - 1
   End If
   Debug.Print Txt
   Select Case Txt
      Case "Hello"
         ReDim Prog(0)
         ReDim Tipke(0)
         MSFlexGrid1.Rows = 1
         Picture1.Cls
         DoEvents
         Send "StartSendConfig"
         
      Case "SendConfigStarted"
         If Text1.Text = "" Then
            Send "Today"
         Else
            Datum = Text1
            Send "DateFrom: " & Year(Datum) & "-" & Format(Month(Datum), "00") & "-" & Format(Day(Datum), "00")
         End If
      
      Case "DateTo?"
         If Text2 = "" Then
            Datum = Date
            Send "DateTo: " & Year(Datum) & "-" & Format(Month(Datum), "00") & "-" & Format(Day(Datum), "00")
         Else
            Datum = Text2
            Send "DateTo: " & Year(Datum) & "-" & Format(Month(Datum), "00") & "-" & Format(Day(Datum), "00")
         End If
               
      Case "DataReady"
         Send "StartSend"
         Timer1.Enabled = False
         Timer1.Interval = 500
         sexes = sexes + 1
         
      Case "Done"
         MDIForm1.StatusBar1.Panels(1).Text = "Drawing graf..."
         MDIForm1.StatusBar1.Refresh
         MyBar1.Reset
         MyBar1.Section = 0
         MyBar1.Sections = 0
         MSFlexGrid1.Redraw = False
         For x = 1 To UBound(Prog)
            If Ispusti(Prog(x).ime) = False Then
               If Prog(x).Date <> Datum Then
                  Datum = Prog(x).Date
                  MyBar1.Sections = MyBar1.Sections + 1
                  MyBar1.Section = MyBar1.Sections
                  MyBar1.Bars = 0
                  MyBar1.SecName = Prog(x).Date
                  ime = ""
               End If
               If Trim(Prog(x).ime) <> Trim(ime) Then
                  ime = Prog(x).ime
                  MyBar1.Bars = MyBar1.Bars + 1
                  MyBar1.Bar = MyBar1.Bars
                  MyBar1.BarName = Prog(x).ime
                  If MyBar1.Color = 0 Then
                     MyBar1.Color = RGB((Rnd * 25) * 10, (Rnd * 25) * 10, (Rnd * 25) * 10)
                  End If
                  MyBar1.Value = Prog(x).Duration
               Else
                  MyBar1.Value = MyBar1.Value + Prog(x).Duration
               End If
               MSFlexGrid1.AddItem Prog(x).Date & vbTab & Prog(x).ime & vbTab & Prog(x).Title & vbTab & Prog(x).Duration
               MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
               For n = 0 To MSFlexGrid1.Cols - 1
                  MSFlexGrid1.Col = n
                  MSFlexGrid1.CellForeColor = MyBar1.Color
               Next n
            End If
         Next x
         If MSFlexGrid1.Rows > 1 Then
            MSFlexGrid1.Row = 1
            MSFlexGrid1.Col = 0
         End If
         MSFlexGrid1.Redraw = True
         MyBar1.Draw
         If UBound(Tipke) > 0 Then
            d = DateDiff("n", Tipke(1).Date, DateAdd("n", Tipke(UBound(Tipke)).Duration, Tipke(UBound(Tipke)).Date))
            If d > 0 Then
               w = Picture1.Width / d
               Picture1.Tag = w
               For x = 1 To UBound(Tipke)
                  LineFrom = DateDiff("n", Tipke(1).Date, Tipke(x).Date) * w
                  LineTo = DateDiff("n", Tipke(1).Date, DateAdd("n", Tipke(x).Duration, Tipke(x).Date)) * w
                  If Tipke(x).State = False Then
                     Picture1.Line (LineFrom, 0)-(LineTo, Picture1.Height), RGB(200, 0, 0), BF
                  Else
                     Picture1.Line (LineFrom, 0)-(LineTo, Picture1.Height), RGB(0, 200, 0), BF
                  End If
               Next x
            End If
            s = 1
            If d > 10 Then s = 1
            If d > 20 Then s = 2
            If d > 60 Then s = 5
            If d > 120 Then s = 10
            If d > 180 Then s = 15
            If d > 360 Then s = 30
            If d > 1200 Then
               s = 60
               Do Until s * 12 > d
                  s = s + 60
               Loop
            End If
            
            For x = 1 To d Step s
               Picture1.Line ((x * w) - 10, 0)-((x * w) + 10, Picture1.Height), RGB(200, 200, 200), BF
            Next x
            Label3(0).Caption = Format(Tipke(1).Date, "dd.mm. h:nn")
            Label3(1).Caption = Format(DateAdd("n", Tipke(UBound(Tipke)).Duration, Tipke(UBound(Tipke)).Date), "dd.mm. h:nn")
            If s > 60 Then
               Label3(2).Caption = "1.block = " & s / 60 & "h"
            Else
               Label3(2).Caption = "1.block = " & s & "min"
            End If
            Label3(0).Left = Picture1.Left
            Label3(1).Left = Picture1.Left + Picture1.Width - Label3(1).Width
            Label3(2).Left = Picture1.Left + Picture1.Width / 2 - Label3(2).Width / 2
         End If
         MDIForm1.StatusBar1.Panels(1).Text = "Ready"
      
      Case Else
         If Mid(Txt, 1, 4) = "Prog" Then
            ReDim Preserve Prog(UBound(Prog) + 1)
            Prog(UBound(Prog)).ime = GetProgName(Txt)
            Prog(UBound(Prog)).Title = GetProgTitle(Txt)
            Prog(UBound(Prog)).Duration = GetProgDuration(Txt)
            Prog(UBound(Prog)).Date = GetProgDate(Txt)
            Send "Next"
         
         ElseIf Mid(Txt, 1, 4) = "Keys" Then
            ReDim Preserve Tipke(UBound(Tipke) + 1)
            Tipke(UBound(Tipke)).Date = GetKeyDate(Txt)
            Tipke(UBound(Tipke)).State = GetKeyState(Txt)
            Tipke(UBound(Tipke)).Duration = GetKeyDuration(Txt)
            Send "Next"
         End If
   End Select
   Exit Sub
Napaka:
   MsgBox Error, , "Watcher: Error"
End Sub

Private Function GetProgName(Txt As String) As String
   Dim n As Integer, m As Integer
   n = InStr(1, Txt, "Word1:") + 6
   m = InStr(n, Txt, ":Word2:") - n
   GetProgName = Mid(Txt, n, m)
End Function

Function GetProgTitle(Txt As String) As String
   Dim n As Integer, m As Integer
   n = InStr(1, Txt, "Word2:") + 6
   m = InStr(n, Txt, ":Word3:") - n
   GetProgTitle = Mid(Txt, n, m)
End Function
Function GetProgDuration(Txt As String) As String
   Dim n As Integer, m As Integer
   n = InStr(1, Txt, "Word4:") + 6
   m = InStr(n, Txt, ":") - n
   GetProgDuration = Mid(Txt, n, m)
End Function
Function GetProgDate(Txt As String) As String
   Dim n As Integer, m As Integer
   n = InStr(1, Txt, "Word3:") + 6
   m = InStr(n, Txt, ":Word4:") - n
   GetProgDate = Mid(Txt, n, m)
End Function

Function GetKeyState(Txt As String) As String
   Dim n As Integer, m As Integer
   n = InStr(1, Txt, "Word1:") + 6
   m = InStr(n, Txt, ":Word2:") - n
   GetKeyState = Mid(Txt, n, m)
End Function
Function GetKeyDate(Txt As String) As String
   Dim n As Integer, m As Integer
   n = InStr(1, Txt, "Word2:") + 6
   m = InStr(n, Txt, ":Word3:") - n
   GetKeyDate = Mid(Txt, n, m)
End Function
Function GetKeyDuration(Txt As String) As String
   Dim n As Integer, m As Integer
   n = InStr(1, Txt, "Word3:") + 6
   m = InStr(n, Txt, ":") - n
   GetKeyDuration = Mid(Txt, n, m)
End Function
Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      If IsDate(Text2) = False And Text2 <> "" Then
         Text2 = ""
         MsgBox "Please Enter Date.", vbExclamation
      Else
         Timer1.Enabled = True
      End If
   End If
End Sub

Private Sub Text2_LostFocus()
   If IsDate(Text2) = False And Text2 <> "" Then
      Text2 = ""
      MsgBox "Please Enter Date.", vbExclamation
   End If
End Sub

Function Ispusti(Txt As String) As Boolean
   Dim rec As Recordset
   Set rec = Baza.OpenRecordset("SELECT * FROM BanedWords")
   Do Until rec.EOF
      If InStr(1, LCase(Txt), LCase(rec!beseda)) > 0 Then
         Ispusti = True
         Exit Function
      End If
      rec.MoveNext
   Loop
End Function

Sub Send(Txt As String)
   If DebugLevel < 2 Then Winsock1.SendData Txt
   If DebugLevel > 0 Then
      Form5.List1.AddItem ":<-- " & Txt
      Form5.List1.ListIndex = Form5.List1.ListCount - 1
   End If
   Debug.Print Txt
End Sub
Function MakeTime(Txt As Double) As String
    'h = Fix(txt / 60)
    'm = Fix(txt - h * 60)
    MakeTime = Format(DateAdd("n", Fix(Txt), Tipke(1).Date), "h:nn")
End Function
