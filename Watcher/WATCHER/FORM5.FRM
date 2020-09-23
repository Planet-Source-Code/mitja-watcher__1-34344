VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Master Config"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   4935
      Left            =   45
      TabIndex        =   1
      Top             =   1950
      Width           =   8325
      Begin VB.CommandButton Command6 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2625
         TabIndex        =   26
         Top             =   225
         Width           =   885
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Manual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1350
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1200
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Auto Run"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Value           =   -1  'True
         Width           =   1200
      End
      Begin VB.TextBox Text2 
         Height          =   300
         Left            =   3975
         TabIndex        =   4
         Top             =   210
         Width           =   1635
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   3900
         Left            =   120
         TabIndex        =   3
         Top             =   630
         Width           =   8085
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   105
         TabIndex        =   2
         Top             =   4560
         Width           =   8115
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Using Local Connection"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5220
         TabIndex        =   16
         Top             =   255
         Width           =   2910
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IP :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3585
         TabIndex        =   5
         Top             =   225
         Width           =   345
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1980
      Left            =   60
      TabIndex        =   0
      Top             =   -30
      Width           =   8295
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   6255
         Top             =   1095
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Test"
         Height          =   285
         Left            =   4185
         TabIndex        =   27
         Top             =   1035
         Width           =   1020
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Run debug"
         Height          =   300
         Left            =   6765
         TabIndex        =   25
         Top             =   480
         Width           =   1395
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   1
         Left            =   4740
         TabIndex        =   24
         Text            =   "30010"
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Index           =   0
         Left            =   4740
         TabIndex        =   22
         Text            =   "30000"
         Top             =   270
         Width           =   735
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Open Connection"
         Height          =   300
         Left            =   6765
         TabIndex        =   20
         Top             =   840
         Width           =   1410
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   7800
         Top             =   150
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Load DataBase"
         Height          =   300
         Left            =   6780
         TabIndex        =   19
         Top             =   1200
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Close"
         Height          =   300
         Left            =   6780
         TabIndex        =   18
         Top             =   1560
         Width           =   1410
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   270
         Left            =   2550
         TabIndex        =   17
         Top             =   1635
         Width           =   1035
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   2535
         TabIndex        =   14
         Top             =   1290
         Width           =   1020
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1575
         TabIndex        =   12
         Top             =   945
         Width           =   2430
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1575
         TabIndex        =   10
         Top             =   600
         Width           =   2430
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1575
         TabIndex        =   8
         Top             =   255
         Width           =   2430
      End
      Begin VB.Label Label5 
         Height          =   210
         Left            =   4230
         TabIndex        =   28
         Top             =   1620
         Width           =   2385
      End
      Begin VB.Label Label3 
         Caption         =   "RP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   4290
         TabIndex        =   23
         Top             =   615
         Width           =   405
      End
      Begin VB.Label Label3 
         Caption         =   "LP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   4305
         TabIndex        =   21
         Top             =   285
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Number Of Licences"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   15
         Top             =   1305
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   4050
         X2              =   4050
         Y1              =   225
         Y2              =   1890
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Company Name:"
         Height          =   195
         Index           =   2
         Left            =   285
         TabIndex        =   13
         Top             =   975
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Registration Key:"
         Height          =   195
         Index           =   1
         Left            =   255
         TabIndex        =   11
         Top             =   630
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Registration Code:"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   9
         Top             =   285
         Width           =   1305
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Open App.Path & "\Info.dat" For Output As #1
      Print #1, "--------------------------------------------------------------"
      Print #1, "              INFORMATION FILE by SoftTech-Si                 "
      Print #1, "--------------------------------------------------------------"
      Print #1, "Any Modifycation in this file my couse serius program failure."
      Print #1, "                                                              "
      Print #1, "!Do Not Change Anything Below This Line!                      "
      Print #1, ""
      Print #1, "-Begin-                                                       "
      Print #1, Crypt("ProductName=" & App.ProductName)
      'Print #1, Crypt("SerialNumber=" & GetSerial(App.Path))
      Print #1, Crypt("NumberOfLicences=" & Text6)
      Print #1, Crypt("LicenceTo=" & Text5)
      Print #1, Crypt("CompanyName=SoftTech-si")
      Print #1, Crypt("UserName=" & GetUserName)
      Print #1, Crypt("ComputerName=" & GetComputerName)
      Print #1, Crypt("LocalIP=" & Winsock1.LocalIP)
      Print #1, "-End-"
   Close
End Sub

Private Sub Command2_Click()
   End
End Sub

Private Sub Command4_Click()
'   If Form1.Winsock1.State = 1 Then
'      Form1.Winsock1.Close
'      MDIForm1.Visible = False
'   End If
   Winsock1.Protocol = sckUDPProtocol
   Winsock1.RemotePort = Text7(1)
   Winsock1.LocalPort = Text7(0)
   Winsock1.RemoteHost = Winsock1.LocalIP
   Winsock1.SendData ""
End Sub

Private Sub Command5_Click()
   Winsock1.Close
   Set WrkSpace = DBEngine.CreateWorkspace("Watcher", "admin", "")
   Set Baza = WrkSpace.OpenDatabase(App.Path & "\Data.mdb", False, Flase, ";PWD=enigma")
   MDIForm1.Visible = True
   Label4.Caption = "Main Connection"
End Sub

Private Sub Command6_Click()
   List1.Clear
End Sub

Private Sub Command7_Click()
   sexes = 0
   Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Form_Load()
   Dim Txt As String, cmd As String, con As String
   DebugLevel = 1
   On Error Resume Next
   Open App.Path & "\Info.dat" For Input As #1
   If Err > 0 Then Exit Sub
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
         Case "NumberOfLicences"
            Text6 = con
         Case "LicenceTo"
            Text5 = con
      End Select
   Loop
   Close
End Sub


Private Sub Option1_Click()
   If Option2.Value = True Then
      DebugLevel = 2
   Else
      DebugLevel = 1
   End If
End Sub

Private Sub Option2_Click()
   If Option2.Value = True Then
      DebugLevel = 2
   Else
      DebugLevel = 1
   End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
   Dim n As Boolean
   On Error GoTo Napaka
   If KeyAscii = 13 Then
      Winsock1.RemoteHost = Text2
      List1.AddItem "" 'Text1
      If Winsock1.State = 0 Then
         If IsError(n = Baza.Updatable) Then Exit Sub
         Form1.Winsock1.SendData Text1
      Else
         Winsock1.SendData Text1
      End If
      Text1 = ""
      List1.ListIndex = List1.ListCount - 1
   End If
   Exit Sub
Napaka:
   MsgBox Error, vbCritical
End Sub

Private Sub Text5_Change()
   CalcSerial
End Sub

Private Sub Text6_Change()
   CalcSerial
End Sub

Private Sub Timer1_Timer()
   Static trys As Integer
   Form1.Timer1.Enabled = True
   Label5 = "Ratio " & trys & " / " & sexes
   trys = trys + 1
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
   Dim Txt As String
   Winsock1.GetData Txt
   List1.AddItem Winsock1.RemoteHost & ": " & Txt
   List1.ListIndex = List1.ListCount - 1
End Sub

Sub CalcSerial()
   Dim Txt As String
   Txt = "7791011"
   Txt = Txt & "-" & Format(Val(Text6) * 7 + 121, "0000") & "-"
   For x = 1 To Len(Text5)
      Char = Char + Asc(Mid(Text5, x, 1)) - 92
      Txt = Txt & Format(Asc(Mid(Text5, x, 1)) - 92, "00")
   Next x
   Txt = Txt & Int(Char / 2)
   Text4 = Txt
End Sub
