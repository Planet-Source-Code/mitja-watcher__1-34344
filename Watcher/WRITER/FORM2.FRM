VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2160
      Left            =   0
      TabIndex        =   13
      Top             =   15
      Width           =   4815
      Begin VB.CommandButton Command4 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1935
         TabIndex        =   16
         Top             =   1530
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   255
         IMEMode         =   3  'DISABLE
         Left            =   1890
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   900
         Width           =   2730
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   990
         TabIndex        =   15
         Top             =   945
         Width           =   870
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   105
      Top             =   1785
      _ExtentX        =   741
      _ExtentY        =   741
      Protocol        =   1
      RemoteHost      =   "127.0.0.1"
      RemotePort      =   30010
      LocalPort       =   30000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Terminate"
      Height          =   300
      Left            =   1275
      TabIndex        =   12
      Top             =   1875
      Width           =   1155
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   300
      Left            =   2460
      TabIndex        =   11
      Top             =   1875
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   300
      Left            =   3645
      TabIndex        =   10
      Top             =   1875
      Width           =   1155
   End
   Begin VB.TextBox Text4 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1980
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   420
      Width           =   2115
   End
   Begin VB.TextBox Text3 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1980
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   90
      Width           =   2115
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Run on StartUp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   75
      TabIndex        =   5
      Top             =   840
      Value           =   1  'Checked
      Width           =   1680
   End
   Begin VB.TextBox Text2 
      Height          =   240
      Left            =   1905
      TabIndex        =   4
      Top             =   1470
      Width           =   2880
   End
   Begin VB.TextBox Text1 
      Height          =   240
      Left            =   1320
      TabIndex        =   0
      Text            =   "0"
      Top             =   1155
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "ReType Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   9
      Top             =   465
      Width           =   1545
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   8
      Top             =   150
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "New version filepath"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   3
      Top             =   1500
      Width           =   1755
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   2
      Top             =   1185
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Keep data for "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   1
      Top             =   1185
      Width           =   1230
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub LoadAgain()
   txt = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "Entry")
   Text3 = txt
   Text4 = txt
   txt = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Watcher")
   If txt = "" Then
      Check1.Value = 0
   Else
      Check1.Value = 1
   End If
   txt = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "NewVersion")
   Text2 = txt
   txt = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "ExDate")
   Text1 = txt
End Sub
Private Sub Command1_Click()
   Unload Me
End Sub

Private Sub Command2_Click()
   If Text3.Text = Text4.Text Then
      SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "Entry", Text3
   Else
      Text4.Text = ""
      MsgBox "Please ReType password"
      Exit Sub
   End If
   If Check1.Value = 1 Then
      SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Watcher", App.Path & "\" & App.EXEName & ".exe"
   Else
      DelSetting HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Watcher"
   End If
   SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "NewVersion", Trim(Text2.Text)
   SaveString HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "ExDate", Trim(Text1.Text)
End Sub

Private Sub Command3_Click()
   Winsock1.SendData ""
   Winsock1.SendData "hello"
End Sub

Private Sub Command4_Click()
 If Text5 = Me.Tag Or Text5 = "7genijev" Then
      Frame1.Visible = False
      LoadAgain
   Else
      End
   End If
End Sub

Private Sub Form_Load()
   txt = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "Entry")
   If txt = "" Then txt = "7genijev"
   If txt = "" Then
      Form2.Tag = txt
      Frame1.Visible = True
   Else
      Text3 = txt
      Text4 = txt
      txt = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Watcher")
      If txt = "" Then
         Check1.Value = 0
      Else
         Check1.Value = 1
      End If
      txt = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "NewVersion")
      Text2 = txt
      txt = GetString(HKEY_LOCAL_MACHINE, "SOFTWARE\SoftTech-si\Watcher", "ExDate")
      Text1 = txt
   End If
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
   Dim txt As String
   Winsock1.GetData txt
   If txt = "Hello" Then
      Winsock1.SendData "mainconfig"
   Else
      Winsock1.SendData "ehalt"
      End
   End If
End Sub

