VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5715
      Top             =   945
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   3555
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   529
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7964
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1501
            MinWidth        =   1501
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   2
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "08. 01. 01"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      ScaleHeight     =   690
      ScaleWidth      =   6420
      TabIndex        =   0
      Top             =   0
      Width           =   6480
      Begin VB.CommandButton Command4 
         Cancel          =   -1  'True
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   2670
         Picture         =   "MDIForm1.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Manager"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1755
         Picture         =   "MDIForm1.frx":0694
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Skip"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   870
         Picture         =   "MDIForm1.frx":099E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   885
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   0
         Picture         =   "MDIForm1.frx":0CA8
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   870
      End
      Begin VB.Image Image1 
         Height          =   660
         Left            =   60
         Top             =   15
         Width           =   6330
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Form1.MSFlexGrid1.Visible = Not Form1.MSFlexGrid1.Visible
   Form1.MSFlexGrid1.ZOrder 0
End Sub

Private Sub Command2_Click()
   Form4.Visible = True
   Form4.ZOrder 0
End Sub

Private Sub Command3_Click()
    Form2.Visible = True
    Form2.ZOrder 0
End Sub

Private Sub Command4_Click()
    End
End Sub

Private Sub MDIForm_Load()
    Me.Caption = "Watcher " & App.Major & "." & App.Minor
    Form1.Left = 0
    Form1.Top = 0
    Form1.Width = Me.Width - 200
    Form1.Height = Me.Height - Picture1.Height - 480
    Form1.Visible = True
    Call MDIForm_Resize
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   End
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    Form1.Left = 0
    Form1.Top = 0
    Form1.Width = Me.Width - 200
    Form1.Height = Me.Height - Picture1.Height - (480 + 290)
End Sub

Private Sub Timer1_Timer()
   StatusBar1.Panels(2).Text = Format(Time, "h:nn:ss")
End Sub
