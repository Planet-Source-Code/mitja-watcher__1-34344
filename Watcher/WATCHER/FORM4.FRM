VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ignored Programs"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   105
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   60
      Width           =   4110
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2850
      TabIndex        =   2
      Top             =   2550
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1455
      TabIndex        =   1
      Top             =   2550
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   2550
      Width           =   1395
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Dim txt As String
   txt = InputBox("Please enter string " & vbNewLine & "wich will be ignored.")
   If txt = "" Then Exit Sub
   Baza.Execute "INSERT INTO BanedWords (beseda) VALUES ('" & txt & "')"
   List1.AddItem txt
End Sub

Private Sub Command2_Click()
   If List1.Text = "" Then Exit Sub
   Baza.Execute "DELETE FROM BanedWords WHERE Beseda ='" & List1.Text & "'"
   List1.RemoveItem List1.ListIndex
End Sub

Private Sub Command3_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Dim rec As Recordset
   Set rec = Baza.OpenRecordset("SELECT * FROM BanedWords")
   Do Until rec.EOF
      List1.AddItem rec!beseda
      rec.MoveNext
   Loop
End Sub
