VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add User"
   ClientHeight    =   1395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2775
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1395
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   75
      TabIndex        =   4
      Top             =   1050
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
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
      Height          =   270
      Left            =   1455
      TabIndex        =   3
      Top             =   1050
      Width           =   1275
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1245
      TabIndex        =   2
      Top             =   630
      Width           =   1470
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1260
      TabIndex        =   1
      Text            =   "127.0.0.1"
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   240
      Left            =   1260
      TabIndex        =   0
      Top             =   90
      Width           =   1440
   End
   Begin VB.Label Label1 
      Caption         =   "Group:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   105
      TabIndex        =   7
      Top             =   660
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "IP or DNS:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   390
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   105
      TabIndex        =   5
      Top             =   105
      Width           =   855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim rec As Recordset
    
    If Text1 = "" Then MsgBox "Please enter computer name.", vbExclamation: Exit Sub
    If Text2 = "" Then MsgBox "Please enter computer IP or DNS.", vbExclamation: Exit Sub
    If Me.Caption = "Modify" Then
      Set rec = Baza.OpenRecordset("SELECT * FROM Users WHERE Name = '" & Text1 & "'")
      If rec.EOF = False Then
         If rec!Name = Text1 And rec!Name <> Me.Tag Then
            MsgBox "Computer with name " & Text1 & " allready exist.", vbExclamation
            Exit Sub
         End If
      End If
      If Combo1.Text = "" Then
          Baza.Execute "UPDATE Users SET Name = '" & Text1 & "',IP='" & Text2 & "' WHERE Name = '" & Me.Tag & "'"
      Else
          Set rec = Baza.OpenRecordset("SELECT * FROM Groups WHERE Name = '" & Combo1.Text & "'")
          Baza.Execute "UPDATE Users SET (Name = '" & Text1 & "',IP='" & Text2 & "',InGroup = '" & rec!id & "') WHERE Name = '" & Me.Tag & "'"
      End If
      If Form2.WhoHadFocus = "List2" Then
         Form2.List2.RemoveItem Form2.List2.ListIndex
         Form2.List2.AddItem Text1
      Else
         Form2.List3.RemoveItem Form2.List3.ListIndex
         Form2.List3.AddItem Text1
      End If
   Else
      Set rec = Baza.OpenRecordset("SELECT * FROM Users WHERE Name = '" & Text1 & "'")
      If rec.EOF Then
          If Combo1.Text = "" Then
              Baza.Execute "INSERT INTO Users (Name,IP) VALUES ('" & Text1 & "','" & Text2 & "')"
          Else
              Set rec = Baza.OpenRecordset("SELECT * FROM Groups WHERE Name = '" & Combo1.Text & "'")
              Baza.Execute "INSERT INTO Users (Name,IP,InGroup) VALUES ('" & Text1 & "','" & Text2 & "','" & rec!id & "')"
          End If
          Form2.List3.AddItem Txt
          Unload Me
      Else
          MsgBox "Computer with name " & Text1 & " allready exist.", vbExclamation
      End If
   End If
   Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rec As Recordset
    Text1 = ""
    Text2 = "127.0.0.1"
    Combo1 = ""
    Set rec = Baza.OpenRecordset("SELECT * FROM Groups")
    Do Until rec.EOF
        Combo1.AddItem rec!Name
        rec.MoveNext
    Loop
End Sub
