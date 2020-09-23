VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Computer Manager"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command9 
      Caption         =   "Close"
      Height          =   330
      Left            =   3900
      TabIndex        =   14
      Top             =   4800
      Width           =   945
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show All"
      Height          =   330
      Left            =   2955
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   885
   End
   Begin VB.Frame Frame2 
      Height          =   4650
      Left            =   2355
      TabIndex        =   1
      Top             =   90
      Width           =   2520
      Begin VB.CommandButton Command4 
         Caption         =   "Add"
         Height          =   255
         Left            =   105
         TabIndex        =   8
         Top             =   3795
         Width           =   690
      End
      Begin VB.ListBox List3 
         Height          =   1620
         Left            =   120
         TabIndex        =   4
         Top             =   2130
         Width           =   2265
      End
      Begin VB.ListBox List2 
         Height          =   1425
         Left            =   120
         TabIndex        =   3
         Top             =   690
         Width           =   2265
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Modify"
         Height          =   255
         Left            =   780
         TabIndex        =   9
         Top             =   3795
         Width           =   900
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Remove"
         Height          =   255
         Left            =   1665
         TabIndex        =   10
         Top             =   3795
         Width           =   750
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Add to Group"
         Height          =   270
         Left            =   105
         TabIndex        =   11
         Top             =   4035
         Width           =   2310
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Remove from Group"
         Height          =   270
         Left            =   105
         TabIndex        =   12
         Top             =   4290
         Width           =   2310
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Computers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   675
         TabIndex        =   16
         Top             =   285
         Width           =   1380
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   120
         Picture         =   "Form2.frx":030A
         Top             =   195
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4650
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   2325
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   285
         Left            =   105
         TabIndex        =   5
         Top             =   4260
         Width           =   660
      End
      Begin VB.ListBox List1 
         Height          =   3375
         Left            =   120
         TabIndex        =   2
         Top             =   690
         Width           =   2115
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Modify"
         Height          =   285
         Left            =   750
         TabIndex        =   6
         Top             =   4260
         Width           =   720
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Remove"
         Height          =   285
         Left            =   1455
         TabIndex        =   7
         Top             =   4260
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Groups"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   810
         TabIndex        =   15
         Top             =   255
         Width           =   930
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   195
         Picture         =   "Form2.frx":0614
         Top             =   165
         Width           =   480
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WhoHadFocus As String

Private Sub Check1_Click()
    Dim rec As Recordset
    List3.Clear
    If Check1.Value = 0 Then
        Set rec = Baza.OpenRecordset("SELECT * FROM users WHERE ingroup IS Null")
    Else
        Set rec = Baza.OpenRecordset("SELECT * FROM users")
    End If
    Do Until rec.EOF
        List3.AddItem rec!Name
        rec.MoveNext
    Loop
End Sub

Private Sub Command1_Click()
    Dim rec As Recordset
    Dim txt As String
    
    txt = InputBox("Enter Group Name:", "Watcher")
    Set rec = Baza.OpenRecordset("SELECT * FROM Groups WHERE Name = '" & txt & "'")
    If rec.EOF Then
        Baza.Execute "INSERT INTO Groups (Name) VALUES ('" & txt & "')"
        List1.AddItem txt
    Else
        MsgBox "Group with name " & txt & " exist.", vbExclamation
    End If
End Sub

Private Sub Command2_Click()
    Dim rec As Recordset, txt As String
    
    Set rec = Baza.OpenRecordset("SELECT * FROM Groups WHERE Name = '" & List1.Text & "'")
    If rec.EOF Then Exit Sub
    txt = InputBox("Enter New Group Name:", "Watcher", List1.Text)
    Baza.Execute "UPDATE Groups SET Name = '" & txt & "' WHERE ID = " & rec!id
End Sub

Private Sub Command3_Click()
    Dim x As Integer, rec As Recordset
    If List1.Text = "" Then Exit Sub
    x = MsgBox("Are you sure you want to delete group " & List1.Text & "?", vbQuestion + vbYesNo + vbDefaultButton2)
    If x = vbYes Then
        Set rec = Baza.OpenRecordset("SELECT * FROM Groups WHERE Name = '" & List1.Text & "'")
        Baza.Execute "UPDATE Users SET inGroup = Null WHERE inGroup = '" & rec!id & "'"
        Baza.Execute "DELETE FROM Groups WHERE ID = " & rec!id
        List1.RemoveItem List1.ListIndex
    End If
End Sub
    
Private Sub Command4_Click()
    Dim rec As Recordset
    Set rec = Baza.OpenRecordset("SELECT Count(*) FROM users")
    If rec!Expr1000 > Licences Then
      MsgBox "You riched maximum clients." & vbNewLine & "Please contact your local retailer.", vbInformation
      Exit Sub
    End If
    Form3.Visible = True
    Form3.Caption = "Add New"
    Form3.ZOrder 0
End Sub

Private Sub Command5_Click()
    Dim rec As Recordset
    
    Form3.Caption = "Modify"
    If WhoHadFocus = "List2" Then
      Set rec = Baza.OpenRecordset("SELECT * FROM USERS WHERE name = '" & List2.Text & "'")
    Else
      Set rec = Baza.OpenRecordset("SELECT * FROM USERS WHERE name = '" & List3.Text & "'")
    End If
    If rec.EOF = True Then
         Form3.Visible = False
         Exit Sub
    End If
    Form3.Tag = rec!Name
    Form3.Text1 = rec!Name
    Form3.Text2 = rec!IP
    Form3.Visible = True
    Form3.ZOrder 0
   
End Sub

Private Sub Command6_Click()
    Dim x As Integer, txt As String, n As Integer
    If WhoHadFocus = "List2" Then
        txt = List2.Text
        If txt = "" Then Exit Sub
        n = MsgBox("Are you sure you want delete " & txt & "?", vbQuestion + vbYesNo)
        If n = vbNo Then Exit Sub
        List2.RemoveItem List2.ListIndex
        For x = 0 To List3.ListCount - 1
            If List3.List(x) = txt Then
                List3.RemoveItem x
                Exit For
            End If
        Next x
    Else
        txt = List3.Text
        If txt = "" Then Exit Sub
        n = MsgBox("Are you sure you want delete " & txt & "?", vbQuestion + vbYesNo)
        If n = vbNo Then Exit Sub
        List3.RemoveItem List3.ListIndex
        For x = 0 To List2.ListCount - 1
            If List2.List(x) = txt Then
                List2.RemoveItem x
                Exit For
            End If
        Next x
    End If
    Baza.Execute "DELETE FROM users WHERE name = '" & txt & "'"
End Sub

Private Sub Command7_Click()
    Dim rec As Recordset
    Set rec = Baza.OpenRecordset("SELECT * FROM groups WHERE name = '" & List1.Text & "'")
    If rec.EOF Then Exit Sub
    Baza.Execute "UPDATE users SET ingroup = " & rec!id & " WHERE name = '" & List3.Text & "'"
    List2.AddItem List3.Text
    If Check1.Value = 0 Then List3.RemoveItem List3.ListIndex
End Sub

Private Sub Command8_Click()
    If WhoHadFocus = "List2" Then
      Baza.Execute "UPDATE users SET ingroup = null WHERE name = '" & List2.Text & "'"
      List2.RemoveItem List2.ListIndex
    Else
      Baza.Execute "UPDATE users SET ingroup = null WHERE name = '" & List3.Text & "'"
    End If
End Sub

Private Sub Command9_Click()
    Call Form1.Form_Load
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rec As Recordset
    Set rec = Baza.OpenRecordset("SELECT * FROM groups")
    Do Until rec.EOF
        List1.AddItem rec!Name
        rec.MoveNext
    Loop
    Set rec = Baza.OpenRecordset("SELECT * FROM users WHERE ingroup = Null")
    Do Until rec.EOF
        List3.AddItem rec!Name
        rec.MoveNext
    Loop
    If List1.ListCount > 0 Then
        List1.ListIndex = 0
        Call List1_Click
    End If
End Sub

Private Sub List1_Click()
    Dim rec As Recordset
    List2.Clear
    Set rec = Baza.OpenRecordset("SELECT * FROM Groups WHERE name = '" & List1.Text & "'")
    Set rec = Baza.OpenRecordset("SELECT * FROM Users WHERE InGroup = '" & rec!id & "'")
    Do Until rec.EOF
        List2.AddItem rec!Name
        rec.MoveNext
    Loop

End Sub

Private Sub List2_Click()
    Dim x As Integer
    If List2.ListIndex < 0 Then Exit Sub
    WhoHadFocus = "List2"
    x = List2.ListIndex
    If List3.ListIndex >= 0 Then
      List3.Selected(List3.ListIndex) = False
    End If
    List2.Selected(x) = True
    Command7.Enabled = False
End Sub

Private Sub List3_Click()
    Dim x As Integer
    If List3.ListIndex < 0 Then Exit Sub
    WhoHadFocus = "List3"
    x = List3.ListIndex
    If List2.ListIndex >= 0 Then
      List2.Selected(List2.ListIndex) = False
    End If
    List3.Selected(x) = True
    Command7.Enabled = True
    Command8.Enabled = True
End Sub
