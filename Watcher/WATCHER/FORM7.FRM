VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registration"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4785
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2835
      MaxLength       =   4
      TabIndex        =   2
      Top             =   45
      Width           =   480
   End
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
      Height          =   315
      Left            =   3315
      TabIndex        =   5
      Top             =   420
      Width           =   1380
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
      Height          =   315
      Left            =   1875
      TabIndex        =   4
      Top             =   420
      Width           =   1380
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3465
      MaxLength       =   12
      TabIndex        =   3
      Top             =   45
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1875
      MaxLength       =   7
      TabIndex        =   1
      Top             =   45
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   3330
      TabIndex        =   7
      Top             =   30
      Width           =   120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   2715
      TabIndex        =   6
      Top             =   30
      Width           =   120
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Registration Key:"
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
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   1785
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Command1_Click()
   If Text1 = "7791011" Then
      m = (Text2 - 121) / 7
      Licences = m
      For x = 1 To Len(Text3) - 2 Step 2
         n = n + CLng((Mid(Text3, x, 2)))
      Next x
      'n = Fix(n / 6)
      If Fix(n / 2) <> Int(Mid(Text3, x)) Then
         MsgBox "Invalid licence number.", vbExclamation
         End
      End If
      For x = 1 To Len(Text3) - 2 Step 2
         Char = Chr(Mid(Text3, x, 2) + 92)
         If InStr(1, "abcdefghijklmnoprstuvzxywq", Char) = 0 Then
            MsgBox "Invalid licence number.", vbExclamation
            End
         Else
            licenceto = licenceto & Chr(Mid(Text3, x, 2) + 92)
         End If
      Next x
      
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
      'Print #1, Crypt("SerialNumber=" & (App.Path))
      Print #1, Crypt("NumberOfLicences=" & Licences)
      Print #1, Crypt("LicenceTo=" & licenceto)
      Print #1, Crypt("CompanyName=SoftTech-si")
      Print #1, Crypt("UserName=" & GetUserName)
      Print #1, Crypt("ComputerName=" & GetComputerName)
      Print #1, Crypt("Datum=" & Format(Date, "dd.mm.yyyy"))
      Print #1, "-End-"
      Close
      Unload Me
      Main
   Else
      MsgBox "Invalid licence number.", vbExclamation
      End
   End If
End Sub

Private Sub Command2_Click()
   End
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Text3.SetFocus
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then Call Command1_Click
End Sub
