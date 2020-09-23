VERSION 5.00
Begin VB.Form Form9 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Software License Agreement"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "No"
      Height          =   345
      Left            =   4815
      TabIndex        =   2
      Top             =   3390
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Yes"
      Height          =   345
      Left            =   3585
      TabIndex        =   1
      Top             =   3390
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   3195
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form9.frx":0000
      Top             =   75
      Width           =   5940
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Index           =   1
      Left            =   4785
      TabIndex        =   5
      Top             =   3360
      Width           =   1155
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   405
      Index           =   0
      Left            =   3555
      TabIndex        =   4
      Top             =   3360
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Do you accept all these terms of the preceding License Agreement?"
      Height          =   390
      Left            =   120
      TabIndex        =   3
      Top             =   3360
      Width           =   3465
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Me.Visible = False
   Form7.Show
End Sub

Private Sub Command2_Click()
   End
End Sub
