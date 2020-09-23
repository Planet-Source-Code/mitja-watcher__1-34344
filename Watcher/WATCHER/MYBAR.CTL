VERSION 5.00
Begin VB.UserControl MyBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4650
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6525
   ScaleHeight     =   4650
   ScaleWidth      =   6525
   Begin VB.Frame Frame1 
      Height          =   3780
      Left            =   0
      TabIndex        =   0
      Top             =   -75
      Width           =   6495
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   1875
         Top             =   195
      End
      Begin VB.Label TollTipHelp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Toll Tip Help "
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   720
         TabIndex        =   5
         Top             =   390
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Shape XBar 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   2610
         Index           =   0
         Left            =   1140
         Shape           =   4  'Rounded Rectangle
         Top             =   915
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Line HLine 
         BorderColor     =   &H00808080&
         Index           =   0
         Visible         =   0   'False
         X1              =   585
         X2              =   6195
         Y1              =   3180
         Y2              =   3180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
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
         Left            =   60
         TabIndex        =   3
         Top             =   3105
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Title 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title"
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
         Left            =   1380
         TabIndex        =   2
         Top             =   3540
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   570
         X2              =   6375
         Y1              =   3510
         Y2              =   3510
      End
      Begin VB.Line VLine 
         BorderColor     =   &H00808080&
         Index           =   0
         Visible         =   0   'False
         X1              =   2940
         X2              =   2940
         Y1              =   225
         Y2              =   3495
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   570
         X2              =   570
         Y1              =   240
         Y2              =   3480
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2445
         Left            =   2670
         Picture         =   "MyBar.ctx":0000
         Stretch         =   -1  'True
         Top             =   270
         Width           =   3675
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   0
      TabIndex        =   1
      Top             =   3585
      Width           =   6510
      Begin VB.Label Caption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
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
         Left            =   180
         TabIndex        =   4
         Top             =   285
         Visible         =   0   'False
         Width           =   660
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   885
         Left            =   30
         Picture         =   "MyBar.ctx":4B146
         Stretch         =   -1  'True
         Top             =   150
         Width           =   6450
      End
   End
End
Attribute VB_Name = "MyBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim MaxValue As Double

Public Property Get Sections() As Variant
    Sections = Parram.AllSec
End Property

Public Property Let Sections(ByVal vNewValue As Variant)
   On Error GoTo CheckErr
   
   ReDim Preserve SecArray(vNewValue)
   Parram.AllSec = vNewValue
   
   Exit Property
CheckErr:
    MsgBox Error, vbInformation
End Property

Public Property Get Bar() As Variant
   Bar = Parram.CurBar
End Property

Public Property Let Bar(ByVal vNewValue As Variant)
   Parram.CurBar = vNewValue
End Property

Public Property Get Bars() As Variant
   Bars = UBound(SecArray(Parram.CurSec).TBar)
End Property

Public Property Let Bars(ByVal vNewValue As Variant)
   ReDim Preserve SecArray(Parram.CurSec).TBar(vNewValue)
   'Parram.AllBar = vNewValue
End Property

Public Property Get Value() As Variant
   Value = SecArray(Parram.CurSec).TBar(Parram.CurBar).Value
End Property

Public Property Let Value(ByVal vNewValue As Variant)
   SecArray(Parram.CurSec).TBar(Parram.CurBar).Value = vNewValue
End Property

Public Property Get Color() As Variant
   Color = SecArray(Parram.CurSec).TBar(Parram.CurBar).Color
   
End Property

Public Property Let Color(ByVal vNewValue As Variant)
   Dim X As Integer
   On Error GoTo 10
   For X = 0 To UBound(PomoznoPolje)
      If SecArray(Parram.CurSec).TBar(Parram.CurBar).Caption = PomoznoPolje(X).Caption Then
         PomoznoPolje(X).Color = vNewValue
         Exit For
      End If
   Next X
10   SecArray(Parram.CurSec).TBar(Parram.CurBar).Color = vNewValue
End Property

Public Property Get BackPicture() As Variant
   BackPicture = Image1.Picture
End Property

Public Property Let BackPicture(ByVal vNewValue As Variant)
   On Error GoTo napaka
   Image1.Picture = LoadPicture(vNewValue)
   Image2.Picture = LoadPicture(vNewValue)
   Exit Property
napaka:
   MsgBox Error, vbInformation
End Property

Public Property Get Section() As Variant
   Section = Parram.CurSec
End Property

Public Property Let Section(ByVal vNewValue As Variant)
   Parram.CurSec = vNewValue
End Property


Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim n As Integer
   On Error GoTo napaka
   X = X + Image1.Left
   For n = 1 To UBound(PosArray)
   If X >= PosArray(n).X1 And X <= PosArray(n).X2 Then
      If Y >= PosArray(n).Y1 - 300 And Y < PosArray(n).Y2 Then
         TollTipHelp.visible = False
         TollTipHelp.Caption = " " & XBar(n).Tag & " "
         TollTipHelp.Left = X + 100
         Do Until TollTipHelp.Left + TollTipHelp.Width < Image1.Width - 100
            TollTipHelp.Left = TollTipHelp.Left - 100
         Loop
         TollTipHelp.Top = Y - 100
         TollTipHelp.ZOrder 0
         TollTipHelp.visible = True
         Timer1.Enabled = True
         Exit Sub
      End If
   End If
   Next n
   Exit Sub
napaka:
   Exit Sub
End Sub

Private Sub Timer1_Timer()
   Timer1.Enabled = False
   TollTipHelp.visible = False
End Sub

Private Sub UserControl_Initialize()
    ReDim SecArray(0)
    ReDim SecArray(0).TBar(0)
    ReDim PomoznoPolje(0)
End Sub

Private Sub UserControl_Resize()
   Dim nbar As Integer, mainleft As Long
   Dim csec As Integer, Cbar As Integer
   On Error Resume Next
   Frame1.Left = 0
   Frame1.Top = -75
   Frame1.Height = UserControl.Height - 860
   Frame1.Width = UserControl.Width
    
   Image1.Left = 50
   Image1.Top = 120
   Image1.Width = Frame1.Width - 90
   Image1.Height = Frame1.Height - 150
   Image2.Left = Image1.Left
   Image2.Width = Image1.Width
   Frame2.Top = Frame1.Height - (100 * 2)
   Frame2.Width = Frame1.Width
   Frame2.Left = 0
   
   
   
   Line1.Y2 = Frame1.Height - 250
   Line2.Y1 = Line1.Y2
   Line2.Y2 = Line1.Y2
   Line2.X2 = Frame1.Width - 400
   HLine(0).X2 = Line2.X2
   VLine(0).Y1 = Line1.Y1
   VLine(0).Y2 = Line1.Y2
   Title(0).Top = Line2.Y1 + 20
   Frame1.Refresh
   Image1.Refresh
   Line2.ZOrder 0
   
   If XBar.Count = 1 Then Exit Sub
   MakeGrid
   mainleft = ((Line2.X2) / 2) - (((csec + Cbar) * 210) / 2)
   ReDim PosArray(0)
   For csec = 1 To UBound(SecArray)
      VLine(csec).X1 = mainleft + 105
      VLine(csec).X2 = mainleft + 105
      VLine(csec).Y1 = Line1.Y1
      VLine(csec).Y2 = Line1.Y2
      Title(csec).Caption = SecArray(csec).Title
      Title(csec).Left = mainleft + (UBound(SecArray(csec).TBar) * 105) - TextWidth(Title(csec).Caption) * 0.5 + 200
      Title(csec).Top = Line2.Y1 + 10
      mainleft = mainleft + 210
      For Cbar = 1 To UBound(SecArray(csec).TBar)
         nbar = nbar + 1
         XBar(nbar).Left = mainleft
         mainleft = mainleft + 210
         SetPosition nbar, SecArray(csec).TBar(Cbar).Value
         XBar(nbar).ZOrder 0
         ReDim Preserve PosArray(UBound(PosArray) + 1)
         PosArray(UBound(PosArray)).X1 = XBar(nbar).Left
         PosArray(UBound(PosArray)).X2 = XBar(nbar).Left + XBar(nbar).Width
         PosArray(UBound(PosArray)).Y1 = XBar(nbar).Top
         PosArray(UBound(PosArray)).Y2 = XBar(nbar).Top + XBar(nbar).Height
      Next Cbar
   Next csec
   VLine(csec).X1 = mainleft + 105
   VLine(csec).X2 = mainleft + 105
   VLine(csec).Y1 = Line1.Y1
   VLine(csec).Y2 = Line1.Y2
   On Error GoTo 0
End Sub

Public Property Get BarName() As Variant
   BarName = SecArray(Parram.CurSec).TBar(Parram.CurBar).Caption
End Property

Public Property Let BarName(ByVal vNewValue As Variant)
   Dim X As Integer
   SecArray(Parram.CurSec).TBar(Parram.CurBar).Caption = vNewValue
   For X = 1 To UBound(PomoznoPolje)
      If PomoznoPolje(X).Caption = vNewValue Then
         SecArray(Parram.CurSec).TBar(Parram.CurBar).Color = PomoznoPolje(X).Color
         Exit Property
      End If
   Next X
   ReDim Preserve PomoznoPolje(X)
   PomoznoPolje(X).Caption = vNewValue
End Property

Public Property Get SecName() As Variant
   SecName = SecArray(Parram.CurSec).Title
End Property

Public Property Let SecName(ByVal vNewValue As Variant)
   SecArray(Parram.CurSec).Title = vNewValue
End Property

Private Sub MakeVisible(visible As Boolean)
   Dim X As Integer
   For X = 1 To HLine.UBound
      HLine(X).ZOrder 0
      HLine(X).visible = visible
   Next X
   For X = 1 To VLine.UBound
      VLine(X).ZOrder 0
      VLine(X).visible = visible
   Next X
   For X = 1 To Label2.UBound
      Label2(X).ZOrder 0
      Label2(X).visible = visible
   Next X
   For X = 1 To Title.UBound
      Title(X).ZOrder 0
      Title(X).visible = visible
   Next X
   For X = 1 To XBar.UBound
      XBar(X).ZOrder 0
      XBar(X).visible = visible
   Next X
   Line2.ZOrder 0
End Sub

Private Sub DrawBars()
   Dim csec As Integer, Cbar As Integer
   Dim mainleft As Double, pristej As Integer
   Dim capleft As Integer, captop As Integer
   Dim X As Integer
   Count csec, Cbar
   UnloadObjects
   MakeGrid
   mainleft = ((Line2.X2) / 2) - (((csec + Cbar) * 210) / 2)
   ReDim PosArray(0)
'- Line2.X1
   On Error Resume Next
   For csec = 1 To UBound(SecArray)
      Load VLine(csec)
      VLine(csec).X1 = mainleft + 105
      VLine(csec).X2 = mainleft + 105
      VLine(csec).Y1 = Line1.Y1
      VLine(csec).Y2 = Line1.Y2
      Load Title(csec)
      Title(csec).Caption = SecArray(csec).Title
      Title(csec).Left = mainleft + (UBound(SecArray(csec).TBar) * 105) - TextWidth(Title(csec).Caption) * 0.5 + 200
      mainleft = mainleft + 210
      For Cbar = 1 To UBound(SecArray(csec).TBar)
         Load XBar(XBar.Count)
         XBar(XBar.UBound).Left = mainleft
         XBar(XBar.UBound).BackColor = SecArray(csec).TBar(Cbar).Color
         XBar(XBar.UBound).Tag = SecArray(csec).TBar(Cbar).Caption & " [" & MakeTime(SecArray(csec).TBar(Cbar).Value) & "]"
         mainleft = mainleft + 210
         SetPosition XBar.UBound, SecArray(csec).TBar(Cbar).Value
         ReDim Preserve PosArray(UBound(PosArray) + 1)
         PosArray(UBound(PosArray)).X1 = XBar(XBar.UBound).Left
         PosArray(UBound(PosArray)).X2 = XBar(XBar.UBound).Left + XBar(XBar.UBound).Width
         PosArray(UBound(PosArray)).Y1 = XBar(XBar.UBound).Top
         PosArray(UBound(PosArray)).Y2 = XBar(XBar.UBound).Top + XBar(XBar.UBound).Height
      Next Cbar
   Next csec
   Load VLine(csec)
   VLine(csec).X1 = mainleft + 105
   VLine(csec).X2 = mainleft + 105
   VLine(csec).Y1 = Line1.Y1
   VLine(csec).Y2 = Line1.Y2

   On Error Resume Next
   capleft = 250
   captop = 200
   For X = 1 To UBound(PomoznoPolje)
      Load Caption(X)
      Caption(X).Left = capleft
      Caption(X).Top = captop
      Caption(X).Caption = PomoznoPolje(X).Caption
      If pristej < TextWidth(PomoznoPolje(X).Caption) Then
         pristej = TextWidth(PomoznoPolje(X).Caption)
      End If
      Caption(X).ForeColor = PomoznoPolje(X).Color
      If Caption(X).Left + Caption(X).Width > Image2.Width Then
         Unload Caption(X)
         Exit Sub
      End If
      Caption(X).visible = True
      Caption(X).ZOrder 0
      captop = captop + 200
      If X / 4 = Fix(X / 4) Then
         captop = 200
         capleft = capleft + pristej + 650
         pristej = 0
      End If
   Next X
End Sub

Private Sub Count(csec As Integer, Cbar As Integer)
   Dim CCsec As Integer, CCbar As Integer
   MaxValue = 0
   For CCsec = 1 To UBound(SecArray)
      csec = csec + 1
      For CCbar = 1 To UBound(SecArray(CCsec).TBar)
         Cbar = Cbar + 1
         If SecArray(CCsec).TBar(CCbar).Value > MaxValue Then
            MaxValue = SecArray(CCsec).TBar(CCbar).Value
         End If
      Next CCbar
   Next CCsec
End Sub

Private Sub UnloadObjects()
   On Error Resume Next
   Dim X As Integer
   For X = 1 To XBar.UBound
      Unload XBar(X)
   Next X
   For X = 1 To HLine.UBound
      Unload HLine(X)
   Next X
   For X = 1 To VLine.UBound
      Unload VLine(X)
   Next X
   For X = 1 To Caption.UBound
      Unload Caption(X)
   Next X
   For X = 1 To Label2.UBound
      Unload Label2(X)
   Next X
   For X = 1 To Title.UBound
      Unload Title(X)
   Next X
End Sub

Private Sub SetPosition(object As Integer, Val As Long)
   Dim ret As Double
   ret = (Line1.Y2 - Line1.Y1) / MaxValue
   
   XBar(object).Top = Line1.Y2 - ret * Val
   XBar(object).Height = ret * Val
End Sub

Public Sub Draw()
   MakeVisible False
   DrawBars
   MakeVisible True
End Sub

Public Sub Reset()
   ReDim SecArray(0).TBar(0)
   ReDim PomoznoPolje(0)
   ReDim PosArray(0)
End Sub

Sub MakeGrid()
   Dim ras As Double, X As Integer
   Dim n As Integer, s As Integer
   
         
         If MaxValue > 240 Then
            n = 2
            s = 60
            Do Until n * s > MaxValue
               n = n + 1
            Loop
            MaxValue = n * s
         End If
         If MaxValue <= 240 And MaxValue > 180 Then MaxValue = 240: n = 8: s = 30
         If MaxValue <= 180 And MaxValue > 120 Then MaxValue = 180: n = 12: s = 15
         If MaxValue <= 120 And MaxValue > 90 Then MaxValue = 120: n = 12: s = 10
         If MaxValue <= 90 And MaxValue > 60 Then MaxValue = 90: n = 9: s = 10
         If MaxValue <= 60 And MaxValue > 30 Then MaxValue = 60: n = 12: s = 5
         If MaxValue <= 30 And MaxValue > 20 Then MaxValue = 30: n = 6: s = 5
         If MaxValue <= 20 And MaxValue > 10 Then MaxValue = 20: n = 10: s = 2
         If MaxValue <= 10 Then MaxValue = 10: n = 10: s = 1
         ras = (Line1.Y2 - Line1.Y1) / (n)
         On Error Resume Next
         For X = 0 To n
            Load HLine(X)
            HLine(X).X1 = Line2.X1
            HLine(X).X2 = Line2.X2
            HLine(X).Y1 = ras * X + Line1.Y1
            HLine(X).Y2 = ras * X + Line1.Y1
            HLine(X).ZOrder 0
            HLine(X).visible = True
            Load Label2(X)
            Label2(X).Top = ras * X - 100 + Line1.Y1
            Label2(X).Caption = MakeTime((n - X) * s)
            Label2(X).ZOrder 0
            Label2(X).visible = True
            'Debug.Print (n - X) * s, ras * X
         Next X
         Exit Sub
         
End Sub

Private Function MakeTime(txt As Variant) As String
   Dim h As Integer, m As Integer
   h = Fix(txt / 60)
   m = txt - h * 60
   MakeTime = h & ":" & Format(m, "00")
End Function
