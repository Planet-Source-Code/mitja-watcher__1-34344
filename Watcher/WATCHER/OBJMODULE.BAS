Attribute VB_Name = "Module1"
Type ParramType
    CurSec As Integer
    CurBar As Integer
    AllSec As Integer
    AllBar As Integer
End Type

Type BarArray
    Color As Long
    Value As Long
    Caption As String
End Type

Public Type SectionType
    Title As String
    TBar() As BarArray
End Type

Type PosArray
   X1 As Long
   X2 As Long
   Y1 As Long
   Y2 As Long
End Type

Type PomoznoPolje
   Caption As String
   Color As Long
End Type

Public Parram As ParramType
Public SecArray() As SectionType
Public PomoznoPolje() As PomoznoPolje
Public PosArray() As PosArray
