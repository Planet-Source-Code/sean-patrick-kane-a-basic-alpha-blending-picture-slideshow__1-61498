Attribute VB_Name = "AlphaAPI"
Option Explicit

'Mouse functions
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long

'These are API functions and variables responsible for the alpha blending in frmSlideshow
Public Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)

Public Const AC_SRC_OVER = &H0
Public Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

Public Function Pause(interval As Integer)
Dim X As Long
    X = Timer
    Do While Timer - X < Val(interval)
        DoEvents
    Loop
End Function

