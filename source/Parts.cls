VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Parts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Struct
Option Explicit

'===============================================================================

Public AllShapes As ShapeRange
Public Contour As Shape
Public ContourOffsetBottom As Double
Public ContourOffsetLeft As Double
Public ContourOffsetRight As Double
Public ContourOffsetTop As Double
Public ContourOk As Boolean
Public Image As ShapeRange
Public ImageOffsetBottom As Double
Public ImageOffsetLeft As Double
Public ImageOk As Boolean

'===============================================================================

Public Property Get Self() As Parts
    Set Self = Me
End Property
