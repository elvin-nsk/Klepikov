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

Public ImageAndContour As ShapeRange
Public Contour As Shape
Public CropBoxOffsetBottom As Double
Public CropBoxOffsetLeft As Double
Public CropBoxOffsetRight As Double
Public CropBoxOffsetTop As Double
Public ContourValid As Boolean
Public Image As ShapeRange
Public ImageValid As Boolean

'===============================================================================

Public Property Get Self() As Parts
    Set Self = Me
End Property
