Attribute VB_Name = "Klepikov"
'===============================================================================
'   Макрос          : Klepikov
'   Версия          : 2024.05.15
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "Klepikov"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_VERSION As String = "2024.05.15"

'===============================================================================
' # Globals

Private Const MAX_WIDTH As Double = 339

'===============================================================================
' # Entry points

Sub SpreadForCutter()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
    
    Dim Source As ShapeRange
    With InputData.RequestShapes
        If .IsError Then Exit Sub
        Set Source = .Shapes
    End With
       
    ActiveDocument.Unit = cdrMillimeter
    Dim SourceParts As Parts: Set SourceParts = GatherParts(Source)
    If Not CheckParts(SourceParts) Then Exit Sub
    
    Dim Cfg As New Config
    If Not ShowSpreadForCutter(SourceParts, Cfg) Then Exit Sub

    BoostStart "Расклад для реза"
       
    Impose(SourceParts, Cfg).CreateSelection
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================
' # Helpers

Public Property Get CalcPlaces( _
                         ByVal Parts As Parts, _
                         ByVal LeftOffset As Double, _
                         ByVal RightOffset As Double, _
                         ByVal SpreadDistance As Double _
                     ) As Long
    Dim MaxBoxWidth As Double: MaxBoxWidth = _
        MAX_WIDTH - LeftOffset - RightOffset
    Dim ContoursDistance As Double: ContoursDistance = _
        ShapesDistance(Parts, SpreadDistance) _
      + Parts.ImageAndContour.SizeWidth - Parts.Contour.SizeWidth
    CalcPlaces = _
        VBA.Fix( _
            (MaxBoxWidth + ContoursDistance) _
          / (Parts.Contour.SizeWidth + ContoursDistance) _
        )
End Property

Public Function Impose( _
               ByVal Parts As Parts, _
               ByVal Cfg As Config _
           ) As ShapeRange
    With Parts
        Set Impose = CreateShapeRange
        Dim Source As ShapeRange: Set Source = .ImageAndContour.Duplicate
        Dim StartingX As Double, StartingY As Double
        StartingX = _
            .ImageAndContour.RightX _
          + .ImageAndContour.SizeWidth * 0.5 _
          + Cfg.LeftOffset
        StartingY = .ImageAndContour.BottomY - Cfg.BottomOffset
        Dim Count As Long: Count = _
            CalcPlaces( _
                .Self, Cfg.LeftOffset, Cfg.RightOffset, Cfg.SpreadDistance _
            )
        Dim Distance  As Double: Distance = _
            ShapesDistance(.Self, Cfg.SpreadDistance)
        Source.LeftX = _
            StartingX + Cfg.LeftOffset - .ContourOffsetLeft - Distance
        Source.BottomY = _
            StartingY + Cfg.BottomOffset - .ContourOffsetBottom
        Dim i As Long
        For i = 1 To Count
            Impose.AddRange _
                Source.Duplicate((i - 1) * (Source.SizeWidth + Distance))
        Next i
        Impose.Add CreateImpositionRect(Impose, Parts, Cfg)
        Source.Delete
    End With
End Function

Private Function CreateImpositionRect( _
                     ByVal Imposition As ShapeRange, _
                     ByVal Parts As Parts, _
                     ByVal Cfg As Config _
                 ) As Shape
    Set CreateImpositionRect = _
        ActiveLayer.CreateRectangle( _
            Imposition.LeftX + Parts.ContourOffsetLeft - Cfg.LeftOffset, _
            Imposition.TopY - Parts.ContourOffsetTop + Cfg.TopOffset, _
            Imposition.RightX - Parts.ContourOffsetRight + Cfg.RightOffset, _
            Imposition.BottomY + Parts.ContourOffsetBottom - Cfg.BottomOffset _
        )
    CreateImpositionRect.Name = "область раскладки"
    CreateImpositionRect.OrderBackOf Imposition.FirstShape
End Function
 
Private Property Get CheckParts(ByVal Parts As Parts) As Boolean
    With Parts
        If Not .ContourValid Then
            VBA.MsgBox "Не найден контур", vbExclamation
            Exit Property
        End If
        If Not .ImageValid Then
            VBA.MsgBox "Не найдено изображение", vbExclamation
            Exit Property
        End If
    End With
    CheckParts = True
End Property

Private Property Get GatherParts(ByVal Shapes As ShapeRange) As Parts
    With New Parts
        Set GatherParts = .Self
        
        Set .Contour = Shapes.Shapes.FindShape(Type:=cdrCurveShape)
        .ContourValid = Not .Contour Is Nothing
        If Not .ContourValid Then Exit Property
        .ContourOffsetBottom = .Contour.BottomY - Shapes.BottomY
        .ContourOffsetLeft = .Contour.LeftX - Shapes.LeftX
        .ContourOffsetRight = Shapes.RightX - .Contour.RightX
        .ContourOffsetTop = Shapes.TopY - .Contour.TopY
        Set .ImageAndContour = CreateShapeRange
        .ImageAndContour.AddRange Shapes
        Set .Image = CreateShapeRange
        .Image.AddRange Shapes
        .Image.RemoveRange PackShapes(.Contour)
        .ImageValid = (.Image.Count > 0)
    End With
End Property

Private Property Get ShapesDistance( _
                         ByVal Parts As Parts, _
                         ByVal SpreadDistance As Double _
                     ) As Double
    ShapesDistance = _
        Parts.Image.SizeWidth - Parts.ImageAndContour.SizeWidth + SpreadDistance
End Property

Private Function ShowSpreadForCutter( _
                         ByVal Parts As Parts, _
                         ByVal Cfg As Config _
                     ) As Boolean
    With New SpreadForCutterView
        .TopOffset = Cfg.TopOffset
        .LeftOffset = Cfg.LeftOffset
        .RightOffset = Cfg.RightOffset
        .BottomOffset = Cfg.BottomOffset
        .SpreadDistance = Cfg.SpreadDistance
        
        'для вызова калькулятора из формы
        Set .Parts = Parts
        
        .Show vbModal
        
        Cfg.TopOffset = .TopOffset
        Cfg.LeftOffset = .LeftOffset
        Cfg.RightOffset = .RightOffset
        Cfg.BottomOffset = .BottomOffset
        Cfg.SpreadDistance = .SpreadDistance
        
        ShowSpreadForCutter = .IsOk
    End With
End Function

'===============================================================================
' # Tests

Private Sub testSomething()
'
End Sub
