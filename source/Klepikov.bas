Attribute VB_Name = "Klepikov"
'===============================================================================
'   Макрос          : Klepikov
'   Версия          : 2024.06.03
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

'===============================================================================
' # Manifest

Public Const APP_NAME As String = "Klepikov"
Public Const APP_DISPLAYNAME As String = APP_NAME
Public Const APP_VERSION As String = "2024.06.03"

'===============================================================================
' # Globals

Public Const SHAPES_WIDTH_TO_OFFSET_MULT As Double = 1.5
Public Const ADDITIONAL_OFFSET As Double = 80
Public Const ADDITIONAL_SPREAD_DISTANCE As Double = 2

'===============================================================================
' # Entry points

Sub SpreadForCutter()

    #If DebugMode = 0 Then
    On Error GoTo Catch
    #End If
    
    Dim Source As ShapeRange
    With InputData.RequestShapes
        If .IsError Then Exit Sub
        Set Source = .Shapes.UngroupAllEx
    End With
       
    ActiveDocument.Unit = cdrMillimeter
    Dim SourceParts As Parts: Set SourceParts = GatherParts(Source)
    If Not CheckParts(SourceParts) Then Exit Sub
    
    Dim Cfg As New Config
    If Not ShowSpreadForCutter(SourceParts, Cfg) Then Exit Sub

    BoostStart "Расклад для реза"
       
    ImposeAll(SourceParts, Cfg).CreateSelection
    
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
                         ByVal SpreadDistance As Double, _
                         ByVal MaxSheetWidth As Double _
                    ) As Long
    Dim MaxUsableSheetWidth As Double: MaxUsableSheetWidth = _
        MaxSheetWidth - LeftOffset - RightOffset
    CalcPlaces = _
        VBA.Fix( _
            (MaxUsableSheetWidth + SpreadDistance) _
          / (Parts.Image.SizeWidth + SpreadDistance) _
        )
End Property

Private Property Get ShapesDistance( _
                         ByVal Parts As Parts, _
                         ByVal SpreadDistance As Double _
                     ) As Double
    ShapesDistance = _
        Parts.Image.SizeWidth - Parts.ImageAndContour.SizeWidth + SpreadDistance
End Property

Private Function ImposeAll( _
                     ByVal Parts As Parts, _
                     ByVal Cfg As Config _
                 ) As ShapeRange
    With Parts
        Set ImposeAll = CreateShapeRange
        Dim Places As Long: Places = _
            CalcPlaces( _
                .Self, Cfg.LeftOffset, Cfg.RightOffset, _
                Cfg.SpreadDistance, Cfg.MaxSheetWidth _
            )
        Dim Distance As Double: Distance = _
            ShapesDistance(.Self, Cfg.SpreadDistance)
        Dim StartingOffsetX As Double: StartingOffsetX = _
            .ImageAndContour.SizeWidth * SHAPES_WIDTH_TO_OFFSET_MULT _
          + .ImageAndContour.LeftX + StartingOffsetX _
          + Cfg.LeftOffset - .CropBoxOffsetLeft
        If Places > 0 Then
            ImposeAll.AddRange _
                Impose(.Self, Places, Distance, StartingOffsetX, Cfg)
            ImposeAll.Add CreateSheetRect(ImposeAll, .Self, Cfg)
        End If
        If Cfg.AdditionalPlaces > 0 Then
            If ImposeAll.Count = 0 Then
                StartingOffsetX = _
                    .ImageAndContour.RightX _
                  + .ImageAndContour.SizeWidth * SHAPES_WIDTH_TO_OFFSET_MULT
            Else
                StartingOffsetX = ImposeAll.RightX + ADDITIONAL_OFFSET
            End If
            ImposeAll.AddRange _
                Impose(.Self, Cfg.AdditionalPlaces, Distance, StartingOffsetX, Cfg)
        End If
        If Cfg.DeleteImages Then _
            ImposeAll.Shapes.FindShapes(Type:=cdrBitmapShape).Delete
    End With
End Function

Private Function Impose( _
                     ByVal Parts As Parts, _
                     ByVal Places As Long, _
                     ByVal Distance As Double, _
                     ByVal StartingOffsetX As Double, _
                     ByVal Cfg As Config _
                 ) As ShapeRange
    With Parts
        Set Impose = CreateShapeRange
        Dim Source As ShapeRange: Set Source = .ImageAndContour.Duplicate
        Source.LeftX = StartingOffsetX
        Dim i As Long
        For i = 1 To Places
            Impose.AddRange _
                Source.Duplicate((i - 1) * (Source.SizeWidth + Distance))
        Next i
        Source.Delete
    End With
End Function

Private Function CreateSheetRect( _
                     ByVal Imposition As ShapeRange, _
                     ByVal Parts As Parts, _
                     ByVal Cfg As Config _
                 ) As Shape
    Set CreateSheetRect = _
        ActiveLayer.CreateRectangle( _
            Imposition.LeftX + Parts.CropBoxOffsetLeft - Cfg.LeftOffset, _
            Imposition.TopY + Parts.CropBoxOffsetTop + Cfg.TopOffset, _
            Imposition.RightX + Parts.CropBoxOffsetRight + Cfg.RightOffset, _
            Imposition.BottomY + Parts.CropBoxOffsetBottom - Cfg.BottomOffset _
        )
    CreateSheetRect.Name = "область раскладки"
    CreateSheetRect.OrderBackOf GetBottomOrderShape(Imposition)
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
        Set .ImageAndContour = CreateShapeRange
        .ImageAndContour.AddRange Shapes
        Set .Image = CreateShapeRange
        .Image.AddRange Shapes
        .Image.RemoveRange PackShapes(.Contour)
        .ImageValid = (.Image.Count > 0)
        If Not .ImageValid Then Exit Property
        
        .CropBoxOffsetBottom = .Image.BottomY - Shapes.BottomY
        .CropBoxOffsetLeft = .Image.LeftX - Shapes.LeftX
        .CropBoxOffsetRight = .Image.RightX - Shapes.RightX
        .CropBoxOffsetTop = .Image.TopY - Shapes.TopY
    End With
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
        .MaxSheetWidth = Cfg.MaxSheetWidth
        .AdditionalPlaces = Cfg.AdditionalPlaces
        .cbDeleteImages = Cfg.DeleteImages
        
        'для вызова калькулятора из формы
        Set .Parts = Parts
        
        .Show vbModal
        
        Cfg.TopOffset = .TopOffset
        Cfg.LeftOffset = .LeftOffset
        Cfg.RightOffset = .RightOffset
        Cfg.BottomOffset = .BottomOffset
        Cfg.SpreadDistance = .SpreadDistance
        Cfg.MaxSheetWidth = .MaxSheetWidth
        Cfg.AdditionalPlaces = .AdditionalPlaces
        Cfg.DeleteImages = .cbDeleteImages
        
        ShowSpreadForCutter = .IsOk
    End With
End Function

'===============================================================================
' # Tests

Private Sub testSomething()
'
End Sub
