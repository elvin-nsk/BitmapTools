Attribute VB_Name = "lib_elvin"
'===============================================================================
'   Модуль          : lib_elvin
'   Версия          : 2022.03.27
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'   Использован код : dizzy (из макроса CtC), Alex Vakulenko
'                     и др.
'   Описание        : библиотека функций для макросов от elvin-nsk
'   Использование   :
'   Зависимости     : самодостаточный
'===============================================================================

Option Explicit

'===============================================================================
' # приватные переменные модуля
'===============================================================================

Private Type typeLayerProps
    Visible As Boolean
    Printable As Boolean
    Editable As Boolean
End Type

Private StartTime As Double

'===============================================================================
' публичные переменные
'===============================================================================

Public Const CustomError = vbObjectError Or 32

Public Type typeMatrix
    d11 As Double
    d12 As Double
    d21 As Double
    d22 As Double
    tx As Double
    ty As Double
End Type

'===============================================================================
' функции общего назначения
'===============================================================================

'-------------------------------------------------------------------------------
' Функции           : BoostStart, BoostFinish
' Версия            : 2020.04.30
' Авторы            : dizzy, elvin-nsk
' Назначение        : доработанные оптимизаторы от CtC
' Зависимости       : самодостаточные
'
' Параметры:
' ~~~~~~~~~~
'
'
' Использование:
' ~~~~~~~~~~~~~~
'
'-------------------------------------------------------------------------------
Public Sub BoostStart( _
               Optional ByVal UnDo As String = "", _
               Optional ByVal Optimize As Boolean = True _
           )
    If Not UnDo = "" And Not ActiveDocument Is Nothing Then _
        ActiveDocument.BeginCommandGroup UnDo
    If Optimize Then Optimization = True
    EventsEnabled = False
    If Not ActiveDocument Is Nothing Then
        With ActiveDocument
            .SaveSettings
            '.PreserveSelection = False
            .Unit = cdrMillimeter
            .WorldScale = 1
            .ReferencePoint = cdrCenter
        End With
    End If
End Sub
Public Sub BoostFinish(Optional ByVal EndUndoGroup As Boolean = True)
    EventsEnabled = True
    Optimization = False
    If Not ActiveDocument Is Nothing Then
        With ActiveDocument
            .RestoreSettings
            '.PreserveSelection = True
            If EndUndoGroup Then .EndCommandGroup
        End With
        ActiveWindow.Refresh
    End If
    Application.Windows.Refresh
End Sub

'===============================================================================
' функции манипуляций с объектами корела
'===============================================================================

'все объекты на всех страницах, включая мастер-страницу - на один слой
'все страницы прибиваются, все объекты на слоях guides прибиваются
Public Function FlattenPagesToLayer(ByVal LayerName As String) As Layer

    Dim DL As Layer: Set DL = ActiveDocument.MasterPage.DesktopLayer
    Dim DLstate As Boolean: DLstate = DL.Editable
    Dim P As Page
    Dim L As Layer
    
    DL.Editable = False
    
    For Each P In ActiveDocument.Pages
        For Each L In P.Layers
            If L.IsSpecialLayer Then
                L.Shapes.All.Delete
            Else
                L.Activate
                L.Editable = True
                With L.Shapes.All
                    .MoveToLayer DL
                    .OrderToBack
                End With
                L.Delete
            End If
        Next
        If P.Index <> 1 Then P.Delete
    Next
    
    Set FlattenPagesToLayer = ActiveDocument.Pages.First.CreateLayer(LayerName)
    FlattenPagesToLayer.MoveBelow ActiveDocument.Pages.First.GuidesLayer
    
    For Each L In ActiveDocument.MasterPage.Layers
        If Not L.IsSpecialLayer Or L.IsDesktopLayer Then
            L.Activate
            L.Editable = True
            With L.Shapes.All
                .MoveToLayer FlattenPagesToLayer
                .OrderToBack
            End With
            If Not L.IsSpecialLayer Then L.Delete
        Else
            L.Shapes.All.Delete
        End If
    Next
    
    FlattenPagesToLayer.Activate
    DL.Editable = DLstate

End Function

'правильно перемещает Shape или ShapeRange на другой слой
Public Function MoveToLayer(ByVal ShapeOrRange As Object, ByVal Layer As Layer)
    
    Dim tSrcLayer() As Layer
    Dim tProps() As typeLayerProps
    Dim tLayersCol As Collection
    Dim i&
    
    If TypeOf ShapeOrRange Is Shape Then
    
        Set tLayersCol = New Collection
        tLayersCol.Add ShapeOrRange.Layer
        
    ElseIf TypeOf ShapeOrRange Is ShapeRange Then
        
        If ShapeOrRange.Count < 1 Then Exit Function
        Set tLayersCol = ShapeRangeLayers(ShapeOrRange)
        
    Else
    
        Err.Raise 13, Source:="MoveToLayer", _
                  Description:="Type mismatch: ShapeOrRange должен быть Shape или ShapeRange"
        Exit Function
    
    End If
    
    ReDim tSrcLayer(1 To tLayersCol.Count)
    ReDim tProps(1 To tLayersCol.Count)
    For i = 1 To tLayersCol.Count
        Set tSrcLayer(i) = tLayersCol(i)
        LayerPropsPreserveAndReset tSrcLayer(i), tProps(i)
    Next i
    ShapeOrRange.MoveToLayer Layer
    For i = 1 To tLayersCol.Count
        LayerPropsRestore tSrcLayer(i), tProps(i)
    Next i

End Function

'правильно копирует Shape или ShapeRange на другой слой
Public Function CopyToLayer(ByVal ShapeOrRange As Object, ByVal Layer As Layer) As Object

    If Not TypeOf ShapeOrRange Is Shape And Not TypeOf ShapeOrRange Is ShapeRange Then
        Err.Raise 13, Source:="CopyToLayer", _
                  Description:="Type mismatch: ShapeOrRange должен быть Shape или ShapeRange"
        Exit Function
    End If
    
    Set CopyToLayer = ShapeOrRange.Duplicate
    MoveToLayer CopyToLayer, Layer

End Function

'дублировать активную страницу со всеми слоями и объектами
Public Function DuplicateActivePage( _
                    ByVal NumberOfPages As Long, _
                    Optional ByVal ExcludeLayerName As String = "" _
                ) As Page
    Dim tRange As ShapeRange
    Dim tShape As Shape, sDuplicate As Shape
    Dim tProps As typeLayerProps
    Dim i&
    For i = 1 To NumberOfPages
        Set tRange = FindShapesActivePageLayers
        Set DuplicateActivePage = _
            ActiveDocument.InsertPages(1, False, ActivePage.Index)
        DuplicateActivePage.SizeHeight = ActivePage.SizeHeight
        DuplicateActivePage.SizeWidth = ActivePage.SizeWidth
        For Each tShape In tRange.ReverseRange
            If tShape.Layer.Name <> ExcludeLayerName Then
                LayerPropsPreserveAndReset tShape.Layer, tProps
                Set sDuplicate = tShape.Duplicate
                sDuplicate.MoveToLayer _
                    FindLayerDuplicate(DuplicateActivePage, tShape.Layer)
                LayerPropsRestore tShape.Layer, tProps
            End If
        Next tShape
    Next i
End Function

'перекрашивает объект в чёрный или белый в серой шкале,
'в зависимости от исходного цвета
'ДОРАБОТАТЬ
Public Function ContrastShape(ByVal Shape As Shape) As Shape
    With Shape.Fill
        Select Case .Type
            Case cdrUniformFill
                .UniformColor.ConvertToGray
                If .UniformColor.Gray < 128 Then _
                    .UniformColor.GrayAssign 0 Else .UniformColor.GrayAssign 255
            Case cdrFountainFill
                'todo
        End Select
    End With
    With Shape.Outline
        If Not .Type = cdrNoOutline Then
            .Color.ConvertToGray
            If .Color.Gray < 128 Then _
                .Color.GrayAssign 0 Else .Color.GrayAssign 255
        End If
    End With
    Set ContrastShape = Shape
End Function

'обрезать битмап по CropEnvelopeShape, но по-умному,
'сначала кропнув на EXPANDBY пикселей побольше
Public Function TrimBitmap( _
                    ByVal BitmapShape As Shape, _
                    ByVal CropEnvelopeShape As Shape, _
                    Optional ByVal LeaveCropEnvelope As Boolean = True _
                ) As Shape

    Const EXPANDBY& = 2 'px
    
    Dim tCrop As Shape
    Dim tPxW#, tPxH#
    Dim tSaveUnit As cdrUnit

    If Not BitmapShape.Type = cdrBitmapShape Then Exit Function
    
    'save
    tSaveUnit = ActiveDocument.Unit
    
    ActiveDocument.Unit = cdrInch
    tPxW = 1 / BitmapShape.Bitmap.ResolutionX
    tPxH = 1 / BitmapShape.Bitmap.ResolutionY
    BitmapShape.Bitmap.ResetCropEnvelope
    Set tCrop = BitmapShape.Layer.CreateRectangle( _
                    CropEnvelopeShape.LeftX - tPxW * EXPANDBY, _
                    CropEnvelopeShape.TopY + tPxH * EXPANDBY, _
                    CropEnvelopeShape.RightX + tPxW * EXPANDBY, _
                    CropEnvelopeShape.BottomY - tPxH * EXPANDBY _
                )
    Set TrimBitmap = Intersect(tCrop, BitmapShape, False, False)
    If TrimBitmap Is Nothing Then
        tCrop.Delete
        GoTo Finally
    End If
    TrimBitmap.Bitmap.Crop
    Set TrimBitmap = Intersect(CropEnvelopeShape, TrimBitmap, LeaveCropEnvelope, False)
    
Finally:
    'restore
    ActiveDocument.Unit = tSaveUnit
    
End Function

'правильный интерсект
Public Function Intersect( _
                    ByVal SourceShape As Shape, _
                    ByVal TargetShape As Shape, _
                    Optional ByVal LeaveSource As Boolean = True, _
                    Optional ByVal LeaveTarget As Boolean = True _
                ) As Shape
                                     
    Dim tPropsSource As typeLayerProps
    Dim tPropsTarget As typeLayerProps
    
    If Not SourceShape.Layer Is TargetShape.Layer Then _
        LayerPropsPreserveAndReset SourceShape.Layer, tPropsSource
    LayerPropsPreserveAndReset TargetShape.Layer, tPropsTarget
    
    Set Intersect = SourceShape.Intersect(TargetShape)
    
    If Not SourceShape.Layer Is TargetShape.Layer Then _
        LayerPropsRestore SourceShape.Layer, tPropsSource
    LayerPropsRestore TargetShape.Layer, tPropsTarget
    
    If Intersect Is Nothing Then Exit Function
    
    Intersect.OrderFrontOf TargetShape
    If Not LeaveSource Then SourceShape.Delete
    If Not LeaveTarget Then TargetShape.Delete

End Function

'отрезать кусок от Shape по контуру Knife, возвращает отрезанный кусок
Public Function Dissect(ByRef Shape As Shape, ByRef Knife As Shape) As Shape
    Set Dissect = Intersect(Knife, Shape, True, True)
    Set Shape = Knife.Trim(Shape, True, False)
End Function

'инструмент Crop Tool
Public Function CropTool( _
                    ByVal ShapeOrRangeOrPage As Object, _
                    ByVal x1#, ByVal y1#, _
                    ByVal x2#, ByVal y2#, _
                    Optional ByVal Angle = 0 _
                ) As ShapeRange
    If TypeOf ShapeOrRangeOrPage Is Shape Or _
         TypeOf ShapeOrRangeOrPage Is ShapeRange Or _
         TypeOf ShapeOrRangeOrPage Is Page Then
        Set CropTool = ShapeOrRangeOrPage.CustomCommand("Crop", "CropRectArea", x1, y1, x2, y2, Angle)
    Else
        Err.Raise 13, Source:="CropTool", Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Function
    End If
End Function

'инструмент Boundary
Public Function CreateBoundary(ByVal ShapeOrRange As Object) As Shape
    On Error GoTo Catch
    Dim tShape As Shape, tRange As ShapeRange
    'просто объект не ест, надо конкретный тип
    If TypeOf ShapeOrRange Is Shape Then
        Set tShape = ShapeOrRange
        Set CreateBoundary = tShape.CustomCommand("Boundary", "CreateBoundary")
    ElseIf TypeOf ShapeOrRange Is ShapeRange Then
        Set tRange = ShapeOrRange
        Set CreateBoundary = tRange.CustomCommand("Boundary", "CreateBoundary")
    Else
        Err.Raise 13, Source:="CreateBoundary", _
            Description:="Type mismatch: ShapeOrRange должен быть Shape или ShapeRange"
        Exit Function
    End If
    Exit Function
Catch:
    Debug.Print Err.Number
End Function

'инструмент Join Curves
Public Function JoinCurves(ByVal SrcRange As ShapeRange, ByVal Tolerance As Double)
    SrcRange.CustomCommand "ConvertTo", "JoinCurves", Tolerance
End Function

'удаление сегмента
'автор: Alex Vakulenko http://www.oberonplace.com/vba/drawmacros/delsegment.htm
Public Sub SegmentDelete(ByVal Segment As Segment)
    If Not Segment.EndNode.IsEnding Then
        Segment.EndNode.BreakApart
        Set Segment = Segment.SubPath.LastSegment
    End If
    Segment.EndNode.Delete
End Sub

'не работает с поверклипом
Public Sub MatrixCopy(ByVal SourceShape As Shape, ByVal TargetShape As Shape)
    Dim tMatrix As typeMatrix
    With tMatrix
        SourceShape.GetMatrix .d11, .d12, .d21, .d22, .tx, .ty
        TargetShape.SetMatrix .d11, .d12, .d21, .d22, .tx, .ty
    End With
End Sub

'присвоить цвет абриса ренджу
Public Sub SetOutlineColor(ByVal Shapes As ShapeRange, ByVal Color As Color)
    Dim Shape As Shape
    For Each Shape In Shapes
        Shape.Outline.Color.CopyAssign Color
    Next Shape
End Sub

Public Sub FitInside(ByVal ShapeToFit As Shape, ByVal TargetRect As Rect)
    If GetHeightKeepProportions(ShapeToFit.BoundingBox, TargetRect.Width) _
     > TargetRect.Height Then
        ShapeToFit.SetSize , TargetRect.Height
    Else
        ShapeToFit.SetSize TargetRect.Width
    End If
    ShapeToFit.CenterX = TargetRect.CenterX
    ShapeToFit.CenterY = TargetRect.CenterY
End Sub

Public Sub FillInside(ByVal ShapeToFill As Shape, ByVal TargetRect As Rect)
    If GetHeightKeepProportions(ShapeToFill.BoundingBox, TargetRect.Width) _
     > TargetRect.Height Then
        ShapeToFill.SetSize TargetRect.Width
    Else
        ShapeToFill.SetSize , TargetRect.Height
    End If
    ShapeToFill.CenterX = TargetRect.CenterX
    ShapeToFill.CenterY = TargetRect.CenterY
End Sub

'===============================================================================
' функции поиска и получения информации об объектах корела
'===============================================================================

'тестирует на пустой кореловский объект
'для пустого объекта коллекции,
'т. к. для Nothing ошибка может быть уже на этапе вызова
Public Function IsNothing(ByVal Object As Object) As Boolean
    Dim t As Variant
    If Object Is Nothing Then GoTo ExitTrue
    If TypeOf Object Is Document Then
        On Error GoTo ExitTrue
        t = Object.Name
    ElseIf TypeOf Object Is Page Then
        On Error GoTo ExitTrue
        t = Object.Name
    ElseIf TypeOf Object Is Layer Then
        On Error GoTo ExitTrue
        t = Object.Name
    ElseIf TypeOf Object Is Shape Then
        On Error GoTo ExitTrue
        t = Object.Name
    ElseIf TypeOf Object Is Curve Then
        On Error GoTo ExitTrue
        t = Object.Length
    ElseIf TypeOf Object Is SubPath Then
        On Error GoTo ExitTrue
        t = Object.Closed
    ElseIf TypeOf Object Is Segment Then
        On Error GoTo ExitTrue
        t = Object.AbsoluteIndex
    ElseIf TypeOf Object Is Node Then
        On Error GoTo ExitTrue
        t = Object.AbsoluteIndex
    End If
    Exit Function
ExitTrue:
    IsNothing = True
End Function

'находит все шейпы с данным именем, включая шейпы в поверклипах, с рекурсией
Public Function FindShapesByName( _
                    ByVal ShapeRange As ShapeRange, _
                    ByVal Name As String _
                ) As ShapeRange
    Set FindShapesByName = FindAllShapes(ShapeRange).Shapes.FindShapes(Name)
End Function

'находит все шейпы, часть имени которых совпадает с NamePart,
'включая шейпы в поверклипах, с рекурсией
Public Function FindShapesByNamePart( _
                    ByVal ShapeRange As ShapeRange, _
                    ByVal NamePart As String _
                ) As ShapeRange
    Set FindShapesByNamePart = FindAllShapes(ShapeRange).Shapes.FindShapes( _
                                   Query:="@Name.Contains('" & NamePart & "')" _
                               )
End Function

'находит поверклипы, без рекурсии
Public Function FindPowerClips(ByVal ShapeRange As ShapeRange) As ShapeRange
    Set FindPowerClips = CreateShapeRange
    'On Error Resume Next
    'FindPowerClips.AddRange ShapeRange.Shapes.FindShapes(Query:="!@com.PowerClip.IsNull")
    Dim Shape As Shape
    For Each Shape In ShapeRange
        If Not lib_elvin.IsNothing(Shape) Then _
            If Not Shape.PowerClip Is Nothing Then FindPowerClips.Add Shape
    Next Shape
End Function

'находит содержимое поверклипов, без рекурсии
Public Function FindShapesInPowerClips(ByVal ShapeRange As ShapeRange) As ShapeRange
    Dim tShape As Shape
    Set FindShapesInPowerClips = CreateShapeRange
    For Each tShape In FindPowerClips(ShapeRange)
        FindShapesInPowerClips.AddRange tShape.PowerClip.Shapes.All
    Next tShape
End Function

'находит все шейпы, включая шейпы в поверклипах, с рекурсией
Public Function FindAllShapes(ByVal ShapeRange As ShapeRange) As ShapeRange
    Dim tShape As Shape
    Set FindAllShapes = CreateShapeRange
    FindAllShapes.AddRange ShapeRange.Shapes.FindShapes
    For Each tShape In FindPowerClips(ShapeRange)
        FindAllShapes.AddRange FindAllShapes(tShape.PowerClip.Shapes.All)
    Next tShape
End Function

'возвращает все шейпы на всех слоях текущей страницы, по умолчанию - без мастер-слоёв и без гайдов
Public Function FindShapesActivePageLayers( _
                    Optional ByVal GuidesLayers As Boolean, _
                    Optional ByVal MasterLayers As Boolean _
                ) As ShapeRange
    Dim tLayer As Layer
    Set FindShapesActivePageLayers = CreateShapeRange
    For Each tLayer In ActivePage.Layers
        If Not (tLayer.IsGuidesLayer And (GuidesLayers = False)) Then _
            FindShapesActivePageLayers.AddRange tLayer.Shapes.All
    Next
    If MasterLayers Then
        For Each tLayer In ActiveDocument.MasterPage.Layers
            If Not (tLayer.IsGuidesLayer And (GuidesLayers = False)) Then _
                FindShapesActivePageLayers.AddRange tLayer.Shapes.All
    Next
    End If
End Function

'возвращает коллекцию слоёв с текущей страницы, имена которых включают NamePart
Public Function FindLayersActivePageByNamePart( _
                    ByVal NamePart As String, _
                    Optional ByVal SearchMasters = True _
                ) As Collection
    Dim tLayer As Layer
    Dim tLayers As Layers
    If SearchMasters Then
        Set tLayers = ActivePage.AllLayers
    Else
        Set tLayers = ActivePage.Layers
    End If
    Set FindLayersActivePageByNamePart = New Collection
    For Each tLayer In tLayers
        If InStr(tLayer.Name, NamePart) > 0 Then _
            FindLayersActivePageByNamePart.Add tLayer
    Next
End Function

'найти дубликат слоя по ряду параметров (достовернее, чем поиск по имени)
Public Function FindLayerDuplicate( _
                    ByVal PageToSearch As Page, _
                    ByVal SrcLayer As Layer _
                ) As Layer
    For Each FindLayerDuplicate In PageToSearch.AllLayers
        With FindLayerDuplicate
            If (.Name = SrcLayer.Name) And _
                 (.IsDesktopLayer = SrcLayer.IsDesktopLayer) And _
                 (.Master = SrcLayer.Master) And _
                 (.Color.IsSame(SrcLayer.Color)) Then _
                 Exit Function
        End With
    Next
    Set FindLayerDuplicate = Nothing
End Function

'возвращает коллекцию слоёв, на которых лежат шейпы из ренджа
Public Function ShapeRangeLayers(ByVal ShapeRange As ShapeRange) As Collection
    
    Dim tShape As Shape
    Dim tLayer As Layer
    Dim inCol As Boolean
    
    If ShapeRange.Count = 0 Then Exit Function
    Set ShapeRangeLayers = New Collection
    If ShapeRange.Count = 1 Then
        ShapeRangeLayers.Add ShapeRange(1).Layer
        Exit Function
    End If
    
    For Each tShape In ShapeRange
        inCol = False
        For Each tLayer In ShapeRangeLayers
            If tLayer Is tShape.Layer Then
                inCol = True
                Exit For
            End If
        Next tLayer
        If inCol = False Then ShapeRangeLayers.Add tShape.Layer
    Next tShape

End Function

'возвращает бОльшую сторону шейпа/рэйнджа/страницы
Public Function GreaterDim(ByVal ShapeOrRangeOrPage As Object) As Double
    If Not TypeOf ShapeOrRangeOrPage Is Shape And Not TypeOf ShapeOrRangeOrPage Is ShapeRange And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="GreaterDim", _
                  Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Function
    End If
    If ShapeOrRangeOrPage.SizeWidth > ShapeOrRangeOrPage.SizeHeight Then
        GreaterDim = ShapeOrRangeOrPage.SizeWidth
    Else
        GreaterDim = ShapeOrRangeOrPage.SizeHeight
    End If
End Function

'возвращает меньшую сторону шейпа/рэйнджа/страницы
Public Function LesserDim(ByVal ShapeOrRangeOrPage As Object) As Double
    If Not TypeOf ShapeOrRangeOrPage Is Shape _
   And Not TypeOf ShapeOrRangeOrPage Is ShapeRange _
   And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="LesserDim", _
                  Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Function
    End If
    If ShapeOrRangeOrPage.SizeWidth < ShapeOrRangeOrPage.SizeHeight Then
        LesserDim = ShapeOrRangeOrPage.SizeWidth
    Else
        LesserDim = ShapeOrRangeOrPage.SizeHeight
    End If
End Function

'возвращает среднее сторон шейпа/рэйнджа/страницы
Public Function AverageDim(ByVal ShapeOrRangeOrPage As Object) As Double
    If Not TypeOf ShapeOrRangeOrPage Is Shape And Not TypeOf ShapeOrRangeOrPage Is ShapeRange And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="AverageDim", _
                  Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Function
    End If
    AverageDim = (ShapeOrRangeOrPage.SizeWidth + ShapeOrRangeOrPage.SizeHeight) _
               / 2
End Function

'возвращает Rect, равный габаритам объекта плюс Space со всех сторон
Public Function SpaceBox(ByVal ShapeOrRange As Object, Space#) As Rect
    If Not TypeOf ShapeOrRange Is Shape _
   And Not TypeOf ShapeOrRange Is ShapeRange Then
        Err.Raise 13, Source:="SpaceBox", _
                  Description:="Type mismatch: ShapeOrRange должен быть Shape или ShapeRange"
        Exit Function
    End If
    Set SpaceBox = ShapeOrRange.BoundingBox.GetCopy
    SpaceBox.Inflate Space, Space, Space, Space
End Function

'является ли шейп/рэйндж/страница альбомным
Public Function IsLandscape(ByVal ShapeOrRangeOrPage As Object) As Boolean
    If Not TypeOf ShapeOrRangeOrPage Is Shape _
   And Not TypeOf ShapeOrRangeOrPage Is ShapeRange _
   And Not TypeOf ShapeOrRangeOrPage Is Page Then
        Err.Raise 13, Source:="IsLandscape", _
                  Description:="Type mismatch: ShapeOrRangeOrPage должен быть Shape, ShapeRange или Page"
        Exit Function
    End If
    If ShapeOrRangeOrPage.SizeWidth > ShapeOrRangeOrPage.SizeHeight Then
        IsLandscape = True
    Else
        IsLandscape = False
    End If
End Function

'являются ли кривые дубликатами, находящимися друг над другом в одном месте
'(underlying dubs)
Public Function IsSameCurves( _
                    ByVal Curve1 As Curve, _
                    ByVal Curve2 As Curve _
                ) As Boolean
    Dim tNode As Node
    Dim Tolerance As Double
    'допуск = 0.001 мм
    Tolerance = ConvertUnits(0.001, cdrMillimeter, ActiveDocument.Unit)
    IsSameCurves = False
    If Not Curve1.Nodes.Count = Curve2.Nodes.Count Then Exit Function
    If Abs(Curve1.Length - Curve2.Length) > Tolerance Then Exit Function
    For Each tNode In Curve1.Nodes
        If Curve2.FindNodeAtPoint( _
               tNode.PositionX, _
               tNode.PositionY, _
               Tolerance * 2 _
           ) Is Nothing Then Exit Function
    Next
    IsSameCurves = True
End Function

'todo: ПРОВЕРИТЬ КАК СЛЕДУЕТ
Public Function IsOverlap( _
                    ByVal FirstShape As Shape, _
                    ByVal SecondShape As Shape _
                ) As Boolean
    
    Dim tIS As Shape
    Dim tShape1 As Shape, tShape2 As Shape
    Dim tBound1 As Shape, tBound2 As Shape
    Dim tProps As typeLayerProps
    
    If FirstShape.Type = cdrConnectorShape _
    Or SecondShape.Type = cdrConnectorShape Then _
        Exit Function
    
    'запоминаем какой слой был активным
    Dim tLayer As Layer: Set tLayer = ActiveLayer
    'запоминаем состояние первого слоя
    FirstShape.Layer.Activate
    LayerPropsPreserveAndReset FirstShape.Layer, tProps
    
    If IsIntersectReady(FirstShape) Then
        Set tShape1 = FirstShape
    Else
        Set tShape1 = CreateBoundary(FirstShape)
        Set tBound1 = tShape1
    End If
    
    If IsIntersectReady(SecondShape) Then
        Set tShape2 = SecondShape
    Else
        Set tShape2 = CreateBoundary(SecondShape)
        Set tBound2 = tShape2
    End If
    
    Set tIS = tShape1.Intersect(tShape2)
    If tIS Is Nothing Then
        IsOverlap = False
    Else
        tIS.Delete
        IsOverlap = True
    End If
    
    On Error Resume Next
        tBound1.Delete
        tBound2.Delete
    On Error GoTo 0
    
    'возвращаем всё на место
    LayerPropsRestore FirstShape.Layer, tProps
    tLayer.Activate

End Function

'IsOverlap здорового человека - меряет по габаритам,
'но зато стабильно работает и в большинстве случаев его достаточно
Public Function IsOverlapBox( _
                    ByVal FirstShape As Shape, _
                    ByVal SecondShape As Shape _
                ) As Boolean
    Dim tShape As Shape
    Dim tProps As typeLayerProps
    'запоминаем какой слой был активным
    Dim tLayer As Layer: Set tLayer = ActiveLayer
    'запоминаем состояние первого слоя
    FirstShape.Layer.Activate
    LayerPropsPreserveAndReset FirstShape.Layer, tProps
    Dim tRect As Rect
    Set tRect = FirstShape.BoundingBox.Intersect(SecondShape.BoundingBox)
    If tRect.Width = 0 And tRect.Height = 0 Then
        IsOverlapBox = False
    Else
        IsOverlapBox = True
    End If
    'возвращаем всё на место
    LayerPropsRestore FirstShape.Layer, tProps
    tLayer.Activate
End Function

Public Function GetWidthKeepProportions( _
                    ByVal Rect As Rect, _
                    ByVal Height As Double _
                ) As Double
    Dim WidthToHeight As Double
    WidthToHeight = Rect.Width / Rect.Height
    GetWidthKeepProportions = Height * WidthToHeight
End Function

Public Function GetHeightKeepProportions( _
                    ByVal Rect As Rect, _
                    ByVal Width As Double _
                ) As Double
    Dim WidthToHeight As Double
    WidthToHeight = Rect.Width / Rect.Height
    GetHeightKeepProportions = Width / WidthToHeight
End Function

'===============================================================================
' функции работы с файлами
'===============================================================================

'находит временную папку
Public Function GetTempFolder() As String
    GetTempFolder = AddProperEndingToPath(VBA.Environ$("TEMP"))
    If FileExists(GetTempFolder) Then Exit Function
    GetTempFolder = AddProperEndingToPath(VBA.Environ$("TMP"))
    If FileExists(GetTempFolder) Then Exit Function
    GetTempFolder = "c:\temp\"
    If FileExists(GetTempFolder) Then Exit Function
    GetTempFolder = "c:\windows\temp\"
    If FileExists(GetTempFolder) Then Exit Function
End Function

'полное имя временного файла
Public Function GetTempFile() As String
    GetTempFile = GetTempFolder & GetTempFileName
End Function

'имя временного файла
Public Function GetTempFileName() As String
    GetTempFileName = "elvin_" & CreateGUID & ".tmp"
End Function

'сохраняет строку Content в файл, перезаписывая, делая в процессе temp файл,
'и оставляя бэкап, если необходимо
Public Sub SaveStrToFile( _
               ByVal Content As String, _
               ByVal File As String, _
               Optional ByVal KeepBak As Boolean = False _
           )

    Dim tFileNum&: tFileNum = FreeFile
    Dim tBak$: tBak = SetFileExt(File, "bak")
    Dim tTemp$
    
    If KeepBak Then
        If FileExists(File) Then FileCopy File, tBak
    Else
        If FileExists(File) Then
            tTemp = GetFilePath(File) & GetTempFileName
            FileCopy File, tTemp
        End If
    End If
        
    Open File For Output Access Write As #tFileNum
    Print #tFileNum, Content
    Close #tFileNum
    
    On Error Resume Next
        If Not KeepBak Then Kill tTemp
    On Error GoTo 0

End Sub

'загружает файл в строку
Public Function LoadStrFromFile(ByVal File As String) As String
    Dim tFileNum&: tFileNum = FreeFile
    Open File For Input As #tFileNum
    LoadStrFromFile = Input(LOF(tFileNum), tFileNum)
    Close #tFileNum
End Function

'заменяет расширение файлу на заданное
Public Function SetFileExt( _
                    ByVal SourceFile As String, _
                    ByVal NewExt As String _
                ) As String
    If Right(SourceFile, 1) <> "\" And Len(SourceFile) > 0 Then
        SetFileExt = GetFileNameNoExt(SourceFile$) & "." & NewExt
    End If
End Function

'возвращает имя файла без расширения
Public Function GetFileNameNoExt(ByVal FileName As String) As String
    If Right(FileName, 1) <> "\" And Len(FileName) > 0 Then
        GetFileNameNoExt = Left(FileName, _
            Switch _
                (InStr(FileName, ".") = 0, _
                    Len(FileName), _
                InStr(FileName, ".") > 0, _
                    InStrRev(FileName, ".") - 1))
    End If
End Function

'создаёт папку, если не было
'возвращает Path обратно (для inline-использования)
Public Function MakeDir(ByVal Path As String) As String
    If Dir(Path, vbDirectory) = "" Then MkDir Path
    MakeDir = Path
End Function

'существует ли файл или папка (папка должна заканчиваться на "\")
Public Function FileExists(ByVal File As String) As Boolean
    If File = "" Then Exit Function
    FileExists = VBA.Len(VBA.Dir(File)) > 0
End Function

Public Function AddProperEndingToPath(ByVal Path As String) As String
    If Not VBA.Right$(Path, 1) = "\" Then AddProperEndingToPath = Path & "\" _
    Else: AddProperEndingToPath = Path
End Function

'---------------------------------------------------------------------------------------
' Procedure         : GetFileName
' Author            : CARDA Consultants Inc.
' Website           : http://www.cardaconsultants.com
' Purpose           : Return the filename from a path\filename input
' Copyright         : The following may be altered and reused as you wish so long as the
'                     copyright notice is left unchanged (including Author, Website and
'                     Copyright).    It may not be sold/resold or reposted on other sites (links
'                     back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFile - string of a path and filename (ie: "c:\temp\Test.xls")
'
' Revision History:
' Rev               Date(yyyy/mm/dd)              Description
' **************************************************************************************
' 1                 2008-Feb-06                   Initial Release
'---------------------------------------------------------------------------------------
Public Function GetFileName(ByVal sFile As String)
On Error GoTo Err_Handler
 
        GetFileName = Right(sFile, Len(sFile) - InStrRev(sFile, "\"))
 
Exit_Err_Handler:
        Exit Function
 
Err_Handler:
        MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
                     "Error Number: " & Err.Number & vbCrLf & _
                     "Error Source: GetFileName" & vbCrLf & _
                     "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
        GoTo Exit_Err_Handler
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetFilePath
' Author            : CARDA Consultants Inc.
' Website           : http://www.cardaconsultants.com
' Purpose           : Return the path from a path\filename input
' Copyright         : The following may be altered and reused as you wish so long as the
'                     copyright notice is left unchanged (including Author, Website and
'                     Copyright).    It may not be sold/resold or reposted on other sites (links
'                     back to this site are allowed).
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sFile - string of a path and filename (ie: "c:\temp\Test.xls")
'
' Revision History:
' Rev               Date(yyyy/mm/dd)              Description
' **************************************************************************************
' 1                 2008-Feb-06                   Initial Release
'---------------------------------------------------------------------------------------
Public Function GetFilePath(ByVal sFile As String)
On Error GoTo Err_Handler
 
        GetFilePath = Left(sFile, InStrRev(sFile, "\"))
 
Exit_Err_Handler:
        Exit Function
 
Err_Handler:
        MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
                     "Error Number: " & Err.Number & vbCrLf & _
                     "Error Source: GetFilePath" & vbCrLf & _
                     "Error Description: " & Err.Description, vbCritical, "An Error has Occurred!"
        GoTo Exit_Err_Handler
End Function

'===============================================================================
' прочие функции
'===============================================================================

Public Sub AssignUnknown(ByRef Destination As Variant, ByRef Value As Variant)
    If VBA.IsObject(Value) Then
        Set Destination = Value
    Else
        Destination = Value
    End If
End Sub

'устарело
Public Sub CopyCollection(ByVal Source As Collection, _
                                                    ByVal Target As Collection)
    Dim Element As Variant
    For Each Element In Source
        Target.Add Element
    Next Element
End Sub

Public Function GetCollectionCopy(ByVal Source As Collection) As Collection
    Set GetCollectionCopy = New Collection
    Dim Item As Variant
    For Each Item In Source
        GetCollectionCopy.Add Item
    Next Item
End Function

Public Function GetCollectionFromDictionary( _
                    ByVal Dictionary As Scripting.IDictionary _
                ) As Collection
    Set GetCollectionFromDictionary = New Collection
    Dim Item As Variant
    For Each Item In Dictionary.Items
        GetCollectionFromDictionary.Add Item
    Next Item
End Function

Public Function GetDictionaryCopy( _
                    ByVal Source As Scripting.IDictionary _
                ) As Scripting.Dictionary
    Set GetDictionaryCopy = New Scripting.Dictionary
    Dim Key As Variant
    For Each Key In Source.Keys
        GetDictionaryCopy.Add Key, Source.Item(Key)
    Next Key
End Function

'https://www.codegrepper.com/code-examples/vb/excel+vba+generate+guid+uuid
Public Function CreateGUID( _
                    Optional ByVal Lowercase As Boolean, _
                    Optional ByVal Parens As Boolean _
                ) As String
    Dim k As Long, H As String
    CreateGUID = VBA.Space(36)
    For k = 1 To VBA.Len(CreateGUID)
        VBA.Randomize
        Select Case k
            Case 9, 14, 19, 24:         H = "-"
            Case 15:                    H = "4"
            Case 20:                    H = VBA.Hex(VBA.Rnd * 3 + 8)
            Case Else:                  H = VBA.Hex(VBA.Rnd * 15)
        End Select
        Mid(CreateGUID, k, 1) = H
    Next
    If Lowercase Then CreateGUID = VBA.LCase$(CreateGUID)
    If Parens Then CreateGUID = "{" & CreateGUID & "}"
End Function

Public Function FindMaxValue(ByVal Collection As Collection) As Variant
    Dim Item As Variant
    For Each Item In Collection
        If VBA.IsNumeric(Item) Then
            If Item > FindMaxValue Then FindMaxValue = Item
        End If
    Next Item
End Function

Public Function FindMinValue(ByVal Collection As Collection) As Variant
    Dim Item As Variant
    For Each Item In Collection
        If VBA.IsNumeric(Item) Then
            If Item < FindMinValue Then
                Debug.Print "aaa"
                FindMinValue = Item
            End If
        End If
    Next Item
End Function

Public Function FindMaxItemNum(ByRef Collection As Collection) As Long
    FindMaxItemNum = 1
    Dim i As Long
    For i = 1 To Collection.Count
        If VBA.IsNumeric(Collection(i)) Then
            If Collection(i) > Collection(FindMaxItemNum) Then _
                FindMaxItemNum = i
        End If
    Next i
End Function

Public Function FindMinItemNum(ByRef Collection As Collection) As Long
    FindMinItemNum = 1
    Dim i As Long
    For i = 1 To Collection.Count
        If VBA.IsNumeric(Collection(i)) Then
            If Collection(i) < Collection(FindMinItemNum) Then _
                FindMinItemNum = i
        End If
    Next i
End Function

Public Function MinOfTwo( _
                    ByVal Value1 As Variant, _
                    ByVal Value2 As Variant _
                ) As Variant
    If Value1 < Value2 Then MinOfTwo = Value1 Else MinOfTwo = Value2
End Function

Public Function MaxOfTwo( _
                    ByVal Value1 As Variant, _
                    ByVal Value2 As Variant _
                ) As Variant
    If Value1 > Value2 Then MaxOfTwo = Value1 Else MaxOfTwo = Value2
End Function

Public Function IsSame( _
                    ByRef Value1 As Variant, _
                    ByRef Value2 As Variant _
                ) As Boolean
    If VBA.IsObject(Value1) And VBA.IsObject(Value2) Then
        IsSame = Value1 Is Value2
    ElseIf Not VBA.IsObject(Value1) And Not VBA.IsObject(Value2) Then
        IsSame = (Value1 = Value2)
    End If
End Function

'функция отсюда: https://stackoverflow.com/questions/38267950/check-if-a-value-is-in-an-array-or-not-with-excel-vba
Public Function IsStrInArr( _
                    ByVal StringToBeFound As String, _
                    Arr As Variant _
                ) As Boolean
        Dim i As Long
        For i = LBound(Arr) To UBound(Arr)
                If Arr(i) = StringToBeFound Then
                        IsStrInArr = True
                        Exit Function
                End If
        Next i
        IsStrInArr = False
End Function

'является ли число чётным :) Что такое Even и Odd запоминать лень...
Public Function IsChet(ByVal X As Variant) As Boolean
    If X Mod 2 = 0 Then IsChet = True Else IsChet = False
End Function

'делится ли Number на Divider нацело
Public Function IsDivider( _
                    ByVal Number As Long, _
                    ByVal Divider As Long _
                ) As Boolean
    If Number Mod Divider = 0 Then IsDivider = True Else IsDivider = False
End Function

Public Sub RemoveElementFromCollection( _
               ByVal Element As Variant, _
               ByVal Collection As Collection _
           )
    If Collection.Count = 0 Then Exit Sub
    Dim i As Long
    For i = 1 To Collection.Count
        If IsSame(Element, Collection(i)) Then
            Collection.Remove i
            Exit Sub
        End If
    Next i
End Sub

'случайное целое от LowerBound до UpperBound
Public Function RndInt( _
                    ByVal LowerBound As Long, _
                    ByVal UpperBound As Long _
                ) As Long
    RndInt = Int((UpperBound - LowerBound + 1) * Rnd + LowerBound)
End Function

Public Function MeasureStart()
    StartTime = Timer
End Function
Public Function MeasureFinish(Optional ByVal Message As String = "")
    Debug.Print Message & CStr(Round(Timer - StartTime, 3)) & " секунд"
End Function

'===============================================================================
' # приватные функции модуля
'===============================================================================

Private Sub LayerPropsPreserve(ByVal L As Layer, ByRef Props As typeLayerProps)
    With Props
        .Visible = L.Visible
        .Printable = L.Printable
        .Editable = L.Editable
    End With
End Sub
Private Sub LayerPropsReset(ByVal L As Layer)
    With L
        If Not .Visible Then .Visible = True
        If Not .Printable Then .Printable = True
        If Not .Editable Then .Editable = True
    End With
End Sub
Private Sub LayerPropsRestore(ByVal L As Layer, ByRef Props As typeLayerProps)
    With Props
        If L.Visible <> .Visible Then L.Visible = .Visible
        If L.Printable <> .Printable Then L.Printable = .Printable
        If L.Editable <> .Editable Then L.Editable = .Editable
    End With
End Sub
Private Sub LayerPropsPreserveAndReset( _
                ByVal L As Layer, _
                ByRef Props As typeLayerProps _
            )
    LayerPropsPreserve L, Props
    LayerPropsReset L
End Sub

'для IsOverlap
Private Function IsIntersectReady(ByVal Shape As Shape) As Boolean
    With Shape
        If .Type = cdrCustomShape Or _
             .Type = cdrBlendGroupShape Or _
             .Type = cdrOLEObjectShape Or _
             .Type = cdrExtrudeGroupShape Or _
             .Type = cdrContourGroupShape Or _
             .Type = cdrBevelGroupShape Or _
             .Type = cdrConnectorShape Or _
             .Type = cdrMeshFillShape Or _
             .Type = cdrTextShape Then
            IsIntersectReady = False
        Else
            IsIntersectReady = True
        End If
    End With
End Function

Private Sub ThrowIfNotCollectionOrArray(ByRef CollectionOrArray As Variant)
    If VBA.IsObject(CollectionOrArray) Then _
        If TypeOf CollectionOrArray Is Collection Then Exit Sub
    If VBA.IsArray(CollectionOrArray) Then Exit Sub
    VBA.Err.Raise 13, Source:="lib_elvin", _
                  Description:="Type mismatch: CollectionOrArray должен быть Collection или Array"
End Sub
