VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BitmapProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================

Private Const CustomError = vbObjectError Or 32

Private Type typeThis
  Shape As Shape
  Bitmap As Bitmap
  Layer As Layer
  Doc As Document
  WidthPx As Long
  HeightPx As Long
  CleanWidthInches As Long
  CleanHeightInches As Long
  CleanWidth As Double
  CleanHeight As Double
  WidthMultiplier As Double
  HeightMultiplier As Double
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal BitmapShape As Shape) As BitmapProcessor
  If Not BitmapShape.Type = cdrBitmapShape Then _
    Err.Raise CustomError, "BitmapProcessor", "BitmapShape must be Bitmap type"
  With New BitmapProcessor
    Set Create = .Init(BitmapShape)
  End With
End Function

Friend Function Init(ByVal BitmapShape As Shape) As BitmapProcessor
  With This
    Set .Shape = BitmapShape
    Set .Bitmap = .Shape.Bitmap
    Set .Layer = .Shape.Layer
    Set .Doc = .Shape.Page.Parent.Parent
  End With
  CalculateSizes
  Set Init = Me
End Function

'===============================================================================

Property Get Shape() As Shape
  Set Shape = This.Shape
End Property

Property Get Bitmap() As Bitmap
  Set Bitmap = This.Bitmap
End Property

Public Function Flatten() As Shape

  If Not This.Shape.Bitmap.Transparent Then Exit Function
  
  Dim BmpCopy As Shape
  Set BmpCopy = This.Shape.Duplicate
  ResetShape BmpCopy
  With BmpCopy.Bitmap
    Set BmpCopy = BmpCopy.ConvertToBitmapEx _
                  (.Mode, , False, .ResolutionX, cdrNoAntiAliasing, False)
  End With
  ReplaceBitmapWith BmpCopy

End Function

Public Function SavePSD(ByVal FileName As String)
  Save FileName, cdrPSD
End Function

Public Function SaveTIF(ByVal FileName As String)
  Save FileName, cdrTIFF
End Function

Public Function ReplaceBitmap(ByVal FileName As String)
  With This
    .Layer.Import FileName
    If .Shape Is Nothing Then
      Init .Doc.ActiveShape
    Else
      ReplaceBitmapWith .Doc.ActiveShape
    End If
  End With
End Function

Public Property Get Self() As BitmapProcessor
  Set Self = Me
End Property

'===============================================================================

Private Sub CalculateSizes()
  With This
    .WidthPx = .Bitmap.SizeWidth
    .HeightPx = .Bitmap.SizeHeight
    .CleanWidthInches = .Bitmap.SizeWidth / .Bitmap.ResolutionX
    .CleanHeightInches = .Bitmap.SizeHeight / .Bitmap.ResolutionY
    .CleanWidth = InchesToCurrentUnits(.CleanWidthInches)
    .CleanHeight = InchesToCurrentUnits(.CleanHeightInches)
    .WidthMultiplier = 1
    .HeightMultiplier = 1
  End With
End Sub

Private Sub ResetShape(ByVal Shape As Shape)
  Dim W#, Ratio#
  W = Shape.SizeWidth
  With Shape
    ClearEffects Shape
    .ClearTransformations
    .Bitmap.ResetCropEnvelope
    CalculateSizes
    If Not .Bitmap.ResolutionX = .Bitmap.ResolutionY Then
      Ratio = .Bitmap.ResolutionX / .Bitmap.ResolutionY
      '.SetSize w
      .SetSize .SizeWidth * Ratio, .SizeHeight
      'This.WidthMultiplier = .SizeWidth / This.CleanWidth
      'This.HeightMultiplier = .SizeHeight / This.CleanHeight
    End If
  End With
End Sub

Private Function ReplaceBitmapWith(ByVal NewBitmapShape As Shape)
  
  If Not NewBitmapShape.Type = cdrBitmapShape Then Exit Function
  
  With This
  
    Dim Env As Curve
    Set Env = .Bitmap.CropEnvelope.GetCopy
    .Bitmap.ResetCropEnvelope
    
    CopyShapeSizeAndPosition .Shape, NewBitmapShape
    NewBitmapShape.Bitmap.CropEnvelope.CopyAssign Env
    CopyShapeEffects .Shape, NewBitmapShape
    CopyShapeName .Shape, NewBitmapShape
    NewBitmapShape.OrderBackOf .Shape
    
    .Shape.Delete
    Init NewBitmapShape
  
  End With
  
End Function

Private Function Save(ByVal FileName As String, ByVal Filter As cdrFilter)
    
  Dim BmpCopy As Shape
  
  If This.Shape.Type = cdrBitmapShape Then
    Set BmpCopy = This.Shape.Duplicate
    ResetShape BmpCopy
    If This.Shape.Bitmap.Transparent Then
      BmpCopy.CreateSelection
      With BmpCopy.Bitmap
        This.Doc.ExportBitmap(FileName, Filter, cdrSelection, _
                              .Mode, .SizeWidth, .SizeHeight, _
                              .ResolutionX, .ResolutionY, _
                              cdrNoAntiAliasing, False, True, True, False _
                              ).Finish
      End With
    Else
      BmpCopy.Bitmap.SaveAs(FileName, Filter).Finish
    End If
    BmpCopy.Delete
  End If
  
End Function

Private Sub CopyShapeSizeAndPosition(ByVal Source As Shape, ByVal Target As Shape)

  Dim d11 As Double, d12 As Double, d21 As Double, d22 As Double
  Dim X As Double, Y As Double, W As Double, H As Double
  Dim tx As Double, ty As Double
  
  'позиция и геометрические искажения
  Source.GetMatrix d11, d12, d21, d22, tx, ty
  Target.SetMatrix d11 / This.WidthMultiplier, d12 / This.WidthMultiplier, d21 / This.HeightMultiplier, d22 / This.HeightMultiplier, tx, ty
  'небольшой хак, чтобы размер был совсем точным (без учёта Multiplier пока что, надо дорабатывать эту концепцию)
  Target.SizeWidth = Source.SizeWidth
  Target.SizeHeight = Source.SizeHeight
  
  Target.OrderBackOf Source

End Sub

Private Sub CopyShapeName(ByVal Source As Shape, ByVal Target As Shape)
  Target.Name = Source.ObjectData("Name").FormattedValue
End Sub

'за основу взята функция из макроса ExImPhotoshop by Dizzy
Private Sub CopyShapeEffects(ByVal Source As Shape, ByVal Target As Shape)
    
  Dim i As Long
  
  On Error Resume Next

  If Not Source.Effects.DropShadowEffect Is Nothing Then
    With Source.Effects.DropShadowEffect.DropShadow
      Target.CreateDropShadow .Type, .Opacity, .Feather, _
                              .OffsetX * Target.AbsoluteHScale, _
                              .OffsetY * Target.AbsoluteVScale, _
                              .Color, .FeatherType, .FeatherEdge, _
                              .PerspectiveAngle, .PerspectiveStretch, .Fade
                                      
    End With
  End If
  
  If Source.Transparency.Type = cdrUniformTransparency Then
    Target.Transparency.ApplyUniformTransparency (Source.Transparency.Uniform)
    Target.Transparency.MergeMode = Source.Transparency.MergeMode
  ElseIf Source.Transparency.Type = cdrFountainTransparency Then
    With Source.Transparency
      Target.Transparency _
        .ApplyFountainTransparency .Start, .End, .Fountain.Type, .Fountain.Angle, _
                                   .Fountain.Steps, .Fountain.EdgePad, .Fountain.MidPoint
      Target.Transparency.MergeMode = Source.Transparency.MergeMode
      Target.Transparency.Fountain.StartX = .Fountain.StartX
      Target.Transparency.Fountain.StartY = .Fountain.StartY
      Target.Transparency.Fountain.EndX = .Fountain.EndX
      Target.Transparency.Fountain.EndY = .Fountain.EndY
      If Source.Transparency.Fountain.Colors.Count > 0 Then
        For i = 1 To .Fountain.Colors.Count
          Target.Transparency.Fountain.Colors.Add .Fountain.Colors(i).Color, .Fountain.Colors(i).Position
        Next i
      End If
    End With
      
   ElseIf Source.Transparency.Type = cdrPatternTransparency Then
    Target.Transparency.ApplyPatternTransparency Source.Transparency.Pattern.Type, _
    Source.Transparency.Pattern.FilePath, Source.Transparency.Pattern.Canvas.Index, _
    Source.Transparency.Start, Source.Transparency.End, Source.Transparency.Pattern.TransformWithShape
    Target.Transparency.MergeMode = Source.Transparency.MergeMode
    Target.Transparency.Pattern.BackColor = Source.Transparency.Pattern.BackColor
    Target.Transparency.Pattern.BackColor = Source.Transparency.Pattern.BackColor
    Target.Transparency.Pattern.FrontColor = Source.Transparency.Pattern.FrontColor
    Target.Transparency.Pattern.OriginX = Source.Transparency.Pattern.OriginX
    Target.Transparency.Pattern.OriginY = Source.Transparency.Pattern.OriginY
    Target.Transparency.Pattern.RotationAngle = Source.Transparency.Pattern.RotationAngle
    Target.Transparency.Pattern.SkewAngle = Source.Transparency.Pattern.SkewAngle
    Target.Transparency.Pattern.MirrorFill = Source.Transparency.Pattern.MirrorFill
    Target.Transparency.Pattern.TileHeight = Source.Transparency.Pattern.TileHeight
    Target.Transparency.Pattern.TileWidth = Source.Transparency.Pattern.TileWidth
    Target.Transparency.Pattern.TransformWithShape = Source.Transparency.Pattern.TransformWithShape
    Target.Transparency.Pattern.TileOffsetType = Source.Transparency.Pattern.TileOffsetType
    Target.Transparency.Pattern.TileOffset = Source.Transparency.Pattern.TileOffset
    Target.Transparency.Pattern = Source.Transparency.Pattern
  ElseIf Source.Transparency.Type = cdrTextureTransparency Then
    Target.Transparency.ApplyTextureTransparency Source.Transparency.Texture.TextureName, _
    Source.Transparency.Texture.LibraryName, Source.Transparency.Start, Source.Transparency.End
    Target.Transparency.MergeMode = Source.Transparency.MergeMode
  End If

End Sub

Private Sub ClearEffects(ByVal Shape As Shape)
  If Shape.Effects.Count = 0 Then Exit Sub
  On Error Resume Next
  Dim i As Long
  For i = 1 To Shape.Effects.Count
    Shape.Effects(i).Clear
  Next
  On Error GoTo 0
End Sub

Private Function InchesToCurrentUnits(ByVal Size As Double) As Double
  InchesToCurrentUnits = ConvertUnits(Size, cdrInch, This.Doc.Unit)
End Function

Private Function CurrentUnitsToInches(ByVal Size As Double) As Double
  CurrentUnitsToInches = ConvertUnits(Size, This.Doc.Unit, cdrInch)
End Function
