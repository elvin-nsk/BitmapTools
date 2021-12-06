VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BitmapContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Public Document As Document
Public BitmapShape As Shape

'===============================================================================

Public Function CreateOrNotify() As IEither
  With New BitmapContext
    Set .Document = ActiveDocument
    If .Document Is Nothing Then
      Set CreateOrNotify = Either.Create
      Exit Function
    End If
    Set .BitmapShape = ActiveShape
    If Not Helpers.ValidateShapeOrNotify(.BitmapShape) Then
      Set CreateOrNotify = Either.Create
      Exit Function
    End If
    Set CreateOrNotify = Either.Create(.Self)
  End With
End Function

'===============================================================================

Public Property Get Self() As BitmapContext
  Set Self = Me
End Property
