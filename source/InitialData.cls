VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InitialData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
Option Explicit

'===============================================================================

Public Document As Document
Public Layer As ILayer
Public BitmapShape As Shape

'===============================================================================

Public Function CreateOrNotifyUser() As IEither
    With New InitialData
        Set .Document = ActiveDocument
        If .Document Is Nothing Then
            Set CreateOrNotifyUser = Either.Create
            Exit Function
        End If
        Set .BitmapShape = ActiveShape
        If Not Helpers.ValidateShapeOrNotifyUser(.BitmapShape) Then
            Set CreateOrNotifyUser = Either.Create
            Exit Function
        End If
        Set .Layer = LayerAdapter.Create(ActiveLayer)
        Set CreateOrNotifyUser = Either.Create(.Self)
    End With
End Function

'===============================================================================

Public Property Get Self() As InitialData
    Set Self = Me
End Property
