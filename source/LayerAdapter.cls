VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LayerAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

Implements ILayer

'===============================================================================

Private Type typeThis
    NativeLayer As Layer
    SavedVisibility As Boolean
    SavedPrintability As Boolean
    SavedEditability As Boolean
End Type
Private This As typeThis

'===============================================================================

Public Function Create(ByVal Layer As Layer) As ILayer
    With New LayerAdapter
        Set Create = .Init(Layer)
    End With
End Function

Friend Function Init(ByVal Layer As Layer) As ILayer
    With This
        Set .NativeLayer = Layer
    End With
    Set Init = Me
End Function

'===============================================================================

Private Sub ILayer_Delete()
    This.NativeLayer.Delete
End Sub

Private Property Get ILayer_Exists() As Boolean
    ILayer_Exists = Not This.NativeLayer Is Nothing
End Property

Private Function ILayer_IsNameContains( _
                     ByVal NamePart As String, _
                     Optional ByVal Compare As VbCompareMethod = vbBinaryCompare _
                 ) As Boolean
    If VBA.InStr(1, This.NativeLayer.Name, NamePart, Compare) Then _
        ILayer_IsNameContains = True
End Function

Private Function ILayer_Match( _
                     Optional ByVal Name As String = "", _
                     Optional ByVal Visible As VbTriState = vbUseDefault, _
                     Optional ByVal Printable As VbTriState = vbUseDefault, _
                     Optional ByVal Editable As VbTriState = vbUseDefault, _
                     Optional ByVal Master As VbTriState = vbUseDefault, _
                     Optional ByVal Special As VbTriState = vbUseDefault, _
                     Optional ByVal Desktop As VbTriState = vbUseDefault, _
                     Optional ByVal Guides As VbTriState = vbUseDefault, _
                     Optional ByVal Grid As VbTriState = vbUseDefault, _
                     Optional ByVal Color As Color, _
                     Optional ByVal OverrideColor As VbTriState = vbUseDefault _
                 ) As Boolean
    With This.NativeLayer
        If Name <> "" And .Name <> Name Then Exit Function
        If Visible = vbTrue And Not .Visible Then Exit Function
        If Visible = vbFalse And .Visible Then Exit Function
        If Printable = vbTrue And Not .Printable Then Exit Function
        If Printable = vbFalse And .Printable Then Exit Function
        If Editable = vbTrue And Not .Editable Then Exit Function
        If Editable = vbFalse And .Editable Then Exit Function
        If Master = vbTrue And Not .Master Then Exit Function
        If Master = vbFalse And .Master Then Exit Function
        If Special = vbTrue And Not .IsSpecialLayer Then Exit Function
        If Special = vbFalse And .IsSpecialLayer Then Exit Function
        If Desktop = vbTrue And Not .IsDesktopLayer Then Exit Function
        If Desktop = vbFalse And .IsDesktopLayer Then Exit Function
        If Guides = vbTrue And Not .IsGuidesLayer Then Exit Function
        If Guides = vbFalse And .IsGuidesLayer Then Exit Function
        If Grid = vbTrue And Not .IsGridLayer Then Exit Function
        If Grid = vbFalse And .IsGridLayer Then Exit Function
        If Not Color Is Nothing Then
            If Not .Color.IsSame(Color) Then Exit Function
        End If
        If OverrideColor = vbTrue And Not .OverrideColor Then Exit Function
        If OverrideColor = vbFalse And .OverrideColor Then Exit Function
    End With
    ILayer_Match = True
End Function

Private Sub ILayer_MoveAbove(ByVal Layer As ILayer)
    This.NativeLayer.MoveAbove Layer.NativeLayer
End Sub

Private Sub ILayer_MoveBelow(ByVal Layer As ILayer)
    This.NativeLayer.MoveBelow Layer.NativeLayer
End Sub

Private Property Get ILayer_Name() As String
    ILayer_Name = This.NativeLayer.Name
End Property
Private Property Let ILayer_Name(ByVal Value As String)
    This.NativeLayer.Name = Value
End Property

Private Property Get ILayer_NativeLayer() As Layer
    Set ILayer_NativeLayer = This.NativeLayer
End Property

Private Sub ILayer_ResetStates()
    With This.NativeLayer
        If Not .Visible Then .Visible = True
        If Not .Printable Then .Printable = True
        If Not .Editable Then .Editable = True
    End With
End Sub

Private Sub ILayer_RestoreStates()
    With This
        If Not .SavedVisibility = .NativeLayer.Visible Then _
            .NativeLayer.Visible = .SavedVisibility
        If Not .SavedPrintability = .NativeLayer.Printable Then _
            .NativeLayer.Printable = .SavedPrintability
        If Not .SavedEditability = .NativeLayer.Editable Then _
            .NativeLayer.Editable = .SavedEditability
    End With
End Sub

Private Sub ILayer_SaveAndResetStates()
    Self.SaveStates
    Self.ResetStates
End Sub

Private Sub ILayer_SaveStates()
    With This
        .SavedVisibility = .NativeLayer.Visible
        .SavedPrintability = .NativeLayer.Printable
        .SavedEditability = .NativeLayer.Editable
    End With
End Sub

Private Property Get ILayer_Self() As ILayer
    Set ILayer_Self = Me
End Property

'===============================================================================

Private Property Get Self() As ILayer
    Set Self = Me
End Property
