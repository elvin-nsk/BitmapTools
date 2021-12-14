VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ILayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Interface
Option Explicit

'===============================================================================

Public Property Get Exists() As Boolean
End Property

Public Function IsNameContains(ByVal NamePart As String, _
                               Optional ByVal Compare As VbCompareMethod = vbBinaryCompare) As Boolean
End Function

Public Function Match(Optional ByVal Name As String = "", _
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
End Function

Public Sub MoveAbove(ByVal Layer As ILayer)
End Sub

Public Sub MoveBelow(ByVal Layer As ILayer)
End Sub

Public Property Get Name() As String
End Property
Public Property Let Name(ByVal Value As String)
End Property

Public Property Get NativeLayer() As Layer
End Property

Public Sub ResetStates()
End Sub

Public Sub RestoreStates()
End Sub

Public Sub SaveStates()
End Sub

Public Sub SaveAndResetStates()
End Sub

Public Sub Delete()
End Sub

'===============================================================================

Public Property Get Self() As ILayer
End Property
