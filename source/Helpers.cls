VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Helpers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Singleton
'@PredeclaredId
Option Explicit

'===============================================================================

Public Sub SendToEditor( _
               ByVal Editor As String, _
               ByVal BitmapShape As Shape, _
               ByVal BitmapFile As String _
           )
    With BitmapProcessor.Create(BitmapShape)
        .SaveTIF BitmapFile
    End With
    OpenInExternalEditor Editor, BitmapFile
End Sub

Public Sub UpdateAfterEdit( _
               ByVal BitmapShape As Shape, _
               ByVal BitmapFile As String _
           )
    With BitmapProcessor.Create(BitmapShape)
        .ReplaceBitmap BitmapFile
    End With
    VBA.Kill BitmapFile
End Sub

Public Function OpenInExternalEditor( _
                    ByVal Editor As String, _
                    ByVal BitmapFileSpec As String _
                )
    Dim CallString As String
    CallString = _
        VBA.Chr(34) & Editor & VBA.Chr(34) _
      & " " & VBA.Chr(34) & BitmapFileSpec & VBA.Chr(34)
    VBA.CreateObject("WScript.Shell").Run CallString
End Function

Public Function GetNewTempBitmapFileSpec( _
                    ByVal DocumentFileName As String, _
                    ByVal StaticID As Long _
                ) As IEither
    Dim TempFolder As String
    TempFolder = GetTempFolder
    If Not FileExists(TempFolder) Then
        Set GetNewTempBitmapFileSpec = Either.Create()
        Exit Function
    End If
    Dim TempFile As String
    TempFile = TempFolder _
             & BuildTempBitmapFileName(DocumentFileName, StaticID)
    Set GetNewTempBitmapFileSpec = Either.Create(TempFile)
End Function

Public Function GetCurrentTempBitmapFileSpec( _
                    ByVal DocumentFileName As String, _
                    ByVal StaticID As Long _
                ) As IEither
    Dim TempFile As String
    TempFile = GetTempFolder _
             & BuildTempBitmapFileName(DocumentFileName, StaticID)
    If Not FileExists(TempFile) Then
        Set GetCurrentTempBitmapFileSpec = Either.Create()
        Exit Function
    End If
    Set GetCurrentTempBitmapFileSpec = Either.Create(TempFile)
End Function

Private Function BuildTempBitmapFileName( _
                     ByVal DocumentFileName As String, _
                     ByVal StaticID As Long _
                 ) As String
    BuildTempBitmapFileName = "tmp_" _
                            & DocumentFileName & "_" _
                            & VBA.CStr(StaticID) _
                            & "_bitmap.tif"
End Function

Public Function ValidateShapeOrNotifyUser(ByVal Shape As Shape) As Boolean
    If Shape Is Nothing Then
        VBA.MsgBox LocalizedStrings("Common.ErrNothingSelected"), vbInformation
        Exit Function
    End If
    If Not Shape.Type = cdrBitmapShape Then
        VBA.MsgBox LocalizedStrings("Common.ErrShapeIsNotBitmap"), vbInformation
        Exit Function
    End If
    ValidateShapeOrNotifyUser = True
End Function

Public Function GetConfig( _
                    Optional ByVal ReadOnly As Boolean = False _
                ) As PresetsConfig
    If ReadOnly Then
        Set GetConfig = PresetsConfig.CreateReadOnly(CONFIG_NAME)
    Else
        Set GetConfig = PresetsConfig.Create(CONFIG_NAME, CreateDefaultPreset)
    End If
End Function

Public Function CreateDefaultPreset() As Dictionary
    Set CreateDefaultPreset = New Dictionary
    With CreateDefaultPreset
        .Add "Editor", "%ProgramFiles%\Adobe\Adobe Photoshop 2023\Photoshop.exe"
    End With
End Function
