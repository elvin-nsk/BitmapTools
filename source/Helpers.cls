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

Private Sub Class_Initialize()
    If Not Me Is Helpers Then Err.Raise 425 'Invalid object use
End Sub

'===============================================================================

Public Sub SendToEditor(ByVal BitmapShape As Shape, ByVal BitmapFile As String)
    With BitmapProcessor.Create(BitmapShape)
        .SaveTIF BitmapFile
    End With
    OpenInExternalEditor BitmapFile
End Sub

Public Sub UpdateAfterEdit(ByVal BitmapShape As Shape, ByVal BitmapFile As String)
    With BitmapProcessor.Create(BitmapShape)
        .ReplaceBitmap BitmapFile
    End With
    VBA.Kill BitmapFile
End Sub

Public Function OpenInExternalEditor(ByVal BitmapFileSpec As String)
    VBA.CreateObject("Photoshop.Application").Open BitmapFileSpec
End Function

Public Function GetNewTempBitmapFileSpec( _
                    ByVal DocumentFileName As String, _
                    ByVal StaticID As Long _
                ) As IEither
    Dim TempFolder As String
    TempFolder = lib_elvin.GetTempFolder
    If Not lib_elvin.FileExists(TempFolder) Then
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
    TempFile = lib_elvin.GetTempFolder _
             & BuildTempBitmapFileName(DocumentFileName, StaticID)
    If Not lib_elvin.FileExists(TempFile) Then
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
        VBA.MsgBox LocalizedStrings("Common_ErrNothingSelected"), vbInformation
        Exit Function
    End If
    If Not Shape.Type = cdrBitmapShape Then
        VBA.MsgBox LocalizedStrings("Common_ErrShapeIsNotBitmap"), vbInformation
        Exit Function
    End If
    ValidateShapeOrNotifyUser = True
End Function
