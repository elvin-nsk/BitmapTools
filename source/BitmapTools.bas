Attribute VB_Name = "BitmapTools"
'===============================================================================
' Макрос           : BitmapTools
' Версия           : 2022.03.03
' Сайты            : https://vk.com/elvin_macro/BitmapTools
'                    https://github.com/elvin-nsk/BitmapTools
' Автор            : elvin-nsk (me@elvin.nsk.ru, https://vk.com/elvin_macro)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

'===============================================================================

Public LocalizedStrings As IStringLocalizer

'===============================================================================

Sub SendToEditor()

  If RELEASE Then On Error GoTo Catch
  
  LocalizedStringsInit
  
  Dim Context As InitialData
  With InitialData.CreateOrNotify
    If .IsError Then
      GoTo Finally
    Else
      Set Context = .SuccessValue
    End If
  End With
    
  Dim BitmapFile As String
  With Helpers.GetNewTempBitmapFileSpec(Context.Document.FileName, _
                                        Context.BitmapShape.StaticID)
    If .IsError Then
      VBA.MsgBox LocalizedStrings("BTools_ErrTempFileCreate"), vbCritical
      GoTo Finally
    Else
      BitmapFile = .SuccessValue
    End If
  End With
  
  lib_elvin.BoostStart LocalizedStrings("BTools_SendToEditorUndo"), _
                       BitmapTools.RELEASE
  
  Context.Layer.SaveAndResetStates
  Helpers.SendToEditor Context.BitmapShape, BitmapFile
  Context.Layer.RestoreStates
      
Finally:
  lib_elvin.BoostFinish
  Set LocalizedStrings = Nothing
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
  Resume Finally
  
End Sub

Sub UpdateAfterEdit()

  If RELEASE Then On Error GoTo Catch
  
  LocalizedStringsInit
  
  Dim Context As InitialData
  With InitialData.CreateOrNotify
    If .IsError Then
      GoTo Finally
    Else
      Set Context = .SuccessValue
    End If
  End With
    
  Dim BitmapFile As String
  With Helpers.GetCurrentTempBitmapFileSpec(Context.Document.FileName, _
                                            Context.BitmapShape.StaticID)
    If .IsError Then
      VBA.MsgBox LocalizedStrings("BTools_ErrTempFileFind"), vbCritical
      GoTo Finally
    Else
      BitmapFile = .SuccessValue
    End If
  End With
  
  lib_elvin.BoostStart LocalizedStrings("BTools_UpdateAfterEditUndo"), _
                       BitmapTools.RELEASE
  
  Context.Layer.SaveAndResetStates
  Helpers.UpdateAfterEdit Context.BitmapShape, BitmapFile
  Context.Layer.RestoreStates
      
Finally:
  lib_elvin.BoostFinish
  Set LocalizedStrings = Nothing
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
  Resume Finally
  
End Sub

Sub SendToEditorAndUpdate()

  If RELEASE Then On Error GoTo Catch
  
  LocalizedStringsInit

  Dim Context As InitialData
  With InitialData.CreateOrNotify
    If .IsError Then
      GoTo Finally
    Else
      Set Context = .SuccessValue
    End If
  End With
    
  Dim BitmapFile As String
  With Helpers.GetNewTempBitmapFileSpec(Context.Document.FileName, _
                                        Context.BitmapShape.StaticID)
    If .IsError Then
      VBA.MsgBox "Не удалось создать временный файл", vbCritical
      GoTo Finally
    Else
      BitmapFile = .SuccessValue
    End If
  End With

  lib_elvin.BoostStart LocalizedStrings("BTools_SendToEditorUndo"), _
                       BitmapTools.RELEASE
                        
  Context.Layer.SaveAndResetStates
  Helpers.SendToEditor Context.BitmapShape, BitmapFile
  
  With New UpdateDialogView
    .FileSpec = BitmapFile
    .Show
    If .IsCancel Then
      GoTo Finally
    ElseIf .IsCancelAndDelete Then
      VBA.Kill BitmapFile
      GoTo Finally
    End If
  End With
  
  If Not lib_elvin.FileExists(BitmapFile) Then
    VBA.MsgBox LocalizedStrings("BTools_ErrTempFileFind"), vbCritical
    GoTo Finally
  End If
  Helpers.UpdateAfterEdit Context.BitmapShape, BitmapFile
  
  Context.Layer.RestoreStates
  
Finally:
  lib_elvin.BoostFinish
  Set LocalizedStrings = Nothing
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
  Resume Finally

End Sub

Sub RemoveCroppingPath()

  If RELEASE Then On Error GoTo Catch
  
  LocalizedStringsInit
  
  Dim Context As InitialData
  With InitialData.CreateOrNotify
    If .IsError Then
      GoTo Finally
    Else
      Set Context = .SuccessValue
    End If
  End With
  
  lib_elvin.BoostStart _
    LocalizedStrings("BTools_RemoveCroppingPathUndo"), False
  Context.Layer.SaveAndResetStates
  Context.BitmapShape.Bitmap.ResetCropEnvelope
  Context.Layer.RestoreStates
  
Finally:
  lib_elvin.BoostFinish
  Set LocalizedStrings = Nothing
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
  Resume Finally

End Sub

Sub RemoveTransparency()

  If RELEASE Then On Error GoTo Catch
  
  LocalizedStringsInit
  
  Dim Context As InitialData
  With InitialData.CreateOrNotify
    If .IsError Then
      GoTo Finally
    Else
      Set Context = .SuccessValue
    End If
  End With
  
  If Not Context.BitmapShape.Bitmap.Transparent Then Exit Sub
  
  lib_elvin.BoostStart LocalizedStrings("BTools_RemoveTransparencyUndo"), RELEASE
  Context.Layer.SaveAndResetStates
  With BitmapProcessor.Create(Context.BitmapShape)
    .Flatten
    .Shape.CreateSelection
  End With
  Context.Layer.RestoreStates
  
Finally:
  lib_elvin.BoostFinish
  Set LocalizedStrings = Nothing
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
  Resume Finally

End Sub

Sub CheckTransparency()

  LocalizedStringsInit
  
  If ActiveDocument Is Nothing Then Exit Sub

  Dim Context As InitialData
  With InitialData.CreateOrNotify
    If .IsError Then
      GoTo Finally
    Else
      Set Context = .SuccessValue
    End If
  End With
  
  With New TransparencyView
    .Transparent = Context.BitmapShape.Bitmap.Transparent
    .Show
  End With
  
Finally:
  lib_elvin.BoostFinish
  Set LocalizedStrings = Nothing
  Exit Sub
  
End Sub

'===============================================================================

Private Sub LocalizedStringsInit()
  With StringLocalizer.Builder(cdrEnglishUS, LocalizedStringsEN.Strings)
    .WithLocale cdrRussian, LocalizedStringsRU.Strings
    Set LocalizedStrings = .Build
  End With
End Sub
