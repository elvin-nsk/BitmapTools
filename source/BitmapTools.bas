Attribute VB_Name = "BitmapTools"
'===============================================================================
'   Макрос          : BitmapTools
'   Версия          : 2023.10.11
'   Сайты           : https://vk.com/elvin_macro/BitmapTools
'                     https://github.com/elvin-nsk/BitmapTools
'   Автор           : elvin-nsk (me@elvin.nsk.ru, https://vk.com/elvin_macro)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

Public Const CONFIG_NAME As String = "elvin_BitmapTools"
Public Const EDITOR_KEY As String = "Editor"
Public Const DEFAULT_PRESET As String = "Default"

Public LocalizedStrings As IStringLocalizer

Sub SendToEditor()

    If RELEASE Then On Error GoTo Catch
    
    LocalizedStringsInit
    
    Dim Context As InputData
    Set Context = GetContext
    If Context.IsError Then GoTo Finally
    Dim Layer As ILayer
    Set Layer = LayerAdapter.Create(Context.Layer)
    
    Dim Cfg As PresetsConfig
    Set Cfg = Helpers.GetConfig(True)
        
    Dim BitmapFile As String
    With Helpers.GetNewTempBitmapFileSpec( _
                     Context.Document.FileName, _
                     Context.Shape.StaticID _
                 )
        If .IsError Then
            VBA.MsgBox LocalizedStrings("BTools.ErrTempFileCreate"), vbCritical
            GoTo Finally
        Else
            BitmapFile = .SuccessValue
        End If
    End With
    
    BoostStart _
        LocalizedStrings("BTools.SendToEditorUndo"), BitmapTools.RELEASE
    
    Layer.SaveAndResetStates
    Helpers.SendToEditor Cfg(EDITOR_KEY), Context.Shape, BitmapFile
    Layer.RestoreStates
            
Finally:
    BoostFinish
    Set LocalizedStrings = Nothing
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
    
End Sub

Sub UpdateAfterEdit()

    If RELEASE Then On Error GoTo Catch
    
    LocalizedStringsInit
    
    Dim Context As InputData
    Set Context = GetContext
    If Context.IsError Then GoTo Finally
    Dim Layer As ILayer
    Set Layer = LayerAdapter.Create(Context.Layer)
        
    Dim BitmapFile As String
    With Helpers.GetCurrentTempBitmapFileSpec( _
                     Context.Document.FileName, _
                     Context.Shape.StaticID _
                 )
        If .IsError Then
            VBA.MsgBox LocalizedStrings("BTools.ErrTempFileFind"), vbCritical
            GoTo Finally
        Else
            BitmapFile = .SuccessValue
        End If
    End With
    
    BoostStart _
        LocalizedStrings("BTools.UpdateAfterEditUndo"), BitmapTools.RELEASE
    
    Layer.SaveAndResetStates
    Helpers.UpdateAfterEdit Context.Shape, BitmapFile
    Layer.RestoreStates
            
Finally:
    BoostFinish
    Set LocalizedStrings = Nothing
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally
    
End Sub

Sub SendToEditorAndUpdate()

    If RELEASE Then On Error GoTo Catch
    
    LocalizedStringsInit

    Dim Context As InputData
    Set Context = GetContext
    If Context.IsError Then GoTo Finally
    Dim Layer As ILayer
    Set Layer = LayerAdapter.Create(Context.Layer)
        
    Dim BitmapFile As String
    With Helpers.GetNewTempBitmapFileSpec( _
                     Context.Document.FileName, _
                     Context.Shape.StaticID _
                 )
        If .IsError Then
            VBA.MsgBox LocalizedStrings("BTools.ErrTempFileCreate"), vbCritical
            GoTo Finally
        Else
            BitmapFile = .SuccessValue
        End If
    End With
    
    Dim Cfg As PresetsConfig
    Set Cfg = Helpers.GetConfig(True)

    BoostStart _
        LocalizedStrings("BTools.SendToEditorUndo"), BitmapTools.RELEASE
                                                
    Layer.SaveAndResetStates
    Helpers.SendToEditor Cfg(EDITOR_KEY), Context.Shape, BitmapFile
    
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
    
    If Not FileExists(BitmapFile) Then
        VBA.MsgBox LocalizedStrings("BTools.ErrTempFileFind"), vbCritical
        GoTo Finally
    End If
    Helpers.UpdateAfterEdit Context.Shape, BitmapFile
    
    Layer.RestoreStates
    
Finally:
    BoostFinish
    Set LocalizedStrings = Nothing
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

Sub RemoveCroppingPath()

    If RELEASE Then On Error GoTo Catch
    
    LocalizedStringsInit
    
    Dim Context As InputData
    Set Context = GetContext
    If Context.IsError Then GoTo Finally
    Dim Layer As ILayer
    Set Layer = LayerAdapter.Create(Context.Layer)
    
    BoostStart _
        LocalizedStrings("BTools.RemoveCroppingPathUndo"), False
    Layer.SaveAndResetStates
    Context.Shape.Bitmap.ResetCropEnvelope
    Layer.RestoreStates
    
Finally:
    BoostFinish
    Set LocalizedStrings = Nothing
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

Sub RemoveTransparency()

    If RELEASE Then On Error GoTo Catch
    
    LocalizedStringsInit
    
    Dim Context As InputData
    Set Context = GetContext
    If Context.IsError Then GoTo Finally
    Dim Layer As ILayer
    Set Layer = LayerAdapter.Create(Context.Layer)
    
    If Not Context.Shape.Bitmap.Transparent Then Exit Sub
    
    BoostStart LocalizedStrings("BTools.RemoveTransparencyUndo"), RELEASE
    Layer.SaveAndResetStates
    With BitmapProcessor.Create(Context.Shape)
        .Flatten
        .Shape.CreateSelection
    End With
    Layer.RestoreStates
    
Finally:
    BoostFinish
    Set LocalizedStrings = Nothing
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

Sub CheckTransparency()

    LocalizedStringsInit
    
    If ActiveDocument Is Nothing Then Exit Sub

    Dim Context As InputData
    Set Context = GetContext
    If Context.IsError Then GoTo Finally
    
    With New TransparencyView
        .Transparent = Context.Shape.Bitmap.Transparent
        .Show
    End With
    
Finally:
    BoostFinish
    Set LocalizedStrings = Nothing
    Exit Sub
    
End Sub

Sub Settings()
    LocalizedStringsInit
    With New SettingsView
        .Show
    End With
    Set LocalizedStrings = Nothing
End Sub

'===============================================================================

Private Sub LocalizedStringsInit()
    With StringLocalizer.Builder(cdrEnglishUS, New LocalizedStringsEN)
        .WithLocale cdrRussian, New LocalizedStringsRU
        .WithLocale cdrBrazilianPortuguese, New LocalizedStringsBR
        Set LocalizedStrings = .Build
    End With
End Sub

Private Function GetContext() As InputData
    Set GetContext = _
        InputData.RequestShapes( _
            ErrNoDocument:=LocalizedStrings("Common.ErrNoDocument"), _
            ErrLayerDisabled:=LocalizedStrings("Common.ErrLayerDisabled"), _
            ErrNoSelection:=LocalizedStrings("Common.ErrNoSelection") _
        )
End Function

'===============================================================================

Private Sub TestPresets()
    Dim PresetsVew As New SettingsView
    Dim Dic As New Dictionary
    With Dic
        .Add "one", "1st content"
        .Add "two", "2nd content"
        .Add "another one", "3d content"
    End With
    Dim PresetsManager As PresetsController
    Set PresetsManager = _
        PresetsController.Create(PresetsVew, Dic, "two")
    'PresetsManager.SetProtectedKey "one"
    PresetsVew.Show
End Sub

Private Sub TestPresetsConfig()
    With PresetsConfig.Create(CONFIG_NAME, Helpers.CreateDefaultPreset)
        Show .CurrentPreset("Editor")
    End With
End Sub

Private Sub TestDic()
    Dim Dic As New Dictionary
    Dim a As Variant
    a = Dic("x")
    Show Dic.Count
End Sub
