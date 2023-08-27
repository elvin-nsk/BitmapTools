VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalizedStringsEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' cdrEnglishUS (1033)
'===============================================================================

Option Explicit

Implements ILocalizedStrings

Private Strings As Scripting.IDictionary

Private Sub Class_Initialize()
    Set Strings = New Scripting.Dictionary
    With Strings
        
        .Add "Common.ErrNothingSelected", "Select bitmap"
        .Add "Common.ErrShapeIsNotBitmap", "Selected shape is not bitmap"
        .Add "Common.ErrNoDocument", "No active document"
        .Add "Common.ErrLayerDisabled", "Active layer is locked"
        .Add "Common.ErrNoSelection", "No selection"
    
        .Add "BTools.SendToEditorUndo", "Edit bitmap in external editor"
        .Add "BTools.UpdateAfterEditUndo", "Update bitmap"
        .Add "BTools.RemoveCroppingPathUndo", "Release from croping path"
        .Add "BTools.RemoveTransparencyUndo", "Remove transparency"
        .Add "BTools.ErrTempFileCreate", "Unable to create temporary file"
        .Add "BTools.ErrTempFileFind", "Temporary file not found"
    
        .Add "TransView.Caption", "Check transparency"
        .Add "TransView.Transparent", "Bitmap is transparent"
        .Add "TransView.Opaque", "Bitmap is opaque"
        .Add "TransView.BtnOK", "OK"
        
        .Add "UDView.Caption", "Edit bitmap"
        .Add "UDView.Update", "Update bitmap when finished editing"
        .Add "UDView.Cancel", "Do not update, leave file {0} in temporary folder"
        .Add "UDView.Delete", "Do not update, remove temporary file"
        .Add "UDView.BtnUpdateStateWait", "Wait"
        .Add "UDView.BtnUpdateStateReady", "Update"
        .Add "UDView.BtnCancel", "Cancel"
        .Add "UDView.BtnDelete", "Delete"
        
        .Add "Settings.Caption", "Settings"
        .Add "Settings.PresetsFrame", "Presets"
        .Add "Settings.WarningPresetNameAlreadyExists", "A preset with the same name already exists"
        .Add "Settings.WarningPresetNameIsEmpty", "Preset name cannot be empty"
        .Add "Settings.WarningPresetRemove", "Delete preset?"
        .Add "Settings.txtEditor", "Select the image editor executable"
        .Add "Settings.btnOk", "Close"

    End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
    Set ILocalizedStrings_Strings = Strings
End Property
