VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SettingsView 
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6975
   OleObjectBlob   =   "SettingsView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SettingsView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'===============================================================================

Private WithEvents PresetsManager As PresetsController
Attribute PresetsManager.VB_VarHelpID = -1
Private Cfg As PresetsConfig

'===============================================================================

Private Sub UserForm_Initialize()
    Set Cfg = Helpers.GetConfig
    Set PresetsManager = _
        PresetsController.Create(Me, Cfg.Presets, DEFAULT_PRESET)
    With PresetsManager
        .SetProtectedKey DEFAULT_PRESET
        .SetWarningPresetNameAlreadyExists _
            LocalizedStrings("Settings.WarningPresetNameAlreadyExists")
        .SetWarningPresetNameIsEmpty _
            LocalizedStrings("Settings.WarningPresetNameIsEmpty")
        .SetWarningPresetRemove _
            LocalizedStrings("Settings.WarningPresetRemove")
    End With
    Me.Caption = LocalizedStrings("Settings.Caption")
    tboxEditor.Value = Cfg("Editor")
    PresetsFrame.Caption = LocalizedStrings("Settings.PresetsFrame")
    txtEditor.Caption = LocalizedStrings("Settings.txtEditor")
    btnOk.Caption = LocalizedStrings("Settings.btnOk")
End Sub

Private Sub UserForm_Activate()
    '
End Sub

Private Sub PresetsManager_OnSelectionChange(outNewKey As String)
    Cfg.CurrentName = outNewKey
    tboxEditor.Value = Cfg("Editor")
End Sub

Private Sub btnBrowseEditor_Click()
    With New FileBrowser
        .Filter = "exe" & Chr(0) & "*.exe"
        .MultiSelect = False
        Dim Result As Collection
        Set Result = .ShowFileOpenDialog
        If Result.Count > 0 Then
            tboxEditor.Value = Result(1)
        End If
    End With
End Sub

Private Sub btnOk_Click()
    FormŒ 
End Sub

'===============================================================================

Private Sub FormŒ ()
    Me.Hide
    Cfg("Editor") = tboxEditor.Value
    Cfg.ForceSave
End Sub

'===============================================================================

Private Sub UserForm_QueryClose(—ancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        —ancel = True
        FormŒ 
    End If
End Sub
