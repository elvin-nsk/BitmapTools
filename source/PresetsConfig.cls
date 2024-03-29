VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PresetsConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@PredeclaredId
Option Explicit

'===============================================================================
' # Declarations

Private Const CURRENT_KEY As String = "Current"
Private Const DEFAULT_PRESET_NAME As String = "Default"
Private Const PRESETS_DICTIONARY_KEY As String = "Presets"

Private Type typeThis
    ConfigDictionary As Scripting.IDictionary
    CurrentPresetName As String
    FallbackPreset As Scripting.IDictionary
    FileSpec As String
    ReadOnly As Boolean
End Type
Private This As typeThis

'===============================================================================
' # Constructors

Public Function Create( _
                    ByVal FileBaseName As String, _
                    ByVal FallbackPreset As Scripting.IDictionary, _
                    Optional ByVal DefaultPresetName As String _
                        = DEFAULT_PRESET_NAME _
                ) As PresetsConfig
    With New PresetsConfig
        Set Create = _
            .Init(FileBaseName, FallbackPreset, DefaultPresetName, False)
    End With
End Function

Public Function CreateReadOnly( _
                    ByVal FileBaseName As String _
                ) As PresetsConfig
    With New PresetsConfig
        Set CreateReadOnly = _
            .Init(FileBaseName, Nothing, vbNullString, True)
    End With
End Function

Friend Function Init( _
                    ByVal FileBaseName As String, _
                    ByVal FallbackPreset As Scripting.IDictionary, _
                    ByVal DefaultPresetName As String, _
                    ByVal ReadOnly As Boolean _
                ) As PresetsConfig
    With This
        .ReadOnly = ReadOnly
        Set .FallbackPreset = FallbackPreset
        Set .ConfigDictionary = _
            GetConfig(FileBaseName, DefaultPresetName)
        .CurrentPresetName = .ConfigDictionary(CURRENT_KEY)
    End With
    Set Init = Me
End Function

Private Sub Class_Terminate()
    If Not This.FileSpec = vbNullString Then
        ForceSave
    End If
End Sub

'===============================================================================
' # Public

Public Property Get Config() As Scripting.IDictionary
    Set Config = This.ConfigDictionary
End Property

Public Property Get CurrentName() As String
    CurrentName = This.CurrentPresetName
End Property
Public Property Let CurrentName(RHS As String)
    This.CurrentPresetName = RHS
End Property

Public Property Get CurrentPreset() As Scripting.IDictionary
    CheckCurrentPreset
    Set CurrentPreset = Presets(This.CurrentPresetName)
End Property

'@DefaultMember
Public Property Get Item(ByVal Key As String) As Variant
    CheckCurrentPreset
    Assign _
        Item, Presets(This.CurrentPresetName)(Key)
End Property
Public Property Let Item(ByVal Key As String, RHS As Variant)
Attribute Item.VB_UserMemId = 0
    CheckCurrentPreset
    Presets(This.CurrentPresetName)(Key) = RHS
End Property
Public Property Set Item(ByVal Key As String, RHS As Variant)
    CheckCurrentPreset
    Set Presets(This.CurrentPresetName)(Key) = RHS
End Property

Public Property Get Presets() As Scripting.IDictionary
    Set Presets = This.ConfigDictionary(PRESETS_DICTIONARY_KEY)
End Property

Public Sub ForceSave()
    If This.ReadOnly Then Exit Sub
    WriteFile _
        JsonConverter.Create.ConvertToJson(This.ConfigDictionary), _
        This.FileSpec
End Sub

Public Property Get Self() As PresetsConfig
    Set Self = Me
End Property

'===============================================================================
' # Helpers

Private Function GetConfig( _
                     ByVal FileBaseName As String, _
                     ByVal DefaultPresetName As String _
                ) As Dictionary
    This.FileSpec = GetConfigFile(FileBaseName)
    If FileExists(This.FileSpec) Then
        Set GetConfig = JsonConverter.Create.ParseJson(ReadFile(This.FileSpec))
    Else
        Set GetConfig = CreateConfig(DefaultPresetName)
    End If
End Function

Private Function GetConfigFile(ByVal FileBaseName As String) As String
    GetConfigFile = GMSManager.UserGMSPath & FileBaseName & ".json"
End Function

Private Function CreateConfig( _
                    ByVal DefaultPresetName As String _
                ) As Scripting.IDictionary
    Set CreateConfig = New Scripting.Dictionary
    With CreateConfig
        .Add CURRENT_KEY, DefaultPresetName
        .Add _
            PRESETS_DICTIONARY_KEY, _
            CreatePreset(DefaultPresetName)
    End With
End Function

Private Function CreatePreset( _
                    ByVal PresetName As String _
                ) As Scripting.IDictionary
    Set CreatePreset = New Scripting.Dictionary
    With CreatePreset
        .Add PresetName, This.FallbackPreset
    End With
End Function

Private Sub CheckCurrentPreset()
    If VBA.IsObject(Presets(This.CurrentPresetName)) Then
        If TypeOf Presets(This.CurrentPresetName) Is Scripting.IDictionary Then Exit Sub
    End If
    Set Presets(This.CurrentPresetName) = New Scripting.Dictionary
End Sub
