VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalizedStringsBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' cdrBrazilianPortuguese (1046) by Elieser Silveira (https://corelnaveia.com)
'===============================================================================

Option Explicit

Implements ILocalizedStrings

Private Strings As Scripting.IDictionary

Private Sub Class_Initialize()
    Set Strings = New Scripting.Dictionary
    With Strings
        
        .Add "Common.ErrNothingSelected", "Selecionar um bitmap"
        .Add "Common.ErrShapeIsNotBitmap", "A forma selecionada não é um bitmap"
        .Add "Common.ErrNoDocument", "Nenhum documento ativo"
        .Add "Common.ErrLayerDisabled", "A camada ativa está bloqueada"
        .Add "Common.ErrNoSelection", "Nenhuma seleção"
        
        .Add "BTools.SendToEditorUndo", "Editar bitmap no editor externo"
        .Add "BTools.UpdateAfterEditUndo", "Atualizar bitmap"
        .Add "BTools.RemoveCroppingPathUndo", "Liberar do caminho de corte"
        .Add "BTools.RemoveTransparencyUndo", "Remover transparência"
        .Add "BTools.ErrTempFileCreate", "Não foi possível criar arquivo temporário"
        .Add "BTools.ErrTempFileFind", "Arquivo temporário não encontrado"
        
        .Add "TransView.Caption", "Verificar transparência"
        .Add "TransView.Transparent", "Bitmap é transparente"
        .Add "TransView.Opaque", "Bitmap não é transparente"
        .Add "TransView.BtnOK", "OK"
        
        .Add "UDView.Caption", "Editar bitmap"
        .Add "UDView.Update", "Atualizar bitmap quando terminar de editar"
        .Add "UDView.Cancel", "Não atualize, deixe o arquivo {0} na pasta temporária"
        .Add "UDView.Delete", "Não atualizar, remover arquivo temporário"
        .Add "UDView.BtnUpdateStateWait", "Aguarde..."
        .Add "UDView.BtnUpdateStateReady", "Atualizar"
        .Add "UDView.BtnCancel", "Cancelar"
        .Add "UDView.BtnDelete", "Deletar"
        
        .Add "Settings.Caption", "Configurações"
        .Add "Settings.PresetsFrame", "Predefinições"
        .Add "Settings.WarningPresetNameAlreadyExists", "Já existe uma predefinição com o mesmo nome"
        .Add "Settings.WarningPresetNameIsEmpty", "O nome da predefinição não pode ficar vazio"
        .Add "Settings.WarningPresetRemove", "Excluir predefinição?"
        .Add "Settings.txtEditor", "Selecione o executável do editor de imagens"
        .Add "Settings.btnOk", "Fechar"

    End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
    Set ILocalizedStrings_Strings = Strings
End Property
