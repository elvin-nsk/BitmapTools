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
        .Add "Common.ErrShapeIsNotBitmap", "A forma selecionada n�o � um bitmap"
        .Add "Common.ErrNoDocument", "Nenhum documento ativo"
        .Add "Common.ErrLayerDisabled", "A camada ativa est� bloqueada"
        .Add "Common.ErrNoSelection", "Nenhuma sele��o"
        
        .Add "BTools.SendToEditorUndo", "Editar bitmap no editor externo"
        .Add "BTools.UpdateAfterEditUndo", "Atualizar bitmap"
        .Add "BTools.RemoveCroppingPathUndo", "Liberar do caminho de corte"
        .Add "BTools.RemoveTransparencyUndo", "Remover transpar�ncia"
        .Add "BTools.ErrTempFileCreate", "N�o foi poss�vel criar arquivo tempor�rio"
        .Add "BTools.ErrTempFileFind", "Arquivo tempor�rio n�o encontrado"
        
        .Add "TransView.Caption", "Verificar transpar�ncia"
        .Add "TransView.Transparent", "Bitmap � transparente"
        .Add "TransView.Opaque", "Bitmap n�o � transparente"
        .Add "TransView.BtnOK", "OK"
        
        .Add "UDView.Caption", "Editar bitmap"
        .Add "UDView.Update", "Atualizar bitmap quando terminar de editar"
        .Add "UDView.Cancel", "N�o atualize, deixe o arquivo {0} na pasta tempor�ria"
        .Add "UDView.Delete", "N�o atualizar, remover arquivo tempor�rio"
        .Add "UDView.BtnUpdateStateWait", "Aguarde..."
        .Add "UDView.BtnUpdateStateReady", "Atualizar"
        .Add "UDView.BtnCancel", "Cancelar"
        .Add "UDView.BtnDelete", "Deletar"
        
        .Add "Settings.Caption", "Configura��es"
        .Add "Settings.PresetsFrame", "Predefini��es"
        .Add "Settings.WarningPresetNameAlreadyExists", "J� existe uma predefini��o com o mesmo nome"
        .Add "Settings.WarningPresetNameIsEmpty", "O nome da predefini��o n�o pode ficar vazio"
        .Add "Settings.WarningPresetRemove", "Excluir predefini��o?"
        .Add "Settings.txtEditor", "Selecione o execut�vel do editor de imagens"
        .Add "Settings.btnOk", "Fechar"

    End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
    Set ILocalizedStrings_Strings = Strings
End Property
