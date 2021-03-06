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
        
        .Add "Common_ErrNothingSelected", "Selecionar um bitmap"
        .Add "Common_ErrShapeIsNotBitmap", "A forma selecionada n�o � um bitmap"
        
        .Add "BTools_SendToEditorUndo", "Editar bitmap no editor externo"
        .Add "BTools_UpdateAfterEditUndo", "Atualizar bitmap"
        .Add "BTools_RemoveCroppingPathUndo", "Liberar do caminho de corte"
        .Add "BTools_RemoveTransparencyUndo", "Remover transpar�ncia"
        .Add "BTools_ErrTempFileCreate", "N�o foi poss�vel criar arquivo tempor�rio"
        .Add "BTools_ErrTempFileFind", "Arquivo tempor�rio n�o encontrado"
        
        .Add "TransView_Caption", "Verificar transpar�ncia"
        .Add "TransView_Transparent", "Bitmap � transparente"
        .Add "TransView_Opaque", "Bitmap n�o � transparente"
        .Add "TransView_BtnOK", "OK"
        
        .Add "UDView_Caption", "Editar bitmap"
        .Add "UDView_Update", "Atualizar bitmap quando terminar de editar"
        .Add "UDView_Cancel", "N�o atualize, deixe o arquivo {0} na pasta tempor�ria"
        .Add "UDView_Delete", "N�o atualizar, remover arquivo tempor�rio"
        .Add "UDView_BtnUpdateStateWait", "Aguarde..."
        .Add "UDView_BtnUpdateStateReady", "Atualizar"
        .Add "UDView_BtnCancel", "Cancelar"
        .Add "UDView_BtnDelete", "Deletar"

    End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
    Set ILocalizedStrings_Strings = Strings
End Property
