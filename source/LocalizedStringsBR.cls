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
        .Add "Common_ErrShapeIsNotBitmap", "A forma selecionada não é um bitmap"
        
        .Add "BTools_SendToEditorUndo", "Editar bitmap no editor externo"
        .Add "BTools_UpdateAfterEditUndo", "Atualizar bitmap"
        .Add "BTools_RemoveCroppingPathUndo", "Liberar do caminho de corte"
        .Add "BTools_RemoveTransparencyUndo", "Remover transparência"
        .Add "BTools_ErrTempFileCreate", "Não foi possível criar arquivo temporário"
        .Add "BTools_ErrTempFileFind", "Arquivo temporário não encontrado"
        
        .Add "TransView_Caption", "Verificar transparência"
        .Add "TransView_Transparent", "Bitmap é transparente"
        .Add "TransView_Opaque", "Bitmap não é transparente"
        .Add "TransView_BtnOK", "OK"
        
        .Add "UDView_Caption", "Editar bitmap"
        .Add "UDView_Update", "Atualizar bitmap quando terminar de editar"
        .Add "UDView_Cancel", "Não atualize, deixe o arquivo {0} na pasta temporária"
        .Add "UDView_Delete", "Não atualizar, remover arquivo temporário"
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
