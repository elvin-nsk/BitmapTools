VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalizedStringsRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' cdrRussian (1049)
'===============================================================================

Option Explicit

Implements ILocalizedStrings

Private Strings As Scripting.IDictionary

Private Sub Class_Initialize()
    Set Strings = New Scripting.Dictionary
    With Strings
        
        .Add "Common.ErrNothingSelected", "Выделите растровый объект"
        .Add "Common.ErrShapeIsNotBitmap", "Выделенный объект не является растром"
        .Add "Common.ErrNoDocument", "Нет активного документа"
        .Add "Common.ErrLayerDisabled", "Текущий слой закрыт"
        .Add "Common.ErrNoSelection", "Выделите объекты"
    
        .Add "BTools.SendToEditorUndo", _
            "Редактирование изображения во внешнем редакторе"
        .Add "BTools.UpdateAfterEditUndo", "Обновление изображения"
        .Add "BTools.RemoveCroppingPathUndo", "Освободить из кропа"
        .Add "BTools.RemoveTransparencyUndo", "Убрать прозрачность"
        .Add "BTools.ErrTempFileCreate", "Не удалось создать временный файл"
        .Add "BTools.ErrTempFileFind", "Не удалось найти временный файл"
    
        .Add "TransView.Caption", "Проверка прозрачности"
        .Add "TransView.Transparent", "Изображение с прозрачностью"
        .Add "TransView.Opaque", "Изображение без прозрачности"
        .Add "TransView.BtnOK", "ОК"
        
        .Add "UDView.Caption", "Редактирование изображения"
        .Add "UDView.Update", _
            "Обновите изображение из временного файла, когда закончите редактирование"
        .Add "UDView.Cancel", "Не обновлять, оставить файл {0} во временной папке"
        .Add "UDView.Delete", "Не обновлять, удалить временный файл"
        .Add "UDView.BtnUpdateStateWait", "Ожидание"
        .Add "UDView.BtnUpdateStateReady", "Обновить"
        .Add "UDView.BtnCancel", "Отмена"
        .Add "UDView.BtnDelete", "Удалить"
        
        .Add "Settings.Caption", "Настройки"
        .Add "Settings.PresetsFrame", "Предустановки"
        .Add "Settings.WarningPresetNameAlreadyExists", "Пресет с таким названием уже существует"
        .Add "Settings.WarningPresetNameIsEmpty", "Название пресета не может быть пустым"
        .Add "Settings.WarningPresetRemove", "Удалить пресет?"
        .Add "Settings.txtEditor", "Выберите исполняемый файл редактора изображений"
        .Add "Settings.btnOk", "Закрыть"

    End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
    Set ILocalizedStrings_Strings = Strings
End Property
