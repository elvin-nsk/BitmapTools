VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalizedStringsRU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' cdrRussian (1049)
'===============================================================================

'@PredeclaredId
Option Explicit

Public Strings As Scripting.IDictionary

Private Sub Class_Initialize()
  Set Strings = New Scripting.Dictionary
  With Strings
    
    .Add "Common_ErrNothingSelected", "Выделите растровый объект"
    .Add "Common_ErrShapeIsNotBitmap", "Выделенный объект не является растром"
  
    .Add "BTools_SendToEditorUndo", _
      "Редактирование изображения во внешнем редакторе"
    .Add "BTools_UpdateAfterEditUndo", "Обновление изображения"
    .Add "BTools_RemoveCroppingPathUndo", "Освободить из кропа"
    .Add "BTools_RemoveTransparencyUndo", "Убрать прозрачность"
    .Add "BTools_ErrTempFileCreate", "Не удалось создать временный файл"
    .Add "BTools_ErrTempFileFind", "Не удалось найти временный файл"
  
    .Add "TransView_Caption", "Проверка прозрачности"
    .Add "TransView_Transparent", "Изображение с прозрачностью"
    .Add "TransView_Opaque", "Изображение без прозрачности"
    .Add "TransView_BtnOK", "ОК"
    
    .Add "UDView_Caption", "Редактирование изображения"
    .Add "UDView_Update", _
      "Обновите изображение из временного файла, когда закончите редактирование"
    .Add "UDView_Cancel", "Не обновлять, оставить файл {0} во временной папке"
    .Add "UDView_Delete", "Не обновлять, удалить временный файл"
    .Add "UDView_BtnUpdateStateWait", "Ожидание"
    .Add "UDView_BtnUpdateStateReady", "Обновить"
    .Add "UDView_BtnCancel", "Отмена"
    .Add "UDView_BtnDelete", "Удалить"

  End With
End Sub
