VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LocalizedStringsEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================
' cdrEnglishUS (1033)
'===============================================================================

'@PredeclaredId
Option Explicit

Public Strings As Scripting.IDictionary

Private Sub Class_Initialize()
  Set Strings = New Scripting.Dictionary
  With Strings
    
    .Add "Common_ErrNothingSelected", "Select bitmap"
    .Add "Common_ErrShapeIsNotBitmap", "Selected shape is not bitmap"
  
    .Add "BTools_SendToEditorUndo", "Edit bitmap in external editor"
    .Add "BTools_UpdateAfterEditUndo", "Update bitmap"
    .Add "BTools_RemoveCroppingPathUndo", "Release from croping path"
    .Add "BTools_RemoveTransparencyUndo", "Remove transparency"
    .Add "BTools_ErrTempFileCreate", "Unable to create temporary file"
    .Add "BTools_ErrTempFileFind", "Temporary file not found"
  
    .Add "TransView_Caption", "Check transparency"
    .Add "TransView_Transparent", "Bitmap is transparent"
    .Add "TransView_Opaque", "Bitmap is opaque"
    .Add "TransView_BtnOK", "OK"
    
    .Add "UDView_Caption", "Edit bitmap"
    .Add "UDView_Update", "Update bitmap when finished editing"
    .Add "UDView_Cancel", "Do not update, leave file {0} in temporary folder"
    .Add "UDView_Delete", "Do not update, remove temporary file"
    .Add "UDView_BtnUpdateStateWait", "Wait"
    .Add "UDView_BtnUpdateStateReady", "Update"
    .Add "UDView_BtnCancel", "Cancel"
    .Add "UDView_BtnDelete", "Delete"

  End With
End Sub
