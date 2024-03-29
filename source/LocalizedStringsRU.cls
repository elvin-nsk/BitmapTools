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
        
        .Add "Common.ErrNothingSelected", "�������� ��������� ������"
        .Add "Common.ErrShapeIsNotBitmap", "���������� ������ �� �������� �������"
        .Add "Common.ErrNoDocument", "��� ��������� ���������"
        .Add "Common.ErrLayerDisabled", "������� ���� ������"
        .Add "Common.ErrNoSelection", "�������� �������"
    
        .Add "BTools.SendToEditorUndo", _
            "�������������� ����������� �� ������� ���������"
        .Add "BTools.UpdateAfterEditUndo", "���������� �����������"
        .Add "BTools.RemoveCroppingPathUndo", "���������� �� �����"
        .Add "BTools.RemoveTransparencyUndo", "������ ������������"
        .Add "BTools.ErrTempFileCreate", "�� ������� ������� ��������� ����"
        .Add "BTools.ErrTempFileFind", "�� ������� ����� ��������� ����"
    
        .Add "TransView.Caption", "�������� ������������"
        .Add "TransView.Transparent", "����������� � �������������"
        .Add "TransView.Opaque", "����������� ��� ������������"
        .Add "TransView.BtnOK", "��"
        
        .Add "UDView.Caption", "�������������� �����������"
        .Add "UDView.Update", _
            "�������� ����������� �� ���������� �����, ����� ��������� ��������������"
        .Add "UDView.Cancel", "�� ���������, �������� ���� {0} �� ��������� �����"
        .Add "UDView.Delete", "�� ���������, ������� ��������� ����"
        .Add "UDView.BtnUpdateStateWait", "��������"
        .Add "UDView.BtnUpdateStateReady", "��������"
        .Add "UDView.BtnCancel", "������"
        .Add "UDView.BtnDelete", "�������"
        
        .Add "Settings.Caption", "���������"
        .Add "Settings.PresetsFrame", "�������������"
        .Add "Settings.WarningPresetNameAlreadyExists", "������ � ����� ��������� ��� ����������"
        .Add "Settings.WarningPresetNameIsEmpty", "�������� ������� �� ����� ���� ������"
        .Add "Settings.WarningPresetRemove", "������� ������?"
        .Add "Settings.txtEditor", "�������� ����������� ���� ��������� �����������"
        .Add "Settings.btnOk", "�������"

    End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
    Set ILocalizedStrings_Strings = Strings
End Property
