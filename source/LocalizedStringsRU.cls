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
        
        .Add "Common_ErrNothingSelected", "�������� ��������� ������"
        .Add "Common_ErrShapeIsNotBitmap", "���������� ������ �� �������� �������"
    
        .Add "BTools_SendToEditorUndo", _
            "�������������� ����������� �� ������� ���������"
        .Add "BTools_UpdateAfterEditUndo", "���������� �����������"
        .Add "BTools_RemoveCroppingPathUndo", "���������� �� �����"
        .Add "BTools_RemoveTransparencyUndo", "������ ������������"
        .Add "BTools_ErrTempFileCreate", "�� ������� ������� ��������� ����"
        .Add "BTools_ErrTempFileFind", "�� ������� ����� ��������� ����"
    
        .Add "TransView_Caption", "�������� ������������"
        .Add "TransView_Transparent", "����������� � �������������"
        .Add "TransView_Opaque", "����������� ��� ������������"
        .Add "TransView_BtnOK", "��"
        
        .Add "UDView_Caption", "�������������� �����������"
        .Add "UDView_Update", _
            "�������� ����������� �� ���������� �����, ����� ��������� ��������������"
        .Add "UDView_Cancel", "�� ���������, �������� ���� {0} �� ��������� �����"
        .Add "UDView_Delete", "�� ���������, ������� ��������� ����"
        .Add "UDView_BtnUpdateStateWait", "��������"
        .Add "UDView_BtnUpdateStateReady", "��������"
        .Add "UDView_BtnCancel", "������"
        .Add "UDView_BtnDelete", "�������"

    End With
End Sub

'===============================================================================

Private Property Get ILocalizedStrings_Strings() As Scripting.IDictionary
    Set ILocalizedStrings_Strings = Strings
End Property
