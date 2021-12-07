Attribute VB_Name = "BitmapTools"
'===============================================================================
' ������           : BitmapTools
' ������           : 2021.12.07
' �����            : https://vk.com/elvin_macro/BitmapTools
'                    https://github.com/elvin-nsk/BitmapTools
' �����            : elvin-nsk (me@elvin.nsk.ru, https://vk.com/elvin_macro)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

'===============================================================================

Sub SendToEditor()

  If RELEASE Then On Error GoTo Catch
  
  Dim Context As BitmapContext
  With BitmapContext.CreateOrNotify
    If .IsError Then
      Exit Sub
    Else
      Set Context = .SuccessValue
    End If
  End With
    
  Dim BitmapFile As String
  With Helpers.GetNewTempBitmapFileSpec(Context.Document.FileName, _
                                        Context.BitmapShape.StaticID)
    If .IsError Then
      VBA.MsgBox "�� ������� ������� ��������� ����", vbCritical
      Exit Sub
    Else
      BitmapFile = .SuccessValue
    End If
  End With
  
  lib_elvin.BoostStart "�������������� ����������� �� ������� ���������", _
                       BitmapTools.RELEASE
  
  Helpers.SendToEditor Context.BitmapShape, BitmapFile
      
Finally:
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "������"
  Resume Finally
  
End Sub

Sub UpdateAfterEdit()

  If RELEASE Then On Error GoTo Catch
  
  Dim Context As BitmapContext
  With BitmapContext.CreateOrNotify
    If .IsError Then
      Exit Sub
    Else
      Set Context = .SuccessValue
    End If
  End With
  
  Dim BitmapFile As String
  With Helpers.GetCurrentTempBitmapFileSpec(Context.Document.FileName, _
                                            Context.BitmapShape.StaticID)
    If .IsError Then
      VBA.MsgBox "�� ������� ����� ��������� ����", vbCritical
      Exit Sub
    Else
      BitmapFile = .SuccessValue
    End If
  End With
  
  lib_elvin.BoostStart "���������� �����������", _
                       BitmapTools.RELEASE
  
  Helpers.UpdateAfterEdit Context.BitmapShape, BitmapFile
      
Finally:
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "������"
  Resume Finally
  
End Sub

Sub SendToEditorAndUpdate()

  If RELEASE Then On Error GoTo Catch

  Dim Context As BitmapContext
  With BitmapContext.CreateOrNotify
    If .IsError Then
      Exit Sub
    Else
      Set Context = .SuccessValue
    End If
  End With
    
  Dim BitmapFile As String
  With Helpers.GetNewTempBitmapFileSpec(Context.Document.FileName, _
                                        Context.BitmapShape.StaticID)
    If .IsError Then
      VBA.MsgBox "�� ������� ������� ��������� ����", vbCritical
      Exit Sub
    Else
      BitmapFile = .SuccessValue
    End If
  End With

  lib_elvin.BoostStart "�������������� ����������� �� ������� ���������", _
                        BitmapTools.RELEASE
                        
  Helpers.SendToEditor Context.BitmapShape, BitmapFile
  
  With New UpdateDialogView
    .FileSpec = BitmapFile
    .Show
    If .IsCancel Then
      GoTo Finally
    ElseIf .IsCancelAndDelete Then
      VBA.Kill BitmapFile
      GoTo Finally
    End If
  End With
  
  If Not lib_elvin.FileExists(BitmapFile) Then
    VBA.MsgBox "��������� ���� �� ������", vbCritical
    GoTo Finally
  End If
  Helpers.UpdateAfterEdit Context.BitmapShape, BitmapFile
  
Finally:
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "������"
  Resume Finally

End Sub

Sub RemoveCroppingPath()

  If RELEASE Then On Error GoTo Catch
  
  Dim Context As BitmapContext
  With BitmapContext.CreateOrNotify
    If .IsError Then
      Exit Sub
    Else
      Set Context = .SuccessValue
    End If
  End With
  
  lib_elvin.BoostStart "���������� �� �����", False
  Context.BitmapShape.Bitmap.ResetCropEnvelope
  
Finally:
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "������"
  Resume Finally

End Sub

Sub RemoveTransparency()

  If RELEASE Then On Error GoTo Catch
  
  Dim Context As BitmapContext
  With BitmapContext.CreateOrNotify
    If .IsError Then
      Exit Sub
    Else
      Set Context = .SuccessValue
    End If
  End With
  
  If Not Context.BitmapShape.Bitmap.Transparent Then Exit Sub
  
  lib_elvin.BoostStart "������ ������������", RELEASE
  With BitmapProcessor.Create(Context.BitmapShape)
    .Flatten
    .Shape.CreateSelection
  End With
  
Finally:
  lib_elvin.BoostFinish
  Exit Sub

Catch:
  VBA.MsgBox VBA.Err.Description, vbCritical, "������"
  Resume Finally

End Sub

Sub CheckTransparency()
  
  If ActiveDocument Is Nothing Then Exit Sub

  Dim Context As BitmapContext
  With BitmapContext.CreateOrNotify
    If .IsError Then
      Exit Sub
    Else
      Set Context = .SuccessValue
    End If
  End With
  
  With New TransparencyView
    .Transparent = Context.BitmapShape.Bitmap.Transparent
    .Show
  End With
  
End Sub
