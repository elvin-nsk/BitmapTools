VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UpdateDialogView 
   Caption         =   "Редактирование изображения"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   OleObjectBlob   =   "UpdateDialogView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UpdateDialogView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

#If VBA7 Then
  Private Declare PtrSafe Sub Sleep Lib "kernel32" _
    (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
  Private Declare Sub Sleep Lib "kernel32" _
    (ByVal dwMilliseconds as Long) 'For 32 Bit Systems
#End If

'===============================================================================

Public FileSpec As String

Public IsUpdate As Boolean
Public IsCancel As Boolean
Public IsCancelAndDelete As Boolean

'===============================================================================

Private Sub UserForm_Initialize()
  '
End Sub

Private Sub UserForm_Activate()
  txtUpdate = "Обновите изображение из временного файла, когда закончите редактирование"
  txtCancel = "Не обновлять, оставить файл " & FileSpec & " во временной папке"
  txtDelete = "Не обновлять, удалить временный файл"
  With btnUpdate
    .Enabled = False
    .Caption = "Ожидание"
    DoEvents
    Sleep 2000
    .Caption = "Обновить"
    .Enabled = True
    .SetFocus
  End With
End Sub

Private Sub btnUpdate_Click()
  FormОК
End Sub

Private Sub btnCancel_Click()
  FormCancel
End Sub

Private Sub btnDelete_Click()
  FormCancelAndDelete
End Sub

'===============================================================================

Private Sub FormОК()
  Me.Hide
  IsUpdate = True
End Sub

Private Sub FormCancel()
  Me.Hide
  IsCancel = True
End Sub

Private Sub FormCancelAndDelete()
  Me.Hide
  IsCancelAndDelete = True
End Sub

'===============================================================================

Private Sub GuardInt(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case VBA.Asc("0") To VBA.Asc("9")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub GuardNum(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case VBA.Asc("0") To VBA.Asc("9")
    Case VBA.Asc(",")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub GuardRangeDbl(TextBox As MSForms.TextBox, ByVal Min As Double, Optional ByVal Max As Double = 2147483647)
  With TextBox
    If .Value = "" Then .Value = VBA.CStr(Min)
    If VBA.CDbl(.Value) > Max Then .Value = VBA.CStr(Max)
    If VBA.CDbl(.Value) < Min Then .Value = VBA.CStr(Min)
  End With
End Sub

Private Sub GuardRangeLng(TextBox As MSForms.TextBox, ByVal Min As Long, Optional ByVal Max As Long = 2147483647)
  With TextBox
    If .Value = "" Then .Value = VBA.CStr(Min)
    If VBA.CLng(.Value) > Max Then .Value = VBA.CStr(Max)
    If VBA.CLng(.Value) < Min Then .Value = VBA.CStr(Min)
  End With
End Sub

Private Sub UserForm_QueryClose(Сancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Сancel = True
    FormCancel
  End If
End Sub

