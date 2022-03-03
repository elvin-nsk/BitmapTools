VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UpdateDialogView 
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
  Me.Caption = LocalizedStrings("UDView_Caption")
  btnUpdate.Caption = LocalizedStrings("UDView_BtnUpdateStateReady")
  btnCancel.Caption = LocalizedStrings("UDView_BtnCancel")
  btnDelete.Caption = LocalizedStrings("UDView_BtnDelete")
End Sub

Private Sub UserForm_Activate()
  txtUpdate = LocalizedStrings("UDView_Update")
  txtCancel = LocalizedStrings("UDView_Cancel", FileSpec)
  txtDelete = LocalizedStrings("UDView_Delete")
  With btnUpdate
    .Enabled = False
    .Caption = LocalizedStrings("UDView_BtnUpdateStateWait")
    DoEvents
    Sleep 2000
    .Caption = LocalizedStrings("UDView_BtnUpdateStateReady")
    .Enabled = True
    .SetFocus
  End With
End Sub

Private Sub btnUpdate_Click()
  FormÎÊ
End Sub

Private Sub btnCancel_Click()
  FormCancel
End Sub

Private Sub btnDelete_Click()
  FormCancelAndDelete
End Sub

'===============================================================================

Private Sub FormÎÊ()
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

Private Sub GuardRangeDbl(TextBox As MSForms.TextBox, _
                          ByVal Min As Double, _
                          Optional ByVal Max As Double = 1.79769313486231E+308)
  With TextBox
    If .Value = "" Then .Value = VBA.CStr(Min)
    If VBA.CDbl(.Value) > Max Then .Value = VBA.CStr(Max)
    If VBA.CDbl(.Value) < Min Then .Value = VBA.CStr(Min)
  End With
End Sub

Private Sub GuardRangeLng(TextBox As MSForms.TextBox, _
                          ByVal Min As Long, _
                          Optional ByVal Max As Long = 2147483647)
  With TextBox
    If .Value = "" Then .Value = VBA.CStr(Min)
    If VBA.CLng(.Value) > Max Then .Value = VBA.CStr(Max)
    If VBA.CLng(.Value) < Min Then .Value = VBA.CStr(Min)
  End With
End Sub

Private Sub UserForm_QueryClose(Ñancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Ñancel = True
    FormCancel
  End If
End Sub

