VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TransparencyView 
   Caption         =   "Проверка прозрачности"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4620
   OleObjectBlob   =   "TransparencyView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TransparencyView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public Transparent As Boolean

Public IsOk As Boolean
Public IsCancelled As Boolean

'===============================================================================

Private Sub UserForm_Initialize()
  '
End Sub

Private Sub UserForm_Activate()
  If Transparent Then
    imgNonTransparent.Visible = False
    Text = "Изображение с прозрачностью"
  Else
    imgTransparent.Visible = False
    Text = "Изображение без прозрачности"
  End If
End Sub


Private Sub btnCancel_Click()
  FormCancel
End Sub

Private Sub btnOK_Click()
  FormОК
End Sub

'===============================================================================

Private Sub FormОК()
  Me.Hide
  IsOk = True
End Sub

Private Sub FormCancel()
  Me.Hide
  IsCancelled = True
End Sub

'===============================================================================

Private Sub UserForm_QueryClose(Сancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Сancel = True
    FormCancel
  End If
End Sub

