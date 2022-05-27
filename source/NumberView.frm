VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NumberView 
   Caption         =   "Íîìåð"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2160
   OleObjectBlob   =   "NumberView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NumberView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public IsOK As Boolean
Public Number As Long

'===============================================================================

Private Sub UserForm_Activate()
  With tbNumber
    .Value = 0
    .SelStart = 0
    .SelLength = VBA.Len(.Text)
  End With
End Sub
Private Sub tbNumber_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  If KeyCode = 13 Then FormÎÊ
End Sub
Private Sub tbNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  OnlyInt KeyAscii
End Sub
Private Sub tbNumber_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
  tbNumber_AfterUpdate
End Sub
Private Sub tbNumber_AfterUpdate()
  CheckRangeLng tbNumber, 0
End Sub

Private Sub cbOverlap_Click()
  IsOverlap = cbOverlap.Value
  VisibilityCheck
End Sub

Private Sub btnCancel_Click()
  FormCancel
End Sub

Private Sub btnOK_Click()
  FormÎÊ
End Sub

'===============================================================================

Private Sub FormÎÊ()
  Me.Hide
  IsOK = True
  Number = CLng(tbNumber.Value)
End Sub

Private Sub FormCancel()
  Me.Hide
End Sub

'===============================================================================

Private Sub OnlyInt(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub OnlyNum(ByVal KeyAscii As MSForms.ReturnInteger)
  Select Case KeyAscii
    Case Asc("0") To Asc("9")
    Case Asc(",")
    Case Else
      KeyAscii = 0
  End Select
End Sub

Private Sub CheckRangeDbl(TextBox As MSForms.TextBox, ByVal Min As Double, Optional ByVal Max As Double = 2147483647)
  With TextBox
    If .Value = "" Then .Value = CStr(Min)
    If CDbl(.Value) > Max Then .Value = CStr(Max)
    If CDbl(.Value) < Min Then .Value = CStr(Min)
  End With
End Sub

Private Sub CheckRangeLng(TextBox As MSForms.TextBox, ByVal Min As Long, Optional ByVal Max As Long = 2147483647)
  With TextBox
    If .Value = "" Then .Value = CStr(Min)
    If CLng(.Value) > Max Then .Value = CStr(Max)
    If CLng(.Value) < Min Then .Value = CStr(Min)
  End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  If CloseMode = VbQueryClose.vbFormControlMenu Then
    Cancel = True
    FormCancel
  End If
End Sub
