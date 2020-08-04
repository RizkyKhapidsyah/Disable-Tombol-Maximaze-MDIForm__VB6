VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Disable Tombol Maximaze di MDIForm"
   ClientHeight    =   4185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32" _
Alias "GetWindowLongA" (ByVal hWnd As Long, _
ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" _
Alias "SetWindowLongA" (ByVal hWnd As Long, _
ByVal nIndex As Long, ByVal dwNewLong As Long) _
As Long

Private Const GWL_STYLE = (-16)
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000

Public Sub NoMaxBox(f As MDIForm)
  Dim l As Long
  l = GetWindowLong(f.hWnd, GWL_STYLE)
  l = l And Not (WS_MAXIMIZEBOX)
  l = SetWindowLong(f.hWnd, GWL_STYLE, l)
End Sub

Public Sub NoMinBox(f As MDIForm)
  Dim l As Long
  l = GetWindowLong(f.hWnd, GWL_STYLE)
  l = l And Not (WS_MINIMIZEBOX)
  l = SetWindowLong(f.hWnd, GWL_STYLE, l)
End Sub

Private Sub MDIForm_Load()
  NoMaxBox Me
  'NoMinBox Me
End Sub

Private Sub MDIForm_Resize()
  'Ganti 4800 di bawah dgn lebar form yg fix Anda
  'tentukan Ganti 3600 di bawah dgn tinggi form yg fix
  'Anda tentukan
  If Me.Width <> 4800 Then Me.Width = 4800
  If Me.Height <> 3600 Then Me.Height = 3600
End Sub


