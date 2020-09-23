VERSION 5.00
Begin VB.Form frmClient 
   Caption         =   "frmClient"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    'Set form title
    Me.Caption = modShared.CLIENT_TITLE
    HookForm Me.hwnd
End Sub


Private Sub Form_Unload(Cancel As Integer)
    UnhookForm Me.hwnd
End Sub
