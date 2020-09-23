VERSION 5.00
Begin VB.Form frmClient 
   Caption         =   "frmClient"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   3765
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox listMessages 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   120
      TabIndex        =   1
      Top             =   1020
      Width           =   3555
   End
   Begin VB.Label Label1 
      Caption         =   $"frmClient.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3555
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***********************************************************************
' FRMCLIENT
' This form is hooked (subclassed) so that all windows messages sent to
' the form go through our function (which is in modClient). This can
' then check for our own custom message to arrive. When it does, it adds
' it to the listbox
'***********************************************************************

Private Sub Form_Load()
    'Set form title. This comes form the modShared module. This is important,
    'as the server identifies this window by it's title.
    Me.Caption = modShared.CLIENT_TITLE
    
    'Hooks the form.
    HookForm Me.hwnd
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Unhooks the form. Whenever you subclass a control, you should always kill
    'it at runtime
    UnhookForm Me.hwnd
End Sub

'This sub is called from the modClient when our message has been detected
Public Sub GotMessage(value As Long)
    listMessages.AddItem "Message recieved! Value : " & CStr(value)
End Sub
