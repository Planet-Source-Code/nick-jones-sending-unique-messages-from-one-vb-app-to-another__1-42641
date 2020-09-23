VERSION 5.00
Begin VB.Form frmServer 
   Caption         =   "Server Program"
   ClientHeight    =   1995
   ClientLeft      =   2955
   ClientTop       =   4740
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   3900
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1500
      TabIndex        =   0
      Text            =   "12345"
      Top             =   1380
      Width           =   915
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send Message"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   1380
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Enter value :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   $"frmServer.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   180
      TabIndex        =   2
      Top             =   120
      Width           =   3555
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************
' FRMSERVER
' This form, when the Send Message button is clicked, tries to find the
' windows handle of the client program, and then sends it a message.
'***********************************************************************

'----------------------------------------------------------------------------
' API CALLS
'----------------------------------------------------------------------------
'Sends a windows message to a specific form. hwnd is the windows handle we
'want to send the message to. The wMsg is the name of the message we are sending
'and wParam and lParam can contain any parameters we want to send with it. In
'this case we stick the value entered in txtMessage in wParam. This can be
'easily modified to send text if required.
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

'Finds a window. We pass it the name of the form, and it returns the 'handle'
'(hwnd) of the form. We can then use this handle in the SendMessageLong API
'call to send that form a message
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'----------------------------------------------------------------------------

'----------------------------------------------------------------------------
' SEND MESSAGE
' Tries to send a message to our client window
'----------------------------------------------------------------------------
Private Sub cmdSend_Click()
    'Check that the user entered a valid number to send...
    If Not IsNumeric(Me.txtMessage) Then
        MsgBox "Try entering a valid number..."
    Else
        'Convert to a long
        Me.txtMessage = CLng(Me.txtMessage)
    
        'Holds the hwnd of the client window
        Dim hwndTarget As Long
        
        'Uses the FindWindow API to try and find the client window using the title
        'of the form. Returns 0 if the window wasn't found
        hwndTarget = FindWindow(vbNullString, modShared.CLIENT_TITLE)
        
        'Did we find the window?
        If hwndTarget <> 0 Then
            'If we did, then send a message with our message and the value
            Call SendMessageLong(hwndTarget, modShared.WM_MYMESSAGE, Me.txtMessage, 0)
        Else
            'If not, then show an error...
            MsgBox "The target windows could not be found. Make sure you have the client window running and try again.", vbOKOnly + vbInformation, "Send message failure..."
        End If
    End If
End Sub

