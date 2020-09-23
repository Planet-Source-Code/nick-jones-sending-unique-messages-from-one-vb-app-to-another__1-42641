Attribute VB_Name = "modShared"
Option Explicit

'***********************************************************************
' MODSHARED
' This module is shared by both the client and server projects. This
' holds various const values and functions that both projects use.
'***********************************************************************

'-----------------------------------------------------------------------
' API CALLS
'-----------------------------------------------------------------------
'Registers a message as windows, returning a unique ID that can be used
'to identify the message
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
' VARIABLES etc
'-----------------------------------------------------------------------
'This is the name of the message we're going to send
Private Const MSG_MYMESSAGE = "MSG_TESTMESSAGE"

'This holds the title (caption property) of the client form. This is used
'when it comes to send a message to we can identify the program
Public Const CLIENT_TITLE = "Client Windows"
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
' FUNCTIONS
'-----------------------------------------------------------------------
'This function returns the msg id of the windows message
Public Function WM_MYMESSAGE() As Long
    'Static variable that holds the unique id of our message
    Static msg As Long
    
    'If this is the first time we're running the function,
    'register the message. Results the unique ID of the registerd
    'message
    If msg = 0 Then
        msg = RegisterWindowMessage(MSG_MYMESSAGE)
    End If
    
    'Return the result
    WM_MYMESSAGE = msg
End Function
'-----------------------------------------------------------------------
