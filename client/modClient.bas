Attribute VB_Name = "modClient"
Option Explicit

'***********************************************************************
' MODCLIENT
' This module subclasses the client form. This means that every time
' windows sends a message to the form, it gets passed through one of
' our functions. We can then deal with the messages as appropriate.
'***********************************************************************

'-----------------------------------------------------------------------
' API CALLS
'These are used to subclass a control.
'-----------------------------------------------------------------------
'Changes the address for the windows procedure, and returns the original value
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'Passes the message information on to the specified windows procedure
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
' VARIABLES
'-----------------------------------------------------------------------
'This const tells the SetWindowLong to change the address of the message
'handler
Const GWL_WNDPROC = -4

'This holds the memory location of the original windows handler
Private ProcPrev As Long
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
' WINDOWS HOOK/UNHOOK
Public Function HookForm(hwnd As Long)
    'This basically says that instead of using the normal windows handler
    'to deal with messages, use our WndProc function instead. It stores
    'the original memory location of the windows handler in ProcPrev
    ProcPrev = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WndProc)
End Function
Public Function UnhookForm(hwnd As Long)
    'Unhooks the form. Instead of going via our function, go back to
    'using your original windows handler
    SetWindowLong hwnd, GWL_WNDPROC, ProcPrev
End Function
'-----------------------------------------------------------------------

'-----------------------------------------------------------------------
' WINDOWS PROCEDURE
' Our own windows procedure. When HookForm is called, all messages are
' passed through this procedure.
'-----------------------------------------------------------------------
Function WndProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    'Depending on the message that is being sent, we decide how to deal with it...
    Select Case msg
        'Is it the message sent from the server form?
        Case modShared.WM_MYMESSAGE
            'If so, then call the function on frmClient
            frmClient.GotMessage wParam
                     
        Case Else
            'All other messages we don't care about, so we want to send
            'it on to it's normal location (which is stored in ProcPrev)
            WndProc = CallWindowProc(ProcPrev, hwnd, msg, wParam, lParam)
    End Select
End Function
