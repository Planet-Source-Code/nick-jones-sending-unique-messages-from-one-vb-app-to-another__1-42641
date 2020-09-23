-------------------------------------------------------
Sending Unique Messages From one VB App To Another
-------------------------------------------------------
Nick Jones :: nickjones@eidosnet.co.uk
-------------------------------------------------------

Demonstrates how to communicate between your applications
by sending your own windows messages along with your
own variable. Properly commented code. I couldn't find a 
similar program on PSC, so I thought I'd write this up. 
Hope someone finds this useful.

The code includes 2 projects, 'a client' and a 'server'.
Using the RegisterWindowMessage API the server creates
a unique windows message, then through the FindWindow 
API it checks to see if the 'client' is running. If it
is, it sends a message using SendMessageLong. The client
is subclassed and adds any messages it detects to a listbox.

The code is well commented, full explaining subclassing, 
the API's it uses and why, so even a notice should be
able to understand all this.


-------------------------------------------------------

To run the project, boot 'em both up in seperate VB
windows, and run them both. Get both of them on the
screen (like in the screenshot). 

-------------------------------------------------------