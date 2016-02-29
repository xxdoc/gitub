Attribute VB_Name = "dropfile"
  Option Explicit
'更多实用源码请访问[VB梦工厂]WWW.51xue8xue8.com
  Public Const MAX_PATH As Long = 260&

  Public Const WM_DROPFILES As Long = &H233&

  Public procOld As Long
  
  '
  Public Declare Function CallWindowProc& Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc&, _
                                                    ByVal hWnd&, ByVal Msg&, ByVal wParam&, ByVal lParam&)

  Public Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hWnd&, ByVal fAccept&)
                               
  Public Declare Function DragQueryFile& Lib "shell32.dll" Alias "DragQueryFileA" (ByVal hDrop&, ByVal iFile&, _
                                                                                  ByVal lpszFile$, ByVal cch&)
  Public Declare Sub DragFinish Lib "shell32.dll" (ByVal hDrop&)
  
    Public Const GWL_WNDPROC As Long = (-4&)
    
  ' API call to alter the class data for this window
  Public Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hWnd&, ByVal nIndex&, ByVal dwNewLong&)



'WARNING!!!! WARNING!!!! WARNING!!!! WARNING!!!! WARNING!!!! WARNING!!!!
'
' Do NOT try to step through this function in debug mode!!!!
' You WILL crash!!!  Also, do NOT set any break points in this function!!!
' You WILL crash!!!  Subclassing is non-trivial and should be handled with
' EXTREAME care!!!
'
' There are ways to use a "Debug" dll to allow you to set breakpoints in
' subclassed code in the IDE but this was not implimented for this demo.
'
'WARNING!!!! WARNING!!!! WARNING!!!! WARNING!!!! WARNING!!!! WARNING!!!!
  
Public Function WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, _
                                              ByVal wParam As Long, ByVal lParam As Long) As Long
    
  ' this is *our* implimentation of the message handling routine
  
  ' determine which message was recieved
  Select Case iMsg
    
    ' grab the message that tells us when a file was dropped on the picturebox
    Case WM_DROPFILES
      ' call the sup that we defined in the form module passing wParam which is the handle to the file
    Mainform.DropFiles wParam
      
      ' return zero to windows and get out
      WindowProc = False
      Exit Function
      
  End Select
  
  ' pass all messages on to VB and then return the value to windows
  WindowProc = CallWindowProc(procOld, hWnd, iMsg, wParam, lParam)

End Function

' What is subclassing anyway?
'
' Windows runs on "messages".  A message is a unique value that, when
' recieved by a window or the operating system, tells either that
' something has happened and that an action of some sort needs to be
' taken.  Sort of like your nervous system passing feeling messages
' to your brain and the brain passing movement messages to your body.
'
' So, each window has what's called a message handler.  This is a
' function where all of the messages FROM Windows are recieved.  Every
' window has one.  I mean EVERY window.  That means every button, textbox,
' picturebox, form, etc...  Windows keeps track of where the message
' handler (called a WindowProc [short for PROCedure]) in a "Class"
' structure associated with each window handle (otherwise known as hWnd).
'
' What happens when a window is subclassed is that you insert a new
' window procedure in line with the original window procedure.  In other
' words, Windows sends the messages for the given window to YOUR WindowProc
' FIRST where you are responsible for handling any messages you want to
' handle.  Then you MUST pass the remaining messages on to the default
' WindoProc.  So it looks like this:
'
'  Windows Message Sender --> Your WindowProc --> Default WindowProc
'
' A window can be subclassed MANY times so it could look like this:
'
'  Windows Message Sender --> Your WindowProc --> Another WindowProc _
'  --> Yet Another WindowProc --> Default WindowProc
'
' You can also change the order of when you respond to a message by
' where in your routine you pass the message on to the default WindowProc.
' Let's say that you want to draw something on the window AFTER the
' default WindowProc handles the WM_PAINT message.  This is easily done
' by calling the default proc before you do your drawing.   Like so:
'
' Public Function WindowProc(Byval hWnd, Byval etc....)
'
'   Select Case iMsg
'     Case SOME_MESSAGE
'       DoSomeStuff
'
'     Case WM_PAINT
'       ' pass the message to the defproc FIRST
'       WindowProc = CallWindowProc(procOld, hWnd, iMsg, wParam, lParam)
'
'       DoDrawingStuff ' <- do your drawing
'
'       Exit Function ' <- exit since we already passed the
'                     '    measage to the defproc
'
'   End Select
'
'   ' pass all messages on to VB and then return the value to windows
'   WindowProc = CallWindowProc(procOld, hWnd, iMsg, wParam, lParam)
'
' End Function
'
'
' This is just a basic overview of subclassing but I hope it helps if
' you were fuzzy about the subject before reading this.
'

















