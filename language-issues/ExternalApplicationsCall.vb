Shell("file:///C|/Andy/My%20Documents/oo/tmp/h.bat",2) 'URL notation uses /
Shell("C:\Andy\My%20Documents\oo\tmp\h.bat",2) 'Windows notation uses \
Shell("/home/andy/foo.ksh", 10, """one argument"" another") ' two arguments


'0 Focus is on the hidden program window.
'1 Focus is on the program window in standard size.
'2 Focus is on the minimized program window.
'3 Focus is on the maximized program window.
'4 Standard size program window, without focus.
'6 Minimized program window, but focus remains on the active window.
'10 Full-screen display.

