
set x = WScript.CreateObject("WScript.Shell")
mySecret = inputbox("Enter text to be Encrypted")
mySecret = StrReverse(mySecret)
msgbox "Please Wait While Text is Encrypted",64,"Encryption"
x.Run "notepad.exe"
wscript.sleep 5000
x.sendkeys encode(mySecret)
function encode(s)
For i = 1 To Len(s)
newtxt = Mid(s, i, 1)
newtxt = Chr(Asc(newtxt)+3)
coded = coded & newtxt
Next
encode = coded
msgbox "Compelted..."
End Function