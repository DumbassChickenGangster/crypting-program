set x = WScript.CreateObject("WScript.Shell")
mySecret = inputbox("enter text to be encoded")
mySecret = StrReverse(mySecret)
x.Run "%windir%\notepad"
wscript.sleep 5000
x.sendkeys encode(mySecret)'this function encodes the text by advancing each character 3 letters'
function encode(s)
For i = 1 To Len(s)
newtxt = Mid(s, i, 1)
newtxt = Chr(Asc(newtxt)+2)
coded = coded & newtxt
Next
encode = coded
End Function