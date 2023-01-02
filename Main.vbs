Dim Ws,times,Str 

Set Ws = WScript.CreateObject("WScript.Shell")

WScript.Sleep 2000

times = Inputbox ("How many times would you like to send? " , "Discord Sender" , 20)

Str = InputBox ("ÄÚÈÝ" , "Discord Sender" , "")

For i = 1 To times

Ws.SendKeys "{ENTER}"

Ws.SendKeys "^v"

Ws.SendKeys "{ENTER}"

WScript.Sleep 50

Ws.SendKeys "{ENTER}"

WScript.Sleep 50

Ws.SendKeys "{ENTER}"

WScript.Sleep 50

Ws.SendKeys "{ENTER}"

WScript.Sleep 600

Next

MsgBox "Done! " , vbInformation , "Discord Sender"