Dim Ws,times,Info 
Set Ws = WScript.CreateObject("WScript.Shell")
'定义变量及对象

WScript.Sleep 2000

times = Inputbox ("How many times would you like to send? " , "Discord Sender" , 20)
Info = InputBox ("What is the message you want to send?" , "Discord Sender" , "")
'收集将要发送的信息与次数

Ws.run "MsHta vbscript:ClipBoardData.SetData("+""""+"text"+""""+","+""""&Info&""""+")(close)" , 0 , true
'复制‘Info’到剪贴板

WScript.Sleep 1000          '停止1000ms，用于切换窗口

For i = 1 To times          '循环直到i=times

Ws.SendKeys "{ENTER}"       '发送上一条尚未发送的消息
Ws.SendKeys "^V"            '粘贴剪贴板最新的内容
Ws.SendKeys "{ENTER}"       '这里和下面的按下回车键都是用于发送消息以及处理速度过快弹窗
WScript.Sleep 50
Ws.SendKeys "{ENTER}"
WScript.Sleep 50
Ws.SendKeys "{ENTER}"
WScript.Sleep 50
Ws.SendKeys "{ENTER}"
WScript.Sleep 600           '等待600ms重新循环发送下一条消息

Next

MsgBox "Command completed successfully! " , vbInformation , "Discord Sender"
'报告程序完成