Dim Ws,times,Info 
Set Ws = WScript.CreateObject("WScript.Shell")
'�������������

WScript.Sleep 2000

times = Inputbox ("How many times would you like to send? " , "Discord Sender" , 20)
Info = InputBox ("What is the message you want to send?" , "Discord Sender" , "")
'�ռ���Ҫ���͵���Ϣ�����

Ws.run "MsHta vbscript:ClipBoardData.SetData("+""""+"text"+""""+","+""""&Info&""""+")(close)" , 0 , true
'���ơ�Info����������

WScript.Sleep 1000          'ֹͣ1000ms�������л�����

For i = 1 To times          'ѭ��ֱ��i=times

Ws.SendKeys "{ENTER}"       '������һ����δ���͵���Ϣ
Ws.SendKeys "^V"            'ճ�����������µ�����
Ws.SendKeys "{ENTER}"       '���������İ��»س����������ڷ�����Ϣ�Լ������ٶȹ��쵯��
WScript.Sleep 50
Ws.SendKeys "{ENTER}"
WScript.Sleep 50
Ws.SendKeys "{ENTER}"
WScript.Sleep 50
Ws.SendKeys "{ENTER}"
WScript.Sleep 600           '�ȴ�600ms����ѭ��������һ����Ϣ

Next

MsgBox "Command completed successfully! " , vbInformation , "Discord Sender"
'����������