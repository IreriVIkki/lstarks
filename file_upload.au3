WinActivate("File Upload")
WinWaitActive("File Upload")
ControlSend("File Upload", "", "Edit1", $CmdLine[1]);
ControlClick("File Upload","","Button1")