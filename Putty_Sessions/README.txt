https://fabianlee.org/2018/10/07/putty-bulk-import-putty-session-definitions-into-the-registry-using-powershell/

----------------------------------------------------------------------------------------------------------------------
Fabian Lee Script
----------------------------------------------------------------------------------------------------------------------

1)	Create list of hosts
	The first thing we need is a list of the hosts that we want added to Putty. 
	Create a text file named “puttyhosts.txt”. 
	The format we will use is a text file where each line represents a host, and each line has the format:
	<name>,<IP|hostname>
	Here is an example: myhost,192.168.1.10

2)	Populate PuTTy session list
	Download createPuttySessions.ps1 and template.reg from my github project and put it into the same directory as the “puttyhosts.txt” you created in the last step.

3)	cmd -> cd..\Putty_Sessions -> run powershell -executionpolicy bypass .\createPuttySessions.ps1
	This will go through each line in your “puttyhosts.txt” and use the “template.reg” as a blueprint for creating registry files that are then imported using the standard Windows utility reg.exe.  
	Close PuTTy and reopen it, and now you should see all your host sessions listed.




