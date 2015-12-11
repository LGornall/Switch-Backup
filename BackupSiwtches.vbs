Option Explicit

Dim objShell, objExec, ObjFSO, objTxtFile, strCurrLine, strLastLine, strCommandText, objSaveFile, intTimeOutCount, objMessage, strEmailTxt
Dim arrSwitches(2,1), x, strCurrSwitchName, strCurrSwitchAddress, strSwitchBackupPath, boolConnected, strSwitchOutputFile
dim intTryCount, strSwitchBackupFile

strSwitchOutputFile = "C:\temp\switch.txt"


'Creating Objects and Arrays 
set objShell = CreateObject("WScript.Shell")
set objExec = objShell.exec("%comspec%")
set objFSO = CreateObject("Scripting.FileSystemObject")
Set objMessage = CreateObject("CDO.Message") 

arrSwitches(0,0) = "Switch Name 1"
arrSwitches(0,1) = "Switch IP 1" 
arrSwitches(1,0) = "Switch Name 2"
arrSwitches(1,1) = "Switch IP 2"
arrSwitches(2,0) = "Switch Name 3"
arrSwitches(2,1) = "Switch IP 3"

'sendEmail false, "test", strCurrSwitchName

For x = 0 to 2
	strCurrSwitchName = arrSwitches(x,0)
	strCurrSwitchAddress = arrSwitches(x,1)
	
	strSwitchBackupPath = "backups location \" & strCurrSwitchName & "\" & year(date) & "\" & monthName(month(date),True)
	strSwitchBackupFile = strSwitchBackupPath & "\Lucf " & replace(date,"/", "-") & ".txt"
	
	
	intTryCount = 0
	connectToSwitch(strCurrSwitchAddress)
	'check the output from the lucf file has finished 
	intTimeOutCount = 0
	if boolConnected = True Then
		do 
			if readyForInput = True then
				saveCommands
				closeConnection
				sendEmail True, "", strCurrSwitchName
				exit do 
			elseif intTimeOutCount > 4999 then
				sendEmail False, "Script timed out when running RUCF command", strCurrSwitchName
				exit do 
			else 
				wscript.sleep 500
			end if 
		loop while readyForInput = False or intTimeOutCount < 4999
	end if 
Next 

'Function to connect to a Switch 
sub connectToSwitch(argSwitchIP)
	intTryCount = intTryCount + 1
	objExec.StdIn.Write "path to plink.exe" & argSwitchIP & " -telnet > " & strSwitchOutputFile & vbCrlf
	wscript.sleep 200	
	set objTxtFile = objFSO.openTextFile(strSwitchOutputFile)
	intTimeOutCount = 0
	do 
		if readyForInput = True Then
			loginAndRunCommands
			boolConnected = True
			exit do	
		elseif intTryCount > 4 then 'has tied to connect 5 times... give up
			sendEmail False, "Script timed out when connecting to Switch", strCurrSwitchName
			exit do
		elseif intTimeOutCount > 2999 then
			objShell.run("taskkill /im plink.exe /f")
			wscript.sleep (500)
			connectToSwitch(strCurrSwitchAddress)
			exit do 
		else 
			wscript.sleep 500
			intTimeOutCount = intTimeOutCount + 500
		end if 
	loop while readyForInput = False or intTimeOutCount < 2999

end sub

'closing the textFile then opens it and rereads the contents to get last line and look for RUCF in lines
sub updateTxtfile
	objTxtFile.close
	strCommandText = ""
	set objTxtFile = objFSO.openTextFile(strSwitchOutputFile)	
	Do Until objTxtFile.AtEndOfStream
		strCurrline = replace(objTxtFile.ReadLine, "    "," ")
		if instr(strCurrline,"RUCF") > 0 then
			strCommandText = strCommandText & strCurrLine & VBCrlf
		end if
		If objTxtFile.atEndOfStream Then
			strLastLine = strCurrLine
		End If
    Loop
end sub

'Function to login with details and run LUCF commands
function loginAndRunCommands
	'msgbox "running comms"
	objExec.StdIn.Write "osl" & VBcr
	objExec.StdIn.Write "Log in ID" & VBcr
	objExec.StdIn.Write "Log in ID" & VBcr
	objExec.StdIn.Write "Log in ID" & VBcr
	objExec.StdIn.Write "lucf" & VBcr
	wscript.sleep 500
end function

'returns true of false depending if he 
function readyForInput
	updateTxtfile
	if strLastLine = "?" then
		readyForInput = True
	else 
		readyForInput = False
	end if 
End function

'saves command text to network drive
sub saveCommands
	set objSaveFile = nothing
	if objFSO.FolderExists(strSwitchBackupPath) Then
		set objSaveFile = objFSO.createTextFile(strSwitchBackupFile)
		objSaveFile.writeline(strCommandText)
	else 
		objFSO.createFolder(strSwitchBackupPath)
		set objSaveFile = objFSO.createTextFile(strSwitchBackupFile)
		objSaveFile.writeline(strCommandText)
	
	end if 
	

end sub

sub closeConnection
	objExec.StdIn.Write "bye" & VBcr
	objShell.run("taskkill /im plink.exe /f")
	boolConnected = False
	strCommandText = ""
	strLastLine = ""
	
end sub

sub sendEmail(argWasSuccessful,argMessage,argSwitchName)
	if argWasSuccessful = True then
		objMessage.Subject = argSwitchName & " - Switches backed up with no errors" 
		strEmailTxt = "SwitchBackup Script has ran has not come across any errors" & vbcrlf & vbcrlf & "Here is the backup path: " & strSwitchBackupPath
	else 
		objMessage.Subject = argSwitchName & " - !UNSUCCESSFUL BACKUP!" 
		strEmailTxt = "Errors have occured," & vbcrlf & vbCrlf & argMessage & vbCrlf & vbcrlf & "Log has been attached."
		objMessage.AddAttachment strSwitchOutputFile
	end if 
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "Exchange Server"
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
	objMessage.Configuration.Fields.Update
	objMessage.from = "Script@myDomain.com" 
	objMessage.To = "Me@myDomain.com" 
	objMessage.TextBody = strEmailTxt 
	objMessage.Send
end sub
