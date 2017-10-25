#$Language="VBScript"
#$Interface="1.0"

' The purpose of this file is to serve as a pre-install configuration script.  This script will follow
' internal processes for network device delivery.  This script can be used as a stand-alone or by supplying arguments
' via command line.  By default arguments will be required for this script to work, or edit the Config Output Variables listed below.
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.

'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <http://www.gnu.org/licenses/>.




Dim g_ConfigFile, g_DestinationDirectory, g_DestinationFile, g_DeviceType, g_Prompt, g_PromptExtension, g_ConfigFileName, g_LogToPrompt, g_UpdateIOSFile
g_DestinationDirectory = "\\server\share\directory\"
g_ConfigFile = "z:\UserFiles\tcurtis\Output Files\CORP_Switch_L2_L3_Data\Sample Site 2 (L2)\corp-3560cx-L2-sw1.txt"


Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.GetFile(g_ConfigFile)

' Config Output Variables
g_DeviceType = "SWITCH" 												' Set to specify device type (Router|Switch|ASA)
g_DeviceHostname = "corp-c3560cx-L2-sw1"								' Set to specify device hostname (final hostname)
g_DeviceIOS	= "c3560cx-universalk9-tar.152-4.E.tar"						' Set to specify IOS image file
g_DeviceIOSInstall = "Archive"											' Set to specify IOS image deployment type (Standard|Archive|Software Install)
g_DeviceProvisioningIF = "VLAN1"										' Set to specify Provisioning interface (Gig0/1, Gig1/0/1/, Vlan1)
g_DeviceFlash = "flash"													' Set to specify flash device name (e.g. flash, bootdisk, disk0, etc) (will set to flash if not specified)
g_SCP_Host = "10.1.1.1"													' Set to SCP Server hostname/IP for file transfer of IOS/config file
g_SCP_Username = "provisioning"											' SCP Username
g_SCP_Password = "mypass123"											' SCP Password


' ** Log File Configuration
g_Log_Enabled = True													' Enable/Disable Log File
g_Log_DateStamp = True													' Prepend Date Stamp to each line in log file
g_Log_DateStampFileName = FALSE											' Add Date Stamp to log file name (unique file names)
g_Log_FileLocation = "C:\_Logs\"										' Log File directory (include trailing \)
g_Log_FileName = crt.window.caption & "-logtofiletest2.txt"				' Log File Name
g_Log_OverwriteOrAppend = "append"										' Append or Overwrite



Sub Main
	crt.Screen.Synchronous = True
	LogToFile ("---------------------------------------- BEGIN LOG  -  " & g_DeviceHostname & "  --------------------------------------------")
	If Not crt.session.connected then
        LogToFile (g_DeviceHostname & " - FAILED - Not Connected.  Please connect before running this script.")
		LogToFile (g_DeviceHostname & " - FAILED - Exiting script.")
		LogToFile (" ================================== END LOG  -  " & g_DeviceHostname & "  ====================================")
		msgBox "FAILED TO CONNECT" & vbcrlf & vbcrlf & g_DeviceHostname & " - Not Completed." & vbcrlf &_
				"Session must be connected to work." & vbcrlf & "Connect and try again."
		exit sub
    end if
 
	
	Get_Arguments
	g_ConfigFileName = objFSO.GetFileName(g_ConfigFile)
	'msgbox g_ConfigFileName
	g_DestinationFile = g_DestinationDirectory & g_ConfigFileName
	Call CopyFile (g_ConfigFile, g_DestinationFile)		
	Get_Prompt 		' Get Command Prompt
	if g_PromptExtension <> "#" Then Call Get_Enable											
	Get_PrepareDevice
	Get_ConfigDevice
	Get_VerifyIOS
	Get_CopySCPFiles
	Get_InstallFiles
	SendCommand (g_LogToPrompt)
	LogToFile ("  -------------- Script Completed ------------- ")
	sLogToPrompt = Replace(g_LogToPrompt,vbcr,vbCrlf)
	LogToFile (vbcrlf & vbcrlf & vbcrlf & sLogToPrompt)	
	LogToFile (" ==================== END LOG   =============================")
End Sub

Sub Get_Prompt
			If crt.session.Connected = "-1" Then
               Do
                    ' Simulate pressing "Enter" so the prompt appears again...
                    crt.Screen.Send vbcr & "term len 0" & vbCr
                    ' Attempt to detect the command prompt heuristically by
                    ' waiting for the cursor to stop moving... (the timeout for
                    ' WaitForCursor above might not be enough for slower-
                    ' responding hosts, so you will need to adjust the timeout
                    ' value above to meet your system's specific timing
                    ' requirements).
                    Do
                        bCursorMoved = crt.Screen.WaitForCursor(1)
                    Loop Until bCursorMoved = False
                    ' Once the cursor has stopped moving for about a second,
                    ' we'll assume it's safe to start interacting with the
                    ' remote system. Get the shell prompt so that we can know
                    ' what to look for when determining if the command is
                    ' completed. Won't work if the prompt is dynamic (e.g.,
                    ' changes according to current working folder, etc.)
                    nRow = crt.Screen.CurrentRow
                    g_Prompt = crt.screen.Get(nRow, _
                                               0, _
                                               nRow, _
                                               crt.Screen.CurrentColumn - 1)
                    ' Loop until we actually see a line of text appear:
                    g_Prompt = Trim(g_Prompt)
					
					set rePrompt = New RegExp
					set rePromptExtension = New RegExp
					rePrompt.pattern = "(.*)[\>|#]"
					rePromptExtension.pattern = ".*([\>|#])"
					
					Set matches = rePrompt.Execute(g_Prompt)
						For each match in matches
							g_hostname = match.SubMatches(0)
						Next
					
					Set matches = rePromptExtension.Execute(g_Prompt)
						For each match in matches
							g_PromptExtension = match.SubMatches(0)
						Next	
                    If g_Prompt <> "" Then Exit Do
					
					
                Loop
			End If
			
			SendCommand ("! From GET_PROMPT Subroutine ")
			SendCommand ("! g_Prompt is : " & g_Prompt)
			SendCommand ("! g_PromptExtension : " & g_PromptExtension)
							
			
	
 End Sub
 
Sub Get_Enable
 ' ** Execute this sub if device is not in enable mode.  The program is designed to 
 ' ** only execute generic password of cisco or error out. The device should not have a running config to work properly.
 '
 
	crt.screen.send "enable" & vbCrlf
	crt.sleep 500
	nRow = crt.Screen.CurrentRow
	sPrompt = crt.screen.Get(nRow, _
							   0, _
							   nRow, _
							   crt.Screen.CurrentColumn - 1)


	if LCase(left(sPrompt,4)) = "pass" Then 
			'msgBox "Password script"
			crt.screen.send "cisco" & vbCrlf
			crt.sleep 500
			crt.screen.send vbCrlf
			nRow = crt.Screen.CurrentRow
					sPrompt = crt.screen.Get(nRow, _
											   0, _
											   nRow, _
											   crt.Screen.CurrentColumn - 1)	
					
					set rePromptExtension = New RegExp
					rePromptExtension.pattern = ".*([\>|#])"


					Set matches = rePromptExtension.Execute(sPrompt)
						For each match in matches
							g_PromptExtension = match.SubMatches(0)
						Next	
				if g_PromptExtension <> "#" Then LogToFile  (g_DeviceHostname & " - PREPARATION - Error with enable password " & g_PromptExtension)
		End If					

 End Sub
 

Sub Get_Arguments()
	If crt.Arguments.Count = 1 Then
	
		If left(crt.Arguments(0),2) = "/help" Then
			' Display help menu
			Dim sErrorMessage
			sErrorMessage = "Install Pre-install.vbs" & vbcrlf & vbcrlf &_
			"This script requires arguments to function properly. " & vbcrlf &_
			" At a minimum an input (-i:) and config file (-c:) have to be specified." & vbcrlf & vbcrlf &_
			"These items should be entered into individual strings using /arg to deliniate each one." & vbcrlf &_
			"e.g. SecureCRT /script myscript.vbs /arg -i:install.vbs /arg -c:config.txt" & vbcrlf & vbcrlf &_
			"Available Switches:" & vbcrlf &_
			" -c:" & vbTab & "Specify the configuration file to be used with input file. (Full path required)." &vbcrlf & vbTab & "(Default to My Documents if path not specified)" & vbcrlf &_
			" -i:" & vbTab & "IOS Filename (Needs to be on SCP Server)" & vbcrlf &_
			" -h:" & vbTab & "Device Hostname" & vbcrlf &_
			" -t:" & vbTab & "Install Type (Standard, Archive, Software Install) " &vbcrlf & vbTab & "(COM1,COM2,etc)"
			MsgBox sErrorMessage
		 End If
	End if
		 
	If crt.Arguments.Count > 0 Then
		' Split the variable into an array using semicolon as delimeter
		Dim sArgs
		sArgs = crt.Arguments.Count
		'msgBox "sArgs = " & sArgs
		Dim oArg()
		sArgs = sArgs - 1
		ReDim oArg (sArgs)
		'msgBox " sArgs is set to : " & sArgs & vbcrlf & crt.Arguments (sArgs)
		Do While sArgs => 0
			oArg (sArgs) = crt.Arguments(sArgs)
		'	msgBox "Argument is: " & oArg (sArgs) & vbcrlf & _
		'			" sArgs is set to : " & sArgs
			sArgs = sArgs - 1
			if sArgs = -1 then Exit Do
		Loop
	
	
	
	 For Each i in oArg
		  sSwitch = LCase(Left(i,3))
		  sValue = Right(i,Len(i)-3)
				'msgbox i
			Select Case sSwitch
				Case "-i:"
					g_DeviceIOS = sValue
				Case "-c:"
					g_ConfigFile = sValue
				Case "-t:"
					g_DeviceIOSInstall = sValue
				Case "-h:"
					g_DeviceHostname = sValue
			End Select
		Next


	 'crt.dialog.MessageBox "Connection Type: " & sConnectionType
	 
	 If g_DeviceIOS = "" Then sError = "IOS Missing.  Specify manually or by using -i:<filename>" & vbcrlf
	 If g_DeviceIOSInstall = "" Then sError = sError & "IOS Install Type Missing.  Specify manually or by" & vbcrlf &_
						" using -t:<Standard|Software Install|Archive>" & vbCrlf
	 If g_ConfigFile = "" Then sError = sError & "Filename Missing.  Specify manually or by using -c:<filename>" & vbCrlf
	 If g_DeviceHostname = "" Then sError = sError & "Hostname Missing.  Specify manually or by using -h:<hostname>" & vbCrlf
		if sError <> "" Then msgbox sError
		
	 
	 
	 	 
	End If 
	 
 End Sub


Sub Get_VerifyIOS
	
	crt.Screen.Synchronous = True	
	
	crt.screen.send "dir | include " & left(g_DeviceIOS,8) & ".*" & right (g_DeviceIOS,4) & vbCr
	nResult = crt.Screen.ReadString (g_Prompt,30)
	'msgbox "Dir output is: " & nresult & vbcrlf & "It should be: " & g_DeviceIOS
 	
	'msgbox nresult & vbcrlf & g_DeviceIOS
 
	set reIOSPattern = New RegExp
	reIOSPattern.pattern = ".*("& g_DeviceIOS & ")"
		
	Set matches = reIOSPattern.Execute(nResult)
		For each match in matches
			sIOSPattern = match.SubMatches(0)
		Next	
	
	if sIOSPattern = g_DeviceIOS Then sUpdateIOSFile = "no"
	
	
 End Sub
 


Sub Get_PrepareDevice
 
 
 	crt.Screen.Synchronous = True
	
	' Delete VLAN.dat if device is a switch
	If lcase(g_DeviceType) = "switch" Then
		crt.screen.send "delete flash:vlan.dat" & vbCr
		nResult = crt.screen.WaitForStrings ("Delete fil", g_Prompt, 1)
		If nResult = 1  Then SendCommand (vbcrlf)
		nResult = crt.screen.WaitForStrings ("[confirm]", g_Prompt, 1)
		If nresult = 1 Then crt.screen.send "y"
		crt.screen.WaitForString g_Prompt, 3
		LogToFile (g_DeviceHostname & " - PREPARATION - VLAN.dat deleted - Success")
		LogToPrompt ("VLAN.dat Deleted:                SUCCESS")
	End If

	crt.screen.send "erase startup-config" & vbCr
	nResult = crt.screen.WaitForStrings ("[confirm]" , 20)
	'msgbox ("Erase Config found confirm prompt: " & nResult)
	if nresult = 1 Then
		crt.screen.send "y"
		crt.screen.WaitforString "complete", 20
		LogToFile (g_DeviceHostname & " - PREPARATION - NVRAM Erase - Success")
		LogToPrompt ("NVRAM Deleted:                    SUCCESS")
	End If
	
	crt.screen.send vbCr
	crt.screen.WaitForString vbCr
	
	' ** Reload the device once flash has been cleared.
	crt.screen.send ("reload") & vbCr
	crt.screen.WaitForString vbcr, 1
	
	
	nResult = crt.screen.WaitForStrings ("[yes/no]","[confirm]",2)
	
	' ** Don't save the configuration if it was modified
	If nResult = 1 Then 
		crt.screen.send "no" & vbCr
	End if
	
	LogToFile (g_DeviceHostname & " - PREPARATION - Reload - Did not save running config - Success")
	
	' ** Confirm reboot
	crt.screen.send "y" & vbCrlf
		
	Get_AnswerRebootPrompts
		
	
	SendCommand ("! Router has now rebooted.  Beginning setup.")
	LogToFile (g_DeviceHostname & " - PREPARATION - Router has now rebooted.  Beginning setup.")
	LogToPrompt ("Device Reload:                    SUCCESS")
	
 End Sub 
 
Sub Get_AnswerRebootPrompts ()
 
	nResult = crt.screen.WaitForStrings ("Press RETURN to get started!", 600)
	crt.sleep 10000			' Sleep 10 seconds
	
	' ** Send a few carriage returns after successful reboot.
	crt.screen.send vbcr
	crt.screen.WaitForString vbcr, 10
	crt.screen.send vbcr
	crt.screen.WaitForString vbcr, 1
	crt.screen.send vbcr
	crt.screen.WaitForString vbcr, 1
	crt.sleep 3000
	crt.screen.send vbcr
	crt.screen.WaitForString vbcr, 1
	crt.screen.send vbcr
	crt.screen.WaitForString vbcr, 1
	crt.screen.send vbcr
	crt.screen.WaitForString vbcr, 1
	crt.screen.send vbcr
	nResult = crt.screen.WaitForStrings ("yes/no", "'yes' or 'no'", "[Router|Switch]>", 1000)
	'msgBox "nresult = " & nresult
	if nresult = 2 Then nresult = 1
	' ** If result is yes/no, then send the no string to terminal
	if nResult = 1 Then
		crt.screen.send "no" & vbCr
		iResult = crt.screen.WaitForStrings (">","[yes]", 5)
		if iResult = 2 then 
			crt.screen.send "yes" & vbCr
			crt.screen.WaitForStrings "Press Return to get started!",">",10
			SendCommand (vbcr)
			SendCommand (vbcr)
			
		End If
	End if
	
	
	SendCommand ("enable")

	SendCommand ("! Router has now rebooted.  Beginning setup.")
	LogToFile (g_DeviceHostname & " - PREPARATION - Router has now rebooted.  Beginning setup.")
	LogToPrompt ("Device Reload:                    SUCCESS")

	
 End Sub	
 
   
Sub Get_ConfigDevice ()
 
  
 
    'crt.Screen.IgnoreEscape = True
	
	SendCommand ("term width 120")
	SendCommand ("config t")
	SendCommand ("no ip domain-lookup")
	SendCommand ("line con 0")
	SendCommand ("exec-timeout 360")
	SendCommand ("no logging synchronous")
	SendCommand ("logging console")
	SendCommand ("ip domain-name duke-energy.com")
	SendCommand ("hostname " & g_DeviceHostname)
	g_Prompt = g_DeviceHostname
	SendCommand ("interface " & g_DeviceProvisioningIF)
	SendCommand ("ip address dhcp")
	SendCommand ("no shut")
	
	'crt.sleep 3000
	'SendCommand (vbcr)
	nResult = crt.screen.ReadString ("hostname " & g_DeviceHostname, 90)
	'msgbox nresult	
	
	If nResult <> "" Then
		'SendCommand (vbcr)
		'SendCommand ("!DHCP Address received")
	set reAddress = New RegExp
	reAddress.pattern = ".*DHCP address (([0-9]{1,3}\.){3}[0-9]{1,3})"
	
	set reMask = New RegExp
	reMask.pattern = ".*, mask (([0-9]{1,3}\.){3}[0-9]{1,3})"
	
	Set matches = reAddress.Execute(nResult)
			For each match in matches
				sAddress = match.SubMatches(0)
			Next	

	Set matches = reMask.Execute(nResult)
			For each match in matches
				sMask = match.SubMatches(0)
			Next
			
		LogToFile (g_DeviceHostname & " - PREPARATION -  DHCP Address received - " & sAddress & " mask " & sMask)
		LogToPrompt ("DHCP Address: " & "                     SUCCESS")
	Else
		'SendCommand (vbcr)
		'SendCommand ("! Timeout waiting for DHCP Address.  Correct cable issue and try again.")
		LogToFile (g_DeviceHostname & " - FAILED - ! Timeout waiting for DHCP Address.  Correct cable issue and try again.")
		LogToPrompt ("DHCP Address: TIMEOUT" & "         FAILED")
	End If
	
	SendCommand ("end")
	
 End Sub


Sub Get_CopySCPFiles ()
 
	crt.Screen.Synchronous = True
	
	'crt.screen.send vbcrlf & "!g_Prompt : " & g_Prompt & vbcrlf &_
	'		"!g_PromptExtension : " & g_PromptExtension & vbCrlf
	'msgbox "copy scp://" & g_SCP_Username & ":" & g_SCP_Password & "@" g_SCP_Host & "/" & objFSO.GetFileName(objFile) & " flash:"
	
	crt.screen.send vbCr
	crt.screen.WaitforString g_PromptExtension
	crt.screen.send vbCr
	crt.screen.WaitforString g_PromptExtension
	
	SendCommand ("!Starting Get_CopySCPFiles procedure")
	'msgbox "Config filename is : " & g_ConfigFileName
	LogToFile (g_DeviceHostname & " - CONFIG FILE COPY - Begin Config File (" & g_ConfigFileName & ") Copy to " & g_DeviceFlash & ":")
	crt.screen.send "copy scp://" & g_SCP_Username & ":" & g_SCP_Password & "@" & g_SCP_Host & "/" &_
								g_ConfigFileName & " " & g_DeviceFlash & ":" & vbCr
	nResult = crt.screen.WaitForStrings ("Destination filename",10)
	if nResult = 1 Then crt.screen.send vbCrlf
	nresult = crt.screen.WaitForStrings ("[confirm]", "!",3)
	'crt.screen.ReadString ("[confirm]", g_Prompt & g_PromptExtension,3)
	if nResult = 1 Then	crt.screen.send vbCrlf
	
	nResult = crt.screen.WaitForStrings ("bytes copied in", "Error", 90)
	crt.screen.WaitForString g_Prompt, 1
	'nerror = crt.screen.ReadString (g_Prompt & g_PromptExtension, 3)
	if nresult = 1 Then
	
		LogToFile (g_DeviceHostname & " - CONFIG FILE COPY - Config File: " & g_ConfigFileName & " copied successfully to " & g_DeviceFlash & ":")
		LogToPrompt ("Copy Config to Flash:" & "              SUCCESS")
		'msgbox "File copied successfully"
	Elseif nResult = 2 Then
		LogToFile (g_DeviceHostname & " - CONFIG FILE COPY - Config File: " & g_ConfigFileName & " FAILED copying file to " & g_DeviceFlash & ":")
		LogToPrompt ("Copy Config to Flash:" & "            FAILED")
	End If	

	'SendCommand (vbcr)
	crt.sleep 2000
	SendCommand ("! Sync back up in script")
	
	crt.screen.send vbCr
	crt.screen.WaitforString g_PromptExtension
	crt.screen.send vbCr
	crt.screen.WaitforString g_PromptExtension
	
	If g_UpdateIOSFile = "yes" Then 
		LogToFile (g_DeviceHostname & " - IOS FILE COPY - Begin IOS File (" & g_DeviceIOS & ") Copy to " & g_DeviceFlash & ":")
		crt.screen.send "copy scp://" & g_SCP_Username & ":" & g_SCP_Password & "@" & g_SCP_Host & "/" &_
							g_DeviceIOS & " " & g_DeviceFlash & ":" & vbCrlf & vbcrlf
		'SendCommand ("! IOS Would have copied here.")
		nResult = crt.screen.WaitForStrings ("bytes copied in", "Error", 2500)
		if nResult = 1 Then 
			LogToFile (g_DeviceHostname & " - IOS FILE COPY - IOS Copied successfully.")
			LogToPrompt ("IOS File Copy:                  SUCCESS")
		End If
	Else
		SendCommand ("! IOS FIle " & g_DeviceIOS & " already exists.  Skipping this step.")
		LogToPrompt ("IOS File Copy:                 N/A (Already exists)")
		LogToFile (g_DeviceHostname & " - IOS FILE COPY - File:" & g_DeviceIOS & " - Already exists.")
	End If
	
 End Sub

Sub Get_InstallFiles ()
 
	SendCommand (vbcr)
	SendCommand (vbcr)
	SendCommand ("!Starting Get_InstallFiles procedure")
	
	If lcase(g_DeviceIOSInstall) = "archive" Then
		LogToFile (g_DeviceHostname & " - SOFTWARE INSTALL - Archive Download-sw process started.")
		SendCommand ("! Starting the ARCHIVE Process.  ")
		SendCommand ("! This could take up to 20 minutes.")
		SendCommand ("! Do not power off or disconnect from ")
		SendCommand ("!    the network during this period.")
		
		crt.screen.send "archive download-sw /overwrite " & g_DeviceFlash & ":/" & g_DeviceIOS & vbCr
		nResult = crt.screen.WaitForStrings ("All software images installed.", 1200)
		
		if nResult = 1 Then
			LogToFile (g_DeviceHostname & " - SOFTWARE INSTALL - Archive Download-sw process completed successfully.")
			LogToPrompt ("Software Install: Archive         SUCCESS")
		Else
			LogToFile (g_DeviceHostname & " - SOFTWARE INSTALL - Archive process FAILED.")
			LogToPrompt ("Software Install: Archive          FAILED")
		End If
		
		
		'crt.screen.WaitForString vbCr, 1
		'SendCommand ("! would have archived software here.")
	ElseIf lcase(g_DeviceIOSInstall) = "software_install" Then
		crt.screen.send "software install file " & g_DeviceFlash & ":" & g_DeviceIOS & " force" & vbCr
		nResult = crt.screen.WaitForStrings ("operation aborted", "Do you want to proceed with reload? [yes/no]", 240)
		if nresult = 1 Then
			LogToFile (g_DeviceHostname & " - SOFTWARE INSTALL - Install Processess Aborted - FAILED")
			LogToPrompt ("Software Install:	Install         	FAILED")
		elseif nResult = 2 Then
			crt.screen.send "yes" & vbcr
			iResult = crt.screen.WaitForStrings (g_Prompt, "[yes/no]", 5)
			if iResult = "2" Then SendCommand ("no")
			LogToFile (g_DeviceHostname & " - SOFTWARE INSTALL - Install Processess Successfully installed " &_
						g_DeviceFlash & ":" & g_DeviceIOS & " - SUCCESS")
			LogToPrompt ("Software Install: Install           SUCCESS")
		End If
		' ** Call sub to wait for reboot and answer reboot prompts.
		Get_AnswerRebootPrompts
	ElseIf lcase(g_DeviceIOSInstall) = "standard" Then
		LogToFile (g_DeviceHostname & " - SOFTWARE INSTALL - Standard IOS Image (.bin).  No action necessary.")
		LogToPrompt ("Software Install: Standard             N/A")
	End If
	
	
	
 
	' ** Copy config file to running-config
	crt.screen.send "copy " & g_DeviceFlash & ":" & g_ConfigFileName & " running-config" & vbCr
	crt.screen.WaitForString "Destination filename",1
	crt.screen.send vbCrlf
	nresult = crt.screen.WaitForStrings ("[confirm]", "bytes/sec","No such file or directory", 3)
	if nResult = 1 Then	crt.screen.send "y"
	' ** If File is not found, write error to log.
	if nResult = 3 Then LogToFile (g_DeviceHostname & " - CONFIG INSTALL - Configuration FAILED to running-config. - FAILED")
	'msgbox g_Prompt
	sResult = crt.screen.ReadString (g_Prompt, 90)
	if nResult <> 3 Then LogToFile (g_DeviceHostname & " - CONFIG INSTALL - Configuration copied to running-config.")
	
	'msgBox sresult
	'msgBox nresult
	sResult = split(sResult,vbcrlf)
	
	
	' ** Add date/time stamp to each line in the output from copying config to running
	sresult3 = vbcrlf
		
	For each i in sResult
		sResult3 = sResult3 & Now & "   " & g_DeviceHostname & " - CONFIG INSTALL - " & i & vbCrlf
	next
	
	LogToFile (sresult3)
	
	
	
	'SendCommand ("! Config file would have copied into running")
	crt.screen.send "write mem" & vbCr
	' Verify OK Prompt is received
	nresult = crt.screen.WaitForStrings ("[OK]",20)
	If nResult = 1 Then
		' Write to log confirming config copied succesfully
		LogToFile (g_DeviceHostname & " - CONFIG INSTALL - Configuration written to memory successfully.")
	Else
		LogToFile (g_DeviceHostname & " - CONFIG INSTALL - FAILED - Copy error writing to NVRAM.")
	End If
	
	SendCommand (vbcr)
	
 End Sub


Sub CopyFile(SourceFile, DestinationFile)
 
    Set fso = CreateObject("Scripting.FileSystemObject")

    'Check to see if the file already exists in the destination folder
    Dim wasReadOnly
    wasReadOnly = False
    If fso.FileExists(DestinationFile) Then
        'Check to see if the file is read-only
        If fso.GetFile(DestinationFile).Attributes And 1 Then 
            'The file exists and is read-only.
            'MsgBox "Removing the read-only attribute"
            'Remove the read-only attribute
            fso.GetFile(DestinationFile).Attributes = fso.GetFile(DestinationFile).Attributes - 1
            wasReadOnly = True
        End If
    End If

    'Copy the file
    'msgBox "Copying " & SourceFile & " to " & DestinationFile
    fso.CopyFile SourceFile, DestinationFile, True
	'If fso.FileExists(DestinationFile) Then msgBox "File exists now"
    If wasReadOnly Then
        'Reapply the read-only attribute
        fso.GetFile(DestinationFile).Attributes = fso.GetFile(DestinationFile).Attributes + 1
    End If

    Set fso = Nothing

 End Sub


Function SendCommand (strComment)
 
    crt.screen.Send(strComment & vbcr)
    crt.screen.WaitForString strComment & vbcr, 3
  	crt.screen.WaitForString "#", 1
	
End Function

Sub LogToPrompt(strMessage)
	if g_LogToPrompt = "" Then g_LogToPrompt = "!  Configuration of Device " & g_DeviceHostname & " completed. Status listed below." & vbcr &_
					"!  ----------------------------------------------------------------------" & vbcr & vbcr & vbcr
	g_LogToPrompt = g_LogToPrompt & "! " & strMessage & vbCr

 End Sub

 
Sub LogToFile(strMessage)
 
   If g_Log_Enabled = False Then Exit Sub
 
    Const ForReading = 1
    Const ForWriting = 2
    Const ForAppending = 8
 
    Set oLogFSO = CreateObject("Scripting.FileSystemObject")
  
    If g_Log_FileLocation = "relative" Then
        Set oLogShell = CreateObject("Wscript.Shell")
        g_Log_FileLocation = oLogShell.CurrentDirectory & "\"
        Set oLogShell = Nothing
    End If
   
    If g_Log_DateStampFileName Then
        sNow = Replace(Replace(Now(),"/","-"),":",".")
        sLog_FileName = sNow & " - " & g_Log_FileName
        g_Log_DateStampFileName		= False      
	Else
		sLogFileName = g_Log_FileName
    End If
 
	
    sLog_File = g_Log_FileLocation & sLog_FileName
   
    If g_Log_OverWriteOrAppend = "overwrite" Then
        Set oLogFile = oLogFSO.OpenTextFile(sLog_File, ForWriting, True)
        g_Log_OverWriteOrAppend = "append"
    Else
        Set oLogFile = oLogFSO.OpenTextFile(sLog_File, ForAppending, True)
    End If
 
    If g_Log_DateStamp Then
        strMessage = Now & "   " & strMessage
    End If
 
    oLogFile.WriteLine(strMessage)
    oLogFile.Close
    
    oLogFSO = Null

 End Sub
 