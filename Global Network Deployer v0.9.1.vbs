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
Dim g_ConnectionType, g_ConnectionPort, g_ApplicationFailed
g_DestinationDirectory = "\\wcltdnsbnap1\D\DEVSETUP\"
g_ConfigFile = "C:\temp\Carolinas West\Test5\lab-3850-l2-sw4.txt"
 
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.GetFile(g_ConfigFile)
 
objFileName = objFSO.GetFileName(g_ConfigFile)
 
Set g_objTab = crt.GetScriptTab()
g_objTab.Caption = "Deploy.IT - " & objFileName
 
g_objTab.Screen.Synchronous = True
 
' Config Output Variables
g_DeviceType = "L2_Switch"
if g_DeviceType <> "" Then g_DeviceType = lcase(right(g_DeviceType, 6))
g_DeviceHostname = "lab-3850-l2-sw4"
g_DeviceIOS = "cat3k_caa-universalk9.SPA.03.06.06.E.152-2.E6.bin"
g_DeviceIOSInstall = "Software Install"
g_DeviceDeleteCatFlash = "" 																' if device requires 
g_DeviceProvisioningIF = "Vlan1"
g_DeviceFlash = "flash"
g_SCP_Host = "139.46.114.21"
g_SCP_Username = "devsetup"
g_SCP_Password = "S72ctAd_sWe96pe"
g_DeviceHardware = "WS-C3850-12S-E"
if g_DeviceHardware = "WS-C4507" Then g_DeviceDeleteCatFlash = "Yes"			' Delete catflash for all 4500 series switches


g_ConnectionType = "Console"
g_ConnectionPort = "COM1"
 
 
' ** Log File Configuration
g_Log_Enabled = True                                                                 ' Enable/Disable Log File
g_Log_DateStamp = True                                                                 ' Prepend Date Stamp to each line in log file
g_Log_DateStampFileName = FALSE                                                       ' Add Date Stamp to log file name (unique file names)
g_Log_FileLocation = "Relative"
g_Log_FileName = g_ObjTab.caption & "-log.txt"                    				' Log File Name
g_Log_OverwriteOrAppend = "Overwrite"
 
 
 
 
Sub Main
     StartTime = Now()
     g_objTab.Screen.Synchronous = True
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
 
 
    If lcase(g_ConnectionType) = "cisco" Then Get_Enable
	
    If lcase(g_ConnectionType) = "mrv" Then
		Call Get_ConnectConsolePort (g_ConnectionType, g_ConnectionPort, g_DeviceHostname)
	ElseIf lcase(g_ConnectionType) = "cisco" Then 
		Call Get_ConnectConsolePort (g_ConnectionType, g_ConnectionPort, "1.1.1.1")
	ElseIf lcase(g_ConnectionType) = "perle" Then 
		'Placeholder for future case
	ElseIf lcase(g_ConnectionType) = "serial" Then 
          'Call Connect("", g_ConnectionPort, "SERIAL","","")
    End If
 
 
	if lcase(g_ApplicationFailed) <> "yes" Then Get_AnswerBootPrompts
    if lcase(g_ApplicationFailed) <> "yes" Then Get_Prompt           ' Get Command Prompt
    if g_PromptExtension <> "#" Then Call Get_Enable                                                       
    if lcase(g_ApplicationFailed) <> "yes" Then Get_PrepareDevice
    if lcase(g_ApplicationFailed) <> "yes" Then Get_ConfigDevice
    if lcase(g_ApplicationFailed) <> "yes" Then Get_VerifyIOS
    if lcase(g_ApplicationFailed) <> "yes" Then Get_CopySCPFiles
    if lcase(g_ApplicationFailed) <> "yes" Then Get_InstallIOSFiles
	if lcase(g_ApplicationFailed) <> "yes" Then Get_InstallConfigFiles
    EndTime = Now()
    dTime = DateDiff("s", StartTime, EndTime)
    dTime = ConvertTime (dTime)
	if lcase(g_ApplicationFailed) = "yes" Then 
		LogToPrompt (" ")
		LogToPrompt (" ")
		LogToPrompt (" ")
		LogToPrompt ("Deployment Failed.  Please correct last failure issue and try again.")
		LogToFile (g_DeviceHostname & " - APP FAILURE - Deployment FAILED.  Correct last failure issue and try again.")
	End If
	
    LogToPrompt ("Total time:             " & dTime)
    SendCommand (g_LogToPrompt)
    LogToFile ("  -------------- Script Completed ------------- ")
    sLogToPrompt = Replace(g_LogToPrompt,vbcr,vbCrlf)
    LogToFile (vbcrlf & vbcrlf & vbcrlf & sLogToPrompt)     
    LogToFile("Total run time: " & dTime)     
    LogToFile (" ==================== END LOG   =============================")
   
 End Sub


Sub Get_Prompt
	
	' Heuristically determine the shell's prompt.  Crt.Screen.Synchronous must
	' already have been set to True.  In general, Crt.Screen.Synchronous should
	' be set to True immediately after a successful crt.Session.Connect().  In
	' This script, SecureCRT should already be connected -- otherwise, a script
	' error will occur.
    Do													' Simulate pressing "Enter" so the prompt appears again...
        g_objtab.Screen.Send vbcr & vbcr & vbcr
						' Attempt to detect the command prompt heuristically by waiting for the
						' cursor to stop moving... (the timeout for WaitForCursor above might
						' not be enough for slower- responding hosts, so you will need to adjust
						' the timeout value above to meet your system's specific timing
						' requirements).
        Do
            bCursorMoved = g_objtab.Screen.WaitForCursor(1)
        Loop Until bCursorMoved = False
																				' Once the cursor has stopped moving for about a second, it's assumed
																				' it's safe to start interacting with the remote system. Get the shell
																				' prompt so that it's known what to look for when determining if the
																				' command is completed. Won't work if the prompt is dynamic (e.g.,
																				' changes according to current working folder, etc.)
        nRow = g_objtab.Screen.CurrentRow
        g_Prompt = g_objtab.Screen.Get(nRow, 0, nRow, crt.Screen.CurrentColumn - 1)
																				' Loop until a line of non-whitespace text actually appears:
        g_Prompt = Trim(g_Prompt)
        If g_Prompt <> "" Then Exit Do
    Loop
    
	
	if right(g_Prompt,1) = "#" or right(g_Prompt,1) = ">" Then
	
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
	Else
		Get_AnswerBootPrompts
	End If	
	SendCommand ("! Prompt is now set to " & g_Prompt)
	
	
 End Sub
 
 
Sub Get_PromptOLD
	   If g_objtab.session.Connected = "-1" Then
	   Do
			' Simulate pressing "Enter" so the prompt appears again...
			crt.Screen.Send "! Test" & vbCr
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
			
			Get_CheckPrompt
			
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
 
 
 
Sub Get_CheckPrompt
	Dim sResult
	g_objTab.screen.send vbcr
	sResult = g_objTab.screen.ReadString (vbcr)
	
	
	
 End Sub	
 
Sub Get_Enable
 ' ** Execute this sub if device is not in enable mode.  The program is designed to 
 ' ** only execute generic password of cisco or error out. The device should not have a running config to work properly.
 '
 
     g_objtab.screen.send "enable" & vbCrlf
     crt.sleep 500
     nRow = g_objtab.Screen.CurrentRow
     sPrompt = g_objtab.screen.Get(nRow, _
                                      0, _
                                      nRow, _
                                      g_objtab.Screen.CurrentColumn - 1)
 
 
     if LCase(left(sPrompt,4)) = "pass" Then 
               'msgBox "Password script"
               g_objtab.screen.send "cisco" & vbCr
               g_objtab.screen.WaitForString vbcr, 10
               nRow = g_objtab.Screen.CurrentRow
                         sPrompt = g_objtab.screen.Get(nRow, _
                                                          0, _
                                                          nRow, _
                                                          g_objtab.Screen.CurrentColumn - 1)     
 
                         set rePromptExtension = New RegExp
                         rePromptExtension.pattern = ".*([\>|#])"
 
 
                         Set matches = rePromptExtension.Execute(sPrompt)
                              For each match in matches
                                   g_PromptExtension = match.SubMatches(0)
                              Next     
                    if g_PromptExtension <> "#" Then LogToFile  (g_DeviceHostname & " - PREPARATION - Error with enable password " & g_PromptExtension)
          End If                         
	
	Get_Prompt
	
 End Sub
 
 
 
 Function Get_ConnectConsolePort (strConnectionType, strPortNumber, strHostname)
           g_objtab.Screen.Synchronous = True
         Select Case UCase(strConnectionType)
        Case "MRV"
               g_objtab.Screen.Send vbCr
               g_objtab.screen.WaitForString ">", 5
               g_objtab.Screen.Send vbCr
               g_objtab.screen.WaitForString ">", 5               
            strCmd = "connect port async " & strPortNumber & vbcr & vbcr & vbcr
            g_objtab.Screen.Send strCmd
            g_objtab.screen.WaitForString vbcr, 5
			   
            If g_objtab.Session.Connected <> True Then 
				LogToFile (g_DeviceHostname & " - CONNECTION - Connection to terminal server FAILED.  Correct issue and try again. - " &_
							strConnectionType & "@" & strHostname & ":" & strPortNumber)
				Exit Function
			End If
 
            'crt.Screen.WaitForString(">")
            Get_ConnectConsolePort = True
          Case "CISCO_TERM"
            strCmd = "telnet " & strHostname & " " & strPortNumber & vbcrlf
            g_objtab.Screen.Send strCmd
            g_objtab.Screen.Synchronous = True
            crt.sleep 1000
               g_objtab.screen.send vbcrlf & vbCrlf
           If g_objtab.Session.Connected <> True Then 
				LogToFile (g_DeviceHostname & " - CONNECTION - Connection to terminal server FAILED.  Correct issue and try again. - " &_
							strConnectionType & "@" & strHostname & ":" & strPortNumber)
				Exit Function
			End If
 
            g_objtab.Screen.WaitForString("Open"), 15
            Get_ConnectConsolePort = True    
			
          Case "CISCO"
            strCmd = "telnet " & strHostname & " " & strPortNumber & vbcrlf
            g_objtab.Screen.Send strCmd
            g_objtab.Screen.Synchronous = True
            crt.sleep 1000
               g_objtab.screen.send vbcrlf & vbCrlf
           If g_objtab.Session.Connected <> True Then 
				LogToFile (g_DeviceHostname & " - CONNECTION - Connection to terminal server FAILED.  Correct issue and try again. - " &_
							strConnectionType & "@" & strHostname & ":" & strPortNumber)
				Exit Function
			End If
 
            g_objtab.Screen.WaitForString("Open"), 15
            Get_ConnectConsolePort = True    
  
          Case "CONSOLE"
            g_objtab.Screen.Send strCmd
            g_objtab.Screen.Synchronous = True
 
           If g_objtab.Session.Connected <> True Then 
				LogToFile (g_DeviceHostname & " - CONNECTION - Connection to console port FAILED.  Correct issue and try again. - " &_
							strConnectionType & "@" & strHostname & ":" & strPortNumber)
				Exit Function
			End If
 
            'g_objtab.Screen.WaitForString(">")
               Get_ConnectConsolePort = True
 
        Case "SSH"
            ' Copy running config to SCP server, backup to local device
 
 
        Case Else
            ' Unsupported connection Type
            g_strError = "Unsupported connection Type: " & strProtocol
            Exit Function
    End Select
 
 
 End Function
 
 
 
Sub Get_Arguments2()
 
     If crt.Arguments.Count = 0 Then
          ' Display help menu
          Dim sErrorMessage
          sErrorMessage = "Deploy-IT Batch Jobs.vbs" & vbcrlf & vbcrlf &_
          "This script requires arguments to function properly. " & vbcrlf &_
          " At a minimum an input (-i:) and config file (-c:) have to be specified." & vbcrlf & vbcrlf &_
          "These items should be entered into individual strings using /arg to deliniate each one." & vbcrlf &_
          "e.g. SecureCRT /script myscript.vbs /arg -i:install.vbs /arg -c:config.txt" & vbcrlf & vbcrlf &_
          "Available Switches:" & vbcrlf &_
          " -i:" & vbTab & "Specify path to input file. (Default to My Documents " & vbcrlf & vbTab & "if path not specified)" & vbcrlf &_
          " -c:" & vbTab & "Specify the configuration file to be used with input file." &vbcrlf & vbTab & "(Default to My Documents if path not specified)" & vbcrlf &_
          " -u:" & vbTab & "Username" & vbcrlf &_
          " -p:" & vbTab & "Password" & vbcrlf &_
          " -t:" & vbTab & "Connection Type (Options are SSH, Cisco, MRV, or Console)" & vbcrlf &_
          " -h:" & vbTab & "Hostname/IP for connection (Required for SSH, Cisco, or MRV)" & vbcrlf &_
          " -a:" & vbTab & "Specify Terminal Server port for Cisco, MRV, or Console " &vbcrlf & vbTab & "(COM1,COM2,etc)"
          MsgBox sErrorMessage
      End If
 
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
          '     msgBox "Argument is: " & oArg (sArgs) & vbcrlf & _
          '               " sArgs is set to : " & sArgs
               sArgs = sArgs - 1
               if sArgs = -1 then Exit Do
          Loop
 
      For Each i in oArg
            sSwitch = LCase(Left(i,3))
            sValue = Right(i,Len(i)-3)
               
               Select Case sSwitch
                    Case "-i:"
                         sInput = sValue
                    Case "-c:"
                         sConfigFile = sValue
                    Case "-p:"
                         sPassword = sValue
                    Case "-u:"
                         sUsername = sValue
                    Case "-t:"
                         sConnectionType = sValue
                    Case "-a:"
                         sPortNumber = sValue
                    Case "-h:"
                         sHostname = sValue
               End Select
          Next
 
 
      'crt.dialog.MessageBox "Connection Type: " & sConnectionType
      
           If (lCase(trim(sConnectionType)) <> "serial" AND trim(sUsername) = "") Then
               Do While sUsername = ""
                    sUserName = crt.Dialog.Prompt("Please specify a username (Required)", "Specify Username", "Username",False)
               Loop
           End If
 
          If (lCase(trim(sConnectionType)) <> "serial" AND trim(sPassword) = "") Then
               Do While sPassword = ""
                    sPassword = crt.Dialog.Prompt("Password not specified on command line.  Please specify global password variable." &_
                                                  vbcrlf & "Default enable password will be set to the same if none specified globally.", "Specify Password", "", True)
               Loop
           End If 
 
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
          '     msgBox "Argument is: " & oArg (sArgs) & vbcrlf & _
          '               " sArgs is set to : " & sArgs
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
 
      g_UpdateIOSFile = "Yes"
     g_objtab.Screen.Synchronous = True     
 
     g_objtab.screen.send "dir | include " & left(g_DeviceIOS,8) & ".*" & right (g_DeviceIOS,4) & vbCr
     nResult = g_objtab.Screen.ReadString (g_Prompt,30)
     'msgbox "Dir output is: " & nresult & vbcrlf & "It should be: " & g_DeviceIOS
 
     'msgbox nresult & vbcrlf & g_DeviceIOS
 
     set reIOSPattern = New RegExp
     reIOSPattern.pattern = ".*("& g_DeviceIOS & ")"
 
     Set matches = reIOSPattern.Execute(nResult)
          For each match in matches
               sIOSPattern = match.SubMatches(0)
          Next     
 
     if sIOSPattern = g_DeviceIOS Then 
          g_UpdateIOSFile = "no"
          SendCommand ("! IOS File does not need to be updated because the following file was found on flash: " & sIOSPattern)
     Else
          SendCommand ("! IOS Files Needs to be updated:  " & g_UpdateIOSFile)
     End If
 
 End Sub
 
Sub Get_PrepareDevice
 
 
      crt.Screen.Synchronous = True
 
	
     ' Delete VLAN.dat if device is a switch
 
     If lcase(g_DeviceType) = "switch" Then
			
		If lcase(g_DeviceDeleteCatFlash) = "yes" Then
			' if catflash then delete cat4000 instead of regular flash
			g_objtab.screen.send "delete cat4000_flash:vlan.dat" & vbCr
		Else	 
		  g_objtab.screen.send "delete " & g_DeviceFlash & ":vlan.dat" & vbCr
		End If
		nResult = g_objtab.screen.WaitForStrings ("Delete fil", g_Prompt, 1)
		If nResult = 1  Then SendCommand (vbcrlf)
		nResult = g_objtab.screen.WaitForStrings ("[confirm]", g_Prompt, 1)
		If nresult = 1 Then g_objtab.screen.send "y"
		g_objtab.screen.WaitForString g_Prompt, 3
		LogToFile (g_DeviceHostname & " - PREPARATION - VLAN.dat deleted - Success")
		LogToPrompt ("VLAN.dat Deleted:                SUCCESS")
     End If

     g_objtab.screen.send "erase startup-config" & vbCr
     nResult = g_objtab.screen.WaitForStrings ("[confirm]" , 20)
     'msgbox ("Erase Config found confirm prompt: " & nResult)
     if nresult = 1 Then
          g_objtab.screen.send "y"
          g_objtab.screen.WaitforString "complete", 20
          LogToFile (g_DeviceHostname & " - PREPARATION - NVRAM Erase - Success")
          LogToPrompt ("NVRAM Deleted:                    SUCCESS")
     End If
	 
	if g_DeviceHardware = "WS-C4507" Then
		g_objtab.screen.send "license right-to-use activate ipbase accepteula" & vbCr
		g_objtab.screen.WaitForString g_Prompt, 1
	End If
	
     g_objtab.screen.send vbCr
     g_objtab.screen.WaitForString vbCr
     
     ' ** Reload the device once flash has been cleared.
     g_objtab.screen.send ("reload") & vbCr
     g_objtab.screen.WaitForString vbcr, 1
     
     
     nResult = g_objtab.screen.WaitForStrings ("[yes/no]","[confirm]",2)
     
     ' ** Don't save the configuration if it was modified
     If nResult = 1 Then 
          g_objtab.screen.send "no" & vbCr
     End if
     
     LogToFile (g_DeviceHostname & " - PREPARATION - Reload - Did not save running config - Success")
     
     ' ** Confirm reboot
     g_objtab.screen.send "y" & vbCrlf
          
     Get_AnswerRebootPrompts
          
     
 End Sub 
 
Sub Get_AnswerRebootPrompts ()
	Dim strCleanBootPrompt
	strCleanBootPrompt = Array("Router#","Switch#","Router>","Switch>")
	g_objTab.Screen.Synchronous = True

	nResult = g_objtab.screen.WaitForStrings ("Press RETURN to get started!", "%PNP-6-HTTP_CONNECTING", "%Error opening tftp:", "System Configuration Dialog","yes/no","Router#","Switch#","Router>","Switch>", 2000)
     'msgBox "First reboot nresult = " & nresult
     'crt.sleep 10000               ' Sleep 10 seconds
     
     ' ** Send a few carriage returns after successful reboot.
     g_objtab.screen.send vbcr
     g_objtab.screen.WaitForString vbcr, 10
     g_objtab.screen.send vbcr
     g_objtab.screen.WaitForString vbcr, 1
     g_objtab.screen.send vbcr
     g_objtab.screen.WaitForString vbcr, 1
     'crt.sleep 3000
     g_objtab.screen.send vbcr
     g_objtab.screen.WaitForString vbcr, 3
     g_objtab.screen.send vbcr
     g_objtab.screen.WaitForString vbcr, 3
     g_objtab.screen.send vbcr
     g_objtab.screen.WaitForString vbcr, 3
     g_objtab.screen.send vbcr
     nResult = g_objtab.screen.WaitForStrings ("yes/no", "'yes' or 'no'","Router#","Switch#","Router>","Switch>" , "%PNP-6-HTTP_CONNECTING", "%Error opening tftp:", 2000)
	 'msgBox "Second reboot nresult = " & nresult
     if nresult = 2 Then nresult = 1
     ' ** If result is yes/no, then send the no string to terminal
     if nResult = 1 Then
          g_objtab.screen.send "no" & vbCr
          iResult = g_objtab.screen.WaitForStrings (">","terminate autoinstall?", 5)
          if iResult = 2 then 
               g_objtab.screen.send "yes" & vbCr
               g_objtab.screen.WaitForStrings "Press Return to get started!",">",10
               SendCommand (vbcr)
               SendCommand (vbcr)
               
          End If
     End if
     g_objtab.screen.send vbcr
     nResult = g_objtab.screen.WaitForStrings ("Router#","Switch#","Router>","Switch>", 100)
	Get_Prompt
	if nResult <> "" Then
		 
		SendCommand ("enable")
		SendCommand ("! Router has now rebooted.  Beginning setup.")
		LogToFile (g_DeviceHostname & " - PREPARATION - Router has now rebooted.  Beginning setup.")
		LogToPrompt ("Device Reload:                    SUCCESS")
	Else
		LogToFile (g_DeviceHostname & " - PREPARATION - Answer Prompts after reboot has FAILED.  Correct issue and try again.")
		LogToFile (g_DeviceHostname & " - PREPARATION - Answer Prompts after reboot has FAILED.  Contact techincal support.")
		LogToPrompt ("Device Reload:    FAILED TO ANSWER REBOOT PROMPTS")
		g_ApplicationFailed = "yes"
	End If
	
	
 End Sub     
 
Sub Get_AnswerBootPrompts ()
	Dim strCleanBootPrompt
	strCleanBootPrompt = Array("Router#","Switch#","Router>","Switch>")
	g_objtab.screen.synchronous = True
	
	g_objtab.screen.send vbCr
	nResult = g_objtab.screen.WaitForStrings ("Press RETURN to get started!", "%PNP-6-HTTP_CONNECTING", "%Error opening tftp:", "System Configuration Dialog","yes/no", "Router#","Switch#","Router>","Switch>", 1000)
	'msgbox nResult
     ' ** Send a few carriage returns after successful reboot.
     g_objtab.screen.send vbcr
     g_objtab.screen.WaitForString vbcr, 10
     g_objtab.screen.send vbcr
     g_objtab.screen.WaitForString vbcr, 1
     g_objtab.screen.send vbcr
     g_objtab.screen.WaitForString vbcr, 1
     g_objtab.screen.send vbcr
     nResult = g_objtab.screen.WaitForStrings ("initial configuration dialog?", "'yes' or 'no'", "Router#","Switch#","Router>","Switch>", "%PNP-6-HTTP_CONNECTING", "%Error opening tftp:", 60)
     'msgBox "nresult = " & nresult
     if nresult = 2 Then nresult = 1
     ' ** If result is yes/no, then send the no string to terminal
     if nResult = 1 Then
          g_objtab.screen.send "no" & vbCr
          iResult = g_objtab.screen.WaitForStrings (">","[yes]", 5)
          if iResult = 2 then 
               g_objtab.screen.send "yes" & vbCr
               g_objtab.screen.WaitForStrings "Press Return to get started!",">",10
               SendCommand (vbcr)
               SendCommand (vbcr)
               
          End If
     End if
     
     Get_Prompt
	 
	 
 End Sub    
Sub Get_ConfigDevice ()
 
 
 
    'crt.Screen.IgnoreEscape = True
     SendCommand ("term length 0")    
     SendCommand ("term width 120")
     SendCommand ("config t")
     SendCommand ("no ip domain-lookup")
     SendCommand ("line con 0")
     SendCommand ("exec-timeout 360")
     SendCommand ("no logging synchronous")
     SendCommand ("logging console")
     SendCommand ("ip domain-name duke-energy.com")
     SendCommand ("hostname " & g_DeviceHostname)
	 strOldPrompt = g_Prompt
	 strOldPrompt = replace(strOldPrompt,"#","")
	 strOldPrompt = replace(strOldPrompt,">","")
     g_Prompt = g_DeviceHostname & "#"
     SendCommand ("interface " & g_DeviceProvisioningIF)
     SendCommand ("shut")
     SendCommand ("no shut")
	 SendCommand ("ip address dhcp")
	 
	 
     'crt.sleep 3000
     'SendCommand (vbcr)
     sResult = crt.screen.ReadString ("hostname " & g_DeviceHostname, 60) ' look for output that DHCP address received.  Wait 60 seconds.
     'msgbox sresult     
     
     If sResult <> "" Then
          'SendCommand (vbcr)
          'SendCommand ("!DHCP Address received")
     set reAddress = New RegExp
     reAddress.pattern = ".*DHCP address (([0-9]{1,3}\.){3}[0-9]{1,3})"
     
     set reMask = New RegExp
     reMask.pattern = ".*, mask (([0-9]{1,3}\.){3}[0-9]{1,3})"
     
     Set matches = reAddress.Execute(sResult)
               For each match in matches
                    sAddress = match.SubMatches(0)
               Next     
     Set matches = reMask.Execute(sResult)
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
		  g_ApplicationFailed = "yes"
		  SendCommand ("hostname " & strOldPrompt)
     End If
     
     SendCommand ("end")
     
 End Sub
Sub Get_CopySCPFiles ()
	
     g_objtab.Screen.Synchronous = True
     
     'crt.screen.send vbcrlf & "!g_Prompt : " & g_Prompt & vbcrlf &_
     '          "!g_PromptExtension : " & g_PromptExtension & vbCrlf
     'msgbox "copy scp://" & g_SCP_Username & ":" & g_SCP_Password & "@" g_SCP_Host & "/" & objFSO.GetFileName(objFile) & " flash:"
     
     g_objtab.screen.send vbCr
     g_objtab.screen.WaitforString g_PromptExtension, 3
     g_objtab.screen.send vbCr
     g_objtab.screen.WaitforString g_PromptExtension, 3
     
     SendCommand ("!Starting Get_CopySCPFiles procedure")
     'msgbox "Config filename is : " & g_ConfigFileName
     LogToFile (g_DeviceHostname & " - CONFIG FILE COPY - Begin Config File (" & g_ConfigFileName & ") Copy to " & g_DeviceFlash & ":")
     g_objtab.screen.send "copy scp://" & g_SCP_Username & ":" & g_SCP_Password & "@" & g_SCP_Host & "/" &_
                                        g_ConfigFileName & " " & g_DeviceFlash & ":" & vbCr
     nResult = g_objtab.screen.WaitForStrings ("Destination filename",10)
     if nResult = 1 Then 
		g_objtab.screen.send vbCrlf
		nresult2 = g_objtab.screen.WaitForStrings ("[confirm]", "% Destination unreachable", "bytes/sec", 20)
		if nResult2 = 1 Then     
			g_objtab.screen.send vbCrlf
			nResult3 = g_objtab.screen.WaitForStrings ("bytes/sec", "Error", 90)
			if nResult3 = 1 Then
				LogToFile (g_DeviceHostname & " - CONFIG FILE COPY - Config File: " & g_ConfigFileName & " copied successfully to " & g_DeviceFlash & ":")
				LogToPrompt ("Copy Config to Flash:" & "              SUCCESS")
				'msgbox "File copied successfully"
			Else
				LogToFile (g_DeviceHostname & " - CONFIG FILE COPY - Config File: " & g_ConfigFileName & " FAILED copying file to " & g_DeviceFlash & ":")
				LogToPrompt ("Copy Config to Flash:" & "            FAILED")
			End If
		ElseIf nResult2 = 2 Then
				LogToFile (g_DeviceHostname & " - CONFIG FILE COPY - Config File: " & g_ConfigFileName & " FAILED copying file to " & g_DeviceFlash & ":")
				LogToPrompt ("Copy Config to Flash:" & "            FAILED")
				g_ApplicationFailed = "yes"
		End If
	End If
		'SendCommand (vbcr)
     SendCommand ("! Sync back up in script")
     
     g_objtab.screen.send vbCr
     g_objtab.screen.WaitforString g_PromptExtension, 10
     g_objtab.screen.send vbCr
     g_objtab.screen.WaitforString g_PromptExtension, 10
     
     If lcase(g_UpdateIOSFile) = "yes" Then 
          LogToFile (g_DeviceHostname & " - IOS FILE COPY - Begin IOS File (" & g_DeviceIOS & ") Copy to " & g_DeviceFlash & ":")
          g_objtab.screen.send "copy scp://" & g_SCP_Username & ":" & g_SCP_Password & "@" & g_SCP_Host & "/" &_
                                   g_DeviceIOS & " " & g_DeviceFlash & ":" & vbCrlf & vbcrlf
          'SendCommand ("! IOS Would have copied here.")
          nResult = g_objtab.screen.WaitForStrings ("bytes copied in", "Error", 2500)
          if nResult = 1 Then 
               LogToFile (g_DeviceHostname & " - IOS FILE COPY - IOS Copied successfully.")
               LogToPrompt ("IOS File Copy:                  SUCCESS")
			Else
				LogToFile (g_DeviceHostname & " - IOS FILE COPY - IOS Copied FAILED to copy.")
				LogToPrompt ("IOS File Copy:                  FAILED")
				g_ApplicationFailed = "yes"
          End If
     Else
          SendCommand ("! IOS FIle " & g_DeviceIOS & " already exists.  Skipping this step.")
          LogToPrompt ("IOS File Copy:                 N/A (Already exists)")
          LogToFile (g_DeviceHostname & " - IOS FILE COPY - File:" & g_DeviceIOS & " - Already exists.")
     End If
     
 End Sub
 
Sub Get_InstallIOSFiles ()
 
     SendCommand (vbcr)
     SendCommand (vbcr)
     SendCommand ("!Starting Get_InstallIOS Files procedure")
     
     If lcase(g_DeviceIOSInstall) = "archive" Then
          LogToFile (g_DeviceHostname & " - SOFTWARE INSTALL - Archive Download-sw process started.")
          SendCommand ("! Starting the ARCHIVE Process.  ")
          SendCommand ("! This could take up to 20 minutes.")
          SendCommand ("! Do not power off or disconnect from ")
          SendCommand ("!    the network during this period.")
          
          g_objtab.screen.send "archive download-sw /overwrite " & g_DeviceFlash & ":/" & g_DeviceIOS & vbCr
          nResult = g_objtab.screen.WaitForStrings ("All software images installed.","ERROR:", 1200)
          
          if nResult = 1 Then
               LogToFile (g_DeviceHostname & " - SOFTWARE INSTALL - Archive Download-sw process completed successfully.")
               LogToPrompt ("Software Install: Archive         SUCCESS")
          Else
               LogToFile (g_DeviceHostname & " - SOFTWARE INSTALL - Archive process FAILED.")
               LogToPrompt ("Software Install: Archive          FAILED")
			   g_ApplicationFailed = "yes"
          End If
          
          
          'g_objtab.screen.WaitForString vbCr, 1
          'SendCommand ("! would have archived software here.")
     ElseIf lcase(g_DeviceIOSInstall) = "software install" Then
          g_objtab.screen.send "software install file " & g_DeviceFlash & ":" & g_DeviceIOS & " force" & vbCr
          nResult = g_objtab.screen.WaitForStrings ("operation aborted", "Do you want to proceed with reload? [yes/no]", 400)
		if nresult = 1 Then
				strIOSInstallStatus = "Failed"
				LogToFile (g_DeviceHostname & " - SOFTWARE INSTALL - Install Processess Aborted - FAILED")
				LogToPrompt ("Software Install: Software Install              FAILED")
			elseif nResult = 2 Then
               g_objtab.screen.send "yes" & vbcr
               iResult = g_objtab.screen.WaitForStrings (g_Prompt, "[yes/no]", 5)
				if iResult = 2 Then 
					g_objtab.screen.send "no" & vbcr
					nResult2 = g_objtab.screen.WaitForStrings ("Reloading", g_Prompt,30)
					if nResult = 2 Then
						strIOSInstallStatus = "Success"
						LogToFile (g_DeviceHostname & " - SOFTWARE INSTALL - Install Processess Successfully installed " &_
									  g_DeviceFlash & ":" & g_DeviceIOS & " - SUCCESS")
						LogToPrompt ("Software Install: Install           SUCCESS")
					Else
						strIOSInstallStatus = "Failed"		
						LogToFile (g_DeviceHostname & " - SOFTWARE INSTALL - Install Processess Aborted - FAILED")
						LogToPrompt ("Software Install: Software Install              FAILED")
						g_ApplicationFailed = "yes"
					End If
				End If
		End If
          ' ** Call sub to wait for reboot and answer reboot prompts.
          if nresult <> 1 Then Get_AnswerRebootPrompts
     ElseIf lcase(g_DeviceIOSInstall) = "standard" Then
          LogToFile (g_DeviceHostname & " - SOFTWARE INSTALL - Standard IOS Image (.bin).  No action necessary.")
          LogToPrompt ("Software Install: Standard             N/A")
     End If
          
 End Sub
 
Sub Get_InstallConfigFiles ()
 
     SendCommand (vbcr)
     SendCommand (vbcr)
     SendCommand ("!Starting Get_InstallConfigFiles procedure")
	 SendCommand (vbcr)
     
 
     ' ** Copy config file to running-config
     g_objtab.screen.send "copy " & g_DeviceFlash & ":" & g_ConfigFileName & " running-config" & vbCr
     g_objtab.screen.WaitForString "Destination filename",1
     g_objtab.screen.send vbCrlf
     nresult = g_objtab.screen.WaitForStrings ("[confirm]", "No such file or directory", "The name", 10)
     if nResult = 1 Then     g_objtab.screen.send "y"
     ' ** If File is not found, write error to log.
	if nResult = 2 Then 
		'MSgbox "failed"
		crt.screen.send vbCr
	End If
     'msgbox g_Prompt
     sResult = g_objtab.screen.ReadString (g_Prompt, 120)
     
     if nResult <> 2  Then 				' If install was successful perform these actions.
        LogToFile (g_DeviceHostname & " - CONFIG INSTALL - Configuration copied to running-config.")
        LogToPrompt ("Copy Config to Running:           SUCCESS")
		sResult = split(sResult,vbcr)
		 
		 
		' ** Add date/time stamp to each line in the output from copying config to running
		sresult3 = vbcrlf
			  
		For each i in sResult
			  sResult3 = sResult3 & Now & "   " & g_DeviceHostname & " - CONFIG INSTALL - " & i & vbCrlf
		next
		 
		LogToFile (sresult3)
		'SendCommand ("! Config file would have copied into running")
		g_objtab.screen.send "write mem" & vbCr
		' Verify OK Prompt is received
		nresult = g_objtab.screen.WaitForStrings ("[OK]",20)
		If nResult = 1 Then
			' Write to log confirming config copied succesfully
			LogToFile (g_DeviceHostname & " - CONFIG INSTALL - Configuration written to memory successfully.")
			LogToPrompt ("Save to startup-config:           SUCCESS")
		Else
			LogToFile (g_DeviceHostname & " - CONFIG INSTALL - FAILED - Copy error writing to NVRAM.")
			LogToPrompt ("Save to startup-config:           FAILED")
			g_ApplicationFailed = "yes"
		End If
		
     Else
        LogToFile (g_DeviceHostname & " - CONFIG INSTALL - Configuration FAILED to running-config. File Not Found - FAILED.")
        LogToPrompt ("Copy Config to Running:           FAILED")
		LogToPrompt ("Save to startup-config:           N/A")		  
		g_ApplicationFailed = "yes"
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
     If fso.FileExists(DestinationFile) Then 
          LogToFile (g_DeviceHostname & " - FILE COPY - File " & SourceFile & " copied successfully to " & DestinationFile)
          LogToPrompt ("File Copy:          Success")
     Else 
          LogToFile (g_DeviceHostname & " - FILE COPY - File " & SourceFile & " DID NOT COPY TO " & DestinationFile & " FAILED")
          LogToPrompt ("File Copy:          Failed")
		  g_ApplicationFailed = "yes"
     End If
    If wasReadOnly Then
        'Reapply the read-only attribute
        fso.GetFile(DestinationFile).Attributes = fso.GetFile(DestinationFile).Attributes + 1
    End If
    Set fso = Nothing
 End Sub
 
 
 Function SendCommand (strComment)
	if strComment <> vbcr Then strComment = strComment & vbcr
    g_objtab.screen.Send(strComment)
       g_objtab.screen.WaitForString g_Prompt, 3
     
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
  
    If lcase(g_Log_FileLocation) = "relative" Then
        Set oLogShell = CreateObject("Wscript.Shell")
        g_Log_FileLocation = oLogShell.CurrentDirectory & "\"
        Set oLogShell = Nothing
    End If
   
   'msgbox g_Log_FileLocation
   
    If g_Log_DateStampFileName Then
        sNow = Replace(Replace(Now(),"/","-"),":",".")
        sLog_FileName = sNow & " - " & g_Log_FileName
        g_Log_DateStampFileName          = False      
     Else
          sLog_FileName = g_Log_FileName
    End If
 
     'msgbox "Log file is: " & g_Log_FileLocation & sLog_FileName
     'Check if log file location exists.  If not create folder.
	 
	If oLogFSO.FolderExists(g_Log_FileLocation) Then
		' Do Nothing
	Else
		oLogFSO.CreateFolder (g_Log_FileLocation)
	End If

	
	 if right(g_Log_FileLocation,1) <> "\" then g_Log_FileLocation = g_Log_FileLocation & "\"
	
    sLog_File = g_Log_FileLocation & sLog_FileName
   
   
   
   
   
   
    If g_Log_OverWriteOrAppend = "overwrite" Then
        Set oLogFile = oLogFSO.OpenTextFile(sLog_File, ForWriting, True)
        g_Log_OverWriteOrAppend = "append"
    Else
		'msgbox sLog_File
        Set oLogFile = oLogFSO.OpenTextFile(sLog_File, ForAppending, True)
    End If
 
    If g_Log_DateStamp Then
        strMessage = Now & "   " & strMessage
    End If
 
    oLogFile.WriteLine(strMessage)
    oLogFile.Close
    
    oLogFSO = Null
End Sub



Function ConvertTime(intTotalSecs)
	Dim intHours,intMinutes,intSeconds,Time
	intHours = intTotalSecs \ 3600 
	intMinutes = (intTotalSecs Mod 3600) \ 60
	intSeconds = intTotalSecs Mod 60
	If intHours = 0 Then intHours = "0"&intHours 
	If intMinutes = 0 Then intMinutes = "0"&intMinutes 
	If intSeconds = 0 Then intSeconds = "0"&intSeconds 
	ConvertTime = LPad(intHours) & ":" & LPad(intMinutes) & ":" & LPad(intSeconds)
End Function


Function LPad(v) 
    LPad = Right("00" & v, 2) 
End Function
