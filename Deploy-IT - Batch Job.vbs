#$Language="VBScript"
#$Interface="1.0"

' The purpose of this script is to execute Installation files that follow Gen.IT output.
Dim g_AppName 
 g_AppName = "Install Gen.IT Files v0.2"
Dim g_shell
Set g_shell = CreateObject("WSCript.Shell")

Dim oArg, sInput, sConfigFile, sUsername, sPassword, sSwitch, sValue, sConnectionType, sPortNumber, sHostname

Dim g_Prompt, g_hostname


' Check to see if Arguments are specified on the command line.  If not, then 
Sub Main ()
	' Get Arguments from command line.  Requires input file and config file.
	Get_Arguments
	' Test to validate connection details are appropriate
	'If crt.Arguments.Count > 0 Then Display_Connection_Details
	
	
		If lcase(sConnectionType) = "mrv" Then
			Call Connect(sHostname, 22, "SSH2", sUsername, sPassword)
			Call Get_ConnectConsolePort (sConnectionType, sPortNumber, sHostname)
		End If
		
		If lcase(sConnectionType) = "cisco" Then 
			Call Connect(sHostname, 22, "SSH2", sUsername, sPassword)
			Call Get_ConnectConsolePort (sConnectionType, sPortNumber, sHostname)
		End If
		
		If lcase(sConnectionType) = "serial" Then 
			Call Connect("", sPortNumber, "SERIAL","","")
		End If
		
		'Get_Prompt
		' Execute enable subrouting if not an MRV
		'If lcase(sConnectionType) <> "mrv" Then Get_Enable

		'If crt.Session.Connected then
		'				' if > is found in prompt, execute enable script.
		'				set reEnableTest = New RegExp
		'				reEnableTest.pattern = ".*(\>)$"
		'				
		'				If reEnableTest.Test(g_Prompt) = TRUE Then Get_Enable
		'	
		'End If	
	
 End Sub

 
 
 
 
Sub Get_Prompt
			If strConnect = "true" Then
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
					rePrompt.pattern = "(.*)[\>|#]"

					
					Set matches = rePrompt.Execute(g_Prompt)
						For each match in matches
							g_hostname = match.SubMatches(0)
						Next
					g_objWriteSheet.Cells(nRowIndex, g_HOSTNAME_COL).Value = g_hostname	
                    If g_Prompt <> "" Then Exit Do
					
					
					
                Loop
			End If
		
 End Sub
 
 
Sub Get_Enable
	'check if prompt is at enabled mode.  If not, type in enable.


	
	
 End Sub
	
	
	
Function Get_ConnectConsolePort (strConnectionType, strPortNumber, strHostname)
	    Select Case UCase(strConnectionType)
        Case "MRV"
			crt.sleep 1500
            strCmd = "connect port async " & strPortNumber & vbcrlf & vbcrlf & vbcrlf 
            crt.Screen.Send strCmd
            crt.Screen.Synchronous = True
            
            If crt.Session.Connected <> True Then Exit Function
            
            'crt.Screen.WaitForString(">")
            Get_ConnectConsolePort = True
          Case "CISCO"
            strCmd = "telnet " & strHostname & " " & strPortNumber & vbcrlf
            crt.Screen.Send strCmd
            crt.Screen.Synchronous = True
            crt.sleep 1000
			crt.screen.send vbcrlf & vbCrlf
            If crt.Session.Connected <> True Then Exit Function
            
            crt.Screen.WaitForString("Open"), 15
            Get_ConnectConsolePort = True    
			
		Case "CONSOLE"
            crt.Screen.Send strCmd
            crt.Screen.Synchronous = True
            
            If crt.Session.Connected <> True Then Exit Function
            
            'crt.Screen.WaitForString(">")
			Get_ConnectConsolePort = True
			
        Case "SSH"
            ' Copy running config to SCP server, backup to local device
           
           
        Case Else
            ' Unsupported connection Type
            g_strError = "Unsupported connection Type: " & strProtocol
            Exit Function
    End Select
 
 
 End Function
 

Sub Get_Arguments()
 
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
		'	msgBox "Argument is: " & oArg (sArgs) & vbcrlf & _
		'			" sArgs is set to : " & sArgs
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

 

 Sub Display_Connection_Details ()
	 crt.dialog.MessageBox "Input File:   " & vbTab & sInput & vbcrlf &_
				"Config File: " & vbTab & sConfigFile & vbcrlf &_
				"Username: " & vbTab & sUsername & vbcrlf &_
				"Password: " & vbTab & sPassword & vbcrlf &_
				"Connection Type:" & vbTab & sConnectionType & vbcrlf &_
				"Port Number:" & vbTab & sPortNumber & vbcrlf &_
				"Hostname:" & vbTab & sHostname & vbcrlf
 
 End Sub			



Function Connect(strIP, strPort, strProtocol, strUsername, strPassword)
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 ' Workaround that uses "On Error Resume Next" VBScript directive to detect
 ' Errors that might occur from the crt.Session.Connect call and instead of
 ' closing the script, allow for error handling within the script as the
 ' script author desires.
    On Error Resume Next
	'msgBox strPort & vbTab & strProtocol
    ' First disconnect if we're already connected.
	If strPort = "" Then Exit Function
    If crt.Session.Connected Then crt.Session.Disconnect

    g_strError = ""
    Err.Clear
    Dim strCmd
    Select Case UCase(strProtocol)
        Case "TELNET"
            strCmd = "/TELNET " & strIP & " " & strPort
            crt.Session.Connect strCmd
            crt.Screen.Synchronous = True
            
            If crt.Session.Connected <> True Then Exit Function
            
            ' Look for username and password prompts
            crt.Screen.WaitForString("sername:")
            crt.Screen.Send strUsername & vbcr
			crt.sleep 1000
            crt.Screen.WaitForString("ssword:")
            crt.Screen.Send strPassword & vbcr
			crt.sleep 1000
            Connect = True
            
		Case "SERIAL"
            strCmd = "/SERIAL " &  strPort & " /BAUD 9600 /DATA 8 /PARITY NONE /STOP 0 "
            crt.Session.Connect strCmd
			crt.window.caption = "Provision - " & sConfigFile
			crt.Screen.Send strPassword & vbcr & vbcr
            crt.Screen.Synchronous = True
            
            If crt.Session.Connected <> True Then Exit Function
            
            ' Look for username and password prompts
            'crt.Screen.WaitForString("ogin:")
            'crt.Screen.Send strUsername & vbcr
            'crt.Screen.WaitForString("ssword:")
            'crt.Screen.Send strPassword & vbcr
            Connect = True
			
        Case "SSH2", "SSH1"
			strCmd = _
                " /" & UCase(strProtocol) & " " & _
                " /L " & strUsername & _
                " /PASSWORD " & strPassword & _
                " /P " & strPort & _
                " /ACCEPTHOSTKEYS " & _
                strIP
			crt.Session.Connect strCmd
			crt.window.caption = "Provision - " & sConfigFile
            'msgBox "Connected? " & crt.session.connected
			If crt.Session.Connected <> True Then Exit Function
            crt.Screen.Synchronous = True
            Connect = True
           
           
        Case Else
            ' Unsupported protocol
            g_strError = "Unsupported protocol: " & strProtocol
            Exit Function
    End Select
	
    
    If Err.Number <> 0 Then
        g_strError = Err.Description
    End If
	
    On Error Goto 0
End Function ' End Function 