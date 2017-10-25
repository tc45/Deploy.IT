' GetInventory.vbs
' 
' This example Script shows how to read (and write) data to (and from) an
' Excel spreadsheet using VBScript running within SecureCRT.
'
' This specific example uses an excel spreadsheet that has the (Current spreadsheet column order changed.  Updated incidated below)
' following general format:
'          A       |   D   |     C    |     E    |     F     |    B    |  G   |    H    |
'   +--------------+-------+----------+----------+-----------+---------+------+---------+
' 1 | IP Address   | Port  | Protocol | Username | Password  | Active? | Date | Results | 
'   +--------------+-------+----------+----------+-----------+---------+------+---------+
' 2 | 192.168.0.1  | 22    | SSH2     | admin    | p4$$w0rd  |   Yes   | 1/11 |  Succ   |
'   +--------------+-------+----------+----------+-----------+---------+------+---------+
' 3 | 192.168.0.2  | 23    | Telnet   | root     | NtheCl33r |   No    | 1/11 |  Fail   |
'   +--------------+-------+----------+----------+-----------+---------+------+---------+
' 4 | 192.168.0.3  | 22    | SSH2     | root     | s4f3rN0w! |   Yes   | 1/11 |  Succ   |
'   +--------------+-------+----------+----------+-----------+---------+------+---------+
'
'	To-Do List --------------------------------------------------------------------------
'	Add Show Int Description; show int status to Interfaces tab.
'	Add SNMP Location info to Main tab
'	Parse show int desc|show int status, get interface name, split to create regex, compare to show ip int brief.  Output should include interface name, description, trunk/access, 
'			vlan info, Port type, speed, Duplex, and status. 
'	Add Settings tab and have variable input
'	Add ARG to specify the path to the XLS file rather than hard coding in the VBS.
'	
'	Version Info
' 	v1.5.5 - Snapshot with completed LogOutputOfSpecificCommand into individual text files.  
' 				Read Excel command tab for list of commands and add to array (next step - replace hard coded commands with Excel file commands.)
'	v1.5.6.1 - Completed Get_Interfaces subroutine. 
'	v1.5.6.2 - Completed most of Get_show_ver subroutine.  IOS REGEX still not working.
'	v1.5.6.3 - Parse full list of devices, username, and password.  If missing any, prompt for username/password and enable.
'	v1.5.6.4 - Added enable handler.  Works as long as the password is correct.  Need to update to add additional prompt if password doesn't work.
'	v1.5.6.5 - Updated columns in XLS spreadsheet.  Active is now Column B, Protocol C, Port D, Username E, and Password F.  Copy input XLSX to new XLSX file 
'				from output file name specified in MAIN tab on XLSX spreadsheet.  Updated g-objExcel.Workbooks to specify a READ
'				and a WRITE workbook seperately.  Added Global Username/Password/Enable Pass function which will override
'				prompts if username not specified on the line item.  
'	v1.5.6.6 - Removed Ports as required field.  SSH2 defaults to 22, Telnet to 23, Blank console or incorrect protocol write error to spreadsheet.
'				Parse CDP data and add to new tab in spreadsheet.
'	v1.5.6.7.2 - Cut to production version GetInventory_v0.5				

Dim g_shell
Set g_shell = CreateObject("WSCript.Shell")

Dim g_fso
Set g_fso = CreateObject("Scripting.FileSystemObject")

Const ForReading   = 1
Const ForWriting   = 2
Const ForAppending = 8

Dim g_strMyDocs, g_strSpreadSheetPath
g_strMyDocs = g_shell.SpecialFolders("MyDocuments")
g_strSpreadsheet = "ExcelSample.xlsx"
g_strSpreadSheetPath = g_strMyDocs & g_strSpreadsheet

If crt.Arguments.Count < 1 Then
	If Not g_fso.FileExists(g_strSpreadSheetPath) Then
		msgBox "No spreadsheet specified.  Please specify one now."
		g_strSpreadSheetPath = crt.Dialog.FileOpenDialog("Please specify an input spreadsheet", "Open", g_strMyDocs & "\ExcelSample.xlsx",_
			"Excel Files (*.xlsx)|*.xlsx|Excel 2003-2007 Files (*.xls)|(*.xls)||")
	End If

 Else
	g_strSpreadSheetPath = crt.Arguments(0)
End If


Dim g_objExcel
Dim g_Prompt, g_hostname, g_IF_k, g_CDPnRowIndex, g_IF_first

Dim g_IP_COL, g_PORT_COL, g_USER_COL, g_PASS_COL, g_ACT_COL
Dim g_Username, g_Password, g_EnablePass

Set g_objExcel = CreateObject("Excel.Application")

Dim g_objReadWkBook
Set g_objReadWkBook = g_objExcel.Workbooks.Open(g_strSpreadSheetPath) 
Dim g_objReadSheet
Set g_objReadSheet = g_objReadWkBook.Sheets("Main")

Dim g_objWriteWkBook, g_objWriteSheet




Dim nRowIndex
' Skip the header row.  If your sheet doesn't have a header row, change the
' value of nRowIndex to '1'.
nRowIndex = 9

' Create global variables for row numbers used for Global username/password/enable specified on the MAIN tab of the spreadsheet.  
Dim nUsernameIndex, nPasswordIndex, nEnableIndex, nSettingsColIndex

nUsernameIndex = 5
nPasswordIndex = 6
nEnableIndex = 7
' Specify the column that the variables are stored in on the MAIN tab of the spreadsheet.
nSettingsColIndex = 3





' Convert Letter column indicators to numerical references
g_IP_COL           = Asc("A") - 64
g_ACT_COL          = Asc("B") - 64
g_PROTO_COL        = Asc("C") - 64
g_PORT_COL         = Asc("D") - 64
g_USER_COL         = Asc("E") - 64
g_PASS_COL         = Asc("F") - 64
g_ACT_DATE_COL     = Asc("G") - 64
g_ACT_RES_COL      = Asc("H") - 64
g_HOSTNAME_COL     = Asc("I") - 64
g_COMMAND_COL	   = Asc("A") - 64


Dim strMyDate, strY, strM, strD

strY = Year(Date) 
strM = Month(Date) : If Len(strM)=1 Then strM = "0" & strM : End If
strD = Day(Date) : If Len(strD)=1 Then strD = "0" & strD : End If
strMyDate=strY&"-"&strM&"-"&strD

' Pulled from LogOutputofSpecificCommand-UseReadString.vbs



Dim g_szLogFile, g_szFirstLogFilePath, objTab, 	g_objFile
g_szLogFile = GetMyDocumentsFolder & "\Command#__NUM___Results.txt"

Set objTab = crt.GetScriptTab

Dim g_vCommands(100)

Dim g_vCommandsSet, g_vVersionSet, g_vInterfaceSet, g_vCDPSet, g_vSFPSet, g_vSNMPSet
Dim g_vCPUSet

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Sub Main() 
 
 
	
    ' Loop continues until we find an empty row
    Dim strIP, strPort, strProtocol, strUsername, strPassword, strActive
	
	' Get file output info from original XLSX, copy workbook to new file and open new Workbook.
	strOutputFileName = Trim(g_objReadSheet.Cells(3, 3).Value)
	'Verify if the filename contains .xls or .xlsx.  If not, add .xlsx
	set reFilename = New RegExp
	reFilename.pattern = "(.*)\.[xls|xlsx|xlsm]"
	If reFilename.Test(strOutputFileName) <> TRUE Then
		strOutputFileName = strOutputFileName & "-" & strMyDate & ".xlsx"
		'msgBox "Pattern did not match " & reFilename.Pattern & vbcrlf & vbcrlf & "New filename is : " & strOutputFileName
	Else
		Set matches = reFilename.Execute(strOutputFileName)
		For each match in matches
			strOutputFileName = match.SubMatches(0) & "-" & strMyDate & ".xlsx"
		Next
	End If
	' If no filename was specified, we will use the same input filename.	
	if strOutputFileName = "" Then strOutputFilename = g_strSpreadsheet
	
	'Read output directory specified in Main tab on the input spreadsheet.
	strOutputFileDir  = Trim(g_objReadSheet.Cells(4, 3).Value)
	If strOutputFileDir = "" Then strOutputFileDir = g_strSpreadSheetPath
	'Create full path using the string specified.
	strOutputFullPath = g_strMyDocs & strOutputFileDir & strOutputFileName
	
	'msgBox "Output directory is : " & strOutputFileDir
	
	MyMkDir (g_strMyDocs & strOutputFileDir)
	
		
	'msgBox "Original XLSX file will be copied to: " & vbcrlf & strOutputFullPath
	
	g_objReadWkBook.SaveAs (strOutputFullPath)

	'Copy existing XLSX file to a new file using the path specified.
	
	g_objReadWkBook.Close
	g_strSpreadSheetPath = g_strMyDocs & strOutputFileDir
	
	Set g_objWriteWkBook = g_objExcel.Workbooks.Open(strOutputFullPath) 
	Set g_objWriteSheet = g_objWriteWkBook.Sheets("Main")
	set strMainSheet = g_objWriteWkBook.Sheets("Main")
	
	strMainSheet.Unprotect

	' Get settings from settings tab.
	Get_Settings		
	' Run subroutine to check for login data. If login info not found, prompt for username/password.
	Get_LoginData
	'Get_ParseCommands is a required subroutine before Get_LogOutputOfSpecificCommandUseReadString is used.
	Get_ParseCommands

	
	
    Do
        ' If you find an empty value in column #1, exit the loop
        strIP = Trim(g_objWriteSheet.Cells(nRowIndex, g_IP_COL).Value)
        If strIP = "" Then Exit Do
		
		
        
        strActive = Trim(g_objWriteSheet.Cells(nRowIndex, g_ACT_COL).Value)
        'If Not Continue("Debug: Active(" & nRowIndex & ") = " & strActive) Then
        '    Exit Do
        'End If
            
        ' Connect to machines using the information we've gathered... only on
        ' condition of value in the g_ACT_COL column.
        If LCase(strActive) = "yes" Then
            strPort     = Trim(g_objWriteSheet.Cells(nRowIndex, g_PORT_COL).Value)
            strProtocol = Trim(g_objWriteSheet.Cells(nRowIndex, g_PROTO_COL).Value)
			'Check to see if Port was specified in spreadsheet.  If it was not specified then default SSH to 22 and Telnet to 23.  
			' Error on CONSOLE or unknown protocol.  Write results to spreadsheet on error.
			
			
			If strPort = "" Then
				Select Case ucase(strProtocol)
					Case "SSH2"
						strPort = "22"
						g_objWriteSheet.Cells(nRowIndex, g_PORT_COL).Value = "22"
					Case "TELNET"
						strPort = "23"
						g_objWriteSheet.Cells(nRowIndex, g_PORT_COL).Value = "23"
					Case "CONSOLE"
						g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = _
							"Failure: Console Port not specified."
					Case Else
						g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = _
							"Failure: Unknown Protocol Specified."
				End Select
			End If
            strUsername = Trim(g_objWriteSheet.Cells(nRowIndex, g_USER_COL).Value)
            strPassword = Trim(g_objWriteSheet.Cells(nRowIndex, g_PASS_COL).Value)
			
			If strUsername = "" Then strUsername = g_Username
			If strPassword = "" Then strPassword = g_Password
			If g_EnablePass = "" Then strEnablePass = strPassword
			
			'msgBox "StrUsername is " & strUsername & vbcrlf & "strPassword is : " & strPassword & vbcrlf & "Device ID : " & strIP & vbcrlf & vbcrlf & _
			'		"strPort is " & strPort & vbcrlf & "strProtocol is " & strProtocol
            
            ' For debugging output (msgbox), uncomment the following group of lines:
            'If Not Continue("Debug: Here's the information from row #" & _
            '    nRowIndex & ":" & vbcrlf & _
            '    "Column #" & g_IP_COL & ": " & strIP & vbcrlf & _
            '    "Column #" & g_PORT_COL & ": " & strPort & vbcrlf & _
            '    "Column #" & g_USER_COL & ": " & strUsername & vbcrlf & _
            '    "Column #" & g_PASS_COL & ": " & strPassword) Then Exit Do
			Dim strConnect
			If strProtocol = "" Then
				If Not ConnectWithFallback (strIP, strPort, strUsername, strPassword) Then
                ' Log an error, send e-mail?  For now, just mark it in the
                ' spreadsheet.
					If g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL).Value = "" Then
						g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = _
							"Failure: Unable to Connect"
					End If
					g_objWriteSheet.Cells(nRowIndex, g_HOSTNAME_COL).Value = "N/A"
					strConnect = "false"
					'msgBox "Connection failed"
				Else
					strConnect = "true"
				End If
			Else
				If Not Connect(_
					strIP, _
					strPort, _
					strProtocol, _
					strUsername, _
					strPassword) Then
					
					' Log an error, send e-mail?  For now, just mark it in the
					' spreadsheet.
					If g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL).Value = "" Then
						g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = _
							"Failure: Unable to Connect"
					End If
					g_objWriteSheet.Cells(nRowIndex, g_HOSTNAME_COL).Value = "N/A"
					
					' Heuristically determine the shell's prompt
					strConnect = "false"
				Else
					strConnect = "true"				
				End If
			End If
			
            
 
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
				
				If objTab.Session.Connected then
					' if > is found in prompt, execute enable script.
					set reEnableTest = New RegExp
					reEnableTest.pattern = ".*(\>)$"
					
					If reEnableTest.Test(g_Prompt) = TRUE Then Get_Enable
		
					'Get_Config
					If g_vVersionSet = "" Then g_vVersionSet = "On"
					If g_vVersionSet = "On" Then Get_show_ver
					If g_vInterfaceSet = "" Then g_vInterfaceSet = "On"
					If g_vInterfaceSet = "On" Then Get_Interfaces				
					If g_vSFPSet = "" Then g_vSFPSet = "On"
					If g_vSFPSet = "On" Then Get_SFPs
					If g_vSNMPSet = "" Then g_vSNMPSet = "On"
					If g_vSNMPSet = "On" Then Get_show_snmp
					If g_vCDPSet = "" Then g_vCDPSet = "On"
					If g_vCDPSet = "On" Then Get_ShowCDPneigh
					If g_vCPUSet = "" Then g_vCPUSet = "On"
					If g_vCPUSet = "On" Then Get_show_proc_cpu
					If g_vCommandsSet = "" Then g_vCommandsSet = "On"
					If g_vCommandsSet = "On" Then Get_LogCommands
					'crt.Screen.Send chr(26) & vbcr					
					'crt.Screen.Send "exit" & vbcr
					
					'If objTab.Session.Config.GetOption("Force Close On Exit") <> True Then
					'	crt.Dialog.MessageBox "(Before)" & vbcrlf & _
				'			"Force Close On Exit: " & objTab.Session.Config.GetOption("Force Close On Exit")
				'		objTab.Session.Config.SetOption "Force Close On Exit", True
				'		crt.Dialog.MessageBox "(After)" & vbcrlf & _
				'			"Force Close On Exit: " & objTab.Session.Config.GetOption("Force Close On Exit")
				'	End If
	
					objTab.Session.Disconnect
					
					
					'objTab.Session.Config.SetOption "Force Close On Exit", True
					'objTab.Session.Disconnect
					'objTab.Session.Disconnect
					'objTab.Session.Config.SetOption "Force Close On Exit", False
				end if				
				
            End If
			
			
 
			
        Else
            ' mark the skipped ones in the spreadsheet
            g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = "Skipped"
			g_objWriteSheet.Cells(nRowIndex, g_HOSTNAME_COL).Value  = "N/A"
            
        End If
        
        ' We always record the date of action status
        g_objWriteSheet.Cells(nRowIndex, g_ACT_DATE_COL).Value = Now
		
        
        ' move down to the next row in the spreadsheet
        nRowIndex = nRowIndex + 1
		g_objWriteWkBook.Save
    Loop
	
    g_objWriteWkBook.Close
        
    g_objExcel.Quit
    Set g_objExcel = Nothing
    
    g_shell.Run Chr(34) & strOutputFullPath & Chr(34)
	'g_shell.Run "explorer /e,/select,""" & g_szFirstLogFilePath & """"
    
End Sub
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~




Sub Get_Settings
 
	Set g_objWriteSheet = g_objWriteWkBook.Sheets("Settings")
 
    ' Loop continues until we find an empty row
    Dim strCommmands, nRowIndexCommand, nCommandIndex
	
	' Pull in values from the third column into global variables for settings.  Values hard coded.
	g_vCommandsSet = g_objWriteSheet.Range("vGetCommands").Value
	g_vVersionSet = g_objWriteSheet.Range("vGetVersion").Value
	g_vInterfaceSet = g_objWriteSheet.Range("vGetInterface").Value
	g_vCDPSet = g_objWriteSheet.Range("vGetCDP").Value
	g_vSFPSet = g_objWriteSheet.Range("vGetSFP").Value
	g_vCPUSet = g_objWriteSheet.Range("vGetCPU").Value
	
	
	
	
	
	'g_vCommandsSet = g_objWriteSheet.Cells(3, 3).Value
	'g_vVersionSet = g_objWriteSheet.Cells(4, 3).Value
	'g_vInterfaceSet = g_objWriteSheet.Cells(5, 3).Value
	'g_vCDPSet = g_objWriteSheet.Cells(6, 3).Value
	'g_vSFPSet = g_objWriteSheet.Cells(7, 3).Value
	'g_vCPUSet = g_objWriteSheet.Cells(8, 3).Value
		
	Set g_objWriteSheet = g_objWriteWkBook.Sheets("Main")
 
End Sub


Public Sub MyMkDir(sPath)
    Dim iStart
    Dim aDirs
    Dim sCurDir
    Dim i
 
    If sPath <> "" Then
        aDirs = Split(sPath, "\")
        If Left(sPath, 2) = "\\" Then
            iStart = 3
        Else
            iStart = 1
        End If
 
        sCurDir = Left(sPath, InStr(iStart, sPath, "\"))
 
        For i = iStart To UBound(aDirs)
            sCurDir = sCurDir & aDirs(i) & "\"
			If Not g_fso.FolderExists(sCurDir) Then
		'BuildFullPath g_fso.GetParentFolderName(FullPath)
			g_fso.CreateFolder sCurDir
            End If
        Next
    End If
End Sub


Sub Get_Hostname
 
  set reHostname = New RegExp
   reHostname.Pattern = "(.*)#"
   
 If reHostname.Test(g_Prompt) = TRUE Then	
	Set prompt_matches = reHostname.Execute(g_Prompt)
	For each match in prompt_matches
		g_hostname = trim(match.SubMatches(0))
		g_hostname = Replace(g_hostname,vbCrLf, "")
		g_hostname = Replace(g_hostname,vbCr, "")
		'msgBox "g_Prompt is now " & g_Hostname
	 Next
	g_objWriteSheet.Cells(nRowIndex, g_HOSTNAME_COL).Value = g_hostname
	'msgBox "Hostname Found" & vbcrlf & "Current Row is " & nRowIndex & vbcrlf & "Current Column is " & g_HOSTNAME_COL
	Else
		g_objWriteSheet.Cells(nRowIndex, g_HOSTNAME_COL).Value = "N/A"
	'msgBox "Hostname Found" & vbcrlf & "Current Row is " & nRowIndex & vbcrlf & "Current Column is " & g_HOSTNAME_COL
	
	End If
	
End Sub

Sub Get_LogCommands()
    ' Instruct WaitForString and ReadString to ignore escape sequences when
    ' detecting and capturing data received from the remote (this doesn't
    ' affect the way the data is displayed to the screen, only how it is handled
    ' by the WaitForString, WaitForStrings, and ReadString methods associated
    ' with the Screen object.
    objTab.Screen.IgnoreEscape = True
    objTab.Screen.Synchronous = True
   
    If Not objTab.Session.Connected then
        'crt.Dialog.MessageBox _
        '    "Not Connected.  Please connect before running this script."
        exit sub
    end if

    Dim szCommand, szPrompt, nRow, szLogFileName, nIndex

    ' If this script is run as a login script, there will likely be data
    ' arriving from the remote system.  This is one way of detecting when it's
    ' safe to start sending data. If this script isn't being run as a login
    ' script, then the worst it will do is seemingly pause for one second
    ' before determining what the prompt is.
    ' If you plan on supplying login information by waiting for username and
    ' password prompts within this script, do so right before this do..loop.
    Do
        bCursorMoved = objTab.Screen.WaitForCursor(1)
    Loop until bCursorMoved = False
    ' Once the cursor has stopped moving for about a second, we'll
    ' assume it's safe to start interacting with the remote system.
    
    ' Get the shell prompt so that we can know what to look for when
    ' determining if the command is completed. Won't work if the prompt
    ' is dynamic (e.g. changes according to current working folder, etc)
    nRow = objTab.Screen.CurrentRow
    szPrompt = objTab.screen.Get(nRow, _
                                 0, _
                                 nRow, _
                                 objTab.Screen.CurrentColumn - 1)
    szPrompt = Trim(szPrompt)

    Dim szLogFile
    nIndex = 0
	
	strFilename = g_hostname & "__cmds__" & strMyDate & ".txt"
	'msgBox "strFilename is " & strFilename
	
	g_szLogFile = g_strSpreadSheetPath & strFilename
	'msgBox "g_szLogFile is " & g_szLogFile
	        
	' Store the path for our first log file for later use (see end of this
	' function...
	if g_szFirstLogFilePath = "" then g_szFirstLogFilePath = Replace(g_szLogFile,strFilename,"")
	
	'Check if path exist, and if not, create it.
	Get_BuildFullPath g_szFirstLogFilePath
	Set g_objFile = g_fso.OpenTextFile(g_szLogFile, ForAppending, True)
	
	' Make sure you are receiving command prompt  before sending commands.
	' Send carriage return, wait for prompt to return
	crt.Screen.Send "show clock" & vbcr
	
	' Check for success of command(s) (modify according your
	' scenario), capturing the output of the command for storing in
	' the excel spreadsheet.  Make sure that the success case is the
	' first string in the array.
	vStringsToWaitFor = Array(_
		g_Prompt, _
		"Type help or '?' for a list of available commands.")
	strResults = crt.Screen.ReadString(vStringsToWaitFor)
	strResults = Trim(strResults)
	
	' Get the current line so that we can peel it off of from the
	' results that were captured from the data sent by the remote.
	strCurLine = crt.Screen.Get(_
		crt.Screen.CurrentRow, _
		0, _
		crt.Screen.CurrentRow, _
		crt.Screen.CurrentColumn)
	'msgBox "strResults is " & strResults
	
	strResults = Left(strResults, Len(strResults) - Len(strCurLine))
	
	Set objCell = g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL)
	
	If crt.Screen.MatchIndex = 1 Then bSuccess = True
	
	If bSuccess Then
		objCell.Value  = g_szLogFile
		objCell.ClearComments
		'objCell.AddComment strResults
		'objCell.Comment.Shape.Width = 300
		'objCell.Comment.Shape.Height = 100
	    
		Do
			szCommand = Trim(g_vCommands(nIndex))
			' If the command is empty, then we should be done issuing commands
			' (there's nothing else in our command array g_vCommands)
			if szCommand = "" then Exit Do

			' Send the command text to the remote
			objTab.Screen.Send szCommand & vbcr

			' Wait for the command to be echo'd back to us.
			objTab.Screen.WaitForString vbcr, 1
			objTab.Screen.WaitForString vblf, 1        

			Dim szResult
			' Use the ReadString() method to get the text displayed while the
			' command was runnning.  Note that the ReadString usage shown below
			' is not documented properly in SecureCRT help files included in
			' SecureCRT versions prior to 6.0 Official.  Note also that the
			' ReadString() method captures escape sequences sent from the remote
			' machine as well as displayed text.  As mentioned earlier in comments
			' above, if you want to suppress escape sequences from being captured,
			' set the Screen.IgnoreEscape property = True.
			szResult = objTab.Screen.ReadString(szPrompt)
		   

			' If you don't want the command logged along with the results, comment
			' out the very next line
			'If command was successful write contents to the file.
			g_objFile.Write "****************************************************************************************************" & vbcrlf
			g_objFile.Write "***************************************   START OF COMMAND   ***************************************" & vbcrlf
			g_objFile.Write "****************************************************************************************************" & vbcrlf & vbCrlf
			g_objFile.Write "Device Hostname : " & g_Hostname & vbcrlf
			g_objFile.Write "Command Issued : " & szCommand & vbcrlf & vbCrlf
			g_objFile.Write "****************************************************************************************************" & vbcrlf
			g_objFile.Write "****************************************************************************************************" & vbcrlf
			g_objFile.Write "****************************************************************************************************" & vbcrlf & vbCrlf
			g_objFile.Write szResult & vbCrlf & vbcrlf
			g_objFile.Write "END----------------------------------------------END---------------------------------------------END" & vbcrlf & vbcrlf & vbcrlf & vbcrlf & vbcrlf
			
			' Move on to the next command in our command array g_vCommands
			nIndex = nIndex + 1
		Loop	
	Else
		objCell.Value  = _
			"Failure: Command failed. Matchindex = " & _
			crt.Screen.MatchIndex
		objCell.AddComment szResult
		objCell.Comment.Shape.Width = 200
		objCell.Comment.Shape.Height = 100
	End If
	
 
		' Close the log file once loop is completed.
        g_objFile.Close
		
End Sub

Sub Get_ParseCommands
 
	Set g_objWriteSheet = g_objWriteWkBook.Sheets("Commands")
 
    ' Loop continues until we find an empty row
    Dim strCommmands, nRowIndexCommand, nCommandIndex
	
	'Start looking in row two of the Commands tab in the spreadsheet.  Assumes row 1 is for column headers
	nRowIndexCommand = 2
	nCommandsIndex = 0
	
    Do
        ' If you find an empty value in column #1, exit the loop
				strCommand = Trim(g_objWriteSheet.Cells(nRowIndexCommand, g_COMMAND_COL).Value)
				' Add Command string to g_vCommand array.  Incremend index by one with each pass.
				g_vCommands(nCommandsIndex) = strCommand

				'msgBox "g_vCommands was updated to include: " & strCommand
		  
		If strCommand = "" Then Exit Do
		
		nRowIndexCommand = nRowIndexCommand + 1
		nCommandsIndex = nCommandsIndex + 1
		'msgBox "nRowIndexCommand is " & nRowIndexCommand & "." & vbcrlf & vbcrlf & "starting loop over."
		
	Loop
		
	Set g_objWriteSheet = g_objWriteWkBook.Sheets("Main")
 
End Sub

Sub Get_BuildFullPath(ByVal FullPath)
	'msgBox "FullPath is " & FullPath
	If Not g_fso.FolderExists(FullPath) Then
		'BuildFullPath g_fso.GetParentFolderName(FullPath)
	g_fso.CreateFolder FullPath
	End If
End Sub

Sub Get_Interfaces
	On Error Resume Next
 ' Expect that Get_Hostname subroutine has already been called and populated the global variable for g_hostname
  
    ' Instruct WaitForString and ReadString to ignore escape sequences when
    ' detecting and capturing data received from the remote (this doesn't
    ' affect the way the data is displayed to the screen, only how it is handled
    ' by the WaitForString, WaitForStrings, and ReadString methods associated
    ' with the Screen object.
    objTab.Screen.IgnoreEscape = True
    objTab.Screen.Synchronous = True
	
  'Set Active Worksheet to the Interfaces tab.
    Set g_objWriteSheet = g_objWriteWkBook.Sheets("Interfaces")
	
	' Define numeric values for column identifiers in sheet Interfaces.
	vIF_LINE_COL       	= Asc("A") - 64
	vIF_HOSTNAME_COL	= Asc("B") - 64
	vIF_IPADD_COL		= Asc("C") - 64
	vIF_INTERFACE_COL	= Asc("D") - 64
	vIF_LINESTATUS_COL	= Asc("E") - 64
	vIF_PROTOCOL_COL	= Asc("F") - 64
	vIF_TEST_COL		= Asc("G") - 64
  
	' Get IP Interface Output
	'msgBox "starting Get_Interfaces"
	crt.Screen.Send "show ip interface brief" & vbcr
	'Wait for command prompt to return
	
	' Check for success of command(s) (modify according your
	' scenario), capturing the output of the command for storing in
	' the excel spreadsheet.  Make sure that the success case is the
	' first string in the array.
	vStringsToWaitFor = Array(_
		g_Prompt, _
		"Type help or '?' for a list of available commands.")
	strResults = crt.Screen.ReadString(vStringsToWaitFor)

	
	' Get the current line so that we can peel it off of from the
	' results that were captured from the data sent by the remote.
	'strCurLine = crt.Screen.Get(_
	'	crt.Screen.CurrentRow, _
	'	0, _
	'	crt.Screen.CurrentRow, _
	'	crt.Screen.CurrentColumn)
	'strInterfaces = Left(strResults, Len(strResults) - Len(strCurLine))
	strInterfaces = Trim(strResults)  

  vLines = Split(strInterfaces, vbcrlf)
	
  'Create regex for field values, determine lines in output, cycle through each line and evaluate regex.  If match, add to string.
  set reInterface = New RegExp
  set reIPAdd = New RegExp
  set reLineStatus = New RegExp
  set reProtocol = New RegExp
  
  	'reInterfaceOLD.pattern = "^([Se|Et|Fa|Gi|Te|Tu|Lo|Vl|Vi]([a-zA-Z]+)[0-9].*)\s.*"
	reInterface.pattern = "^(\D+\d+((/\d+)+(\.\d+)?)?)"
	reIPAdd.pattern = "(\d+\.\d+\.\d+\.\d+|unassigned)"			
	reLineStatus.pattern = "[NVRAM|unset|manual|TFTP]\s+(.*)\s+[up|down]"
	reProtocol.pattern = "[NVRAM|unset|manual|TFTP]\s+[a-zA-Z]*\s+(up|down)\s+"
  
  nLineCount = uBound(vLines)
  'msgBox "nLineCount initial : " & nLineCount
  nLineCount = nLineCount + 1
	ReDim vDataArray (nLineCount, 5)

  
	For i = 0 to uBound(vLines)
 
		If g_IF_first = "" Then g_IF_first = 0
		If g_IF_k = "" Then g_IF_k = 1
		'msgBox "g_IF_first is now set to" & g_IF_first
		vRowNumber = g_IF_k
		
		  ' More work to be done here, but if row 1 is blank, populate it with a header.  Increase K to 2.
		  If g_IF_first = 0 Then
			'msgBox "i = " & i
			vDataArray(i,0) = "Line #"
			vDataArray(i,1) = "Hostname"		
			vDataArray(i,2) = "IP Address"
			vDataArray(i,3) = "Interface"
			vDataArray(i,4) = "Line Status"
			vDataArray(i,5) = "Protocol Status"
			'g_objWriteSheet.Cells(g_IF_k, vIF_LINE_COL).Value = "Line #"
			'g_objWriteSheet.Cells(g_IF_k, vIF_HOSTNAME_COL).Value = "Hostname"
			'g_objWriteSheet.Cells(g_IF_k, vIF_IPADD_COL).Value = "IP Address"
			'g_objWriteSheet.Cells(g_IF_k, vIF_INTERFACE_COL).Value = "Interface"
			'g_objWriteSheet.Cells(g_IF_k, vIF_LINESTATUS_COL).Value = "Line Status"
			'g_objWriteSheet.Cells(g_IF_k, vIF_PROTOCOL_COL).Value = "Protocol Status"
			vRowNumber = vRowNumber + 1
			g_IF_first = g_IF_first + 1
		   End If	
		
	  
	  
		If g_IF_first > 0 Then
			
			'msgBox "i = " & i
		 
			For Each strLine In vLines
				'do something with the line variable
		 
			  ' Define new g_IF_k value to increment row number by 1
				'MsgBox "Starting loop.  Index number is now " & nIFIndex
		 
				
				If reInterface.Test(strLine) <> TRUE Then	
				  'MsgBox "Pattern """ & reInterface.Pattern & """ wasn't found within the following text: " & _
				  'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
				 Else
					vDataArray(i,0) = vRowNumber
					vDataArray(i,1) = g_hostname
					
					'g_objWriteSheet.Cells(g_IF_k, vIF_LINE_COL).Value = g_IF_k
					'g_objWriteSheet.Cells(g_IF_k, vIF_HOSTNAME_COL).Value = g_hostname
					Set matches = reInterface.Execute(strLine)
					For each match in matches
						'MsgBox "Pattern matched for Interfaces for iteration " & g_IF_k & " : " & match.SubMatches(0)
						'g_objWriteSheet.Cells(g_IF_k, vIF_INTERFACE_COL).Value = match.SubMatches(0)
						vDataArray(i,3) = match.SubMatches(0)
					Next
					
					If reLineStatus.Test(strLine) <> TRUE Then	
						'MsgBox "Pattern """ & reLineStatus.Pattern & """ wasn't found within the following text: " & _
						'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
					 Else
						Set matches = reLineStatus.Execute(strLine)
						For each match in matches
							'MsgBox "Pattern matched for LineStatus for iteration " & g_IF_k & " : " & match.SubMatches(0)
							'g_objWriteSheet.Cells(g_IF_k, vIF_LINESTATUS_COL).Value = match.SubMatches(0)
							vDataArray(i,4) = match.SubMatches(0)
						Next
					End If
					
					If reIPAdd.Test(strLine) <> TRUE Then	
						'MsgBox "Pattern """ & reIPAdd.Pattern & """ wasn't found within the following text: " & _
						'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
						Else
							Set matches = reIPAdd.Execute(strLine)
								For each match in matches
									'MsgBox "Patter matched for IP Address for iteration " & g_IF_k & " : " & match.SubMatches(0)
									'g_objWriteSheet.Cells(g_IF_k, vIF_IPADD_COL).Value = match.SubMatches(0)
									vDataArray(i,2) = match.SubMatches(0)								
								Next
					End If
					
					If reProtocol.Test(strLine) <> TRUE Then	
						'MsgBox "Pattern """ & reProtocol.Pattern & """ wasn't found within the following text: " & _
						'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
						Else
							Set matches = reProtocol.Execute(strLine)
								For each match in matches
									'MsgBox "Pattern matched for Protocol Status for iteration " & g_IF_k & " : " & match.SubMatches(0)
									'g_objWriteSheet.Cells(g_IF_k, vIF_PROTOCOL_COL).Value = match.SubMatches(0)
									vDataArray(i,5) = match.SubMatches(0)								
								Next
					End If
				vRowNumber = vRowNumber + 1				
				End If
				
				i = i + 1

			Next
		End If
    Next
 
	'Populate array data to excel
	'msgBox "vDataArray uBound is " & uBound(vDataArray, 1)
	For j = 0 to uBound(vDataArray,1)
		if vDataArray(j,0) <> "" Then
			'msgBox "writing data values for row " & j
			g_objWriteSheet.Cells(g_IF_k, vIF_LINE_COL).Value = vDataArray(j,0)
			g_objWriteSheet.Cells(g_IF_k, vIF_HOSTNAME_COL).Value = vDataArray(j,1)			
			g_objWriteSheet.Cells(g_IF_k, vIF_IPADD_COL).Value = vDataArray(j,2)
			g_objWriteSheet.Cells(g_IF_k, vIF_INTERFACE_COL).Value = vDataArray(j,3)				
			g_objWriteSheet.Cells(g_IF_k, vIF_LINESTATUS_COL).Value = vDataArray(j,4)
			g_objWriteSheet.Cells(g_IF_k, vIF_PROTOCOL_COL).Value = vDataArray(j,5)
			g_IF_k = g_IF_k + 1		
		End If
	Next
		
 'Return Worksheet to Main tab for other functions 
  Set g_objWriteSheet = g_objWriteWkBook.Sheets("Main")
 
 'msgBox "Done with Get_Interfaces.  Headed back to main subroutine."
 
End Sub 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub Get_InterfacesBACKUP
 ' Expect that Get_Hostname subroutine has already been called and populated the global variable for g_hostname
  
  
  
    ' Instruct WaitForString and ReadString to ignore escape sequences when
    ' detecting and capturing data received from the remote (this doesn't
    ' affect the way the data is displayed to the screen, only how it is handled
    ' by the WaitForString, WaitForStrings, and ReadString methods associated
    ' with the Screen object.
    objTab.Screen.IgnoreEscape = True
    objTab.Screen.Synchronous = True
	
  'Set Active Worksheet to the Interfaces tab.
    Set g_objWriteSheet = g_objWriteWkBook.Sheets("Interfaces")
	
	' Define numeric values for column identifiers in sheet Interfaces.
	vIF_LINE_COL       	= Asc("A") - 64
	vIF_HOSTNAME_COL	= Asc("B") - 64
	vIF_IPADD_COL		= Asc("C") - 64
	vIF_INTERFACE_COL	= Asc("D") - 64
	vIF_LINESTATUS_COL	= Asc("E") - 64
	vIF_PROTOCOL_COL	= Asc("F") - 64
	vIF_TEST_COL		= Asc("G") - 64
  
	' Get IP Interface Output
	'msgBox "starting Get_Interfaces"
	crt.Screen.Send "show ip interface brief" & vbcr
	'Wait for command prompt to return
	
	' Check for success of command(s) (modify according your
	' scenario), capturing the output of the command for storing in
	' the excel spreadsheet.  Make sure that the success case is the
	' first string in the array.
	vStringsToWaitFor = Array(_
		g_Prompt, _
		"Type help or '?' for a list of available commands.")
	strResults = crt.Screen.ReadString(vStringsToWaitFor)

	
	' Get the current line so that we can peel it off of from the
	' results that were captured from the data sent by the remote.
	'strCurLine = crt.Screen.Get(_
	'	crt.Screen.CurrentRow, _
	'	0, _
	'	crt.Screen.CurrentRow, _
	'	crt.Screen.CurrentColumn)
	'strInterfaces = Left(strResults, Len(strResults) - Len(strCurLine))
	strInterfaces = Trim(strResults)  
	  
  'Create regex for field values, determine lines in output, cycle through each line and evaluate regex.  If match, add to string.
  set reInterface = New RegExp
  set reIPAdd = New RegExp
  set reLineStatus = New RegExp
  set reProtocol = New RegExp
  
  	'reInterfaceOLD.pattern = "^([Se|Et|Fa|Gi|Te|Tu|Lo|Vl|Vi]([a-zA-Z]+)[0-9].*)\s.*"
	reInterface.pattern = "^(\D+\d+((/\d+)+(\.\d+)?)?)"
	reIPAdd.pattern = "(\d+\.\d+\.\d+\.\d+|unassigned)"			
	reLineStatus.pattern = "[NVRAM|unset|manual|TFTP]\s+(.*)\s+[up|down]"
	reProtocol.pattern = "[NVRAM|unset|manual|TFTP]\s+[a-zA-Z]*\s+(up|down)\s+"
  
  

  Dim nRowIndex2
  
  
	If g_IF_k = "" Then g_IF_k = 1
	'msgBox "g_IF_k is now set to" & g_IF_k
	
  ' More work to be done here, but if row 1 is blank, populate it with a header.  Increase K to 2.
  If g_IF_k = 1 Then
	g_objWriteSheet.Cells(g_IF_k, vIF_LINE_COL).Value = "Line #"
	g_objWriteSheet.Cells(g_IF_k, vIF_HOSTNAME_COL).Value = "Hostname"
	g_objWriteSheet.Cells(g_IF_k, vIF_IPADD_COL).Value = "IP Address"
	g_objWriteSheet.Cells(g_IF_k, vIF_INTERFACE_COL).Value = "Interface"
	g_objWriteSheet.Cells(g_IF_k, vIF_LINESTATUS_COL).Value = "Line Status"
	g_objWriteSheet.Cells(g_IF_k, vIF_PROTOCOL_COL).Value = "Protocol Status"
	g_IF_k = g_IF_k + 1
   End If

  ' Determine total number of lines in the strInterfaces.  Don't need header Since we created one.
  vLines = Split(strInterfaces, vbcrlf)
  nIFIndex = 2
	'msgBox "strInterfaces includes the following lines: " & strInterfaces
  'Cycle through each 
  'MsgBox "Starting nIndex at: " & g_IF_k & vbcrlf & vbcrlf & "Total number of lines in the Interface output is " & (UBound(VLines) - 1)
  
  
  'Dim strLineArray(uBound(vLines))
	'strLineArray() = Split(strInterfaces,vbcrlf)
	 
	For Each strLine In vLines
	    'do something with the line variable
 
	  ' Define new g_IF_k value to increment row number by 1
		'MsgBox "Starting loop.  Index number is now " & nIFIndex
 
		
		If reInterface.Test(strLine) <> TRUE Then	
		  'MsgBox "Pattern """ & reInterface.Pattern & """ wasn't found within the following text: " & _
		  'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
		 Else
			g_objWriteSheet.Cells(g_IF_k, vIF_LINE_COL).Value = g_IF_k
			g_objWriteSheet.Cells(g_IF_k, vIF_HOSTNAME_COL).Value = g_hostname
			Set matches = reInterface.Execute(strLine)
			For each match in matches
				'MsgBox "Pattern matched for Interfaces for iteration " & g_IF_k & " : " & match.SubMatches(0)
				g_objWriteSheet.Cells(g_IF_k, vIF_INTERFACE_COL).Value = match.SubMatches(0)
			Next
			
			If reLineStatus.Test(strLine) <> TRUE Then	
				'MsgBox "Pattern """ & reLineStatus.Pattern & """ wasn't found within the following text: " & _
				'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
			 Else
				Set matches = reLineStatus.Execute(strLine)
				For each match in matches
					'MsgBox "Pattern matched for LineStatus for iteration " & g_IF_k & " : " & match.SubMatches(0)
					g_objWriteSheet.Cells(g_IF_k, vIF_LINESTATUS_COL).Value = match.SubMatches(0)
				Next
			End If
			
			If reIPAdd.Test(strLine) <> TRUE Then	
				'MsgBox "Pattern """ & reIPAdd.Pattern & """ wasn't found within the following text: " & _
				'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
				Else
					Set matches = reIPAdd.Execute(strLine)
						For each match in matches
							'MsgBox "Patter matched for IP Address for iteration " & g_IF_k & " : " & match.SubMatches(0)
							g_objWriteSheet.Cells(g_IF_k, vIF_IPADD_COL).Value = match.SubMatches(0)
						Next
			End If
			
			If reProtocol.Test(strLine) <> TRUE Then	
				'MsgBox "Pattern """ & reProtocol.Pattern & """ wasn't found within the following text: " & _
				'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
				Else
					Set matches = reProtocol.Execute(strLine)
						For each match in matches
							'MsgBox "Pattern matched for Protocol Status for iteration " & g_IF_k & " : " & match.SubMatches(0)
							g_objWriteSheet.Cells(g_IF_k, vIF_PROTOCOL_COL).Value = match.SubMatches(0)
						Next
			End If
		
			g_IF_k = g_IF_k + 1
		
		End If
		
		
		nIFIndex = nIFIndex + 1
		'msgBox "Line # " & nIFIndex & " of string strInterfaces processed"
	
	Next
 
 'Return Worksheet to Main tab for other functions
 
  Set g_objWriteSheet = g_objWriteWkBook.Sheets("Main")
 
 'msgBox "Done with Get_Interfaces.  Headed back to main subroutine."
 
End Sub 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub Get_show_ver
 ' Expect that Get_Hostname subroutine has already been called and populated the global variable for g_hostname
 
    ' Instruct WaitForString and ReadString to ignore escape sequences when
    ' detecting and capturing data received from the remote (this doesn't
    ' affect the way the data is displayed to the screen, only how it is handled
    ' by the WaitForString, WaitForStrings, and ReadString methods associated
    ' with the Screen object.
    objTab.Screen.IgnoreEscape = True
    objTab.Screen.Synchronous = True
	
  'Set Active Worksheet to the Main tab.
    Set g_objWriteSheet = g_objWriteWkBook.Sheets("Main")
	
	' Define numeric values for column identifiers in sheet Interfaces.
	vSERIAL_IF_COL				= Asc("J") - 64
	vSERIAL_ACTIVE_COL			= Asc("K") - 64
	vETHERNET_IF_COL			= Asc("L") - 64
	vETHERNET_ACTIVE_COL		= Asc("M") - 64
	vFASTETHERNET_IF_COL		= Asc("N") - 64
	vFASTETHERNET_ACTIVE_COL	= Asc("O") - 64
	vGIGETHERNET_IF_COL			= Asc("P") - 64
	vGIGETHERNET_ACTIVE_COL		= Asc("Q") - 64
	vTENGIGETHERNET_IF_COL		= Asc("R") - 64
	vTENGIGETHERNET_ACTIVE_COL	= Asc("S") - 64
	vFLASH_COL					= Asc("T") - 64
	vMEMORY_COL 				= Asc("U") - 64
	vMODEL_COL					= Asc("V") - 64
	vSTACK_COL					= Asc("W") - 64
	vIOS_COL					= Asc("X") - 64
	vSERIALNUM_COL				= Asc("Y") - 64
	vUPTIME_COL					= 27
  
	' Get IP Interface Output
	'msgBox "starting Get_Interfaces"
	crt.Screen.Send "show version" & vbcr
	'Wait for command prompt to return
	
	' Check for success of command(s) (modify according your
	' scenario), capturing the output of the command for storing in
	' the excel spreadsheet.  Make sure that the success case is the
	' first string in the array.
	'msgBox "Inside show ver subrouting " & vbcrlf & "g_Prompt is : " & g_Prompt
	vStringsToWaitFor = Array(_
		g_Prompt, _
		"Type help or '?' for a list of available commands.")
	strResults = crt.Screen.ReadString(vStringsToWaitFor)
	' Get the current line so that we can peel it off of from the
	' results that were captured from the data sent by the remote.
	strCurLine = crt.Screen.Get(_
		crt.Screen.CurrentRow, _
		0, _
		crt.Screen.CurrentRow, _
		crt.Screen.CurrentColumn)
	strVersion = Left(strResults, Len(strResults) - Len(strCurLine))
	strVersion = Trim(strVersion)  
	  
	  
	  
  'Create regex for field values, determine lines in output, cycle through each line and evaluate regex.  If match, add to string.
	set reSERIAL_IF = New RegExp
	set reETHERNET_IF = New RegExp
	set reFASTETHERNET_IF = New RegExp
	set reGIGETHERNET_IF = New RegExp
	set reTENGIGETHERNET_IF = New RegExp
	set reFLASH = New RegExp
	set reMEMORY = New RegExp
	set reMODEL = New RegExp
	set reIOS = New RegExp
	set reSERIALNUM = New RegExp
	set reSTACKTEST = New RegExp
	set reCOUNT = New RegExp
	set reUP = New RegExp
	set reUPTIME = New RegExp
	set reDATETIME = New RegExp
  
	reETHERNET_IF.pattern = "([0-9]+) Ethernet [Ii]nterfaces"
	reFASTETHERNET_IF.pattern = "([0-9]+) Fast.*Ethernet [iI]nterfaces"									' Tested and works
	reGIGETHERNET_IF.pattern = "([0-9]+) [gG]igabit.*[eE]thernet.*[iI]nterfaces"						' Tested and works
	reTENGIGETHERNET_IF.pattern = "([0-9]+) [tT]en [gG]igabit [eE]thernet [iI]nterfaces"				' Tested and works
	reSERIAL_IF.pattern = "([0-9]+).*[sS]erial.*[iI]nterfaces"
	reFLASH.pattern = "([0-9]+)[kK] bytes of .*Flash.*"
	reIOS.pattern = "System image file is.*:(.*)\"""
	'System image file is "flash:c3750-ipservicesk9-mz.122-55.SE3.bin"
	reMODEL.pattern = "[cC]isco\s(.[^\s]+)\s.*[mM]emory"													' Tested and works
	reSERIALNUM.pattern = "Processor board ID\s(.*)"														' Tested and works
	reMEMORY.pattern = "[cC]isco.*with\s(.*[^\s])\sbytes of [mM]emory"										' Tested and works
	reSTACKTEST.pattern = "^(WS-C3750).*"
	reCOUNT.pattern = "^(\*[0-9]| [0-9]).*"
 	reUP.pattern = ".*(up.*up).*"
	reUPTIME.pattern = ".*uptime is (.*)$"
	reDATETIME.pattern = "uptime is (.*[0-9]+:[0-9]+)$"   


  ' Determine total number of lines in the strInterfaces.  Don't need header Since we created one.
  vLines = Split(strVersion, vbcrlf)
  
  'Cycle through each 
  'MsgBox "Starting nIndex at: " & g_IF_k & vbcrlf & vbcrlf & "Total number of lines in the Interface output is " & (UBound(VLines) - 1)
  
	 
	For Each strLine In vLines
	    'do something with the line variable

		If reETHERNET_IF.Test(strLine) <> TRUE Then	
			'MsgBox "Pattern """ & reLineStatus.Pattern & """ wasn't found within the following text: " & _
			'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
		 Else
			pattern = "^[Ee]thernet.*[0-9]+(/[0-9]+)+ .*up.*up"
			crt.Screen.Send "show ip int bri | inc " & pattern & vbcr
			'Wait for command prompt to return
			strResultsCount = crt.Screen.ReadString(g_Prompt)
			trim(strResultsCount)
			vCount = split(strResultsCount, vbcrlf)
			strCount = 0
			reUP.pattern = pattern	
			For each count in vCount
				'msgBox "Current line is " & vbcrlf & count & vbcrlf & "Counter set to " & strCount
				If reUP.test(count) = TRUE Then
					
					strCount = strCount + 1
					'msgBox "pattern match reCOUNT " & reCOUNT.pattern & vbcrlf & "Incrementing counter:  " & strCount
				End If
			Next
			g_objWriteSheet.Cells(nRowIndex, vFASTETHERNET_ACTIVE_COL).Value = strCount		 			
			Set matches = reETHERNET_IF.Execute(strLine)
			For each match in matches
				'MsgBox "Pattern matched for LineStatus for iteration " & g_IF_k & " : " & match.SubMatches(0)
				g_objWriteSheet.Cells(nRowIndex, vETHERNET_IF_COL).Value = match.SubMatches(0)
			Next
		End If
		
		If reFASTETHERNET_IF.Test(strLine) <> TRUE Then	
 
		 Else
		 	pattern = "^[Ff]ast.*[0-9]+(/[0-9]+)+ .*up.*up"
			crt.Screen.Send "show ip int bri | inc " & pattern & vbcr
			'Wait for command prompt to return
			strResultsCount = crt.Screen.ReadString(g_Prompt)
			trim(strResultsCount)
			vCount = split(strResultsCount, vbcrlf)
			strCount = 0
			reUP.pattern = pattern
			For each count in vCount
				'msgBox "Current line is " & vbcrlf & count & vbcrlf & "Counter set to " & strCount
				If reUP.test(count) = TRUE Then
					
					strCount = strCount + 1
					'msgBox "pattern match reCOUNT " & reCOUNT.pattern & vbcrlf & "Incrementing counter:  " & strCount
				End If
			Next
			g_objWriteSheet.Cells(nRowIndex, vFASTETHERNET_ACTIVE_COL).Value = strCount
			
			Set matches = reFASTETHERNET_IF.Execute(strLine)
			For each match in matches
				g_objWriteSheet.Cells(nRowIndex, vFASTETHERNET_IF_COL).Value = match.SubMatches(0)
			Next
		End If
		
		If reGIGETHERNET_IF.Test(strLine) <> TRUE Then	
 
		 Else
			pattern = "^[Gg]ig.*[0-9]+(/[0-9]+)+ .*up.*up"
			crt.Screen.Send "show ip int bri | inc " & pattern & vbcr
			'Wait for command prompt to return
			strResultsCount = crt.Screen.ReadString(g_Prompt)
			trim(strResultsCount)
			vCount = split(strResultsCount, vbcrlf)
			strCount = 0
			reUP.pattern = pattern
			For each count in vCount
				'msgBox "Current line is " & vbcrlf & count & vbcrlf & "Counter set to " & strCount
				If reUP.test(count) = TRUE Then
					
					strCount = strCount + 1
					'msgBox "pattern match reUP " & reUP.pattern & "line is " & count & vbcrlf & "Incrementing counter:  " & strCount
				End If
			Next
			g_objWriteSheet.Cells(nRowIndex, vGIGETHERNET_ACTIVE_COL).Value = strCount	
					 
			Set matches = reGIGETHERNET_IF.Execute(strLine)
			For each match in matches
				g_objWriteSheet.Cells(nRowIndex, vGIGETHERNET_IF_COL).Value = match.SubMatches(0)
			Next
		End If
		
		If reTENGIGETHERNET_IF.Test(strLine) <> TRUE Then	
 
		 Else
			pattern = "^[Tt]e.*[0-9]+(/[0-9]+)+ .*up.*up"
			crt.Screen.Send "show ip int bri | inc " & pattern & vbcr
			'Wait for command prompt to return
			strResultsCount = crt.Screen.ReadString(g_Prompt)
			trim(strResultsCount)
			vCount = split(strResultsCount, vbcrlf)
			strCount = 0
			reUP.pattern = pattern
			For each count in vCount
				'msgBox "Current line is " & vbcrlf & count & vbcrlf & "Counter set to " & strCount
				If reUP.test(count) = TRUE Then
					
					strCount = strCount + 1
					'msgBox "pattern match reCOUNT " & reCOUNT.pattern & vbcrlf & "Incrementing counter:  " & strCount
				End If
			Next		 
		 	g_objWriteSheet.Cells(nRowIndex, vTENGIGETHERNET_ACTIVE_COL).Value = strCount
					 
			Set matches = reTENGIGETHERNET_IF.Execute(strLine)
			For each match in matches
				g_objWriteSheet.Cells(nRowIndex, vTENGIGETHERNET_IF_COL).Value = match.SubMatches(0)
			Next
		End If

		If reSERIAL_IF.Test(strLine) <> TRUE Then	

		Else
			pattern = "^[Ss]er.*[0-9]+(/[0-9]+)+ .*up.*up"
			crt.Screen.Send "show ip int bri | inc " & pattern & vbcr
			'Wait for command prompt to return
			strResultsCount = crt.Screen.ReadString(g_Prompt)
			trim(strResultsCount)
			vCount = split(strResultsCount, vbcrlf)
			strCount = 0
			reUP.pattern = pattern
			For each count in vCount
				'msgBox "Current line is " & vbcrlf & count & vbcrlf & "Counter set to " & strCount
				If reUP.test(count) = TRUE Then
					
					strCount = strCount + 1
					'msgBox "pattern match reCOUNT " & reCOUNT.pattern & vbcrlf & "Incrementing counter:  " & strCount
				End If
			Next		 		

			g_objWriteSheet.Cells(nRowIndex, vSERIAL_ACTIVE_COL).Value = strCount
					
			Set matches = reSERIAL_IF.Execute(strLine)
			For each match in matches
				g_objWriteSheet.Cells(nRowIndex, vSERIAL_IF_COL).Value = match.SubMatches(0)
			Next
		End If
		
		If reFLASH.Test(strLine) <> TRUE Then	
 
		 Else
			Set matches = reFLASH.Execute(strLine)
			For each match in matches
				g_objWriteSheet.Cells(nRowIndex, vFLASH_COL).Value = match.SubMatches(0)
			Next
		End If
		
		If reMEMORY.Test(strLine) <> TRUE Then	
 
		 Else
			Set matches = reMEMORY.Execute(strLine)
			For each match in matches
				g_objWriteSheet.Cells(nRowIndex, vMEMORY_COL).Value = match.SubMatches(0)
			Next
		End If
		
		If reMODEL.Test(strLine) <> TRUE Then	
 
		 Else
			Set matches = reMODEL.Execute(strLine)
			For each match in matches
				g_objWriteSheet.Cells(nRowIndex, vMODEL_COL).Value = match.SubMatches(0)
				' Look for stacked switches.  If stack switch is matched, count number in stack.
				If reSTACKTEST.Test(match.SubMatches(0)) = TRUE Then
					crt.Screen.Send "sh switch" & vbcr
					'Wait for command prompt to return
					strResultsCount = crt.Screen.ReadString(g_Prompt)
					
					vCount = split(strResultsCount, vbcrlf)
					strCount = 0
					For each count in vCount
						'msgBox "Current line is " & vbcrlf & count & vbcrlf & "Counter set to " & strCount
						If reCOUNT.test(count) = TRUE Then
							
							strCount = strCount + 1
							'msgBox "pattern match reCOUNT " & reCOUNT.pattern & vbcrlf & "Incrementing counter:  " & strCount
						End If
					Next
						g_objWriteSheet.Cells(nRowIndex, vSTACK_COL).Value = strCount
				End If
			Next
		End If	

		If reIOS.Test(strLine) <> TRUE Then	
 
		 Else
			Set matches = reIOS.Execute(strLine)
			For each match in matches
				g_objWriteSheet.Cells(nRowIndex, vIOS_COL).Value = match.SubMatches(0)
			Next
		End If				
		
		If reSERIALNUM.Test(strLine) <> TRUE Then	
 
		 Else
			Set matches = reSERIALNUM.Execute(strLine)
			For each match in matches
				g_objWriteSheet.Cells(nRowIndex, vSERIALNUM_COL).Value = match.SubMatches(0)
			Next
		End If						

		If reUPTIME.Test(strLine) <> TRUE Then	
 
		 Else
			Set matches = reUPTIME.Execute(strLine)
			For each match in matches
				strMatch = Replace(match, " years, ","y")
				strMatch = Replace(strMatch, " year, ","y")
				strMatch = Replace(strMatch, " weeks, ","w")
				strMatch = Replace(strMatch, " week, ","w")
				strMatch = Replace(strMatch, " days, ","d_")
				strMatch = Replace(strMatch, " day, ","d_")
				strMatch = Replace(strMatch, " hours, ",":")
				strMatch = Replace(strMatch, " hour, ",":")
				strMatch = Replace(strMatch, " minutes","")
				strMatch = Replace(strMatch, " minute","")
				Set matches2 = reDATETIME.Execute(strMatch)
				For each match2 in matches2
					g_objWriteSheet.Cells(nRowIndex, vUPTIME_COL).Value = match2.SubMatches(0)
				Next
			Next
		End If			
  Next

 

End Sub 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub Get_show_proc_cpu
	
	' Define numeric values for column identifiers in sheet Interfaces.
	vCPU_ONEMIN_COL			= 28
	vCPU_FIVEMIN_COL		= 29
	strCommand	= "show proc cpu | inc CPU utilization"
	crt.Screen.Send strCommand & vbcr
	'Wait for command prompt to return
	
	' Check for success of command(s) (modify according your
	' scenario), capturing the output of the command for storing in
	' the excel spreadsheet.  Make sure that the success case is the
	' first string in the array.
	vWait = Array(_
		g_Prompt, _
		"Type help or '?' for a list of available commands.")
	strResults = crt.Screen.ReadString(vWait)
	strResults = Trim(strResults)
	vLines = Split(strResults, vbcrlf)
	
	set reONEMINCPU = New RegExp
	set reFIVEMINCPU = New RegExp
  
	reONEMINCPU.pattern = "^CPU utilization.*one minute: (.*);.*"
	reFIVEMINCPU.pattern = "^CPU utilization.*five minutes: (.*)$"
	
	
	Dim strCPU
	
	For each line in vLines
		if line <> (strCommand) Then
			If line <> "" Then
				If strCPU = "" Then
					strCPU = line
				Else
					strCPU = strCPU & vbcrlf & line
				End If
			End If
		End If
	Next
	
	'msgBox "strCPU = " & strCPU
	
	If reONEMINCPU.Test(strCPU) <> TRUE Then	
 
	 Else
		Set matches = reONEMINCPU.Execute(strCPU)
		For each match in matches
			g_objWriteSheet.Cells(nRowIndex, vCPU_ONEMIN_COL).Value = match.SubMatches(0)
		Next
	End If	
	
	If reFIVEMINCPU.Test(strCPU) <> TRUE Then	
 
	 Else
		Set matches = reFIVEMINCPU.Execute(strCPU)
		For each match in matches
			g_objWriteSheet.Cells(nRowIndex, vCPU_FIVEMIN_COL).Value = match.SubMatches(0)
		Next
	End If	

End Sub 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~



Sub Get_show_snmp
 ' Expect that Get_Hostname subroutine has already been called and populated the global variable for g_hostname
 
    ' Instruct WaitForString and ReadString to ignore escape sequences when
    ' detecting and capturing data received from the remote (this doesn't
    ' affect the way the data is displayed to the screen, only how it is handled
    ' by the WaitForString, WaitForStrings, and ReadString methods associated
    ' with the Screen object.
    objTab.Screen.IgnoreEscape = True
    objTab.Screen.Synchronous = True
	
	
	' Define numeric values for column identifiers in sheet Interfaces.
	vSNMP_LOC_COL			= Asc("Z") - 64
	strCommand	= "show snmp location"
	crt.Screen.Send strCommand & vbcr
	'Wait for command prompt to return
	
	' Check for success of command(s) (modify according your
	' scenario), capturing the output of the command for storing in
	' the excel spreadsheet.  Make sure that the success case is the
	' first string in the array.
	vStringsToWaitFor = Array(_
		g_Prompt, _
		"Type help or '?' for a list of available commands.")
	strResults = crt.Screen.ReadString(vStringsToWaitFor)
	strResults = Trim(strResults)
	vLines = Split(strResults, vbcrlf)
	
	Dim strSNMP
	
	For each line in vLines
		if line <> (strCommand) Then
			If line <> "" Then
				strSNMP = line
			End If
		End If
	Next
	
	g_objWriteSheet.Cells(nRowIndex, vSNMP_LOC_COL).Value = strSNMP

End Sub 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Sub Get_SFPs		' Get SFP count and populate Excel spreadsheet
 ' Expect that Get_Hostname subroutine has already been called and populated the global variable for g_hostname
 
    ' Instruct WaitForString and ReadString to ignore escape sequences when
    ' detecting and capturing data received from the remote (this doesn't
    ' affect the way the data is displayed to the screen, only how it is handled
    ' by the WaitForString, WaitForStrings, and ReadString methods associated
    ' with the Screen object.
    objTab.Screen.IgnoreEscape = True
    objTab.Screen.Synchronous = True
	Set g_objWriteSheet = g_objWriteWkBook.Sheets("Main")
	
	' Define numeric values for column identifiers in sheet Interfaces.
    v100M_FX_SFP_COL			= 30
    v100M_LX_SFP_COL			= 31
    v100M_TX_SFP_COL			= 32
	v1G_SH_SFP_COL			    = 33
	v1G_LH_SFP_COL			    = 34
	v1G_TX_SFP_COL			    = 35
	v1G_CX_SFP_COL				= 36	
	v10G_SH_SFP_COL			    = 37	
	v10G_LH_SFP_COL			    = 38	
	v10G_X2_SFP_COL			    = 39
	vOther_SFP_COL			    = 40
	strCommand	= "sh int | inc media type.*SFP|10G[Bb]ase| 1000Base"
	crt.Screen.Send strCommand & vbcr
	'Wait for command prompt to return
	
	' Check for success of command(s) (modify according your
	' scenario), capturing the output of the command for storing in
	' the excel spreadsheet.  Make sure that the success case is the
	' first string in the array.
	vStringsToWaitFor = Array(_
		g_Prompt, _
		"Type help or '?' for a list of available commands.")
	strResults = crt.Screen.ReadString(vStringsToWaitFor)
	strResults = Trim(strResults)
	vLines = Split(strResults, vbcrlf)
	
	set reSFP = New RegExp
	set reSFP2 = New RegExp
	set re10G = New RegExp
	
	reSFP.pattern = "media type is (.*) SFP.*"
	reSFP2.pattern = "media type is SFP-(.*)"	
	re10G.pattern = "media type is (.*[Bb]ase.*)$"
	
	
	' Declare variables populate with spreadsheet data (should be "")
 	str100MFX = g_objWriteSheet.Cells(nRowIndex, v100M_FX_SFP_COL).Value
	str100MLX = g_objWriteSheet.Cells(nRowIndex, v100M_LX_SFP_COL).Value	
	str100MTX = g_objWriteSheet.Cells(nRowIndex, v100M_TX_SFP_COL).Value   
	str1GSH = g_objWriteSheet.Cells(nRowIndex, v1G_SH_SFP_COL).Value
	str1GLH = g_objWriteSheet.Cells(nRowIndex, v1G_LH_SFP_COL).Value	
	str1GTX = g_objWriteSheet.Cells(nRowIndex, v1G_TX_SFP_COL).Value	
	str10GSH = g_objWriteSheet.Cells(nRowIndex, v10G_SH_SFP_COL).Value	
	str10GLH = g_objWriteSheet.Cells(nRowIndex, v10G_LH_SFP_COL).Value	
	str10GX2 = g_objWriteSheet.Cells(nRowIndex, v10G_X2_SFP_COL).Value	
	str1GCX = g_objWriteSheet.Cells(nRowIndex, v1G_CX_SFP_COL).Value	
	strOther = g_objWriteSheet.Cells(nRowIndex, vOther_SFP_COL).Value		
	
	For each line in vLines
		If reSFP.Test(line) = TRUE Then
			Set matches = reSFP.Execute(line)
			For each match in matches
				Select Case Trim(UCase(match.submatches(0)))
					Case "1000BASESX"
						'Multimode 1G fiber
						If str1GSH = "" Then
							str1GSH = 1
						Else
							str1GSH = str1GSH + 1
						End If
					Case "1000BASELX", "1000BASELH"
						'SingleMode 1G fiber
						If str1GLH = "" Then
							str1GLH = 1
						Else
							str1GLH = str1GLH + 1
						End If
					Case "10/100/1000BASETX" 
						'Copper 1G fiber
						If str1GTX = "" Then
							str1GTX = 1
						Else
							str1GTX = str1GTX + 1
						End If
 					Case "100BASEFX" 
						'100 Meg MultiMode Fiber SFP
						If str100MFX = "" Then
							str100MFX = 1
						Else
							str100MFX = str100MFX + 1
						End If
  					Case "100BASELX" 
						'100 Meg SingleMode Fiber SFP
						If str100MLX = "" Then
							str100MLX = 1
						Else
							str100MLX = str100MLX + 1
						End If	                       
  					Case "100BASETX" 
						'100 Meg SingleMode Fiber SFP
						If str100MTX = "" Then
							str100MTX = 1
						Else
							str100MTX = str100MTX + 1
						End If	 
  					Case "1000BaseCX" 
						'1G Twinax SFP Cable
						If str1GCX = "" Then
							str1GCX = 1
						Else
							str1GCX = str1GCX + 1
						End If							
					Case Else
						If strOther = "" Then
							strOther = 1
						Else
							strOther = strOther + 1
						End If						
				End Select
			Next
		ElseIf reSFP2.Test(line) = TRUE Then
			Set matches = reSFP2.Execute(line)
			For each match in matches
				Select Case UCase(match.submatches(0))
					Case "10GBASE-SR"
						'Multimode fiber 10G
						If str10GSH = "" Then
							str10GSH = 1
						Else
							str10GSH = str10GSH + 1
						End If
					Case "10GBASE-LR" 
						'SingleMode fiber 10G
						If str10GLH = "" Then
							str10GLH = 1
						Else
							str10GLH = str10GLH + 1
						End If
					Case "10GBASE-LX4" 
						'X2 10G SFP
						If str10GX2 = "" Then
							str10GX2 = 1
						Else
							str10GX2 = str10GX2 + 1
						End If	
					Case "1000BASESX"
						'Multimode 1G fiber
						If str1GSH = "" Then
							str1GSH = 1
						Else
							str1GSH = str1GSH + 1
						End If
					Case "1000BASELX", "1000BASELH"
						'SingleMode 1G fiber
						If str1GLH = "" Then
							str1GLH = 1
						Else
							str1GLH = str1GLH + 1
						End If
					Case "10/100/1000BASETX" 
						'Copper 1G fiber
						If str1GTX = "" Then
							str1GTX = 1
						Else
							str1GTX = str1GTX + 1
						End If		
  					Case "1000BaseCX" 
						'1G Twinax SFP Cable
						If str1GCX = "" Then
							str1GCX = 1
						Else
							str1GCX = str1GCX + 1
						End If							
					Case Else
						If strOther = "" Then
							strOther = 1
						Else
							strOther = strOther + 1
						End If
				End Select		
			Next
		ElseIf re10G.Test(line) = TRUE Then
			Set matches = re10G.Execute(line)
			For each match in matches
				Select Case UCase(match.submatches(0))
					Case "10GBASE-SR"
						'Multimode fiber 10G
						If str10GSH = "" Then
							str10GSH = 1
						Else
							str10GSH = str10GSH + 1
						End If
					Case "10GBASE-LR" 
						'SingleMode fiber 10G
						If str10GLH = "" Then
							str10GLH = 1
						Else
							str10GLH = str10GLH + 1
						End If
					Case "10GBASE-LX4"
						'X2 10G SFP
						If str10GX2 = "" Then
							str10GX2 = 1
						Else
							str10GX2 = str10GX2 + 1
						End If	
					Case "1000BASESX"
						'Multimode 1G fiber
						If str1GSH = "" Then
							str1GSH = 1
						Else
							str1GSH = str1GSH + 1
						End If
					Case "1000BASELX", "1000BASELH"
						'SingleMode 1G fiber
						If str1GLH = "" Then
							str1GLH = 1
						Else
							str1GLH = str1GLH + 1
						End If
					Case "10/100/1000BASETX" 
						'Copper 1G fiber
						If str1GTX = "" Then
							str1GTX = 1
						Else
							str1GTX = str1GTX + 1
						End If	
  					Case "1000BaseCX" 
						'1G Twinax SFP Cable
						If str1GCX = "" Then
							str1GCX = 1
						Else
							str1GCX = str1GCX + 1
						End If							
					Case Else
						If strOther = "" Then
							strOther = 1
						Else
							strOther = strOther + 1
						End If
				End Select
			Next
		End If		
	Next
	
	'msgBox "Total SFP Count" & vbcrlf & "1G SH - " & str1GSH & vbcrlf & _
	'		"1G LH - " & str1GLH & vbcrlf & "1G TX" & str1GTX & vbcrlf & _
	'		"10G SH - " & str10GSH & vbcrlf & "10G LH - " & str10GLH & vbcrlf &_ 
	'		"10G X2 - " & str10GX2
	g_objWriteSheet.Cells(nRowIndex, v100M_FX_SFP_COL).Value = str100MFX
	g_objWriteSheet.Cells(nRowIndex, v100M_LX_SFP_COL).Value = str100MLX
	g_objWriteSheet.Cells(nRowIndex, v100M_TX_SFP_COL).Value = str100MTX
	g_objWriteSheet.Cells(nRowIndex, v1G_SH_SFP_COL).Value = str1GSH
	g_objWriteSheet.Cells(nRowIndex, v1G_LH_SFP_COL).Value = str1GLH
	g_objWriteSheet.Cells(nRowIndex, v1G_TX_SFP_COL).Value = str1GTX
	g_objWriteSheet.Cells(nRowIndex, v1G_CX_SFP_COL).Value = str1GCX	
	g_objWriteSheet.Cells(nRowIndex, v10G_SH_SFP_COL).Value = str10GSH
	g_objWriteSheet.Cells(nRowIndex, v10G_LH_SFP_COL).Value = str10GLH
	g_objWriteSheet.Cells(nRowIndex, v10G_X2_SFP_COL).Value = str10GX2
	
 
End Sub 
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Sub Get_LoginData
 
	Set g_objWriteSheet = g_objWriteWkBook.Sheets("Main")
 
    ' Loop continues until we find an empty row
    Dim strCommmands, nRowIndexCommand, nCommandIndex, strUsername, strPassword, strEnablePass
	
	'Start looking in row two of the Commands tab in the spreadsheet.  Assumes row 1 is for column headers
	nRowIndexCommand = nRowIndex
	nLoginDataIndex = 0
	
	strUsername = Trim(g_objWriteSheet.Cells(nUsernameIndex, nSettingsColIndex).Value)
	strPassword = Trim(g_objWriteSheet.Cells(nPasswordIndex, nSettingsColIndex).Value)
	strEnablePass = Trim(g_objWriteSheet.Cells(nEnableIndex, nSettingsColIndex).Value)
	
	If strUsername <> "" Then g_Username = strUsername
	If strPassword <> "" Then 
		g_Password = strPassword
		g_objWriteSheet.Cells(nPasswordIndex, nSettingsColIndex).Value = ""
	End If
	If strEnablePass <> "" Then 
		g_EnablePass = strEnablePass
		g_objWriteSheet.Cells(nEnableIndex, nSettingsColIndex).Value = ""
	End If
	
	If (g_Username = "" or g_Password = "") Then
		Do
		
			' If you find an empty value in column #1, exit the loop
			'msgBox "Cell selection will be read from " & nRowIndexCommand & ", " & g_IP_COL
			strIP = Trim(g_objWriteSheet.Cells(nRowIndexCommand, g_IP_COL).Value)
			If strIP = "" Then Exit Do
			
			strActive = Trim(g_objWriteSheet.Cells(nRowIndexCommand, g_ACT_COL).Value)
			strProtocol = Trim(g_objWriteSheet.Cells(nRowIndexCommand, g_PROTO_COL).Value)
		
			If LCase(strActive) = "yes" Then
				If (strProtocol = "SSH1" OR strProtocol = "SSH2" or strProtocol = "Telnet") Then
					strUsername = Trim(g_objWriteSheet.Cells(nRowIndexCommand, g_USER_COL).Value)
					strPassword = Trim(g_objWriteSheet.Cells(nRowIndexCommand, g_PASS_COL).Value)	
					
					If (strUsername = "" or strPassword = "") Then 
						
					
						If g_Username = "" Then 
							g_Username = crt.Dialog.Prompt("Username not specified on one line.  Please specify global username variable.", "Specify Username", "Username", False)
						End If
						
						If g_Password = "" Then 
							g_Password = crt.Dialog.Prompt("Password not specified on one line.  Please specify global password variable." &_
										vbcrlf & "Default enable password will be set to the same if none specified globally.", "Specify Password", "", True)
							If g_EnablePass = "" Then g_EnablePass = g_Password
						End If
					Else	
						'msgBox "Device : " & strIP & vbcrlf & "Protocol : " & strProtocol & vbcrlf &_
						'		"Username : " & strUsername & vbcrlf & "Password : " & strPassword
					 End If
				End If

			 End If
			 
			'Increment index by one with each pass.
			nRowIndexCommand = nRowIndexCommand + 1
			nLoginIndex = nLoginIndex + 1
			'msgBox "nRowIndexCommand is " & nRowIndexCommand & "." & vbcrlf & vbcrlf & "starting loop over."
	 
		Loop
	Else
	
	End If
	
		
	Set g_objWriteSheet = g_objWriteWkBook.Sheets("Main")
	'MsgBox "Username is set to: " & g_Username & vbcrlf & "Password is set to: " & g_Password & vbcrlf & "Enable pass set to : " & g_EnablePass
	
 
End Sub


Sub Get_Enable
 
	'	NEW			NEW				NEW				NEW				NEW
	'  This is all new lines of code to add enable aware function.
	' Add regular expression to look for > prompt.
 
	set reEnable = New RegExp
	reEnable.Pattern = g_hostname & ".*(\>)"
 
	If reEnable.Test(g_Prompt) = TRUE Then
		'msgBox "Regex matched " & reEnable.Pattern & vbcrlf & vbcrlf & "Prompt : " & g_prompt
		'If RegEx has a hit for a non-priveledge prompt, check if global password exists.
		strEnablePass = Trim(g_objWriteSheet.Cells(nRowIndex, g_PASS_COL).Value)
		If g_EnablePass = "" Then
			If strEnablePass <> "" Then
				g_EnablePass = g_objWriteSheet.Cells(nRowIndex, g_PASS_COL).Value
				strEnablePass = g_objWriteSheet.Cells(nRowIndex, g_PASS_COL).Value
			Else
				g_EnablePass = crt.Dialog.Prompt("Please enter the enable password:", "Enable Pass?", "", True)
				If g_EnablePass = "" Then 
					g_EnablePass = crt.Dialog.Prompt("Please enter enable password - last chance:", "Enable Pass?", "", True)
					If g_EnablePass = "" Then 
						'msgBox "Enable password required to connect to " & strIP & vbcrlf & vbcrlf & "No Password entered.  Moving to next item."
						g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = "Enable password not entered"
						Exit Sub
					End If
				End If

			End If
		Else
			strEnablePass = g_EnablePass
		End If
		
		crt.Screen.Send "enable" & vbcr
		
		crt.Screen.WaitforString("assword:")
		'msgBox "strEnablePass is currently set to " & strEnablePass & vbcrlf & "g_EnablePass is " & g_EnablePass
		crt.Screen.Send strEnablePass & vbCrlf
		crt.Screen.WaitForString vbcrlf, 2
		
		'If crt.Screen.WaitForString g_hostname, 2) = True Then
 
		
			Do
				' Simulate pressing "Enter" so the prompt appears again...
				crt.Screen.Send vbcr & "show clock | inc 2016" & vbCr
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
				strPrompt = crt.screen.Get(nRow, _
										   0, _
										   nRow, _
										   crt.Screen.CurrentColumn - 1)
				' Loop until we actually see a line of text appear:
				strPrompt = Trim(strPrompt)
				If strPrompt <> "" Then Exit Do
			Loop
			'msgBox "Matched hostname"
			
			'msgBox "strPrompt is " & strPrompt & vbcrlf & "g_hostname is: " & g_hostname
			If strPrompt = g_hostname & "#" Then
				'msgBox "Updating g_Prompt to " & strPrompt
				g_Prompt = strPrompt
				Exit Sub
			Else
				g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = "Enable password not correct."	
			End If
		'End If
	Else
		Exit Sub
	End If
	
	'	NEW			NEW				NEW				NEW				NEW
					
End Sub


Sub Get_ShowCDPneigh
 
    ' Instruct WaitForString and ReadString to ignore escape sequences when
    ' detecting and capturing data received from the remote (this doesn't
    ' affect the way the data is displayed to the screen, only how it is handled
    ' by the WaitForString, WaitForStrings, and ReadString methods associated
    ' with the Screen object.
    objTab.Screen.IgnoreEscape = True
    objTab.Screen.Synchronous = True
	
	
	' Define numeric values for column identifiers in sheet Interfaces.
   	vHOSTNAME_COL			= Asc("A") - 64
	vCDPNEIGHID_COL			= Asc("B") - 64
	vIPADD_COL				= ASC("C") - 64
	vMGMTIP_COL				= Asc("D") - 64
	vPLATFORM_COL			= Asc("E") - 64
	vIFLOCAL_COL			= Asc("F") - 64
	vIFREMOTE_COL			= Asc("G") - 64
	vREMOTEIOS_COL			= Asc("H") - 64
	vPOWER_COL				= Asc("I") - 64

	' Get IP Interface Output
	'msgBox "starting Get_Interfaces"
	crt.Screen.Send "show cdp neighbor detail" & vbcr
	'Wait for command prompt to return
	
	' Check for success of command(s) (modify according your
	' scenario), capturing the output of the command for storing in
	' the excel spreadsheet.  Make sure that the success case is the
	' first string in the array.
	'msgBox "Inside show ver subrouting " & vbcrlf & "g_Prompt is : " & g_Prompt
	vStringsToWaitFor = Array(_
		g_Prompt, _
		"Type help or '?' for a list of available commands.")
	strCDP = crt.Screen.ReadString(vStringsToWaitFor)
	' Get the current line so that we can peel it off of from the
	' results that were captured from the data sent by the remote.
	'strCurLine = crt.Screen.Get(_
	'	crt.Screen.CurrentRow, _
	'	0, _
	'	crt.Screen.CurrentRow, _
	'	crt.Screen.CurrentColumn)
	'strCDP = Left(strCDP, Len(strCDP) - Len(strCurLine))
	strCDP = Trim(strCDP)  
	strCDP = strCDP & vbcrlf & "-------------------------"
	
	'Set Active Worksheet to the Interfaces tab.
    Set g_objWriteSheet = g_objWriteWkBook.Sheets("CDP")
	
	
	set reCDPNEIGHID_CDP = New RegExp
	set reIPADD_CDP = New RegExp
	'set reIPv6ADD_CDP = New RegExp
	set rePLATFORM_CDP = New RegExp
	set reIFLOCAL_CDP = New RegExp
	set reIFREMOTE_CDP = New RegExp
	set reREMOTEIOS_CDP = New RegExp
	set rePOWER_CDP = New RegExp
	set reENDMARK_CDP = New RegExp
	set reMGMTADD_CDP = New RegExp
	
	reCDPNEIGHID_CDP.pattern = "Device ID: (.*)"							' Tested and works
	reIPADD_CDP.pattern = "IP address: (\d+\.\d+\.\d+\.\d+)"						'Tested and works					' Tested and works
	'reIPv6ADD_CDP.pattern = "(([0-9a-fA-F]{1,4}:){1,7})"							' Doesn't work.  Skipping for now.	
	rePLATFORM_CDP.pattern = "Platform: [cC]isco (.*),.*"							' Tested and works
	reIFLOCAL_CDP.pattern = "[iI]nterface: (.*),.*"											' Tested and works
	reIFREMOTE_CDP.pattern = "[iI]nterface:.*port\): (.*)"							'Tested and works
	reREMOTEIOS_CDP.pattern = "Cisco IOS Software.*Version (.*)[,| ].*"													
	rePOWER_CDP.pattern = "Power drawn: (.* [wW]).*"											
	reMGMTADD_CDP.pattern = "(Management address).*"
	reENDMARK_CDP.pattern = "-----.*-----"										' Tested and works
  
  


  ' Determine total number of lines in the strInterfaces.  Don't need header Since we created one.
  vLines = Split(strCDP, vbcrlf)
  
  'Cycle through each 
  'MsgBox "Starting nIndex at: " & g_CDPnRowIndex & vbcrlf & vbcrlf & "Total number of lines in the Interface output is " & (UBound(VLines) - 1)
  
  If g_CDPnRowIndex = "" Then g_CDPnRowIndex = 1
	'msgBox "g_CDPnRowIndex is now set to" & g_CDPnRowIndex
	
  If g_CDPnRowIndex = 1 Then
	g_objWriteSheet.Cells(g_CDPnRowIndex, vHOSTNAME_COL).Value = "Hostname"
	g_objWriteSheet.Cells(g_CDPnRowIndex, vCDPNEIGHID_COL).Value = "Remote Hostname"
	g_objWriteSheet.Cells(g_CDPnRowIndex, vIPADD_COL).Value = "Remote IP"
	g_objWriteSheet.Cells(g_CDPnRowIndex, vMGMTIP_COL).Value = "Remote Mgmt Address"
	
	g_objWriteSheet.Cells(g_CDPnRowIndex, vPLATFORM_COL).Value = "Remote Model"
	g_objWriteSheet.Cells(g_CDPnRowIndex, vIFLOCAL_COL).Value = "Local Interface"
	g_objWriteSheet.Cells(g_CDPnRowIndex, vIFREMOTE_COL).Value = "Remote Interface"
	g_objWriteSheet.Cells(g_CDPnRowIndex, vREMOTEIOS_COL).Value = "IOS"
	g_objWriteSheet.Cells(g_CDPnRowIndex, vPOWER_COL).Value = "Power Drawn"
	g_CDPnRowIndex = g_CDPnRowIndex + 1
    End If

  ' Determine total number of lines in the strInterfaces.  Don't need header Since we created one.
  vLines = Split(strCDP, vbcrlf)
	i = 0
	For Each strLine In vLines
	
		If reCDPNEIGHID_CDP.Test(strLine) <> TRUE Then	
			'MsgBox "Pattern """ & reLineStatus.Pattern & """ wasn't found within the following text: " & _
			'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
		 Else
				'Set variable to notate new line starting.
				i = 1
				g_objWriteSheet.Cells(g_CDPnRowIndex, vHOSTNAME_COL).Value = g_hostname
				Set matches = reCDPNEIGHID_CDP.Execute(strLine)
				For each match in matches
					'MsgBox "Pattern matched for LineStatus for iteration " & g_CDPnRowIndex & " : " & match.SubMatches(0)
					g_objWriteSheet.Cells(g_CDPnRowIndex, vCDPNEIGHID_COL).Value = match.SubMatches(0)
				Next

		End If
		
		'Only look for associated data if we find an applicable Device ID.
		If i = 1 Then
			If reIPADD_CDP.Test(strLine) <> TRUE Then	
				'MsgBox "Pattern """ & reLineStatus.Pattern & """ wasn't found within the following text: " & _
				'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
			Else
				Set matches = reIPADD_CDP.Execute(strLine)
				For each match in matches
					'MsgBox "Pattern matched for LineStatus for iteration " & g_CDPnRowIndex & " : " & match.SubMatches(0)
					g_objWriteSheet.Cells(g_CDPnRowIndex, vIPADD_COL).Value = match.SubMatches(0)
				Next
			End If
			
			If rePLATFORM_CDP.Test(strLine) <> TRUE Then	
				'MsgBox "Pattern """ & reLineStatus.Pattern & """ wasn't found within the following text: " & _
				'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
			Else
				Set matches = rePLATFORM_CDP.Execute(strLine)
				For each match in matches
					'MsgBox "Pattern matched for LineStatus for iteration " & g_CDPnRowIndex & " : " & match.SubMatches(0)
					g_objWriteSheet.Cells(g_CDPnRowIndex, vPLATFORM_COL).Value = match.SubMatches(0)
				Next
			End If
			
			If reIFLOCAL_CDP.Test(strLine) <> TRUE Then	
				'MsgBox "Pattern """ & reLineStatus.Pattern & """ wasn't found within the following text: " & _
				'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
			Else
				Set matches = reIFLOCAL_CDP.Execute(strLine)
				For each match in matches
					'MsgBox "Pattern matched for LineStatus for iteration " & g_CDPnRowIndex & " : " & match.SubMatches(0)
					g_objWriteSheet.Cells(g_CDPnRowIndex, vIFLOCAL_COL).Value = match.SubMatches(0)
				Next
			End If
			
			If reIFREMOTE_CDP.Test(strLine) <> TRUE Then	
				'MsgBox "Pattern """ & reLineStatus.Pattern & """ wasn't found within the following text: " & _
				'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
			Else
				Set matches = reIFREMOTE_CDP.Execute(strLine)
				For each match in matches
					'MsgBox "Pattern matched for LineStatus for iteration " & g_CDPnRowIndex & " : " & match.SubMatches(0)
					g_objWriteSheet.Cells(g_CDPnRowIndex, vIFREMOTE_COL).Value = match.SubMatches(0)
				Next
			End If
			
			If reREMOTEIOS_CDP.Test(strLine) <> TRUE Then	
				'MsgBox "Pattern """ & reLineStatus.Pattern & """ wasn't found within the following text: " & _
				'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
			Else
				Set matches = reREMOTEIOS_CDP.Execute(strLine)
				For each match in matches
					'MsgBox "Pattern matched for LineStatus for iteration " & g_CDPnRowIndex & " : " & match.SubMatches(0)
					g_objWriteSheet.Cells(g_CDPnRowIndex, vREMOTEIOS_COL).Value = match.SubMatches(0)
				Next
			End If
			
			If rePOWER_CDP.Test(strLine) <> TRUE Then	
				'MsgBox "Pattern """ & reLineStatus.Pattern & """ wasn't found within the following text: " & _
				'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
			Else
				Set matches = rePOWER_CDP.Execute(strLine)
				For each match in matches
					'MsgBox "Pattern matched for LineStatus for iteration " & g_CDPnRowIndex & " : " & match.SubMatches(0)
					g_objWriteSheet.Cells(g_CDPnRowIndex, vPOWER_COL).Value = match.SubMatches(0)
				Next
			End If
			
			If reMGMTADD_CDP.Test(strLine) <> TRUE Then	
				'MsgBox "Pattern """ & reLineStatus.Pattern & """ wasn't found within the following text: " & _
				'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
			Else
				Set matches = reMGMTADD_CDP.Execute(strLine)
				For each match in matches
					'MsgBox "Pattern matched for LineStatus for iteration " & g_CDPnRowIndex & " : " & match.SubMatches(0)
					i = 2
				Next
			End If
			
			If reENDMARK_CDP.Test(strLine) <> TRUE Then	
				'MsgBox "Pattern """ & reLineStatus.Pattern & """ wasn't found within the following text: " & _
				'vbcrlf & vbcrlf & vbcrlf & """" & strLine & """"
			Else
				' Reset i to 0 to evaluate a new line.  Increment row counter.
				i = 0
				g_CDPnRowIndex = g_CDPnRowIndex + 1
			End If
			
		Else 
			if i = 2 Then
				
				If reENDMARK_CDP.Test(strLine) = TRUE Then	
					' Reset i to 0 to evaluate a new line.  Increment row counter.
					i = 0
					g_CDPnRowIndex = g_CDPnRowIndex + 1				
				ElseIf reIPADD_CDP.Test(strLine) = TRUE Then	
					Set matches = reIPADD_CDP.Execute(strLine)
					For each match in matches
						'MsgBox "Pattern matched for LineStatus for iteration " & g_CDPnRowIndex & " : " & match.SubMatches(0)
						g_objWriteSheet.Cells(g_CDPnRowIndex, vMGMTIP_COL).Value = match.SubMatches(0)
					Next
					i = 1 			' Send back to main part of loop with i variable reset to 1.
				Else
					i = 1
				End If
			
			End If
		End If

	Next
 
	Set g_objWriteSheet = g_objWriteWkBook.Sheets("Main")

End Sub


Function Connect(strIP, strPort, strProtocol, strUsername, strPassword)
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 ' Workaround that uses "On Error Resume Next" VBScript directive to detect
 ' Errors that might occur from the crt.Session.Connect call and instead of
 ' closing the script, allow for error handling within the script as the
 ' script author desires.
    On Error Resume Next

    ' First disconnect if we're already connected.
	If strPort = "" Then Exit Function
    If crt.Session.Connected Then crt.Session.Disconnect
	If strUsername = "" Then crt.dialog.prompt 

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
            crt.Screen.WaitForString("ssword:")
            crt.Screen.Send strPassword & vbcr
            Connect = True
            
		Case "CONSOLE"
            strCmd = "/TELNET " & strIP & " " & strPort
            crt.Session.Connect strCmd
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

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Function ConnectOLD(strIP, strPort, strProtocol, strUsername, strPassword)
 '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
 ' Workaround that uses "On Error Resume Next" VBScript directive to detect
 ' Errors that might occur from the crt.Session.Connect call and instead of
 ' closing the script, allow for error handling within the script as the
 ' script author desires.
    On Error Resume Next

    ' First disconnect if we're already connected.
    If crt.Session.Connected Then crt.Session.Disconnect
	If strUsername = "" Then crt.dialog.prompt 

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
            crt.Screen.WaitForString("ssword:")
            crt.Screen.Send strPassword & vbcr
			
			
			
			Do
				bCursorMoved = crt.Screen.WaitForCursor(2)
			Loop Until bCursorMoved = False
			nRow = objTab.Screen.CurrentRow
			msgBox "nRow is " & nRow
			szPrompt = objTab.screen.Get(nRow, _
						 0, _
						 nRow, _
						 objTab.Screen.CurrentColumn - 1)
			szPrompt = Trim(szPrompt)
			
			If szPrompt = "Password:" Then
				Connect = False
			Else
				Connect = True
				crt.Screen.Synchronous = True
			End If
			
		Case "CONSOLE"
            strCmd = "/TELNET " & strIP & " " & strPort
            crt.Session.Connect strCmd
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
            If crt.Session.Connected <> True Then Exit Function
			
			msgBox "Current line is " & objTab.screen.Get(nRow, 0, nRow, objTab.Screen.CurrentColumn - 1)
 
			Do
				bCursorMoved = crt.Screen.WaitForCursor(2)
			Loop Until bCursorMoved = False
			nRow = objTab.Screen.CurrentRow
			szPrompt = objTab.screen.Get(nRow, _
						 0, _
						 nRow, _
						 objTab.Screen.CurrentColumn - 1)
			szPrompt = Trim(szPrompt)
			
			If szPrompt = "Password:" Then
				Connect = False
			Else
				Connect = True
				crt.Screen.Synchronous = True
			End If

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

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~


Function Continue(strMsg)
    Continue = True
    WScript.Sleep 200
    If msgBox(strMsg, vbYesno) <> vbYes Then Continue = False
    WScript.Sleep 400
End Function

Function GetMyDocumentsFolder()
    Dim myShell
    Set myShell = CreateObject("WScript.Shell")

    GetMyDocumentsFolder = myShell.SpecialFolders("MyDocuments")
End Function

Function NN(nNumber, nDesiredDigits)
 ' Normalizes a single digit number to have nDesiredDigits 0s in front of it
    Dim nIndex, nOffbyDigits, szResult
    nOffbyDigits = nDesiredDigits - len(nNumber)

    szResult = nNumber

    For nIndex = 1 to nOffByDigits
        szResult = "0" & szResult
    Next
    NN = szResult
End Function

Function GetPlainTextInput(strPrompt, strTitle, strDefaultValue)
    GetPlainTextInput = _
        crt.Dialog.Prompt(strPrompt, strTitle, strDefaultValue)
End Function

'-------------------------------------------------------------------------------
Function GetPasswordInput(strPrompt, strTitle, strDefaultValue)
    GetPasswordInput = _
        crt.Dialog.Prompt(strPrompt, strTitle, strDefaultValue, True)
End Function


Function GetLastLines(s, nLinesToDisplay, lineBreakChar)
  ' Used for testing only.  Not functional yet.
    'Split the string into an array
    Dim splitString()
	msgBox "LineBreakChar is " & lineBreakChar & vbcrlf & "Displaying last " & nLinestoDisplay & " lines."
    splitString = Split(s, vbcrlf)

    'How many lines are there?
    Dim nLines
    nLines = UBound(splitString) + 1

    If nLines <= nLinesToDisplay Then
        'No need to remove anything. Get out.
        GetLastLines = s
        Exit Function
    End If

    'Collect last N lines in a new array
    Dim lastLines()
    'ReDim lastLines(0 To nLinesToDisplay - 1)
    Dim i
    For i = 0 To UBound(lastLines)
        lastLines(i) = splitString(i + nLines - nLinesToDisplay)
    Next

    'Join the lines array into a single string
    GetLastLines = Join(lastLines, lineBreakChar)
End Function


'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' Definition of our main subroutine
Function ConnectWithFallback(strIP, strPort, strUsername, strPassword)
    ' This script will only work with SecureCRT version 5.0 or later.
	On Error Resume Next
    If Int(Left(crt.Version, 1)) < 5 then
        MsgBox "This script works with SecureCRT 5.0 or later." & vbcrlf & _
               "The version you are running is " & crt.Version
        exit function
    end if

'    g_szSessionName = crt.Session.Path

    ' Read in the session configuration from the .ini file...
'    g_szConfigData = ReadSessionConfig
'    if g_szConfigData = "" then
'        Err.Raise 50001, "Sub ConnectWithFallback()" & vbcrlf & _
 '           "Unable to read session configuration."
'        exit function
'    end if

    ' In order to make subsequent connection attempts, we'll
    ' need to gain access to the destination host from
    ' within the session's .ini file.

    if strIP = "" then exit function

    if crt.Session.Connected then
        'MsgBox "Disconnecting"
        crt.Session.Disconnect
    end if

    ' MsgBox szHost & vbcrlf & vbcrlf & "Connected: " & crt.Session.Connected

    crt.Screen.Synchronous = true
    Dim szSecureShellIdent
    If Not CheckRemoteSSHConnectivity(szSecureShellIdent, strIP) then
        If not ConnectWithTelnetProtocol (strIP, strUsername, strPassword) Then
			ConnectWithFallback = False
		Else
			ConnectWithFallback = TRUE
		End If
        exit function
    end if

    if  Instr(szSecureShellIdent, "SSH-2") > 0 then
        If not ConnectWithSSH2Protocol (strIP, strUsername, strPassword) Then
			ConnectWithFallback = False
		Else
			ConnectWithFallback = TRUE
		End If
    ElseIf Instr(szSecureShellIdent, "SSH-1.99") > 0 then
        ' 1.99 indicates the server should support both... let's choose
        ' the more robust SSH2 protocol.
        If not ConnectWithSSH2Protocol (strIP, strUsername, strPassword) Then
			ConnectWithFallback = False
		Else
			ConnectWithFallback = TRUE
		End If
    Elseif Instr(szSecureShellIdent, "SSH-1") > 0 then
        If not ConnectWithSSH1Protocol (strIP, strUsername, strPassword) Then
			ConnectWithFallback = False
		Else
			ConnectWithFallback = TRUE
		End If
    end if

End function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function GetConfigPath()
    On Error Resume Next
    szRegPath = "HKCU\Software\VanDyke\SecureCRT\Config Path"
    Err.Clear
    GetConfigPath = g_shell.RegRead(szRegPath)
    If Err.Number <> 0 then
        On Error Goto 0
        ' Fancy way of bailing out of the script early...
        Err.Raise 50001, "Function GetConfigPath()", _
            "SecureCRT 5.x Configuration folder not found in the registry." & _
            vbcrlf & "A hard-coded path is needed in place of the " & _
            "GetConfigPath() function." & _
            vbcrlf
    end if
    On Error Goto 0
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function GetConfigParameter(szConfigData, szOptionName)
 ' Returns the value of an option line from a SecureCRT .ini file
 ' that matches the following pattern:
 '   S:"Option Name"=Option Value
 
 
    Set re = new RegExp
    re.Pattern = "^S\:""" & szOptionName & """=(.+)$"
    re.Global = False
    re.Multiline = True
    re.IgnoreCase = True

    if re.Test(szConfigData) <> True then
        Err.Raise 50002, "Function GetConfigParameter()" & vbcrlf & _
            "String Option """ & szOptionName & _
            """ not found in session configuration: " & g_szSessionName
        exit function
    end if

    Set matches = re.Execute(szConfigData)
    For each Match in Matches
        szParam = match.Submatches(0)
        szParam = Replace(szParam, vbcr, "")
        szParam = Replace(szParam, vblf, "")
        szParam = Trim(szParam)
        GetConfigParameter = szParam
        exit function
    Next
end function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function ReadSessionConfig()
    szINIfilename = GetConfigPath & "\Sessions\" & g_szSessionName & ".ini"
    if Not g_fso.FileExists(szINIfilename) then
        Err.Raise 50003, "Function ReadSessionConfig()" & vbcrlf & _
            "Session's .ini file not found: " & vbcrlf & _
            vbtab & szINIfilename
        exit function
    end if

    Set objFile = g_fso.OpenTextFile(szINIfilename, ForReading, False)
    ReadSessionConfig = objFile.ReadAll
    objFile.Close
    Set objFile = nothing
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Function CheckRemoteSSHConnectivity(byRef szIdent, strIP)
    On Error Resume Next

    ' Telnet to port 22 to determine if there's anything running there.
    crt.Session.Connect "/TELNET " & strIP & " 22"

    ' Check the error level... Connect will fail if there
    ' isn't anything listening on port 22.
    if Err.Number <> 0 then
        ' There must not be any SSH connectivity (or there
        ' is a problem connecting in general to the remote
        ' machine).  In Either case, there's no SSH
        ' connectivity, so we should exit/return false
        ' (default value for any VBScript function)
        ' MsgBox "Error connecting to host """ & g_szHost & """: " & _
        '     Err.Description
        exit function
    end if

    ' If we're still here in this function, then we must
    ' have successfully connected to port 22 on the remote
    ' machine. Now we need to wait for an SSH ident string
    ' to appear so we have a better idea as to the version
    ' of SSH that is available to us...
    crt.Screen.Synchronous = True
    nResult = crt.Screen.WaitForString("SSH-", 10)
    nResult = crt.Screen.WaitForString("-", 10)

    ' If we don't get the SSH ident string within 10
    ' seconds, we likely don't have any connectivity.  Bail
    ' out now.  However, you may need to adjust the timeout
    ' parameter in the WaitForString() call above to meet
    ' the highest latency possible among all of your connections.
    if nResult = 0 then
        crt.Session.Disconnect
        exit function
    end if

    szIdent = crt.Screen.Get(crt.Screen.CurrentRow, _
                             0, _
                             crt.Screen.CurrentRow, _
                             crt.Screen.Columns)

    crt.Session.Disconnect

    CheckRemoteSSHConnectivity = True
    On Error Goto 0
End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function ConnectWithTelnetProtocol(strIP, strUsername, strPassword)
	On Error Resume Next
    ' MsgBox "Connecting to " & chr(34) & g_szSessionName & chr(34) & _
    '   " using Telnet."
	strCmd = "/TELNET " & strIP & " 23"
	crt.Session.Connect strCmd
	If crt.Session.Connected <> True Then 
		g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = _
		"Failure: TELNET and SSH connection Failed."
		Exit Function
	End If
	crt.Screen.Synchronous = True
	 
 
	Do
		' Look for username and password prompts
		vWaitFors = Array("sername:", "assword:", "ogin:", "Incorrect Password")
		nResult = crt.Screen.WaitForStrings (vWaitFors, 5)
		'msgBox "Starting Select Case" & vbcrlf & "nResult = " & nResult
		Select Case nResult
			Case 0
				ConnectWithTelnetProtocol = False
				g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = _
				"Failure: TELNET and SSH connection Failed."	
				Exit Function
			Case 1
				crt.Screen.Send strUsername & vbcr
				crt.Screen.WaitForString "ssword:", 5
				crt.Screen.Send strPassword & vbcr
				nResult2 = crt.Screen.WaitForStrings (vWaitFors, 5)
				If (nResult2 = 2 OR nResult = 4) Then 
					ConnectWithTelnetProtocol = False
					g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = _
						"Failure: TELNET connection Failed. Password not correct"					
					Exit Function
				End If
				Exit Do
			Case 2
				crt.Screen.Send strPassword & vbcr
				nResult2 = crt.Screen.WaitForStrings (vWaitFors, 5)
				If (nResult2 = 2 OR nResult = 4) Then 
					ConnectWithTelnetProtocol = False
					g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = _
						"Failure: TELNET connection Failed. Password not correct"						
					Exit Function
				End If
				Exit Do
			Case 3
				crt.Screen.Send strUsername & vbcr
				crt.Screen.WaitForString "ssword:", 5
				crt.Screen.Send strPassword & vbcr
				nResult2 = crt.Screen.WaitForStrings (vWaitFors, 5)
				If (nResult2 = 2 OR nResult = 4) Then 
					ConnectWithTelnetProtocol = False
					g_objWriteSheet.Cells(nRowIndex, g_ACT_RES_COL).Value  = _
						"Failure: TELNET connection Failed. Password not correct"								
					Exit Function
				End If
				Exit Do
		End Select
	Loop
	
	
	ConnectWithTelnetProtocol = True
	g_objWriteSheet.Cells(nRowIndex, g_PORT_COL).Value = "23"
	g_objWriteSheet.Cells(nRowIndex, g_PROTO_COL).Value = "Telnet"

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function ConnectWithSSH1Protocol(strIP, strUsername, strPassword)
	On Error Resume Next
        szConnectCommand = "/SSH1 /P 22 /L " & strUsername & " /PASSWORD " & strPassword &_ 
		" /ACCEPTHOSTKEYS " & strIP
	g_objWriteSheet.Cells(nRowIndex, g_PORT_COL).Value = "22"
	g_objWriteSheet.Cells(nRowIndex, g_PROTO_COL).Value = "SSH1"


    crt.Session.Connect szConnectCommand

End Function

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Function ConnectWithSSH2Protocol(strIP, strUsername, strPassword)
	On Error Resume Next
    szConnectCommand = "/SSH2 /P 22 /L " & strUsername & " /PASSWORD " & strPassword &_
	"  /ACCEPTHOSTKEYS " &   strIP


    'MsgBox "Connecting to " & chr(34) & g_szSessionName & chr(34) & _
    '    " using SSH2 with the following connect string:" & vbcrlf & _
    '    szConnectCommand

    crt.Session.Connect szConnectCommand
	If crt.Session.Connected <> True Then Exit Function
	g_objWriteSheet.Cells(nRowIndex, g_PORT_COL).Value = "22"
	g_objWriteSheet.Cells(nRowIndex, g_PROTO_COL).Value = "SSH2"
	crt.Screen.Synchronous = True
	ConnectWithSSH2Protocol = True
End Function