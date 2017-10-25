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
