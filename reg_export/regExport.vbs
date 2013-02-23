'**************************************** SCRIPT SETTINGS ****************************************

Const intUnicode = -1
Const intForReading = 1
Const intForWriting = 2
Const intForAppending = 8
	
Dim regPathsFile, regTemplateFile, flagHeadLine
Dim regFormatLine,templateCheckQ
Dim errorMessage1,errorMessage2,errorMessage3,errorMessage4,errorMessage5,errorMessage6,errorMessage7,errorMessage8,errorMessage9

regPathsFile = "c:\Release\App\Registry\regPaths.txt"																'Set regPathsFile name
regTemplateFile = "c:\Release\App\Registry\template.reg"																'Set templateFile name

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

regFormatLine = "Windows Registry Editor Version 5.00"
templateCheckQ = "It seems that a template file already exists. What would you like to do?"
errorMessage1 = "All regPaths.txt Lines must start with:" & vbCr _									
				& """;"" - to ignore a line" & vbCr _
				& """#LINESPACE#"" - to separate between registry entry bulks" & vbCr _
				& "or ""H"" to provide a registry entry"
errorMessage2 = "Key not found in registry file or there is a problem with hive exporting"
errorMessage3 = "Key not found in registry file"
errorMessage4 = "Value not found in registry file"
errorMessage5 = "You're trying to export a hive while specificying a value. The Hive WILL be exported to the Template File"
errorMessage6 = "You're trying to export a value without specificying one"
errorMessage7 = "Unexpected error occured - Check regPaths file for syntax errors"
errorMessage8 = "UNKNOWN line prefix"
errorMessage9 = "Invalid Choice - exiting"

'**************************************** SCRIPT ****************************************
Main
	
Sub Main
	call templateFileCheck(regTemplateFile)
	call RegExport(CreateRegArray(regPathsFile))
End Sub




'**************************************** FUNCTIONS ****************************************

'****************************** creates an array from the registry table file ******************************
Function CreateRegArray(regPathsFile)															'Function creates an array from the registry table file 
	Set objPathsFile = objFSO.OpenTextFile(regPathsFile, intForReading) 											'Open regPathsFile for reading
	
	Dim arrRegPaths() 																				'Define arrRegPaths array.
	i = 0 																							'Loop variable
	Do Until objPathsFile.AtEndOfStream 															'Loop regPathsFile until the file ends 																	
		strLin = objPathsFile.ReadLine																	'Read Line in regPathsFile
		If Not Left(strLin, 1) = "" AND Not Left(strLin, 1) = ";" Then 									'If first char of Line in regPathsFile is NOT ";" And NOT "nothing"																																	
			ReDim Preserve arrRegPaths(i)																	'Preserve next line of arrRegPaths
			arrRegPaths(i) = strLin																			'Set i place in arrRegPaths to be the whole line in the registry paths file	 																						'loop advancement
			i = i+1																							'advance Line in newly created Array
		End If
	Loop 																							
	objPathsFile.Close 																				'Close regPathsFile
	CreateRegArray=arrRegPaths 																		'Return the Array to Main
End Function

'****************************** Exports Registry hives and values ******************************
Sub RegExport(arrRegPaths())																	'Sub Exports Registry keys and values
	
	Dim regTempKeyFile,strCommand																'...
	Dim strRegKey,strRegValue																		'Set regKey and regValue vars
	
	If objFSO.fileExists(regTemplateFile) Then														'If templateFile exists
		Set objTemplateFile = objFSO.OpenTextFile(regTemplateFile, intForReading, False, intUnicode)	'Open templateFile forReading
		strLine = objTemplateFile.ReadLine																'Read 1st Line of templateFile
		If strLine = regFormatLine Then																	'If 1st Line of templateFile is regFormatLine
			flagFormatLine = True																			'Raise flagFormatLine
		End If																							'End If																
		objTemplateFile.Close																			'Close templateFile
	End If																							'End If
	
	Set objTemplateFile = objFSO.OpenTextFile(regTemplateFile, intForAppending, True, intUnicode)	'Open templateFile
	If Not flagFormatLine Then																		'If flagFormatLine is not Raised
		objTemplateFile.WriteLine regFormatLine															'Write regFormatLine to templateFile
	End If																							'End If
	objTemplateFile.Close																			'Close Template File	
	
	For Each strRegPath In arrRegPaths 																'For each Line in the arrRegPaths
		on Error Resume Next
		arrRegEntry = Split(strRegPath,",") 															'split each line to Key and Value
		If Err.Number <> 0 Then																			'If error occurs
			wscript.echo errorMessage7 & ": " & strRegPath													'echo errorMessage7
		End If
		strRegKey = arrRegEntry(1) 																		'set Key string
		strRegValue = arrRegEntry(2)																	'set Value string
		on Error Goto 0
		
		regTempKeyFile = Replace(strRegKey, "\", "_") & ".reg"									 	'set temporary regTempKeyFile
		strCommand = "cmd /c REG EXPORT """ & strRegKey & """ """ & regTempKeyFile & "" 			'set batch command for registry export to a regTempKeyFile
		objShell.Run strCommand, 0, True 																'run the export command
		
		If objFSO.FileExists(regTempKeyFile) = True Then						 					'If a regTempKeyFile EXISTS
			Select Case Left(strRegPath, 1)																	'Select Case according to the first char of the strLine read 
				Case "h"																						'Case "h" (Hive)
					If Not strRegValue = "" Then																	'If regValue IS defined
						wscript.echo errorMessage5 & ": " & strRegPath													'echo errorMessage5
					End If																							'End If
					call PrintHiveToTemplate(regTempKeyFile,regTemplateFile)									'call Sub to export hive
				Case "v"																						'Case "v" (Value)
					If strRegValue = "" Then																		'If regValue IS NOT defined
						wscript.echo errorMessage6 & ": " & strRegKey													'echo errorMessage6
					End If																							'End If
					call PrintValueToTemplate(regTemplateFile,findRegInText(strRegKey,strRegValue,regTempKeyFile))	'call Sub to export value
				Case "x"																						'Case "x" (Exclude)
					wscript.echo "NOT YET IMPLEMENTED"
					wscript.Quit
				
					'Set objTemplateFile = objFSO.OpenTextFile(regTemplateFile, intForReading, False, intUnicode)
					'Do Until objTemplateFile.AtEndOfStream
					'	strCheck = objTemplateFile.ReadLine
					'	If strCheck = "[" & strRegKey & "]" Then
					'		flagKeyExists = True
					'	End If
					'Loop
					
					'If flagKeyExists Then
					'	call DeleteExcluded(strRegKey,strRegValue,regTemplateFile)
					'Else
					'	call PrintHiveToTemplate(regTempKeyFile,regTemplateFile)									'call Sub to export hive
					'	call DeleteExcluded(strRegKey,strRegValue,regTemplateFile)
					'End If
				Case Else																						'Case Else
					wscript.echo errorMessage8 & ": " & strRegPath													'echo errorMessage8
			End Select																						'End Select
			objFSO.DeleteFile regTempKeyFile, True														'Delete regTempKeyFile
		Else																							'Else a regTempKeyFile DOES NOT EXIST
			wscript.echo errorMessage2 & ": " & regTempKeyFile											'Echo errorMessage2
		End If																							'End If
	Next																							'Next Loop
	
	strRegKey = NULLKEY		 																		'set Key string
	strRegValue = NULLVALUE 																		'set Value string
	Set objTemplateFile = Nothing																	'Recycle templateFile variable
End Sub																							'End Sub



'****************************** Checks for template file existance ******************************
Sub templateFileCheck(regTemplateFile)
	Dim bullet, choice
	
	bullet = Chr(10) & "   " & Chr(149) & " "
	If objFSO.fileExists(regTemplateFile) Then
		choice = InputBox(templateCheckQ & Chr(10) _ 
		& bullet & "1.) Delete Old Template file and Continue" _
		& bullet & "2.) Append to existing template file" _ 
		& bullet & "3.) Exit" _ 
		& Chr(10), "Select Action")
		
		Select Case choice
		Case "1" 
			objFSO.DeleteFile regTemplateFile
		Case "2"
			Exit Sub
		Case "3"
			wscript.Quit
		Case Else
			wscript.echo errorMessage9
			wscript.Quit
		End Select
	End If
End Sub

'****************************** Find value and its key in a file a return the result ******************************
Function findRegInText(strKeyToFind,strValueToFind,fileInput)									'Function Find A string in a text file and copy its line

	Dim strLine,indValueLine																		'define Line variable and Line index
	Dim flagKeyFound,flagValueFound,flagMultiline													'define flags
	
	flagKeyFound = False																			'set key flag found
	flagValueFound = False																			'set value flag found
	
	ReDim arrRegValueData(0)																		'define arrRegValueData
	indValueLine = Ubound(arrRegValueData) 															'Set 0th Multiline Array Index
	ReDim Preserve arrRegValueData(indValueLine+1) 													'Preserve next Line of Multiline Array for next Line of registry value
	
	Set regInputFile = objFSO.OpenTextFile(fileInput, intForReading, False, intUnicode) 			'open the regTempKeyFile for parsing
	regInputFile.SkipLine																			'SkipLine to skip on regFormatLine
	
	Do until regInputFile.atEndOfStream 															'Loop until ending of regTempKeyFile
		
		strLine = regInputFile.ReadLine 																'Read line of regTempKeyFile
		Select Case Left(strLine, 1) 																	'Select Case according to the first char of the strLine read
		
		'******************** Case line starts with PARANTHESIS ********************
		Case ""
			
		Case "["																						'Case line starts with PARANTHESIS
			If InStr(strLine,chr(91) & strKeyToFind & chr(93)) > 0 Then 									'If relevant text WAS found in Line
				flagKeyFound = True																				'Enable flagKeyFound
				arrRegValueData(indValueLine) = strLine															'Set the registry key to the 1st Line of the Array
				indValueLine = Ubound(arrRegValueData)															'Set 1st Multiline Array Index
				ReDim Preserve arrRegValueData(indValueLine+1)													'Preserve next Line of Multiline Array for next strLine of registry value 3
			End If																							'End If
			
		'******************** Case line starts with DOUBLE QUOTES ********************
		Case chr(34)																					'Case strLine starts with DOUBLE QUOTES
			If InStr(strLine,chr(34) & strValueToFind & chr(34)) > 0 Then 									'If relevant text WAS found in strLine
				flagValueFound = True																			'Enable flagValueFound
				arrRegValueData(indValueLine) = strLine 														'Set the first (or only) Line of the registry value to the 2nd Line of the Multiline Array 
				indValueLine = Ubound(arrRegValueData) 															'Set 2nd Multiline Array Index
				ReDim Preserve arrRegValueData(indValueLine+1)													'Preserve next Line of Multiline Array for next Line of registry value
				If Right(strLine, 1) = "\" Then 																'If strLine CONTAINS an ending \ (meaning Multiline Value)
					flagMultiline = True 																			'Enable Multiline Flag
				Else 																							'Else strLine DOES NOT CONTAIN an ending \ (meaning Singleline Value)
					findRegInText = arrRegValueData 																'Return Singleline String of registry value to function
					Exit Do 																						'exit loop
				End If																							'End If
			End If																							'End If
			
		'******************** Case line starts with SOMETHING ELSE ********************
		Case Else																						'Case strLine starts with SOMETHING ELSE
			If flagMultiline Then	 																		'If this strLine IS NOT a part of a Multiline Value
				If Right(strLine, 1) = "\" Then 																'If strLine CONTAINS an ending \ (meaning Multiline Value)
					arrRegValueData(indValueLine) = strLine															'Set following strLine of registry value to following Line of Multiline Array
					indValueLine = Ubound(arrRegValueData) 															'Set following Multiline Array Index
					ReDim Preserve arrRegValueData(indValueLine+1)													'Preserve next Line of Multiline Array for next strLine of registry value
				Else 																							'Else strLine DOES NOT CONTAIN an ending \ (meaning Mutliline Value Ended)
					arrRegValueData(indValueLine) = strLine															'Set following strLine of registry value to following Line of Multiline Array
					findRegInText = arrRegValueData 																'Return Multiline Array of registry value to function
					flagMultiline = False 																			'Disable Multiline Flag
					Exit Do 																						'exit loop
				End If																							'End If
			End If																							'End If
		End Select																						'End Select
	Loop																							'Next Loop
	
	If Not flagKeyFound Then																		'If
		wscript.echo errorMessage3 & ": " & strKeyToFind 
		'wscript.Quit 3
	End If																							'End If
	If Not flagValueFound Then																		'If
		wscript.echo errorMessage4 & ": " & strValueToFind
		wscript.Quit 4								
	End If																							'End If
	regInputFile.Close 																				'close the regTempKeyFile
End Function																					'End Function

'****************************** Prints the registry hive to the template file ******************************
Sub PrintHiveToTemplate(regTempKeyFile,regTemplateFile)									'Sub Prints the registry Hive to the templateFile
	Set objKeyFile = objFSO.OpenTextFile(regTempKeyFile, intForReading, False, intUnicode) 	'Open the regTempKeyFile				
	objKeyFile.SkipLine 																			'Skip reading the regFormatLine
	Set objTemplateFile = objFSO.OpenTextFile(regTemplateFile, intForAppending, True, intUnicode)	'Open templateFile for appending
	objTemplateFile.Write objKeyFile.ReadAll 														'Write everything else to the templateFile
	objTemplateFile.Close																			'Close templateFile
	objKeyFile.Close																				'Close temporaryRegKeyFile
End Sub																							'End Sub

'****************************** Prints the registry key and value to the template ******************************
Sub PrintValueToTemplate(regTemplateFile,arrRegValueData())										'Sub Prints the registry key/value to the templateFile							
	Set regOutputFile = objFSO.OpenTextFile(regTemplateFile, intForAppending, True, intUnicode) 	'Open templateFile for appending
	'regOutputFile.Write vbCrLf																		'Print a line break to templateFile
	For Each regValueLine in arrRegValueData														'For each registry value line in registry value Array
		regOutputFile.Write regValueLine & vbCrLf														'Print the registry value line to the templateFile
	Next																							'Next Loop
	regOutputFile.Close 																			'Close templateFile
End Sub																							'End Sub

'****************************** Deletes Excluded Values from Template File *******************************
Sub DeleteExcluded(regTemplateFile,regKeyToSearchIn,regValueToDelete)
'
'
End Sub

















'************************************************************************************************
Sub ProcessExcludeList(arrExclude())
	Dim oFS	:	Set oFS = CreateObject("Scripting.FileSystemObject")
	Dim oRegExp	:	Set oRegExp = New RegExp
	Dim oMatches
	Dim strFullPath, ExcludeFile, strLine, intLast
	strFullPath = strScriptPath + fileExclude
	
	On Error Resume Next
	Set ExcludeFile = oFS.OpenTextFile(strFullPath, ForReading, False, UseSysDefualt)
	If Err.Number <> 0 Then Call ReturnMessage(Err.Number & " - " & Err.Description, OtherExit)
	On Error Goto 0
	
	oRegExp.IgnoreCase = True
	oRegExp.Global = False
	oRegExp.Pattern = "(.*)\t(.*)\t(.*)"
	
	Do While ExcludeFile.AtEndOfStream <> True
		strLine = ExcludeFile.ReadLine
		If (Len(strLine) > 0) And (StrComp(Left(strLine, 1), "#", vbTextCompare) <> 0) Then 
			Set oMatches = oRegExp.Execute(strLine)
			If oMatches.Count <> 1 Then Call ReturnMessage(EMSG_INVEXCFILE, CriticalExit)
			
			intLast = Ubound(arrExclude)
			Set arrExclude(intLast) = New ExcludeItem
			
			arrExclude(intLast).ValueName = oMatches(0).SubMatches(0)
			arrExclude(intLast).Regex = oMatches(0).SubMatches(1)
			arrExclude(intLast).Action = oMatches(0).SubMatches(2)
			
			ReDim Preserve arrExclude(intLast+1)
		End If
	Loop
	ReDim Preserve arrExclude(intLast)
	ExcludeFile.Close
	Set oFS = Nothing
	Set oRegExp = Nothing
	Set oMatches = Nothing
End Sub
'************************************************************************************************
Function checkExcluded(oRegItem, arrExcludeItems(), strRegData)
	Dim oRegExpKey	:	Set oRegExpKey = New RegExp
	Dim oRegExpComputer	:	Set oRegExpComputer = New RegExp
	Dim item, arrAction, strIP
	
	oRegExpKey.IgnoreCase = True
	oRegExpKey.Global = False
	oRegExpComputer.IgnoreCase = True
	oRegExpComputer.Global = False	

	For Each item In arrExcludeItems
		oRegExpKey.Pattern = item.ValueName
		oRegExpComputer.Pattern = item.Regex
		If (oRegExpKey.Test(oRegItem.Key & "\" & oRegItem.Value)) And (oRegExpComputer.Test(strComputerName)) Then
			checkExcluded = 1
			On Error Resume Next
			arrAction = Split(item.Action,":",2)
			If Err.Number <> 0 Then Call ReturnMessage(EMSG_INVEXCFILE, OtherExit)
			On Error Goto 0
			Select Case arrAction(0)
				Case "IGNORE"
					Exit Function
				Case "IP"
					strIP = GetWANIP
					If StrComp(strRegData, strIP, vbTextCompare) <> 0 Then checkExcluded = 2
					Exit Function
				Case "COMPUTERNAME"
					If StrComp(strRegData, strComputerName & arrAction(1), vbTextCompare) <> 0 Then checkExcluded = 2
					Exit Function
				Case "VALUE"
				'todo add handle to non string values (for example if you need to pass array)
					If StrComp(strRegData, arrAction(1), vbTextCompare) <> 0 Then checkExcluded = 2
					Exit Function
				Case Else
			End Select
		End If
	Next
	
	Set oRegExpKey = Nothing
	Set oRegExpComputer = Nothing
	checkExcluded = 0
End Function