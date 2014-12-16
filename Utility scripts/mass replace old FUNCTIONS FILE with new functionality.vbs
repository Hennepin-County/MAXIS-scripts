section = "NAV"		'Should be NAV, NOTES, BULK, ACTIONS, MEMOS, DAIL

'Puts the new header into an array, to be used later
new_header_array = Array("url = ""https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs""", _
	"SET req = CreateObject(""Msxml2.XMLHttp.6.0"")				'Creates an object to get a URL", _
	"req.open ""GET"", url, FALSE									'Attempts to open the URL", _
	"req.send													'Sends request", _
	"IF req.Status = 200 THEN									'200 means great success", _
	"	Set fso = CreateObject(""Scripting.FileSystemObject"")	'Creates an FSO", _
	"	Execute req.responseText								'Executes the script code", _
	"ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).", _
	"	MsgBox 	""Something has gone wrong. The code stored on GitHub was not able to be reached."" & vbCr &_ ", _
	"			vbCr & _", _
	"			""Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com."" & vbCr &_", _
	"			vbCr & _", _
	"			""If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:"" & vbCr &_", _
	"			vbTab & ""- The name of the script you are running."" & vbCr &_", _
	"			vbTab & ""- Whether or not the script is """"erroring out"""" for any other users."" & vbCr &_", _
	"			vbTab & ""- The name and email for an employee from your IT department,"" & vbCr & _", _
	"			vbTab & vbTab & ""responsible for network issues."" & vbCr &_", _
	"			vbTab & ""- The URL indicated below (a screenshot should suffice)."" & vbCr &_", _
	"			vbCr & _", _
	"			""Veronica will work with your IT department to try and solve this issue, if needed."" & vbCr &_ ", _
	"			vbCr &_", _
	"			""URL: "" & url", _
	"			script_end_procedure(""Script ended due to error connecting to GitHub."")", _
	"END IF")

'This is the old header. Later it will apply a filter to see if a specific line matches any of this content.
old_header_array = Array("Set run_another_script_fso = CreateObject(""Scripting.FileSystemObject"")", _
	"Set fso_command = run_another_script_fso.OpenTextFile(""C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs"")",_
	"text_from_the_other_script = fso_command.ReadAll",_
	"fso_command.Close",_
	"Execute text_from_the_other_script")

'Sets some constants for later
Const ForReading = 1
Const ForWriting = 2

'sFolder is the folder I'm currently working on. 
sFolder = "C:\DHS-MAXIS-Scripts\Script Files\" & section & "\"

'Creates an FSO
Set oFSO = CreateObject("Scripting.FileSystemObject")

'For each file in the folder, it will do the following.
For Each oFile In oFSO.GetFolder(sFolder).Files
	If UCase(oFSO.GetExtensionName(oFile.Name)) = "VBS" Then
		'This MsgBox was inserted for easy testing. It should be removed to make this run very quickly. 
		display = MsgBox(replace(oFile, sFolder, ""), vbOKCancel)
		If display = vbCancel then wscript.quit
		
		'Opens the file
		Set oFile2 = oFSO.OpenTextFile(oFile.path, ForReading)
		
		'Do the magic
		Do until oFile2.AtEndOfStream
			'Declares the strline
			strLine = oFile2.ReadLine
			'Replaces "name_of_script" with the filename, as that'll be VERY VERY helpful later on when running statistics
			If left(strLine, 14) = "name_of_script" then strLine = "name_of_script = """ & replace(oFile, sFolder, "") & """"
			
			'If it finds the routine functions, it inserts the array from above
			If instr(strLine, "LOADING ROUTINE FUNCTIONS") <> 0 then
				strLine = "'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------"
				For each hLine in new_header_array
					strLine = strLine & vbCrLf & hLine 
				Next
			End if
			filter_results = Filter(old_header_array, strLine)								'Checks to see if the line is in the original header
			If ubound(filter_results) <> 0 then strText = strText & strLine & vbCrLf		'Only writes the line if it isn't part of the original header (skips original header)
		Loop
		
		'Sets a new path variable for later, closes current file and clears the variable
		sFile = oFile.path
		oFile2.Close
		set oFile2 = nothing
		
		'If it's a NAV script, it creates redirects. Here's an array with the contents.
		If left(replace(oFile, sFolder, ""), 3) = "NAV" then 
			redirect_array = Array("'LOADING GLOBAL VARIABLES--------------------------------------------------------------------", _
				"Set run_another_script_fso = CreateObject(""Scripting.FileSystemObject"")", _
				"Set fso_command = run_another_script_fso.OpenTextFile(""C:\DHS-MAXIS-Scripts\Script Files\SETTINGS - GLOBAL VARIABLES.vbs"")", _
				"text_from_the_other_script = fso_command.ReadAll", _
				"fso_command.Close", _
				"Execute text_from_the_other_script", _
				"", _
				"'LOADING SCRIPT", _
				"url = script_repository & ""/" & section & "/" & replace(oFile, sFolder, "") & """", _
				"SET req = CreateObject(""Msxml2.XMLHttp.6.0"")				'Creates an object to get a URL", _
				"req.open ""GET"", url, FALSE									'Attempts to open the URL", _
				"req.send													'Sends request", _
				"IF req.Status = 200 THEN									'200 means great success", _
				"	Set fso = CreateObject(""Scripting.FileSystemObject"")	'Creates an FSO", _
				"	Execute req.responseText								'Executes the script code", _
				"ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).", _
				"	MsgBox 	""Something has gone wrong. The code stored on GitHub was not able to be reached."" & vbCr &_ ", _
				"			vbCr & _", _
				"			""Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com."" & vbCr &_", _
				"			vbCr & _", _
				"			""If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:"" & vbCr &_", _
				"			vbTab & ""- The name of the script you are running."" & vbCr &_", _
				"			vbTab & ""- Whether or not the script is """"erroring out"""" for any other users."" & vbCr &_", _
				"			vbTab & ""- The name and email for an employee from your IT department,"" & vbCr & _", _
				"			vbTab & vbTab & ""responsible for network issues."" & vbCr &_", _
				"			vbTab & ""- The URL indicated below (a screenshot should suffice)."" & vbCr &_", _
				"			vbCr & _", _
				"			""Veronica will work with your IT department to try and solve this issue, if needed."" & vbCr &_ ", _
				"			vbCr &_", _
				"			""URL: "" & url", _
				"			script_end_procedure(""Script ended due to error connecting to GitHub."")", _
				"END IF")
			
			'Creates a redirect script if the script is a NAV script. Stole some of the code from @RobertFewins-Kalb.
			SET create_redirect_file_fso = CreateObject("Scripting.FileSystemObject")
			SET create_redirect_file_command = create_redirect_file_fso.CreateTextFile("C:\DHS-MAXIS-Scripts\Script Files\REDIRECT - " & replace(oFile, sFolder, ""), 2)
			For each redirect_line in redirect_array
				create_redirect_file_command.Write(redirect_line & vbCrLf)
			Next
			create_redirect_file_command.Close
		End if
		
		'Writes new version of the file
		Set oFile = oFSO.OpenTextFile(sFile ,  ForWriting)
		oFile.Write strText
		oFile.Close
		
		'Clears variables
		Set oFile = Nothing
		strText = ""
		
	End if
Next
