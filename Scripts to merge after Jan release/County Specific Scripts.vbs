EMConnect ""

Set objFSO = CreateObject("Scripting.FileSystemObject")
	'------------------ THIS IS THE ONLY THING THAT SHOULD BE CHANGED PER COUNTY ------------
objStartFolder = "C:\Users\shanleyl\Desktop\Scripts"
	'----------------------------------------------------------------------------------------
	
Set objFolder = objFSO.GetFolder(objStartFolder)
Set colFiles = objFolder.Files

file_count = 0

	'Creates an array to detect what is pressed at the end
Dim checked_scripts(100)
Dim script_names(100)
script_number = 0

For Each objFile in colFiles
	if right(objFile.Name,4) = ".vbs" then
		checked_scripts(script_number) = objFile.Name
		script_names(script_number) = objFile.Name
		script_number = script_number + 1
		file_count = file_count + 1
	end if
Next

vert_shift = 13

If file_count > 1 then
	vert_shift = (file_count * 13)
end If

offset = 3
on_script = 0

BeginDialog county_script_library, 0, 0, 218, 27 + vert_shift, "County Custom Scripts"
  	'Script Pages Here
	For Each objFile in colFiles
			'Only Lists .vbs files
		if right(objFile.Name,4) = ".vbs" then
				'Removes .vbs from the title to clean up the naming
			file_type_remo = len(objFile.Name) - 4
			script_title = left(objFile.Name,file_type_remo)
				'Creates checkbox
			CheckBox 3, offset, 150, 10, script_title, checked_scripts(on_script)
				'Changes Checkbox offset for next script
			offset = offset + 13
			on_script = on_script + 1
		end if
	next		
	ButtonGroup ButtonPressed
    OkButton 197, 13 + vert_shift, 19, 12
    CancelButton 166, 13 + vert_shift, 28, 12
EndDialog

dialog county_script_library
	if buttonpressed = 0 then stopscript

	'use checked_scripts array to scan
check_script = 0
number_selected = 0

if buttonpressed = -1 then
	For i = 1 to file_count
			if checked_scripts(check_script) = 1 then
					number_selected = number_selected + 1				
			end if
			check_script = check_script + 1
	Next	
	if number_selected > 1 then
		msgbox "You must only select 1 script at a time. You are receiving this error because you selected more that one script. Please choose again.","Script Selection Error"
	end if
	if number_selected = 0 then
		msgbox "You must a script. You are receiving this error because script were chosen to be run. Please make a selection again.","Script Selection Error"
	end if
	if number_selected = 1 then	
		check_script = 0
		For i = 1 to file_count
			if checked_scripts(check_script) = 1 then
				Set county_script_run = CreateObject("Scripting.FileSystemObject")
				Set fso_command_crs = county_script_run.OpenTextFile(objStartFolder&"\"&script_names(check_script))
				count_specific_script = fso_command_crs.ReadAll
				fso_command_crs.Close
				Execute count_specific_script
			end if
			check_script = check_script + 1
		Next
	end if
end if

stopscript 
