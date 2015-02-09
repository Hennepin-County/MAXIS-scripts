EMConnect ""

Set objFSO = CreateObject("Scripting.FileSystemObject")
	'------------------ THIS IS THE ONLY THING THAT SHOULD BE CHANGED PER COUNTY ------------
	objStartFolder = default_directory & "AGENCY CUSTOMIZED\"
	'----------------------------------------------------------------------------------------

Public folder_array(100), dir, objCount, objFolders, objScripts, script_number, checked_scripts(100), script_names(100), checked_folders(100), i, objFile, main_folder, folder_list, colFiles, dia_width, vert_shift
Public horza_offset, on_item, offset, on_button, buttonpressed, folder_level, file_count

dir = objStartFolder
folder_level = 0

'Changes the directory
function dir_nav(change)
	Dim colon_finder, dir_root, dir_length, dir_path, dir_spacer, len_space, spaces, current_dir 
		'Add a \ to the end of the dir if one is not there
	If Right(dir,1) <> "\" then dir = dir & "\"
		'Find the current drive location
	colon_finder = InStr(dir,":\")
	dir_length = Len(dir)
	dir_root = Left(dir,colon_finder+1)
		'Remove the drive naming and preserve the spaces using a carrot
	dir_path = replace(replace(dir,dir_root,""), " ", "^")
		'Explode remaining directory to target your current folder
	dir_spacer = Len(dir_path)
	len_space = dir_spacer + 10
	for i = 1 to len_space
		spaces = spaces & " "
	next
		'This will equal your current folder and remove the carrots set prior
	current_dir = Replace(Trim(Right(trim(replace(dir_path,Right(dir,1),spaces)),dir_spacer)),"^", " ") & "\"
	if change = "back" then
		dir = replace(dir,current_dir,"")
		folder_level = folder_level - 1
	else
		dir = dir & change & "\"
		folder_level = folder_level + 1
	end if
end function

'Sets the dimensions of the dialog box
function reset_dialog()
	vert_shift = 13
	'Sets the height based on the number of objects counted
	If objCount > 1 AND objCount < 25 then
		vert_shift = (objCount * 13)
	elseif objCount > 24 then
		vert_shift = 325
	end If
	
		'Sets the width based on the number of objects counted
	dia_width = 0
	If objCount > 24 then dia_width = 153
	If objCount > 49 then dia_width = 306	
End Function

Function folder_contents(dir)
	objCount = 0
		'List folders in dir
	Dim main_folder, folder, folder_name
	set main_folder = objFSO.GetFolder(dir)
	set folder_list = main_folder.SubFolders
	Erase checked_folders
	folder_name = 0
	For Each folder in folder_list
		folder_array(folder_name) = folder.name
		folder_name = folder_name + 1
		objCount = objCount + 1
	Next
	
	'List files in dir
	Set colFiles = main_folder.Files
	objScripts = 0
		'Cleans the arrays
	Erase	checked_scripts
	Erase	script_names
	script_number = 0	
	file_count = 0
	For Each objFile in colFiles
		if right(objFile.Name,4) = ".vbs" then
			checked_scripts(script_number) = objFile.Name
			script_names(script_number) = objFile.Name
			script_number = script_number + 1
			objCount = objCount + 1
			file_count = file_count + 1
		end if
	Next
End Function

Function main_dialog()
	BeginDialog county_script_library, 0, 0, 218 + dia_width, 27 + vert_shift, "County Custom Scripts"
		offset = 3
		on_item = 0
		horza_offset = 0
		'List Folders
		ButtonGroup folderselected
		on_button = 10
		For Each folder in folder_list
			PushButton 3 + horza_offset, offset, 25, 10, "Select", on_button
			Text 30 + horza_offset, 1 + offset, 120, 10, folder.name
			offset = offset + 13
			on_item = on_item + 1
			on_button = on_button + 1
			if on_item = 25 then 
				horza_offset = 153
				offset = 3
			elseif on_item = 50 then
				horza_offset = 306
				offset = 3
			end if
		Next
			
		on_item = 0
			OptionGroup RadioGroup1
		'Script Pages Here
		For Each objFile in colFiles
				'Only Lists .vbs files
			if right(objFile.Name,4) = ".vbs" then
					'Removes .vbs from the title to clean up the naming
				file_type_remo = len(objFile.Name) - 4
				script_title = left(objFile.Name,file_type_remo)
					'Creates Radio Buttons for script files
					RadioButton 3 + horza_offset, offset, 150, 10, script_title, checked_scripts(on_item)
					'Changes Radio Button offset for next script
				offset = offset + 13
				on_item = on_item + 1
				if on_item = 25 then 
					horza_offset = 153
					offset = 3
				elseif on_item = 50 then
					horza_offset = 306
					offset = 3
				end if
			end if
		next		
		ButtonGroup ButtonPressed
			OkButton 197 + horza_offset, 13 + vert_shift, 19, 12
			CancelButton 166 + horza_offset, 13 + vert_shift, 28, 12
			If folder_level <> 0 then PushButton 118 + horza_offset, 13 + vert_shift, 45, 12, "Back Folder", -10
	EndDialog
End Function

Do 
	call folder_contents(dir)
	call reset_dialog()
	call main_dialog
	dialog county_script_library
		if buttonpressed = 0 then stopscript
		'use checked_scripts array to scan 
		if buttonpressed = -10 then dir_nav("back")
	for i = 10 to 110
		if buttonpressed = i then 
			call dir_nav(folder_array(i-10))
		end if
	next
Loop until buttonpressed = -1

check_script = 0
number_selected = 0

if buttonpressed = -1 then
	For i = 1 to file_count
			if checked_scripts(check_script) = 1 then
					number_selected = number_selected + 1				
			end if
			check_script = check_script + 1
	Next	
	check_script = 0
	For i = 1 to file_count
		if checked_scripts(check_script) = 1 then
			Set county_script_run = CreateObject("Scripting.FileSystemObject")
				if Right(dir,1) <> "\" then dir = dir & "\"
			Set fso_command_crs = county_script_run.OpenTextFile(dir&script_names(check_script))
			count_specific_script = fso_command_crs.ReadAll
			fso_command_crs.Close
			Execute count_specific_script
		end if
		check_script = check_script + 1
	Next
end if

stopscript
