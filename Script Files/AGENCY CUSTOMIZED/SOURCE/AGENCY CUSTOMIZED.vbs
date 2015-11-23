'------------------------------------HEADER STARTS HERE------------------------------------

EMConnect ""

ReDim folder_array(0), script_names(0)
Dim dir, dia_width, vert_shift, horza_offset, offset, folder_level, buttonpressed, count_custom_script_library
Dim button_assignment, i

	'Allows users to navigate through the folders in the root directory
Function dir_nav(change)
	Dim colon_finder, dir_root, dir_length, dir_path, dir_spacer, len_space, spaces, current_dir
		'Add a \ to the End of the dir If one is not there
	If Right(dir,1) <> "\" Then dir = dir & "\"
		'Find the current drive location
	colon_finder = InStr(dir,":\")
	dir_length = Len(dir)
	dir_root = Left(dir,colon_finder+1)
		'Remove the drive naming and preserve the spaces using a carrot
	dir_path = replace(replace(dir,dir_root,""), " ", "^")
		'Explode remaining directory to target your current folder
	dir_spacer = Len(dir_path)
	len_space = dir_spacer + 10
	For i = 1 to len_space
		spaces = spaces & " "
	Next
		'This will equal your current folder and remove the carrots Set prior
	current_dir = Replace(Trim(Right(trim(replace(dir_path,Right(dir,1),spaces)),dir_spacer)),"^", " ") & "\"
	If change = "back" Then
		dir = replace(dir,current_dir,"")
		folder_level = folder_level - 1
	Else
		dir = dir & change & "\"
		folder_level = folder_level + 1
	End If
End Function

	'Gathers the vbs files and the folders from a directory
Function folder_contents(dir)
	Dim main_folder, folder, colFiles, folder_list, objFile, objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
		'List folders in dir
	Set main_folder = objFSO.GetFolder(dir)
	Set folder_list = main_folder.SubFolders
	ReDim folder_array(0)
	For Each folder in folder_list
		If folder_array(0) <> "" Then ReDim Preserve folder_array(UBound(folder_array)+1) 
		folder_array(UBound(folder_array)) = folder.name
	Next
		'List files in dir
	Set colFiles = main_folder.Files
	ReDim script_names(0)
	For Each objFile in colFiles
		If right(objFile.Name,4) = ".vbs" Then
			If script_names(0) <> "" Then ReDim Preserve script_names(UBound(script_names)+1)
			script_names(UBound(script_names)) = objFile.Name
		End If
	Next
End Function

	'Resets the dimensions of the dialog box
Sub reset_dialog
	vert_shift = 13
	'Sets the height based on the number of objects counted
	If (UBound(folder_array)+UBound(script_names)+2) > 1 AND (UBound(folder_array)+UBound(script_names)+2) < 25 Then
		vert_shift = ((UBound(folder_array)+UBound(script_names)+2) * 13)
	Elseif (UBound(folder_array)+UBound(script_names)+2) > 24 Then
		vert_shift = 325
	End If
		'Sets the width based on the number of objects counted
	dia_width = 0
	If (UBound(folder_array)+UBound(script_names)+2) > 24 Then dia_width = 153
	If (UBound(folder_array)+UBound(script_names)+2) > 49 Then dia_width = 306
End Sub

	'Builds the dynamic dialogs
Sub build_dialog
	BeginDialog count_custom_script_library, 0, 0, 218 + dia_width, 27 + vert_shift, "County Custom Scripts"
		offset = 3
		horza_offset = 0
		ButtonGroup ButtonPressed 'All buttons are contained in a single ButtonGroup
				'List Folders
			button_assignment = 10
			If folder_array(LBound(folder_array)) <> "" Then
				For i = LBound(folder_array) to UBound(folder_array)
					PushButton 3 + horza_offset, offset, 25, 11, "Select", button_assignment
					Text 30 + horza_offset, 1 + offset, 120, 10, folder_array(i)
					offset = offset + 13
					button_assignment = button_assignment + 1
					If i = 24 Then 
						horza_offset = 153
						offset = 3
					Elseif i = 49 Then
						horza_offset = 306
						offset = 3
					End If
				Next
			End If
				'List Scripts		
			button_assignment = 61
			If script_names(LBound(script_names)) <> "" Then
				For i = LBound(script_names) to UBound(script_names)
					PushButton 3 + horza_offset, offset, 25, 11, "Run", button_assignment
					Text 30 + horza_offset, 1 + offset, 120, 10, left(script_names(i),(len(script_names(i)) - 4)) 'Removes the file extension when showing the name in the dialog
					offset = offset + 13
					button_assignment = button_assignment + 1
					If (UBound(folder_array) + i) = 24 Then 
						horza_offset = 153
						offset = 3
					Elseif (UBound(folder_array) + i) = 49 Then
						horza_offset = 306
						offset = 3
					End If
				Next
			End If
			CancelButton 188 + horza_offset, 13 + vert_shift, 28, 12
			If folder_level <> 0 Then PushButton 140 + horza_offset, 13 + vert_shift, 45, 12, "Back Folder", -10
	EndDialog
End Sub

'------------------------------------End OF HEADER------------------------------------

	'This sets the default directory and makes it the navigation root
dir = default_directory & "AGENCY CUSTOMIZED\"
folder_level = 0

Do 
	Call folder_contents(dir)
	reset_dialog
	build_dialog
	dialog count_custom_script_library
		If buttonpressed = 0 Then stopscript
		If buttonpressed = -10 Then dir_nav("back")
	For i = 10 to 110
		If buttonpressed = i Then 	
			If i < 61 Then 
				Call dir_nav(folder_array(i-10))
			Elseif i > 60 Then 
				Dim county_script_run, fso_command_crs, count_specific_script
				Set county_script_run = CreateObject("Scripting.FileSystemObject")
				If Right(dir,1) <> "\" Then dir = dir & "\"
				Set fso_command_crs = county_script_run.OpenTextFile(dir&script_names(i-61))
				count_specific_script = fso_command_crs.ReadAll
				fso_command_crs.Close
				Execute count_specific_script
				stopscript
			End If
		End If
	Next
Loop

stopscript
