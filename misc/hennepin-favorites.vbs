'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "Hennepin Favorites.vbs"
start_time = timer
run_locally = TRUE
'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================
function edit_favorites

	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>> SECTION 1 <<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>> The gobbins that happen before the user sees anything. <<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


	'Warning/instruction box
	MsgBox  "This section will display a dialog with various scripts on it. Any script you check will be added to your favorites menu. Scripts you un-check will be removed. Once you are done making your selection hit ""OK"" and your menu will be updated. " & vbNewLine & vbNewLine &_
			"Note: you will be unable to edit the list of new scripts."

	REDIM scripts_edit_favs_array(ubound(script_array), script_hotkey)

	'determining the number of each kind of script...by category
	number_of_scripts = 0
	actions_scripts = 0
	bulk_scripts = 0
	notc_scripts = 0
	notes_scripts = 0
	utilities_scripts = 0
	admin_scripts = 0
	qi_scripts = 0
	bz_scripts = 0
	FOR i = 0 TO ubound(script_array)
		number_of_scripts = i
		IF script_array(i).category = "ACTIONS" THEN
			actions_scripts = actions_scripts + 1
		ELSEIF script_array(i).category = "BULK" THEN
			bulk_scripts = bulk_scripts + 1
		ELSEIF script_array(i).category = "NOTICES" THEN
			notc_scripts = notc_scripts + 1
		ELSEIF script_array(i).category = "NOTES" THEN
			notes_scripts = notes_scripts + 1
		ELSEIF script_array(i).category = "UTILTIES" THEN
	        utilities_scripts = utilities_scripts + 1
		ELSEIF script_array(i).category = "ADMIN" THEN
			For each review_group in script_array(i).tags
				If review_group = "" Then admin_scripts = admin_scripts + 1
				If review_group = "QI" Then qi_scripts = qi_scripts + 1
				If review_group = "BZ" Then bz_scripts = bz_scripts + 1
			Next
	    End if
	NEXT

	'>>> If the user has already selected their favorites, the script will open that file and
	'>>> and read it, storing the contents in the variable name ''favorites_text_file_array''
	SET oTxtFile = (CreateObject("Scripting.FileSystemObject"))
	With oTxtFile
		If .FileExists(favorites_text_file_location) THEN
            If .GetFile(favorites_text_file_location).size <> 0 Then
    			Set fav_scripts = CreateObject("Scripting.FileSystemObject")
    			Set fav_scripts_command = fav_scripts.OpenTextFile(favorites_text_file_location)
    			fav_scripts_array = fav_scripts_command.ReadAll
    			IF fav_scripts_array <> "" THEN favorites_text_file_array = fav_scripts_array
    			fav_scripts_command.Close
            End If
		END IF
	END WITH

	'>>> Determining the width of the dialog from the number of scripts that are available...
	'the dialog starts with a width of 400
	dia_width = 750

	'>>> Building the dialog
    Dialog1 = ""
	BeginDialog Dialog1, 0, 0, dia_width, 380, "Select your favorites"
		ButtonGroup ButtonPressed
			OkButton 5, 5, 50, 15
			CancelButton 55, 5, 50, 15
			PushButton 165, 5, 70, 15, "Reset Favorites", reset_favorites_button
		'>>> Creating the display of all scripts for selection (in checkbox form)
		script_position = 0		' <<< This value is tied to the number_of_scripts variable


		col = 10
		row = 30
		Text col, row, 175, 10, "---------- ACTIONS SCRIPTS ----------"
		row = row + 10

		FOR i = 0 to ubound(script_array)
			IF script_array(i).category = "ACTIONS" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 360 THEN
					col = col + 125
					If col > 250 Then
						row = 10
					Else
						row = 30
					End If
				END IF
				'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(favorites_text_file_array, "-", " ")), UCASE(replace(script_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, add_checkbox) = checked
				ELSE
					scripts_edit_favs_array(script_position, add_checkbox) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, fav_script_name)   = script_array(i).script_name
				' scripts_edit_favs_array(script_position, proper_name)       = script_array(i).file_name
				scripts_edit_favs_array(script_position, cat_as_direct)     = script_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 120, 10, scripts_edit_favs_array(script_position, fav_script_name), scripts_edit_favs_array(script_position, add_checkbox)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 10	'Padding for the new section
		'Account for overflow
		IF row >= 360 THEN
			col = col + 125
			If col > 250 Then
				row = 10
			Else
				row = 30
			End If
		END IF
		Text col, row, 175, 10, "---------- BULK SCRIPTS ----------"
		row = row + 10

		'BULK script laying out
		FOR i = 0 to ubound(script_array)
			IF script_array(i).category = "BULK" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 360 THEN
					col = col + 125
					If col > 250 Then
						row = 10
					Else
						row = 30
					End If
				END IF
                '>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(favorites_text_file_array, "-", " ")), UCASE(replace(script_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, add_checkbox) = checked
				ELSE
					scripts_edit_favs_array(script_position, add_checkbox) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, fav_script_name)   = script_array(i).script_name
				' scripts_edit_favs_array(script_position, proper_name)       = script_array(i).file_name
				scripts_edit_favs_array(script_position, cat_as_direct)     = script_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 120, 10, scripts_edit_favs_array(script_position, fav_script_name), scripts_edit_favs_array(script_position, add_checkbox)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 10	'Padding for the new section
		'Account for overflow
		IF row >= 360 THEN
			col = col + 125
			If col > 250 Then
				row = 10
			Else
				row = 30
			End If
		END IF
		Text col, row, 175, 10, "---------- NOTICES SCRIPTS ----------"
		row = row + 10

		'CALCULATOR script laying out
		FOR i = 0 to ubound(script_array)
			IF script_array(i).category = "NOTICES" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 360 THEN
					col = col + 125
					If col > 250 Then
						row = 10
					Else
						row = 30
					End If
				END IF
                '>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(favorites_text_file_array, "-", " ")), UCASE(replace(script_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, add_checkbox) = checked
				ELSE
					scripts_edit_favs_array(script_position, add_checkbox) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, fav_script_name)   = script_array(i).script_name
				' scripts_edit_favs_array(script_position, proper_name)       = script_array(i).file_name
				scripts_edit_favs_array(script_position, cat_as_direct)     = script_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 120, 10, scripts_edit_favs_array(script_position, fav_script_name), scripts_edit_favs_array(script_position, add_checkbox)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 10	'Padding for the new section
		'Account for overflow
		IF row >= 360 THEN
			col = col + 125
			If col > 250 Then
				row = 10
			Else
				row = 30
			End If
		END IF
		Text col, row, 175, 10, "---------- NOTES SCRIPTS ----------"
		row = row + 10

		'NOTES script laying out
		FOR i = 0 to ubound(script_array)
			IF script_array(i).category = "NOTES" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 360 THEN
					col = col + 125
					If col > 250 Then
						row = 10
					Else
						row = 30
					End If
				END IF
                '>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(favorites_text_file_array, "-", " ")), UCASE(replace(script_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, add_checkbox) = checked
				ELSE
					scripts_edit_favs_array(script_position, add_checkbox) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, fav_script_name)   = script_array(i).script_name
				' scripts_edit_favs_array(script_position, proper_name)       = script_array(i).file_name
				scripts_edit_favs_array(script_position, cat_as_direct)     = script_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 120, 10, scripts_edit_favs_array(script_position, fav_script_name), scripts_edit_favs_array(script_position, add_checkbox)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 10	'Padding for the new section
		'Account for overflow
		IF row >= 360 THEN
			col = col + 125
			If col > 250 Then
				row = 10
			Else
				row = 30
			End If
		END IF
		Text col, row, 175, 10, "---------- UTILITIES SCRIPTS ----------"
		row = row + 10

		'UTILITIES script laying out
	    FOR i = 0 to ubound(script_array)
			IF script_array(i).category = "UTILITIES" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 360 THEN
					col = col + 125
					If col > 250 Then
						row = 10
					Else
						row = 30
					End If
				END IF
                '>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
				IF InStr(UCASE(replace(favorites_text_file_array, "-", " ")), UCASE(replace(script_array(i).script_name, "-", " "))) <> 0 THEN
					scripts_edit_favs_array(script_position, add_checkbox) = checked
				ELSE
					scripts_edit_favs_array(script_position, add_checkbox) = unchecked
				END IF

				'Sets the file name and category
				scripts_edit_favs_array(script_position, fav_script_name)   = script_array(i).script_name
				' scripts_edit_favs_array(script_position, proper_name)       = script_array(i).file_name
				scripts_edit_favs_array(script_position, cat_as_direct)     = script_array(i).category & "/"

				'Displays the checkbox
				CheckBox col, row, 120, 10, scripts_edit_favs_array(script_position, fav_script_name), scripts_edit_favs_array(script_position, add_checkbox)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 10	'Padding for the new section
		'Account for overflow
		IF row >= 360 THEN
			col = col + 125
			If col > 250 Then
				row = 10
			Else
				row = 30
			End If
		END IF
		Text col, row, 175, 10, "---------- ADMIN SCRIPTS ----------"
		row = row + 10

		'ADMIN script laying out
	    FOR i = 0 to ubound(script_array)
			IF script_array(i).category = "ADMIN" THEN
				For each review_group in script_array(i).tags
					If review_group = "" Then
						'>>> Determining the positioning of the checkboxes.
						'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
						IF row >= 360 THEN
							col = col + 125
							If col > 250 Then
								row = 10
							Else
								row = 30
							End If
						END IF
		                '>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
						IF InStr(UCASE(replace(favorites_text_file_array, "-", " ")), UCASE(replace(script_array(i).script_name, "-", " "))) <> 0 THEN
							scripts_edit_favs_array(script_position, add_checkbox) = checked
						ELSE
							scripts_edit_favs_array(script_position, add_checkbox) = unchecked
						END IF

						'Sets the file name and category
						scripts_edit_favs_array(script_position, fav_script_name)   = script_array(i).script_name
						' scripts_edit_favs_array(script_position, proper_name)       = script_array(i).file_name
						scripts_edit_favs_array(script_position, cat_as_direct)     = script_array(i).category & "/"

						'Displays the checkbox
						CheckBox col, row, 120, 10, scripts_edit_favs_array(script_position, fav_script_name), scripts_edit_favs_array(script_position, add_checkbox)

						'Increments the row and script_position
						row = row + 10
						script_position = script_position + 1
					End If
				Next
			END IF
		NEXT

		If qi_staff = TRUE Then
			'Section header
			row = row + 10	'Padding for the new section
			'Account for overflow
			IF row >= 360 THEN
				col = col + 125
				If col > 250 Then
					row = 10
				Else
					row = 30
				End If
			END IF
			Text col, row, 175, 10, "---------- QI SCRIPTS ----------"
			row = row + 10

			'ADMIN script laying out
		    FOR i = 0 to ubound(script_array)
				IF script_array(i).category = "ADMIN" THEN
					For each review_group in script_array(i).tags
						If review_group = "QI" Then
							'>>> Determining the positioning of the checkboxes.
							'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
							IF row >= 360 THEN
								col = col + 125
								If col > 250 Then
									row = 10
								Else
									row = 30
								End If
							END IF
			                '>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
							IF InStr(UCASE(replace(favorites_text_file_array, "-", " ")), UCASE(replace(script_array(i).script_name, "-", " "))) <> 0 THEN
								scripts_edit_favs_array(script_position, add_checkbox) = checked
							ELSE
								scripts_edit_favs_array(script_position, add_checkbox) = unchecked
							END IF

							'Sets the file name and category
							scripts_edit_favs_array(script_position, fav_script_name)   = script_array(i).script_name
							' scripts_edit_favs_array(script_position, proper_name)       = script_array(i).file_name
							scripts_edit_favs_array(script_position, cat_as_direct)     = script_array(i).category & "/"

							'Displays the checkbox
							CheckBox col, row, 120, 10, scripts_edit_favs_array(script_position, fav_script_name), scripts_edit_favs_array(script_position, add_checkbox)

							'Increments the row and script_position
							row = row + 10
							script_position = script_position + 1
						End If
					Next
				END IF
			NEXT
		End If

		If bz_staff = TRUE Then
			'Section header
			row = row + 10	'Padding for the new section
			'Account for overflow
			IF row >= 360 THEN
				col = col + 125
				If col > 250 Then
					row = 10
				Else
					row = 30
				End If
			END IF
			Text col, row, 175, 10, "---------- BZST SCRIPTS ----------"
			row = row + 10

			'ADMIN script laying out
			FOR i = 0 to ubound(script_array)
				IF script_array(i).category = "ADMIN" THEN
					For each review_group in script_array(i).tags
						If review_group = "BZ" Then
							'>>> Determining the positioning of the checkboxes.
							'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
							IF row >= 360 THEN
								col = col + 125
								If col > 250 Then
									row = 10
								Else
									row = 30
								End If
							END IF
							'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
							IF InStr(UCASE(replace(favorites_text_file_array, "-", " ")), UCASE(replace(script_array(i).script_name, "-", " "))) <> 0 THEN
								scripts_edit_favs_array(script_position, add_checkbox) = checked
							ELSE
								scripts_edit_favs_array(script_position, add_checkbox) = unchecked
							END IF

							'Sets the file name and category
							scripts_edit_favs_array(script_position, fav_script_name)   = script_array(i).script_name
							' scripts_edit_favs_array(script_position, proper_name)       = script_array(i).file_name
							scripts_edit_favs_array(script_position, cat_as_direct)     = script_array(i).category & "/"

							'Displays the checkbox
							CheckBox col, row, 120, 10, scripts_edit_favs_array(script_position, fav_script_name), scripts_edit_favs_array(script_position, add_checkbox)

							'Increments the row and script_position
							row = row + 10
							script_position = script_position + 1
						End If
					Next
				END IF
			NEXT
		End If

	EndDialog

	'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>> SECTION 2 <<<<<<<<<<<<<<<<<<<<<
	'>>> The gobbins that the user sees and makes do. <<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<

	DO
		DO
			'>>> Running the dialog
            DIALOG Dialog1

			'>>> Cancel confirmation
			IF ButtonPressed = 0 THEN
				confirm_cancel = MsgBox("Are you sure you want to cancel? Press YES to cancel the script. Press NO to return to the script.", vbYesNo)
				IF confirm_cancel = vbYes THEN StopScript
			END IF
			'>>> If the user selects to reset their favorites selections, the script
			'>>> will go through the multi-dimensional array and reset all the values
			'>>> for position 1, thereby clearing the favorites from the display.
			IF ButtonPressed = reset_favorites_button THEN
				FOR i = 0 to number_of_scripts
					scripts_edit_favs_array(i, add_checkbox) = unchecked
				NEXT
			END IF
		'>>> The exit condition for the first do/loop is the user pressing 'OK'
		LOOP UNTIL ButtonPressed <> 0 AND ButtonPressed <> reset_favorites_button
		'>>> Validating that the user does not select more than a prescribed number of scripts.
		'>>> Exceeding the limit will cause an exception access violation for the Favorites script when it runs.
		'>>> Currently, that value is 30. That is lower than previous because of the larger number of new scripts. (-Robert, 04/20/2016)

		double_check_array = ""
		FOR i = 0 to number_of_scripts
			IF scripts_edit_favs_array(i, add_checkbox) = checked THEN double_check_array = double_check_array & scripts_edit_favs_array(i, fav_script_name) & "~"
		NEXT
		double_check_array = split(double_check_array, "~")
		IF ubound(double_check_array) > 24 THEN MsgBox "Your favorites menu is too large. Please limit the number of favorites to no greater than 25."
		'>>> Exit condition is the user having fewer than 30 scripts in their favorites menu.
	LOOP UNTIL ubound(double_check_array) <= 25
    ' dialog1 = "1 - " & dialog1
	'>>> Getting ready to write the user's selection to a text file and save it on a prescribed location on the network.
	'>>> Building the content of the text file.
	FOR i = 0 to number_of_scripts - 1
		IF scripts_edit_favs_array(i, add_checkbox) = checked THEN favorite_scripts = favorite_scripts & scripts_edit_favs_array(i, cat_as_direct) & scripts_edit_favs_array(i, fav_script_name) & vbNewLine
	NEXT

	'>>> After the user selects their favorite scripts, we are going to write (or overwrite) the list of scripts
	'>>> stored at H:\my favorite scripts.txt.
	IF favorite_scripts <> "" THEN
		SET updated_fav_scripts_fso = CreateObject("Scripting.FileSystemObject")
		SET updated_fav_scripts_command = updated_fav_scripts_fso.CreateTextFile(favorites_text_file_location, 2)
		updated_fav_scripts_command.Write(favorite_scripts)
		updated_fav_scripts_command.Close
		MsgBox "Success!! Your Favorites Menu has been updated. Please click your favorites list button to re-load them."
		StopScript
		' script_end_procedure("Success!! Your Favorites Menu has been updated. Please click your favorites list button to re-load them.")
	ELSE
		'>>> OR...if the user has selected no scripts for their favorite, the file will be deleted to
		'>>> prevent the Favorites Menu from erroring out.
		'>>> Experience with worker_signature automation tells us that if the text file is blank, the favorites menu doth not work.
		oTxtFile.DeleteFile(favorites_text_file_location)
		MsgBox "You have updated your Favorites Menu, but you haven't selected any scripts. The next time you use the Favorites scripts, you will need to select your favorites."
		StopScript
		' script_end_procedure("You have updated your Favorites Menu, but you haven't selected any scripts. The next time you use the Favorites scripts, you will need to select your favorites.")
	END IF

end function


'>>> Custom function that builds the Favorites Main Menu dialog.
'>>> the array of the user's scripts
FUNCTION favorite_menu(favorites_text_file_string, script_to_run)
	'>>> Splitting the array of all scripts.
    favorites_text_file_array = split(favorites_text_file_string, vbNewLine)


	num_of_scripts = num_of_user_scripts + num_of_new_scripts

    script_counter = 0
    for each script_data_from_complete_list in script_array
        FOR EACH script_path IN favorites_text_file_array
            script_path = trim(script_path)
            divider_position = InStr(script_path, "/")
            favorite_script_category = left(script_path, divider_position - 1)
            favorite_script_name = right(script_path, len(script_path) - divider_position)

            If favorite_script_category = script_data_from_complete_list.category AND favorite_script_name = script_data_from_complete_list.script_name Then
                ReDim Preserve favorite_scripts_array(script_counter)
                SET favorite_scripts_array(script_counter) = NEW script_bowie
                favorite_scripts_array(script_counter).script_name		= script_data_from_complete_list.script_name
                favorite_scripts_array(script_counter).category			= script_data_from_complete_list.category
                favorite_scripts_array(script_counter).description		= script_data_from_complete_list.description
                favorite_scripts_array(script_counter).release_date		= script_data_from_complete_list.release_date
                favorite_scripts_array(script_counter).tags		        = script_data_from_complete_list.tags
                favorite_scripts_array(script_counter).dlg_keys		    = script_data_from_complete_list.dlg_keys
                favorite_scripts_array(script_counter).keywords		    = script_data_from_complete_list.keywords
                favorite_scripts_array(script_counter).hot_topic_date   = script_data_from_complete_list.hot_topic_date
                script_counter = script_counter + 1
            End If
        Next
    Next


	dlg_height = 0
    dlg_height = dlg_height + 15 + ((ubound(favorite_scripts_array, 1) + 1) * 12)

	'>>> Adjusting the height if the user has fewer scripts selected (in the left column) than what is "new" (in the right column).
	dlg_height = dlg_height + 15 + (12 * num_of_new_scripts)
    dlg_height = dlg_height + 15 + (12 * num_featured_scripts)

	'>>> A nice decoration for the user. If they have used Update Worker Signature, then their signature is built into the dialog display.
	IF worker_full_name <> "" THEN
		dlg_name = worker_full_name & "'s Favorite Scripts"
	ELSEIF worker_signature <> "" THEN
		dlg_name = worker_signature & "'s Favorite Scripts"
	ELSE
		dlg_name = "My Favorite Scripts"
	END IF

	'>>> The dialog
    ' MsgBox dlg_height & " - 4"
	' MsgBox "2 - " & dialog1
	' dialog1 = "2 - " & dialog1
    Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 600, dlg_height, dlg_name & ""
  	  ButtonGroup ButtonPressed

		'>>> User's favorites
		'>>> This iterates through an array to display the scripts from the favorites text file, in buttons which can be pressed and will run the script.

		'Defining these variables before the start of the loop
		number_of_scripts_in_this_category = 1
		button_placeholder = 100
        SIR_button_placeholder = 200
        update_favorites_button = 400
        hot_topics_button = 500
        update_hotkeys_button = 600

        vert_button_position = 10

        If featured_scripts_array(1).script_name <> "" Then
            GroupBox 5, vert_button_position, 590, 15 + (12 * num_featured_scripts), "These scripts have been recently featured in HOT TOPICS"
            PushButton 205, vert_button_position, 50, 10, "HOT TOPICS", hot_topics_button
            vert_button_position = vert_button_position + 12
            for i = 1 to num_featured_scripts
                script_keys_combine = ""
                If featured_scripts_array(i).dlg_keys(0) <> "" Then script_keys_combine = join(featured_scripts_array(i).dlg_keys, ":")
                PushButton 		10, 					vert_button_position, 	10, 		10, 			"?", 												SIR_button_placeholder
                PushButton 		23,						vert_button_position, 	120, 		10, 			featured_scripts_array(i).script_name, 			button_placeholder
                Text 			152, 				    vert_button_position, 	40, 		10, 			"-- " & script_keys_combine & " --"
                ' PushButton      175,                    vert_button_position,   10,         10,             "+",                                                add_to_favorites_button_placeholder
                Text            190,                    vert_button_position,   450,        10,             "Featured on " & featured_scripts_array(i).hot_topic_date & " --- " & featured_scripts_array(i).description

                featured_scripts_array(i).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
                featured_scripts_array(i).SIR_instructions_button = SIR_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.

                vert_button_position = vert_button_position + 12
                button_placeholder = button_placeholder + 1			'This gets passed to ButtonPressed where it can be refigured as the selected item in the array by subtracting 100
                SIR_button_placeholder = SIR_button_placeholder + 1			'This gets passed to ButtonPressed where it can be refigured as the selected item in the array by subtracting 100

            NEXT
            vert_button_position = vert_button_position + 5
        End If

        Text 5, vert_button_position, 500, 10, "---------------------------------------------------------------------- FAVORITE SCRIPTS ------------------------------------------------------------------------"
        vert_button_position = vert_button_position + 12
        FOR i = 0 TO (ubound(favorite_scripts_array, 1))
            If favorite_scripts_array(i).script_name <> "" Then
                script_keys_combine = ""
                If favorite_scripts_array(i).dlg_keys(0) <> "" Then script_keys_combine = join(favorite_scripts_array(i).dlg_keys, ":")
                PushButton 		5, 						vert_button_position, 	10, 		10, 			"?", 												SIR_button_placeholder
                PushButton 		18,						vert_button_position, 	120, 		10, 			favorite_scripts_array(i).script_name, 			button_placeholder
                Text 			143, 				    vert_button_position, 	40, 		10, 			"-- " & script_keys_combine & " --"
                ' PushButton      175,                    vert_button_position,   10,         10,             "+",                                                add_to_favorites_button_placeholder
                Text            185,                    vert_button_position,   450,        10,             favorite_scripts_array(i).description

                favorite_scripts_array(i).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
                favorite_scripts_array(i).SIR_instructions_button = SIR_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.

                vert_button_position = vert_button_position + 12
                button_placeholder = button_placeholder + 1			'This gets passed to ButtonPressed where it can be refigured as the selected item in the array by subtracting 100
                SIR_button_placeholder = SIR_button_placeholder + 1			'This gets passed to ButtonPressed where it can be refigured as the selected item in the array by subtracting 100

            End If
		NEXT
        vert_button_position = vert_button_position + 5
		'>>> Placing new scripts on the list! This happens in the right-hand column of the dialog.

		right_col_y_pos = dialog_margin + (groupbox_margin * 2)


		'>>> Now we increment through the new scripts, and create buttons for them
        ' MsgBox new_scripts_array(1).script_name
        If new_scripts_array(1).script_name = "" Then
            Text 185, vert_button_position, 200, 10, "*****         No new scripts added in the past 60 Days         *****"
        Else
            Text 5, vert_button_position, 500, 10, "------------------------------------------------------------------------------ NEW SCRIPTS --------------------------------------------------------------------------------"
            vert_button_position = vert_button_position + 12
    		for i = 1 to num_of_new_scripts
                script_keys_combine = ""
                If new_scripts_array(i).dlg_keys(0) <> "" Then script_keys_combine = join(new_scripts_array(i).dlg_keys, ":")
                PushButton 		5, 						vert_button_position, 	10, 		10, 			"?", 												SIR_button_placeholder
                PushButton 		18,						vert_button_position, 	120, 		10, 			new_scripts_array(i).script_name, 			button_placeholder
                Text 			143, 				    vert_button_position, 	40, 		10, 			"-- " & script_keys_combine & " --"
                ' PushButton      175,                    vert_button_position,   10,         10,             "+",                                                add_to_favorites_button_placeholder
                Text            185,                    vert_button_position,   450,        10,             "NEW " & new_scripts_array(i).release_date & "!!! --- " & new_scripts_array(i).description

                new_scripts_array(i).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
                new_scripts_array(i).SIR_instructions_button = SIR_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.

                vert_button_position = vert_button_position + 12
                button_placeholder = button_placeholder + 1			'This gets passed to ButtonPressed where it can be refigured as the selected item in the array by subtracting 100
                SIR_button_placeholder = SIR_button_placeholder + 1			'This gets passed to ButtonPressed where it can be refigured as the selected item in the array by subtracting 100

    		NEXT
        End If
		' GroupBox 210, dialog_margin, button_width + (dialog_margin * 2), 5 + (button_height * (UBound(new_scripts_array) + 1)), "NEW SCRIPTS (LAST 60 DAYS)"


		' PushButton 210, dlg_height - 25, 60, 15, "Update Hotkeys", update_hotkeys_button						<<<<< SEE ISSUE #765
		PushButton 475, dlg_height - 25, 65, 15, "Update Favorites", update_favorites_button
		CancelButton 545, dlg_height - 25, 50, 15
	EndDialog

    Do
    	'>>> Loading the favorites dialog
    	DIALOG Dialog1

    	'>>> Cancelling the script if ButtonPressed = 0
    	IF ButtonPressed = 0 THEN stopscript

        If ButtonPressed = hot_topics_button Then run "C:\Program Files\Internet Explorer\iexplore.exe https://dept.hennepin.us/hsphd/sa/ews/afepages/Adults%20and%20Families%20Eligibility.aspx#BlueZone%20Update"

    	'>>> Giving user has the option of updating their favorites menu.
    	'>>> We should try to incorporate the chainloading function of the new script_end_procedure to bring the user back to their favorites.
    	IF buttonpressed = update_favorites_button THEN
    		call edit_favorites
    		StopScript
    	ElseIf buttonpressed = update_hotkeys_button then
    		call edit_hotkeys
    		StopScript
    	End if

        For i = 1 to ubound(featured_scripts_array)
            If ButtonPressed = featured_scripts_array(i).SIR_instructions_button then call open_URL_in_browser(featured_scripts_array(i).SharePoint_instructions_URL)
            If ButtonPressed = featured_scripts_array(i).button then
                script_to_run = featured_scripts_array(i).script_URL
                Exit For
            End If
        Next
        For i = 0 to ubound(favorite_scripts_array)
            If ButtonPressed = favorite_scripts_array(i).SIR_instructions_button then call open_URL_in_browser(favorite_scripts_array(i).SharePoint_instructions_URL)
            If ButtonPressed = favorite_scripts_array(i).button then
                script_to_run = favorite_scripts_array(i).script_URL
                Exit For
            End If
        Next
        For i = 1 to ubound(new_scripts_array)
            If ButtonPressed = new_scripts_array(i).SIR_instructions_button then call open_URL_in_browser(new_scripts_array(i).SharePoint_instructions_URL)
            If ButtonPressed = new_scripts_array(i).button then
                script_to_run = new_scripts_array(i).script_URL
                Exit For
            End If
        Next

    Loop until script_to_run <> ""
    ' MsgBox script_URL
	' MsgBox "3 - " & dialog1
	' dialog1 = "1 - " & dialog1

END FUNCTION

function open_favorites_file()
    'Needs to determine MyDocs directory before proceeding.
    Set wshshell = CreateObject("WScript.Shell")
    user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"
    favorites_text_file_location = user_myDocs_folder & "\scripts-favorites.txt"

    Dim oTxtFile
    With (CreateObject("Scripting.FileSystemObject"))
    	'>>> If the file exists, we will grab the list of the user's favorite scripts and run the favorites menu.
    	If .FileExists(favorites_text_file_location) Then
            If .GetFile(favorites_text_file_location).size <> 0 Then
        		Set fav_scripts = CreateObject("Scripting.FileSystemObject")
        		Set fav_scripts_command = fav_scripts.OpenTextFile(favorites_text_file_location)
        		fav_scripts_array = fav_scripts_command.ReadAll
        		IF fav_scripts_array <> "" THEN favorites_text_file_array = fav_scripts_array
        		fav_scripts_command.Close
            End If
    	ELSE
    		'>>> ...otherwise, if the file does not exist, the script will require the user to select their favorite scripts.
    		call edit_favorites
    	END IF
    END WITH

    favorites_text_file_string = trim(favorites_text_file_array)
    favorites_text_file_array = split(favorites_text_file_string, vbNewLine)

    num_of_user_scripts = ubound(favorites_text_file_array)

    review_string = "~"
    favorites_text_file_string = replace(favorites_text_file_string, vbNewLine, "~")
    favorites_text_file_string = "~" & favorites_text_file_string & "~"
    count_of_favorites = 0
    For each fav_file in favorites_text_file_array
        fav_file = trim(fav_file)
        If fav_file <> "" Then
            If InStr(review_string, "~" & fav_file & "~") = 0 Then
                review_string = review_string & fav_file & "~"
                count_of_favorites = count_of_favorites + 1
            End If
        End If
    Next
    If count_of_favorites > 25 Then
        MsgBox "Your favorites menu is too large. Please limit the number of favorites to no greater than 25." & vbNewLine & "The script will now run the favorites selection functionality. please choose no more than 25."
        call edit_favorites
    End If
    ' MsgBox "Review String:" & vbNewLine & review_string & vbNewLine & vbNewLine & "Favorites String:" & vbNewLine & favorites_text_file_string
    If review_string <> favorites_text_file_string Then
        review_string = right(review_string, len(review_string) - 1)
        review_string = left(review_string, len(review_string) - 1)

        review_string = replace(review_string, "~", vbNewLine)

        SET updated_fav_scripts_fso = CreateObject("Scripting.FileSystemObject")
        SET updated_fav_scripts_command = updated_fav_scripts_fso.CreateTextFile(favorites_text_file_location, 2)
        updated_fav_scripts_command.Write(review_string)
        updated_fav_scripts_command.Close

        favorites_text_file_string = trim(review_string)
        favorites_text_file_array = split(favorites_text_file_string, vbNewLine)

        num_of_user_scripts = ubound(favorites_text_file_array)

    Else
        review_string = right(review_string, len(review_string) - 1)
        review_string = left(review_string, len(review_string) - 1)

        review_string = replace(review_string, "~", vbNewLine)
        favorites_text_file_string = trim(review_string)
        favorites_text_file_array = split(favorites_text_file_string, vbNewLine)

        num_of_user_scripts = ubound(favorites_text_file_array)
    End If
end function


'DEFINING SOME VARIABLES ===================================================================================================
button_height = 12
button_width = 145
dialog_margin = 5
groupbox_margin = 5
Dim favMenu
'END VARIABLES =============================================================================================================

const fav_script_name   = 0
const add_checkbox      = 1
const cat_as_direct     = 2
const proper_name       = 3
const script_hotkey     = 4
'====================================================================================
'====================================================================================
'This VERY VERY long function contains all of the logic behind editing the favorites.
'====================================================================================
'====================================================================================
' MsgBox "1 - " & dialog1
' dialog1 = "1 - " & dialog1

'======================================

'The script starts HERE!!!-------------------------------------------------------------------------------------------------------------------------------------

Set objNet = CreateObject("WScript.NetWork")
windows_user_ID = objNet.UserName
user_ID_for_validation = ucase(windows_user_ID)

tester_found = FALSE
qi_staff = FALSE
bz_staff = FALSE
For each tester in tester_array
    If user_ID_for_validation = tester.tester_id_number Then
        tester_found = TRUE
        worker_full_name            = tester.tester_full_name
        worker_first_name           = tester.tester_first_name
        worker_last_name            = tester.tester_last_name
        worker_email                = tester.tester_email
        worker_id_number            = tester.tester_id_number
        worker_x_number             = tester.tester_x_number
        worker_supervisor           = tester.tester_supervisor_name
        worker_supervisor_email     = tester.tester_supervisor_email
        worker_population           = tester.tester_population
        worker_region               = tester.tester_region
        worker_test_groups          = tester.tester_groups
        worker_test_scripts         = tester.tester_scripts
        For each group in worker_test_groups
            If group = "QI" Then
                qi_staff = TRUE
            End If
            If group = "BZ" Then
                qi_staff = TRUE
                bz_staff = TRUE
            End If
        Next
    End If
Next

If tester_found = FALSE Then
    tags_coming_soon = MsgBox("***            Coming soon!            ***" & vbNewLine & vbNewLine & "We are updating how we engage with the script tools. One of these ways is with a new system of tagging." & vbNewLine & "This button will have functionality to preview the new menu using these tags. It is not available just yet as we develop and test the functionality." & vbNewLine & vbNewLine & "Come back later to try this new functionality.", vbOk, "New Tags Menu Coming Soon.")
    stopscript
End If

'Needs to determine MyDocs directory before proceeding.
Set wshshell = CreateObject("WScript.Shell")
user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

favorites_text_file_location = user_myDocs_folder & "\scripts-favorites.txt"
hotkeys_text_file_location = user_myDocs_folder & "\scripts-hotkeys.txt"

' ' script_list_URL = script_repository & "COMPLET%20LIST%20OF%20SCRIPTS.vbs"
' script_list_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/COMPLETE%20LIST%20OF%20SCRIPTS.vbs"
' Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' Set fso_command = run_another_script_fso.OpenTextFile(script_list_URL)
' text_from_the_other_script = fso_command.ReadAll
' fso_command.Close
' Execute text_from_the_other_script

'>>> favorited_scripts_array will be built from the contents of the user's text file
favorited_scripts_array = ""

'>>> Building the array of new scripts
num_of_new_scripts = 0
num_featured_scripts = 0
Dim featured_scripts_array()
Dim new_scripts_array()
ReDim featured_scripts_array(1)
ReDim new_scripts_array(1)
'>>> Looking through the scripts array and determining all of the new ones.
FOR i = 0 TO Ubound(script_array)
    IF DateDiff("D", script_array(i).release_date, date) < 60 THEN
		show_this_one = TRUE
		If script_array(i).category = "ADMIN" Then
			show_this_one = FALSE
			For each review_group in script_array(i).tags
				If bz_staff = TRUE AND review_group = "BZ" Then show_this_one = TRUE
				If qi_staff = TRUE AND review_group = "QI" Then show_this_one = TRUE
			Next
		End If
		If show_this_one = TRUE Then
	        num_of_new_scripts = num_of_new_scripts + 1
	        ReDim Preserve new_scripts_array(num_of_new_scripts)
	        SET new_scripts_array(num_of_new_scripts) = NEW script_bowie
	        new_scripts_array(num_of_new_scripts).script_name		= script_array(i).script_name
	        new_scripts_array(num_of_new_scripts).category			= script_array(i).category
	        new_scripts_array(num_of_new_scripts).description		= script_array(i).description
	        new_scripts_array(num_of_new_scripts).release_date		= script_array(i).release_date
	        new_scripts_array(num_of_new_scripts).tags		        = script_array(i).tags
	        new_scripts_array(num_of_new_scripts).dlg_keys		    = script_array(i).dlg_keys
	        new_scripts_array(num_of_new_scripts).keywords		    = script_array(i).keywords
		End If
    end if

    If IsDate(script_array(i).hot_topic_date) = TRUE Then
        IF DateDiff("d", script_array(i).hot_topic_date, date) < 7 Then
            num_featured_scripts = num_featured_scripts + 1
            ReDim Preserve featured_scripts_array(num_featured_scripts)
            SET featured_scripts_array(num_featured_scripts) = NEW script_bowie
            featured_scripts_array(num_featured_scripts).script_name		= script_array(i).script_name
            featured_scripts_array(num_featured_scripts).category			= script_array(i).category
            featured_scripts_array(num_featured_scripts).description		= script_array(i).description
            featured_scripts_array(num_featured_scripts).release_date		= script_array(i).release_date
            featured_scripts_array(num_featured_scripts).tags		        = script_array(i).tags
            featured_scripts_array(num_featured_scripts).dlg_keys		    = script_array(i).dlg_keys
            featured_scripts_array(num_featured_scripts).keywords		    = script_array(i).keywords
            featured_scripts_array(num_featured_scripts).hot_topic_date		= script_array(i).hot_topic_date
        End If
    End If
    If num_featured_scripts = 0 Then SET featured_scripts_array(1) = NEW script_bowie
    If num_of_new_scripts = 0 Then SET new_scripts_array(1) = NEW script_bowie
NEXT

' '>>> This handles what happens if there are no new scripts (it'll happen)
' if num_of_new_scripts = 0 then
'     num_of_new_scripts = 1
'     ReDim Preserve new_scripts_array(num_of_new_scripts)
'     SET new_scripts_array(num_of_new_scripts) = NEW script_bowie
'     new_scripts_array(num_of_new_scripts).script_name		= "no new scripts found."
'     new_scripts_array(num_of_new_scripts).category			= "none"
'     new_scripts_array(num_of_new_scripts).description		= ""
'     new_scripts_array(num_of_new_scripts).release_date		= "none"
'     new_scripts_array(num_of_new_scripts).tags		        = "none"
'     new_scripts_array(num_of_new_scripts).dlg_keys		    = ""
'     new_scripts_array(num_of_new_scripts).keywords		    = "none"
' end if



Dim favorites_text_file_array, num_of_user_scripts, favorites_text_file_string

Call open_favorites_file

'>>> Calling the function that builds the favorites menu.
CALL favorite_menu(favorites_text_file_string, script_to_run)
' MsgBox "4 - " & dialog1
' dialog1 =  "1 - " & dialog1
' dialog1 = ""
Dialog1 = ""
' ReDim scripts_edit_favs_array(0)
' ReDim script_array(0)
' ReDim fav_scripts_array(0)
' ReDIm favorite_scripts_array(0)
' ReDim featured_scripts_array(0)
' ReDim new_scripts_array(0)

If IsArray(scripts_edit_favs_array) = TRUE Then Erase scripts_edit_favs_array
If IsArray(script_array) = TRUE Then Erase script_array
If IsArray(fav_scripts_array) = TRUE Then Erase fav_scripts_array
If IsArray(favorite_scripts_array) = TRUE Then Erase favorite_scripts_array
If IsArray(featured_scripts_array) = TRUE Then Erase featured_scripts_array
If IsArray(new_scripts_array) = TRUE Then Erase new_scripts_array 

'>>> Running the script
MsgBox script_to_run

' list_of_things_to_remove = array("OPTION EXPLICIT", _
' 								"option explicit", _
' 								"Option Explicit", _
' 								"dim case_number", _
' 								"DIM case_number", _
' 								"Dim case_number")
' If run_locally = "" or run_locally = False then					'Runs the script from GitHub if we're not set up to run locally.
' 	Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
' 	req.open "GET", script_to_run, False									'Attempts to open the URL
' 	req.send													'Sends request
' 	If req.Status = 200 Then									'200 means great success
' 		Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
' 		script_contents = req.responseText						'Empties the response into a variable called script_contents
' 		'Uses a for/next to remove the list_of_things_to_remove
' 		FOR EACH phrase IN list_of_things_to_remove
' 			script_contents = replace(script_contents, phrase, "")
' 		NEXT
' 		ExecuteGlobal script_contents							'Executes the remaining script code
' 	ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
' 		critical_error_msgbox = MsgBox ("Something has gone wrong. The code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
' 										"The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
' 										vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
' 		script_end_procedure("Script ended due to error connecting to GitHub.")
' 	END IF
' ELSE
' 	Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' 	Set fso_command = run_another_script_fso.OpenTextFile(script_to_run)
' 	text_from_the_other_script = fso_command.ReadAll
' 	fso_command.Close
' 	ExecuteGlobal text_from_the_other_script
'
' 	call run_another_script(url)
' END IF

CALL run_from_GitHub(script_to_run)
