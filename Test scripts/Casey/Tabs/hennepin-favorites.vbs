'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "Hennepin Favorites.vbs"
start_time = timer
run_locally = true
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


'DEFINING SOME VARIABLES ===================================================================================================
button_height = 12
button_width = 145
dialog_margin = 5
groupbox_margin = 5
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
function edit_favorites

	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>> SECTION 1 <<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>> The gobbins that happen before the user sees anything. <<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
	'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
    '
	' If list_of_scripts_ran <> true then
	' 	' Looks up the script details online (or locally if you're a scriptwriter)
	' 	If run_locally <> true then
	' 		SET get_all_scripts = CreateObject("Msxml2.XMLHttp.6.0")				' Creating the object to the URL a la text file
	' 		get_all_scripts.open "GET", all_scripts_repo, FALSE						' Creating an AJAX object
	' 		get_all_scripts.send													' Opening the URL for the given main menu
	' 		IF get_all_scripts.Status = 200 THEN									' 200 means great success
	' 			Set filescriptobject = CreateObject("Scripting.FileSystemObject")	' Create an FSO for the script object
	' 			Execute get_all_scripts.responseText								' Execute the script (building an array of all scripts)
	' 		ELSE																	' If the script cannot open the URL provided...
	' 			MsgBox 	"Something went wrong with the URL: " & all_scripts_repo	' Tell the worker
	' 			stopscript															' Stop the script
	' 		END IF
	' 	ELSE																		' If it's set as run_locally...
	' 		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")	' Create an FSO to read the script list file
	' 		Set fso_command = run_another_script_fso.OpenTextFile(all_scripts_repo)	' Create a command object to run the script list object
	' 		text_from_the_other_script = fso_command.ReadAll						' Read the file
	' 		fso_command.Close														' Close the file
	' 		Execute text_from_the_other_script										' Execute the script (building an array of all scripts)
	' 	END If
	' end if

	'Warning/instruction box
	MsgBox  "This section will display a dialog with various scripts on it. Any script you check will be added to your favorites menu. Scripts you un-check will be removed. Once you are done making your selection hit ""OK"" and your menu will be updated. " & vbNewLine & vbNewLine &_
			"Note: you will be unable to edit the list of new scripts."

	'An array containing details about the list of scripts, including how they are displayed and stored in the favorites tag
	'0 => The script name
	'1 => The checked/unchecked status (based on the dialog list)
	'2 => The script category, and a "/" so that it's presented in a URL
	'3 => The proper script file name
	'4 => The hotkey the user has associated with the script


	REDIM scripts_edit_favs_array(ubound(script_array), script_hotkey)

	'determining the number of each kind of script...by category
	number_of_scripts = 0
	actions_scripts = 0
	bulk_scripts = 0
	notc_scripts = 0
	notes_scripts = 0
	utilities_scripts = 0
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
	    End if
	NEXT


	'>>> If the user has already selected their favorites, the script will open that file and
	'>>> and read it, storing the contents in the variable name ''favorites_text_file_array''
	SET oTxtFile = (CreateObject("Scripting.FileSystemObject"))
	With oTxtFile
		If .FileExists(favorites_text_file_location) Then
			Set fav_scripts = CreateObject("Scripting.FileSystemObject")
			Set fav_scripts_command = fav_scripts.OpenTextFile(favorites_text_file_location)
			fav_scripts_array = fav_scripts_command.ReadAll
			IF fav_scripts_array <> "" THEN favorites_text_file_array = fav_scripts_array
			fav_scripts_command.Close
		END IF
	END WITH

	'>>> Determining the width of the dialog from the number of scripts that are available...
	'the dialog starts with a width of 400
	dia_width = 750

	'VKC - removed old functionality to determine dynamically the width. This will need to be redetermined based on the number of scripts, but I am holding off on this until I know all of the content I'll jam in here. -11/29/2016

	'>>> Building the dialog
	BeginDialog build_new_favorites_dialog, 0, 0, dia_width, 380, "Select your favorites"
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
					row = 30
					col = col + 150
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
				CheckBox col, row, 140, 10, scripts_edit_favs_array(script_position, fav_script_name), scripts_edit_favs_array(script_position, add_checkbox)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 20	'Padding for the new section
		'Account for overflow
		IF row >= 360 THEN
			row = 30
			col = col + 150
		END IF
		Text col, row, 175, 10, "---------- BULK SCRIPTS ----------"
		row = row + 10

		'BULK script laying out
		FOR i = 0 to ubound(script_array)
			IF script_array(i).category = "BULK" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 360 THEN
					row = 30
					col = col + 150
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
				CheckBox col, row, 140, 10, scripts_edit_favs_array(script_position, fav_script_name), scripts_edit_favs_array(script_position, add_checkbox)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 20	'Padding for the new section
		'Account for overflow
		IF row >= 360 THEN
			row = 30
			col = col + 150
		END IF
		Text col, row, 175, 10, "---------- NOTICES SCRIPTS ----------"
		row = row + 10

		'CALCULATOR script laying out
		FOR i = 0 to ubound(script_array)
			IF script_array(i).category = "NOTICES" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 360 THEN
					row = 30
					col = col + 150
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
				CheckBox col, row, 140, 10, scripts_edit_favs_array(script_position, fav_script_name), scripts_edit_favs_array(script_position, add_checkbox)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 20	'Padding for the new section
		'Account for overflow
		IF row >= 360 THEN
			row = 30
			col = col + 150
		END IF
		Text col, row, 175, 10, "---------- NOTES SCRIPTS ----------"
		row = row + 10

		'NOTES script laying out
		FOR i = 0 to ubound(script_array)
			IF script_array(i).category = "NOTES" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 360 THEN
					row = 30
					col = col + 150
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
				CheckBox col, row, 140, 10, scripts_edit_favs_array(script_position, fav_script_name), scripts_edit_favs_array(script_position, add_checkbox)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT

		'Section header
		row = row + 20	'Padding for the new section
		'Account for overflow
		IF row >= 360 THEN
			row = 30
			col = col + 150
		END IF
		Text col, row, 175, 10, "---------- UTILITIES SCRIPTS ----------"
		row = row + 10

		'UTILITIES script laying out
	    FOR i = 0 to ubound(script_array)
			IF script_array(i).category = "UTILITIES" THEN
				'>>> Determining the positioning of the checkboxes.
				'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
				IF row >= 360 THEN
					row = 30
					col = col + 150
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
				CheckBox col, row, 140, 10, scripts_edit_favs_array(script_position, fav_script_name), scripts_edit_favs_array(script_position, add_checkbox)

				'Increments the row and script_position
				row = row + 10
				script_position = script_position + 1
			END IF
		NEXT
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
			Dialog build_new_favorites_dialog
			'>>> Cancel confirmation
			IF ButtonPressed = 0 THEN
				confirm_cancel = MsgBox("Are you sure you want to cancel? Press YES to cancel the script. Press NO to return to the script.", vbYesNo)
				IF confirm_cancel = vbYes THEN script_end_procedure("~PT: Script cancelled.")
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
		IF ubound(double_check_array) > 29 THEN MsgBox "Your favorites menu is too large. Please limit the number of favorites to no greater than 30."
		'>>> Exit condition is the user having fewer than 30 scripts in their favorites menu.
	LOOP UNTIL ubound(double_check_array) <= 29

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
		script_end_procedure("Success!! Your Favorites Menu has been updated. Please click your favorites list button to re-load them.")
	ELSE
		'>>> OR...if the user has selected no scripts for their favorite, the file will be deleted to
		'>>> prevent the Favorites Menu from erroring out.
		'>>> Experience with worker_signature automation tells us that if the text file is blank, the favorites menu doth not work.
		oTxtFile.DeleteFile(favorites_text_file_location)
		script_end_procedure("You have updated your Favorites Menu, but you haven't selected any scripts. The next time you use the Favorites scripts, you will need to select your favorites.")
	END IF

end function


'Needs to determine MyDocs directory before proceeding.
Set wshshell = CreateObject("WScript.Shell")
user_myDocs_folder = wshShell.SpecialFolders("MyDocuments") & "\"

favorites_text_file_location = user_myDocs_folder & "\scripts-favorites.txt"
hotkeys_text_file_location = user_myDocs_folder & "\scripts-hotkeys.txt"

script_list_URL = "C:\MAXIS-scripts\Test scripts\Casey\Tabs\COMPLETE LIST OF SCRIPTS.vbs"
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile(script_list_URL)
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'>>> favorited_scripts_array will be built from the contents of the user's text file
favorited_scripts_array = ""

'>>> Building the array of new scripts
num_of_new_scripts = 0

'>>> Looking through the scripts array and determining all of the new ones.
FOR i = 0 TO Ubound(script_array)
	IF DateDiff("D", script_array(i).release_date, date) < 60 THEN
		num_of_new_scripts = num_of_new_scripts + 1
		ReDim Preserve new_scripts_array(num_of_new_scripts)
		SET new_scripts_array(num_of_new_scripts) = NEW cs_script
		new_scripts_array(num_of_new_scripts).script_name		= script_array(i).script_name
		new_scripts_array(num_of_new_scripts).category			= script_array(i).category
		new_scripts_array(num_of_new_scripts).description		= script_array(i).description
		new_scripts_array(num_of_new_scripts).release_date		= script_array(i).release_date
        new_scripts_array(num_of_new_scripts).tags		        = script_array(i).tags
        new_scripts_array(num_of_new_scripts).dlg_keys		    = script_array(i).dlg_keys
        new_scripts_array(num_of_new_scripts).keywords		    = script_array(i).keywords
	end if
NEXT

'>>> This handles what happens if there are no new scripts (it'll happen)
if num_of_new_scripts = 0 then
	num_of_new_scripts = 1
	ReDim Preserve new_scripts_array(num_of_new_scripts)
	SET new_scripts_array(num_of_new_scripts) = NEW script_bowie
	new_scripts_array(num_of_new_scripts).script_name		= "no new scripts found."
	new_scripts_array(num_of_new_scripts).category			= "none"
	new_scripts_array(num_of_new_scripts).description		= "none"
	new_scripts_array(num_of_new_scripts).release_date		= "none"
    new_scripts_array(num_of_new_scripts).tags		        = "none"
    new_scripts_array(num_of_new_scripts).dlg_keys		    = "none"
    new_scripts_array(num_of_new_scripts).keywords		    = "none"
end if

'>>> Custom function that builds the Favorites Main Menu dialog.
'>>> the array of the user's scripts

' const fav_script_name   = 0
const script_directory  = 1
const fav_script_btn    = 2
const fav_category      = 3
const fav_name_wo_cat   = 4
const fav_redirect      = 5
const display_name      = 6
FUNCTION favorite_menu(favorites_text_file_array, script_URL)
	'>>> Splitting the array of all scripts.
	favorites_text_file_array = trim(favorites_text_file_array)
	favorites_text_file_array = split(favorites_text_file_array, vbNewLine)

	num_of_user_scripts = ubound(favorites_text_file_array)


	num_of_scripts = num_of_user_scripts + num_of_new_scripts

	ReDim favorited_scripts_array(num_of_scripts, display_name)
	'position 0 = script name
	'position 1 = script directory
	'position 2 = button
	'position 3 = category
	'position 4 = script name without category
	'position 5 = state-supported true/false
	'position 6 = friendly name

	scripts_pos = 0
	FOR EACH script_path IN favorites_text_file_array
        script_path = trim(script_path)
        ' MsgBox script_path
		IF script_path <> "" THEN
			favorited_scripts_array(scripts_pos, fav_script_name) = script_path
			'>>> Creating the correct URL for the github call
			IF left(script_path, 5) = "NOTES" THEN
				favorited_scripts_array(scripts_pos, script_directory)  = script_path
				favorited_scripts_array(scripts_pos, fav_category)      = "NOTES"
				favorited_scripts_array(scripts_pos, fav_name_wo_cat)   = right(script_path, len(script_path) - 6)
			ELSEIF left(script_path, 7) = "ACTIONS" THEN
				favorited_scripts_array(scripts_pos, script_directory)  = script_path
				favorited_scripts_array(scripts_pos, fav_category)      = "ACTIONS"
				favorited_scripts_array(scripts_pos, fav_name_wo_cat)   = right(script_path, len(script_path) - 8)
			ELSEIF left(script_path, 4) = "BULK" THEN
				favorited_scripts_array(scripts_pos, script_directory)  = script_path
				favorited_scripts_array(scripts_pos, fav_category)      = "BULK"
				favorited_scripts_array(scripts_pos, fav_name_wo_cat)   = right(script_path, len(script_path) - 5)
			ELSEIF left(script_path, 7) = "NOTICES" THEN
				favorited_scripts_array(scripts_pos, script_directory)  = script_path
				favorited_scripts_array(scripts_pos, fav_category)      = "NOTICES"
				favorited_scripts_array(scripts_pos, fav_name_wo_cat)   = right(script_path, len(script_path) - 8)
            ELSEIF left(script_path, 9) = "UTILITIES" THEN
    			favorited_scripts_array(scripts_pos, script_directory)   = script_path
    			favorited_scripts_array(scripts_pos, fav_category)       = "UTILITIES"
    			favorited_scripts_array(scripts_pos, fav_name_wo_cat)    = right(script_path, len(script_path) - 10)
			END IF
            ' MsgBox script_path
			' 'This part reads the complete list of scripts, and then stores the "friendly name" as an item in the array (makes the dialog prettier down the road)
			for each script_data_from_complete_list in script_array
                ' favorited_scripts_array(scripts_pos, display_name) = script_data_from_complete_list.category & " - " & script_data_from_complete_list.script_name
                ' MsgBox "fav array - " & favorited_scripts_array(scripts_pos, fav_name_wo_cat) & vbNewLine & "array name - " & script_data_from_complete_list.script_name
				if favorited_scripts_array(scripts_pos, fav_name_wo_cat) = script_data_from_complete_list.script_name then
                    favorited_scripts_array(scripts_pos, display_name) = script_data_from_complete_list.script_name
                    favorited_scripts_array(scripts_pos, fav_redirect) = script_data_from_complete_list.script_URL
                End If
			next

			scripts_pos = scripts_pos + 1
		END IF
	NEXT

	'>>> Determining the height parameters to enable the group boxes.
	actions_count = 0
	bulk_count = 0
	notc_count = 0
	notes_count = 0
	utilities_count = 0
	FOR i = 0 TO (ubound(favorites_text_file_array) - 1)
		IF favorited_scripts_array(i, fav_category) = "ACTIONS" THEN
			actions_count = actions_count + 1
		ELSEIF favorited_scripts_array(i, fav_category) = "BULK" THEN
			bulk_count = bulk_count + 1
		ELSEIF favorited_scripts_array(i, fav_category) = "NOTICES" THEN
			notc_count = notc_count + 1
		ELSEIF favorited_scripts_array(i, fav_category) = "NOTES" THEN
			notes_count = notes_count + 1
        ELSEIF favorited_scripts_array(i, fav_category) = "UTILITIES" THEN
    		utilities_count = utilities_count + 1
		END IF
	NEXT

	'>>> Determining the height of the dialog.
	'>>> Each groupbox will require a minimum of 25 pixels. That is the height of the groupbox with 1 script PushButton
	'>>> The groupboxes need to grow 10 for each script pushbutton, so the dialog also needs to grow 10 for each script push button. However,
	'>>> 	the size of each groupbox will always be 15 plus (10 times the number of that kind of script)...
	dlg_height = 0
	IF actions_count <> 0 THEN dlg_height = 15 + (button_height * actions_count)
	IF bulk_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (button_height * bulk_count))
	IF notc_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (button_height * bulk_count))
	IF notes_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (button_height * notes_count))
    IF utilities_count <> 0 THEN dlg_height = dlg_height + 15 + (15 + (button_height * utilities_count))
	dlg_height = dlg_height + 5
	'>>> The dialog needs to be at least 185 pixels tall. If it is not...because the user has not selected a sufficient number of scripts...then
	'>>> the dialog needs to grow to 185.

	'>>> Adjusting the height if the user has fewer scripts selected (in the left column) than what is "new" (in the right column).
	right_col_dlg_height = 60 + (button_height * (Ubound(new_scripts_array) + 1))
	IF right_col_dlg_height > dlg_height THEN dlg_height = right_col_dlg_height



	'>>> Defining some variables for use in the display
	groupbox_y_pos = dialog_margin
	left_col_y_pos = groupbox_y_pos + (groupbox_margin * 2)
	left_col_x_pos = dialog_margin + groupbox_margin



	'>>> A nice decoration for the user. If they have used Update Worker Signature, then their signature is built into the dialog display.
	IF worker_signature <> "" THEN
		dlg_name = worker_signature & "'s Favorite Scripts"
	ELSE
		dlg_name = "My Favorite Scripts"
	END IF

	'>>> The dialog
	BeginDialog favorites_dialog, 0, 0, 411, dlg_height, dlg_name & " "
  	  ButtonGroup ButtonPressed

		'>>> User's favorites
		'>>> This iterates through an array to display the scripts from the favorites text file, in buttons which can be pressed and will run the script.

		'Defining these variables before the start of the loop
		number_of_scripts_in_this_category = 1
		button_placeholder = 100

		'The actual array (this goes through the text file and creates scripts and buttons)
        ' MsgBox ubound(favorites_text_file_array)
		' FOR i = 0 TO (ubound(favorites_text_file_array) - 1)
        FOR i = 0 TO (ubound(favorites_text_file_array))
			'Defines the current category for comparison purposes, and to write out the labels.
			current_script_category_from_list = favorited_scripts_array(i, fav_category)

			'Determines the next script category, but only if it's not at the end of the array (because then we're out and the ubound would error out)
			if i + 1 < ubound(favorited_scripts_array) - 1 then
				next_script_category_from_list = favorited_scripts_array(i + 1, fav_category)
			end if

			'Adding the button
			PushButton left_col_x_pos + groupbox_margin, left_col_y_pos, button_width, button_height, favorited_scripts_array(i, display_name), button_placeholder
			button_placeholder = button_placeholder + 1			'This gets passed to ButtonPressed where it can be refigured as the selected item in the array by subtracting 100

			'If the current category differs from the next, it's time to make the groupbox
			if current_script_category_from_list <> next_script_category_from_list OR i = (ubound(favorites_text_file_array) - 1) then
				left_col_y_pos = groupbox_y_pos + ((groupbox_margin * 3) + (button_height * number_of_scripts_in_this_category))		'Margin x 3 because we need extra padding on the top, and half the extra on the bottom
				GroupBox left_col_x_pos, groupbox_y_pos, button_width + (groupbox_margin * 2), (groupbox_margin * 3) + (button_height * number_of_scripts_in_this_category), current_script_category_from_list
				number_of_scripts_in_this_category = 1
				groupbox_y_pos = left_col_y_pos
				left_col_y_pos = left_col_y_pos + (groupbox_margin * 2)
			else
				left_col_y_pos = left_col_y_pos + button_height
				number_of_scripts_in_this_category = number_of_scripts_in_this_category + 1
			end if

		NEXT

		'>>> Placing new scripts on the list! This happens in the right-hand column of the dialog.

		right_col_y_pos = dialog_margin + (groupbox_margin * 2)


		'>>> Now we increment through the new scripts, and create buttons for them
		for i = 1 to num_of_new_scripts
			PushButton 215, right_col_y_pos, button_width, button_height, ucase(new_scripts_array(i).category) & " - " & new_scripts_array(i).script_name, button_placeholder
			button_placeholder = button_placeholder + 1			'This gets passed to ButtonPressed where it can be refigured as the selected item in the array by subtracting 100
			right_col_y_pos = right_col_y_pos + button_height
		NEXT

		GroupBox 210, dialog_margin, button_width + (dialog_margin * 2), 5 + (button_height * (UBound(new_scripts_array) + 1)), "NEW SCRIPTS (LAST 60 DAYS)"


		' PushButton 210, dlg_height - 25, 60, 15, "Update Hotkeys", update_hotkeys_button						<<<<< SEE ISSUE #765
		PushButton 285, dlg_height - 25, 65, 15, "Update Favorites", update_favorites_button
		CancelButton 355, dlg_height - 25, 50, 15
	EndDialog

	'>>> Loading the favorites dialog
	DIALOG favorites_dialog
	'>>> Cancelling the script if ButtonPressed = 0
	IF ButtonPressed = 0 THEN stopscript



	'>>> Giving user has the option of updating their favorites menu.
	'>>> We should try to incorporate the chainloading function of the new script_end_procedure to bring the user back to their favorites.
	IF buttonpressed = update_favorites_button THEN
		call edit_favorites
		StopScript
	ElseIf buttonpressed = update_hotkeys_button then
		call edit_hotkeys
		StopScript
	End if

	'Determining the script that was selected, simply by subtracting 100 from the button_placeholder we'd previously defined. This corresponds with the array item selected.
	selected_script = ButtonPressed - 100

	'If it's a new script, it'll be larger than the text file array since it's displayed after, so this will create a variable for that too.
	selected_new_script = selected_script - (ubound(favorites_text_file_array) - 1)

    ' MsgBox "Button Pressed - " & ButtonPressed & vbNewLine & "Selected Script - " & selected_script & vbNewLine & "Name - " & favorited_scripts_array(selected_script, display_name)
	'This part takes the selected script integer and determines the file path for it
	if selected_script < ubound(favorited_scripts_array, 1) then
		script_URL = favorited_scripts_array(selected_script, fav_redirect)  '!!!! This works in conjunction with the button_placeholder that's used and incremented for each button. It won't work otherwise.
	else
		script_URL = new_scripts_array(selected_new_script).category & "/" & new_scripts_array(selected_new_script).file_name
	end if
    ' MsgBox script_URL
END FUNCTION
'======================================

'The script starts HERE!!!-------------------------------------------------------------------------------------------------------------------------------------

'>>> The gobbins of the script that the user sees and makes do.
'>>> Declaring the text file storing the user's favorite scripts list.
Dim oTxtFile
With (CreateObject("Scripting.FileSystemObject"))
	'>>> If the file exists, we will grab the list of the user's favorite scripts and run the favorites menu.
	If .FileExists(favorites_text_file_location) Then
		Set fav_scripts = CreateObject("Scripting.FileSystemObject")
		Set fav_scripts_command = fav_scripts.OpenTextFile(favorites_text_file_location)
		fav_scripts_array = fav_scripts_command.ReadAll
		IF fav_scripts_array <> "" THEN favorites_text_file_array = fav_scripts_array
		fav_scripts_command.Close
	ELSE
		'>>> ...otherwise, if the file does not exist, the script will require the user to select their favorite scripts.
		call edit_favorites
	END IF
END WITH


'>>> Calling the function that builds the favorites menu.
CALL favorite_menu(favorites_text_file_array, script_URL)

''Figuring out where the script goes...
'if selected_script < ubound(favorites_text_file_array) then
'	script_location = favorited_scripts_array(selected_script, 1)
'end if
'


'>>> Running the script
CALL run_from_GitHub(script_URL)
