'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "SELECT FAVORITES.vbs"
start_time = timer

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

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'>>>>>>>>>>>>>>>>>>>>>>>>> SECTION 1 <<<<<<<<<<<<<<<<<<<<<<<<<<
'>>> The gobbins that happen before the user sees anything. <<<
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<


'>>> A class for each script item
'>>> This is needed because it will enable the script to build the array of scripts.
'>>> This needs to be removed when the class is added to FuncLib
class script
	public script_name
	public file_name
	public description
	public button

	public property get button_size	'This part determines the size of the button dynamically by determining the length of the script name, multiplying that by 3.5, rounding the decimal off, and adding 10 px
		button_size = round ( len( script_name ) * 3.5 ) + 10
	end property
end class

'>>> Determining the location of the user's favorites list.
'>>> This value should be stored in Global Variables for state-wide deployment.
network_location_of_favorites_text_file = "H:\my favorite scripts.txt"

'>>> Telling the script main menus that we do not need to call FuncLib or build the dialogs
run_from_favorites = TRUE

'Creating the object to the URL a la text file
SET get_all_scripts = CreateObject("Msxml2.XMLHttp.6.0")

'Grabbing all the actions scripts
actions_url = "https://raw.githubusercontent.com/RobertFewins-Kalb/Anoka-Specific-Scripts/master/GLOBAL-FAVORITES/ACTIONS-MAIN%20MENU.vbs"
'Grabbing all the bulk scripts
bulk_url = "https://raw.githubusercontent.com/RobertFewins-Kalb/Anoka-Specific-Scripts/master/GLOBAL-FAVORITES/BULK-MAIN%20MENU.vbs"
'grabbing all the Notes scripts
notes_url = "https://raw.githubusercontent.com/RobertFewins-Kalb/Anoka-Specific-Scripts/master/GLOBAL-FAVORITES/NOTES-MAIN%20MENU.vbs"
'grabbing all the notices scripts
notices_url = "https://raw.githubusercontent.com/RobertFewins-Kalb/Anoka-Specific-Scripts/master/GLOBAL-FAVORITES/NOTICES-MAIN%20MENU.vbs"
'Creating an array of URLs
all_url_array = actions_url & "UUDDLRLRBA" & bulk_url & "UUDDLRLRBA" & notes_url & "UUDDLRLRBA" & notices_url
all_url_array = split(all_url_array, "UUDDLRLRBA")

'Building an array of all scripts
FOR EACH menu_url IN all_url_array
	'Opening the URL for the given main menu
	get_all_scripts.open "GET", menu_url, FALSE
	get_all_scripts.send
	IF get_all_scripts.Status = 200 THEN
		Set filescriptobject = CreateObject("Scripting.FileSystemObject")
		Execute get_all_scripts.responseText
	ELSE
		'If the script cannot open the URL provided...
		MsgBox 	"Something went wrong with the URL: " & menu_url
		stopscript
	END IF
NEXT

'>>> Building the array of all scripts from the arrays within the main menus. These arrays should
'>>> probably be cleaned up for pushing this script outside Anoka.
'-----------------------
'>>> I would like to see something like all_bulk_scripts, all_notes_scripts, all_notices_scripts
all_scripts_array = ""
FOR EACH specific_script IN script_array_ACTIONS_main
	all_scripts_array = all_scripts_array & specific_script.file_name & "~~~"
NEXT
FOR EACH specific_script IN script_array_BULK_main
	all_scripts_array = all_scripts_array & specific_script.file_name & "~~~"
NEXT
FOR EACH specific_script IN script_array_BULK_list
	all_scripts_array = all_scripts_array & specific_script.file_name & "~~~"
NEXT
FOR EACH specific_script IN script_array_0_to_C
	all_scripts_array = all_scripts_array & specific_script.file_name & "~~~"
NEXT
FOR EACH specific_script IN script_array_D_to_F
	all_scripts_array = all_scripts_array & specific_script.file_name & "~~~"
NEXT
FOR EACH specific_script IN script_array_G_to_L
	all_scripts_array = all_scripts_array & specific_script.file_name & "~~~"
NEXT
FOR EACH specific_script IN script_array_M_to_Q
	all_scripts_array = all_scripts_array & specific_script.file_name & "~~~"
NEXT
FOR EACH specific_script IN script_array_R_to_Z
	all_scripts_array = all_scripts_array & specific_script.file_name & "~~~"
NEXT
FOR EACH specific_script IN script_array_LTC
	all_scripts_array = all_scripts_array & specific_script.file_name & "~~~"
NEXT
FOR EACH specific_script IN script_array_NOTICES_main
	all_scripts_array = all_scripts_array & specific_script.file_name & "~~~"
NEXT
FOR EACH specific_script IN script_array_NOTICES_list
	all_scripts_array = all_scripts_array & specific_script.file_name & "~~~"
NEXT

'Warning/instruction box
MsgBox "This script will display a dialog with various scripts on it."  & vbNewLine &_
		"Any script you check will be added to your favorites menu.  " & vbNewline &_
		"Scripts you un-check will be removed. Once you are done " & vbNewLine &_
		"making your selection hit OK and your menu will be updated. " & vbNewLine & vbNewLine &_
		"- You will be unable to edit NEW Scripts and Recommended Scripts."

'Removing .vbs for readability for the end user...they don't need to see that
all_scripts_array = all_scripts_array & "~THE~END~"
all_scripts_array = replace(all_scripts_array, "~THE~END~", "")
all_scripts_array = replace(all_scripts_array, ".vbs", "")

all_scripts_array = trim(all_scripts_array)
all_scripts_array = split(all_scripts_array, "~~~")

number_of_scripts = UBound(all_scripts_array)

'Creating a new multi-dimensional array for this script. We need to hang on to the checked/unchecked value.
ReDim scripts_multidimensional_array(number_of_scripts, 1)

'Converting the single-dimensional array into a multi-dimensional array.
scr_pos = 0
FOR EACH scriptName IN all_scripts_array
	scripts_multidimensional_array(scr_pos, 0) = scriptName
	scripts_multidimensional_array(scr_pos, 1) = 0
	scr_pos = scr_pos + 1
NEXT

'>>> If the user has already selected their favorites, the script will open that file and
'>>> and read it, storing the contents in the variable name ''user_scripts_array''
SET oTxtFile = (CreateObject("Scripting.FileSystemObject"))
With oTxtFile
	If .FileExists(network_location_of_favorites_text_file) Then
		Set fav_scripts = CreateObject("Scripting.FileSystemObject")
		Set fav_scripts_command = fav_scripts.OpenTextFile(network_location_of_favorites_text_file)
		fav_scripts_array = fav_scripts_command.ReadAll
		IF fav_scripts_array <> "" THEN user_scripts_array = fav_scripts_array
		fav_scripts_command.Close
	END IF
END WITH

'>>> Determining the width of the dialog from the number of scripts that are available...
IF number_of_scripts <= 39 THEN
	dia_width = 205
ELSEIF number_of_scripts >= 40 AND number_of_scripts <=79 THEN
	dia_width = 400
ELSEIF number_of_scripts >= 80 AND number_of_scripts <= 119 THEN
	dia_width = 605
ELSEIF number_of_scripts >= 120 AND number_of_scripts <= 159 THEN
	dia_width = 800
END IF

'>>> Building the dialog
BeginDialog fav_dlg, 0, 0, dia_width, 440, "Select your favorites"
	ButtonGroup ButtonPressed
		OkButton 5, 5, 50, 15
		CancelButton 55, 5, 50, 15
		PushButton 165, 5, 70, 15, "Reset Favorites", reset_favorites_button
	row = 30
	'>>> Creating the display of all scripts for selection (in checkbox form)
	FOR i = 0 to number_of_scripts
		IF scripts_multidimensional_array(i, 0) <> "" THEN
			'>>> Determining the positioning of the checkboxes.
			'>>> For some reason, even though we exceed 65 objects, we do not hit any issues with missing scripts. Oh well.
			IF i <= 39 THEN
				col = 10
			ELSEIF i >= 40 AND i <= 79 THEN
				col = 205
			ELSEIF i >= 80 AND i <= 119 THEN
				col = 400
			ELSEIF i > 119 THEN
				col = 605
			END IF
			IF row = 430 THEN row = 30
			'>>> If the script in question is already known to the list of scripts already picked by the user, the check box is defaulted to checked.
			IF InStr(user_scripts_array, scripts_multidimensional_array(i, 0)) <> 0 THEN
				scripts_multidimensional_array(i, 1) = 1
			ELSE
				scripts_multidimensional_array(i, 1) = 0
			END IF
			CheckBox col, row, 185, 10, scripts_multidimensional_array(i, 0), scripts_multidimensional_array(i, 1)
			row = row + 10
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
		Dialog fav_dlg
			'>>> Cancel confirmation
			IF ButtonPressed = 0 THEN
				confirm_cancel = MsgBox("Are you sure you want to cancel? Press YES to cancel the script. Press NO to return to the script.", vbYesNo)
				IF confirm_cancel = vbYes THEN script_end_procedure("Script cancelled.")
			END IF
			'>>> If the user selects to reset their favorites selections, the script
			'>>> will go through the multi-dimensional array and reset all the values
			'>>> for position 1, thereby clearing the favorites from the display.
			IF ButtonPressed = reset_favorites_button THEN
				FOR i = 0 to number_of_scripts
					scripts_multidimensional_array(i, 1) = 0
				NEXT
			END IF
	'>>> The exit condition for the first do/loop is the user pressing 'OK'
	LOOP UNTIL ButtonPressed <> 0 AND ButtonPressed <> reset_favorites_button
	'>>> Validating that the user does not select more than a prescribed number of scripts.
	'>>> Exceeding the limit will cause an exception access violation for the Favorites script when it runs.
	'>>> Currently, that value is 38. That is lower than previous because of the larger number of new scripts. (-Robert, 04/20/2016)
	double_check_array = ""
	FOR i = 0 to number_of_scripts
		IF scripts_multidimensional_array(i, 1) = 1 THEN double_check_array = double_check_array & scripts_multidimensional_array(i, 0) & "~"
	NEXT
	double_check_array = split(double_check_array, "~")
	IF ubound(double_check_array) > 37 THEN MsgBox "Your favorites menu is too large. Please limit the number of favorites to no greater than 37."
	'>>> Exit condition is the user having fewer than 38 scripts in their favorites menu.
LOOP UNTIL ubound(double_check_array) <= 37

'>>> Getting ready to write the user's selection to a text file and save it on a prescribed location on the network.
'>>> Building the content of the text file.
favorite_scripts = ""
FOR i = 0 to number_of_scripts - 1
	IF scripts_multidimensional_array(i, 1) = 1 THEN favorite_scripts = favorite_scripts & scripts_multidimensional_array(i, 0) & "~~~"
NEXT

'>>> After the user selects their favorite scripts, we are going to write (or overwrite) the list of scripts
'>>> stored at H:\my favorite scripts.txt.
IF favorite_scripts <> "" THEN
	SET updated_fav_scripts_fso = CreateObject("Scripting.FileSystemObject")
	SET updated_fav_scripts_command = updated_fav_scripts_fso.CreateTextFile(network_location_of_favorites_text_file, 2)
	updated_fav_scripts_command.Write(favorite_scripts)
	updated_fav_scripts_command.Close
	script_end_procedure("Success!! Your Favorites Menu has been updated.")
ELSE
	'>>> OR...if the user has selected no scripts for their favorite, the file will be deleted to
	'>>> prevent the Favorites Menu from erroring out.
	'>>> Experience with worker_signature automation tells us that if the text file is blank, the favorites menu doth not work.
	oTxtFile.DeleteFile(network_location_of_favorites_text_file)
	script_end_procedure("You have updated your Favorites Menu, but you haven't selected any scripts. The next time you use the Favorites scripts, you will need to select your favorites.")
END IF
