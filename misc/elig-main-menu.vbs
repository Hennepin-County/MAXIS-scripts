'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ELIG - MAIN MENU.vbs"
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("06/04/2021", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

class subcat
	public subcat_name
	public subcat_button
End class

Function declare_elig_main_menu_dialog()

	'Runs through each script in the array and generates a list of subcategories based on the category located in the function. Also modifies the script description if it's from the last two months, to include a "NEW!!!" notification.
	For current_script = 0 to ubound(script_array)
		'Subcategory handling (creating a second list as a string which gets converted later to an array)
		script_array(current_script).show_script = FALSE
		If script_array(current_script).used_for_elig = True then script_array(current_script).show_script = TRUE

		If IsDate(script_array(current_script).retirement_date) = TRUE Then
			If DateDiff("d", date, script_array(current_script).retirement_date) =< 0 Then script_array(current_script).show_script = FALSE
		End If

		Call script_array(current_script).show_button(see_the_button)
		If see_the_button = FALSE Then script_array(current_script).show_script = FALSE
	Next

    dlg_len = 60
    For current_script = 0 to ubound(script_array)
		If script_array(current_script).show_script = TRUE Then dlg_len = dlg_len + 15
    next

    dialog1 = ""
	BeginDialog dialog1, 0, 0, 600, dlg_len, "ELIG scripts main menu dialog"
	 	Text 5, 5, 435, 10, "ELIG scripts main menu: select the script to run from the choices below."
		EditBox 700, 700, 50, 15, holderbox				'This sits here as the first control element so the fisrt button listed doesn't have the blue box around it.
	  	ButtonGroup ButtonPressed

		Text 5, 25, 435, 10, "-------------------------------------------- NOTES --------------------------------------------"

		'SCRIPT LIST HANDLING--------------------------------------------

		'' 	PushButton 445, 10, 65, 10, "SIR instructions", 	SIR_instructions_button
		'This starts here, but it shouldn't end here :)
		vert_button_position = 35
		For current_script = 0 to ubound(script_array)


            If script_array(current_script).show_script = TRUE AND InStr(script_array(current_script).script_name,"FIAT") = 0 Then

				SIR_button_placeholder = button_placeholder + 1	'We always want this to be one more than the button_placeholder

				'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
				'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
				PushButton 		5, 						vert_button_position, 	10, 		12, 			"?", 												SIR_button_placeholder
				PushButton 		18,						vert_button_position, 	120, 		12, 			script_array(current_script).script_name, 			button_placeholder
				Text 			120 + 23, 				vert_button_position+1, 500, 		14, 			"--- " & script_array(current_script).description
				'----------
				vert_button_position = vert_button_position + 15	'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
				'----------
				script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
				script_array(current_script).SIR_instructions_button = SIR_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
				button_placeholder = button_placeholder + 2
			End if

		next

		Text 5, vert_button_position, 435, 10, "-------------------------------------------- FIATers --------------------------------------------"
		vert_button_position = vert_button_position + 15
		For current_script = 0 to ubound(script_array)


            If script_array(current_script).show_script = TRUE AND InStr(script_array(current_script).script_name,"FIAT") <> 0 Then

				SIR_button_placeholder = button_placeholder + 1	'We always want this to be one more than the button_placeholder

				'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
				'FUNCTION		HORIZ. ITEM POSITION	VERT. ITEM POSITION		ITEM WIDTH	ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
				PushButton 		5, 						vert_button_position, 	10, 		12, 			"?", 												SIR_button_placeholder
				PushButton 		18,						vert_button_position, 	120, 		12, 			script_array(current_script).script_name, 			button_placeholder
				Text 			120 + 23, 				vert_button_position+1, 500, 		14, 			"--- " & script_array(current_script).description
				'----------
				vert_button_position = vert_button_position + 15	'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
				'----------
				script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
				script_array(current_script).SIR_instructions_button = SIR_button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
				button_placeholder = button_placeholder + 2
			End if

		next

		CancelButton 540, dlg_len - 20, 50, 15
	EndDialog
End function

'Starting these with a very high number, higher than the normal possible amount of buttons.
'	We're doing this because we want to assign a value to each button pressed, and we want
'	that value to change with each button. The button_placeholder will be placed in the .button
'	property for each script item. This allows it to both escape the Function and resize
'	near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'	to force the value of ButtonPressed to hold in near infinite iterations.
button_placeholder 			= 24601
subcat_button_placeholder 	= 1701

'Other pre-loop and pre-function declarations
subcategory_array = array()
subcategory_string = ""
subcategory_selected = "MAIN"

'Displays the dialog
Do

	'Creates the dialog
	call declare_elig_main_menu_dialog()

	'At the beginning of the loop, we are not ready to exit it. Conditions later on will impact this.
	ready_to_exit_loop = false

	'Displays dialog, if cancel is pressed then stopscript
	dialog
	If ButtonPressed = 0 then stopscript

	'Runs through each script in the array... if the user selected script instructions (via ButtonPressed) it'll open_URL_in_browser to those instructions
	For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).SIR_instructions_button then call open_URL_in_browser(script_array(i).SharePoint_instructions_URL)
	Next

	'Runs through each script in the array... if the user selected the actual script (via ButtonPressed), it'll run_from_GitHub
	For i = 0 to ubound(script_array)
		If ButtonPressed = script_array(i).button then
			ready_to_exit_loop = true		'Doing this just in case a stopscript or script_end_procedure is missing from the script in question
			script_to_run = script_array(i).script_URL
			Exit for
		End if
	Next


Loop until ready_to_exit_loop = true

call run_from_GitHub(script_to_run)

stopscript
