'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = " IVE - MAIN MENU.vbs"
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
call changelog_update("09/22/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CUSTOM FUNCTIONS===========================================================================================================
Function declare_IVE_menu_dialog(script_array)
	BeginDialog IVE_dialog, 0, 0, 516, 120, "IV-E Scripts"
	 	Text 5, 5, 435, 10, "IV-E scripts main menu: select the script to run from the choices below."
	  	EditBox 700, 700, 50, 15, holderbox				'This sits here as the first control element so the fisrt button listed doesn't have the blue box around it.
		ButtonGroup ButtonPressed
		 	'PushButton 015, 35, 40, 15, "CA", 				IVE_main_button
		 	'PushButton 445, 10, 65, 10, "SIR instructions", 	SIR_instructions_button

		'This starts here, but it shouldn't end here :)
		vert_button_position = 25

		For current_script = 0 to ubound(script_array)
			'Displays the button and text description-----------------------------------------------------------------------------------------------------------------------------
			'FUNCTION		HORIZ. ITEM POSITION								VERT. ITEM POSITION		ITEM WIDTH									ITEM HEIGHT		ITEM TEXT/LABEL										BUTTON VARIABLE
			PushButton 		5, 													vert_button_position, 	script_array(current_script).button_size, 	10, 			script_array(current_script).script_name, 			button_placeholder
			Text 			script_array(current_script).button_size + 10, 		vert_button_position, 	500, 										10, 			"--- " & script_array(current_script).description
			'----------
			vert_button_position = vert_button_position + 15	'Needs to increment the vert_button_position by 15px (used by both the text and buttons)
			'----------
			script_array(current_script).button = button_placeholder	'The .button property won't carry through the function. This allows it to escape the function. Thanks VBScript.
			button_placeholder = button_placeholder + 1
		next

		CancelButton 455, 95, 50, 15
		'GroupBox 5, 20, 205, 35, "IVE Sub-Menus"
	EndDialog
End function
'END CUSTOM FUNCTIONS=======================================================================================================

'VARIABLES TO DECLARE=======================================================================================================

'Declaring the variable names to cut down on the number of arguments that need to be passed through the function.
DIM ButtonPressed
'DIM SIR_instructions_button
dim IVE_dialog

script_array_IVE_main = array()
'script_array_IVE_list = array()


'END VARIABLES TO DECLARE===================================================================================================

'LIST OF SCRIPTS================================================================================================================

'INSTRUCTIONS: simply add your new script below. Scripts are listed in alphabetical order. Copy a block of code from above and paste your script info in. The function does the rest.

'-------------------------------------------------------------------------------------------------------------------------IVE MAIN MENU

'Resetting the variable
script_num = 0
ReDim Preserve script_array_IVE_main(script_num)
Set script_array_IVE_main(script_num) = new script
script_array_IVE_main(script_num).script_name 			= "Adoption Assistance"																'Script name
script_array_IVE_main(script_num).file_name 			= "ive-adoption-assistance.vbs"															'Script URL
script_array_IVE_main(script_num).description 			= "Case note template for Adoption Assistance cases."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_IVE_main(script_num)			'Resets the array to add one more element to it
Set script_array_IVE_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_IVE_main(script_num).script_name 			= "IV-E"																'Script name
script_array_IVE_main(script_num).file_name 			= "ive-ive.vbs"															'Script URL
script_array_IVE_main(script_num).description 			= "Case note template for IV-E."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_IVE_main(script_num)			'Resets the array to add one more element to it
Set script_array_IVE_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_IVE_main(script_num).script_name 			= "Elig Review"																		'Script name
script_array_IVE_main(script_num).file_name 			= "ive-elig-review.vbs"															'Script URL
script_array_IVE_main(script_num).description 			= "Case note template for IV-E eligibilty review cases."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_IVE_main(script_num)			'Resets the array to add one more element to it
Set script_array_IVE_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_IVE_main(script_num).script_name			= "MA Needy Child"													'needs spaces to generate button width properly.
script_array_IVE_main(script_num).file_name				= "ive-ma-needy-child.vbs"
script_array_IVE_main(script_num).description			= "Case note template for MA Needy Child cases."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_IVE_main(script_num)			'Resets the array to add one more element to it
Set script_array_IVE_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_IVE_main(script_num).script_name 			= "Permanency Review"																		'Script name
script_array_IVE_main(script_num).file_name 			= "ive-perm-review.vbs"															'Script URL
script_array_IVE_main(script_num).description 			= "Case note template for IV-E permanency review cases."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_IVE_main(script_num)			'Resets the array to add one more element to it
Set script_array_IVE_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_IVE_main(script_num).script_name			= "Trial Home Visit"													'needs spaces to generate button width properly.
script_array_IVE_main(script_num).file_name				= "ive-trial-home-visit.vbs"
script_array_IVE_main(script_num).description			= "Case note template for Trial Home Visit cases."


'Starting these with a very high number, higher than the normal possible amount of buttons.
'	We're doing this because we want to assign a value to each button pressed, and we want
'	that value to change with each button. The button_placeholder will be placed in the .button
'	property for each script item. This allows it to both escape the Function and resize
'	near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'	to force the value of ButtonPressed to hold in near infinite iterations.
button_placeholder 	= 24601
'IVE_main_button		= 1000
'SNAP_WCOMS_button	= 2000


'Displays the dialog
Do
	If ButtonPressed = "" or ButtonPressed = IVE_main_button then declare_IVE_menu_dialog(script_array_IVE_main)
	dialog IVE_dialog
	If ButtonPressed = 0 then stopscript

    'Opening the SIR Instructions
	'IF buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/Notices%20scripts.aspx")
Loop until 	ButtonPressed <> IVE_main_button
'MsgBox buttonpressed = script_array_IVE_main(0).button

'Runs through each script in the array... if the selected script (buttonpressed) is in the array, it'll run_from_GitHub
For i = 0 to ubound(script_array_IVE_main)
	If ButtonPressed = script_array_IVE_main(i).button then call run_from_GitHub(script_repository & "iv-e/" & script_array_IVE_main(i).file_name)
Next

'For i = 0 to ubound(script_array_IVE_list)
'	If ButtonPressed = script_array_IVE_list(i).button then call run_from_GitHub(script_repository & "iv-e/" & script_array_IVE_list(i).file_name)
'Next

stopscript
