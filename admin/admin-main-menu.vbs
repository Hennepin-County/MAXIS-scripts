'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - MAIN MENU.vbs"
start_time = timer

''LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
'IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
'	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
'		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
'			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
'		Else											'Everyone else should use the release branch.
'			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
'		End if
'		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
'		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
'		req.send													'Sends request
'		IF req.Status = 200 THEN									'200 means great success
'			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
'			Execute req.responseText								'Executes the script code
'		ELSE														'Error message
'			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
'                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
'                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
'                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
'            StopScript
'		END IF
'	ELSE
'		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
'		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
'		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
'		text_from_the_other_script = fso_command.ReadAll
'		fso_command.Close
'		Execute text_from_the_other_script
'	END IF
'END IF
''END FUNCTIONS LIBRARY BLOCK================================================================================================
'
'LOADING FUNC LIB
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Scripts\Script Files\MASTER FUNCTIONS LIBRARY.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/30/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CUSTOM FUNCTIONS===========================================================================================================
Function declare_admin_menu_dialog(script_array)
	BeginDialog admin_dialog, 0, 0, 516, 280, "Admin Scripts"
	 	Text 5, 5, 435, 10, "Admin scripts main menu: select the script to run from the choices below."
	  	ButtonGroup ButtonPressed
		 	'PushButton 015, 35, 40, 15, "CA", 				admin_main_button
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

		CancelButton 455, 260, 50, 15
		'GroupBox 5, 20, 205, 35, "admin Sub-Menus"
	EndDialog
End function
'END CUSTOM FUNCTIONS=======================================================================================================

'VARIABLES TO DECLARE=======================================================================================================

'Declaring the variable names to cut down on the number of arguments that need to be passed through the function.
DIM ButtonPressed
'DIM SIR_instructions_button
dim admin_dialog

script_array_admin_main = array()
'script_array_admin_list = array()


'END VARIABLES TO DECLARE===================================================================================================

'LIST OF SCRIPTS================================================================================================================

'INSTRUCTIONS: simply add your new script below. Scripts are listed in alphabetical order. Copy a block of code from above and paste your script info in. The function does the rest.

'-------------------------------------------------------------------------------------------------------------------------admin MAIN MENU

'Resetting the variable
script_num = 0
ReDim Preserve script_array_admin_main(script_num)
Set script_array_admin_main(script_num) = new script
script_array_admin_main(script_num).script_name 		= " ABAWD Report "											'Script name
script_array_admin_main(script_num).file_name 			= "abawd-report.vbs"										'Script URL
script_array_admin_main(script_num).description 		= "BULK script that gathers ABAWD/FSET codes for members on SNAP/MFIP active cases."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)
Set script_array_admin_main(script_num) = new script
script_array_admin_main(script_num).script_name 		= "Auto-Dialer Case Status"											'Script name
script_array_admin_main(script_num).file_name 			= "auto-dialer-case-status.vbs"										'Script URL
script_array_admin_main(script_num).description 		= "BULK script that gathers case status for cases with recerts for SNAP/MFIP the previous month."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)
Set script_array_admin_main(script_num) = new script
script_array_admin_main(script_num).script_name 		= "Banked Months Closure"											'Script name
script_array_admin_main(script_num).file_name 			= "banked-months-closure.vbs"										'Script URL
script_array_admin_main(script_num).description 		= "Completes closure case notes and applicable worker comments for banked months cases."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "Carl Discrepancy Report"													'Script name
script_array_admin_main(script_num).file_name			= "carl-discrepancy-report.vbs"												'Script URL
script_array_admin_main(script_num).description			= "Processes a list of cases that are missing a CARL assignment."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "DAIL Decimator"													'Script name
script_array_admin_main(script_num).file_name			= "dail-decimator.vbs"												'Script URL
script_array_admin_main(script_num).description			= "BULK script that deletes specific DAILS based on content, and collects them into an Excel spreadsheet."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "DAIL report"														'Script name
script_array_admin_main(script_num).file_name			= "dail-report.vbs"													'Script URL
script_array_admin_main(script_num).description			= "Specialty BULK DAIL report for use by the Quality Improvement team."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "DISA Dr. PEPR"													'Script name
script_array_admin_main(script_num).file_name			= "disa-dr-pepr.vbs"												'Script URL
script_array_admin_main(script_num).description			= "Adds additional information to an existing list of cases applicable to DAIL PEPR DAILS."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "Interview Required"													'Script name
script_array_admin_main(script_num).file_name			= "interview-required.vbs"												'Script URL
script_array_admin_main(script_num).description			= "BULK script to collect case information for cases that require an interview for SNAP/MFIP."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "Language Stats"													'Script name
script_array_admin_main(script_num).file_name			= "language-stats.vbs"												'Script URL
script_array_admin_main(script_num).description			= "Collects language statistics by language and region. Take approximately 10 hours to run."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "MFIP Sanction FIATer"											'Script name
script_array_admin_main(script_num).file_name			= "mfip-sanction-fiater.vbs"										'Script URL
script_array_admin_main(script_num).description			= "FIATs MFIP sanction actions for CS, ES and both types of sanctions."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "Pull Cases Into Excel"											'Script name
script_array_admin_main(script_num).file_name			= "pull-cases-into-excel.vbs"										'Script URL
script_array_admin_main(script_num).description			= "Creates a list of information not available in other BULK scripts."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "Sanction Member Info"										'Script name
script_array_admin_main(script_num).file_name			= "sanction-member-info.vbs"									'Script URL
script_array_admin_main(script_num).description			= "BULK script to gather information for for MFIP participants on REPT/MFCM."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "Send CBO Manual Referrals"										'Script name
script_array_admin_main(script_num).file_name			= "send-cbo-manual-referrals.vbs"									'Script URL
script_array_admin_main(script_num).description			= "Sends manual referrals for a list of cases provided by Employment and Training."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "UNEA Updater"										'Script name
script_array_admin_main(script_num).file_name			= "unea-updater.vbs"									'Script URL
script_array_admin_main(script_num).description			= "BULK script that updates UNEA information and sends SPEC/MEMO for VA cases at ER."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "Update Banked Month Case Review List"										'Script name
script_array_admin_main(script_num).file_name			= "update-BM-case-review-list.vbs"									'Script URL
script_array_admin_main(script_num).description			= "BULK script that updates monthly banked month statuses on INAC and most EXEMPT cases."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_admin_main(script_num)		'Resets the array to add one more element to it
Set script_array_admin_main(script_num) = new script	'Set this array element to be a new script. Script details below...
script_array_admin_main(script_num).script_name			= "WF1 Case Status"													'Script name
script_array_admin_main(script_num).file_name			= "wf1-case-status.vbs"												'Script URL
script_array_admin_main(script_num).description			= "Updates a list of cases from Excel with current case and ABAWD status inforamtion."


'Starting these with a very high number, higher than the normal possible amount of buttons.
'	We're doing this because we want to assign a value to each button pressed, and we want
'	that value to change with each button. The button_placeholder will be placed in the .button
'	property for each script item. This allows it to both escape the Function and resize
'	near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'	to force the value of ButtonPressed to hold in near infinite iterations.
button_placeholder 	= 24601
'admin_main_button		= 1000
'SNAP_WCOMS_button	= 2000


'Displays the dialog
Do
	If ButtonPressed = "" or ButtonPressed = admin_main_button then declare_admin_menu_dialog(script_array_admin_main)
	dialog admin_dialog
	If ButtonPressed = 0 then stopscript
	
    'Opening the SIR Instructions
	'IF buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/Notices%20scripts.aspx")
Loop until 	ButtonPressed <> admin_main_button 
'MsgBox buttonpressed = script_array_admin_main(0).button

'Runs through each script in the array... if the selected script (buttonpressed) is in the array, it'll run_from_GitHub
For i = 0 to ubound(script_array_admin_main)
	If ButtonPressed = script_array_admin_main(i).button then call run_from_GitHub(script_repository & "admin/" & script_array_admin_main(i).file_name)
Next

'For i = 0 to ubound(script_array_admin_list)
'	If ButtonPressed = script_array_admin_list(i).button then call run_from_GitHub(script_repository & "Admin/" & script_array_admin_list(i).file_name)
'Next

stopscript
