'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - MAIN MENU.vbs"
start_time = timer 

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
			MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_
					vbCr & _
					"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
					vbCr & _
					"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
					vbTab & "- The name of the script you are running." & vbCr &_
					vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
					vbTab & "- The name and email for an employee from your IT department," & vbCr & _
					vbTab & vbTab & "responsible for network issues." & vbCr &_
					vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
					vbCr & _
					"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_
					vbCr &_
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================


'CUSTOM FUNCTIONS===========================================================================================================
Function declare_ACTIONS_menu_dialog(script_array)
	BeginDialog ACTIONS_dialog, 0, 0, 516, 340, "NOTICES Scripts"
	 	Text 5, 5, 435, 10, "Notices scripts main menu: select the script to run from the choices below."
	  	ButtonGroup ButtonPressed
		 	PushButton 015, 35, 40, 15, "ACTIONS", 				ACTIONS_main_button
		 	'PushButton 055, 35, 60, 15, "Additional Actions", 				Addtional_actions_button    'to be used in the future when a split is needed
		 	PushButton 445, 10, 65, 10, "SIR instructions", 	SIR_instructions_button
			PushButton 395, 10, 45, 10, "UTILITIES",            UTILITIES_SCRIPTS_button
		'This starts here, but it shouldn't end here :)
		vert_button_position = 70

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

		CancelButton 460, 320, 50, 15
		GroupBox 5, 20, 205, 35, "NOTICES Sub-Menus"
	EndDialog
End function
'END CUSTOM FUNCTIONS=======================================================================================================

'VARIABLES TO DECLARE=======================================================================================================

'Declaring the variable names to cut down on the number of arguments that need to be passed through the function.
DIM ButtonPressed
DIM UTILITIES_SCRIPTS_button
DIM SIR_instructions_button
dim ACTIONS_dialog

script_array_ACTIONS_main = array()
script_array_ACTIONS_list = array()


'END VARIABLES TO DECLARE===================================================================================================

'CLASSES TO DEFINE==========================================================================================================

'A class for each script item
class script

	public script_name
	public file_name
	public description
	public button

	public property get button_size	'This part determines the size of the button dynamically by determining the length of the script name, multiplying that by 3.5, rounding the decimal off, and adding 10 px
		button_size = round ( len( script_name ) * 3.5 ) + 10
	end property

end class

'END CLASSES TO DEFINE==========================================================================================================

'LIST OF SCRIPTS================================================================================================================

'INSTRUCTIONS: simply add your new script below. Scripts are listed in alphabetical order. Copy a block of code from above and paste your script info in. The function does the rest.

'-------------------------------------------------------------------------------------------------------------------------ACTIONS MAIN MENU

'Resetting the variable
script_num = 0
ReDim Preserve script_array_ACTIONS_main(script_num)
Set script_array_ACTIONS_main(script_num) = new script
script_array_ACTIONS_main(script_num).script_name 			= "   ABAWD BANKED MONTHS FIATer   "																		'Script name
script_array_ACTIONS_main(script_num).file_name 			= "ACTIONS - ABAWD BANKED MONTHS FIATER.vbs"															'Script URL
script_array_ACTIONS_main(script_num).description 			= "FIATS SNAP eligibility, income and deductions for HH members using banked months."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name 			= " ABAWD FSET Exemption Check "																		'Script name
script_array_ACTIONS_main(script_num).file_name 			= "ACTIONS - ABAWD FSET EXEMPTION CHECK.vbs"															'Script URL
script_array_ACTIONS_main(script_num).description 			= "Double checks a case to see if any possible ABAWD/FSET exemptions exist."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= " ABAWD Screening Tool"													'needs spaces to generate button width properly.
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - ABAWD SCREENING TOOL.vbs"
script_array_ACTIONS_main(script_num).description			= "A tool to walk through a screening to determine if client is ABAWD."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= "BILS Updater"
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - BILS UPDATER.vbs"
script_array_ACTIONS_main(script_num).description			= "Updates a BILS panel with reoccurring or actual BILS received."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= " Check EDRS"
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - CHECK EDRS.vbs"
script_array_ACTIONS_main(script_num).description			= "Checks EDRS for HH members with disqualifications on a case."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= "Copy Panels to Word"													'needs spaces to generate button width properly.
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - COPY PANELS TO WORD.vbs"
script_array_ACTIONS_main(script_num).description			= "Copies MAXIS panels to Word en masse for a case for easier review."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= " FSET SANCTION "
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - FSET SANCTION.vbs"
script_array_ACTIONS_main(script_num).description			= "Updates the WREG panel, and case notes when imposing or resolving a FSET sanction."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= "   HG MONY/CHCK ISSUANCE   "													'needs spaces to generate button width properly.
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - HOUSING GRANT MONY CHCK ISSUANCE.vbs"
script_array_ACTIONS_main(script_num).description			= "-- New 04/2016!!! Issues a housing grant in MONY/CHCK for cases that should have been issued in prior months."


script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= "LTC Spousal Allocation FIATer"													'needs spaces to generate button width properly.
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - LTC - SPOUSAL ALLOCATION FIATER.vbs"
script_array_ACTIONS_main(script_num).description			= "FIATs a spousal allocation across a budget period."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= " MA-EPD Earned Income FIATer "
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - MA-EPD EI FIAT.vbs"
script_array_ACTIONS_main(script_num).description			= "FIATs MA-EPD earned income (JOBS income) to be even across an entire budget period."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= "New Job Reported"
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - NEW JOB REPORTED.vbs"
script_array_ACTIONS_main(script_num).description			= "Creates a JOBS panel, CASE/NOTE and TIKL when a new job is reported. Use the DAIL scrubber for new hire DAILs."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= "PA Verif Request"
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - PA VERIF REQUEST.vbs"
script_array_ACTIONS_main(script_num).description			= "Creates a Word document with PA benefit totals for other agencies to determine client benefits."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= "Paystubs Recieved"
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - PAYSTUBS RECEIVED.vbs"
script_array_ACTIONS_main(script_num).description			= "Enter in pay stubs, and puts it on JOBS (both retro & pro if applicable), as well as the PIC and HC pop-up, and case note."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= "Shelter Expense Verif Recv'd"
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - SHELTER EXPENSE VERIF RECEIVED.vbs"
script_array_ACTIONS_main(script_num).description			= "Enter shelter expense/address information in a dialog and the script updates SHEL, HEST, and ADDR and case notes."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= "Send SVES"
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - SEND SVES.vbs"
script_array_ACTIONS_main(script_num).description			= "Sends a SVES/QURY."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= "Transfer Case"
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - TRANSFER CASE.vbs"
script_array_ACTIONS_main(script_num).description			= "SPEC/XFERs a case, and can send a client memo. For in-agency as well as out-of-county XFERs."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_ACTIONS_main(script_num)		'Resets the array to add one more element to it
Set script_array_ACTIONS_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_ACTIONS_main(script_num).script_name			= "TYMA TIKLer"
script_array_ACTIONS_main(script_num).file_name				= "ACTIONS - TYMA TIKLER.vbs"
script_array_ACTIONS_main(script_num).description			= "TIKLS for TYMA report forms to be sent."



'-------------------------------------------------------------------------------------------------------------------------Additional ACTIONS scripts   'to be used in the future when a non-alpha split is needed MUST UNCOMMENT ALL RELATED REFERENCES
'Resetting the variable
'script_num = 0
'ReDim Preserve script_array_ACTIONS_list(script_num)
'Set script_array_ACTIONS_list(script_num) = new script
'script_array_ACTIONS_list(script_num).script_name 			= " ABAWD with Child in HH WCOM "'needs spaces to generate button width properly.				'Script name
'script_array_ACTIONS_list(script_num).file_name			= "NOTICES - ABAWD WITH CHILD IN HH WCOM.vbs"
'script_array_ACTIONS_list(script_num).description 			= "Adds a WCOM to a notice for an ABAWD adult receiving child under 18 exemption."
'
'script_num = script_num + 1								'Increment by one
'ReDim Preserve script_array_ACTIONS_list(script_num)		'Resets the array to add one more element to it
'Set script_array_ACTIONS_list(script_num) = new script		'Set this array element to be a new script. Script details below...
'script_array_ACTIONS_list(script_num).script_name 			= "  Banked Month WCOMS "
'script_array_ACTIONS_list(script_num).file_name			= "NOTICES - BANKED MONTH WCOMS.vbs"
'script_array_ACTIONS_list(script_num).description 			= "Adds various WCOMS to a notice for regarding banked month approvals/closure."
'
'script_num = script_num + 1								'Increment by one
'ReDim Preserve script_array_ACTIONS_list(script_num)		'Resets the array to add one more element to it
'Set script_array_ACTIONS_list(script_num) = new script		'Set this array element to be a new script. Script details below...
'script_array_ACTIONS_list(script_num).script_name 			= "Duplicate assistance WCOM"
'script_array_ACTIONS_list(script_num).file_name			= "NOTICES - DUPLICATE ASSISTANCE WCOM.vbs"
'script_array_ACTIONS_list(script_num).description 			= "Adds a WCOM to a notice for duplicate assistance explaining why the client was ineligible."
'
'script_num = script_num + 1								'Increment by one
'ReDim Preserve script_array_ACTIONS_list(script_num)		'Resets the array to add one more element to it
'Set script_array_ACTIONS_list(script_num) = new script		'Set this array element to be a new script. Script details below...
'script_array_ACTIONS_list(script_num).script_name 			= "Postponed WREG Verif"
'script_array_ACTIONS_list(script_num).file_name			= "NOTICES - POSTPONED WREG VERIFS.vbs"
'script_array_ACTIONS_list(script_num).description 			= "Sends a WCOM informing the client of postponed verifications that MAXIS won't add to notice correctly by itself."




'Starting these with a very high number, higher than the normal possible amount of buttons.
'	We're doing this because we want to assign a value to each button pressed, and we want
'	that value to change with each button. The button_placeholder will be placed in the .button
'	property for each script item. This allows it to both escape the Function and resize
'	near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'	to force the value of ButtonPressed to hold in near infinite iterations.
button_placeholder 	= 24601
ACTIONS_main_button		= 1000
'Addtional_actions_button		= 2000



'Displays the dialog
Do
	If ButtonPressed = "" or ButtonPressed = ACTIONS_main_button then
		declare_ACTIONS_menu_dialog(script_array_ACTIONS_main)
	'ElseIf ButtonPressed = Addtional_actions_button then
	'	declare_ACTIONS_menu_dialog(script_array_ACTIONS_list)
	End if

	dialog ACTIONS_dialog

	If ButtonPressed = 0 then stopscript
    'Opening the SIR Instructions
	IF buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/Actions%20scripts.aspx")
	'Opening utilities
	If ButtonPressed = UTILITIES_SCRIPTS_button then 
		call run_from_GitHub(script_repository & "/UTILITIES/UTILITIES - MAIN MENU.vbs")
		stopscript
	END IF
Loop until 	ButtonPressed <> SIR_instructions_button and _
			ButtonPressed <> ACTIONS_main_button 'and _
			'ButtonPressed <> Addtional_actions_button

'MsgBox buttonpressed = script_array_ACTIONS_main(0).button

'Runs through each script in the array... if the selected script (buttonpressed) is in the array, it'll run_from_GitHub
For i = 0 to ubound(script_array_ACTIONS_main)
	If ButtonPressed = script_array_ACTIONS_main(i).button then call run_from_GitHub(script_repository & "/ACTIONS/" & script_array_ACTIONS_main(i).file_name)
Next

For i = 0 to ubound(script_array_ACTIONS_list)
	If ButtonPressed = script_array_ACTIONS_list(i).button then call run_from_GitHub(script_repository & "/ACTIONS/" & script_array_ACTIONS_list(i).file_name)
Next

stopscript





