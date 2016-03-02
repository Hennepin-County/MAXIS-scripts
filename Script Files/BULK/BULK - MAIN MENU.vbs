'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - MAIN MENU.vbs"
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
Function declare_BULK_menu_dialog(script_array)
	BeginDialog BULK_dialog, 0, 0, 516, 340, "BULK Scripts"
	 	Text 5, 5, 435, 10, "Bulk scripts main menu: select the script to run from the choices below."
	  	ButtonGroup ButtonPressed
		 	PushButton 015, 35, 30, 15, "BULK", 				BULK_main_button
		 	PushButton 045, 35, 50, 15, "BULK LISTS", 				BULK_lists_button
		 	PushButton 445, 10, 65, 10, "SIR instructions", 	SIR_instructions_button

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
		GroupBox 5, 20, 205, 35, "BULK Sub-Menus"
	EndDialog
End function
'END CUSTOM FUNCTIONS=======================================================================================================

'VARIABLES TO DECLARE=======================================================================================================

'Declaring the variable names to cut down on the number of arguments that need to be passed through the function.
DIM ButtonPressed
DIM SIR_instructions_button
dim BULK_dialog

script_array_BULK_main = array()
script_array_BULK_list = array()


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

'-------------------------------------------------------------------------------------------------------------------------BULK MAIN MENU

'Resetting the variable
script_num = 0
ReDim Preserve script_array_BULK_main(script_num)
Set script_array_BULK_main(script_num) = new script
script_array_BULK_main(script_num).script_name 			= "CASE/NOTE from List"																		'Script name
script_array_BULK_main(script_num).file_name 				= "BULK - CASE NOTE FROM LIST.vbs"															'Script URL
script_array_BULK_main(script_num).description 			= "Creates the same case note on cases listed in REPT/ACTV, manually entered, or from an Excel spreadsheet of your choice."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_main(script_num)			'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name 			= "Case Transfer"																		'Script name
script_array_BULK_main(script_num).file_name 				= "BULK - CASE TRANSFER.vbs"															'Script URL
script_array_BULK_main(script_num).description 			= "Searches caseload(s) by selected parameters. Transfers a specified number of those cases to another worker. Creates list of these cases."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_main(script_num)			'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name				= "CEI Premium Noter"
script_array_BULK_main(script_num).file_name				= "BULK - CEI PREMIUM NOTER.vbs"
script_array_BULK_main(script_num).description				= "Case notes recurring CEI premiums on multiple cases simultaneously."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_main(script_num)			'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name				= "COLA Auto-approved Dail Noter"
script_array_BULK_main(script_num).file_name				= "BULK - COLA AUTO APPROVED DAIL NOTER.vbs"
script_array_BULK_main(script_num).description				= "Case notes all cases on DAIL/DAIL with Auto-approved COLA message, creates list of these messages, deletes the DAIL."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_main(script_num)			'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name				= "INAC Scrubber"
script_array_BULK_main(script_num).file_name				= "BULK - INAC SCRUBBER.vbs"
script_array_BULK_main(script_num).description				= "Checks cases on REPT/INAC (for criteria see SIR) case notes if passes criteria, and transfers if agency uses closed-file worker number. "

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_main(script_num)			'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name				= "MEMO from List"
script_array_BULK_main(script_num).file_name				= "BULK - MEMO FROM LIST.vbs"
script_array_BULK_main(script_num).description				= "Creates the same MEMO on cases listed in REPT/ACTV, manually entered, or from an Excel spreadsheet of your choice."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_main(script_num)			'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name				= "Returned Mail"
script_array_BULK_main(script_num).file_name				= "BULK - RETURNED MAIL.vbs"
script_array_BULK_main(script_num).description				= "Case notes that returned mail (without a forwarding address) was received for up to 60 cases, TIKLs for 10-day return."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_main(script_num)			'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name				= " REVW/MONT Closures "													'needs spaces to generate button width properly.
script_array_BULK_main(script_num).file_name				= "BULK - REVW-MONT CLOSURES.vbs"
script_array_BULK_main(script_num).description				= "Case notes all cases on REPT/REVW or REPT/MONT that are closing for missing or incomplete CAF/HRF/CSR/HC ER."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_main(script_num)			'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name				= "TIKL from List"
script_array_BULK_main(script_num).file_name				= "BULK - TIKL FROM LIST.vbs"
script_array_BULK_main(script_num).description				= "Creates the same TIKL on cases listed in REPT/ACTV, manually entered, or from an Excel spreadsheet of your choice."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_main(script_num)			'Resets the array to add one more element to it
Set script_array_BULK_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_main(script_num).script_name				= "Update EOMC List"
script_array_BULK_main(script_num).file_name				= "BULK - UPDATE EOMC LIST.vbs"
script_array_BULK_main(script_num).description				= "Updates a saved REPT/EOMC excel file from previous month with current case status."

'-------------------------------------------------------------------------------------------------------------------------BULK LISTS
'Resetting the variable
script_num = 0
ReDim Preserve script_array_BULK_list(script_num)
Set script_array_BULK_list(script_num) = new script
script_array_BULK_list(script_num).script_name 			= "ADDR"																		'Script name
script_array_BULK_list(script_num).file_name			= "BULK - ADDRESS REPORT.vbs"
script_array_BULK_list(script_num).description 			= "Creates a list of all addresses from a caseload(or entire county)."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= "ACTV"
script_array_BULK_list(script_num).file_name			= "BULK - REPT-ACTV LIST.vbs"
script_array_BULK_list(script_num).description 			= "Pulls a list of cases in REPT/ACTV into an Excel spreadsheet."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= "ARST"
script_array_BULK_list(script_num).file_name			= "BULK - REPT-ARST LIST.vbs"
script_array_BULK_list(script_num).description 			= "Pulls a list of cases in REPT/ARST into an Excel spreadsheet."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= " Check SNAP for GA/RCA "													'needs spaces to generate button width properly.
script_array_BULK_list(script_num).file_name			= "BULK - CHECK SNAP FOR GA RCA.vbs"
script_array_BULK_list(script_num).description 			= "Compares the amount of GA and RCA FIAT'd into SNAP and creates a list of the results."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= "DAIL"
script_array_BULK_list(script_num).file_name			= "BULK - DAIL REPORT.vbs"
script_array_BULK_list(script_num).description 			= "Pulls a list of DAILS in DAIL/DAIL into an Excel spreadsheet."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= " EOMC "													'needs spaces to generate button width properly.
script_array_BULK_list(script_num).file_name			= "BULK - REPT-EOMC LIST.vbs"
script_array_BULK_list(script_num).description 			= "Pulls a list of cases in REPT/EOMC into an Excel spreadsheet."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= "Find Updated Panels"
script_array_BULK_list(script_num).file_name			= "BULK - FIND PANEL UPDATE DATE.vbs"
script_array_BULK_list(script_num).description 			= "Creates a list of cases from a caseload(s) showing when selected panels have been updated."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= "INAC"
script_array_BULK_list(script_num).file_name			= "BULK - REPT-INAC LIST.vbs"
script_array_BULK_list(script_num).description 			= "Pulls a list of cases in REPT/INAC into an Excel spreadsheet."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= "LTC-GRH List Gen"
script_array_BULK_list(script_num).file_name			= "BULK - LTC-GRH LIST GENERATOR.vbs"
script_array_BULK_list(script_num).description 			= "Creates a list of FACIs, AREPs, and waiver types assigned to the various cases in a caseload(s)."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= "MA-EPD/Medi Pt B CEI"
script_array_BULK_list(script_num).file_name			= "BULK - FIND MAEPD MEDI CEI.vbs"
script_array_BULK_list(script_num).description 			= "Creates a list of cases and clients active on MA-EPD and Medicare Part B that are eligible for Part B reimbursement."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= " MFCM "													'needs spaces to generate button width properly.
script_array_BULK_list(script_num).file_name			= "BULK - REPT-MFCM LIST.vbs"
script_array_BULK_list(script_num).description 			= "Pulls a list of cases in REPT/MFCM into an Excel spreadsheet."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= "Non-MAGI HC Info"
script_array_BULK_list(script_num).file_name			= "BULK - NON-MAGI HC INFO.vbs"
script_array_BULK_list(script_num).description 			= "Creates a list of cases with non-MAGI HC/PDED information."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= "PND1"
script_array_BULK_list(script_num).file_name			= "BULK - REPT-PND1 LIST.vbs"
script_array_BULK_list(script_num).description 			= "Pulls a list of cases in REPT/PND1 into an Excel spreadsheet."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= "PND2"
script_array_BULK_list(script_num).file_name			= "BULK - REPT-PND2 LIST.vbs"
script_array_BULK_list(script_num).description 			= "Pulls a list of cases in REPT/PND2 into an Excel spreadsheet."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= "POLI TEMP LIST"
script_array_BULK_list(script_num).file_name			= "BULK - POLI TEMP LIST.vbs"
script_array_BULK_list(script_num).description 			= "Creates a list of current POLI TEMP topics, the TEMP reference and the revised date."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= " REVS "													'needs spaces to generate button width properly.
script_array_BULK_list(script_num).file_name			= "BULK - REPT-REVS LIST.vbs"
script_array_BULK_list(script_num).description 			= "Pulls a list of cases in REPT/REVS into an Excel spreadsheet."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= " REVW "													'needs spaces to generate button width properly.
script_array_BULK_list(script_num).file_name			= "BULK - REPT-REVW LIST.vbs"
script_array_BULK_list(script_num).description 			= "Pulls a list of cases in REPT/REVW into an Excel spreadsheet."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_BULK_list(script_num)		'Resets the array to add one more element to it
Set script_array_BULK_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_BULK_list(script_num).script_name 			= "SWKR List Gen"
script_array_BULK_list(script_num).file_name			= "BULK - SWKR LIST GENERATOR.vbs"
script_array_BULK_list(script_num).description 			= "Creates a list of SWKRs assigned to the various cases in a caseload(s)."



'Starting these with a very high number, higher than the normal possible amount of buttons.
'	We're doing this because we want to assign a value to each button pressed, and we want
'	that value to change with each button. The button_placeholder will be placed in the .button
'	property for each script item. This allows it to both escape the Function and resize
'	near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'	to force the value of ButtonPressed to hold in near infinite iterations.
button_placeholder 	= 24601
BULK_main_button		= 1000
BULK_lists_button		= 2000


'Displays the dialog
Do
	If ButtonPressed = "" or ButtonPressed = BULK_main_button then
		declare_BULK_menu_dialog(script_array_BULK_main)
	ElseIf ButtonPressed = BULK_lists_button then
		declare_BULK_menu_dialog(script_array_BULK_list)
	End if

	dialog BULK_dialog

	If ButtonPressed = 0 then stopscript
    'Opening the SIR Instructions
	IF buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/Bulk%20scripts.aspx")
Loop until 	ButtonPressed <> SIR_instructions_button and _
			ButtonPressed <> BULK_main_button and _
			ButtonPressed <> BULK_lists_button

'MsgBox buttonpressed = script_array_BULK_main(0).button

'Runs through each script in the array... if the selected script (buttonpressed) is in the array, it'll run_from_GitHub
For i = 0 to ubound(script_array_BULK_main)
	If ButtonPressed = script_array_BULK_main(i).button then call run_from_GitHub(script_repository & "/BULK/" & script_array_BULK_main(i).file_name)
Next

For i = 0 to ubound(script_array_BULK_list)
	If ButtonPressed = script_array_BULK_list(i).button then call run_from_GitHub(script_repository & "/BULK/" & script_array_BULK_list(i).file_name)
Next

stopscript
