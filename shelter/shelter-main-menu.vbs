'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "SHELTER - MAIN MENU.vbs"
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
call changelog_update("06/04/2021", "Retired GRH APPROVAL and SINGLE CLIENT INTERVIEW scripts.", "Ilse Ferris, Hennepin County")
call changelog_update("07/05/2018", "Updates to add scripts per shelter team request..", "MiKayla Handley")
call changelog_update("01/05/2018", "Updates to CES-Screening Appt per shelter team request..", "MiKayla Handley")
call changelog_update("09/23/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CUSTOM FUNCTIONS===========================================================================================================
Function declare_NOTES_menu_dialog(script_array)
	BeginDialog NOTES_dialog, 0, 0, 516, 340, "Shelter Team Scripts"
	 	Text 5, 5, 435, 10, "Shelter scripts main menu: select the script to run from the choices below."
	  	ButtonGroup ButtonPressed
		 	PushButton 015, 35, 30, 15, "# - L", 				a_to_n_button
		 	PushButton 045, 35, 30, 15, "M - Z", 				p_to_z_button

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
		GroupBox 5, 20, 85, 35, "Shelter Sub-Menus"
	EndDialog
End function
'END CUSTOM FUNCTIONS=======================================================================================================

'VARIABLES TO DECLARE=======================================================================================================

'Declaring the variable names to cut down on the number of arguments that need to be passed through the function.
DIM ButtonPressed
'DIM SIR_instructions_button
dim NOTES_dialog

script_array_a_to_n = array()
script_array_p_to_z = array()


'END VARIABLES TO DECLARE===================================================================================================

'LIST OF SCRIPTS================================================================================================================

'INSTRUCTIONS: simply add your new script below. Scripts are listed in alphabetical order. Copy a block of code from above and paste your script info in. The function does the rest.

'-------------------------------------------------------------------------------------------------------------------------A through M
'Resetting the variable
script_num = 0
ReDim Preserve script_array_a_to_n(script_num)
Set script_array_a_to_n(script_num) = new script
script_array_a_to_n(script_num).script_name 			= "2 PM Return"																				'Script name
script_array_a_to_n(script_num).file_name 				= "shelter-2-pm-return.vbs"																	'Script URL
script_array_a_to_n(script_num).description 			= "Case note template for documenting details for the 2 PM return process."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_a_to_n(script_num)			'Resets the array to add one more element to it
Set script_array_a_to_n(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_a_to_n(script_num).script_name 			= "311"																						'Script name
script_array_a_to_n(script_num).file_name 				= "shelter-311.vbs"																			'Script URL
script_array_a_to_n(script_num).description 			= "Case note template for documenting 311 information."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_a_to_n(script_num)			'Resets the array to add one more element to it
Set script_array_a_to_n(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_a_to_n(script_num).script_name 			= "ACF Request Pend"																		'Script name
script_array_a_to_n(script_num).file_name 				= "shelter-acf-request-pend.vbs"															'Script URL
script_array_a_to_n(script_num).description 			= "Case note template for documenting details for a ACF pending request."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_a_to_n(script_num)			'Resets the array to add one more element to it
Set script_array_a_to_n(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_a_to_n(script_num).script_name 			= "ACF Used"																				'Script name
script_array_a_to_n(script_num).file_name 				= "shelter-acf-used.vbs"																	'Script URL
script_array_a_to_n(script_num).description 			= "Case note template for documenting details for ACF used."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_a_to_n(script_num)			'Resets the array to add one more element to it
Set script_array_a_to_n(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_a_to_n(script_num).script_name 			= "Bus Ticket Issued"																		'Script name
script_array_a_to_n(script_num).file_name 				= "shelter-bus-ticket-issued.vbs"															'Script URL
script_array_a_to_n(script_num).description 			= "Case note template for documenting details for issuing bus tickets."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_a_to_n(script_num)			'Resets the array to add one more element to it
Set script_array_a_to_n(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_a_to_n(script_num).script_name 			= "Cash Cut-off"																			'Script name
script_array_a_to_n(script_num).file_name 				= "shelter-cash-cut-off.vbs"																'Script URL
script_array_a_to_n(script_num).description 			= "Case note template for documenting details for cash cut-off."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_a_to_n(script_num)			'Resets the array to add one more element to it
Set script_array_a_to_n(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_a_to_n(script_num).script_name 			= "CES Screening Referral"																			'Script name
script_array_a_to_n(script_num).file_name 				= "shelter-ces-screening-appt.vbs"																'Script URL
script_array_a_to_n(script_num).description 			= "Case note template for documenting details for the CES screening appointment."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_a_to_n(script_num)			'Resets the array to add one more element to it
Set script_array_a_to_n(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_a_to_n(script_num).script_name 			= "Client Sheltered by Win A"																'Script name
script_array_a_to_n(script_num).file_name 				= "shelter-client-sheltered-by-win-a.vbs"													'Script URL
script_array_a_to_n(script_num).description 			= "Case note template for documenting details of contact with client at Window A."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_a_to_n(script_num)			'Resets the array to add one more element to it
Set script_array_a_to_n(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_a_to_n(script_num).script_name 			= "Change Reported"																'Script name
script_array_a_to_n(script_num).file_name 				= "shelter-change-reported.vbs"													'Script URL
script_array_a_to_n(script_num).description 			= "Case note template for documenting details of change reported by client."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_a_to_n(script_num)			'Resets the array to add one more element to it
Set script_array_a_to_n(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_a_to_n(script_num).script_name 			= "Diversion Program Referral"																'Script name
script_array_a_to_n(script_num).file_name 				= "shelter-diversion-program-referral.vbs"													'Script URL
script_array_a_to_n(script_num).description 			= "Case note template for documenting details of Diversion Program Referral."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_a_to_n(script_num)			'Resets the array to add one more element to it
Set script_array_a_to_n(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_a_to_n(script_num).script_name 			= "Diversion Program Referral Result"																'Script name
script_array_a_to_n(script_num).file_name 				= "shelter-diversion-program-referral-results.vbs"													'Script URL
script_array_a_to_n(script_num).description 			= "Case note template for documenting details of Diversion Program Referral result."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_a_to_n(script_num)			'Resets the array to add one more element to it
Set script_array_a_to_n(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_a_to_n(script_num).script_name 			= "EA Approved"																				'Script name
script_array_a_to_n(script_num).file_name 				= "shelter-ea-approved.vbs"																	'Script URL
script_array_a_to_n(script_num).description 			= "Case note template for documenting details for the EA approval."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_a_to_n(script_num)			'Resets the array to add one more element to it
Set script_array_a_to_n(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_a_to_n(script_num).script_name 			= "EA Extension"																			'Script name
script_array_a_to_n(script_num).file_name 				= "shelter-ea-extension.vbs"																'Script URL
script_array_a_to_n(script_num).description 			= "Case note template for documenting details for the EA extension."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_a_to_n(script_num)			'Resets the array to add one more element to it
Set script_array_a_to_n(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_a_to_n(script_num).script_name 			= "Homelessness Verified"																	'Script name
script_array_a_to_n(script_num).file_name 				= "shelter-homelessness-verified.vbs"														'Script URL
script_array_a_to_n(script_num).description 			= "Case note template for documenting homelessness information."

'-------------------------------------------------------------------------------------------------------------------------N through Z
'Resetting the variable
'Set this array element to be a new script. Script details below...
script_num = 0							'Sets to zero
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "Mandatory Vendor App'd"																	'Script name
script_array_p_to_z(script_num).file_name 				= "shelter-mandatory-vendor-approved.vbs"													'Script URL
script_array_p_to_z(script_num).description 			= "Memo and case note template for approving mandatory vendor(s)."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "Mandatory Vendor Memo"																	'Script name
script_array_p_to_z(script_num).file_name 				= "shelter-mandatory-vendor-memo.vbs"														'Script URL
script_array_p_to_z(script_num).description 			= "Notice script that sends a Mandatory Vendor MEMO."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= " Money Mismanagement "																	'Script name
script_array_p_to_z(script_num).file_name 				= "shelter-money-mismanagement.vbs"															'Script URL
script_array_p_to_z(script_num).description 			= "Case note template for details for money mismanagement information."
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= " NSPOW Checked "																			'Script name
script_array_p_to_z(script_num).file_name 				= "shelter-nspow-checked.vbs"																'Script URL
script_array_p_to_z(script_num).description 			= "Case note template for details for NSPOW information."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "Partner Calls"																			'Script name
script_array_p_to_z(script_num).file_name				= "shelter-partner-calls.vbs"																'Script URL
script_array_p_to_z(script_num).description 			= "Case note template for documenting partner calls."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "Perm Housing Found"																		'Script name
script_array_p_to_z(script_num).file_name				= "shelter-permanent-housing-found.vbs"														'Script URL
script_array_p_to_z(script_num).description 			= "Case note template for documenting details of permanent housing found."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "Personal Needs"																			'Script name
script_array_p_to_z(script_num).file_name				= "shelter-personal-needs.vbs"																'Script URL
script_array_p_to_z(script_num).description 			= "Case note template for documenting personal needs information."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "P-Note"																					'Script name
script_array_p_to_z(script_num).file_name				= "shelter-p-note.vbs"																		'Script URL
script_array_p_to_z(script_num).description 			= "Template for adding person notes in MAXIS."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "Reim Shelter Account"																	'Script name
script_array_p_to_z(script_num).file_name				= "shelter-reimb-shelter-acct.vbs"															'Script URL
script_array_p_to_z(script_num).description 			= "Case note template for documenting details for reimbursement to the Shelter account."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "Revoucher"																				'Script name
script_array_p_to_z(script_num).file_name				= "shelter-revoucher.vbs"																	'Script URL
script_array_p_to_z(script_num).description 			= "Case note template for documenting details for the revoucher process."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "Self Pay"																				'Script name
script_array_p_to_z(script_num).file_name				= "shelter-selfpay.vbs"																		'Script URL
script_array_p_to_z(script_num).description 			= "Case note template for documenting details for shelter self pay ."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "Shelter Alternative"																		'Script name
script_array_p_to_z(script_num).file_name				= "shelter-shelter-alternative.vbs"															'Script URL
script_array_p_to_z(script_num).description 			= "Case note template for documenting information regarding shelter alternative."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "Shelter Interview"																		'Script name
script_array_p_to_z(script_num).file_name				= "shelter-shelter-interview.vbs"															'Script URL
script_array_p_to_z(script_num).description 			= "Case note template for documenting details of the shelter interview."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "Sheriff Foreclosure"																		'Script name
script_array_p_to_z(script_num).file_name				= "shelter-sheriff-foreclosure.vbs"															'Script URL
script_array_p_to_z(script_num).description 			= "Case note template for documenting sheriff foreclosure information."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "Special EA"																				'Script name
script_array_p_to_z(script_num).file_name				= "shelter-special-ea.vbs"																	'Script URL
script_array_p_to_z(script_num).description 			= "Case note template for documenting details for special EA."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "Utility Info"																			'Script name
script_array_p_to_z(script_num).file_name				= "shelter-utility-information.vbs"															'Script URL
script_array_p_to_z(script_num).description 			= "Case note template for documenting details of utility information."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_p_to_z(script_num)			'Resets the array to add one more element to it
Set script_array_p_to_z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_p_to_z(script_num).script_name 			= "Voucher Extended"																		'Script name
script_array_p_to_z(script_num).file_name				= "shelter-voucher-extended.vbs"															'Script URL
script_array_p_to_z(script_num).description 			= "Case note template for documenting details about the voucher extended process."


'Starting these with a very high number, higher than the normal possible amount of buttons.
'	We're doing this because we want to assign a value to each button pressed, and we want
'	that value to change with each button. The button_placeholder will be placed in the .button
'	property for each script item. This allows it to both escape the Function and resize
'	near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'	to force the value of ButtonPressed to hold in near infinite iterations.
button_placeholder 	= 24601
a_to_n_button		= 1000
p_to_z_button		= 2000

'Displays the dialog
Do
	If ButtonPressed = "" or ButtonPressed = a_to_n_button then
		declare_NOTES_menu_dialog(script_array_a_to_n)
	ElseIf ButtonPressed = p_to_z_button then
		declare_NOTES_menu_dialog(script_array_p_to_z)
	End if

	dialog NOTES_dialog
	If ButtonPressed = 0 then stopscript

    'Opening the SIR Instructions
	'IF buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/Notes%20scripts.aspx")
Loop until 	ButtonPressed <> a_to_n_button and _
			ButtonPressed <> p_to_z_button

'MsgBox buttonpressed = script_array_a_to_n(0).button

'Runs through each script in the array... if the selected script (buttonpressed) is in the array, it'll run_from_GitHub
For i = 0 to ubound(script_array_a_to_n)
	If ButtonPressed = script_array_a_to_n(i).button then call run_from_GitHub(script_repository & "shelter/" & script_array_a_to_n(i).file_name)
Next

For i = 0 to ubound(script_array_p_to_z)
	If ButtonPressed = script_array_p_to_z(i).button then call run_from_GitHub(script_repository & "shelter/" & script_array_p_to_z(i).file_name)
Next

stopscript
