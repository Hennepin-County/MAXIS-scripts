'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - MAIN MENU.vbs"
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
Function declare_NOTICES_menu_dialog(script_array)
	BeginDialog NOTICES_dialog, 0, 0, 516, 340, "NOTICES Scripts"
	 	Text 5, 5, 435, 10, "Notices scripts main menu: select the script to run from the choices below."
	  	ButtonGroup ButtonPressed
		 	PushButton 015, 35, 40, 15, "NOTICES", 				NOTICES_main_button
		 	PushButton 055, 35, 60, 15, "SNAP WCOMS", 				SNAP_WCOMS_button
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
		GroupBox 5, 20, 205, 35, "NOTICES Sub-Menus"
	EndDialog
End function
'END CUSTOM FUNCTIONS=======================================================================================================

'VARIABLES TO DECLARE=======================================================================================================

'Declaring the variable names to cut down on the number of arguments that need to be passed through the function.
DIM ButtonPressed
DIM SIR_instructions_button
dim NOTICES_dialog

script_array_NOTICES_main = array()
script_array_NOTICES_list = array()


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

'-------------------------------------------------------------------------------------------------------------------------NOTICES MAIN MENU

'Resetting the variable
script_num = 0
ReDim Preserve script_array_NOTICES_main(script_num)
Set script_array_NOTICES_main(script_num) = new script
script_array_NOTICES_main(script_num).script_name 			= "12 Month Contact"																		'Script name
script_array_NOTICES_main(script_num).file_name 			= "NOTICES - 12 MONTH CONTACT.vbs"															'Script URL
script_array_NOTICES_main(script_num).description 			= "Sends a MEMO to the client of their reporting responsibilities (required for SNAP 2-yr certifications, per POLI/TEMP TE02.08.165)."

script_num = script_num + 1									'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name 			= "Appointment Letter"																		'Script name
script_array_NOTICES_main(script_num).file_name 			= "NOTICES - APPOINTMENT LETTER.vbs"															'Script URL
script_array_NOTICES_main(script_num).description 			= "Sends a MEMO containing the appointment letter (with text from POLI/TEMP TE02.05.15)."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= " GRH Overpayment"													'needs spaces to generate button width properly.
script_array_NOTICES_main(script_num).file_name				= "NOTICES - GRH OP CL LEFT FACI.vbs"
script_array_NOTICES_main(script_num).description			= "Sends a MEMO to a facility indicating that an overpayment is due because a client left."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= "LTC Asset Transfer"
script_array_NOTICES_main(script_num).file_name				= "NOTICES - LTC - ASSET TRANSFER.vbs"
script_array_NOTICES_main(script_num).description			= "Sends a MEMO to a LTC client regarding asset transfers. "

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= "MA-EPD No Initial Premium Paid"
script_array_NOTICES_main(script_num).file_name				= "NOTICES - MA-EPD NO INITIAL PREMIUM.vbs"
script_array_NOTICES_main(script_num).description			= "Sends a WCOM on a denial for no initial MA-EPD premium."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= " Method B WCOM "													'needs spaces to generate button width properly.
script_array_NOTICES_main(script_num).file_name				= "NOTICES - METHOD B WCOM.vbs"
script_array_NOTICES_main(script_num).description			= "Makes detailed WCOM regarding spenddown vs. recipient amount for method B HC cases."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= "MFIP Orientation"													
script_array_NOTICES_main(script_num).file_name				= "NOTICES - MFIP ORIENTATION.vbs"
script_array_NOTICES_main(script_num).description			= "Sends a MEMO to a client regarding MFIP orientation."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= " MNsure Memo"													'needs spaces to generate button width properly.
script_array_NOTICES_main(script_num).file_name				= "NOTICES - MNSURE MEMO.vbs"
script_array_NOTICES_main(script_num).description			= "Sends a MEMO to a client regarding MNsure."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= "NOMI"
script_array_NOTICES_main(script_num).file_name				= "NOTICES - NOMI.vbs"
script_array_NOTICES_main(script_num).description			= "Sends the SNAP notice of missed interview (NOMI) letter, following rules set out in POLI/TEMP TE02.05.15."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= "Overdue Baby"
script_array_NOTICES_main(script_num).file_name				= "NOTICES - OVERDUE BABY.vbs"
script_array_NOTICES_main(script_num).description			= "Sends a MEMO informing client that they need to report information regarding the status of pregnancy, within 10 days or their case may close."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_NOTICES_main(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_main(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_main(script_num).script_name			= "SNAP E and T Letter"
script_array_NOTICES_main(script_num).file_name				= "NOTICES - SNAP E AND T LETTER.vbs"
script_array_NOTICES_main(script_num).description			= "Sends a SPEC/LETR informing client that they have an Employment and Training appointment."



'-------------------------------------------------------------------------------------------------------------------------SNAP WCOMS LISTS
'Resetting the variable
script_num = 0
ReDim Preserve script_array_NOTICES_list(script_num)
Set script_array_NOTICES_list(script_num) = new script
script_array_NOTICES_list(script_num).script_name 			= " ABAWD with Child in HH WCOM "'needs spaces to generate button width properly.																'Script name
script_array_NOTICES_list(script_num).file_name			= "NOTICES - ABAWD WITH CHILD IN HH WCOM.vbs"
script_array_NOTICES_list(script_num).description 			= "Adds a WCOM to a notice for an ABAWD adult receiving child under 18 exemption."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_NOTICES_list(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_list(script_num).script_name 			= "Duplicate assistance WCOM"
script_array_NOTICES_list(script_num).file_name			= "NOTICES - DUPLICATE ASSISTANCE WCOM.vbs"
script_array_NOTICES_list(script_num).description 			= "Adds a WCOM to a notice for duplicate assistance explaining why the client was ineligible."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_NOTICES_list(script_num)		'Resets the array to add one more element to it
Set script_array_NOTICES_list(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_NOTICES_list(script_num).script_name 			= "Postponed WREG Verif"
script_array_NOTICES_list(script_num).file_name			= "NOTICES - POSTPONED WREG VERIFS.vbs"
script_array_NOTICES_list(script_num).description 			= "Sends a WCOM informing the client of postponed verifications that MAXIS won't add to notice correctly by itself."




'Starting these with a very high number, higher than the normal possible amount of buttons.
'	We're doing this because we want to assign a value to each button pressed, and we want
'	that value to change with each button. The button_placeholder will be placed in the .button
'	property for each script item. This allows it to both escape the Function and resize
'	near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'	to force the value of ButtonPressed to hold in near infinite iterations.
button_placeholder 	= 24601
NOTICES_main_button		= 1000
SNAP_WCOMS_button		= 2000


'Displays the dialog
Do
	If ButtonPressed = "" or ButtonPressed = NOTICES_main_button then
		declare_NOTICES_menu_dialog(script_array_NOTICES_main)
	ElseIf ButtonPressed = SNAP_WCOMS_button then
		declare_NOTICES_menu_dialog(script_array_NOTICES_list)
	End if

	dialog NOTICES_dialog

	If ButtonPressed = 0 then stopscript
    'Opening the SIR Instructions
	IF buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/Notices%20scripts.aspx")
Loop until 	ButtonPressed <> SIR_instructions_button and _
			ButtonPressed <> NOTICES_main_button and _
			ButtonPressed <> SNAP_WCOMS_button 

'MsgBox buttonpressed = script_array_NOTICES_main(0).button

'Runs through each script in the array... if the selected script (buttonpressed) is in the array, it'll run_from_GitHub
For i = 0 to ubound(script_array_NOTICES_main)
	If ButtonPressed = script_array_NOTICES_main(i).button then call run_from_GitHub(script_repository & "/NOTICES/" & script_array_NOTICES_main(i).file_name)
Next

For i = 0 to ubound(script_array_NOTICES_list)
	If ButtonPressed = script_array_NOTICES_list(i).button then call run_from_GitHub(script_repository & "/NOTICES/" & script_array_NOTICES_list(i).file_name)
Next

stopscript
