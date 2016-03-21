'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MAIN MENU.vbs"
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
Function declare_NOTES_menu_dialog(script_array)
	BeginDialog NOTES_dialog, 0, 0, 516, 340, "NOTES Scripts"
	 	Text 5, 5, 435, 10, "Notes scripts main menu: select the script to run from the choices below. Notes with autofill functionality marked with an asterisk (*)."
	  	ButtonGroup ButtonPressed
		 	PushButton 015, 35, 30, 15, "# - C", 				a_to_c_button
		 	PushButton 045, 35, 30, 15, "D - F", 				d_to_f_button
		 	PushButton 075, 35, 30, 15, "G - L", 				g_to_l_button
		 	PushButton 105, 35, 30, 15, "M - Q", 				m_to_q_button
		 	PushButton 135, 35, 30, 15, "R - Z", 				r_to_z_button
		 	PushButton 165, 35, 30, 15, "LTC", 					ltc_button
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
		GroupBox 5, 20, 205, 35, "NOTES Sub-Menus"
	EndDialog
End function
'END CUSTOM FUNCTIONS=======================================================================================================

'VARIABLES TO DECLARE=======================================================================================================

'Declaring the variable names to cut down on the number of arguments that need to be passed through the function.
DIM ButtonPressed
DIM SIR_instructions_button
dim NOTES_dialog

script_array_0_to_C = array()
script_array_D_to_F = array()
script_array_G_to_L = array()
script_array_M_to_Q = array()
script_array_R_to_Z = array()
script_array_LTC    = array()

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

'-------------------------------------------------------------------------------------------------------------------------0 through C

'Resetting the variable
script_num = 0
ReDim Preserve script_array_0_to_C(script_num)
Set script_array_0_to_C(script_num) = new script
script_array_0_to_C(script_num).script_name 			= "Application Received"																		'Script name
script_array_0_to_C(script_num).file_name 				= "NOTES - APPLICATION RECEIVED.vbs"															'Script URL
script_array_0_to_C(script_num).description 			= "Template for documenting details about an application recevied."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_0_to_C(script_num)			'Resets the array to add one more element to it
Set script_array_0_to_C(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_0_to_C(script_num).script_name 			= "Approved programs"																		'Script name
script_array_0_to_C(script_num).file_name 				= "NOTES - APPROVED PROGRAMS.vbs"															'Script URL
script_array_0_to_C(script_num).description 			= "Template for when you approve a client's programs."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_0_to_C(script_num)			'Resets the array to add one more element to it
Set script_array_0_to_C(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_0_to_C(script_num).script_name				= "AREP Form Received"
script_array_0_to_C(script_num).file_name				= "NOTES - AREP FORM RECEIVED.vbs"
script_array_0_to_C(script_num).description				= "NEW 02/2016!!! Template for when you receive an Authorized Representative (AREP) form."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_0_to_C(script_num)			'Resets the array to add one more element to it
Set script_array_0_to_C(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_0_to_C(script_num).script_name				= "Burial assets"
script_array_0_to_C(script_num).file_name				= "NOTES - BURIAL ASSETS.vbs"
script_array_0_to_C(script_num).description				= "Template for burial assets."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_0_to_C(script_num)			'Resets the array to add one more element to it
Set script_array_0_to_C(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_0_to_C(script_num).script_name				= "CAF"
script_array_0_to_C(script_num).file_name				= "NOTES - CAF.vbs"
script_array_0_to_C(script_num).description				= "Template for when you're processing a CAF. Works for intake as well as recertification and reapplication.*"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_0_to_C(script_num)			'Resets the array to add one more element to it
Set script_array_0_to_C(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_0_to_C(script_num).script_name				= "Change Report Form Received"
script_array_0_to_C(script_num).file_name				= "NOTES - CHANGE REPORT FORM RECEIVED.vbs"
script_array_0_to_C(script_num).description				= "Template for case noting information reported from a Change Report Form."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_0_to_C(script_num)			'Resets the array to add one more element to it
Set script_array_0_to_C(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_0_to_C(script_num).script_name				= "Change Reported"
script_array_0_to_C(script_num).file_name				= "NOTES - CHANGE REPORTED.vbs"
script_array_0_to_C(script_num).description				= "Template for case noting HHLD Comp or Baby Born being reported. **More changes to be added in the future**"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_0_to_C(script_num)			'Resets the array to add one more element to it
Set script_array_0_to_C(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_0_to_C(script_num).script_name				= "Citizenship/identity verified"
script_array_0_to_C(script_num).file_name				= "NOTES - CITIZENSHIP-IDENTITY VERIFIED.vbs"
script_array_0_to_C(script_num).description				= "Template for documenting citizenship/identity status for a case."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_0_to_C(script_num)			'Resets the array to add one more element to it
Set script_array_0_to_C(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_0_to_C(script_num).script_name				= "Client contact"
script_array_0_to_C(script_num).file_name				= "NOTES - CLIENT CONTACT.vbs"
script_array_0_to_C(script_num).description				= "Template for documenting client contact, either from or to a client."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_0_to_C(script_num)			'Resets the array to add one more element to it
Set script_array_0_to_C(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_0_to_C(script_num).script_name				= "Client Transportation Costs"
script_array_0_to_C(script_num).file_name				= "NOTES - CLIENT TRANSPORTATION COSTS.vbs"
script_array_0_to_C(script_num).description				= "NEW 02/2016!!! Template for documenting client transportation costs."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_0_to_C(script_num)			'Resets the array to add one more element to it
Set script_array_0_to_C(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_0_to_C(script_num).script_name				= "Closed programs"
script_array_0_to_C(script_num).file_name				= "NOTES - CLOSED PROGRAMS.vbs"
script_array_0_to_C(script_num).description				= "Template for indicating which programs are closing, and when. Also case notes intake/REIN dates based on various selections."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_0_to_C(script_num)			'Resets the array to add one more element to it
Set script_array_0_to_C(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_0_to_C(script_num).script_name				= "Combined AR"
script_array_0_to_C(script_num).file_name				= "NOTES - COMBINED AR.vbs"
script_array_0_to_C(script_num).description				= "Template for the Combined Annual Renewal.*"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_0_to_C(script_num)			'Resets the array to add one more element to it
Set script_array_0_to_C(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_0_to_C(script_num).script_name				= "County Burial Application"
script_array_0_to_C(script_num).file_name				= "NOTES - COUNTY BURIAL APPLICATION.vbs"
script_array_0_to_C(script_num).description				= "Template for the County Burial Application.*"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_0_to_C(script_num)			'Resets the array to add one more element to it
Set script_array_0_to_C(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_0_to_C(script_num).script_name				= "CSR"
script_array_0_to_C(script_num).file_name				= "NOTES - CSR.vbs"
script_array_0_to_C(script_num).description				= "Template for the Combined Six-month Report (CSR).*"




'-------------------------------------------------------------------------------------------------------------------------D through F
'Resetting the variable
script_num = 0
ReDim Preserve script_array_D_to_F(script_num)
Set script_array_D_to_F(script_num) = new script
script_array_D_to_F(script_num).script_name 			= "Denied programs"																		'Script name
script_array_D_to_F(script_num).file_name				= "NOTES - DENIED PROGRAMS.vbs"
script_array_D_to_F(script_num).description 			= "Template for indicating which programs you've denied, and when. Also case notes intake/REIN dates based on various selections."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_D_to_F(script_num)			'Resets the array to add one more element to it
Set script_array_D_to_F(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_D_to_F(script_num).script_name 			= "Docs Received"
script_array_D_to_F(script_num).file_name				= "NOTES - DOCUMENTS RECEIVED.vbs"
script_array_D_to_F(script_num).description 			= "Template for case noting information about documents received."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_D_to_F(script_num)			'Resets the array to add one more element to it
Set script_array_D_to_F(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_D_to_F(script_num).script_name 			= "Drug felon"
script_array_D_to_F(script_num).file_name				= "NOTES - DRUG FELON.vbs"
script_array_D_to_F(script_num).description 			= "Template for noting drug felon info."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_D_to_F(script_num)			'Resets the array to add one more element to it
Set script_array_D_to_F(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_D_to_F(script_num).script_name 			= "DWP budget"
script_array_D_to_F(script_num).file_name				= "NOTES - DWP BUDGET.vbs"
script_array_D_to_F(script_num).description 			= "Template for noting DWP budgets."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_D_to_F(script_num)			'Resets the array to add one more element to it
Set script_array_D_to_F(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_D_to_F(script_num).script_name 			= "EDRS DISQ match found"
script_array_D_to_F(script_num).file_name				= "NOTES - EDRS DISQ MATCH FOUND.vbs"
script_array_D_to_F(script_num).description 			= "NEW 02/2016!!! Template for noting the action steps when a SNAP recipient has an eDRS DISQ per TE02.08.127."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_D_to_F(script_num)			'Resets the array to add one more element to it
Set script_array_D_to_F(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_D_to_F(script_num).script_name 			= "Emergency"
script_array_D_to_F(script_num).file_name				= "NOTES - EMERGENCY.vbs"
script_array_D_to_F(script_num).description 			= "Template for EA/EGA applications.*"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_D_to_F(script_num)			'Resets the array to add one more element to it
Set script_array_D_to_F(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_D_to_F(script_num).script_name 			= "Employment plan or status update"
script_array_D_to_F(script_num).file_name				= "NOTES - EMPLOYMENT PLAN OR STATUS UPDATE.vbs"
script_array_D_to_F(script_num).description 			= "Template for case noting an employment plan or status update for family cash cases."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_D_to_F(script_num)			'Resets the array to add one more element to it
Set script_array_D_to_F(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_D_to_F(script_num).script_name 			= "Employment Verif Recv'd"
script_array_D_to_F(script_num).file_name				= "NOTES - EVF RECEIVED.vbs"
script_array_D_to_F(script_num).description 			= "Template for noting information about an employment verification received by the agency."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_D_to_F(script_num)			'Resets the array to add one more element to it
Set script_array_D_to_F(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_D_to_F(script_num).script_name 			= "ES Referral"
script_array_D_to_F(script_num).file_name				= "NOTES - ES REFERRAL.vbs"
script_array_D_to_F(script_num).description 			= "Template for sending an MFIP or DWP referral to employment services."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_D_to_F(script_num)			'Resets the array to add one more element to it
Set script_array_D_to_F(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_D_to_F(script_num).script_name 			= "Expedited determination"
script_array_D_to_F(script_num).file_name				= "NOTES - EXPEDITED DETERMINATION.vbs"
script_array_D_to_F(script_num).description 			= "Template for noting detail about how expedited was determined for a case."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_D_to_F(script_num)			'Resets the array to add one more element to it
Set script_array_D_to_F(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_D_to_F(script_num).script_name 			= "Expedited screening"
script_array_D_to_F(script_num).file_name				= "NOTES - EXPEDITED SCREENING.vbs"
script_array_D_to_F(script_num).description 			= "Template for screening a client for expedited status."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_D_to_F(script_num)			'Resets the array to add one more element to it
Set script_array_D_to_F(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_D_to_F(script_num).script_name 			= "Explanation of Income Budgeted"
script_array_D_to_F(script_num).file_name				= "NOTES - EXPLANATION OF INCOME BUDGETED.vbs"
script_array_D_to_F(script_num).description 			= "NEW 01/2016!!! Template for explaining the income budgeted for a case."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_D_to_F(script_num)			'Resets the array to add one more element to it
Set script_array_D_to_F(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_D_to_F(script_num).script_name 			= "Foster Care HCAPP"
script_array_D_to_F(script_num).file_name				= "NOTES - FOSTER CARE HCAPP.vbs"
script_array_D_to_F(script_num).description 			= "NEW 02/2016!!! Template for noting foster care HCAPP info."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_D_to_F(script_num)			'Resets the array to add one more element to it
Set script_array_D_to_F(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_D_to_F(script_num).script_name 			= "Foster Care Review"
script_array_D_to_F(script_num).file_name				= "NOTES - FOSTER CARE REVIEW.vbs"
script_array_D_to_F(script_num).description 			= "NEW 02/2016!!! Template for noting foster care review info."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_D_to_F(script_num)			'Resets the array to add one more element to it
Set script_array_D_to_F(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_D_to_F(script_num).script_name 			= "Fraud info"
script_array_D_to_F(script_num).file_name				= "NOTES - FRAUD INFO.vbs"
script_array_D_to_F(script_num).description 			= "Template for noting fraud info."




'-------------------------------------------------------------------------------------------------------------------------G through L
'Resetting the variable
script_num = 0
ReDim Preserve script_array_G_to_L(script_num)
Set script_array_G_to_L(script_num) = new script
script_array_G_to_L(script_num).script_name 			= "Good Cause Claimed"
script_array_G_to_L(script_num).file_name				= "NOTES - GOOD CAUSE CLAIMED.vbs"
script_array_G_to_L(script_num).description				= "Template for requests of good cause to not receive child support."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_G_to_L(script_num)			'Resets the array to add one more element to it
Set script_array_G_to_L(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_G_to_L(script_num).script_name 			= "Good Cause Results"
script_array_G_to_L(script_num).file_name				= "NOTES - GOOD CAUSE RESULTS.vbs"
script_array_G_to_L(script_num).description				= "Template for Good Cause results for determination or renewal.*"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_G_to_L(script_num)			'Resets the array to add one more element to it
Set script_array_G_to_L(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_G_to_L(script_num).script_name 			= "GRH - HRF"
script_array_G_to_L(script_num).file_name				= "NOTES - GRH - HRF.vbs"
script_array_G_to_L(script_num).description				= "Template for GRH HRFs. Case must be post-pay.*"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_G_to_L(script_num)			'Resets the array to add one more element to it
Set script_array_G_to_L(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_G_to_L(script_num).script_name 			= "HC ICAMA"
script_array_G_to_L(script_num).file_name				= "NOTES - HC ICAMA.vbs"
script_array_G_to_L(script_num).description				= "NEW 02/2016!!! Template for HC Interstate Compact on Adoption and Medical Assistance (HC ICAMA)."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_G_to_L(script_num)			'Resets the array to add one more element to it
Set script_array_G_to_L(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_G_to_L(script_num).script_name 			= "HC Renewal"
script_array_G_to_L(script_num).file_name				= "NOTES - HC RENEWAL.vbs"
script_array_G_to_L(script_num).description				= "Template for HC renewals.*"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_G_to_L(script_num)			'Resets the array to add one more element to it
Set script_array_G_to_L(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_G_to_L(script_num).script_name 			= "HCAPP"
script_array_G_to_L(script_num).file_name				= "NOTES - HCAPP.vbs"
script_array_G_to_L(script_num).description				= "Template for HCAPPs.*"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_G_to_L(script_num)			'Resets the array to add one more element to it
Set script_array_G_to_L(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_G_to_L(script_num).script_name 			= "HRF"
script_array_G_to_L(script_num).file_name				= "NOTES - HRF.vbs"
script_array_G_to_L(script_num).description				= "Template for HRFs (for GRH, use the ''GRH - HRF'' script).*"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_G_to_L(script_num)			'Resets the array to add one more element to it
Set script_array_G_to_L(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_G_to_L(script_num).script_name 			= "IEVS Notice Received"
script_array_G_to_L(script_num).file_name				= "NOTES - IEVS NOTICE RECEIVED.vbs"
script_array_G_to_L(script_num).description				= "Template to case note when a IEVS notice is returned."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_G_to_L(script_num)			'Resets the array to add one more element to it
Set script_array_G_to_L(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_G_to_L(script_num).script_name 			= "Incarceration "
script_array_G_to_L(script_num).file_name				= "NOTES - INCARCERATION.vbs"
script_array_G_to_L(script_num).description				= "Template to note details of an incarceration, and also updates STAT/FACI if necessary."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_G_to_L(script_num)			'Resets the array to add one more element to it
Set script_array_G_to_L(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_G_to_L(script_num).script_name 			= "Interview Completed"
script_array_G_to_L(script_num).file_name				= "NOTES - INTERVIEW COMPLETED.vbs"
script_array_G_to_L(script_num).description				= "Template to case note an interview being completed but no stat panels updated."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_G_to_L(script_num)			'Resets the array to add one more element to it
Set script_array_G_to_L(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_G_to_L(script_num).script_name 			= "Interview No Show"
script_array_G_to_L(script_num).file_name				= "NOTES - INTERVIEW NO SHOW.vbs"
script_array_G_to_L(script_num).description				= "NEW 01/2016!!! Template for case noting a client's no-showing their in-office or phone appointment."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_G_to_L(script_num)			'Resets the array to add one more element to it
Set script_array_G_to_L(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_G_to_L(script_num).script_name 			= "LEP - EMA"
script_array_G_to_L(script_num).file_name				= "NOTES - LEP - EMA.vbs"
script_array_G_to_L(script_num).description				= "Template for EMA applications."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_G_to_L(script_num)			'Resets the array to add one more element to it
Set script_array_G_to_L(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_G_to_L(script_num).script_name 			= "LEP - SAVE"
script_array_G_to_L(script_num).file_name				= "NOTES - LEP - SAVE.vbs"
script_array_G_to_L(script_num).description				= "Template for the SAVE system for verifying immigration status."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_G_to_L(script_num)			'Resets the array to add one more element to it
Set script_array_G_to_L(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_G_to_L(script_num).script_name 			= "LEP - Sponsor income"
script_array_G_to_L(script_num).file_name				= "NOTES - LEP - SPONSOR INCOME.vbs"
script_array_G_to_L(script_num).description				= "Template for the sponsor income deeming calculation (it will also help calculate it for you)."




'-------------------------------------------------------------------------------------------------------------------------M through Q
'Resetting the variable
script_num = 0
ReDim Preserve script_array_M_to_Q(script_num)
Set script_array_M_to_Q(script_num) = new script
script_array_M_to_Q(script_num).script_name 			= "Medical Opinion Form Received"
script_array_M_to_Q(script_num).file_name				= "NOTES - MEDICAL OPINION FORM RECEIVED.vbs"
script_array_M_to_Q(script_num).description				= "Template for case noting information about a Medical Opinion Form."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_M_to_Q(script_num)			'Resets the array to add one more element to it
Set script_array_M_to_Q(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_M_to_Q(script_num).script_name 			= "MFIP Sanction Cured"
script_array_M_to_Q(script_num).file_name				= "NOTES - MFIP SANCTION CURED.vbs"
script_array_M_to_Q(script_num).description				= "Template for noting the curing of an MFIP sanction."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_M_to_Q(script_num)			'Resets the array to add one more element to it
Set script_array_M_to_Q(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_M_to_Q(script_num).script_name 			= "MFIP Sanction/DWP Disqualification"
script_array_M_to_Q(script_num).file_name				= "NOTES - MFIP SANCTION AND DWP DISQUALIFICATION.vbs"
script_array_M_to_Q(script_num).description				= "Template for MFIP sanctions and DWP disqualifications, both CS and ES."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_M_to_Q(script_num)			'Resets the array to add one more element to it
Set script_array_M_to_Q(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_M_to_Q(script_num).script_name 			= "MFIP to SNAP Transition"
script_array_M_to_Q(script_num).file_name				= "NOTES - MFIP TO SNAP TRANSITION.vbs"
script_array_M_to_Q(script_num).description				= "Template for noting when closing MFIP and opening SNAP."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_M_to_Q(script_num)			'Resets the array to add one more element to it
Set script_array_M_to_Q(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_M_to_Q(script_num).script_name 			= "MNsure - Documents requested"
script_array_M_to_Q(script_num).file_name				= "NOTES - MNSURE - DOCUMENTS REQUESTED.vbs"
script_array_M_to_Q(script_num).description				= "Template for when MNsure documents have been requested."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_M_to_Q(script_num)			'Resets the array to add one more element to it
Set script_array_M_to_Q(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_M_to_Q(script_num).script_name 			= "MNsure - Retro HC Application"
script_array_M_to_Q(script_num).file_name				= "NOTES - MNSURE RETRO HC APPLICATION.vbs"
script_array_M_to_Q(script_num).description				= "Template for when MNsure retro HC has been requested."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_M_to_Q(script_num)			'Resets the array to add one more element to it
Set script_array_M_to_Q(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_M_to_Q(script_num).script_name 			= "MSQ"
script_array_M_to_Q(script_num).file_name				= "NOTES - MSQ.vbs"
script_array_M_to_Q(script_num).description				= "NEW 02/2016!!! Template for noting Medical Service Questionaires (MSQ)."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_M_to_Q(script_num)			'Resets the array to add one more element to it
Set script_array_M_to_Q(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_M_to_Q(script_num).script_name 			= "MTAF"
script_array_M_to_Q(script_num).file_name				= "NOTES - MTAF.vbs"
script_array_M_to_Q(script_num).description				= "Template for the MN Transition Application form (MTAF)."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_M_to_Q(script_num)			'Resets the array to add one more element to it
Set script_array_M_to_Q(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_M_to_Q(script_num).script_name 			= "OHP Received"
script_array_M_to_Q(script_num).file_name				= "NOTES - OHP RECEIVED.vbs"
script_array_M_to_Q(script_num).description				= "NEW 02/2016!!! Template for noting Out of Home Placement (OHP)."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_M_to_Q(script_num)			'Resets the array to add one more element to it
Set script_array_M_to_Q(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_M_to_Q(script_num).script_name 			= "Overpayment"
script_array_M_to_Q(script_num).file_name				= "NOTES - OVERPAYMENT.vbs"
script_array_M_to_Q(script_num).description				= "Template for noting basic information about overpayments."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_M_to_Q(script_num)			'Resets the array to add one more element to it
Set script_array_M_to_Q(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_M_to_Q(script_num).script_name 			= "Pregnancy Reported"
script_array_M_to_Q(script_num).file_name				= "NOTES - PREGNANCY REPORTED.vbs"
script_array_M_to_Q(script_num).description				= "Template for case noting a pregnancy. This script can update STAT/PREG."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_M_to_Q(script_num)			'Resets the array to add one more element to it
Set script_array_M_to_Q(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_M_to_Q(script_num).script_name 			= "Proof of relationship"
script_array_M_to_Q(script_num).file_name				= "NOTES - PROOF OF RELATIONSHIP.vbs"
script_array_M_to_Q(script_num).description				= "Template for documenting proof of relationship between a member 01 and someone else in the household."




'-------------------------------------------------------------------------------------------------------------------------R through Z
'Resetting the variable
script_num = 0
ReDim Preserve script_array_R_to_Z(script_num)
Set script_array_R_to_Z(script_num) = new script
script_array_R_to_Z(script_num).script_name 			= "REIN Progs"
script_array_R_to_Z(script_num).file_name				= "NOTES - REIN PROGS.vbs"
script_array_R_to_Z(script_num).description				= "Template for noting program reinstatement information."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_R_to_Z(script_num)			'Resets the array to add one more element to it
Set script_array_R_to_Z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_R_to_Z(script_num).script_name 			= "Returned Mail"
script_array_R_to_Z(script_num).file_name				= "NOTES - RETURNED MAIL RECEIVED.vbs"
script_array_R_to_Z(script_num).description				= "Template for noting Returned Mail Received information."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_R_to_Z(script_num)			'Resets the array to add one more element to it
Set script_array_R_to_Z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_R_to_Z(script_num).script_name 			= "Significant Change"
script_array_R_to_Z(script_num).file_name				= "NOTES - SIGNIFICANT CHANGE.vbs"
script_array_R_to_Z(script_num).description				= "Template for noting Significant Change information."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_R_to_Z(script_num)			'Resets the array to add one more element to it
Set script_array_R_to_Z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_R_to_Z(script_num).script_name 			= "SNAP case review"
script_array_R_to_Z(script_num).file_name				= "NOTES - SNAP CASE REVIEW.vbs"
script_array_R_to_Z(script_num).description				= "Template for SNAP reviewers to use that will case note the status SNAP quality review."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_R_to_Z(script_num)			'Resets the array to add one more element to it
Set script_array_R_to_Z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_R_to_Z(script_num).script_name 			= "Submit case for SNAP review"
script_array_R_to_Z(script_num).file_name				= "NOTES - SUBMIT CASE FOR SNAP REVIEW.vbs"
script_array_R_to_Z(script_num).description				= "Template for when a worker sends a case to be reviewed by a supervisor or coworker."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_R_to_Z(script_num)			'Resets the array to add one more element to it
Set script_array_R_to_Z(script_num) = new script		'Set this array element to be a new script. Script details below...
script_array_R_to_Z(script_num).script_name 			= "Verifications needed"
script_array_R_to_Z(script_num).file_name				= "NOTES - VERIFICATIONS NEEDED.vbs"
script_array_R_to_Z(script_num).description				= "Template for when verifications are needed (enters each verification clearly)."




'-------------------------------------------------------------------------------------------------------------------------LTC
'Resetting the variable
script_num = 0
ReDim Preserve script_array_LTC(script_num)
Set script_array_LTC(script_num) = new script
script_array_LTC(script_num).script_name 				= "LTC - 1503"
script_array_LTC(script_num).file_name					= "NOTES - LTC - 1503.vbs"
script_array_LTC(script_num).description				= "Template for processing DHS-1503."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_LTC(script_num)				'Resets the array to add one more element to it
Set script_array_LTC(script_num) = new script			'Set this array element to be a new script. Script details below...
script_array_LTC(script_num).script_name 				= "LTC - 5181"
script_array_LTC(script_num).file_name					= "NOTES - LTC - 5181.vbs"
script_array_LTC(script_num).description				= "Template for processing DHS-5181."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_LTC(script_num)				'Resets the array to add one more element to it
Set script_array_LTC(script_num) = new script			'Set this array element to be a new script. Script details below...
script_array_LTC(script_num).script_name 				= "LTC - Application received"
script_array_LTC(script_num).file_name					= "NOTES - LTC - APPLICATION RECEIVED.vbs"
script_array_LTC(script_num).description				= "Template for initial details of a LTC application.*"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_LTC(script_num)				'Resets the array to add one more element to it
Set script_array_LTC(script_num) = new script			'Set this array element to be a new script. Script details below...
script_array_LTC(script_num).script_name 				= "LTC - Asset assessment"
script_array_LTC(script_num).file_name					= "NOTES - LTC - ASSET ASSESSMENT.vbs"
script_array_LTC(script_num).description				= "Template for the LTC asset assessment. Will enter both person and case notes if desired."

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_LTC(script_num)				'Resets the array to add one more element to it
Set script_array_LTC(script_num) = new script			'Set this array element to be a new script. Script details below...
script_array_LTC(script_num).script_name 				= "LTC - COLA summary 2016"
script_array_LTC(script_num).file_name					= "NOTES - LTC - COLA SUMMARY 2016.vbs"
script_array_LTC(script_num).description				= "Template to summarize actions for the 2016 COLA.*"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_LTC(script_num)				'Resets the array to add one more element to it
Set script_array_LTC(script_num) = new script			'Set this array element to be a new script. Script details below...
script_array_LTC(script_num).script_name 				= "LTC - Intake approval"
script_array_LTC(script_num).file_name					= "NOTES - LTC - INTAKE APPROVAL.vbs"
script_array_LTC(script_num).description				= "Template for use when approving a LTC intake.*"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_LTC(script_num)				'Resets the array to add one more element to it
Set script_array_LTC(script_num) = new script			'Set this array element to be a new script. Script details below...
script_array_LTC(script_num).script_name 				= "LTC - MA approval"
script_array_LTC(script_num).file_name					= "NOTES - LTC - MA APPROVAL.vbs"
script_array_LTC(script_num).description				= "Template for approving LTC MA (can be used for changes, initial application, or recertification).*"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_LTC(script_num)				'Resets the array to add one more element to it
Set script_array_LTC(script_num) = new script			'Set this array element to be a new script. Script details below...
script_array_LTC(script_num).script_name 				= "LTC - Renewal"
script_array_LTC(script_num).file_name					= "NOTES - LTC - RENEWAL.vbs"
script_array_LTC(script_num).description				= "Template for LTC renewals.*"

script_num = script_num + 1								'Increment by one
ReDim Preserve script_array_LTC(script_num)				'Resets the array to add one more element to it
Set script_array_LTC(script_num) = new script			'Set this array element to be a new script. Script details below...
script_array_LTC(script_num).script_name 				= "LTC - Transfer penalty"
script_array_LTC(script_num).file_name					= "NOTES - LTC - TRANSFER PENALTY.vbs"
script_array_LTC(script_num).description				= "Template for noting a transfer penalty."

'Starting these with a very high number, higher than the normal possible amount of buttons.
'	We're doing this because we want to assign a value to each button pressed, and we want
'	that value to change with each button. The button_placeholder will be placed in the .button
'	property for each script item. This allows it to both escape the Function and resize
'	near infinitely. We use dummy numbers for the other selector buttons for much the same reason,
'	to force the value of ButtonPressed to hold in near infinite iterations.
button_placeholder 	= 24601
a_to_c_button		= 1000
d_to_f_button		= 2000
g_to_l_button		= 3000
m_to_q_button		= 4000
r_to_z_button		= 5000
ltc_button			= 6000

'Displays the dialog
Do
	If ButtonPressed = "" or ButtonPressed = a_to_c_button then
		declare_NOTES_menu_dialog(script_array_0_to_C)
	ElseIf ButtonPressed = d_to_f_button then
		declare_NOTES_menu_dialog(script_array_D_to_F)
	ElseIf ButtonPressed = g_to_l_button then
		declare_NOTES_menu_dialog(script_array_G_to_L)
	ElseIf ButtonPressed = m_to_q_button then
		declare_NOTES_menu_dialog(script_array_M_to_Q)
	ElseIf ButtonPressed = r_to_z_button then
		declare_NOTES_menu_dialog(script_array_R_to_Z)
	ElseIf ButtonPressed = ltc_button then
		declare_NOTES_menu_dialog(script_array_LTC)
	End if

	dialog NOTES_dialog

	If ButtonPressed = 0 then stopscript
    'Opening the SIR Instructions
	IF buttonpressed = SIR_instructions_button then CreateObject("WScript.Shell").Run("https://www.dhssir.cty.dhs.state.mn.us/MAXIS/blzn/Script%20Instructions%20Wiki/Notes%20scripts.aspx")
Loop until 	ButtonPressed <> SIR_instructions_button and _
			ButtonPressed <> a_to_c_button and _
			ButtonPressed <> d_to_f_button and _
			ButtonPressed <> g_to_l_button and _
			ButtonPressed <> m_to_q_button and _
			ButtonPressed <> r_to_z_button and _
			ButtonPressed <> ltc_button

'MsgBox buttonpressed = script_array_0_to_C(0).button

'Runs through each script in the array... if the selected script (buttonpressed) is in the array, it'll run_from_GitHub
For i = 0 to ubound(script_array_0_to_C)
	If ButtonPressed = script_array_0_to_C(i).button then call run_from_GitHub(script_repository & "/NOTES/" & script_array_0_to_C(i).file_name)
Next

For i = 0 to ubound(script_array_D_to_F)
	If ButtonPressed = script_array_D_to_F(i).button then call run_from_GitHub(script_repository & "/NOTES/" & script_array_D_to_F(i).file_name)
Next

For i = 0 to ubound(script_array_G_to_L)
	If ButtonPressed = script_array_G_to_L(i).button then call run_from_GitHub(script_repository & "/NOTES/" & script_array_G_to_L(i).file_name)
Next

For i = 0 to ubound(script_array_M_to_Q)
	If ButtonPressed = script_array_M_to_Q(i).button then call run_from_GitHub(script_repository & "/NOTES/" & script_array_M_to_Q(i).file_name)
Next

For i = 0 to ubound(script_array_R_to_Z)
	If ButtonPressed = script_array_R_to_Z(i).button then call run_from_GitHub(script_repository & "/NOTES/" & script_array_R_to_Z(i).file_name)
Next

For i = 0 to ubound(script_array_LTC)
	If ButtonPressed = script_array_LTC(i).button 	 then call run_from_GitHub(script_repository & "/NOTES/" & script_array_LTC(i).file_name)
Next
stopscript
