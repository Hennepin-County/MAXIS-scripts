'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - Report a Duplicate UNEA Panel.vbs"
start_time = timer
STATS_counter = 0                     	'sets the stats counter at one
STATS_manualtime = 45                	'manual run time in seconds
STATS_denomination = "I"       		'C is for each CASE
'END OF stats block=========================================================================================================

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
call changelog_update("10/24/2023", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'THE SCRIPT ================================================================================================================
EMConnect ""											'connect to BZ
Call MAXIS_case_number_finder(MAXIS_case_number)		'See if we can find the case number
Call check_for_MAXIS(False)								'make sure we are in MAXIS

'Create a dialog image to gather the CASE NUMBER and explain the script purpose
BeginDialog Dialog1, 0, 0, 231, 145, "Case Number Dialog"
  EditBox 60, 55, 55, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 175, 105, 50, 15
    CancelButton 175, 125, 50, 15
  Text 70, 10, 95, 10, "EX PARTE SUPPORT ONLY"
  Text 10, 25, 210, 25, "This script is intended to capture information about a case and report to the Automation and Integration Team (the script writers) when it appears a duplicate UNEA panel has been created. "
  Text 10, 60, 50, 10, "Case Number:"
  Text 10, 80, 215, 25, "No additional information is needed. The script will gather details of all UNEA panels and send them to the script team. We will then be able to review the information on the case."
  Text 10, 120, 150, 20, "Once the script run is completed, you can update the case to process as needed."
EndDialog

'here is where we show the dialog
Do
	Do
		err_msg = ""									'reset the error message variable

		dialog Dialog1									'show the dialog
		cancel_without_confirmation						'cancel if the button is pressed

		Call validate_MAXIS_case_number(err_msg, "*")	'makse sure a Case Number has been entered

		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg		'display any error message that might exist

	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)		'make sure we are not passworded out
Loop until are_we_passworded_out = False
Call check_for_MAXIS(False)

'starting the email information
email_subject = "TESTING REPORT - " & MAXIS_case_number & "Ex Parte DUPLICATE UNEA Panel"		'setting the subject text

email_body = "The case - " & MAXIS_case_number & ", appears to have a duplicate UNEA panel created from the Ex Parte Report run." & vbCr		'starting the email body with some base information

'We are always going to be looking in current month plus 1 for Ex Parte cases
MAXIS_footer_monht = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

'Going to STAT/MEMB to get the member information
Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
If is_this_priv = True Then call script_end_procedure("The script has ended because this case appears to be privileged.")

member_list = " "									'creating a list that starts with a space. This is nice for splittling simple things
memb_row = 5										'begin at row 5
Do
	EMReadScreen ref_numb, 2, memb_row, 3			'read each reference number from the list on the left side of MAXIS.
	member_list = member_list & ref_numb & " "		'save to the string with a space between each reference number
	memb_row = memb_row + 1							'go to the next row
	EMReadScreen next_ref_numb, 2, memb_row, 3		'read to see if we are at the end of the list of reference numbers
Loop until next_ref_numb = "  "

'Create an array of the list of reference numbers, removing the spaces at the beginning and end, then split by the spaces
member_list = trim(member_list)
member_array = split(member_list)

'Now we go look at UNEA and save the information into the email body
Call navigate_to_MAXIS_screen("STAT", "UNEA")									'go to UNEA
email_body = email_body & "<BODY style=font-size:11pt;font-family:Courier New>"	'set the email body to courier font type

'Here we find all the UNEA panels
for each memb_numb in member_array												'we are going to check UENA for every HH member
	EMWriteScreen memb_numb, 20, 76												'enter the reference number in the command line
	Call write_value_and_transmit("01", 20, 79)									'enter 01 to get to the first instance and transmit to go to that panel

	EMReadScreen version_number, 1, 2, 78										'read the version number to make sure that any UNEA panel exists for this member
	If version_number <> "0" Then
		Do																		'This do loop is to go through all the panels for a specific worker
			STATS_counter = STATS_counter + 1									'incremember our stats count, which counts up for any panel found
			EMReadScreen instance, 1, 2, 73										'reading the instance for the panel to enter it into the email
			email_body = email_body & "<br>" & "<h3>PANEL INFORMATION: UNEA " & memb_numb & " 0" & instance & ":</h3>"		'This is a header line in the email of the panel information
			For unea_row = 2 to 19												'loop through each row of the panel to read and copy it into the email
				EMReadScreen unea_line, 57, unea_row, 24						'read the line from 24 to the end
				unea_line = replace(unea_line, "    ", "&emsp;&ensp;")			'format the spaces to look better in the email
				unea_line = replace(unea_line, "  ", "&emsp;")
				unea_line = replace(unea_line, " ", "&ensp;")
				email_body = email_body & "&emsp;&emsp;" & unea_line & "<br>"	'add the line information into the body of the email
			Next
			'this enters a horizontal line after the UNEA information in the email
			email_body = email_body & "<br>" & "------------------------------------------------------------------------------------" & "<br>"

			transmit															'try to go to the next UNEA panel for this member
			EMReadScreen enter_a_valid, 13, 24, 2								'read to see if MAXIS moved to the next UNEA or listed and error asking for 'Enter a valid command or PF-Key' indicating we are at the end of the UNEA panels
		Loop until enter_a_valid = "ENTER A VALID"
	End If
Next
email_body = email_body & "</BODY>"												'closing the html body in the email and ending use of Courier New font

call find_user_name(worker_name)				'getting the name of the user for the email
email_body = email_body & "<br>" & "Thank you!" & "<br>" & worker_name & "<br>" & now		'adding a signature to the email

'Sending the email. The parameters can be seen here:
'create_outlook_email(email_from, email_recip, email_recip_CC, email_recip_bcc, email_subject, email_importance, include_flag, email_flag_text, email_flag_days, email_flag_reminder, email_flag_reminder_days, email_body, include_email_attachment, email_attachment_array, send_email)
Call create_outlook_email("", "hsph.ews.bluezonescripts@hennepin.us", "", "", email_subject, 1, False, "", "", "", "", email_body, False, "", True)

'Creating an end message to let the worker know the script ran and the case can be processed
end_msg = "EMAIL SENT TO THE A & I Team"
end_msg = end_msg & vbCr & "No additional report is needed."
end_msg = end_msg & vbCr & "You can now update the case and process as needed."
Call script_end_procedure(end_msg)														'ending the script run

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------10/24/2023
'--Tab orders reviewed & confirmed----------------------------------------------10/24/2023
'--Mandatory fields all present & Reviewed--------------------------------------10/24/2023
'--All variables in dialog match mandatory fields-------------------------------10/24/2023
'Review dialog names for content and content fit in dialog----------------------10/24/2023
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------N/A
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------N/A
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------10/24/2023
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------10/24/2023
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------10/24/2023
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------10/24/2023
'--Incrementors reviewed (if necessary)-----------------------------------------10/24/2023
'--Denomination reviewed -------------------------------------------------------10/24/2023
'--Script name reviewed---------------------------------------------------------10/24/2023
'--BULK - remove 1 incrementor at end of script reviewed------------------------10/24/2023

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------10/24/2023
'--comment Code-----------------------------------------------------------------10/24/2023
'--Update Changelog for release/update------------------------------------------10/24/2023
'--Remove testing message boxes-------------------------------------------------10/24/2023
'--Remove testing code/unnecessary code-----------------------------------------10/24/2023
'--Review/update SharePoint instructions----------------------------------------10/24/2023
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------10/25/2023
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------N/A
'--Complete misc. documentation (if applicable)---------------------------------10/24/2023
'--Update project team/issue contact (if applicable)----------------------------10/24/2023
