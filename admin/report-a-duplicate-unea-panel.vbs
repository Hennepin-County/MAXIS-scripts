'Required for statistical purposes==========================================================================================
name_of_script = "ADMIN - Report a Duplicate UNEA Panel.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 0                	'manual run time in seconds
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

EMConnect ""
Call MAXIS_case_number_finder(MAXIS_case_number)
Call check_for_MAXIS(False)

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

Do
	Do
		err_msg = ""

		dialog Dialog1
		cancel_without_confirmation

		Call validate_MAXIS_case_number(err_msg, "*")

		If err_msg <> "" Then MsgBox "*** Please Resolve to Continue ***" & vbNewLine & err_msg

	Loop until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until are_we_passworded_out = False

email_subject = "TESTING REPORT - Ex Parte DUPLICATE UNEA Panel"

email_body = "The case - " & MAXIS_case_number & ", appears to have a duplicate UNEA panel created from the Ex Parte Report run." & vbCr

MAXIS_footer_monht = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

Call navigate_to_MAXIS_screen("STAT", "MEMB")

member_list = " "
memb_row = 5
Do
	EMReadScreen ref_numb, 2, memb_row, 3
	member_list = member_list & ref_numb & " "
	memb_row = memb_row + 1
	EMReadScreen next_ref_numb, 2, memb_row, 3
Loop until next_ref_numb = "  "

member_list = trim(member_list)
member_array = split(member_list)

Call navigate_to_MAXIS_screen("STAT", "UNEA")
email_body = email_body & "<BODY style=font-size:11pt;font-family:Courier New>"
' "Good Morning;<p>We have completed our main aliasing process for today.  All assigned firms are complete.  Please feel free to respond with any questions.<p>Thank you."

for each memb_numb in member_array
	EMWriteScreen memb_numb, 20, 76
	Call write_value_and_transmit("01", 20, 79)

	EMReadScreen version_number, 1, 2, 78
	If version_number <> "0" Then
		Do
			EMReadScreen instance, 1, 2, 73
			email_body = email_body & "<br>" & "<h3>PANEL INFORMATION: UNEA " & memb_numb & " 0" & instance & ":</h3>"
			For unea_row = 2 to 19
				EMReadScreen unea_line, 57, unea_row, 24
				unea_line = replace(unea_line, "    ", "&emsp;&ensp;")
				unea_line = replace(unea_line, "  ", "&emsp;")
				unea_line = replace(unea_line, " ", "&ensp;")
				email_body = email_body & "&emsp;&emsp;" & unea_line & "<br>"
			Next
			email_body = email_body & "<br>" & "------------------------------------------------------------------------------------" & "<br>"

			transmit
			EMReadScreen enter_a_valid, 13, 24, 2
		Loop until enter_a_valid = "ENTER A VALID"

	End If
Next
email_body = email_body & "</BODY>"

call find_user_name(worker_name)
email_body = email_body & "<br>" & "Thank you!" & "<br>" & worker_name & "<br>" & now

' Function create_outlook_email(email_from, email_recip, email_recip_CC, email_recip_bcc, email_subject, email_importance, include_flag, email_flag_text, email_flag_days, email_flag_reminder, email_flag_reminder_days, email_body, include_email_attachment, email_attachment_array, send_email)
Call create_outlook_email("", "hsph.ews.bluezonescripts@hennepin.us", "", "", email_subject, 1, False, "", "", "", "", email_body, False, "", True)
end_msg = "EMAIL SENT TO THE A & I Team"
end_msg = end_msg & vbCr & "No additional report is needed."
end_msg = end_msg & vbCr & "You can now update the case and process as needed."
Call script_end_procedure(end_msg)

