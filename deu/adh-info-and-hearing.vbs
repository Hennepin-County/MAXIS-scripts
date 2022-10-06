'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - DEU-ADH INFO HEARING.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 120           'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
'Example: CALL changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("10/06/2022", "Update to remove hard coded DEU signature all DEU scripts.", "MiKayla Handley, Hennepin County") '#316
CALL changelog_update("09/16/2022", "Update to ensure Worker Signature is in all scripts that CASE/NOTE.", "MiKayla Handley, Hennepin County") '#316
CALL changelog_update("09/30/2019", "Removed FSS Data Team from automated emails per request.", "Ilse Ferris, Hennepin County")
CALL changelog_update("01/29/2018", "Updated to correct for member # error.", "MiKayla Handley, Hennepin County")
CALL changelog_update("11/08/2017", "Updated to add first name of memb to casenote.", "MiKayla Handley, Hennepin County")
CALL changelog_update("7/07/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The script
EMCONNECT ""
CALL MAXIS_case_number_finder(MAXIS_case_number)
memb_number = "01"

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 176, 85, "ADH INFORMATION"
  EditBox 55, 5, 45, 15, MAXIS_case_number
  EditBox 145, 5, 25, 15, memb_number
  DropListBox 90, 25, 80, 15, "Select One:"+chr(9)+"ADH waiver signed"+chr(9)+"Hearing Held", ADH_option
  EditBox 65, 45, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 65, 65, 50, 15
    CancelButton 120, 65, 50, 15
  Text 5, 10, 45, 10, "Case number:"
  Text 110, 10, 30, 10, "Memb#:"
  Text 5, 30, 75, 10, "Select an ADH option:"
  Text 5, 50, 60, 10, "Worker signature:"
EndDialog


Do
	Do
        err_msg = ""
		Dialog Dialog1
		cancel_confirmation
        Call validate_MAXIS_case_number(err_msg, "*")
		IF IsNumeric(memb_number) = false or len(memb_number) <> 2 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid two-digit member number."
        IF worker_signature = "" THEN err_msg = err_msg & vbCr & "* Please sign your case note."
		IF ADH_option = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select an ADH action."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
 	CALL check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

'------------------------------------------------------------------------------------getting the case name
CALL navigate_to_MAXIS_screen("STAT", "MEMB")
EMwritescreen memb_number, 20, 76
transmit

EMReadScreen MEMB_panel_check, 4, 2, 48
If MEMB_panel_check <> "MEMB" THEN script_end_procedure ("Could not access MEMB.")

EMReadscreen memb_name, 12, 6, 63
memb_name = replace(memb_name, "_", "")

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
IF ADH_option = "ADH waiver signed" THEN
    BeginDialog Dialog1, 0, 0, 326, 100, "ADH waiver signed"
      EditBox 100, 5, 40, 15, date_waiver_signed
      DropListBox 250, 5, 65, 15, "Select One:"+chr(9)+"CASH"+chr(9)+"SNAP"+chr(9)+"CASH & SNAP", Program_droplist
      EditBox 30, 35, 40, 15, start_date
      EditBox 30, 55, 40, 15, end_date
      EditBox 135, 35, 20, 15, months_disq
      EditBox 135, 55, 40, 15, DISQ_begin_date
      EditBox 250, 35, 65, 15, fraud_claim_number
      DropListBox 250, 55, 65, 15, "Select One:"+chr(9)+"Unknown"+chr(9)+"Judy Grandel"+chr(9)+"Chris Gormley"+chr(9)+"Keyatta Hill"+chr(9)+"Amanda Lange"+chr(9)+"Kimberly Littlejohn"+chr(9)+"Jonathan Martin"+chr(9)+"Ryan Swanson"+chr(9)+"Scott Benedict", Fraud_investigator
      EditBox 55, 80, 125, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 215, 80, 50, 15
        CancelButton 270, 80, 50, 15
      Text 75, 40, 55, 10, "Months of DISQ:"
      Text 210, 10, 35, 10, "Programs:"
      Text 10, 40, 20, 10, "Start:"
      Text 75, 60, 60, 10, "DISQ Begin Date:"
      Text 195, 40, 50, 10, "Claim Number:"
      GroupBox 5, 25, 175, 50, "Period of offense:"
      GroupBox 190, 25, 130, 50, "Fraud Information"
      Text 10, 85, 45, 10, "Other Notes:"
      Text 205, 60, 40, 10, "Investigator:"
      Text 10, 60, 15, 10, "End:"
      Text 5, 10, 95, 10, "Date ADH Hearing was held:"
    EndDialog
    DO
    	Do
    		err_msg = ""
    		Dialog Dialog1
    		cancel_confirmation
    		IF isdate(date_waiver_signed) = false THEN err_msg = err_msg & vbNewLine & "* Please enter date waiver was signed."
    		IF program_droplist = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select the program."
    		IF isdate(start_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter start date."
    		IF isdate(end_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter end date."
    		'IF IsNumeric(months_disq) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the amount of disqualification months."
    		IF isdate(DISQ_begin_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter DISQ beign date."
    		IF trim(fraud_claim_number) = "" and Fraud_investigator <> "Select One:" THEN err_msg = err_msg & vbNewLine & "* Enter both the fraud case number AND the Fraud Investigator's name, or clear the non-applicable info."
    		IF trim(fraud_claim_number) <> "" and Fraud_investigator = "Select One:"  THEN err_msg = err_msg & vbNewLine & "* Enter both the fraud case number AND the Fraud Investigator's name, or clear the non-applicable info."
    		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    	Loop until err_msg = ""
    	CALL check_for_password(are_we_passworded_out)
    LOOP UNTIL check_for_password(are_we_passworded_out) = False

    DISQ_end_date = DateAdd("M", months_disq, DISQ_begin_date)
    IF Fraud_investigator = "Select One:"  THEN Fraud_investigator = ""
    'IF Fraud_investigator = "" 						THEN fraud_email = ""
    IF Fraud_investigator = "Judy Grandel" 			THEN fraud_email = "Judith.Grandel@hennepin.us"
    IF Fraud_investigator = "Chris Gormley"		 	THEN fraud_email = "Chris.Gormley@hennepin.us"
    IF Fraud_investigator = "Keyatta Hill" 	 		THEN fraud_email = "Keyatta.Hill@hennepin.us"
    IF Fraud_investigator = "Amanda Lange" 			THEN fraud_email = "Amanda.Lange@hennepin.us"
    IF Fraud_investigator = "Kimberly Littlejohn"	THEN fraud_email = "Kimberly.Littlejohn@Hennepin.us"
    IF Fraud_investigator = "Jonathan Martin" 		THEN fraud_email = "Jonathan.Martin@Hennepin.us"
    IF Fraud_investigator = "Ryan Swanson" 			THEN fraud_email = "Ryan.Swanson@hennepin.us"
    IF Fraud_investigator = "Scott Benedict" 		THEN fraud_email = "Scott.Benedict@hennepin.us"

'The 1st case note-------------------------------------------------------------------------------------------------
 	start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	CALL write_variable_in_case_note("-----1st Fraud DISQ/Claims (" & memb_name & ") ADH Waiver Signed-----")
	CALL write_variable_in_case_note("Client signed ADH waiver on: " & date_waiver_signed & " waiving his/her right to an Administrative Disqualification Hearing for wrongfully obtaining public assistance. This disqualification is not for any other household member and does not affect MA eligibility.")
	CALL write_variable_in_case_note("* Programs: " & program_droplist)
	CALL write_variable_in_case_note("* Period of Offense: " & start_date & " - " & end_date)
	CALL write_variable_in_case_note("* Client is subject to a " & months_disq & " month DISQ from " & DISQ_begin_date & " - "  & DISQ_end_date & ".")
	IF program_droplist <> "SNAP"  THEN CALL write_variable_in_case_note("* Because member " & memb_number & " is DQ'd from MFIP, client is also barred from FS for that same period of time.")
	IF fraud_claim_number <> "" THEN
	 CALL write_variable_in_case_note("----- ----- -----")
	 CALL write_bullet_and_variable_in_case_note("Fraud claim number", fraud_claim_number)
	 CALL write_bullet_and_variable_in_case_note("Fraud Investigator", Fraud_investigator)
	END IF
	CALL write_variable_in_case_note("* Email sent to team: L. Bloomquist, and TTL")
    CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
	CALL write_variable_in_case_note("----- ----- ----- ----- -----")
    CALL write_variable_in_CASE_NOTE(worker_signature)
	'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
	CALL create_outlook_email("Lea.Bloomquist@hennepin.us", "HSPH.ES.TEAM.TTL@hennepin.us;" & fraud_email, "1st Fraud DISQ/Claims--ADH Waiver Signed for #" &  MAXIS_case_number, "Member #: " & memb_number & vbcr & "Client signed ADH waiver on: " & date_waiver_signed & " waiving his/her right to an Administrative Disqualification Hearing for wrongfully obtaining public assistance." & vbcr & "Programs: " & program_droplist & vbcr & "Period of Offense: " & start_date & " - " & end_date & vbcr & "See case notes for further details.", "", False)
END IF

IF ADH_option = "Hearing Held" THEN
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 326, 120, "Hearing Held"
      EditBox 100, 5, 40, 15, hearing_date
      EditBox 100, 25, 40, 15, date_order_signed
      DropListBox 250, 5, 65, 15, "Select One:"+chr(9)+"CASH"+chr(9)+"SNAP"+chr(9)+"CASH & SNAP", Program_droplist
      EditBox 30, 55, 40, 15, end_date
      EditBox 30, 75, 40, 15, start_date
      EditBox 135, 55, 20, 15, months_disq
      EditBox 135, 75, 40, 15, DISQ_begin_date
      EditBox 250, 55, 65, 15, fraud_claim_number
      DropListBox 250, 75, 65, 15, "Select One:"+chr(9)+"Unknown"+chr(9)+"Judy Grandel"+chr(9)+"Chris Gormley"+chr(9)+"Keyatta Hill"+chr(9)+"Amanda Lange"+chr(9)+"Kimberly Littlejohn"+chr(9)+"Jonathan Martin"+chr(9)+"Ryan Swanson"+chr(9)+"Scott Benedict", Fraud_investigator
      EditBox 55, 100, 125, 15, other_notes
      ButtonGroup ButtonPressed
        OkButton 215, 100, 50, 15
        CancelButton 270, 100, 50, 15
      Text 75, 60, 55, 10, "Months of DISQ:"
      Text 210, 10, 35, 10, "Programs:"
      Text 10, 60, 20, 10, "Start:"
      Text 75, 80, 60, 10, "DISQ Begin Date:"
      Text 195, 60, 50, 10, "Claim Number:"
      GroupBox 5, 45, 175, 50, "Period of offense:"
      GroupBox 190, 45, 130, 50, "Fraud Information"
      Text 10, 105, 45, 10, "Other Notes:"
      Text 205, 80, 45, 10, "Investigator:"
      Text 10, 80, 15, 10, "End:"
      Text 20, 30, 80, 10, "Date order was signed:"
      Text 5, 10, 95, 10, "Date ADH Hearing was held:"
    EndDialog
	DO
		Do
			err_msg = ""
			Dialog Dialog1
			cancel_confirmation
			IF isdate(hearing_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the hearing date."
			IF program_droplist = "Select One:" THEN err_msg = err_msg & vbNewLine & "* Please select the program."
			IF isdate(date_order_signed) = false THEN err_msg = err_msg & vbNewLine & "* Please enter date order was signed"
			IF isdate(start_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter start date"
			IF isdate(end_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter end date"
			IF IsNumeric(months_disq) = false THEN err_msg = err_msg & vbNewLine & "* Please enter the amount of disqualification months."
			IF isdate(DISQ_begin_date) = false THEN err_msg = err_msg & vbNewLine & "* Please enter date DISQ begins"
			IF trim(fraud_claim_number) = "" and Fraud_investigator <> "Select One:" THEN err_msg = err_msg & vbNewLine & "* Enter both the fraud case number AND the Fraud Investigator's name, or clear the non-applicable info."
			IF trim(fraud_claim_number) <> "" and Fraud_investigator = "Select One:"  THEN err_msg = err_msg & vbNewLine & "* Enter both the fraud case number AND the Fraud Investigator's name, or clear the non-applicable info."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		Loop until err_msg = ""
		CALL check_for_password(are_we_passworded_out)
	LOOP UNTIL check_for_password(are_we_passworded_out) = False

	DISQ_end_date = DateAdd("M", months_disq, DISQ_begin_date)
	IF Fraud_investigator = "Unknown" THEN Fraud_investigator = ""
	'IF Fraud_investigator = "" 						THEN fraud_email = ""
	IF Fraud_investigator = "Judy Grandel" 			THEN fraud_email = "Judith.Grandel@hennepin.us"
	IF Fraud_investigator = "Chris Gormley"		 	THEN fraud_email = "Chris.Gormley@hennepin.us"
	IF Fraud_investigator = "Keyatta Hill" 	 		THEN fraud_email = "Keyatta.Hill@hennepin.us"
	IF Fraud_investigator = "Amanda Lange" 			THEN fraud_email = "Amanda.Lange@hennepin.us"
	IF Fraud_investigator = "Kimberly Littlejohn"	THEN fraud_email = "Kimberly.Littlejohn@Hennepin.us"
	IF Fraud_investigator = "Jonathan Martin" 		THEN fraud_email = "Jonathan.Martin@Hennepin.us"
	IF Fraud_investigator = "Ryan Swanson" 			THEN fraud_email = "Ryan.Swanson@hennepin.us"
	IF Fraud_investigator = "Scott Benedict" 		THEN fraud_email = "Scott.Benedict@hennepin.us"

'The 2nd case note-------------------------------------------------------------------------------------------------
	start_a_blank_case_note      'navigates to case/note and puts case/note into edit mode
	 CALL write_variable_in_case_note("-----1st Fraud DISQ/Claim (" & memb_name & ")  ADH Hearing Held-----")
	 CALL write_variable_in_case_note("Administrative Disqualification Hearing for Wrongfully Obtaining Public Assistance was held on: " & hearing_date & " " & "Order was signed: " & date_order_signed & "This disqualification is not for any other household member and does not affect MA eligibility.")
	 CALL write_variable_in_case_note("----- ----- -----")
     CALL write_bullet_and_variable_in_case_note("Hearing Date", hearing_date)
     CALL write_bullet_and_variable_in_case_note("Date order signed ", date_order_signed)
	 CALL write_bullet_and_variable_in_case_note("Programs", program_droplist)
	 CALL write_variable_in_case_note("* Period of Offense: " & start_date & " - " & end_date)
	 CALL write_variable_in_case_note("* Client is subject to a " & months_disq & " month DISQ from " & DISQ_begin_date & " - "  & DISQ_end_date)
	 IF program_droplist <> "SNAP"  THEN CALL write_variable_in_case_note("* Because member " & memb_number & " is DQ'd from MFIP, client is also barred from FS for that same period of time.")
	 IF fraud_claim_number <> "" THEN
		 CALL write_variable_in_case_note("----- ----- -----")
		 CALL write_bullet_and_variable_in_case_note("Fraud claim number", fraud_claim_number)
		 CALL write_bullet_and_variable_in_case_note("Fraud Investigator", Fraud_investigator)
	 END IF
	 CALL write_variable_in_case_note("* Email sent to team: L. Bloomquist, and TTL.")
     CALL write_bullet_and_variable_in_case_note("Other Notes", other_notes)
	 CALL write_variable_in_case_note("----- ----- ----- ----- -----")
     CALL write_variable_in_CASE_NOTE(worker_signature)
	 'Drafting an email. Does not send the email!!!!
	 'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
	 CALL create_outlook_email("Lea.Bloomquist@hennepin.us", "HSPH.ES.TEAM.TTL@hennepin.us;" & fraud_email, "1st Fraud DISQ/Claims--ADH Hearing Held for #" &  MAXIS_case_number, "Member #: " & memb_number & vbcr & "Administrative Disqualification Hearing for Wrongfully Obtaining Public Assistance was held on: " & hearing_date & vbcr & "Order was signed: " & date_order_signed & vbcr & "Programs: " & program_droplist & vbcr & "Period of Offense: " & start_date & " - " & end_date & vbcr & "See case notes for further details.", "", False)
END IF
script_end_procedure_with_error_report("Please select the applicable team in the drafted email, any additional notes required and send the email regarding ADH information.")
