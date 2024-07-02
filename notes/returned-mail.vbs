Option Explicit
'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - RETURNED MAIL.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 360          'manual run time in seconds
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
'END FUNCTIONS LIBRARY BLOCK=================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("06/27/2024", "Updated POLI/TEMP links to navigate to SharePoint instead of POLI/TEMP in MAXIS.", "Mark Riegel, Hennepin County")
call changelog_update("05/17/2024", "Updated to reflect unclear information for active, SNAP-only 6-month reporting cases where mail is returned without a forwarding address.", "Mark Riegel, Hennepin County")
call changelog_update("01/26/2023", "Removed term 'ECF' from the case note per DHS guidance, and referencing the case file instead.", "Ilse Ferris, Hennepin County")
call changelog_update("08/03/2022", "Remove ability to select both residential and mailing address. Created an option for received in error and address HC only handling.", "MiKayla Handley, Hennepin County") '#927 '
call changelog_update("06/03/2022", "Updates for HC active/pending procedure and added handling for entering PACT panel.", "MiKayla Handley, Hennepin County") '#427 & #365 '
call changelog_update("03/12/2021", "Updated handling for current address confirmation.", "MiKayla Handley, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris, Hennepin County")
call changelog_update("02/13/2020", "Updated the zip code to only allow for 5 characters.", "MiKayla Handley, Hennepin County")
call changelog_update("06/06/2019", "Initial version. Re-written per POLI/TEMP.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-----------------------------------------------------------------------------------------------------------------
'global variable'
Dim name_of_script, start_time, STATS_counter, STATS_manualtime, STATS_denomination, run_locally, use_master_branch, req, fso, FuncLib_URL, critical_error_msgbox, run_another_script_fso, _
fso_command, text_from_the_other_script

Dim Dialog1, MAXIS_case_number, date_received, METS_case_number, ADDR_actions, worker_signature, is_this_priv, worker_number, notes_on_address, _
resi_addr_line_one, resi_addr_line_two, resi_addr_street_full, resi_addr_city, resi_addr_state, resi_addr_zip, resi_county, addr_verif, homeless_addr, reservation_addr, living_situation, _
reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city_line, mail_state_line, mail_zip_line, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, _
type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted, case_invalid_error, snap_or_cash_case, _
family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, residential_address_confirmed, mailing_address_confirmed, returned_mail, _
verifications_requested, other_notes, due_date, POLI_TEMP_PACT_button, panel_number, panel_number_check, TIKL_note_text, _
SNAP_POLI_TEMP_button, CASH_POLI_TEMP_button, one_source_button, err_msg, are_we_passworded_out, case_active, MAXIS_footer_year, MAXIS_footer_month, case_pending, case_rein, ma_case, _
msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, _
emer_type, case_status, list_active_programs, list_pending_programs, ADHI_button, HSR_manual_button, new_addr_line_one, new_addr_line_two, new_addr_city, new_addr_state, _
new_addr_zip, new_addr_street_full, begining_of_footer_month, resi_street_full, county_code, date_verifications_requested, error_message, mets_addr_correspondence, open_cash1, open_cash2, _
pact_pop_up, closing_message, end_msg, received_error_confirmation, HC_only, row, col, unclear_information_case, no_SNAP, status_row, app_status, vers_number, reporting_status

EMConnect ""                                        'Connecting to BlueZone
CALL MAXIS_case_number_finder(MAXIS_case_number)    'Grabbing the CASE Number
CALL Check_for_MAXIS(false)                         'Ensuring we are not passworded out
CALL back_to_self                                   'added to ensure we have the time to update and send the case in the background

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Running the initial dialog
BeginDialog Dialog1, 0, 0, 221, 225, "RETURNED MAIL"
  EditBox 75, 65, 50, 15, MAXIS_case_number
  EditBox 75, 85, 50, 15, date_received
  EditBox 75, 105, 50, 15, METS_case_number
  DropListBox 75, 125, 140, 15, "Select:"+chr(9)+"Forwarding Address Provided"+chr(9)+"No Forwarding Address Provided"+chr(9)+"No Response Received"+chr(9)+"Address Confirmed - Received in Error", ADDR_actions
  EditBox 70, 185, 145, 15, worker_signature
  ButtonGroup ButtonPressed
    PushButton 135, 65, 80, 15, "SNAP TE02.08.012", SNAP_POLI_TEMP_button
    PushButton 135, 85, 80, 15, "CASH TE02.08.011", CASH_POLI_TEMP_button
    PushButton 135, 105, 80, 15, "ONEsource", one_source_button
    OkButton 110, 205, 50, 15
    CancelButton 165, 205, 50, 15
  Text 5, 70, 50, 10, "Case Number:"
  Text 5, 90, 50, 10, "Date Received:"
  Text 5, 110, 70, 10, "METS Case Number:"
  Text 5, 190, 60, 10, "Worker Signature:"
  Text 10, 15, 195, 10, "Best practice is that you make an attempt to call the resident."
  Text 10, 30, 200, 25, "If you have contacted the resident, the CLIENT CONTACT script should be used, if you were unable to reach the them, continue using this script. "
  GroupBox 5, 140, 210, 40, "Note:"
  Text 10, 150, 160, 10, "Do not make any changes to STAT/ADDR."
  Text 10, 160, 195, 15, "Do not enter a ? or unknown or other county codes on the ADDR panel."
  GroupBox 5, 5, 210, 55, "Request for Contact:"
  Text 5, 130, 50, 10, "Returned Mail:"
EndDialog

DO
    DO
    	err_msg = ""
    	DIALOG Dialog1
    	cancel_confirmation
    	Call validate_MAXIS_case_number(err_msg, "*")
		IF isdate(date_received) = FALSE THEN
			err_msg = err_msg & vbnewline & "Please enter the date."
		Else
			IF Cdate(date_received) > cdate(date) = TRUE THEN err_msg = err_msg & vbnewline & "You must enter an actual date that is not in the future and is in the footer month that you are working in."
		End If
		IF ADDR_actions = "Select:" THEN err_msg = err_msg & vbCr & "Please choose an action for the returned mail."
		IF worker_signature = "" THEN err_msg = err_msg & vbCr & "Please sign your case note."
		IF ButtonPressed = CASH_POLI_TEMP_button THEN run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:b:/r/sites/hs-es-poli-temp/Documents%203/TE%2002.08.011%20RETURNED%20MAIL%20PROCESSING%20-%20CASH.pdf?csf=1&web=1&e=hBQIUx"
		If ButtonPressed = SNAP_POLI_TEMP_button THEN run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:b:/r/sites/hs-es-poli-temp/Documents%203/TE%2002.08.012%20RETURNED%20MAIL%20PROCESSING%20-%20SNAP.pdf?csf=1&web=1&e=QANcp0"
		IF ButtonPressed = one_source_button THEN run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://www.dhs.state.mn.us/cs/login/login.htm"
		IF ButtonPressed = one_source_button or ButtonPressed = SNAP_POLI_TEMP_button or ButtonPressed = CASH_POLI_TEMP_button THEN
			err_msg = "LOOP"
		Else                                                'If the instructions button was NOT pressed, we want to display the error message if it exists.
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		End If
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = False					'loops until user passwords back in

'Defaulting unclear_information to false
unclear_information_case = False

'Determine if case falls under unclear information
If ADDR_actions = "No Forwarding Address Provided" Then
	
	'Pull case details from CASE/CURR, maintains connection to DAIL
	Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

	If case_active = TRUE AND list_active_programs = "SNAP" AND list_pending_programs = "" Then

		'Navigate to ELIG/FS from CASE/CURR to maintain tie to DAIL
		EMWriteScreen "ELIG", 20, 22
		Call write_value_and_transmit("FS  ", 20, 69)

		EMReadScreen no_SNAP, 10, 24, 2
		If no_SNAP <> "NO VERSION" then						'NO SNAP version means no determination
			EMWriteScreen "99", 19, 78
			transmit
			'This brings up the FS versions of eligibility results to search for approved versions
			status_row = 7
			Do
				EMReadScreen app_status, 8, status_row, 50
				app_status = trim(app_status)
				If app_status = "" then
					PF3
					exit do 	'if end of the list is reached then exits the do loop
				End if
				If app_status = "UNAPPROV" Then status_row = status_row + 1
			Loop until app_status = "APPROVED" or app_status = ""

			If app_status = "APPROVED" then
				EMReadScreen vers_number, 1, status_row, 23
				Call write_value_and_transmit(vers_number, 18, 54)
				Call write_value_and_transmit("FSSM", 19, 70)
				EmReadscreen reporting_status, 12, 8, 31
				reporting_status = trim(reporting_status)

				If reporting_status = "SIX MONTH" Then
					'Case falls under unclear information - it is a SNAP-only, 6-month case
					unclear_information_case = True
				End If
			End if
		End If
	End If
	Call back_to_SELF
End If

'If case is an active SNAP-only, 6-month reporting case then it will create CASE/NOTE
If unclear_information_case = True Then

	Dialog1 = "" 'Running the unclear info dialog
	BeginDialog Dialog1, 0, 0, 221, 145, "Undeliverable Returned Mail"
	EditBox 80, 80, 135, 15, returned_mail
	EditBox 80, 100, 135, 15, other_notes
	ButtonGroup ButtonPressed
		OkButton 110, 125, 50, 15
		CancelButton 165, 125, 50, 15
	Text 10, 5, 50, 10, "Case Number:"
	Text 65, 5, 95, 10, MAXIS_case_number
	GroupBox 5, 20, 210, 55, "Request for Contact:"
	Text 10, 30, 200, 40, "This case is an active SNAP-only, 6-month reporter so the undeliverable mail is unclear information. This information should be entered into a CASE/NOTE and should be followed-up on at the next scheduled certification action or periodic report about the potential change in residence."
	Text 10, 85, 65, 10, "What was returned:"
	Text 10, 105, 45, 10, "Other notes:"
	EndDialog

	DO
		DO
			err_msg = ""
			DIALOG Dialog1
			cancel_confirmation

			If trim(returned_mail) = "" Then err_msg = err_msg & vbCr & "Please indicate what the returned mail was."
			IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
		Loop until err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = False					'loops until user passwords back in

	'Navigate to CASE/NOTE
	CALL start_a_blank_CASE_NOTE()
	CALL write_variable_in_CASE_NOTE("RETURNED MAIL RECEIVED - " & ADDR_actions)
	CALL write_bullet_and_variable_in_CASE_NOTE("RECEIVED ON", date_received)
	CALL write_bullet_and_variable_in_CASE_NOTE("WHAT WAS RETURNED", returned_mail)
	CALL write_bullet_and_variable_in_CASE_NOTE("OTHER NOTES", other_notes)
	CALL write_variable_in_case_note("---")
	CALL write_variable_in_case_note("REVIEW POTENTIAL CHANGE IN RESIDENCE WITH RESIDENT AT RENEWAL/RECERTIFICATION AS CASE IS A SNAP 6-MONTH REPORTING CASE. SEE CM 0007.03.02 - SIX-MONTH REPORTING. VERIFICATION REQUEST IS NOT REQUIRED AND NEGATIVE ACTION SHOULD NOT BE TAKEN ON THE CASE.")
	CALL write_variable_in_case_note("---")
	CALL write_variable_in_case_note(worker_signature)

	script_end_procedure("CASE/NOTE successfully added. No negative action should be taken on case as it is a SNAP-only six-month reporting case. Script will now end.")
End If

'setting the footer month to make the updates in'
CALL convert_date_into_MAXIS_footer_month(date_received, MAXIS_footer_month, MAXIS_footer_year)
MAXIS_footer_month_confirmation

CALL navigate_to_MAXIS_screen_review_PRIV("STAT", "ADDR", is_this_priv)
IF is_this_priv = TRUE THEN script_end_procedure("This case is privileged, please request access to the case and run the script again.")

EMReadScreen worker_number, 7, 21, 21              'reading the current(primary) workers number '
IF left(worker_number, 4) <> "X127" THEN script_end_procedure("*** Out of County ***" & vbCr & worker_number & vbCr & "This case is Out of County the script will now end.")

' CALL read_ADDR_panel(addr_eff_date, addr_future_date, resi_addr_line_one, resi_addr_line_two, resi_addr_city, resi_addr_state, resi_addr_zip, mail_line_one, mail_line_two, mail_city_line, mail_state_line, mail_zip_line, living_situation, living_sit_line, homeless_line, addr_phone_1A)
Call access_ADDR_panel("READ", notes_on_address, resi_addr_line_one, resi_addr_line_two, resi_addr_street_full, resi_addr_city, resi_addr_state, resi_addr_zip, resi_county, addr_verif, homeless_addr, reservation_addr, living_situation, reservation_name, mail_line_one, mail_line_two, mail_street_full, mail_city_line, mail_state_line, mail_zip_line, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)

EMReadScreen case_invalid_error, 72, 24, 2 'if a person enters an invalid footer month for the case the script will attempt to navigate'
case_invalid_error = trim(case_invalid_error)
If case_invalid_error = "ENTER A VALID COMMAND OR PF-KEY" Then case_invalid_error = ""			'this message is for if you press transmit on a single instance panel and does not indicate an error. Ignoring it
IF trim(case_invalid_error) <> "" THEN script_end_procedure("*** NOTICE!!! ***" & vbCr & case_invalid_error & vbCr & "Please resolve for the script to continue.")

CALL determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

snap_or_cash_case = False 'we need to establish if snap or cash are active, pending, for how the case is handled'
If family_cash_case = True Then snap_or_cash_case = True
If adult_cash_case = True Then snap_or_cash_case = True
If grh_case = True Then snap_or_cash_case = True
If snap_case = True Then snap_or_cash_case = True

If ma_case = True or msp_case = true then
    'If other programs are active/pending then no notice is necessary
    If  family_cash_case = True OR _
        mfip_case = True OR _
        dwp_case = True OR _
        adult_cash_case = True OR _
        ga_case = True OR _
        msa_case = True OR _
        grh_case = True OR _
        snap_case = True OR _
        emer_case = True then
        HC_only = False
    Else
        HC_only = True
    End if
End if

'-------------------------------------------------------------------------------------------------DIALOG
IF ADDR_actions <> "No Response Received" THEN
    residential_address_confirmed = "YES"
	mailing_address_confirmed = "NO"
    If mail_line_one <> "" Then
    	residential_address_confirmed = "NO"
    	mailing_address_confirmed = "YES"
    End If
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 566, 155, "RETURNED MAIL "
      DropListBox 220, 15, 55, 15, "YES"+chr(9)+"NO", residential_address_confirmed
      DropListBox 500, 15, 55, 15, "YES"+chr(9)+"NO", mailing_address_confirmed
      EditBox 90, 75, 190, 15, returned_mail
      Text 10, 25, 200, 25, resi_addr_street_full
      Text 10, 35, 210, 15, resi_addr_city &  ", "  & resi_addr_state & " "   & resi_addr_zip
      Text 290, 25, 200, 25, mail_street_full
      If trim(mail_city_line) <> "" THEN Text 290, 35, 200, 10, mail_city_line & ", "  & mail_state_line &  " "  & mail_zip_line
      GroupBox 285, 5, 275, 65, "Mailing Address"
      Text 290, 15, 210, 10, "Is this the address that the agency attempted to deliver mail to?"
      Text 10, 15, 210, 10, "Is this the address that the agency attempted to deliver mail to?"
      IF addr_future_date <> "" THEN Text 10, 55, 50, 10, "Future Date:"
      Text 60, 55, 70, 10, addr_future_date
      Text 345, 55, 70, 10, addr_eff_date
	  Text 290, 55, 50, 10, "Effective Date:"
      Text 5, 80, 65, 10, "What was returned:"
      GroupBox 5, 5, 275, 65, "Residential Address"
      GroupBox 285, 70, 275, 35, "Note:"
      Text 290, 80, 265, 10, "Send out a Request for Contact(Verification Request Form - DHS 2919) to the"
      Text 290, 90, 170, 10, "correct address and ensure STAT/ADDR is updated."
      Text 5, 120, 40, 10, "Other notes:"
      IF ADDR_actions = "Address Confirmed - Received in Error" THEN
      	Text 5, 140, 200, 10, "How did you confirm the returned mail was received in error?"
       	EditBox 220, 135, 215, 15, received_error_confirmation
		EditBox 90, 95, 190, 15, other_notes
		EditBox 90, 115, 190, 15, notes_on_address
		Text 5, 100, 40, 10, "Other notes:"
  		Text 5, 120, 65, 10, "Notes on address:"
      ELSE
      	Text 5, 100, 85, 10, "Verification(s) requested:"
      	Text 5, 140, 35, 10, "Due date:"
		Text 5, 120, 40, 10, "Other notes:"
		Text 290, 120, 65, 10, "Notes on address:"
      	EditBox 90, 95, 190, 15, verifications_requested
		EditBox 90, 115, 190, 15, other_notes
      	EditBox 90, 135, 45, 15, due_date
      	EditBox 365, 115, 195, 15, notes_on_address
      END IF
       ButtonGroup ButtonPressed
       PushButton 220, 50, 55, 15, "CASE/ADHI", ADHI_button
       PushButton 500, 50, 55, 15, "HSR Manual", HSR_manual_button
       CancelButton 505, 135, 55, 15
       OkButton 445, 135, 55, 15
    EndDialog

    DO
        DO
    		err_msg = ""
        	DO
    			DIALOG Dialog1
    			cancel_without_confirmation
    			MAXIS_dialog_navigation
                IF buttonpressed = HSR_manual_button then CreateObject("WScript.Shell").Run("https://hennepin.sharepoint.com/teams/hs-es-manual/SitePages/return_Mail_Processing_for_SNAP.aspx") 'HSR manual policy page
            LOOP until ButtonPressed = -1
 			IF residential_address_confirmed = "YES" and mailing_address_confirmed = "YES" THEN err_msg = err_msg & vbCr & "Please confirm what the address the agency attempted to deliver mail to."
    		IF mailing_address_confirmed = "NO" and residential_address_confirmed = "NO" THEN err_msg = err_msg & vbCr & "Please confirm what the address the agency attempted to deliver mail to."
			IF mailing_address_confirmed = "YES" and mail_street_full = "" THEN err_msg = err_msg & vbCr & "Please confirm what mailing address the agency attempted to deliver mail to, this address appears to be blank."
			IF residential_address_confirmed = "YES" and resi_addr_street_full = "" THEN err_msg = err_msg & vbCr & "Please confirm the residential address the agency attempted to deliver mail to, this address appears to be blank."
			IF trim(returned_mail) = "" or len(returned_mail) < 3 THEN err_msg = err_msg & vbCr & "Please explain in detail what mail was returned."
			IF ADDR_actions = "Address Confirmed - Received in Error" THEN
				IF received_error_confirmation = "" or len(received_error_confirmation) < 5 THEN err_msg = err_msg & vbNewLine & "Please explain in detail how you confirmed the address on file is correct."
			ELSE
				IF isdate(due_date) = FALSE THEN err_msg = err_msg & vbnewline & "Please enter the verification(s) requested due date."
				IF trim(verifications_requested) = "" or len(verifications_requested) < 3 THEN  err_msg = err_msg & vbCr & "Please explain in detail the verification(s) still pending."
			END IF
			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    		LOOP UNTIL err_msg = ""
        CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not pass worded out of MAXIS, allows user to  assword back into MAXIS
    LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

'-------------------------------------------------------------------------------------------------DIALOG
	Dialog1 = "" 'Blanking out previous dialog detail
	IF ADDR_actions = "Forwarding Address Provided" THEN
		new_addr_state = "MN"
	    BeginDialog Dialog1, 0, 0, 206, 125, "RETURNED MAIL "
	      EditBox 40, 15, 155, 15, new_addr_line_one
	      EditBox 40, 35, 155, 15, new_addr_line_two
	      EditBox 40, 55, 155, 15, new_addr_city
	      EditBox 40, 75, 20, 15, new_addr_state
	      EditBox 155, 75, 40, 15, new_addr_zip
	      ButtonGroup ButtonPressed
	        OkButton 95, 105, 50, 15
	        CancelButton 150, 105, 50, 15
	      GroupBox 5, 5, 195, 95, "Forwarding Address:"
	      Text 10, 20, 25, 10, "Street:"
	      Text 120, 80, 30, 10, "Zip code:"
	      Text 10, 60, 15, 10, "City:"
	      Text 10, 80, 20, 10, "State:"
	    EndDialog

    	DO
    		DO
    			err_msg = ""
    			DIALOG Dialog1
    			cancel_without_confirmation
				new_addr_line_one = trim(UCASE(new_addr_line_one))
				new_addr_line_two = trim(UCASE(new_addr_line_two))
				new_addr_city = trim(UCASE(new_addr_city))
				new_addr_state = trim(new_addr_state)
				new_addr_zip = trim(new_addr_zip)
    			IF new_addr_line_one = "" THEN err_msg = err_msg & vbCr & "Please enter the street address."
    			IF new_addr_city = "" THEN err_msg = err_msg & vbCr & "Please enter the city."
    			IF new_addr_state = "" THEN err_msg = err_msg & vbCr & "Please enter the state."
    			IF new_addr_zip = "" OR (new_addr_zip <> "" AND len(new_addr_zip) > 5) THEN err_msg = err_msg & vbNewLine & "Please only enter a 5 digit zip code."     'Makes sure there is a numeric zip
    			IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    		LOOP UNTIL err_msg = ""
    		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not pass worded out of MAXIS, allows user to password back into MAXIS
    	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

		new_addr_street_full = ""
		begining_of_footer_month = MAXIS_footer_month & "/1/" & MAXIS_footer_year
		begining_of_footer_month = DateAdd("d", 0, begining_of_footer_month)
		If DateDiff("d", begining_of_footer_month, addr_eff_date) > 0 Then begining_of_footer_month = addr_eff_date
		Call access_ADDR_panel("WRITE", notes_on_address, resi_addr_line_one, resi_addr_line_two, resi_street_full, resi_addr_city, resi_addr_state, resi_addr_zip, county_code, addr_verif, homeless_addr, reservation_addr, living_situation, reservation_name, new_addr_line_one, new_addr_line_two, new_addr_street_full, new_addr_city, new_addr_state, new_addr_zip, begining_of_footer_month, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
	END IF
END IF 'forwarding address provided

IF ADDR_actions = "No Response Received" THEN
	Dialog1 = "" 'Blanking out previous dialog detail
	BeginDialog Dialog1, 0, 0, 351, 185, "Request for contact to the resident with no response"
  	 EditBox 105, 5, 50, 15, date_verifications_requested
  	 EditBox 55, 145, 285, 15, other_notes
  	 ButtonGroup ButtonPressed
    	 PushButton 235, 5, 105, 15, "PACT TE02.13.10", POLI_TEMP_PACT_button
    	 OkButton 235, 165, 50, 15
    	 CancelButton 290, 165, 50, 15
  	 Text 5, 10, 100, 10, "Date Verifications Requested:"
  	 Text 5, 25, 320, 10, "Allow 10 days for the resident to respond to the Verification Request before terminating benefits."
  	 Text 5, 40, 330, 10, "Approve ineligible results in ELIG. Send a closing notice 10 days before the effective date of closing."
  	 Text 5, 55, 340, 35, "Remember to enter a worker comment in SPEC/WCOM to add a detailed explanation of closure to the notice.  DHS suggested text for the SPEC/WCOM: Your mail has been returned to our agency.  On (insert date) you were sent a request to contact this agency because of this returned mail.  You can avoid having your case closed if you contact this agency by (insert the closing date deadline)."
  	 Text 15, 110, 265, 10, "CASH the script will enter code 3 in the Close/Deny field on the STAT/PACT Panel."
  	 Text 15, 125, 320, 20, "Health Care cannot be denied at this time for whereabouts unknown. Script will navigate directly to CASE/NOTE for HC only cases."
  	 Text 15, 95, 295, 10, "SNAP the script will enter code 4 when mail to the resident has been returned to the agency."
  	 Text 10, 150, 45, 10, "Other Notes:"
	EndDialog

	DO
		DO
	  		err_msg = ""
		    DIALOG Dialog1
		    cancel_without_confirmation
            If HC_only = False then
                IF IsDate(date_verifications_requested) = FALSE THEN
		    	    err_msg = err_msg & vbnewline & "Please enter the date the verifications were requested."
		        ELSE
		    	    IF DateDiff("d", date_verifications_requested, date) < 10 THEN err_msg = err_msg & vbnewline & "You must allow 10 days for the resident to respond to the Verification Request before terminating benefits."
		        END IF
            End if
		    IF ButtonPressed = POLI_TEMP_PACT_button THEN
		    	CALL view_poli_temp("02", "13", "10", "") 'TE02.13.10' STAT:  PACT
		    	err_msg = "LOOP"
	        END IF                                        	'If the instructions button was NOT pressed, we want to display the error message if it exists.
		    IF err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL err_msg = ""
	   	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not pass worded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in

	IF snap_or_cash_case = TRUE THEN 'per POLI/TEMP this only pertains to active cash and snap '
		CALL back_to_self ' so that we done with POLI/TEMP'
       	CALL MAXIS_background_check
		CALL navigate_to_MAXIS_screen("STAT", "PACT") 	'Checking to see if the PACT panel is empty, if not it create a new panel'
       	EMReadScreen panel_number, 1, 02, 73
       	IF panel_number = "0" then
       		EMWriteScreen "NN", 20,79 'cursor is automatically set to 06, 58'
       		TRANSMIT
       	ELSE
			PF9
		END IF
       	EMReadScreen open_cash1, 2, 6, 43 'I have to read these'
		EMReadScreen open_cash2, 2, 8, 43
		IF open_cash1 <> "  " THEN EMWriteScreen "3", 6, 58 'Enter code "3" (Refused/Failed Required Info)'
		IF open_cash2 <> "  " THEN EMWriteScreen "3", 8, 58 'Enter code "3" (Refused/Failed Required Info)'

       	IF snap_case = True THEN
			IF case_pending = True THEN
				EMWriteScreen "3", 12, 58 ''CASE IS PENDING, USE '1' OR '3' TO DENY
			ELSE
				EMWriteScreen "4", 12, 58 'Enter code "4" (Refused/Failed (FS Only))'
			END IF
		END IF
   		IF grh_case = True THEN EMWriteScreen "3", 10, 58 'Enter code "3" (Refused/Failed Required Info)'
       	TRANSMIT
		row = 13
		col = 14
		EMSearch "IS IT", row, col
 	    If row <> 0 Then
	    	Do
	    		EMReadScreen pact_pop_up, 45, row, col 'this always comes up to confirm but moves if there was a panel previously '
	    		IF pact_pop_up = "IS IT CORRECT POLICY TO USE A PACT PANEL? Y/N" THEN
					EMSearch "_", row, col
	    			EMWriteScreen "Y", row, col ' this is a pop up box asking if the selection is correct per poli/temp SEE TEMP TE02.13.10'
	    			TRANSMIT
	    		END IF
	       	Loop until trim(pact_pop_up) = "*" 'it wil never equal blank '
	    	EMReadScreen panel_number_check, 1, 02, 73
	    	IF panel_number_check = "0" THEN
	    		closing_message = "Unable to verify that the PACT panel was updated. Please verify and notify the BlueZone Script team."
	    	END IF
	    	EMReadScreen error_message,  74, 24, 02 'for script_run_lowdown-reading for messages that might be missed if they are not inhibiting'
	    	error_message = trim(error_message)
  	    END IF
    END IF 'if snap or cash are true'
    'we cannot close HC currently but this is the place for that handling'
END IF 'no response received '
    '----------------------------------------------------------------------------------------------------TIKL
IF ADDR_actions = "Forwarding Address Provided" or ADDR_actions = "No Forwarding Address Provided" THEN
    'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
    Call create_TIKL("Returned mail received, contact from resident should have occurred re: address change. If no response-verbal or written, please take appropriate action.", 10, date, True, TIKL_note_text)
END IF
'if there is no forwarding address is provided the only step we can take is to ensure we are sending out a request for verifications and closing in  timely manner"'
'starts a blank case note
'----------------------------------------------------------------------------------------------------CASENOTE
CALL back_to_SELF ' to ensure we are not caught on the dail'
CALL start_a_blank_CASE_NOTE()
CALL write_variable_in_CASE_NOTE("Returned mail received " & ADDR_actions)
CALL write_bullet_and_variable_in_CASE_NOTE("Received on", date_received)
CALL write_bullet_and_variable_in_CASE_NOTE("What was returned", returned_mail)
CALL write_bullet_and_variable_in_CASE_NOTE("METS case number", METS_case_number)
CALL write_bullet_and_variable_in_CASE_NOTE("Confirmed address is correct", received_error_confirmation)
IF mailing_address_confirmed = "YES" THEN  'Address Detail
	CALL write_variable_in_CASE_NOTE("* Returned mail received from (mailing): " & mail_line_one)
	If mail_line_two <> "" Then CALL write_variable_in_CASE_NOTE("                                        " & mail_line_two)
	CALL write_variable_in_CASE_NOTE("                                        " & mail_city_line & ", " & mail_state_line & " " &   mail_zip_line)
END IF
IF residential_address_confirmed = "YES" THEN
	CALL write_variable_in_CASE_NOTE("* Returned mail received from (residential): " & resi_addr_line_one)
	If resi_addr_line_two <> "" Then CALL write_variable_in_CASE_NOTE("                                             " & resi_addr_line_two)
	CALL write_variable_in_CASE_NOTE("                                             " & resi_addr_city & ", " & resi_addr_state & " " & resi_addr_zip)
END IF
IF homeless_addr = "Yes" Then Call write_variable_in_CASE_NOTE("* Household reported as homeless.")
IF reservation_addr = "Yes" THEN CALL write_variable_in_CASE_NOTE("* Reservation " & reservation_name)
CALL write_bullet_and_variable_in_CASE_NOTE("Address detail", notes_on_address)
CALL write_bullet_and_variable_in_case_note("Verification(s) requested", verifications_requested)
CALL write_bullet_and_variable_in_case_note("Verification(s) due", due_date)
IF ADDR_actions = "Forwarding Address Provided" THEN
	CALL write_variable_in_CASE_NOTE("* Forwarding address was on returned mail.")
	CALL write_variable_in_CASE_NOTE("* Mailing address updated:  " & new_addr_line_one)
	If new_addr_line_two <> "" Then CALL write_variable_in_CASE_NOTE("                            " & new_addr_line_two)
	CALL write_variable_in_CASE_NOTE("                            " & new_addr_city & ", " & new_addr_state & " " & new_addr_zip)
	CALL write_variable_in_case_note("* Resident must be provided 10 days to return requested verifications.")
ELSEIF ADDR_actions = "No Response Received" THEN
	CALL write_variable_in_CASE_NOTE ("* Case file reviewed for requested verifications.")
	CALL write_variable_in_CASE_NOTE("* Date verification(s) requested: " & date_verifications_requested)
	CALL write_variable_in_case_note("* Resident was provided 10 days to return requested verifications.")
	IF snap_or_cash_case = True THEN CALL write_variable_in_CASE_NOTE ("* PACT panel entered per POLI/TEMP TE02.13.10.")
	IF ma_case = True THEN CALL write_variable_in_CASE_NOTE ("* Cannot close HC cases for whereabouts unknown during the COVID-19 emergency.")
END IF
CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
CALL write_variable_in_CASE_NOTE ("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

error_message = error_message & ", "
script_run_lowdown = script_run_lowdown & vbCr & " Message: " & vbCr & error_message & vbCr & snap_or_cash_case & " snap or cash" & vbCr & ADDR_actions & " ADDR actions " & vbCr & " notes on address " & notes_on_address & vbCr & " resi address " & resi_addr_line_one & " " & resi_addr_line_two & " " & resi_addr_street_full & " " & resi_addr_city & " " & resi_addr_state & " " & resi_addr_zip & vbCr & "resi_county " & resi_county & vbCr & "addr_verif " & addr_verif & vbCr & "homeless_addr " & homeless_addr & vbCr & " reservation_addr " & reservation_addr & vbCr & " living situation " & living_situation & vbCr & " reservation name " & reservation_name & vbCr & " mail address " & mail_line_one & " " & mail_line_two & " " & mail_street_full & " " & mail_city_line & " " & mail_state_line & " " & mail_zip_line & vbCr & "addr_eff_date & addr_future_date " & addr_eff_date & addr_future_date & vbCr & "phone " & phone_one & phone_two & phone_three & vbCr & "addr_email " & addr_email & vbCr & "verif received " & verif_received & vbCr & " original information " & original_information & vbCr & "update attempted " & update_attempted & vbCr & "Verification Requested " & verifications_requested & vbCr & "new addr " & new_addr_line_one & " " & new_addr_line_two & " " & new_addr_city  & " " & new_addr_state & " " & new_addr_zip & vbCr & "county list" & county_code & vbCr & "Mets " & mets_addr_correspondence & mets_case_number & vbCr & "Other Notes " & other_notes
'Checks if this is a METS case and pops up a message box with instructions if the ADDR is incorrect.
IF METS_case_number <> "" THEN end_msg = "Please update the METS ADDR if you are able to. If unable, please forward the new ADDR information to the correct area (i.e. Change In Circumstance Process)"

IF ADDR_actions = "Forwarding Address Provided" or ADDR_actions = "No Forwarding Address Provided" THEN
	closing_message = closing_message & "Success! TIKL has been set for the ADDR verification requested. Reminder:  When a change reporting unit reports a change over the telephone or in person, the unit is not required to also report the change on a Change Report from. "  & vbCr & end_msg 'FOR EVERYTHING ELSE'
END IF
IF ADDR_actions = "No Response Received" THEN
    IF snap_or_cash_case = TRUE THEN
		closing_message = closing_message & "Success! The PACT panel and case note have been entered, please approve ineligible results in ELIG & enter using NOTICES SPEC/WCOM adding worker comments." & vbCr & end_msg 'WILL ONLY RUN IF SNAP OR CASH AND NO RESPONSE RCVD'
	ELSE
		closing_message = closing_message & "Success! A case note has been entered." & vbCr & end_msg 'this meets the requirement for HC'
	END IF
END IF
IF ADDR_actions = "Address Confirmed - Received in Error" THEN closing_message = closing_message & "Success! A case note has been entered." & vbCr & end_msg 'this meets the requirement for HC'

CALL script_end_procedure_with_error_report(closing_message)
'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step---------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------05/11/2022
'--Tab orders reviewed & confirmed----------------------------------------------05/11/2022
'--Mandatory fields all present & Reviewed--------------------------------------05/11/2022
'--All variables in dialog match mandatory fields-------------------------------06/14/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------05/11/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------05/11/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------06/14/2022
'--write_variable_in_CASE_NOTE function: confirm proper punctuation is used-----09/06/2022

'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------05/11/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------05/11/2022
'--PRIV Case handling reviewed -------------------------------------------------05/11/2022
'--Out-of-County handling reviewed----------------------------------------------05/11/2022	discussed with Ilse
'--script_end_procedures (w/ or w/o error messaging)----------------------------05/11/2022
'--BULK - review output of statistics and run time/count (if applicable)--------05/11/2022
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---09/06/2022

'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------05/14/2022
'--Incrementors reviewed (if necessary)-----------------------------------------N/A
'--Denomination reviewed -------------------------------------------------------N/A
'--Script name reviewed---------------------------------------------------------06/14/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------09/06/2022
'--Comment code-----------------------------------------------------------------06/14/2022
'--Update changelog for release/update------------------------------------------06/14/2022
'--Remove testing message boxes-------------------------------------------------06/14/2022
'--Remove testing code/unnecessary code-----------------------------------------09/06/2022
'--Review/update SharePoint instructions----------------------------------------06/14/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------06/14/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------06/14/2022
'--Complete misc. documentation (if applicable)---------------------------------06/14/2022
'--Update project team/issue contact (if applicable)----------------------------06/14/2022
'--Other Note-------------------------------------------------------------------SNAP 2. On STAT/ADDR, enter the new address from the returned mail envelope.  Enter "OT" in the verification field. We are not updating OT as it is in the residential area of the script
