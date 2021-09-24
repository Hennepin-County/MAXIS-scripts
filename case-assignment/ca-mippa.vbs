'===========================================================================================STATS
name_of_script = "CA - MIPPA.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 500
STATS_denominatinon = "C"
'===========================================================================================END OF STATS BLOCK

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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

' ===========================================================================================================CHANGELOG BLOCK
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")CALL changelog_update("04/16/2019", "Added requested language for denial dates and extra help.", "MiKayla Handley, Hennepin County")
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
CALL changelog_update("08/02/2019", "Removed error reporting, due to system limitations.", "MiKayla Handley, Hennepin County")
CALL changelog_update("04/16/2019", "Added requested language for denial dates and extra help.", "MiKayla Handley, Hennepin County")
CALL changelog_update("03/19/2019", "Added an error reporting option at the end of the script run.", "Casey Love, Hennepin County")
call changelog_update("11/06/2017", "Updates to handle when there are multiple PMI associated with the same client.", "MiKayla Handley, Hennepin County")
call changelog_update("10/10/2017", "Updates to correct dialog box error message and ensure the correct case number pulls through the whole script.", "MiKayla Handley, Hennepin County")
call changelog_update("10/10/2017", "Updates to correct action when case noting and updating REPT/MLAR.", "MiKayla Handley, Hennepin County")
call changelog_update("09/29/2017", "Updates to correct action if HC is already pending.", "MiKayla Handley, Hennepin County")
call changelog_update("08/21/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK
'---------------------------------------------------------------script
EMConnect ""
'Navigates to MIPPA Lis Application-Medicare Improvement for Patients and Providers (MIPPA)

CALL navigate_to_MAXIS_screen("REPT", "MLAR")
EMReadscreen current_panel_check, 4, 2, 54
IF current_panel_check <> "MLAR" THEN
    'script_end_procedure_with_error_report ("You are not on a MIPPA message. This script will stop")
    '-------------------------------------------------------------------------------------------------DIALOG
    Dialog1 = "" 'Blanking out previous dialog detail
    BeginDialog Dialog1, 0, 0, 176, 85, "REPT - MIPPA"
      OkButton 65, 65, 50, 15
      CancelButton 120, 65, 50, 15
      ButtonGroup ButtonPressed
      GroupBox 10, 5, 160, 55, "About this script:"
      Text 15, 20, 150, 35, "This script navigates to REPT/MLAR and guides you through the MIPPA process. This dialog will ensure if you are passworded out."
    EndDialog

    Do
    	dialog Dialog1
    	Cancel_without_confirmation
    	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
    Loop until are_we_passworded_out = false					'loops until user passwords back in
END IF

EMReadscreen appl_status, 2, 5, 17
IF appl_status <> "NO" THEN
 EMWriteScreen "NO", 5, 17
 TRANSMIT
END IF

'THERE ARE NO RECORDS THAT MATCH THE CRITERIA
EMReadScreen error_check, 75, 24, 2
error_check = TRIM(error_check)
IF error_check = "THERE ARE NO RECORDS THAT MATCH THE CRITERIA" THEN script_end_procedure(error_check & vbcr & "The script will now end.") '-------option to read from REPT need to checking for error msg'


row = 11 'this part should be a for next?' can we jsut do a cursor read for now?
DO
	EMReadScreen MLAR_maxis_name, 21, row, 5
	MLAR_maxis_name = TRIM(MLAR_maxis_name)
	    MLAR_info_confirmation = MsgBox("Press YES to confirm this is the MIPPA you wish to clear." & vbNewLine & "For the next match, press NO." & vbNewLine & vbNewLine & _
		"   " & MLAR_maxis_name, vbYesNoCancel, "Please confirm this match")
			IF MLAR_info_confirmation = vbNo THEN
				row = row + 1
				'msgbox "row: " & row
				IF row = 19 THEN
					PF8
					row = 7
				END IF
			END IF
			IF MLAR_info_confirmation = vbCancel THEN script_end_procedure_with_error_report ("It appears you have not made a selection please run the script again.")
			IF MLAR_info_confirmation = vbYes THEN 	EXIT DO
LOOP UNTIL MLAR_info_confirmation = vbYes

EMwritescreen "X", row, 03 'this will take us to REPT/MLAD'
TRANSMIT
	'navigates to MLAD
EMReadScreen MLAD_maxis_name, 22, 6, 20
MLAD_maxis_name = TRIM(MLAD_maxis_name)
EMReadScreen MLAD_SSN_number, 11, 7, 20
MLAD_SSN_number = trim(MLAD_SSN_number)
MLAD_SSN_number  = replace(MLAD_SSN_number , " ", "-")
EMReadScreen SSN_first, 3, 7, 20
EMReadScreen SSN_mid, 2, 7, 24
EMReadScreen SSN_last, 4, 7, 27
EMReadScreen appl_date, 8, 11, 20
appl_date = replace(appl_date, " ", "/")
'----------------------------------------------------------------------used for the dialog to appl
EMReadScreen client_dob, 8, 8, 20
client_dob = replace(client_dob, " ", "/")
EMReadScreen medi_number, 10, 10, 20
EMReadScreen rcvd_date, 8, 12, 20
rcvd_date = replace(rcvd_date, " ", "/")
EMReadScreen gender_ask, 1, 9, 20
EMReadScreen MLAR_addr_street, 19, 8, 56
EMReadScreen MLAR_addr_streetII, 19, 9, 56
EMReadScreen MLAR_addr_city, 22, 12, 56
EMReadScreen MLAR_addr_state, 2, 13, 56
EMReadScreen MLAR_addr_zip, 5, 13, 65
EMReadScreen addr_county, 22, 14, 56
EMReadScreen MLAR_addr_phone, 12, 15, 56
'EMReadScreen appl_status, 2, 4, 20 'this is not used anywhere else in the script'
'--------------------------------------------------------------------navigates to PERS and writing SSN'
PF2
EMwritescreen SSN_first, 14, 36
EMwritescreen SSN_mid, 14, 40
EMwritescreen SSN_last, 14, 43
TRANSMIT
'msgbox " appl: " & appl_date & " rcvd: " & rcvd_date
EMReadScreen error_check, 18, 24, 2
error_check = trim(error_check)
IF error_check = "SSN DOES NOT EXIST" THEN script_end_procedure_with_error_report ("Unable to find person in SSN search." & vbNewLine & "Please do a PERS search using the client's name." & vbNewLine & "Case may need to be APPLd.")
'This will take us to certain places based on PERS search'
EMReadscreen current_panel_check, 4, 2, 51
IF current_panel_check = "PERS" THEN script_end_procedure_with_error_report ("Please search by person name and run script again.")

Row = 8
IF current_panel_check = "MTCH" THEN
	DO
		EMReadScreen PMI_number, 7, row, 71
		IF trim(PMI_number) = "" THEN script_end_procedure_with_error_report("A PMI could not be found. The script will now end.")
		PERS_check = MsgBox("Multiple matches found. Ensure duplicate PMIs have been reported, APPL using oldest PMI." & vbNewLine & "Press YES to confirm this is the PERS match you wish to act on." & vbNewLine & "For the next PERS match, press NO." & vbNewLine & vbNewLine & _
		"   " & PMI_number, vbYesNoCancel, "Please confirm this PERS match")
		If PERS_check = vbYes THEN
			EMWriteScreen "x", row, 5
			TRANSMIT
			EMReadscreen current_panel_check, 4, 2, 51
			IF current_panel_check = "DSPL" THEN
			 	EXIT DO
			ELSE
				msgbox("Unable to access DSPL screen. Please review your case, and process manually if necessary.")
				EXIT DO
			END IF
		END IF
		IF PERS_check = vbNo THEN
			row = row + 1
			'msgbox "row: " & row
			IF row = 16 THEN
				PF8
				row = 8
			END IF
		END IF
		IF PERS_check = vbCancel THEN script_end_procedure_with_error_report ("The script has ended. The match has not been acted on.")
	LOOP UNTIL PERS_check = vbYes
'ELSE
END IF
'msgbox "Where am I this should be DSPL if there is a match"
IF current_panel_check = "DSPL" THEN
	EMwritescreen "HC", 07, 22 'drilling down for accuracy '
	TRANSMIT
	EMReadScreen error_msg, 23, 24, 2
	error_msg = trim(error_msg)
	IF error_msg = "NO RECORDS EXIST FOR HC" THEN
		EMwritescreen "MA", 07, 22
		TRANSMIT
	END IF
	IF MAXIS_case_number = "" THEN
		EMwritescreen "  ", 07, 22
		TRANSMIT
		EMReadScreen MAXIS_case_number, 8, row, 06 'not sure about this part'
		EMReadscreen case_status, 4, row, 35
		EMReadScreen case_status, 4, row, 53
	END IF
    '-------------------------------------------------------------------checking for an active case
	row = 10
    DO
		EMReadScreen MAXIS_case_number, 8, row, 06
    	'second loop to ensure we are acting on the correct case number'
		EMReadScreen primary_appl, 1, 10, 61
		IF primary_appl = "Y" THEN
		    MLAR_case_number_check = MsgBox("Client is the primary applicant on multiple case matches. Ensure duplicate PMIs have been reported, APPL or update using current or pending case." & vbNewLine & "Press YES to confirm this is the case you wish to act on." & vbNewLine & "For the next case, press NO." & vbNewLine & vbNewLine & _
    	    "   " & MAXIS_case_number, vbYesNoCancel, "Please confirm this case.")
    	    IF MLAR_case_number_check = vbYes THEN
    	    	EMWriteScreen "x", row, 5
    	    	TRANSMIT
		    END IF
    	    If MLAR_case_number_check = vbNo THEN
    	    	row = row + 1
    	    	'msgbox "row: " & row
		    	IF row = 19 THEN
		    		PF8
		    		row = 7
		    	END IF
    	    END IF
		    IF MLAR_case_number_check = vbCancel THEN script_end_procedure_with_error_report ("The script has ended. The case has not been acted on.")
		ELSE
			MLAR_case_number_check = MsgBox("The client is known to MAXIS but is not the primary applicant, please review to ensure case accurancy. Ensure duplicate PMIs have been reported, APPL or update using current or pending case." & vbNewLine & "Press YES to confirm this is the case you wish to act on." & vbNewLine & "For the next case, press NO." & vbNewLine & vbNewLine & _
			"   " & MAXIS_case_number, vbYesNoCancel, "Please confirm this case.")
			IF MLAR_case_number_check = vbYes THEN
				EMWriteScreen "x", row, 5
				TRANSMIT
				END IF
			IF MLAR_case_number_check = vbNo THEN
				row = row + 1
				'msgbox "row: " & row
				IF row = 19 THEN
					PF8
					row = 7
				END IF
			END IF
			IF MLAR_case_number_check = vbCancel THEN script_end_procedure_with_error_report ("The script has ended. The case has not been acted on.")
		END IF
	LOOP UNTIL MLAR_case_number_check = vbYes

    IF case_status = "CURRENT" THEN
    	EMReadScreen appl_date, 8, row, 25
    	APPL_box = MsgBox("This information is read from REPT/MLAR:" & vbcr & MLAD_maxis_name & vbcr & appl_date & vbcr & maxis_name & vbcr & client_dob & vbcr & gender_ask & vbcr & MLAR_addr_street & MLAR_addr_street & MLAR_addr_city & MLAR_addr_state & "" & MLAR_addr_zip & vbcr & MLAR_addr_phone & vbcr & "APPL case and click OK if you wish to continue running the script and CANCEL if you want to exit." & vbcr & "HCRE must be updated when adding HC", vbOKCancel)
    	IF APPL_box = vbCancel then script_end_procedure_with_error_report("The script has ended. Please review the REPT/MLAR as you indicated that you wish to exit the script")
    ELSEIF case_status = "PEND" THEN
    	EMReadScreen pend_date, 5, row, 47
    	PEND_box = MsgBox("This information is read from REPT/MLAR:" & vbcr & MLAD_maxis_name & vbcr & appl_date & vbcr & maxis_name & vbcr & client_dob & vbcr & gender_ask & vbcr & MLAR_addr_street & MLAR_addr_street & MLAR_addr_city & MLAR_addr_state & "" & MLAR_addr_zip & vbcr & MLAR_addr_phone & vbcr & "APPL case and click OK if you wish to continue running the script and CANCEL if you want to exit." & vbcr & "HCRE must be updated when adding HC", vbOKCancel)
    	IF PEND_box = vbCancel then script_end_procedure_with_error_report("The script has ended. Please review the REPT/MLAR as you indicated that you wish to exit the script")
    ELSEIF case_status = "CAF " THEN
    	MsgBox "Please ensure case is in a PEND II status"
    	EMReadScreen end_date, 5, row, 53
    END IF
    'Call navigate_to_MAXIS_screen("CASE", "CURR")

    IF case_status = "CURRENT" or case_status = "PEND" THEN
		Call access_ADDR_panel(access_type, notes_on_address, addr_line_1, addr_line_2, resi_street_full, city, State, Zip_code, resi_county, addr_verif, addr_homeless, addr_reservation, addr_living_sit, reservation_name, mailing_addr_line_1, mailing_addr_line_2, mail_street_full, mailing_city, mailing_State, mailing_Zip_code, addr_eff_date, addr_future_date, phone_one, phone_two, phone_three, type_one, type_two, type_three, text_yn_one, text_yn_two, text_yn_three, addr_email, verif_received, original_information, update_attempted)
	END IF
	'IF current_panel_check <> "ADDR" THEN MsgBox(current_panel_check)
END IF
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 376, 180, "MIPPA"
  EditBox 55, 5, 35, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    PushButton 110, 5, 50, 15, "Geocoder", Geo_coder_button
  CheckBox 170, 30, 160, 10, "Check if case does not need to be transferred", transfer_case_checkbox
  EditBox 55, 25, 20, 15, spec_xfer_worker
  DropListBox 250, 5, 120, 15, "Select One:"+chr(9)+"YES - Update MLAD"+chr(9)+"NO - APPL (Known to MAXIS)"+chr(9)+"NO - APPL (Not known to MAXIS)"+chr(9)+"NO - ADD A PROGRAM", select_answer
  ButtonGroup ButtonPressed
    OkButton 275, 160, 45, 15
    CancelButton 325, 160, 45, 15
  Text 170, 10, 75, 10, "Active on Health Care?"
  Text 5, 30, 40, 10, "Transfer to:"
  Text 5, 10, 50, 10, "Case Number:"
  Text 80, 30, 60, 10, " (last 3 digit of X#)"
  Text 15, 60, 215, 10, "Case Name: "   & MLAD_maxis_name
  Text 15, 70, 110, 10, "APPL date: "  & appl_date
  Text 15, 100, 80, 10, "DOB: "   & client_dob
  Text 170, 140, 100, 10, "Phone: "   & MLAR_addr_phone
  Text 15, 115, 110, 10, "Received Date: "  & rcvd_date
  Text 15, 135, 120, 10, "MEDI Number: "  &  medi_number
  Text 15, 90, 120, 10, "SSN: "  & MLAD_SSN_number
  Text 15, 125, 110, 10, "Gender Marker: "   & gender_ask
  Text 170, 90, 195, 10, "Addr: "   & MLAR_addr_streetII & MLAR_addr_street
  Text 170, 110, 90, 10, "State: "  & MLAR_addr_state
  Text 170, 100, 195, 10, "City: "  &  MLAR_addr_city
  Text 170, 120, 110, 10, "Zip: " & MLAR_addr_zip
  Text 170, 130, 110, 10, "County: " & addr_county
  GroupBox 5, 45, 365, 110, "MLAR Information"
EndDialog

'--------------------------------------------------------------------------------------------------script
DO 'Password DO loop
	DO 'Conditional handling DO loop
       DO  'External resource DO loop
           dialog Dialog1
           cancel_without_confirmation
           If ButtonPressed = Geo_coder_button then CreateObject("WScript.Shell").Run("https://hcgis.hennepin.us/agsinteractivegeocoder/default.aspx")
       Loop until ButtonPressed = -1
	   err_msg = ""
	   IF select_answer = "Select One:" THEN err_msg = err_msg & vbnewline & "* Select at least one option."
	   IF transfer_case_checkbox = CHECKED and spec_xfer_worker <> "" THEN err_msg = err_msg & vbnewline & "* Only check if the case does NOT need to be transferred."
	   IF transfer_case_checkbox = UNCHECKED and spec_xfer_worker = "" THEN err_msg = err_msg & vbnewline & "* You must advise of basket to transfer to (last 3 digits of worker number)."
	   IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)
LOOP UNTIL are_we_passworded_out = FALSE

'-------------------------------------------------------------------------------------Transfers the case to the assigned worker if this was selected in the second dialog box
'Determining if a case will be transferred or not. All cases will be transferred except addendum app types. THIS IS NOT CORRECT AND NEEDS TO BE DISCUSSED WITH QI
IF spec_xfer_worker <> "" THEN
    CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
    EMWriteScreen "x", 7, 16
    TRANSMIT
    PF9
    EMWriteScreen "X127" & spec_xfer_worker, 18, 61
    TRANSMIT
    EMReadScreen worker_check, 9, 24, 2

    IF worker_check = "SERVICING" THEN
    	action_completed = False
    	PF10
    END IF

    EMReadScreen transfer_confirmation, 16, 24, 2
    IF transfer_confirmation = "CASE XFER'D FROM" then
    	action_completed = True
    Else
    	action_completed = False
    End if
END IF

denial_date = DateAdd("d", 45, date)

'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
IF select_answer = "YES - Update MLAD" or select_answer = "NO - ADD A PROGRAM" THEN
    Call create_TIKL("~ Client submitted intent to apply for MA/MSP on " & appl_date & " Certain Populations App mailed on " & date & " If client has not responded, HC request should be denied. If client is disabled, give an additional 15 days.", 0, denial_date, False, TIKL_note_text)
Else
    Call create_TIKL("~ Client submitted intent to apply for MA/MSP. Case is pending or active on HC. Ensure the Date of Application is " & appl_date & " or according to the HC app on file, whichever is oldest.", 0, denial_date, False, TIKL_note_text)
End if

'----------------------------------------------------------------------------------case note
start_a_blank_case_note
'Client submitted intent to apply for MA/MSP. Case is already pending or active on Health Care in MAXIS. Please ensure that the Date of Application is 1/22/2019 or according to the Health Care Application on file, whichever is oldest.
CALL write_variable_in_case_note("~ MIPPA/Extra Help request received via REPT/MLAR on " & rcvd_date & " ~")
'IF select_answer <> "YES - Update MLAD" THEN CALL write_variable_in_case_note("* Case APPL'd based on the intent to apply date.  Application mailed out by the Case Assignment team on " & date)
IF select_answer <> "YES - Update MLAD" THEN CALL write_variable_in_case_note("* Applicant is not active on Health Care in MAXIS.  Case APPLed based on the intent to apply date of " & appl_date & ". ")
IF select_answer = "YES - Update MLAD" THEN CALL write_variable_in_case_note("* Client submitted intent to apply for MA/MSP. Case is already pending or active on Health Care in MAXIS. Please ensure that the Date of Application is " & appl_date & " or according to the Health Care Application on file, whichever is oldest.")
IF select_answer = "NO - APPL (Known to MAXIS)" THEN CALL write_variable_in_case_note("* APPL'd case using the MIPPA record and case information applicant is known to MAXIS.")
IF select_answer = "NO - APPL (Not known to MAXIS)" THEN CALL write_variable_in_case_note("* APPL'd case using the MIPPA record and case information applicant is not known to MAXIS.")
IF select_answer = "NO - ADD A PROGRAM" THEN
	CALL write_variable_in_case_note("* APPL'd case using the MIPPA record and case information applicant is known to MAXIS and may be active on other programs.")
	IF end_date <> "" THEN CALL write_variable_in_case_note ("* HC Ended on: " & end_date)
END IF
CALL write_variable_in_case_note ("* REPT/MLAR APPL Date: " & appl_date)
IF select_answer <> "YES - Update MLAD" THEN
	CALL write_variable_in_case_note ("*  Application mailed out by the Case Assignment team & pended on: " & date)
	CALL write_variable_in_case_note("* The case should not be denied until: " & denial_date & " If client is disabled, please give an additional 15 days for the application to be returned.")
	CALL write_variable_in_case_note ("* TIKL set for " & denial_date & " to review for eligibility 45 days after the application was mailed out.")
END IF
IF spec_xfer_worker <> "" THEN CALL write_variable_in_case_note ("* Case transferred to basket " & spec_xfer_worker & ".")
CALL write_variable_in_case_note ("* MIPPA rcvd and acted on per: TE 02.07.459")
CALL write_variable_in_case_note ("---")
CALL write_variable_in_case_note (worker_signature)
PF3

'------------------------------------------------------------------------Naviagetes to REPT/MLAR'
'Navigates back to MIPPA to clear the match
CALL navigate_to_MAXIS_screen("REPT", "MLAR")
row = 11 'this part should be a for next?' can we jsut do a cursor read for now?
EMReadscreen msg_check, 1, row, 03
IF msg_check <> "_" THEN script_end_procedure_with_error_report("You are not on a MIPPA message. This script will stop.")
DO
	EMReadScreen MLAR_maxis_name, 21, row, 5
	MLAR_maxis_name = TRIM(MLAR_maxis_name)
	END_info_confirmation = MsgBox("Press YES to confirm this is the MIPPA  you wish to clear." & vbNewLine & "For the next match, press NO." & vbNewLine & vbNewLine & _
	"   " & MLAR_maxis_name, vbYesNoCancel, "Please confirm this match")
	IF END_info_confirmation = vbNo THEN
		row = row + 1
		'msgbox "row: " & row
		IF row = 19 THEN
			PF8
			row = 7
		END IF
	END IF
	IF END_info_confirmation = vbCancel THEN script_end_procedure_with_error_report ("The script has ended. The match has not been acted on.")
	IF END_info_confirmation = vbYes THEN
		EMwritescreen "X", row, 03
		TRANSMIT
		PF9
		'this is the updated for MLAD'
		IF select_answer = "YES - Update MLAD" or select_answer = "NO - ADD A PROGRAM" THEN
			EMwritescreen "AP", 4, 20
		ELSE
			EMwritescreen "PN", 4, 20
		END IF
		TRANSMIT
		PF3
		PF3
	END IF
LOOP UNTIL END_info_confirmation = vbYes
'to help check on app rcvd'
'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
'CALL create_outlook_email("pahoua.vang@hennepin.us;", "", maxis_name & maxis_case_number & " MIPPA case need Application sent EOM.", "", "", TRUE)
'msgbox "where am i ending?"
script_end_procedure("MIPPA CASE NOTE HAS BEEN UPDATED. PLEASE ENSURE THE CASE IS CLEARED on REPT/MLAR & THE FORMS HAVE BEEN MAILED. ")
