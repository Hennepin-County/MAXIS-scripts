'===========================================================================================STATS
name_of_script = "CA - MIPPA.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 200
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("10/10/2017", "Updates to correct dialog box error message and ensure the correct case number pulls through the whole script/", "MiKayla Handley, Hennepin County")
call changelog_update("10/10/2017", "Updates to correct action when case noting and updating REPT/MLAR", "MiKayla Handley, Hennepin County")
call changelog_update("09/29/2017", "Updates to correct action if HC is already pending", "MiKayla Handley, Hennepin County")
call changelog_update("08/21/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK 

'---------------------------------------------------------------script
'option explicit ask Ilse 
EMConnect ""

'Navigates to MIPPA Lis Application-Medicare Improvement for Patients and Providers (MIPPA) 
CALL navigate_to_MAXIS_screen("REPT", "MLAR")

MAXIS_row = 11 'this part should be a for next?' can we jsut do a cursor read for now?
EMReadscreen msg_check, 1, MAXIS_row, 03
IF msg_check <> "_" THEN script_end_procedure("You are not on a MIPPA message. This script will stop.")
EMReadScreen maxis_name, 22, row, 5
maxis_name = TRIM(maxis_name)

EMwritescreen "X", MAXIS_row, 03 'this will take us to REPT/MLAD'
transmit
	'navigates to MLAD
EMReadScreen MLAD_maxis_name, 22, 6, 20
	MLAD_maxis_name = TRIM(MLAD_maxis_name)
EMReadScreen SSN_first, 3, 7, 20
EMReadScreen SSN_mid, 2, 7, 24
EMReadScreen SSN_last, 4, 7, 27
EMReadScreen appl_date, 8, 11, 20
	appl_date = replace(appl_date, " ", "/")
'----------------------------------------------------------------------used for the dialog to appl
EMReadScreen birth_date, 8, 8, 20
EMReadScreen medi_number, 10, 10, 20 
EMReadScreen rcvd_date, 8, 12, 20
EMReadScreen gender_ask, 1, 9, 20
EMReadScreen addr_street, 19, 9, 56
EMReadScreen apt_addr, 19, 8, 56
EMReadScreen addr_city, 22, 12, 56
EMReadScreen addr_state, 2, 13, 56
EMReadScreen addr_zip, 5, 13, 65
EMReadScreen addr_county, 22, 14, 56
EMReadScreen addr_phone, 12, 15, 56
EMReadScreen appl_status, 2, 4, 20
'--------------------------------------------------------------------navigates to PERS'
PF2 
EMwritescreen SSN_first, 14, 36
EMwritescreen SSN_mid, 14, 40
EMwritescreen SSN_last, 14, 43 

transmit

EMReadScreen error_msg, 18, 24, 2
error_msg = trim(error_msg)
IF error_msg = "SSN DOES NOT EXIST" THEN script_end_procedure ("Unable to find person in SSN search." & vbNewLine & "Please do a PERS search using the client's name." & vbNewLine & "Case may need to be APPLd.")

EMReadscreen PAGE_confirmation, 4, 2, 51 
IF PAGE_confirmation = "PERS" THEN script_end_procedure ("Please search by person name and run script again.")
IF PAGE_confirmation = "MTCH" THEN script_end_procedure ("PMI NBR ASSIGNED THRU SMI OR PMIN - NO MAXIS CASE EXISTS - Ensure duplicate PMIs have been reported if found, APPL using oldest PMI")
IF PAGE_confirmation <> "DSPL" THEN script_end_procedure("Unable to access DSPL screen. Please review your case, and process manually if necessary.")

EMwritescreen "HC", 07, 22 
transmit
EMReadScreen error_msg, 23, 24, 2
error_msg = trim(error_msg)
IF error_msg = "NO RECORDS EXIST FOR HC" THEN 
	EMwritescreen "MA", 07, 22
	transmit
END IF 
'-------------------------------------------------------------------checking for an active case
MAXIS_row = 10 
EMReadScreen MAXIS_case_number, 8, MAXIS_row, 06 
MAXIS_case_number = trim(MAXIS_case_number)
EMReadscreen current_case, 7, MAXIS_row, 35 
EMReadScreen pending_case, 4, MAXIS_row, 53
IF MAXIS_case_number = "" THEN 
	EMwritescreen "AP", 07, 22
	transmit
	EMReadScreen MAXIS_case_number, 8, MAXIS_row, 06 
	EMReadscreen current_case, 7, MAXIS_row, 35 
	EMReadScreen pending_case, 4, MAXIS_row, 53
END IF 

IF current_case = "Current" THEN 
	EMReadScreen appl_date, 8, MAXIS_row, 25
	transfer_case = FALSE
ELSEIF pending_case = "PEND" THEN 
	EMReadScreen pend_date, 5, MAXIS_row, 47
ELSEIF pending_case = "CAF " THEN 
	MsgBox "Please ensure case is in a PEND II status" 	
	EMReadScreen end_date, 5, MAXIS_row, 53
END IF	

'------------------------------------------------------------------------------------------------dialogs'
BeginDialog MIPPA_active_dialog, 0, 0, 206, 70, "MIPAA"
  EditBox 55, 5, 55, 15, MAXIS_case_number
  CheckBox 115, 10, 85, 10, "Check to transfer case", transfer_case_checkbox
  DropListBox 85, 30, 115, 15, "Select One..."+chr(9)+"YES - Update MLAD"+chr(9)+"NO - APPL (Known to MAXIS)"+chr(9)+"NO - ADD A PROGRAM"+chr(9)+"NO - APPL (Not known to MAXIS)", select_answer
  ButtonGroup ButtonPressed
    OkButton 105, 50, 45, 15
    CancelButton 155, 50, 45, 15
  Text 5, 10, 50, 10, "Case Number: "
  Text 5, 35, 75, 10, "Active on Health Care?"
EndDialog

'----------------------------------------------------------------------------------------TRANSFER DIALOG'
BeginDialog transfer_dialog, 0, 0, 141, 110, "MIPAA Transfer"
  ButtonGroup ButtonPressed
    PushButton 20, 25, 100, 15, "Geocoder", Geo_coder_button
  Text 5, 10, 50, 10, "Case Number:"
  EditBox 55, 5, 65, 15, maxis_case_number
  Text 5, 50, 100, 10, "Transfer to (last 3 digit of X#):"
  EditBox 105, 45, 30, 15, spec_xfer_worker
  Text 5, 70, 90, 10, "Assigned to (3 digit team #):"
  EditBox 105, 65, 30, 15, team_number
  ButtonGroup ButtonPressed
    OkButton 40, 90, 45, 15
    CancelButton 90, 90, 45, 15
EndDialog

'----------------------------------------------------------------------------------------Case number DIALOG'
BeginDialog case_number_dialog, 0, 0, 116, 75, "MIPPA case note"
  EditBox 55, 5, 55, 15, MAXIS_case_number
  ButtonGroup ButtonPressed
    OkButton 15, 55, 45, 15
    CancelButton 65, 55, 45, 15
  Text 5, 10, 50, 10, "Case Number: "
  Text 5, 25, 110, 25, "To ensure accuracy please confirm that case # is correct for recent APPL or updates to MLAR"
EndDialog

'--------------------------------------------------------------------------------------------------script	
Do
	Do
		err_msg = ""
		dialog MIPPA_active_dialog
		IF select_answer = "Select One..." THEN err_msg = err_msg & vbnewline & "* Select at least one option."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)  
LOOP UNTIL are_we_passworded_out = false

IF select_answer <> "YES - Update MLAD" THEN APPL_box = MsgBox("This information is read from REPT/MLAR:" & vbcr & MLAD_maxis_name & vbcr & appl_date & vbcr & maxis_name & vbcr & SSN_first & SSN_mid & SSN_last & vbcr & birth_date & vbcr & gender_ask & vbcr & addr_street & apt_addr & addr_city & addr_state & "" & addr_zip & vbcr & addr_phone & vbcr & "APPL case and click OK if you wish to continue running the script and CANCEL if you want to exit." & vbcr & "HCRE must be updated when adding HC", vbOKCancel)
IF APPL_box = vbCancel then script_end_procedure("The script has ended. Please review the REPT/MLAR as you indicated that you wish to exit the script")

Call MAXIS_case_number_finder(MAXIS_case_number)
'-------------------------------------------------------------------------------------Transfers the case to the assigned worker if this was selected in the second dialog box
'Determining if a case will be transferred or not. All cases will be transferred except addendum app types. THIS IS NOT CORRECT AND NEEDS TO BE DISCUSSED WITH QI
IF transfer_case_checkbox = UNCHECKED THEN 		
	transfer_case = FALSE
ELSE 
	transfer_case = TRUE
END IF 

IF transfer_case = TRUE THEN 
    DO
    	DO
    	   err_msg = ""
    		DO
    			dialog transfer_dialog
    			cancel_confirmation
    			IF buttonpressed = Geo_coder_button THEN CreateObject("WScript.Shell").Run("https://hcgis.hennepin.us/agsinteractivegeocoder/default.aspx")
    		LOOP UNTIL buttonpressed = -1	
            IF spec_xfer_worker = "" then err_msg = err_msg & vbCr & "You must have a caseload # (SPEC/XFER) to continue."
            IF len(spec_xfer_worker) <> 3 then err_msg = err_msg & vbCr & "You only need to include last 3 digit of X127#"
            IF team_number = "" then err_msg = err_msg & vbCr & "You must have a 3 digit team # (email) to continue."
            IF err_msg <> "" THEN Msgbox err_msg
    	LOOP UNTIL err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	LOOP UNTIL are_we_passworded_out = false
	
    	
    CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
    EMWriteScreen "x", 7, 16
    transmit
    PF9
    EMWriteScreen "X127" & spec_xfer_worker, 18, 61
    transmit
    EMReadScreen worker_check, 9, 24, 2
    IF worker_check = "SERVICING" THEN
    	MsgBox "The correct worker number was not entered, this X-Number is not a valid worker in MAXIS. You will need to transfer the case manually"
		PF10
    	transfer_case = unchecked
    END IF
END IF
	
CALL check_for_password(are_we_passworded_out)
'--------------------------------------------------------------------------FOR each case we have to come back to clear/update the MLAD screen

MAXIS_background_check

'------------------------------------------------------------------------Naviagetes to REPT/MLAR'

CALL navigate_to_MAXIS_screen("REPT", "MLAR")

MAXIS_row = 11
	DO
		EMReadScreen MLAD_maxis_name, 25, row, 5
		MLAD_maxis_name = trim(MLAD_maxis_name)
		IF MLAD_maxis_name = maxis_name THEN 
			EXIT DO
		ELSE
			row = row + 1
			IF row = 17 THEN 
				PF8
				ROW = 11
			END IF 
		END IF
	LOOP UNTIL case_number = ""
	

EMwritescreen "X", MAXIS_row, 03
transmit
PF9	
 
IF select_answer = "YES - Update MLAD" or select_answer = "NO - ADD A PROGRAM" THEN 
	EMwritescreen "AP", 4, 20
	transmit
	PF3
	PF3
    EMWriteScreen MAXIS_case_number, 18, 43
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	CALL create_MAXIS_friendly_date(date, 0, 5, 18)
	CALL write_variable_in_TIKL("~ A MIPPA record was recieved please check case information for consistency and follow-up with any inconsistent information, as appropriate.")
	transmit
	PF3
ELSE 
	EMwritescreen "PN", 4, 20	
	transmit
	PF3
	PF3
    EMWriteScreen MAXIS_case_number, 18, 43
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	CALL create_MAXIS_friendly_date(date, 0, 5, 18)
	CALL write_variable_in_TIKL("~ Please review the MIPPA record and case information for consistency and follow-up with any inconsistent information, as appropriate.")
	transmit
	PF3			
END IF 
 

'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
'CALL create_outlook_email("mikayla.handley@hennepin.us;", "", maxis_name & maxis_case_number & " MIPPA case need Application sent EOM.", "", "", TRUE)	
'-----------------------------------------------------------------------initial case number dialog
Do 
	DO 
		err_msg = ""
	    dialog case_number_dialog
        if ButtonPressed = 0 Then StopScript
        if IsNumeric(maxis_case_number) = false or len(maxis_case_number) > 8 THEN err_msg = err_msg & vbNewLine & "* Please enter a valid case number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine		
	Loop until err_msg = ""	
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

back_to_self

EMWriteScreen MAXIS_case_number, 18, 43
'----------------------------------------------------------------------------------case note

start_a_blank_CASE_NOTE 
IF select_answer = "YES - Update MLAD" THEN 
	CALL write_variable_in_CASE_NOTE("~ MIPAA received via REPT/MLAR on " & rcvd_date & " ~")
	CALL write_variable_in_CASE_NOTE("** Please review the MIPPA record and case information for consistency and follow-up with any inconsistent information, as appropriate.")
ELSE CALL write_variable_in_CASE_NOTE ("~ HC PENDED - MIPAA received via REPT/MLAR on " & appl_date & " ~")
    IF select_answer = "NO - APPL (Known to MAXIS)" THEN 
    	CALL write_variable_in_CASE_NOTE("** APPL'd case using the MIPPA record and case information applicant is   known to MAXIS by SSN or name search.")
    	CALL write_variable_in_CASE_NOTE ("* Pended on: " & date)
		CALL write_variable_in_CASE_NOTE ("* Application mailed using automated system per DHS: " & rcvd_date)
    ELSEIF select_answer = "NO - APPL (Not known to MAXIS)" THEN 
    	CALL write_variable_in_CASE_NOTE("** APPL'd case using the MIPPA record and case information applicant is not     known to MAXIS by SSN or name search.")
    	CALL write_variable_in_CASE_NOTE ("* Pended on: " & date)
		CALL write_variable_in_CASE_NOTE ("* Application mailed using automated system per DHS: " & rcvd_date)
    ELSEIF select_answer = "NO - ADD A PROGRAM" THEN 
    	CALL write_variable_in_CASE_NOTE("** APPL'd case using the MIPPA record and case information applicant is      known to MAXIS and may be active on other programs.")
		CALL write_variable_in_CASE_NOTE ("* Application mailed using automated system per DHS: " & rcvd_date)
    	CALL write_variable_in_CASE_NOTE ("* HC Ended on: " & end_date)
	END IF
END IF	
CALL write_variable_in_case_NOTE ("* Requesting: HC")
CALL write_variable_in_CASE_NOTE ("* REPT/MLAR APPL Date: " & appl_date)
IF transfer_case = TRUE THEN CALL write_variable_in_CASE_NOTE ("* Case transferred to Team " & team_number & " in MAXIS(" & spec_xfer_worker & ").")
CALL write_variable_in_CASE_NOTE ("* MIPPA rcvd and acted on per: TE 02.07.459")
CALL write_variable_in_CASE_NOTE ("---")
CALL write_variable_in_CASE_NOTE (worker_signature)

script_end_procedure("MIPPA CASE NOTE HAS BEEN UPDATED. PLEASE ENSURE THE CASE IS CLEARED on REPT/MLAR.")
