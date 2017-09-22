'===========================================================================================STATS
name_of_script = "CA-MIPPA.vbs"
start_time = timer
STATS_counter = 1
STATS_manualtime = 200
STATS_denominatinon = "C"
'===========================================================================================END OF STATS BLOCK
'===========================================================================LOADING FUNCTIONS LIBRARY FROM GITHUB 
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib IF it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'IF the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'IF the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End IF
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
'================================================================================================END FUNCTIONS LIBRARY BLOCK
' ===========================================================================================================CHANGELOG BLOCK
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("08/21/2017", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'=======================================================================================================END CHANGELOG BLOCK 

'---------------------------------------------------------------script
'option explicit ask Ilse 
EMConnect ""
CALL check_for_MAXIS(FALSE)

'Navigates to MIPPA Lis Application-Medicare Improvement for Patients and Providers (MIPPA) 
CALL navigate_to_MAXIS_screen("REPT", "MLAR")

MAXIS_row = 11 'this part should be a for next?' can we jsut do a cursor read for now?
EMReadscreen msg_check, 1, MAXIS_row, 03
If msg_check <> "_" then script_end_procedure("You are not on a MIPPA message. This script will stop.")

EMwritescreen "X", MAXIS_row, 03 'this will take us to REPT/MLAD'
transmit
	'navigates to MLAD
EMReadScreen maxis_name, 25, 6, 20
EMReadScreen SSN_first, 3, 7, 20
EMReadScreen SSN_mid, 2, 7, 24
EMReadScreen SSN_last, 4, 7, 27
EMReadScreen appl_date, 8, 11, 20

'used for the dialog to appl
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

appl_date = replace(appl_date, " ", "/")
PF2 'navigates to PERS'


EMwritescreen SSN_first, 14, 36
EMwritescreen SSN_mid, 14, 40
EMwritescreen SSN_last, 14, 43 

transmit

EMReadScreen error_msg, 18, 24, 2
error_msg = trim(error_msg)
IF error_msg = "SSN DOES NOT EXIST" THEN script_end_procedure ("Unable to find person in SSN search." & vbNewLine & "  Please do a PERS search using the client's name"  & vbNewLine & "     Case may need to be APPLd.")

EMReadscreen PAGE_confirmation, 4, 2, 51 
IF PAGE_confirmation = "PERS" THEN script_end_procedure ("Please search by person name and run script again.")
IF PAGE_confirmation = "MTCH" THEN script_end_procedure ("PMI NBR ASSIGNED - Ensure duplicate PMIs have been reported if found, APPL using oldest PMI")
IF PAGE_confirmation <> "DSPL" THEN script_end_procedure("Unable to access DSPL screen. Please review your case, and process manually if necessary.")

EMwritescreen "HC", 07, 22 
transmit
EMReadScreen error_msg, 23, 24, 2
error_msg = trim(error_msg)
IF error_msg = "NO RECORDS EXIST FOR HC" THEN 
	EMwritescreen "MA", 07, 22
	transmit
END IF 

MAXIS_row = 10 'checking for an active case
EMReadScreen MAXIS_case_number, 8, MAXIS_row, 06 
EMReadscreen current_case, 7, MAXIS_row, 35 'need a or read for pend here
EMReadScreen pending_case, 4, MAXIS_row, 53
IF trim(MAXIS_case_number) = "" THEN 
	EMwritescreen "AP", 07, 22
	transmit
	EMReadScreen MAXIS_case_number, 8, MAXIS_row, 06 
	EMReadscreen current_case, 7, MAXIS_row, 35 'need a or read for pend here
	EMReadScreen pending_case, 4, MAXIS_row, 53
END IF 
'checking for an active case

'script_end_procedure ("Please search by person name and if no case can be found - APPL case - then run script again and select NO-APPL(not known)")

IF current_case = "Current" THEN EMReadScreen appl_date, 8, MAXIS_row, 25
IF pending_case = "PEND" THEN EMReadScreen pend_date, 5, MAXIS_row, 47
IF pending_case = "CAF " THEN 
	MsgBox "Please ensure case is in a PEND II status" 	
	EMReadScreen end_date, 5, MAXIS_row, 53
END IF	

'------------------------------------------------------------------------------------------------dialog'
BeginDialog MIPPA_active_dialog, 0, 0, 186, 65, "MIPAA"
  EditBox 85, 5, 55, 15, MAXIS_case_number
  DropListBox 85, 25, 95, 15, "Select One..."+chr(9)+"YES - Update MLAD"+chr(9)+"NO - APPL (Known)"+chr(9)+"NO - APPL (Not known)"+chr(9)+"NO - ADD A PROGRAM", select_answer
  ButtonGroup ButtonPressed
    OkButton 75, 45, 50, 15
    CancelButton 130, 45, 50, 15
  Text 5, 30, 75, 10, "Active on Health Care?"
  Text 35, 10, 50, 10, "Case Number: "
EndDialog

BeginDialog transfer_dialog, 0, 0, 111, 155, "MIPAA Transfer"
  	ButtonGroup ButtonPressed
	   PushButton 5, 5, 100, 15, "Geocoder", Geo_coder_button
  	EditBox 55, 25, 50, 15, team_region
  	EditBox 55, 45, 50, 15, population_team
  	EditBox 55, 65, 50, 15, worker_to_transfer_to
  	EditBox 55, 85, 50, 15, worker_team_number
  	ButtonGroup ButtonPressed
	   OkButton 10, 135, 45, 15
       CancelButton 60, 135, 45, 15
  	Text 15, 110, 100, 20, "* Script will transfer case to                assigned worker *"
  	Text 15, 50, 40, 10, "Population:"
  	Text 25, 30, 25, 10, "Region:"
  	Text 5, 90, 50, 10, "Team Number:"
  	Text 10, 70, 40, 10, "Assigned to:"
EndDialog

'--------------------------------------------------------------------------------------------------script	
Do
	Do
		err_msg = ""
		dialog MIPPA_active_dialog
		cancel_confirmation
		IF select_answer = "Select One..." THEN err_msg = err_msg & vbnewline & "* Select at least one option."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)  
LOOP UNTIL are_we_passworded_out = false

IF select_answer = "NO - APPL (Known)" or select_answer = "NO - APPL (Not known)" THEN APPL_box = MsgBox("This information is read from REPT/MLAR:" & vbcr & appl_date & vbcr & maxis_name & vbcr & SSN_first & SSN_mid & SSN_last & vbcr & birth_date & vbcr & gender_ask & vbcr & addr_street & apt_addr & addr_city & addr_state & "" & addr_zip & vbcr & addr_phone & vbcr & "APPL case and click OK if you wish to continue running the script and CANCEL if you want to exit.",  vbOKCancel)
	If APPL_box = vbCancel then script_end_procedure("The script has ended. Please review the REPT/MLAR as you indicated that you wish to exit the script")

DO
	DO
		err_msg = ""
		DO
			dialog transfer_dialog
			cancel_confirmation
			IF buttonpressed = Geo_coder_button THEN CreateObject("WScript.Shell").Run("https://hcgis.hennepin.us/agsinteractivegeocoder/default.aspx")
			IF team_region = "" then err_msg = err_msg & vbCr & "Please provide team region (CNE, SS, NW) to continue."
			IF population_team = "" then err_msg = err_msg & vbCr & "Please provide population (ADAD, FAD) to continue."
			IF worker_team_number = "" then err_msg = err_msg & vbCr & "You must have a 3 digit team email to continue."
			IF len(worker_to_transfer_to) <> 3 then err_msg = err_msg & vbCr & "You only need to include last 3 digit of X127#"
			IF err_msg <> "" THEN Msgbox err_msg
		LOOP UNTIL buttonpressed = -1	
	LOOP UNTIL err_msg = ""	
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false	

	'-------------------------------------------------------------------------------------Transfers the case to the assigned worker if this was selected in the second dialog box
 
 IF select_answer = "NO - ADD A PROGRAM" THEN 
	MsgBox "The script will help you add this program to current maxis case."
	transfer_case = False 
Elseif select_answer = "YES - Update MLAD" THEN 
	MsgBox "Please update the appropriate status on REPT/MLAR."
	transfer_case = False
else 
	transfer_case = True
End if 

If transfer_case = true then 	
	CALL navigate_to_MAXIS_screen ("SPEC", "XFER")
	EMWriteScreen "x", 7, 16
	transmit
	PF9
	EMWriteScreen "X127" & worker_to_transfer_to, 18, 61
	transmit
	EMReadScreen worker_check, 9, 24, 2
				
	IF worker_check = "SERVICING" THEN
		MsgBox "The correct worker number was not entered, this X-Number is not a valid worker in MAXIS. You will need to transfer the case manually"
		PF10
	END IF
End if 	

'FOR each case we have to come back
CALL navigate_to_MAXIS_screen("REPT", "MLAR")
'For each case_new in MLAR_case

MAXIS_row = 11
 'this part should be a for next?' can we jsut do a cursor read for now?
EMwritescreen "X", MAXIS_row, 03 'this will take us to REPT/MLAD'
transmit
PF9	

IF select_answer = "YES - Update MLAD" THEN 
	EMwritescreen "AP", 4, 20
	transmit
	CALL navigate_to_MAXIS_screen("DAIL", "WRIT")
	CALL create_MAXIS_friendly_date(date, 0, 5, 18)
	CALL write_variable_in_TIKL("~** Please review the MIPPA record and case information for consistency and follow-up with any inconsistent information, as appropriate.")
	transmit
	PF3
Else
	EMwritescreen "PN", 4, 20
END IF 

'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)

CALL create_outlook_email("", "mikayla.handley@hennepin.us;", maxis_name & maxis_case_number & " MIPPA case need Application sent EOM.", "", "", FALSE)	
	
'---------------------------------------------------------------------------------------------------------------------------------case note
start_a_blank_CASE_NOTE 'appl_date needs to be fixed to read 08/08/08 instead of 08 08 08'
CALL write_variable_in_CASE_NOTE ("~ HC PENDED - MIPAA received via REPT/MLAR on " & appl_date & " ~")
IF select_answer = "NO - APPL (Known)" THEN 
	CALL write_variable_in_CASE_NOTE("** APPL'd case using the MIPPA record and case information applicant is   known to MAXIS by SSN or name search.")
	CALL write_variable_in_CASE_NOTE ("* Pended on: " & date)
ELSEIF select_answer = "NO - APPL (Not known)" THEN 
	CALL write_variable_in_CASE_NOTE("** APPL'd case using the MIPPA record and case information applicant is not     known to MAXIS by SSN or name search.")
	CALL write_variable_in_CASE_NOTE ("* Pended on: " & date)
ELSEIF select_answer = "NO-ADD A PROGRAM" THEN 
	CALL write_variable_in_CASE_NOTE("** APPL'd case using the MIPPA record and case information applicant is      known to MAXIS and may be active on other programs.")
	CALL write_variable_in_CASE_NOTE ("* HC Ended on: " & end_date)
END IF	

IF select_answer = "YES - Update MLAD" THEN CALL write_variable_in_CASE_NOTE("** Please review the MIPPA record and case information for consistency and follow-up with any inconsistent information, as appropriate.")

CALL write_variable_in_case_NOTE ("* Requesting: HC")
CALL write_variable_in_CASE_NOTE ("* REPT/MLAR APPL Date: " & appl_date)
CALL write_variable_in_CASE_NOTE ("* Application mailed: " & date) 'this we do not want if team is mailing HCAPP enahance to send email possibly'
If transfer_case = true THEN CALL write_variable_in_CASE_NOTE ("* Case transferred to Team " & worker_to_transfer_to & " in MAXIS. ("  & worker_team_number & " " & team_region & " " & population_team & ")")
CALL write_variable_in_CASE_NOTE ("* MIPPA rcvd and acted on per: TE 02.07.459")
CALL write_variable_in_CASE_NOTE ("---")
CALL write_variable_in_CASE_NOTE (worker_signature)

script_end_procedure("MIPPA CASE NOTE HAS BEEN UPDATED. PLEASE ENSURE THE CASE IS CLEARED on REPT/MLAR." & vbcr & "MAXIS case number: " & MAXIS_case_number)