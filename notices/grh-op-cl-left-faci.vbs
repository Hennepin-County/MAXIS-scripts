'Required for statistical purposes==========================================================================================
name_of_script = "NOTICES - GRH OP CL LEFT FACI.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 150                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Checks for county info from global variables, or asks if it is not already defined.
get_county_code

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog GRH_OP_LEAVING_FACI_dialog, 0, 0, 326, 190, "GRH overpayment due to leaving facility dialog"
  EditBox 50, 5, 55, 15, MAXIS_case_number
  EditBox 165, 5, 45, 15, discovery_date
  EditBox 275, 5, 45, 15, established_date
  EditBox 90, 30, 230, 15, OP_reason
  EditBox 55, 55, 45, 15, OP_date_01
  EditBox 125, 55, 30, 15, OP_amt_01
  EditBox 205, 55, 45, 15, OP_date_02
  EditBox 275, 55, 45, 15, OP_amt_02
  EditBox 55, 75, 45, 15, OP_date_03
  EditBox 125, 75, 30, 15, OP_amt_03
  EditBox 205, 75, 45, 15, OP_date_04
  EditBox 275, 75, 45, 15, OP_amt_04
  EditBox 55, 95, 45, 15, OP_date_05
  EditBox 125, 95, 30, 15, OP_amt_05
  EditBox 205, 95, 45, 15, OP_date_06
  EditBox 275, 95, 45, 15, OP_amt_06
  CheckBox 10, 125, 95, 10, "Set follow-up 30 day TIKL ", set_TIKL_check
  EditBox 180, 120, 45, 15, OP_total
  ButtonGroup ButtonPressed
    PushButton 230, 120, 90, 15, "Calculate total facility OP", OP_total_button
  EditBox 240, 140, 80, 15, client_amt
  EditBox 120, 165, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 215, 165, 50, 15
    CancelButton 270, 165, 50, 15
  Text 10, 145, 230, 10, "Total client portion for GRH during all of the OP months (if applicable):"
  Text 165, 60, 40, 10, "Date of OP:"
  Text 255, 60, 15, 10, "Amt:"
  Text 10, 80, 40, 10, "Date of OP:"
  Text 105, 80, 15, 10, "Amt:"
  Text 5, 35, 85, 10, "Reason for overpayment:"
  Text 165, 80, 40, 10, "Date of OP:"
  Text 5, 10, 45, 10, "Case number:"
  Text 255, 80, 15, 10, "Amt:"
  Text 110, 10, 55, 10, "Discovery date:"
  Text 10, 100, 40, 10, "Date of OP:"
  Text 10, 60, 40, 10, "Date of OP:"
  Text 105, 100, 15, 10, "Amt:"
  Text 165, 100, 40, 10, "Date of OP:"
  Text 255, 100, 15, 10, "Amt:"
  Text 215, 10, 60, 10, "Established date:"
  Text 105, 60, 15, 10, "Amt:"
  Text 120, 125, 55, 10, "Facility OP total: "
  Text 55, 170, 60, 10, "Worker signature:"
EndDialog

BeginDialog GRH_OP_LEAVING_FACI_ADDR_dialog, 0, 0, 306, 220, "GRH overpayment due to leaving facility ADDR dialog"
  EditBox 70, 15, 230, 15, facility_name
  EditBox 70, 35, 230, 15, facility_address_line_01
  EditBox 70, 55, 230, 15, facility_address_line_02
  EditBox 70, 75, 80, 15, facility_city
  EditBox 155, 75, 25, 15, facility_state
  EditBox 185, 75, 45, 15, facility_zip
  EditBox 65, 115, 235, 15, county_name_dept
  EditBox 65, 135, 235, 15, county_address_line_01
  EditBox 65, 155, 235, 15, county_address_line_02
  EditBox 65, 175, 80, 15, county_address_city
  EditBox 150, 175, 25, 15, county_address_state
  EditBox 180, 175, 45, 15, county_address_zip
  CheckBox 5, 205, 165, 10, "Send overpayment to DHS (not a county/agency)", send_OP_to_DHS_check
  ButtonGroup ButtonPressed
    OkButton 195, 200, 50, 15
    CancelButton 250, 200, 50, 15
  Text 5, 160, 55, 10, "Address Line 2:"
  Text 5, 140, 55, 10, "Address Line 1:"
  Text 10, 180, 50, 10, "City/State/Zip:"
  Text 20, 80, 50, 10, "City/State/Zip:"
  Text 5, 60, 65, 10, "FACI ADDR Line 2:"
  Text 5, 40, 60, 10, "FACI ADDR line 1:"
  Text 25, 20, 40, 10, "FACI Name:"
  Text 5, 120, 45, 10, "Agency info:"
  GroupBox 0, 5, 305, 90, "**FACILITY ADDRESS WHERE THE OVERPAYMENT WILL BE SENT**"
  GroupBox 0, 105, 305, 90, "**AGENCY ADDRESS WHERE THE OVERPAYMENT WILL BE SENT**"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to MAXIS
EMConnect ""
'searches for a case number
call MAXIS_case_number_finder(MAXIS_case_number)

'Dialog completed by worker.  Worker must enter several mandatory fields, and will loop until worker presses cancel or completes fields.
DO
	DO
		DO
			DO
				DO
					DO
						DO
							DO
								Dialog GRH_OP_LEAVING_FACI_dialog
								cancel_confirmation
								If MAXIS_case_number = "" or isnumeric(MAXIS_case_number) = false THEN MsgBox "You did not enter a valid case number. Please try again."
								If worker_signature = "" THEN MsgBox "You did not sign your case note. Please try again."
								If discovery_date = "" THEN MsgBox "You must enter the discovery date"
								If established_date = "" THEN MsgBox "You must enter the established date"
								If OP_date_01 = "" THEN MsgBox "You must enter at least 1 overpayment date"
								If OP_amt_01 = "" THEN MsgBox "You must enter at least 1 overpayment amount"
								If OP_reason = "" THEN MsgBox "You must enter the reason for the overpayment."
							Loop until MAXIS_case_number <> "" and isnumeric(MAXIS_case_number) = true and worker_signature <> "" and discovery_date <> "" and established_date <> "" and OP_date_01 <> "" and OP_amt_01 <> "" and OP_reason <> ""
							If (OP_date_01 = "" AND OP_amt_01 <> "") OR (OP_date_01 <> "" AND OP_amt_01 = "") THEN MsgBox "You have must complete both an overpayment date AND an overpayment amount."
						LOOP UNTIL(OP_date_01 = "" AND OP_amt_01 = "") OR (OP_date_01 <> "" AND OP_amt_01 <> "")
						If (OP_date_02 = "" AND OP_amt_02 <> "") OR (OP_date_02 <> "" AND OP_amt_02 = "") THEN MsgBox "You have must complete both an overpayment date AND an overpayment amount."
					LOOP UNTIL(OP_date_02 = "" AND OP_amt_02 = "") OR (OP_date_02 <> "" AND OP_amt_02 <> "")
					If (OP_date_03 = "" AND OP_amt_03 <> "") OR (OP_date_03 <> "" AND OP_amt_03 = "") THEN MsgBox "You have must complete both an overpayment date AND an overpayment amount."
				LOOP UNTIL (OP_date_03 = "" AND OP_amt_03 = "") OR (OP_date_03 <> "" AND OP_amt_03 <> "")
				If (OP_date_04 = "" AND OP_amt_04 <> "") OR (OP_date_04 <> "" AND OP_amt_04 = "") THEN MsgBox "You have must complete both an overpayment date AND an overpayment amount."
			LOOP UNTIL (OP_date_04 = "" AND OP_amt_04 = "") OR (OP_date_04 <> "" AND OP_amt_04 <> "")
			If (OP_date_05 = "" AND OP_amt_05 <> "") OR (OP_date_05 <> "" AND OP_amt_05 = "") THEN MsgBox "You have must complete both an overpayment date AND an overpayment amount."
		LOOP UNTIL (OP_date_05 = "" AND OP_amt_05 = "") OR (OP_date_05 <> "" AND OP_amt_05 <> "")
		If (OP_date_06 = "" AND OP_amt_06 <> "") OR (OP_date_06 <> "" AND OP_amt_06 = "") THEN MsgBox "You have must complete both an overpayment date AND an overpayment amount."
	LOOP UNTIL (OP_date_06 = "" AND OP_amt_06 = "") OR (OP_date_06 <> "" AND OP_amt_06 <> "")
	If ButtonPressed = OP_total_button THEN
		'makes the overpayment amounts = 0 so the Abs(number) function work
		If OP_amt_01 = "" THEN OP_amt_01 = "0"
		If OP_amt_02 = "" THEN OP_amt_02 = "0"
		If OP_amt_03 = "" THEN OP_amt_03 = "0"
		If OP_amt_04 = "" THEN OP_amt_04 = "0"
		If OP_amt_05 = "" THEN OP_amt_05 = "0"
		If OP_amt_06 = "" THEN OP_amt_06 = "0"
		OP_total = (Abs(OP_amt_01) + Abs(OP_amt_02) + Abs(OP_amt_03) + Abs(OP_amt_04) + Abs(OP_amt_05) + Abs(OP_amt_06)) & ""
	END IF
		'This reverses this logic listed above (makes the overpayment amounts = 0 so the Abs(number) function work)
		If OP_amt_01 = "0" THEN OP_amt_01 = ""
		If OP_amt_02 = "0" THEN OP_amt_02 = ""
		If OP_amt_03 = "0" THEN OP_amt_03 = ""
		If OP_amt_04 = "0" THEN OP_amt_04 = ""
		If OP_amt_05 = "0" THEN OP_amt_05 = ""
		If OP_amt_06 = "0" THEN OP_amt_06 = ""
Loop until ButtonPressed = -1
	Do
		DO
			Dialog GRH_OP_LEAVING_FACI_ADDR_dialog
			cancel_confirmation
			If facility_name = "" THEN MsgBox "You must enter the facility name"
			If facility_address_line_01 = "" THEN MsgBox "You must enter the facility's street address"
			If facility_city = "" THEN MsgBox "You must enter the facility's city"
			If facility_state = "" THEN MsgBox "You must enter the facility's state"
			If facility_zip = "" THEN MsgBox "You must enter the facility's zip code"
		LOOP UNTIL (facility_name <> "" and facility_address_line_01 <> "" and facility_city <> "" and facility_state <> "" and facility_zip <> "")
		IF(send_OP_to_DHS_check = 1 AND (county_name_dept <> "" OR county_address_line_01 <> "" OR county_address_line_02 <> "" OR county_address_city <> "" OR county_address_state <> "" OR county_address_zip <> "")) THEN MsgBox "You must select either 'send the payment to DHS' or enter the county mailing information, not both options."
		IF(send_OP_to_DHS_check = 0 AND (county_name_dept = "" OR county_address_line_01 = "" OR county_address_city = "" OR county_address_state = "" OR county_address_zip = "")) THEN MsgBox "You must select either 'send the payment to DHS' or enter the county mailing information, not both options."
	LOOP UNTIL (send_OP_to_DHS_check = 1 AND (county_name_dept = "" AND county_address_line_01 = "" AND county_address_city = "" AND county_address_state = "" AND county_address_zip = "")) OR _
	(send_OP_to_DHS_check = 0 AND (county_name_dept <> "" AND county_address_line_01 <> "" AND county_address_city <> "" AND county_address_state <> "" AND county_address_zip <> ""))

'Checking to see that we're in MAXIS
Call check_for_MAXIS(False)

'Actions and calculations----------------------------------------------------------------------------------------------------
'Calculate OP total if nothing is entered.
IF OP_total = "" THEN
	If OP_amt_01 = "" THEN OP_amt_01 = "0"
	If OP_amt_02 = "" THEN OP_amt_02 = "0"
	If OP_amt_03 = "" THEN OP_amt_03 = "0"
	If OP_amt_04 = "" THEN OP_amt_04 = "0"
	If OP_amt_05 = "" THEN OP_amt_05 = "0"
	If OP_amt_06 = "" THEN OP_amt_06 = "0"
	OP_total = (Abs(OP_amt_01) + Abs(OP_amt_02) + Abs(OP_amt_03) + Abs(OP_amt_04) + Abs(OP_amt_05) + Abs(OP_amt_06)) & ""
END IF

'Dollar bill symbol will be added to numeric variables
IF OP_total <> "" THEN OP_total = "$" & OP_total
IF OP_amt_01 <> "" THEN OP_amt_01 = "$" & OP_amt_01
IF OP_amt_02 <> "" THEN OP_amt_02 = "$" & OP_amt_02
IF OP_amt_03 <> "" THEN OP_amt_03 = "$" & OP_amt_03
IF OP_amt_04 <> "" THEN OP_amt_04 = "$" & OP_amt_04
IF OP_amt_05 <> "" THEN OP_amt_05 = "$" & OP_amt_05
IF OP_amt_06 <> "" THEN OP_amt_06 = "$" & OP_amt_06
IF client_amt <> "" THEN client_amt = "$" & client_amt
IF client_amt = "" THEN client_amt = "$0"

'Sending the TIKL to the worker
If set_TIKL_check = checked THEN
	'navigates to DAIL/WRIT
	Call navigate_to_MAXIS_screen ("DAIL", "WRIT")
	'The following will generate a TIKL formatted date for 10 days from now.
	Call create_MAXIS_friendly_date(date, 30, 5, 18)
	'Writes TIKL to worker
	Call write_variable_in_TIKL("A letter of overpayment was sent for this case. Please check this case to see if follow up is required. Thank you.")
	'Saves TIKL and enters out of TIKL function
	transmit
	PF3
END If

'Sending the SPEC/MEMO to FACI----------------------------------------------------------------------------------------------------
'Navigates to SPEC/MEMO and selects a new MEMO
call navigate_to_MAXIS_screen("SPEC", "MEMO")
PF5
'Selects "other recipient of your choosing" instead of the client to send the MEMO to
other_row = 6
DO				'loop to search for OTHER recipient
	EMReadscreen find_other, 5, other_row, 12
	If find_other <> "OTHER" THEN other_row = other_row + 1
LOOP until find_other = "OTHER"
EMWritescreen "x", other_row, 10   'writes X on row where the phrase OTHER was found.
transmit
'Writes in Name of Facility and the address which MEMO is being sent
EMWritescreen facility_name, 13, 24
EMWritescreen facility_address_line_01, 14, 24
EMWritescreen facility_address_line_02, 15, 24
EMWritescreen facility_city, 16, 24
EMWritescreen facility_state, 17, 24
EMWritescreen facility_zip, 17, 32
'transmits 3 times to ensure that all edits are acknowledged and moves to blank SPEC/MEMO
transmit
transmit
transmit

'Writes the information in the SPEC/MEMO
Call write_variable_in_SPEC_MEMO ("Due to a change in placement, A GRH overpayment has occurred for this case for the following month(s):")
IF OP_amt_01 <> "" and OP_date_01 <> "" THEN Call write_variable_in_SPEC_MEMO("* " & OP_amt_01 & " for " & OP_date_01)
IF OP_amt_02 <> "" and OP_date_02 <> "" THEN Call write_variable_in_SPEC_MEMO("* " & OP_amt_02 & " for " & OP_date_02)
IF OP_amt_03 <> "" and OP_date_03 <> "" THEN Call write_variable_in_SPEC_MEMO("* " & OP_amt_03 & " for " & OP_date_03)
IF OP_amt_04 <> "" and OP_date_04 <> "" THEN Call write_variable_in_SPEC_MEMO("* " & OP_amt_04 & " for " & OP_date_04)
IF OP_amt_05 <> "" and OP_date_05 <> "" THEN Call write_variable_in_SPEC_MEMO("* " & OP_amt_05 & " for " & OP_date_05)
IF OP_amt_06 <> "" and OP_date_06 <> "" THEN Call write_variable_in_SPEC_MEMO("* " & OP_amt_06 & " for " & OP_date_06)
Call write_variable_in_SPEC_MEMO("The total amount of the overpayment to be returned is: " & OP_total)
Call write_variable_in_SPEC_MEMO("Reason for the overpayment(s):" & OP_reason)
Call write_variable_in_SPEC_MEMO("Amount client is responsible to pay for GRH during overpayment months (this amount is not to be subtracted from the total overpayment amount):" & client_amt)
Call write_variable_in_SPEC_MEMO("Please submit payment to:")
If send_OP_to_DHS_check = 1 THEN
	Call write_variable_in_SPEC_MEMO("Minnesota Department of Human Services")
	Call write_variable_in_SPEC_MEMO("MAXIS Cashier - 211")
	Call write_variable_in_SPEC_MEMO("PO BOX 64835")
	Call write_variable_in_SPEC_MEMO("St. Paul, MN 55164-0835")
ELSE
	Call write_variable_in_SPEC_MEMO(county_name_dept)
	Call write_variable_in_SPEC_MEMO(county_address_line_01)
	Call write_variable_in_SPEC_MEMO(county_address_line_02)
	Call write_variable_in_SPEC_MEMO(county_address_city & ", " & county_address_state & " " & county_address_zip)
END IF
Call write_variable_in_SPEC_MEMO(" ")
Call write_variable_in_SPEC_MEMO("Please include the case name, case number, month(s) of recovery and reason for the recovery with the payment. Please use the contact phone number on this letter if you have any questions. Thank you.")
'Saves and sends the MEMOS
PF4
PF3

'THE CASE NOTE -----------------------------------------------------------------------------------------------------------------
'Navigates to a blank case note
Call start_a_blank_CASE_NOTE

'Writes inforamtion from the dialog into the case note
Call write_variable_in_CASE_NOTE("****GRH OVERPAYMENT SENT DUE TO CLIENT LEAVING FACILITY****")
Call write_bullet_and_variable_in_case_note("Name of facility which OP was issued", facility_name)
Call write_variable_in_case_note("* Address OP was sent to: " & facility_address_line_01 & ", " & facility_address_line_02 & ", " & facility_city & ", " & facility_state & " " & facility_zip)
Call write_variable_in_CASE_NOTE ("*")
Call write_bullet_and_variable_in_case_note("Reason for overpayment(s)", OP_reason)
Call write_bullet_and_variable_in_case_note("Discovery date", discovery_date)
Call write_bullet_and_variable_in_case_note("Established date", established_date)
Call write_bullet_and_variable_in_case_note("Total overpayment amount", OP_total)
Call write_bullet_and_variable_in_case_note("Amount client is responsible to pay for GRH during overpayment months",client_amt)
IF OP_amt_01 <> "" and OP_date_01 <> "" THEN Call write_variable_in_CASE_NOTE("* " & OP_amt_01 & " for " & OP_date_01)
IF OP_amt_02 <> "" and OP_date_02 <> "" THEN Call write_variable_in_CASE_NOTE("* " & OP_amt_02 & " for " & OP_date_02)
IF OP_amt_03 <> "" and OP_date_03 <> "" THEN Call write_variable_in_CASE_NOTE("* " & OP_amt_03 & " for " & OP_date_03)
IF OP_amt_04 <> "" and OP_date_04 <> "" THEN Call write_variable_in_CASE_NOTE("* " & OP_amt_04 & " for " & OP_date_04)
IF OP_amt_05 <> "" and OP_date_05 <> "" THEN Call write_variable_in_CASE_NOTE("* " & OP_amt_05 & " for " & OP_date_05)
IF OP_amt_06 <> "" and OP_date_06 <> "" THEN Call write_variable_in_CASE_NOTE("* " & OP_amt_06 & " for " & OP_date_06)
Call write_bullet_and_variable_in_case_note ("Instructed FACI to send overpayment to", county_name_dept)
If send_OP_to_DHS_check = 1 THEN write_variable_in_CASE_NOTE("* Instructed FACI to send overpayment to DHS.")
If set_TIKL_check = 1 THEN write_variable_in_CASE_NOTE ("* TIKL'd to recheck case in 30 days")
Call write_variable_in_case_note("---")
Call write_variable_in_case_note(worker_signature)
MsgBox "A MEMO has been sent.  Please refer to your agency's overpayment procedure to ensure the overpayment process is complete.  Thank you."

script_end_procedure ""
