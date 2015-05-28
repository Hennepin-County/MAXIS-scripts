'OPTION EXPLICIT

name_of_script = "MEMOS - GRH OP CL LEFT FACI.vbs"
start_time = timer

'DIM name_of_script
'DIM start_time
'DIM FuncLib_URL
'DIM run_locally
'DIM default_directory
'DIM beta_agency
'DIM req
'DIM fso


'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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


'DECLARING VARIABLES----------------------------------------------------------------------------------------------------
'DIM ButtonPressed
'DIM GRH_OP_LEAVING_FACI_dialog
'DIM case_number
'DIM total_OP_amt
'DIM facility_name
'DIM facility_address_line_01
'DIM facility_address_line_02
'DIM facility_city
'DIM facility_state
'DIM facility_zip
'DIM OP_reason
'DIM discovery_date
'DIM established_date
'DIM OP_date_01
'DIM OP_date_02
'DIM OP_date_03
'DIM OP_date_04
'DIM OP_date_05
'DIM OP_date_06
'DIM OP_amt_01
'DIM OP_amt_02
'DIM OP_amt_03
'DIM OP_amt_04
'DIM OP_amt_05
'DIM OP_amt_06
'DIM county_name_dept
'DIM county_address_line_01
'DIM county_address_line_02
'DIM county_address_city
'DIM county_address_state
'DIM county_address_zip
'DIM send_OP_to_DHS_check
'DIM set_TIKL_check
'DIM worker_signature
'DIM first_name
'DIM last_name


'DIALOG----------------------------------------------------------------------------------------------------
BeginDialog GRH_OP_LEAVING_FACI_dialog, 0, 0, 306, 390, "GRH overpayment due to leaving facility dialog"
  EditBox 50, 5, 55, 15, case_number
  EditBox 210, 5, 55, 15, total_OP_amt
  EditBox 70, 45, 230, 15, facility_name
  EditBox 70, 65, 230, 15, facility_address_line_01
  EditBox 70, 85, 230, 15, facility_address_line_02
  EditBox 70, 105, 80, 15, facility_city
  EditBox 155, 105, 25, 15, facility_state
  EditBox 185, 105, 45, 15, facility_zip
  EditBox 90, 135, 210, 15, OP_reason
  EditBox 135, 155, 50, 15, discovery_date
  EditBox 250, 155, 50, 15, established_date
  EditBox 45, 180, 45, 15, OP_date_01
  EditBox 110, 180, 30, 15, OP_amt_01
  EditBox 185, 180, 45, 15, OP_date_02
  EditBox 255, 180, 45, 15, OP_amt_02
  EditBox 45, 200, 45, 15, OP_date_03
  EditBox 110, 200, 30, 15, OP_amt_03
  EditBox 185, 200, 45, 15, OP_date_04
  EditBox 255, 200, 45, 15, OP_amt_04
  EditBox 45, 220, 45, 15, OP_date_05
  EditBox 110, 220, 30, 15, OP_amt_05
  EditBox 185, 220, 45, 15, OP_date_06
  EditBox 255, 220, 45, 15, OP_amt_06
  EditBox 65, 265, 235, 15, county_name_dept
  EditBox 65, 285, 235, 15, county_address_line_01
  EditBox 65, 305, 235, 15, county_address_line_02
  EditBox 65, 325, 80, 15, county_address_city
  EditBox 150, 325, 25, 15, county_address_state
  EditBox 180, 325, 45, 15, county_address_zip
  CheckBox 40, 350, 95, 10, "Send overpayment to DHS", send_OP_to_DHS_check
  CheckBox 155, 350, 125, 10, "Set TIKL to recheck case in 30 days", set_TIKL_check
  EditBox 100, 365, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 195, 365, 50, 15
    CancelButton 250, 365, 50, 15
  Text 65, 250, 235, 10, "**COUNTY ADDRESS WHERE THE OVERPAYMENT WILL BE SENT**"
  Text 40, 370, 60, 10, "Worker signature:"
  Text 5, 310, 55, 10, "Address Line 2:"
  Text 95, 185, 15, 10, "Amt:"
  Text 120, 10, 90, 10, "Total overpayment amount:"
  Text 5, 290, 55, 10, "Address Line 1:"
  Text 145, 185, 40, 10, "Date of OP:"
  Text 5, 350, 25, 10, "**OR**"
  Text 235, 185, 15, 10, "Amt:"
  Text 10, 330, 50, 10, "City/State/Zip:"
  Text 5, 205, 40, 10, "Date of OP:"
  Text 20, 110, 50, 10, "City/State/Zip:"
  Text 95, 205, 15, 10, "Amt:"
  Text 5, 140, 85, 10, "Reason for overpayment:"
  Text 145, 205, 40, 10, "Date of OP:"
  Text 5, 10, 45, 10, "Case number:"
  Text 235, 205, 15, 10, "Amt:"
  Text 80, 160, 55, 10, "Discovery date:"
  Text 5, 225, 40, 10, "Date of OP:"
  Text 5, 185, 40, 10, "Date of OP:"
  Text 95, 225, 15, 10, "Amt:"
  Text 5, 90, 65, 10, "FACI ADDR Line 2:"
  Text 145, 225, 40, 10, "Date of OP:"
  Text 5, 70, 60, 10, "FACI ADDR line 1:"
  Text 235, 225, 15, 10, "Amt:"
  Text 40, 30, 265, 10, "**FACILITY ADDRESS WHERE THE OVERPAYMENT MEMO WILL BE SENT**"
  Text 25, 50, 40, 10, "FACI Name:"
  Text 190, 160, 60, 10, "Established date:"
  Text 0, 270, 65, 10, "County Name/Dept:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to MAXIS
EMConnect ""
'searches for a case number
call MAXIS_case_number_finder(case_number)

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
								If case_number = "" or isnumeric(case_number) = false THEN MsgBox "You did not enter a valid case number. Please try again."
								If worker_signature = "" THEN MsgBox "You did not sign your case note. Please try again."
								If total_OP_amt = "" THEN MsgBox "You must enter the total amount of the overpayment"
								If facility_name = "" THEN MsgBox "You must enter the facility name"
								If facility_address_line_01 = "" THEN MsgBox "You must enter the facility's street address"
								If facility_city = "" THEN MsgBox "You must enter the facility's city"
								If facility_state = "" THEN MsgBox "You must enter the facility's state"
								If facility_zip = "" THEN MsgBox "You must enter the facility's zip code"
								If OP_reason = "" THEN MsgBox "You must enter the reason for the overpayment"
								If discovery_date = "" THEN MsgBox "You must enter the discovery date"
								If established_date = "" THEN MsgBox "You must enter the established date"
								If OP_date_01 = "" THEN MsgBox "You must enter at least one overpayment date"
								If OP_amt_01 = "" THEN MsgBox "You must enter at least one overpayment amount"
							Loop until case_number <> "" and isnumeric(case_number) = true and worker_signature <> "" and total_OP_amt <> "" and facility_name <> "" and facility_address_line_01 <> "" and facility_city <> "" and facility_state <> "" and facility_zip <> "" and OP_reason <> "" and discovery_date <> "" and established_date <> "" and OP_date_01 <> "" and OP_amt_01 <> ""
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
	IF(send_OP_to_DHS_check = 1 AND (county_name_dept <> "" OR county_address_line_01 <> "" OR county_address_line_02 <> "" OR county_address_city <> "" OR county_address_state <> "" OR county_address_zip <> "")) THEN MsgBox "You must select either 'send the payment to DHS' or enter the county mailing information, not both options."
	IF(send_OP_to_DHS_check = 0 AND (county_name_dept = "" OR county_address_line_01 = "" OR county_address_city = "" OR county_address_state = "" OR county_address_zip = "")) THEN MsgBox "You must select either 'send the payment to DHS' or enter the county mailing information, not both options."
LOOP UNTIL (send_OP_to_DHS_check = 1 AND (county_name_dept = "" AND county_address_line_01 = "" AND county_address_city = "" AND county_address_state = "" AND county_address_zip = "")) OR _
  (send_OP_to_DHS_check = 0 AND (county_name_dept <> "" AND county_address_line_01 <> "" AND county_address_city <> "" AND county_address_state <> "" AND county_address_zip <> ""))

'Checking to see that we're in MAXIS
Call check_for_MAXIS(False)

'Actions and calculations----------------------------------------------------------------------------------------------------
'Dollar bill symbol will be added to numeric variables 
IF total_OP_amt <> "" THEN total_OP_amt = "$" & total_OP_amt
IF OP_amt_01 <> "" THEN OP_amt_01 = "$" & OP_amt_01
IF OP_amt_02 <> "" THEN OP_amt_02 = "$" & OP_amt_02
IF OP_amt_03 <> "" THEN OP_amt_03 = "$" & OP_amt_03
IF OP_amt_04 <> "" THEN OP_amt_04 = "$" & OP_amt_04
IF OP_amt_05 <> "" THEN OP_amt_05 = "$" & OP_amt_05
IF OP_amt_06 <> "" THEN OP_amt_06 = "$" & OP_amt_06


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
EMWritescreen "x", 6, 10
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
Call write_variable_in_SPEC_MEMO("The total amount of the overpayment to be returned is: " & total_OP_amt)
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
Call write_bullet_and_variable_in_case_note("Total overpayment amount", total_OP_amt)
IF OP_amt_01 <> "" and OP_date_01 <> "" THEN Call write_variable_in_SPEC_MEMO("* " & OP_amt_01 & " for " & OP_date_01)
IF OP_amt_02 <> "" and OP_date_02 <> "" THEN Call write_variable_in_SPEC_MEMO("* " & OP_amt_02 & " for " & OP_date_02)
IF OP_amt_03 <> "" and OP_date_03 <> "" THEN Call write_variable_in_SPEC_MEMO("* " & OP_amt_03 & " for " & OP_date_03)
IF OP_amt_04 <> "" and OP_date_04 <> "" THEN Call write_variable_in_SPEC_MEMO("* " & OP_amt_04 & " for " & OP_date_04)
IF OP_amt_05 <> "" and OP_date_05 <> "" THEN Call write_variable_in_SPEC_MEMO("* " & OP_amt_05 & " for " & OP_date_05)
IF OP_amt_06 <> "" and OP_date_06 <> "" THEN Call write_variable_in_SPEC_MEMO("* " & OP_amt_06 & " for " & OP_date_06)
Call write_bullet_and_variable_in_case_note ("Instructed FACI to send overpayment to", county_name_dept)
If send_OP_to_DHS_check = 1 THEN write_variable_in_CASE_NOTE("*  Instructed FACI to send overpayment to DHS.")
If set_TIKL_check = 1 THEN write_variable_in_CASE_NOTE ("* TIKL'd to recheck case in 30 days")
Call write_variable_in_case_note("---")
Call write_variable_in_case_note(worker_signature)
MsgBox "A MEMO has been sent.  Please refer to your agency's overpayment procedure to ensure the overpayment process is complete.  Thank you."

script_end_procedure ""