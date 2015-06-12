'OPTION EXPLICIT

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - MFIP SANCTION AND DWP DISQUALIFICATION.vbs"
start_time = timer

'DIM name_of_script
'DIM start_time
'DIM FuncLib_URL
'DIM run_locally
'DIM default_directory
'DIM beta_agency
'DIM req
'DIM fso

''LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
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
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'Dimming variables----------------------------------------------------------------------------------------------------
'DIM MFIP_Sanction_DWP_Disq_Dialog
'DIM case_number
'DIM sanction_status_droplist
'DIM HH_Member_Number
'DIM sanction_type_droplist
'DIM number_occurances_droplist
'DIM Date_Sanction
'DIM Sanction_Percentage_droplist
'DIM sanction_information
'DIM sanction_reason_droplist
'DIM other_sanction_notes
'DIM Memo_to_Client
'DIM Impact_Other_Programs
'DIM Vendor_Information
'DIM Last_Day_Cure
'DIM Update_Sent_ES_Checkbox
'DIM FIAT_check
'DIM Update_Sent_CCA_Checkbox
'DIM mandatory_vendor_check
'DIM TIKL_next_month
'DIM Sent_SPEC_MEMO
'DIM set_TIKL_check
'DIM worker_signature
'DIM ButtonPressed
'DIM TIKL_date


'DIALOGS----------------------------------------------------------------------------------------------------
'MFIP Sanction/DWP Disqualification Dialog Box
BeginDialog MFIP_Sanction_DWP_Disq_Dialog, 0, 0, 341, 275, "MFIP Sanction - DWP Disqualification"
  EditBox 55, 5, 60, 15, case_number
  EditBox 180, 5, 20, 15, HH_Member_Number
  DropListBox 265, 5, 65, 15, "Select one..."+chr(9)+"imposed"+chr(9)+"pending", sanction_status_droplist
  DropListBox 65, 25, 110, 15, "Select one..."+chr(9)+"CS"+chr(9)+"ES"+chr(9)+"No show to orientation"+chr(9)+"Minor mom truancy", sanction_type_droplist
  DropListBox 265, 25, 65, 15, "Select one..."+chr(9)+"1"+chr(9)+"2"+chr(9)+"3"+chr(9)+"4"+chr(9)+"5"+chr(9)+"6"+chr(9)+"7"+chr(9)+"7+", number_occurances_droplist
  DropListBox 50, 45, 65, 15, "Select one..."+chr(9)+"10%"+chr(9)+"30%"+chr(9)+"100%", Sanction_Percentage_droplist
  EditBox 265, 45, 65, 15, Date_Sanction
  DropListBox 90, 65, 240, 15, "Select one..."+chr(9)+"Failed to attend ES overview"+chr(9)+"Failed to develop employment plan"+chr(9)+"Non-compliance with employment plan"+chr(9)+"< 20, failed education requirement"+chr(9)+"Failed to accept suitable employment"+chr(9)+"Quit suitable employment w/o good cause"+chr(9)+"Failure to attend MFIP orientation"+chr(9)+"Non-cooperation with child support", sanction_reason_droplist
  EditBox 90, 85, 240, 15, sanction_information
  EditBox 90, 105, 240, 15, other_sanction_notes
  EditBox 90, 125, 240, 15, Impact_Other_Programs
  EditBox 90, 145, 240, 15, Vendor_Information
  EditBox 175, 165, 60, 15, Last_Day_Cure
  CheckBox 5, 190, 130, 10, "Update sent to Employment Services", Update_Sent_ES_Checkbox
  CheckBox 5, 205, 130, 10, "Update sent to Child Care Assistance", Update_Sent_CCA_Checkbox
  CheckBox 5, 220, 130, 10, "TIKL to change sanction status ", TIKL_next_month
  CheckBox 145, 190, 80, 10, "Case has been FIAT'd", Fiat_check
  CheckBox 145, 205, 140, 10, "Mandatory vendor form mailed to client", mandatory_vendor_check
  CheckBox 145, 220, 190, 10, "Sent MFIP sanction for future closed month SPEC/LETR", Sent_SPEC_MEMO
  EditBox 150, 255, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 230, 255, 50, 15
    CancelButton 285, 255, 50, 15
  Text 5, 70, 80, 10, "Reason for the sanction:"
  Text 85, 260, 60, 10, "Worker signature:"
  Text 5, 170, 170, 10, "Last day to cure (10 day cutoff or last day of month):"
  Text 185, 30, 75, 10, "Number of occurences:"
  Text 5, 150, 65, 10, "Vendor information:"
  Text 5, 130, 85, 10, "Impact to other programs:"
  Text 5, 90, 80, 10, "Sanction info from/how:"
  Text 5, 50, 40, 10, "Sanction %:"
  Text 210, 10, 55, 10, "Sanction status:"
  Text 5, 10, 45, 10, "Case number:"
  Text 125, 50, 140, 10, "Effective Date of Sanction/Disqualification:"
  Text 130, 10, 50, 10, "HH Member #:"
  Text 5, 30, 60, 10, "Type of sanction:"
  Text 5, 110, 70, 10, "Other sanction notes:"
  Text 155, 230, 160, 10, "(See TE10.20 for info on when to use this notice)"
  GroupBox 0, 185, 335, 60, ""
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

'Asks for Case Number
CALL MAXIS_case_number_finder(case_number)

'Shows dialog
DO
	DO
		DO
			DO
				DO
					DO
						DO
							DO
								DO	
									DO
										DO
											DO								
												Dialog MFIP_Sanction_DWP_Disq_Dialog
												cancel_confirmation
												IF IsNumeric(case_number) = FALSE THEN MsgBox "You must type a valid numeric case number"
											LOOP UNTIL IsNumeric(case_number) = TRUE
											IF sanction_status_droplist = "Select one..." THEN MsgBox "You must select a sanction status type"
										LOOP UNTIL Sanction_status_droplist <> "Select one..."
										IF HH_Member_Number = "" THEN MsgBox "You must enter a HH member number"
									LOOP UNTIL HH_Member_Number <> ""
									IF sanction_type_droplist = "Select one..." THEN MsgBox "You must select a sanction type"
								LOOP UNTIL sanction_type_droplist <> "Select one..."
								IF number_occurances_droplist = "Select one..." THEN MsgBox "You must select a number of the sanction occurrence"
							LOOP UNTIL number_occurances_droplist <> "Select one..."
							IF IsDate(Date_Sanction) = FALSE THEN MsgBox "You must type a valid date of sanction"
						LOOP UNTIL IsDate(Date_Sanction) = TRUE
						IF Sanction_Percentage_droplist = "Select one..." THEN MsgBox "You must select a sanction percentage"
					LOOP UNTIL Sanction_Percentage_droplist <> "Select one..."
					IF sanction_information = "" THEN MsgBox "You must enter information about how the sanction information was received"
				LOOP UNTIL sanction_information <> ""
				IF IsDate(Date_Sanction) = FALSE THEN MsgBox "You must type a valid date of sanction"
			LOOP UNTIL IsDate(Date_Sanction) = TRUE
			IF sanction_reason_droplist = "Select One..." THEN MsgBox "You must select a sanction percentage"
		LOOP UNTIL sanction_reason_droplist <> "Select One..."
		IF Last_Day_Cure = "" THEN MsgBox "You must enter the last day to cure the sanction"
	LOOP UNTIL Last_Day_Cure <> ""
	IF worker_signature = "" THEN MsgBox "You must sign your case note"
LOOP UNTIL worker_signature <> ""


'Checks MAXIS for password prompt
Call check_for_MAXIS(True)

'TIKL to change sanction status (check box selected)
If TIKL_next_month = checked THEN 
	'navigates to DAIL/WRIT 
	Call navigate_to_MAXIS_screen ("DAIL", "WRIT")	
	
	TIKL_date = dateadd("m", 1, date)		'Creates a TIKL_date variable with the current date + 1 month (to determine what the month will be next month)
	TIKL_date = datepart("m", TIKL_date) & "/01/" & datepart("yyyy", TIKL_date)		'Modifies the TIKL_date variable to reflect the month, the string "/01/", and the year from TIKL_date, which creates a TIKL date on the first of next month.
	
	'The following will generate a TIKL formatted date for 10 days from now.
	Call create_MAXIS_friendly_date(TIKL_date, 0, 5, 18) 'updates to first day of the next available month dateadd(m, 1)
	'Writes TIKL to worker
	Call write_variable_in_TIKL("A pending sanction was determined last month.  Please review case, and resolve or impose the sanction.")
	'Saves TIKL and enters out of TIKL function
	transmit
	PF3
END If

'Navigates to case note
CALL start_a_blank_CASE_NOTE

'Writes case note
'case noting the droplist and editboxes
Call write_variable_in_case_note("***" & Sanction_Percentage_droplist & " " & sanction_status_droplist & " SANCTION" & " for " & "MEMB " & HH_Member_Number & " eff: " & Date_Sanction & "***")
CALL write_bullet_and_variable_in_case_note("HH member number", HH_Member_Number)
Call write_bullet_and_variable_in_case_note("Sanction status", sanction_status_droplist)
CALL write_bullet_and_variable_in_case_note("Type of Sanction", sanction_type_droplist)
CALL write_bullet_and_variable_in_case_note("Number of occurences", number_occurances_droplist)
CALL write_bullet_and_variable_in_case_note("Sanction Percent is", Sanction_Percentage_droplist)
CALL write_bullet_and_variable_in_case_note("Effective date of sanction/disqualification", Date_Sanction)
CALL write_bullet_and_variable_in_case_note("Sanction information received from", sanction_information)
CALL write_bullet_and_variable_in_case_note ("Reason for the sanction", sanction_reason_droplist)
CALL write_bullet_and_variable_in_case_note("Other sanction notes", other_sanction_notes)
CALL write_bullet_and_variable_in_case_note ("Impact to other programs", Impact_Other_Programs)
CALL write_bullet_and_variable_in_case_note("Vendoring information", Vendor_Information)
CALL write_bullet_and_variable_in_case_note("Last day to cure", Last_Day_Cure)
'case noting check boxes if checked
IF Update_Sent_ES_Checkbox = 1 THEN CALL write_variable_in_case_note("* Status update information was sent to Employment Services.")
IF Update_Sent_CCA_Checkbox = 1 THEN CALL write_variable_in_case_note("* Status update information was sent to Child Care Assistance.")
IF TIKL_next_month = 1 THEN Call write_variable_in_case_note("* A TIKL was set to update the case from pending to imposed for the 1st of         the next month.")
IF FIAT_check = 1 THEN CALL write_variable_in_case_note("* Case has been FIATed.")
IF mandatory_vendor_check = 1 THEN CALL write_variable_in_case_note("* A mandatory vendor form has been mailed to the sanctioned individual.")
IF Sent_SPEC_MEMO = 1 THEN CALL write_variable_in_case_note ("* Sent MFIP sanction for future closed month SPEC/MEMO to the sanctioned           individual.")
CALL write_variable_in_case_note("---")
CALL write_variable_in_case_note(worker_signature)

script_end_procedure ("")
