'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - LTC - 1503.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'Required for statistical purposes==========================================================================================
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 360          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 141, 80, "Case number dialog"
  EditBox 65, 10, 65, 15, case_number
  EditBox 65, 30, 30, 15, MAXIS_footer_month
  EditBox 100, 30, 30, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 25, 55, 50, 15
    CancelButton 80, 55, 50, 15
  Text 10, 30, 50, 15, "Footer month:"
  Text 10, 10, 50, 15, "Case number: "
EndDialog

BeginDialog DHS_1503_dialog, 0, 0, 366, 275, "1503 Dialog"
  EditBox 55, 5, 135, 15, FACI
  DropListBox 255, 5, 95, 15, "30 days or less"+chr(9)+"31 to 90 days"+chr(9)+"91 to 180 days"+chr(9)+"over 180 days", length_of_stay
  DropListBox 105, 25, 45, 15, "SNF"+chr(9)+"NF"+chr(9)+"ICF-MR"+chr(9)+"RTC", level_of_care
  DropListBox 215, 25, 135, 15, "acute-care hospital"+chr(9)+"home"+chr(9)+"RTC"+chr(9)+"other SNF or NF"+chr(9)+"ICF-MR", admitted_from
  EditBox 145, 45, 205, 15, hospital_admitted_from
  EditBox 75, 65, 65, 15, admit_date
  EditBox 275, 65, 75, 15, discharge_date
  CheckBox 15, 85, 155, 10, "If you've processed this 1503, check here.", processed_1503_check
  CheckBox 15, 115, 60, 10, "Updated RLVA?", updated_RLVA_check
  CheckBox 85, 115, 60, 10, "Updated FACI?", updated_FACI_check
  CheckBox 150, 115, 50, 10, "Need 3543?", need_3543_check
  CheckBox 205, 115, 55, 10, "Need 3531?", need_3531_check
  CheckBox 265, 115, 95, 10, "Need asset assessment?", need_asset_assessment_check
  EditBox 130, 130, 225, 15, verifs_needed
  CheckBox 15, 155, 85, 10, "Sent 3050 back to LTCF", sent_3050_check
  CheckBox 165, 155, 100, 10, "Sent verif req? If so, to who:", sent_verif_request_check
  ComboBox 270, 150, 85, 15, "client"+chr(9)+"AREP"+chr(9)+"Client & AREP", sent_request_to
  EditBox 30, 180, 325, 15, notes
  EditBox 170, 200, 75, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 250, 200, 50, 15
    CancelButton 305, 200, 50, 15
  CheckBox 10, 230, 260, 10, "Check here to have the script TIKL out to contact the FACI re: length of stay.", TIKL_check
  CheckBox 10, 245, 155, 10, "Check here to have the script update HCMI.", HCMI_update_check
  CheckBox 10, 260, 150, 10, "Check here to have the script update FACI.", FACI_update_check
  Text 105, 205, 60, 10, "Worker signature:"
  Text 5, 185, 25, 10, "Notes:"
  Text 5, 50, 135, 10, "If hospital, list name/dates of admission:"
  GroupBox 5, 100, 355, 75, "actions/proofs"
  Text 5, 30, 95, 10, "Recommended level of care:"
  Text 10, 135, 115, 10, "Other proofs needed (if applicable):"
  Text 160, 30, 50, 10, "Admitted from:"
  Text 5, 10, 47, 10, "Facility name:"
  Text 5, 70, 65, 10, "Date of admission:"
  Text 165, 70, 105, 10, "Date of discharge (if applicible):"
  Text 200, 10, 50, 10, "Length of stay:"
  GroupBox 5, 220, 260, 55, "Script actions"
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
'connecting to MAXIS & grabs the case number and footer month/year
EMConnect ""
call MAXIS_case_number_finder(case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'The initial dialog----------------------------------------------------------------------------------------------------
DO
	Dialog case_number_dialog
	cancel_confirmation
	IF IsNumeric(case_number) = FALSE THEN MsgBox "You must type a valid case number."
	IF IsNumeric(MAXIS_footer_month) = FALSE THEN MsgBox "You must type a valid footer month."
	IF IsNumeric(MAXIS_footer_year) = FALSE THEN MsgBox "You must type a valid footer year."
LOOP UNTIL IsNumeric(case_number) = TRUE and IsNumeric(MAXIS_footer_month) = TRUE and IsNumeric(MAXIS_footer_year) = True

'THE 1503 MAIN DIALOG----------------------------------------------------------------------------------------------------
Do
	Dialog DHS_1503_dialog
	cancel_confirmation
	IF worker_signature = "" THEN MsgBox "You must sign your case note."
LOOP UNTIL worker_signature <> ""  

'Checks for an active MAXIS session
call check_for_MAXIS(False)
'checking to make sure case is out of background
MAXIS_background_check

'navigating the script to the correct footer month
back_to_self
EMWriteScreen MAXIS_footer_month, 20, 43
EMWriteScreen MAXIS_footer_year, 20, 46
call navigate_to_MAXIS_screen("STAT", "FACI")

'THE TIKL----------------------------------------------------------------------------------------------------
If TIKL_check = 1 then
  If length_of_stay = "30 days or less" then TIKL_multiplier = 30
  If length_of_stay = "31 to 90 days" then TIKL_multiplier = 90
  If length_of_stay = "91 to 180 days" then TIKL_multiplier = 180
  TIKL_date = dateadd("d", TIKL_multiplier, admit_date)
  TIKL_date_DD = datepart("d", TIKL_date)
  If len(TIKL_date_DD) = 1 then TIKL_date_DD = "0" & TIKL_date_DD
  TIKL_date_MM = datepart("m", TIKL_date)
  If len(TIKL_date_MM) = 1 then TIKL_date_MM = "0" & TIKL_date_MM
  TIKL_date_YY = datepart("yyyy", TIKL_date)
  If len(TIKL_date_YY) = 4 then TIKL_date_YY = TIKL_date_YY - 2000
End if

'UPDATING MAXIS PANELS----------------------------------------------------------------------------------------------------
'FACI
If FACI_update_check = 1 then
	call navigate_to_MAXIS_screen("stat", "faci")
	EMReadScreen panel_max_check, 1, 2, 78
	IF panel_max_check = "5" THEN 
		script_end_procedure ("This case has reached the maximum amount of FACI panels.  Please review your case, delete an appropriate FACI panel, and run the script again.  Thank you.")
	ELSE
		EMWriteScreen "nn", 20, 79
		transmit
	END IF 
	EMWriteScreen FACI, 6, 43
	If level_of_care = "NF" then EMWriteScreen "42", 7, 43
	If level_of_care = "RTC" THEN EMWriteScreen "47", 7, 43
	If length_of_stay = "30 days or less" and level_of_care = "SNF" then EMWriteScreen "44", 7, 43
	If length_of_stay = "31 to 90 days" and level_of_care = "SNF" then EMWriteScreen "41", 7, 43
	If length_of_stay = "91 to 180 days" and level_of_care = "SNF" then EMWriteScreen "41", 7, 43
	if length_of_stay = "over 180 days" and level_of_care = "SNF" then EMWriteScreen "41", 7, 43
	If length_of_stay = "30 days or less" and level_of_care = "ICF-MR" then EMWriteScreen "44", 7, 43
	If length_of_stay = "31 to 90 days" and level_of_care = "ICF-MR" then EMWriteScreen "41", 7, 43
	If length_of_stay = "91 to 180 days" and level_of_care = "ICF-MR" then EMWriteScreen "41", 7, 43
	If length_of_stay = "over 180 days" and level_of_care = "ICF-MR" then EMWriteScreen "41", 7, 43
	EMWriteScreen "n", 8, 43
	Call create_MAXIS_friendly_date_with_YYYY(admit_date, 0, 14, 47)
	If discharge_date<> "" then
		Call create_MAXIS_friendly_date_with_YYYY(discharge_date, 0, 14, 71)
		transmit
		transmit
	End if
End if

'HCMI
If HCMI_update_check = 1 THEN
	call navigate_to_MAXIS_screen("stat", "hcmi") 
	EMReadScreen HCMI_panel_check, 1, 2, 78
	IF HCMI_panel_check <> 0 Then
		PF9
	ELSE
		EMWriteScreen "nn", 20, 79
		transmit
	END IF 
	EMWriteScreen "dp", 10, 57
	transmit
	transmit
END IF

'THE TIKL----------------------------------------------------------------------------------------------------
If TIKL_check = 1 then
  call navigate_to_MAXIS_screen("dail", "writ")
  EMWriteScreen TIKL_date_MM, 5, 18
  EMWriteScreen TIKL_date_DD, 5, 21
  EMWriteScreen TIKL_date_YY, 5, 24
  EMSetCursor 9, 3
  write_variable_in_TIKL("Have " & worker_signature & " call " & FACI & " re: length of stay. " & TIKL_multiplier & " days expired.")
  transmit
  PF3
End if

'The CASE NOTE----------------------------------------------------------------------------------------------------
Call start_a_blank_CASE_NOTE
If processed_1503_check = 1 then 
  call write_variable_in_CASE_NOTE("***Processed 1503 from " & FACI & "***")
Else
  call write_variable_in_CASE_NOTE("***Rec'd 1503 from " & FACI & ", DID NOT PROCESS***")
End if
Call write_bullet_and_variable_in_case_note("Length of stay", length_of_stay)
Call write_bullet_and_variable_in_case_note("Recommended level of care", level_of_care)
Call write_bullet_and_variable_in_case_note("Admitted from", admitted_from)
Call write_bullet_and_variable_in_case_note("Hospital admitted from", hospital_admitted_from)
Call write_bullet_and_variable_in_case_note("Admit date", admit_date)
Call write_bullet_and_variable_in_case_note("Discharge date", discharge_date)
Call write_variable_in_CASE_NOTE("---")
If updated_RLVA_check = 1 and updated_FACI_check = 1 then 
Call write_variable_in_CASE_NOTE("* Updated RLVA and FACI.")
Else
  If updated_RLVA_check = 1 then Call write_variable_in_case_note("* Updated RLVA.")
  If updated_FACI_check = 1 then Call write_variable_in_case_note("* Updated FACI.")
End if
If need_3543_check = 1 then Call write_variable_in_case_note("* A 3543 is needed for the client.")
If need_3531_check = 1 then call write_variable_in_CASE_NOTE("* A 3531 is needed for the client.")
If need_asset_assessment_check = 1 then call write_variable_in_CASE_NOTE("* An asset assessment is needed before a MA-LTC determination can be made.")
If sent_3050_check = 1 then call write_variable_in_CASE_NOTE("* Sent 3050 back to LTCF.")
Call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
If sent_verif_request_check = 1 then Call write_variable_in_case_note("* Sent verif request to " & sent_request_to)
If processed_1503_check = 1 then Call write_variable_in_case_note("* Completed & Returned 1503 to LTCF.")
If TIKL_check = 1 then Call write_variable_in_case_note("* TIKLed to recheck length of stay on " & TIKL_date & ".")
Call write_variable_in_case_note("---")
Call write_bullet_and_variable_in_case_note("Notes", notes)
Call write_variable_in_case_note("---")
Call write_variable_in_case_note(worker_signature)

script_end_procedure("")
