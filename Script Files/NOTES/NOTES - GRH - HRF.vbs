'STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - GRH - HRF.vbs"
start_time = timer

'BY USING THIS SCRIPT, WE ARE ASSUMING THE USER IS RUNNING IT ON POST PAY CASES ONLY'-----------------------------------------

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" OR default_directory = "" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
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

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
next_month = dateadd("m", + 1, date)
footer_month = datepart("m", next_month)
If len(footer_month) = 1 then footer_month = "0" & footer_month
footer_year = datepart("yyyy", next_month)
footer_year = "" & footer_year - 2000


'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 181, 100, "Case number dialog"
  EditBox 80, 5, 70, 15, case_number
  EditBox 65, 25, 30, 15, footer_month
  EditBox 140, 25, 30, 15, footer_year
  CheckBox 10, 60, 30, 10, "GRH", cash_check
  CheckBox 50, 60, 30, 10, "HC", HC_check
  ButtonGroup ButtonPressed
    OkButton 35, 80, 50, 15
    CancelButton 95, 80, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 110, 30, 25, 10, "Year:"
  GroupBox 5, 45, 85, 30, "Programs recertifying"
  Text 95, 45, 80, 30, "Please note: this script assumes you're working a postpay case."
EndDialog

BeginDialog HRF_dialog, 0, 0, 451, 240, "HRF dialog"
  EditBox 65, 25, 50, 15, HRF_datestamp
  DropListBox 170, 25, 75, 15, "(select one...)"+chr(9)+"complete"+chr(9)+"incomplete", HRF_status
  EditBox 65, 45, 380, 15, earned_income
  EditBox 70, 65, 375, 15, unearned_income
  EditBox 30, 85, 90, 15, YTD
  EditBox 170, 85, 275, 15, changes
  EditBox 100, 105, 345, 15, FIAT_reasons
  EditBox 50, 125, 395, 15, other_notes
  CheckBox 240, 145, 150, 10, "Sent forms to AREP?", sent_arep_checkbox
  EditBox 240, 160, 205, 15, verifs_needed
  EditBox 235, 180, 210, 15, actions_taken
  CheckBox 125, 205, 180, 10, "Check here to case note grant info from ELIG/GRH.", grab_GRH_info_check
  EditBox 235, 220, 65, 15, date_check_sent_to_facility
  EditBox 380, 200, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 340, 220, 50, 15
    CancelButton 395, 220, 50, 15
    PushButton 280, 15, 20, 10, "GRH", ELIG_MFIP_button
    PushButton 300, 15, 20, 10, "HC", ELIG_HC_button
    PushButton 335, 15, 45, 10, "prev. panel", prev_panel_button
    PushButton 335, 25, 45, 10, "next panel", next_panel_button
    PushButton 395, 15, 45, 10, "prev. memb", prev_memb_button
    PushButton 395, 25, 45, 10, "next memb", next_memb_button
    PushButton 10, 155, 25, 10, "BUSI", BUSI_button
    PushButton 35, 155, 25, 10, "JOBS", JOBS_button
    PushButton 10, 165, 25, 10, "RBIC", RBIC_button
    PushButton 35, 165, 25, 10, "UNEA", UNEA_button
    PushButton 75, 155, 25, 10, "ACCT", ACCT_button
    PushButton 100, 155, 25, 10, "CARS", CARS_button
    PushButton 125, 155, 25, 10, "CASH", CASH_button
    PushButton 150, 155, 25, 10, "OTHR", OTHR_button
    PushButton 75, 165, 25, 10, "REST", REST_button
    PushButton 100, 165, 25, 10, "SECU", SECU_button
    PushButton 125, 165, 25, 10, "TRAN", TRAN_button
    PushButton 10, 195, 25, 10, "MEMB", MEMB_button
    PushButton 35, 195, 25, 10, "MEMI", MEMI_button
    PushButton 60, 195, 25, 10, "MONT", MONT_button
    PushButton 85, 195, 20, 10, "FACI", FACI_button
  GroupBox 275, 5, 50, 25, "ELIG panels"
  GroupBox 330, 5, 115, 35, "STAT-based navigation"
  Text 5, 30, 55, 10, "HRF datestamp:"
  Text 125, 30, 40, 10, "HRF status:"
  Text 5, 50, 55, 10, "Earned income:"
  Text 5, 70, 60, 10, "Unearned income:"
  Text 5, 90, 20, 10, "YTD:"
  Text 130, 90, 35, 10, "Changes?:"
  Text 5, 110, 95, 10, "FIAT reasons (if applicable):"
  Text 5, 130, 45, 10, "Other notes:"
  GroupBox 5, 145, 60, 35, "Income panels"
  GroupBox 70, 145, 110, 35, "Asset panels"
  Text 185, 165, 50, 10, "Verifs needed:"
  Text 185, 185, 50, 10, "Actions taken:"
  GroupBox 5, 185, 105, 25, "other STAT panels"
  Text 315, 205, 65, 10, "Worker signature:"
  Text 70, 225, 160, 10, "If post-pay check sent to facility, enter date sent:"
EndDialog


BeginDialog case_note_dialog, 0, 0, 136, 51, "Case note dialog"
  ButtonGroup ButtonPressed
    PushButton 15, 20, 105, 10, "Yes, take me to case note.", yes_case_note_button
    PushButton 5, 35, 125, 10, "No, take me back to the script dialog.", no_case_note_button
  Text 10, 5, 125, 10, "Are you sure you want to case note?"
EndDialog


'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5
Dim row
Dim col


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connecting to BlueZone
EMConnect ""

'Grabbing case number and footer year/month
call MAXIS_case_number_finder(case_number)
call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)


'Showing case number dialog
Do
	Dialog case_number_dialog
	cancel_confirmation
	If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8

'Checking for MAXIS
call check_for_MAXIS(False)

'NAV to MEMB
call navigate_to_MAXIS_screen("stat", "memb")

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Autofilling info for case note
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "MONT", HRF_datestamp)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "FACI", faci)

'Cleaning up info for case note
HRF_computer_friendly_month = footer_month & "/01/" & footer_year
retro_month_name = monthname(datepart("m", (dateadd("m", -2, HRF_computer_friendly_month))))
pro_month_name = monthname(datepart("m", (HRF_computer_friendly_month)))
HRF_month = retro_month_name & "/" & pro_month_name

'checking for an active MAXIS session
Call check_for_MAXIS(False)

'The case note dialog, complete with panel navigation, reading the ELIG/GRH screen, and navigation to case note, as well as logic for certain sections to be required.
Do
	Do
		DO
			Dialog HRF_dialog
			cancel_confirmation
			MAXIS_dialog_navigation
		LOOP UNTIL ButtonPressed = -1
		If HRF_status = "(select one...)" then MsgBox "You must indicate a HRF status, either ''complete'' or ''incomplete''!"
	Loop until HRF_status <> "(select one...)"
	If HRF_status = " " or earned_income = "" or actions_taken = "" or HRF_datestamp = "" or worker_signature = "" then MsgBox "You need to fill in the datestamp, HRF status, earned income, and actions taken sections, as well as sign your case note. Check these items after pressing ''OK''."
Loop until HRF_status <> " " and earned_income <> "" and actions_taken <> "" and HRF_datestamp <> "" and worker_signature <> ""
	If grab_GRH_info_check = 1 then
		call navigate_to_MAXIS_screen("elig", "grh")
		EMReadScreen GRPR_check, 4, 3, 47
		If GRPR_check <> "GRPR" then
			MsgBox "The script couldn't find ELIG/GRH. It will now jump to case note."
		Else
			EMWriteScreen "GRSM", 20, 71
			transmit
			EMReadScreen GRSM_line_01, 69, 10, 3
			EMReadScreen GRSM_line_02, 69, 11, 3
			EMReadScreen GRSM_line_03, 69, 12, 3
			EMReadScreen GRSM_line_04, 69, 13, 3
			EMReadScreen GRSM_Obligation, 9, 18, 31
		End if		
	End if
	

'Enters the case note
Call start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("***" & HRF_month & " HRF received " & HRF_datestamp & ": " & HRF_status & "***")
call write_bullet_and_variable_in_case_note("Earned income", earned_income)
If unearned_income <> "" then call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
If YTD <> "" then call write_bullet_and_variable_in_case_note("YTD", YTD)
If changes <> "" then call write_bullet_and_variable_in_case_note("Changes", changes)
if FIAT_reasons <> "" then call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
if other_notes <> "" then call write_bullet_and_variable_in_case_note("Other notes", other_notes)
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
if verifs_needed <> "" then call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
If date_check_sent_to_facility <> "" then call write_bullet_and_variable_in_case_note("Date check sent to facility", date_check_sent_to_facility)
call write_variable_in_CASE_NOTE("---")
If GRPR_check = "GRPR" then
	call write_variable_in_CASE_NOTE("   " & GRSM_line_01)
	call write_variable_in_CASE_NOTE("   " & GRSM_line_02)
	call write_variable_in_CASE_NOTE("   " & GRSM_line_03)
	call write_variable_in_CASE_NOTE("   " & GRSM_line_04)
	call write_variable_in_CASE_NOTE("Client Obligation: $" & GRSM_Obligation)
	call write_variable_in_CASE_NOTE("---")
End if
call write_variable_in_CASE_NOTE(worker_signature)

call script_end_procedure("")