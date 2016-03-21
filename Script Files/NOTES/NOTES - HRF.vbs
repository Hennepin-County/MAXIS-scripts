'STATS----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - HRF.vbs"
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
STATS_manualtime = 480          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

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
  CheckBox 10, 60, 30, 10, "MFIP", MFIP_check
  CheckBox 55, 60, 35, 10, "SNAP", SNAP_check
  CheckBox 100, 60, 30, 10, "HC", HC_check
  CheckBox 140, 60, 30, 10, "GA", GA_check
  ButtonGroup ButtonPressed
    OkButton 35, 80, 50, 15
    CancelButton 95, 80, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 110, 30, 25, 10, "Year:"
  GroupBox 5, 45, 170, 30, "Programs recertifying"
EndDialog

BeginDialog HRF_dialog, 0, 0, 451, 250, "HRF dialog"
  EditBox 65, 30, 50, 15, HRF_datestamp
  DropListBox 170, 30, 75, 15, " " +chr(9)+"complete"+chr(9)+"incomplete", HRF_status
  EditBox 65, 50, 380, 15, earned_income
  EditBox 70, 70, 375, 15, unearned_income
  EditBox 30, 90, 90, 15, YTD
  EditBox 170, 90, 275, 15, changes
  EditBox 100, 110, 345, 15, FIAT_reasons
  EditBox 50, 130, 395, 15, other_notes
  CheckBox 190, 150, 60, 10, "10% sanction?", ten_percent_sanction_check
  CheckBox 265, 150, 60, 10, "30% sanction?", thirty_percent_sanction_check
  CheckBox 330, 150, 85, 10, "Sent forms to AREP?", sent_arep_checkbox
  EditBox 240, 165, 205, 15, verifs_needed
  EditBox 235, 185, 210, 15, actions_taken
  CheckBox 125, 205, 180, 10, "Check here to case note grant info from ELIG/MFIP.", grab_MFIP_info_check
  CheckBox 125, 220, 170, 10, "Check here to case note grant info from ELIG/FS. ", grab_FS_info_check
  CheckBox 125, 235, 170, 10, "Check here to case note grant info from ELIG/GA.", grab_GA_info_check
  EditBox 380, 205, 65, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 340, 225, 50, 15
    CancelButton 395, 225, 50, 15
    PushButton 260, 20, 20, 10, "FS", ELIG_FS_button
    PushButton 280, 20, 20, 10, "HC", ELIG_HC_button
    PushButton 300, 20, 20, 10, "MFIP", ELIG_MFIP_button
    PushButton 260, 30, 20, 10, "GA", ELIG_GA_button
    PushButton 335, 20, 45, 10, "prev. panel", prev_panel_button
    PushButton 395, 20, 45, 10, "prev. memb", prev_memb_button
    PushButton 335, 30, 45, 10, "next panel", next_panel_button
    PushButton 395, 30, 45, 10, "next memb", next_memb_button
    PushButton 10, 165, 25, 10, "BUSI", BUSI_button
    PushButton 35, 165, 25, 10, "JOBS", JOBS_button
    PushButton 75, 165, 25, 10, "ACCT", ACCT_button
    PushButton 100, 165, 25, 10, "CARS", CARS_button
    PushButton 125, 165, 25, 10, "CASH", CASH_button
    PushButton 150, 165, 25, 10, "OTHR", OTHR_button
    PushButton 10, 175, 25, 10, "RBIC", RBIC_button
    PushButton 35, 175, 25, 10, "UNEA", UNEA_button
    PushButton 75, 175, 25, 10, "REST", REST_button
    PushButton 100, 175, 25, 10, "SECU", SECU_button
    PushButton 125, 175, 25, 10, "TRAN", TRAN_button
    PushButton 10, 210, 25, 10, "MEMB", MEMB_button
    PushButton 35, 210, 25, 10, "MEMI", MEMI_button
    PushButton 60, 210, 25, 10, "MONT", MONT_button
  GroupBox 255, 5, 70, 40, "ELIG panels:"
  GroupBox 330, 5, 115, 40, "STAT-based navigation"
  Text 5, 35, 55, 10, "HRF datestamp:"
  Text 5, 55, 55, 10, "Earned income:"
  Text 5, 75, 60, 10, "Unearned income:"
  Text 5, 95, 20, 10, "YTD:"
  Text 5, 115, 95, 10, "FIAT reasons (if applicable):"
  Text 5, 135, 45, 10, "Other notes:"
  GroupBox 5, 150, 60, 40, "Income panels"
  GroupBox 70, 150, 110, 40, "Asset panels"
  Text 315, 210, 65, 10, "Worker signature:"
  Text 185, 190, 50, 10, "Actions taken:"
  GroupBox 5, 195, 85, 30, "other STAT panels:"
  Text 185, 170, 50, 10, "Verifs needed:"
  Text 125, 35, 40, 10, "HRF status:"
  Text 130, 95, 35, 10, "Changes?:"
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

'Grabbing case number & footer month/year
call MAXIS_case_number_finder(case_number)

'Showing case number dialog
Do
  Dialog case_number_dialog
  cancel_confirmation
  If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then MsgBox "You need to type a valid case number."
Loop until case_number <> "" and IsNumeric(case_number) = True and len(case_number) <= 8

'Checking for an active MAXIS seesion
Call check_for_MAXIS(False)

'NAV to STAT
call navigate_to_MAXIS_screen("stat", "memb")

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Autofilling info for case note
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "MONT", HRF_datestamp)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

'Cleaning up info for case note
HRF_computer_friendly_month = footer_month & "/01/" & footer_year
retro_month_name = monthname(datepart("m", (dateadd("m", -2, HRF_computer_friendly_month))))
pro_month_name = monthname(datepart("m", (HRF_computer_friendly_month)))
HRF_month = retro_month_name & "/" & pro_month_name

'The case note dialog, complete with panel navigation, reading the ELIG/MFIP screen, and navigation to case note, as well as logic for certain sections to be required.
DO
	DO
		Do
			err_msg = ""
			Dialog HRF_dialog
			cancel_confirmation
			Call check_for_password(are_we_passworded_out)   'Adding functionality for MAXIS v.6 Passworded Out issue'
			MAXIS_dialog_navigation
			IF HRF_status = " " AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter a status for your HRF."
			IF HRF_datestamp = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please indicate the date the HRF was received."
			IF earned_income = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please enter information about earned income."
			IF actions_taken = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please indicate which actions you took."
			IF worker_signature = "" AND ButtonPressed = -1 THEN err_msg = err_msg & vbCr & "* Please sign your case note."
			IF err_msg <> "" AND ButtonPressed = -1 THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
		LOOP UNTIL ButtonPressed = -1 AND err_msg = "" AND are_we_passworded_out = False
		case_note_confirmation = MsgBox("Do you want to case note? Press YES to case note. Press NO to return to the previous dialog. Press CANCEL to stop the script.", vbYesNoCancel)
		IF case_note_confirmation = vbCancel THEN script_end_procedure("You have aborted this script.")
	LOOP UNTIL case_note_confirmation = vbYes
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

'grabbing info from elig----------------------------------------------------------------------------------------------------------------------
If grab_MFIP_info_check = 1 then
	call navigate_to_MAXIS_screen("elig", "mfip")
	EMReadScreen MFPR_check, 4, 3, 47
	If MFPR_check <> "MFPR" then
		MsgBox "The script couldn't find ELIG/MFIP. It will now jump to case note."
	Else
		EMWriteScreen "MFSM", 20, 71
		transmit
		EMReadScreen MFSM_line_01, 37, 12, 44
		EMReadScreen MFSM_line_02, 37, 14, 44
		EMReadScreen MFSM_line_03, 37, 15, 44
		EMReadScreen MFSM_line_04, 37, 16, 44
	End if
End if
If grab_FS_info_check = 1 then
	call navigate_to_MAXIS_screen("elig", "fs__")
	EMReadScreen FS_check, 4, 3, 48
	If FS_check <> "FSPR" then 
		MsgBox "The script couldn't find Elig/FS. It will now jump to case note." 
	Else 
		EMWriteScreen "FSSM", 19, 70
		transmit	
		EMReadScreen FS_line_01, 37, 13, 44
	End if
End If	
If grab_GA_info_check = 1 Then
		call navigate_to_MAXIS_screen("ELIG", "GA__")
		EMReadScreen GAPR_check, 4, 3, 48
		IF GAPR_check <> "GAPR" Then
			MsgBox "The script couldn't find Elig/GA. It will now jump to case note." 
		Else
			EMWriteScreen "GASM", 20, 70
			transmit
			EMReadScreen GA_line_01, 10, 14, 70
		END If
END IF



'Creating program list---------------------------------------------------------------------------------------------
If MFIP_check = 1 Then programs_list = "MFIP "
If SNAP_check = 1 Then programs_list = programs_list & "SNAP "
If HC_check = 1 Then programs_list = programs_list & "HC "
If GA_check = 1 Then programs_list = programs_list & "GA "

'Enters the case note-----------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
Call write_variable_in_case_note("***" & HRF_month & " HRF received " & HRF_datestamp & ": " & HRF_status & "***")
call write_bullet_and_variable_in_case_note("Programs", programs_list)
call write_bullet_and_variable_in_case_note("Earned income", earned_income)
call write_bullet_and_variable_in_case_note("Unearned income", unearned_income)
call write_bullet_and_variable_in_case_note("YTD", YTD)
call write_bullet_and_variable_in_case_note("Changes", changes)
call write_bullet_and_variable_in_case_note("FIAT reasons", FIAT_reasons)
call write_bullet_and_variable_in_case_note("Other notes", other_notes)
If ten_percent_sanction_check = 1 then call write_variable_in_CASE_NOTE("* 10% sanction.")
If thirty_percent_sanction_check = 1 then call write_variable_in_CASE_NOTE("* 30% sanction.")
IF Sent_arep_checkbox = checked THEN CALL write_variable_in_case_note("* Sent form(s) to AREP.")
call write_bullet_and_variable_in_case_note("Verifs needed", verifs_needed)
call write_bullet_and_variable_in_case_note("Actions taken", actions_taken)
call write_variable_in_CASE_NOTE("---")
If MFPR_check = "MFPR" then
  call write_variable_in_CASE_NOTE("   " & MFSM_line_01)
  call write_variable_in_CASE_NOTE("   " & MFSM_line_02)
  call write_variable_in_CASE_NOTE("   " & MFSM_line_03)
  call write_variable_in_CASE_NOTE("   " & MFSM_line_04)
  call write_variable_in_CASE_NOTE("---")
End if
If FS_check = "FSPR" then
	call write_variable_in_CASE_NOTE("       FS " & FS_line_01)
	call write_variable_in_CASE_NOTE("---")
End if
If GAPR_check = "GAPR" Then
	call write_variable_in_CASE_NOTE("       GA Benefit Amount............" & GA_line_01)
	call write_variable_in_CASE_NOTE("---")
End If
call write_variable_in_CASE_NOTE(worker_signature)

call script_end_procedure("")
