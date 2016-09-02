'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - HRF.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 480          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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

'defaulting the MAXIS footer month/year to current month plus one
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 181, 100, "Case number dialog"
  EditBox 80, 5, 70, 15, MAXIS_case_number
  EditBox 65, 25, 30, 15, MAXIS_footer_month
  EditBox 140, 25, 30, 15, MAXIS_footer_year
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

BeginDialog HRF_dialog, 0, 0, 451, 270, "HRF dialog"
  EditBox 65, 30, 50, 15, HRF_datestamp
  DropListBox 170, 30, 75, 15, " "+chr(9)+"complete"+chr(9)+"incomplete", HRF_status
  EditBox 65, 50, 380, 15, earned_income
  EditBox 70, 70, 375, 15, unearned_income
  EditBox 30, 90, 90, 15, YTD
  EditBox 170, 90, 275, 15, changes
  EditBox 30, 110, 415, 15, EMPS
  EditBox 100, 130, 345, 15, FIAT_reasons
  EditBox 50, 150, 395, 15, other_notes
  CheckBox 190, 170, 60, 10, "10% sanction?", ten_percent_sanction_check
  CheckBox 265, 170, 60, 10, "30% sanction?", thirty_percent_sanction_check
  CheckBox 330, 170, 85, 10, "Sent forms to AREP?", sent_arep_checkbox
  EditBox 235, 185, 210, 15, verifs_needed
  EditBox 235, 205, 210, 15, actions_taken
  CheckBox 100, 225, 175, 10, "Check here to case note grant info from ELIG/MFIP.", grab_MFIP_info_check
  CheckBox 100, 240, 170, 10, "Check here to case note grant info from ELIG/FS. ", grab_FS_info_check
  CheckBox 100, 255, 170, 10, "Check here to case note grant info from ELIG/GA.", grab_GA_info_check
  EditBox 340, 225, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 340, 245, 50, 15
    CancelButton 395, 245, 50, 15
    PushButton 260, 20, 20, 10, "FS", ELIG_FS_button
    PushButton 280, 20, 20, 10, "HC", ELIG_HC_button
    PushButton 300, 20, 20, 10, "MFIP", ELIG_MFIP_button
    PushButton 260, 30, 20, 10, "GA", ELIG_GA_button
    PushButton 335, 20, 45, 10, "prev. panel", prev_panel_button
    PushButton 395, 20, 45, 10, "prev. memb", prev_memb_button
    PushButton 335, 30, 45, 10, "next panel", next_panel_button
    PushButton 395, 30, 45, 10, "next memb", next_memb_button
    PushButton 5, 115, 25, 10, "EMPS", EMPS_button
    PushButton 10, 185, 25, 10, "BUSI", BUSI_button
    PushButton 35, 185, 25, 10, "JOBS", JOBS_button
    PushButton 10, 195, 25, 10, "RBIC", RBIC_button
    PushButton 35, 195, 25, 10, "UNEA", UNEA_button
    PushButton 75, 185, 25, 10, "ACCT", ACCT_button
    PushButton 100, 185, 25, 10, "CARS", CARS_button
    PushButton 125, 185, 25, 10, "CASH", CASH_button
    PushButton 150, 185, 25, 10, "OTHR", OTHR_button
    PushButton 75, 195, 25, 10, "REST", REST_button
    PushButton 100, 195, 25, 10, "SECU", SECU_button
    PushButton 125, 195, 25, 10, "TRAN", TRAN_button
    PushButton 10, 230, 25, 10, "MEMB", MEMB_button
    PushButton 35, 230, 25, 10, "MEMI", MEMI_button
    PushButton 60, 230, 25, 10, "MONT", MONT_button
    PushButton 10, 240, 25, 10, "PARE", PARE_button
    PushButton 35, 240, 25, 10, "SANC", SANC_button
    PushButton 60, 240, 25, 10, "TIME", TIME_button
  Text 5, 95, 20, 10, "YTD:"
  Text 5, 135, 95, 10, "FIAT reasons (if applicable):"
  Text 5, 155, 45, 10, "Other notes:"
  GroupBox 5, 170, 60, 40, "Income panels"
  GroupBox 70, 170, 110, 40, "Asset panels"
  Text 280, 230, 60, 10, "Worker signature:"
  Text 185, 210, 50, 10, "Actions taken:"
  GroupBox 5, 215, 85, 40, "other STAT panels:"
  Text 185, 190, 50, 10, "Verifs needed:"
  Text 125, 35, 40, 10, "HRF status:"
  Text 130, 95, 35, 10, "Changes?:"
  GroupBox 330, 5, 115, 40, "STAT-based navigation"
  Text 5, 35, 55, 10, "HRF datestamp:"
  Text 5, 55, 55, 10, "Earned income:"
  Text 5, 75, 60, 10, "Unearned income:"
  GroupBox 255, 5, 70, 40, "ELIG panels:"
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
call MAXIS_case_number_finder(MAXIS_case_number)

'Showing case number dialog
do 
	Do
  		Dialog case_number_dialog
  		cancel_confirmation
  		If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then MsgBox "You need to type a valid case number."
	Loop until MAXIS_case_number <> "" and IsNumeric(MAXIS_case_number) = True and len(MAXIS_case_number) <= 8
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					

'Checking for an active MAXIS seesion
Call check_for_MAXIS(False)

'NAV to STAT
call navigate_to_MAXIS_screen("stat", "memb")

'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'Autofilling info for case note
call autofill_editbox_from_MAXIS(HH_member_array, "BUSI", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "EMPS", EMPS)
call autofill_editbox_from_MAXIS(HH_member_array, "JOBS", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "MONT", HRF_datestamp)
call autofill_editbox_from_MAXIS(HH_member_array, "RBIC", earned_income)
call autofill_editbox_from_MAXIS(HH_member_array, "UNEA", unearned_income)

'Cleaning up info for case note
HRF_computer_friendly_month = MAXIS_footer_month & "/01/" & MAXIS_footer_year
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
call write_bullet_and_variable_in_case_note("EMPS", EMPS)
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

script_end_procedure("")