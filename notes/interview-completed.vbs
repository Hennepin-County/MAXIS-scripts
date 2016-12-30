'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - INTERVIEW COMPLETED.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 90           'manual run time in seconds
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DATE CALCULATIONS----------------------------------------------------------------------------------------------------
MAXIS_footer_month = datepart("m", date)
If len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month
MAXIS_footer_year = "" & datepart("yyyy", date) - 2000

'DIALOGS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 181, 120, "Case number dialog"
  EditBox 80, 5, 70, 15, MAXIS_case_number
  EditBox 65, 25, 30, 15, MAXIS_footer_month
  EditBox 140, 25, 30, 15, MAXIS_footer_year
  CheckBox 10, 60, 30, 10, "cash", cash_checkbox
  CheckBox 50, 60, 30, 10, "HC", HC_checkbox
  CheckBox 90, 60, 35, 10, "SNAP", SNAP_checkbox
  CheckBox 135, 60, 35, 10, "EMER", EMER_checkbox
  DropListBox 70, 80, 75, 15, "Select One..."+chr(9)+"Intake"+chr(9)+"Reapplication"+chr(9)+"Recertification"+chr(9)+"Add program"+chr(9)+"Addendum", CAF_type
  ButtonGroup ButtonPressed
	OkButton 35, 100, 50, 15
	CancelButton 95, 100, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 10, 30, 50, 10, "Footer month:"
  Text 110, 30, 25, 10, "Year:"
  GroupBox 5, 45, 170, 30, "Programs applied for"
  Text 30, 85, 35, 10, "CAF type:"
EndDialog

BeginDialog interview_dialog, 0, 0, 451, 265, "Interview Dialog"
  EditBox 65, 5, 75, 15, caf_datestamp
  EditBox 205, 5, 75, 15, interview_date
  ComboBox 345, 5, 100, 15, "Office"+chr(9)+"Phone", interview_type
  EditBox 45, 25, 400, 15, HH_comp
  EditBox 60, 45, 385, 15, earned_income
  EditBox 75, 65, 370, 15, unearned_income
  EditBox 45, 85, 400, 15, expenses
  EditBox 35, 105, 410, 15, assets
  CheckBox 15, 140, 135, 10, "Check here if this case is expedited.", expedited_checkbox
  EditBox 140, 155, 300, 15, why_xfs
  EditBox 185, 175, 255, 15, reason_expedited_wasnt_processed
  EditBox 50, 200, 395, 15, other_notes
  EditBox 60, 220, 385, 15, verifs_needed
  EditBox 65, 240, 190, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 340, 240, 50, 15
    CancelButton 395, 240, 50, 15
  Text 5, 10, 55, 10, "CAF Datestamp:"
  Text 150, 10, 50, 10, "Interview Date:"
  Text 290, 10, 50, 10, "Interview Type:"
  Text 5, 30, 35, 10, "HH Comp:"
  Text 5, 50, 55, 10, "Earned Income:"
  Text 5, 70, 60, 10, "Unearned Income:"
  Text 5, 90, 35, 10, "Expenses:"
  Text 5, 110, 25, 10, "Assets:"
  GroupBox 5, 125, 440, 70, "Expedited SNAP"
  Text 15, 160, 125, 10, "Explain why case is expedited or not:"
  Text 15, 180, 165, 10, "Reason expedited wasn't processed (if applicable) "
  Text 5, 205, 45, 10, "Other Notes:"
  Text 5, 245, 60, 10, "Worker Signature"
  Text 5, 225, 50, 10, "Verifs Needed:"
EndDialog


'VARIABLES WHICH NEED DECLARING------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
HH_memb_row = 5 'This helps the navigation buttons work!
Dim row
Dim col
application_signed_checkbox = checked 'The script should default to having the application signed.


'GRABBING THE CASE NUMBER, THE MEMB NUMBERS, AND THE FOOTER MONTH------------------------------------------------------------------------------------------------------------------------------------------------
EMConnect ""

call find_variable("Case Nbr: ", MAXIS_case_number, 8)
MAXIS_case_number = trim(MAXIS_case_number)
MAXIS_case_number = replace(MAXIS_case_number, "_", "")
If IsNumeric(MAXIS_case_number) = False then MAXIS_case_number = ""

call find_variable("Month: ", MAXIS_footer_month, 2)
If row <> 0 then
  MAXIS_footer_month = MAXIS_footer_month
  call find_variable("Month: " & MAXIS_footer_month & " ", MAXIS_footer_year, 2)
  If row <> 0 then MAXIS_footer_year = MAXIS_footer_year
End if

MAXIS_case_number = trim(MAXIS_case_number)
MAXIS_case_number = replace(MAXIS_case_number, "_", "")
If IsNumeric(MAXIS_case_number) = False then MAXIS_case_number = ""

Do
  Dialog case_number_dialog 'Runs the first dialog that gathers program information and case number
  cancel_confirmation
  If MAXIS_case_number = "" or IsNumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then MsgBox "You need to type a valid case number."
  If CAF_type = "Select One..." then MsgBox "You must select the type of CAF you interviewed"
Loop until MAXIS_case_number <> "" and IsNumeric(MAXIS_case_number) = True and len(MAXIS_case_number) <= 8 and CAF_type <> "Select One..."
transmit
call check_for_MAXIS(True)


'GRABBING THE DATE RECEIVED AND THE HH MEMBERS---------------------------------------------------------------------------------------------------------------------------------------------------------------------
call navigate_to_MAXIS_screen("stat", "hcre")
EMReadScreen STAT_check, 4, 20, 21
If STAT_check <> "STAT" then script_end_procedure("Can't get in to STAT. This case may be in background. Wait a few seconds and try again. If the case is not in background contact an alpha user for your agency.")


'Creating a custom dialog for determining who the HH members are
call HH_member_custom_dialog(HH_member_array)

'GRABBING THE INFO FOR THE CASE NOTE-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

If CAF_type = "Recertification" then                                                          'For recerts it goes to one area for the CAF datestamp. For other app types it goes to STAT/PROG.
	call autofill_editbox_from_MAXIS(HH_member_array, "REVW", CAF_datestamp)
Else
	call autofill_editbox_from_MAXIS(HH_member_array, "PROG", CAF_datestamp)
End if
IF DateDiff ("d", CAF_datestamp, date) > 60 THEN CAF_datestamp = ""							'This will disregard Application Dates that are older than 60 days. IF and old dste is pulled, the next dialog will require the worker to enter the correct date
If HC_checkbox = checked and CAF_type <> "Recertification" then call autofill_editbox_from_MAXIS(HH_member_array, "HCRE-retro", retro_request)     'Grabbing retro info for HC cases that aren't recertifying
call autofill_editbox_from_MAXIS(HH_member_array, "MEMB", HH_comp)                                                                        'Grabbing HH comp info from MEMB.
If SNAP_checkbox = checked then call autofill_editbox_from_MAXIS(HH_member_array, "EATS", HH_comp)                                                 'Grabbing EATS info for SNAP cases, puts on HH_comp variable
'Removing semicolons from HH_comp variable, it is not needed.
HH_comp = replace(HH_comp, "; ", "")

'MAKING THE GATHERED INFORMATION LOOK BETTER FOR THE CASE NOTE
If cash_checkbox = checked then programs_applied_for = programs_applied_for & "Cash, "
If HC_checkbox = checked then programs_applied_for = programs_applied_for & "HC, "
If SNAP_checkbox = checked then programs_applied_for = programs_applied_for & "SNAP, "
If EMER_checkbox = checked then programs_applied_for = programs_applied_for & "Emergency, "
programs_applied_for = trim(programs_applied_for)
if right(programs_applied_for, 1) = "," then programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

interview_date = date & ""		'Defaults the date of the interview to today's date.

'CASE NOTE DIALOG--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
DO
	Do
		Do
			err_msg = ""
			Dialog interview_dialog			'Displays the Interview Dialog
			cancel_confirmation				'Asks if you're sure you want to cancel, and cancels if you select that.
			If CAF_datestamp = "" or len(CAF_datestamp) > 10 THEN err_msg = "Please enter a valid application datestamp."
			IF worker_signature = "" THEN err_msg = err_msg & vbCr & "Please enter a Worker Signature"
			IF (SNAP_checkbox = checked) AND (why_xfs = "") THEN err_msg = err_msg & vbCr & "SNAP is pending, you must explain your Expedited Determination"
			If err_msg <> "" THEN Msgbox err_msg
		Loop until err_msg = ""
		CALL proceed_confirmation(case_note_confirm)			'Checks to make sure that we're ready to case note.
	Loop until case_note_confirm = TRUE							'Loops until we affirm that we're ready to case note.
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false							'Loops until we affirm that we're ready to case note.

check_for_maxis(FALSE)  'allows for looping to check for maxis after worker has complete dialog box so as not to lose a giant CAF case note if they get timed out while writing.

'Navigates to case note, and checks to make sure we aren't in inquiry.
start_a_blank_CASE_NOTE

'Adding footer month to the recertification case notes
If CAF_type = "Recertification" then CAF_type = MAXIS_footer_month & "/" & MAXIS_footer_year & " Recert"

'THE CASE NOTE-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
CALL write_variable_in_CASE_NOTE("***" & CAF_type & " Interview Completed ***")
CALL write_variable_in_CASE_NOTE ("** Case note for Interview only - full case note of CAF processing to follow.")
IF move_verifs_needed = TRUE THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = True (on FUNCTIONS FILE), it'll case note at the top.
CALL write_bullet_and_variable_in_CASE_NOTE("CAF Datestamp", CAF_datestamp)
CALL write_variable_in_CASE_NOTE("* Interview type: " & interview_type & " - Interview date: " & interview_date)
CALL write_bullet_and_variable_in_CASE_NOTE("Programs applied for", programs_applied_for)
CALL write_bullet_and_variable_in_CASE_NOTE("HH comp/EATS", HH_comp)
CALL write_bullet_and_variable_in_CASE_NOTE("Earned Income", earned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("Unearned Income", unearned_income)
CALL write_bullet_and_variable_in_CASE_NOTE("Expenses", expenses)
CALL write_bullet_and_variable_in_CASE_NOTE("Assets", assets)
'IF application_signed_checkbox = checked THEN 							'Removed this but did not delete in case this functionality is desired for this script
'	CALL write_variable_in_CASE_NOTE("* Application was signed.")
'Else
'	CALL write_variable_in_CASE_NOTE("* Application was not signed.")
'END IF
IF expedited_checkbox = checked THEN CALL write_variable_in_CASE_NOTE("* Expedited SNAP.")
IF (expedited_checkbox = unchecked) AND (SNAP_checkbox = checked) THEN CALL write_variable_in_CASE_NOTE ("* NOT Expedited SNAP")
CALL write_bullet_and_variable_in_CASE_NOTE ("Explanation of Expedited Determination", why_xfs)		'Worker can detail how they arrived at if client is expedited or not - particularly useful if different from screening
CALL write_bullet_and_variable_in_CASE_NOTE("Reason expedited wasn't processed", reason_expedited_wasnt_processed)		'This is strategically placed next to expedited checkbox entry.
CALL write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
IF move_verifs_needed = False THEN CALL write_bullet_and_variable_in_CASE_NOTE("Verifs needed", verifs_needed)			'IF global variable move_verifs_needed = False (on FUNCTIONS FILE), it'll case note at the bottom.
CALL write_variable_in_CASE_NOTE("---")
CALL write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("Success! Interview has been successfully noted. Once processing is completed remember to run the CAF Script for detailed case note.")
