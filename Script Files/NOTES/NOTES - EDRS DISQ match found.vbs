'OPTION EXPLICIT
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - EDRS DISQ MATCH FOUND.vbs"
start_time = timer

'Variables to be DIMMED for FUNC LIB when testing with Option explicit
'DIM name_of_script, start_time, FuncLib_URL, run_locally, default_directory, beta_agency, req, fso, row

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
STATS_manualtime = 235          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'DECLARING VARIABLES
'DIM case_number_dialog, case_number, edrs_disq_dialog, member_name, edrs_status, worker_signature, edrs_disq_dialog, contact_info, DISQ_state
'DIM contact_date, contact_time, DISQ_reason, DISQ_begin, DISQ_end, DISQ_confirmation, IPV_requested, STAT_DISQ, IPV_TIKL, STATS_counter, STATS_manualtime, STATS_denomination

'DIALOG-------------------------------------------------------------------
BeginDialog edrs_disq_dialog, 0, 0, 296, 415, "eDRS DISQ dialog"
  EditBox 55, 25, 50, 15, case_number
  EditBox 110, 45, 170, 15, HH_memb
  EditBox 110, 65, 170, 15, contact_info
  EditBox 50, 85, 25, 15, DISQ_state
  EditBox 125, 85, 50, 15, contact_date
  EditBox 230, 85, 50, 15, contact_time
  EditBox 65, 105, 215, 15, DISQ_reason
  EditBox 65, 125, 50, 15, DISQ_begin
  EditBox 230, 125, 50, 15, DISQ_end
  CheckBox 5, 150, 255, 10, "Verbal confirmation of DISQ rec'd from DISQ state. SNAP will not be issued.", DISQ_confirmation
  CheckBox 5, 165, 265, 10, "IPV (Intentional Program Violation) documentation requested from DISQ state.", IPV_requested
  CheckBox 5, 180, 155, 10, "STAT/DISQ panel has been added/updated.", STAT_DISQ
  CheckBox 5, 225, 120, 10, "Set 20 day TIKL for return of IPV.", IPV_TIKL
  EditBox 70, 200, 95, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 175, 200, 50, 15
    CancelButton 230, 200, 50, 15
  Text 5, 130, 60, 10, "DISQ begin date: "
  Text 180, 90, 45, 10, "Contact time:"
  Text 5, 110, 60, 10, "Reason for DISQ:"
  Text 130, 130, 100, 10, "DISQ end date (if applicable):"
  GroupBox 0, 10, 290, 185, "Details of contact with DISQ state:"
  GroupBox 0, 245, 290, 165, "Since a match was found several steps need to be taken."
  Text 10, 265, 265, 10, "1. Approve SNAP benefits for other eligible SNAP unit members."
  Text 10, 280, 275, 20, "2. Determine if overpayments for the time period that benefits were paid and disqualified from the SNAP programs."
  Text 10, 305, 275, 20, "3. Check Question 1 in the 'Penalty warnings and qualifications questions' section of all CAFs the client has completed since the disqualification began."
  Text 10, 330, 275, 20, "4. Consider making a fraud referral if the client received food support SNAP or MFIP benefits in Minnesota while disqualified in another state."
  Text 10, 360, 275, 45, "NOTE: If the case has been closed, these steps still need to be completed so the client will not receive any benefits they are not eligible for in the future. If you are unable to update the STAT/DISQ panel because the case has been closed, submit a PF11 with the disqualification screen information and the HelpDesk will enter the information."
  Text 5, 70, 105, 10, "Name/phone of contact person:"
  Text 5, 50, 100, 10, "DISQ household memeber(s):"
  Text 5, 90, 40, 10, "DISQ state:"
  Text 5, 30, 45, 10, "Case number:"
  Text 120, 25, 160, 10, "** Procedure for eDRS DISQ per TE02.08.127**"
  Text 80, 90, 45, 10, "Contact date:"
  Text 5, 205, 60, 10, "Worker signature:"
EndDialog

'THE SCRIPT------------------------------------------------------------------------------------------
'Connecting to BlueZone & finding case number
EMConnect ""
call MAXIS_case_number_finder(case_number)
Call check_for_MAXIS(False)	'checking for an active MAXIS session'

'updates the contact_date & contact_time variables to show the current date & time
contact_date = date & ""
contact_time = time & ""

Do			'edrs status dialog
	err_msg = ""			'establishes a blank variable for the DO LOOP
	dialog edrs_disq_dialog
	cancel_confirmation
	If HH_memb = "" 																					then err_msg = err_msg & vbNewLine & "* You must enter the disqualified HH member(s)."
	If contact_info = ""																			then err_msg = err_msg & vbNewLine & "* You must enter the contact name and phone number."
	IF DISQ_state	= "" 																				then err_msg = err_msg & vbNewLine & "* You must enter the state of disqualification."
	If IsDate(contact_date) = False 													then err_msg = err_msg & vbNewLine & "* You must enter a the contact date."
	If contact_time = ""																			then err_msg = err_msg & vbNewLine & "* You must enter a the contact time."
	If DISQ_reason = "" 																			then err_msg = err_msg & vbNewLine & "* You must enter the disqualification reason."
	If DISQ_begin = "" or IsDate(DISQ_begin) = FALSE 					then err_msg = err_msg & vbNewLine & "* You must enter a valid date for the DISQ begin date."
	If DISQ_end <> "" And IsDate(DISQ_end) = FALSE 						then err_msg = err_msg & vbNewLine & "* You must enter a valid date for the DISQ end date."
	If IsNumeric(case_number) = False or Len(case_number) > 8 then err_msg = err_msg & vbNewLine & "* You must enter a valid case number."
	If worker_signature = "" 																	then err_msg = err_msg & vbNewLine & "* You must sign the case note."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
Loop until err_msg = ""

'checking for active MAXIS session
Call check_for_MAXIS(False)

'TIKL for return of IPV if option is selected by user
If IPV_TIKL = 1 then
	call navigate_to_MAXIS_screen("dail", "writ")
	call create_MAXIS_friendly_date(contact_date, 20, 5, 18)
	EMSetCursor 9, 3
	Call write_variable_in_TIKL("Has IPV from " & (DISQ_state) & " been received?  If not, contact " & (DISQ_state) & " again and case note. Refer to TE02.08.127 for procedural information.")
	transmit
	PF3
End if

'Case noting & navigating to a new case note
Call start_a_blank_CASE_NOTE
Call write_variable_in_CASE_NOTE("***eDRS completed/DISQ match found***")
Call write_bullet_and_variable_in_CASE_NOTE("Household members with DISQ match", HH_memb)
Call write_bullet_and_variable_in_CASE_NOTE("Name/phone number of DISQ contact", contact_info)
Call write_bullet_and_variable_in_CASE_NOTE("DISQ State", DISQ_state)
Call write_bullet_and_variable_in_CASE_NOTE("Contact date", contact_date)
Call write_bullet_and_variable_in_CASE_NOTE("Contact time", contact_time)
call write_bullet_and_variable_in_CASE_NOTE("Reason for DISQ", DISQ_reason)
Call write_bullet_and_variable_in_CASE_NOTE("DISQ begin date", DISQ_begin)
Call write_bullet_and_variable_in_CASE_NOTE("DISQ end date", DISQ_end)
Call write_variable_in_CASE_NOTE("---")
If DISQ_confirmation = 1 then call write_variable_in_CASE_NOTE("* Verbal confirmation of DISQ rec'd from " & DISQ_state & ". SNAP will not be issued.")
IF IPV_requested = 1 then call write_variable_in_CASE_NOTE("* IPV (Intentional Program Violation) documentation requested from " & DISQ_state & ".")
IF STAT_DISQ = 1 then Call write_variable_in_CASE_NOTE("* STAT/DISQ panel has been added/updated.")
IF IPV_TIKL = 1 then Call write_variable_in_CASE_NOTE("* A TIKL has been set for 20 days for the return of IPV.")
call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("Please refer to TE02.08.127 in POLI/TEMP for additional procedural information if you have any questions.")
