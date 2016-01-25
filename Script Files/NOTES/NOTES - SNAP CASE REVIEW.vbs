'OPTION EXPLICIT
name_of_script = "NOTES - SNAP CASE REVIEW.vbs"
start_time = timer

'declared varaibles for FuncLib
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
STATS_manualtime = 120          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
'END OF stats block=========================================================================================================

'DECLARING VARIABLES--------------------------------------------------------------------------------------------------------
'DIM case_number_dialog
'DIM ButtonPressed
'DIM case_number
'DIM MAXIS_footer_month
'DIM MAXIS_footer_year
'DIM program_droplist

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog SNAP_case_review_dialog, 0, 0, 276, 80, "SNAP case review dialog"
  EditBox 70, 5, 70, 15, case_number					
  DropListBox 190, 5, 80, 15, "Select one..."+chr(9)+"MFIP"+chr(9)+"SNAP", program_droplist
  EditBox 70, 30, 25, 15, MAXIS_footer_month					
  EditBox 100, 30, 25, 15, MAXIS_footer_year
  DropListBox 190, 30, 80, 15, "Select one..."+chr(9)+"correct & approved"+chr(9)+"error exists", SNAP_status_droplist
  EditBox 70, 55, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 165, 55, 50, 15
    CancelButton 220, 55, 50, 15
  Text 5, 35, 65, 10, "Footer month/year:"
  Text 5, 10, 45, 10, "Case number: "
  Text 145, 35, 45, 10, "SNAP status:"
  Text 150, 10, 30, 10, "Program:"
  Text 5, 60, 60, 10, "Worker signature:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""
'Grabs case number
CALL MAXIS_case_number_finder(case_number)
'Grabbing the footer month/year
Call MAXIS_footer_finder (MAXIS_footer_month, MAXIS_footer_year)

'dialog with err_msg.  Alternative to large 'DO LOOP' as user will not be able to leave the loop until err_msg = "" as set at the begining of the dialog
DO 
	err_msg = ""
	Dialog SNAP_case_review_dialog
	If ButtonPressed = 0 then StopScript
	If case_number = "" or IsNumeric(case_number) = False or len(case_number) > 8 then err_msg = err_msg & vbNewLine & "* Enter a valid case number."
	If worker_signature = "" then err_msg = err_msg & vbNewLine & "* Sign your case note."
	If program_droplist = "Select one..." then err_msg = err_msg & vbNewLine & "* You must a program type."
	If MAXIS_footer_month AND MAXIS_footer_year = "" then err_msg = err_msg & vbNewLine & "* You must enter the footer month and footer year."
	IF SNAP_status_droplist = "Select one..." then err_msg = err_msg & vbNewLine & "* You must select a SNAP status type."
	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
LOOP until err_msg = ""


'Checking to make sure user is still in active MAXIS session
Call check_for_MAXIS(FALSE)

If SNAP_status_droplist = "error exists" THEN 	'navigates right to case note, documents error, and ends script.
	start_a_blank_CASE_NOTE
	Call write_variable_in_CASE_NOTE("~~~" & program_droplist & " case review completed, further action required~~~")
	Call write_variable_in_CASE_NOTE("* An error exists in the SNAP budget or issuance.")  
	Call write_variable_in_CASE_NOTE("* The case has been returned to the worker and/or supervisor for correction.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)

	script_end_procedure("A SNAP error exists.  Please communicate this per your agency's procedure.")
END IF


'Navigates to the ELIG results for SNAP, if the worker desires to have the script auto-fills the case note with SNAP approval information.
IF program_droplist = "SNAP" THEN
	call navigate_to_MAXIS_screen("ELIG", "FS")
	EMWriteScreen MAXIS_footer_month, 19, 54
	EMWriteScreen MAXIS_footer_year, 19, 57
	EMReadScreen program_check, 39, 24, 2
	IF program_check = "NO VERSION OF ELIGIBILITY EXISTS FOR FS" THEN
		script_end_procedure("SNAP is NOT active on this case.  Please review the case for accuracy, and take the appropriate action.") 
	ELSE
		EMWRiteScreen "FSSM", 19, 70
		transmit
		EMReadScreen approved_version, 8, 3, 3
		IF approved_version = "APPROVED" THEN
			EMReadScreen snap_bene_amt, 5, 13, 73
			snap_bene_amt = LTrim(snap_bene_amt)
		ELSEIF approved_version <> "APPROVED" THEN
			script_end_procedure("There is an UNAPPROVED version of SNAP for this case.  Please review the case for accuracy, and approve if appropriate.")
		END IF
	END IF
END IF

'Navigates to the ELIG results for MFIP, if the worker desires to have the script auto-fills the case note with MFIP approval information.
IF program_droplist = "MFIP" THEN
	call navigate_to_MAXIS_screen("ELIG", "MFIP")
	EMWriteScreen MAXIS_footer_month, 19, 54
	EMWriteScreen MAXIS_footer_year, 19, 57
	EMReadScreen program_check, 41, 24, 2
	IF program_check = "NO VERSION OF ELIGIBILITY EXISTS FOR MFIP" Then 
		script_end_procedure("MFIP is NOT active on this case.  Please review the case for accuracy, and take the appropriate action.")
	ELSE
		EMWRiteScreen "MFSM", 20, 71
		transmit
		EMReadScreen cash_approved_version, 8, 3, 3
		IF cash_approved_version = "APPROVED" THEN
			EMReadScreen mfip_net_grant_amount, 8, 13, 73
			EMReadScreen mfip_bene_cash_amt, 8, 14, 73
			EMReadScreen mfip_bene_food_amt, 8, 15, 73
			EMReadScreen mfip_bene_housing_amt, 8, 16, 73
			mfip_net_grant_amount = replace(mfip_net_grant_amount, " ", "0")
			mfip_bene_cash_amt = replace(mfip_bene_cash_amt, " ", "0")
			mfip_bene_food_amt = replace(mfip_bene_food_amt, " ", "0")
			mfip_bene_housing_amt = replace(mfip_bene_housing_amt, " ", "0")
		ELSEIF cash_approved_version <> "APPROVED" THEN
			script_end_procedure("There is an UNAPPROVED version of MFIP for this case.  Please review the case for accuracy, and approve if appropriate.")
		END IF
	END IF 
END IF 


'CASE NOTE for correct and approved cases----------------------------------------------------------------------------------------------------
start_a_blank_CASE_NOTE
IF program_droplist = "SNAP" or program_droplist = "EXP SNAP" THEN 
	Call write_variable_in_CASE_NOTE("~~~SNAP case review complete & app'd for " & MAXIS_footer_month & "/" & MAXIS_footer_year & " of " & FormatCurrency(snap_bene_amt) & "~~~")
	Call write_bullet_and_variable_in_CASE_NOTE("SNAP grant amount", FormatCurrency(snap_bene_amt))
	Call write_variable_in_CASE_NOTE("* SNAP case has been reviewed, and the budget and issuance is correct.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
ELSEIF program_droplist = "MFIP" THEN
	Call write_variable_in_CASE_NOTE("~~~MFIP case review complete & app'd for " & MAXIS_footer_month & "/" & MAXIS_footer_year & " of " & FormatCurrency(mfip_net_grant_amount) & "~~~")
	Call write_bullet_and_variable_in_CASE_NOTE("MFIP net grant amount", FormatCurrency(mfip_net_grant_amount))
	call write_bullet_and_variable_in_CASE_NOTE("MFIP Cash portion", FormatCurrency(mfip_bene_cash_amt))
	call write_bullet_and_variable_in_CASE_NOTE("MFIP Food portion",  FormatCurrency(mfip_bene_food_amt))
	call write_bullet_and_variable_in_CASE_NOTE("MFIP Housing portion", FormatCurrency(mfip_bene_housing_amt))
	Call write_variable_in_CASE_NOTE("* MFIP case has been reviewed, and the budget and issuance is correct.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
END IF

script_end_procedure("")
