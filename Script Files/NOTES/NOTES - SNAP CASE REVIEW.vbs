'OPTION EXPLICIT

name_of_script = "NOTES - SNAP CASE REVIEW.vbs"
start_time = timer

'DIM name_of_script
'DIM start_time
'DIM FuncLib_URL
'DIM run_locally
'DIM default_directory
'DIM beta_agency
'DIM req
'DIM fso
'DIM row

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
'END OF GLOBAL VARIABLES----------------------------------------------------------------------------------------------------

'FUNCTION----------------------------------------------------------------------------------------------------
FUNCTION MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)'Grabbing the footer month/year
	back_to_self
    Call find_variable("Benefit Period (MM YY): ", MAXIS_footer_month, 2)
    If isnumeric(MAXIS_footer_month) = true then               'checking to see if a footer month 'number' is present 
    footer_month = MAXIS_footer_month                
    call find_variable("Benefit Period (MM YY): " & footer_month & " ", MAXIS_footer_year, 2)
    If isnumeric(MAXIS_footer_year) = true then footer_year = MAXIS_footer_year 'checking to see if a footer year 'number' is present
	Else 'If we don’t have one found, we’re going to assign the current month/year.
		MAXIS_footer_month = DatePart("m", date)   'Datepart delivers the month number to the variable
		If len(MAXIS_footer_month) = 1 then MAXIS_footer_month = "0" & MAXIS_footer_month   'If it’s a single digit month, add a zero
		MAXIS_footer_year = right(DatePart("yyyy", date), 2)   'We only need the right two characters of the year for MAXIS
	End if
END FUNCTION

'DECLARING VARIABLES--------------------------------------------------------------------------------------------------------
'DIM SNAP_quality_case_review_dialog
'DIM ButtonPressed
'DIM case_number
'DIM MAXIS_footer_month
'DIM MAXIS_footer_year
'DIM SNAP_status
'DIM grant_amount
'DIM worker_signature
'DIM footer_month
'DIM footer_year


'DIALOG----------------------------------------------------------------------------------------------------
BeginDialog case_number_dialog, 0, 0, 131, 90, "Case number dialog"
  EditBox 70, 5, 55, 15, case_number					
  EditBox 70, 25, 25, 15, MAXIS_footer_month					
  EditBox 100, 25, 25, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 20, 70, 50, 15
    CancelButton 75, 70, 50, 15
  Text 5, 30, 62, 10, "Footer month/year:"
  Text 10, 10, 45, 10, "Case number: "
  Text 10, 50, 30, 10, "Program:"
  DropListBox 70, 45, 55, 15, "Select one..."+chr(9)+"EXP SNAP"+chr(9)+"MFIP"+chr(9)+"SNAP", program_droplist
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""
'Grabs case number
CALL MAXIS_case_number_finder(case_number)
'Grabbing the footer month/year
Call MAXIS_footer_finder (MAXIS_footer_month, MAXIS_footer_year)


DO
	DO
		Do
			DO
				err
				Dialog case_number_dialog
				If ButtonPressed = 0 then StopScript
				IF IsNumeric(case_number) = FALSE THEN MsgBox "You must type a valid case number"
			LOOP UNTIL IsNumeric(case_number) = TRUE
			If worker_signature = "" THEN MsgBox "You must sign the case note."
		LOOP until worker_signature <> ""
		If SNAP_status = "Select one..." THEN MsgBox "You must check either that the case is correct and approved, or an error exists."
	LOOP UNTIL SNAP_status <> "Select one..."
	If (SNAP_status = "correct & approved" AND grant_amount = "") OR (SNAP_status = "error exists" AND grant_amount <> "") THEN Msgbox "You must either select 'error exists', and leave the grant amount blank OR select 'correct & approved', and enter the grant amount. "
LOOP until (SNAP_status = "correct & approved" AND grant_amount <> "") OR (SNAP_status = "error exists" AND grant_amount = "") 	


'Dollar bill symbol will be added to numeric variables (in grant_amount)
IF grant_amount <> "" THEN grant_amount = "$" & grant_amount

'Checking to make sure user is still in active MAXIS session
Call check_for_MAXIS(TRUE)

'The CASE NOTE----------------------------------------------------------------------------------------------------
'navigates to case note and creates a new one
Call start_a_blank_CASE_NOTE
'Case note if case is incorrect
If SNAP_status = "error exists" THEN
	Call write_variable_in_CASE_NOTE("~~~SNAP case review complete, further action required~~~")
	Call write_variable_in_CASE_NOTE("* An error exists in the SNAP budget or issuance.")  
	Call write_variable_in_CASE_NOTE("* The case has been returned to the worker and supervisor for correction.")
	Call write_variable_in_CASE_NOTE("---")
	Call write_variable_in_CASE_NOTE(worker_signature)
	'Case note if case is correct
	ELSEIF SNAP_status = "correct & approved" THEN 
		Call write_variable_in_CASE_NOTE("~~~SNAP case review complete & app'd for " & MAXIS_footer_month & "/" & MAXIS_footer_year & " of " & grant_amount & " SNAP grant~~~")
		Call write_variable_in_CASE_NOTE("* SNAP case has been reviewed, and the budget and issuance is correct.")
		Call write_variable_in_CASE_NOTE("---")
		Call write_variable_in_CASE_NOTE(worker_signature)	
END If

script_end_procedure("")

'Navigates to the ELIG results for SNAP, if the worker desires to have the script autofill the case note with SNAP approval information.
IF program_droplist = "SNAP" or program_droplist = "EXP SNAP" THEN
	snap_month = int(snap_start_mo)
	snap_year = int(snap_start_yr)
	snap_count = 0
	DO
		IF len(snap_month) = 1 THEN snap_month = "0" & snap_month
		call navigate_to_screen("ELIG", "FS")
		EMWriteScreen snap_month, 19, 54
		EMWriteScreen snap_year, 19, 57
		EMWRiteScreen "FSSM", 19, 70
		transmit
		EMReadScreen approved_version, 8, 3, 3
		IF approved_version = "APPROVED" THEN
			EMReadScreen approval_date, 8, 3, 14
			approval_date = cdate(approval_date)
			IF approval_date = date THEN
				EMReadScreen snap_bene_amt, 5, 13, 73
				EMReadScreen current_snap_bene_mo, 2, 19, 54
				EMReadScreen current_snap_bene_yr, 2, 19, 57
				EMReadScreen snap_reporter, 10, 8, 31
				snap_bene_amt = replace(snap_bene_amt, ",", "")
				snap_bene_amt = replace(snap_bene_amt, " ", "0")
				snap_reporter = replace(snap_reporter, " ", "")
				IF len(snap_bene_amt) = 5 THEN snap_bene_amt = right(snap_bene_amt, 4)
				snap_approval_array = snap_approval_array & snap_bene_amt & snap_reporter & current_snap_bene_mo & current_snap_bene_yr & " "
			ELSE
				script_end_procedure("Your most recent SNAP approval for the benefit month chosen is not from today. The script cannot autofill this result. Process manually.")
			END IF
		ELSE
			EMReadScreen approval_versions, 2, 2, 18
			IF trim(approval_versions) = "1" THEN script_end_procedure("You do not have an approved version of SNAP in the selected benefit month. Please approve before running the script.")
			approval_versions = approval_versions * 1
			approval_to_check = approval_versions - 1
			EMWriteScreen approval_to_check, 19, 78
			transmit
			EMReadScreen approval_date, 8, 3, 14
			approval_date = cdate(approval_date)
			IF approval_date = date THEN
				EMReadScreen snap_bene_amt, 5, 13, 73
				EMReadScreen current_snap_bene_mo, 2, 19, 54
				EMReadScreen current_snap_bene_yr, 2, 19, 57
				EMReadScreen snap_reporter, 10, 8, 31
				snap_bene_amt = replace(snap_bene_amt, ",", "")
				snap_bene_amt = replace(snap_bene_amt, " ", "0")
				snap_reporter = replace(snap_reporter, " ", "")
				IF len(snap_bene_amt) = 5 THEN snap_bene_amt = right(snap_bene_amt, 4)
				snap_approval_array = snap_approval_array & snap_bene_amt & snap_reporter & current_snap_bene_mo & current_snap_bene_yr & " "
			ELSE
				script_end_procedure("Your most recent SNAP approval for the benefit month chosen is not from today. The script cannot autofill this result. Process manually.")
			END IF
		END IF	
		snap_month = snap_month + 1
		IF snap_month = 13 THEN
			snap_month = 1
			snap_year = snap_year + 1
		END IF
		snap_count = snap_count + 1
	LOOP UNTIL snap_count = total_snap_months
END IF

snap_approval_array = trim(snap_approval_array)
snap_approval_array = split(snap_approval_array)

'----------This version only autofills CASH.----------
IF autofill_cash_check = checked THEN
	cash_month = int(cash_start_mo)
	IF len(cash_month) = 1 THEN cash_month = "0" & cash_month
	cash_year = int(cash_start_yr)
	cash_count = 0

	DO
		IF len(cash_month) = 1 THEN cash_month = "0" & cash_month
		call navigate_to_screen("ELIG", "SUMM")
		EMWriteScreen cash_month, 19, 56
		EMWriteScreen cash_year, 19, 59
		transmit

		EMReadScreen dwp_elig_summ, 1, 7, 40
		EMReadScreen mfip_elig_summ, 1, 8, 40
		EMReadScreen msa_elig_summ, 1, 11, 40
		EMReadScreen ga_elig_summ, 1, 12, 40

		
		IF mfip_elig_summ <> " " THEN
			EMReadScreen date_of_last_MFIP_version, 8, 8, 48
'			IF date_of_last_MFIP_version = "11/06/14" THEN
			prog_to_check_array = prog_to_check_array & "MF" & cash_month & cash_year & "/"
		END IF
	
		IF dwp_elig_summ = " " AND mfip_elig_summ = " " AND msa_elig_summ = " " AND ga_elig_summ = " " THEN prog_to_check_array = prog_to_check_array & "NO" & cash_month & cash_year & "/"
		
		cash_month = cash_month + 1
		IF cash_month = 13 THEN
			cash_month = 1
			cash_year = cash_year + 1
		END IF
		cash_count = cash_count + 1
	LOOP UNTIL cash_count = total_cash_months

		prog_to_check_array = trim(prog_to_check_array)
		prog_to_check_array = split(prog_to_check_array, "/")


		FOR EACH prog_to_check IN prog_to_check_array

			IF left(prog_to_check, 2) = "NO" THEN
				MsgBox "There are no CASH result found."

			ELSEIF left(prog_to_check, 2) = "MF" THEN
				mfip_housing_start_date = #07/01/2015#
				'MFIP portion
				call navigate_to_screen("ELIG", "MFIP")
				EMWriteScreen left(right(prog_to_check, 4), 2), 20, 56
				EMWriteScreen right(prog_to_check, 2), 20, 59
				EMWRiteScreen "MFSM", 20, 71
				transmit
				EMReadScreen cash_approved_version, 8, 3, 3
				IF cash_approved_version = "APPROVED" THEN
					EMReadScreen cash_approval_date, 8, 3, 14
					IF cdate(cash_approval_date) = date THEN
						EMReadScreen current_cash_bene_mo, 2, 20, 55
						EMReadScreen current_cash_bene_yr, 2, 20, 58
						current_cash_month = current_cash_bene_mo & "/01/" & current_cash_bene_yr
						'Determining the benefit month so that script knows whether or not to be looking for the MFIP housing grant.
						'If the benefit month is 07/15 or later, it will read the housing grant...
						IF DateDiff("D", mfip_housing_start_date, current_cash_month) >= 0 THEN 
							EMReadScreen mfip_bene_cash_amt, 8, 14, 73
							EMReadScreen mfip_bene_food_amt, 8, 15, 73
							EMReadScreen mfip_bene_housing_amt, 8, 16, 73
							mfip_bene_cash_amt = replace(mfip_bene_cash_amt, " ", "0")
							mfip_bene_food_amt = replace(mfip_bene_food_amt, " ", "0")
							mfip_bene_housing_amt = replace(mfip_bene_housing_amt, " ", "0")
							cash_approval_array = cash_approval_array & "MFIP" & mfip_bene_cash_amt & mfip_bene_food_amt & mfip_bene_housing_amt & current_cash_bene_mo & current_cash_bene_yr & " "
						ELSEIF DateDiff("D", mfip_housing_start_date, current_cash_month) < 0 THEN 
							EMReadScreen mfip_bene_cash_amt, 8, 15, 73
							EMReadScreen mfip_bene_food_amt, 8, 16, 73
							mfip_bene_cash_amt = replace(mfip_bene_cash_amt, " ", "0")
							mfip_bene_food_amt = replace(mfip_bene_food_amt, " ", "0")
							cash_approval_array = cash_approval_array & "MFIP" & mfip_bene_cash_amt & mfip_bene_food_amt & current_cash_bene_mo & current_cash_bene_yr & " "
						END IF
					END IF
				ELSE
					EMReadScreen cash_approval_versions, 1, 2, 18
					IF cash_approval_versions = "1" THEN script_end_procedure("You do not have an approved version of CASH in the selected benefit month. Please approve before running the script.")
					cash_approval_versions = int(cash_approval_versions)
					cash_approval_to_check = cash_approval_versions - 1
					EMWriteScreen cash_approval_to_check, 20, 79
					transmit
					EMReadScreen cash_approval_date, 8, 3, 14
					IF cdate(cash_approval_date) = date THEN
						EMReadScreen current_cash_bene_mo, 2, 20, 55
						EMReadScreen current_cash_bene_yr, 2, 20, 58
						current_cash_month = current_cash_bene_mo & "/01/" & current_cash_bene_yr
						'Determining the benefit month so that script knows whether or not to be looking for the MFIP housing grant.
						'If the benefit month is 07/15 or later, it will read the housing grant...
						IF DateDiff("D", mfip_housing_start_date, current_cash_month) >= 0 THEN 
							EMReadScreen mfip_bene_cash_amt, 8, 14, 73
							EMReadScreen mfip_bene_food_amt, 8, 15, 73
							EMReadScreen mfip_bene_housing_amt, 8, 16, 73
							mfip_bene_cash_amt = replace(mfip_bene_cash_amt, " ", "0")
							mfip_bene_food_amt = replace(mfip_bene_food_amt, " ", "0")
							mfip_bene_housing_amt = replace(mfip_bene_housing_amt, " ", "0")
							cash_approval_array = cash_approval_array & "MFIP" & mfip_bene_cash_amt & mfip_bene_food_amt & mfip_bene_housing_amt & current_cash_bene_mo & current_cash_bene_yr & " "
						ELSEIF DateDiff("D", mfip_housing_start_date, current_cash_month) < 0 THEN 
							EMReadScreen mfip_bene_cash_amt, 8, 15, 73
							EMReadScreen mfip_bene_food_amt, 8, 16, 73
							mfip_bene_cash_amt = replace(mfip_bene_cash_amt, " ", "0")
							mfip_bene_food_amt = replace(mfip_bene_food_amt, " ", "0")
							cash_approval_array = cash_approval_array & "MFIP" & mfip_bene_cash_amt & mfip_bene_food_amt & current_cash_bene_mo & current_cash_bene_yr & " "
						END IF
					END IF
				END IF	
			ELSEIF left(prog_to_check, 2) = "GA" THEN
				'GA portion
				call navigate_to_screen("ELIG", "GA")
				EMWriteScreen left(right(prog_to_check, 4), 2), 20, 54
				EMWriteScreen right(prog_to_check, 2), 20, 57
				EMWRiteScreen "GASM", 20, 70
				transmit
				EMReadScreen cash_approved_version, 8, 3, 3
				IF cash_approved_version = "APPROVED" THEN
					EMReadScreen cash_approval_date, 8, 3, 15
					IF cdate(cash_approval_date) = date THEN
						EMReadScreen GA_bene_cash_amt, 8, 14, 72
						EMReadScreen current_cash_bene_mo, 2, 20, 54
						EMReadScreen current_cash_bene_yr, 2, 20, 57
						GA_bene_cash_amt = replace(GA_bene_cash_amt, " ", "0")
						cash_approval_array = cash_approval_array & "GA__" & GA_bene_cash_amt & current_cash_bene_mo & current_cash_bene_yr & " "
					END IF
				ELSE
					EMReadScreen cash_approval_versions, 1, 2, 18
					IF cash_approval_versions = "1" THEN script_end_procedure("You do not have an approved version of CASH in the selected benefit month. Please approve before running the script.")
					cash_approval_versions = int(cash_approval_versions)
					cash_approval_to_check = cash_approval_versions - 1
					EMWriteScreen cash_approval_to_check, 20, 79
					transmit
					EMReadScreen cash_approval_date, 8, 3, 15
					IF cdate(cash_approval_date) = date THEN
						EMReadScreen GA_bene_cash_amt, 8, 14, 72
						EMReadScreen current_cash_bene_mo, 2, 20, 54
						EMReadScreen current_cash_bene_yr, 2, 20, 57
						GA_bene_cash_amt = replace(GA_bene_cash_amt, " ", "0")
						cash_approval_array = cash_approval_array & "GA__" & GA_bene_cash_amt & current_cash_bene_mo & current_cash_bene_yr & " "
					END IF
				END IF
		
			ELSEIF left(prog_to_check, 2) = "MS" THEN
				'MSA portion
				call navigate_to_screen("ELIG", "MSA")
				EMWriteScreen left(right(prog_to_check, 4), 2), 20, 56
				EMWriteScreen right(prog_to_check, 2), 20, 59
				EMWRiteScreen "MSSM", 20, 71
				transmit
				EMReadScreen cash_approved_version, 8, 3, 3
				IF cash_approved_version = "APPROVED" THEN
					EMReadScreen cash_approval_date, 8, 3, 14
					IF cdate(cash_approval_date) = date THEN
						EMReadScreen MSA_bene_cash_amt, 8, 17, 73
						EMReadScreen current_cash_bene_mo, 2, 20, 54
						EMReadScreen current_cash_bene_yr, 2, 20, 57
						MSA_bene_cash_amt = replace(MSA_bene_cash_amt, " ", "0")
						cash_approval_array = cash_approval_array & "MSA_" & MSA_bene_cash_amt & current_cash_bene_mo & current_cash_bene_yr & " "
					END IF
				ELSE
					EMReadScreen cash_approval_versions, 1, 2, 18
					IF cash_approval_versions = "1" THEN script_end_procedure("You do not have an approved version of CASH in the selected benefit month. Please approve before running the script.")
					cash_approval_versions = int(cash_approval_versions)
					cash_approval_to_check = cash_approval_versions - 1
					EMWriteScreen cash_approval_to_check, 20, 79
					transmit
					EMReadScreen cash_approval_date, 8, 3, 14
					IF cdate(cash_approval_date) = date THEN
						EMReadScreen MSA_bene_cash_amt, 8, 17, 73
						EMReadScreen current_cash_bene_mo, 2, 20, 54
						EMReadScreen current_cash_bene_yr, 2, 20, 57
						MSA_bene_cash_amt = replace(MSA_bene_cash_amt, " ", "0")
						cash_approval_array = cash_approval_array & "MSA_" & MSA_bene_cash_amt & current_cash_bene_mo & current_cash_bene_yr & " "
					END IF
				END IF
			ELSEIF left(prog_to_check, 2) = "DW" THEN
				'DWP portion
				call navigate_to_screen("ELIG", "DWP")
				EMWriteScreen left(right(prog_to_check, 4), 2), 20, 56
				EMWriteScreen right(prog_to_check, 2), 20, 59
				EMWRiteScreen "DWSM", 20, 71
				transmit
				EMReadScreen cash_approved_version, 8, 3, 3
				IF cash_approved_version = "APPROVED" THEN
					EMReadScreen cash_approval_date, 8, 3, 14
					IF cdate(cash_approval_date) = date THEN
						EMReadScreen DWP_bene_shel_amt, 8, 13, 73
						EMReadScreen DWP_bene_pers_amt, 8, 14, 73
						EMReadScreen current_cash_bene_mo, 2, 20, 56
						EMReadScreen current_cash_bene_yr, 2, 20, 59
						DWP_bene_shel_amt = replace(DWP_bene_shel_amt, " ", "0")
						DWP_bene_pers_amt = replace(DWP_bene_pers_amt, " ", "0")
						cash_approval_array = cash_approval_array & "DWP_" & DWP_bene_shel_amt & DWP_bene_pers_amt & current_cash_bene_mo & current_cash_bene_yr & " "
					END IF
				ELSE
					EMReadScreen cash_approval_versions, 1, 2, 18
					IF cash_approval_versions = "1" THEN script_end_procedure("You do not have an approved version of CASH in the selected benefit month. Please approve before running the script.")
					cash_approval_versions = int(cash_approval_versions)
					cash_approval_to_check = cash_approval_versions - 1
					EMWriteScreen cash_approval_to_check, 20, 79
					transmit
					EMReadScreen cash_approval_date, 8, 3, 14
					IF cdate(cash_approval_date) = date THEN
						EMReadScreen DWP_bene_shel_amt, 8, 13, 73
						EMReadScreen DWP_bene_pers_amt, 8, 14, 73
						EMReadScreen current_cash_bene_mo, 2, 20, 56
						EMReadScreen current_cash_bene_yr, 2, 20, 59
						DWP_bene_shel_amt = replace(DWP_bene_shel_amt, " ", "0")
						DWP_bene_pers_amt = replace(DWP_bene_pers_amt, " ", "0")
						cash_approval_array = cash_approval_array & "DWP_" & DWP_bene_shel_amt & DWP_bene_pers_amt & current_cash_bene_mo & current_cash_bene_yr & " "
					END IF
				END IF
			END IF
		NEXT
END IF



cash_approval_array = trim(cash_approval_array)
cash_approval_array = split(cash_approval_array)


'Case notes----------------------------------------------------------------------------------------------------
call start_a_blank_CASE_NOTE

IF snap_approved_check = checked THEN approved_programs = approved_programs & "SNAP/"
IF hc_approved_check = checked THEN approved_programs = approved_programs & "HC/"
IF cash_approved_check = checked THEN approved_programs = approved_programs & "CASH/"
IF emer_approved_check = checked THEN approved_programs = approved_programs & "EMER/"
EMSendKey "---Approved " & approved_programs & "<backspace>" & " " & type_of_approval & "---" & "<newline>"
IF benefit_breakdown <> "" THEN call write_bullet_and_variable_in_CASE_NOTE("Benefit Breakdown", benefit_breakdown)
IF autofill_snap_check = checked THEN
	FOR EACH snap_approval_result in snap_approval_array
		bene_amount = left(snap_approval_result, 4)
		report_status = " " & MID(snap_approval_result, 5)
		len_report_status = LEN(report_status)
		report_status = LEFT(report_status, (len_report_status - 4)) & " Reporter"
		benefit_month = left(right(snap_approval_result, 4), 2)
		benefit_year = right(snap_approval_result, 2)
		snap_header = ("SNAP for " & benefit_month & "/" & benefit_year)
		call write_bullet_and_variable_in_CASE_NOTE(snap_header, FormatCurrency(bene_amount) & report_status)
	NEXT
END IF
IF program_droplist = "MFIP" THEN
	FOR EACH cash_approval_result IN cash_approval_array
		IF left(cash_approval_result, 4) = "MFIP" THEN
			mfip_housing_start_date = #07/01/2015#
			curr_cash_bene_mo = left(right(cash_approval_result, 4), 2)
			curr_cash_bene_yr = right(cash_approval_result, 2)
			current_cash_month_for_case_noting_purposes = curr_cash_bene_mo & "/01/" & curr_cash_bene_yr
			'Determining whether the script needs to be concerned about the MFIP housing grant...
			IF DateDiff("D", mfip_housing_start_date, current_cash_month_for_case_noting_purposes) >= 0 THEN
				mfip_cash_amt = right(left(cash_approval_result, 12), 8)
				mfip_food_amt = right(left(cash_approval_result, 20), 8)
				mfip_housing_amt = left(right(cash_approval_result, 12), 8)
				call write_bullet_and_variable_in_CASE_NOTE(("MFIP Cash portion for " & curr_cash_bene_mo & "/" & curr_cash_bene_yr), FormatCurrency(mfip_cash_amt))
				call write_bullet_and_variable_in_CASE_NOTE(("MFIP Food portion for " & curr_cash_bene_mo & "/" & curr_cash_bene_yr), FormatCurrency(mfip_food_amt))
				call write_bullet_and_variable_in_CASE_NOTE(("MFIP Housing grant Amount for " & curr_cash_bene_mo & "/" & curr_cash_bene_yr), FormatCurrency(mfip_housing_amt))
			ELSEIF DateDiff("D", mfip_housing_start_date, current_cash_month_for_case_noting_purposes) < 0 THEN 
				mfip_cash_amt = right(left(cash_approval_result, 12), 8)
				mfip_food_amt = right(left(cash_approval_result, 20), 8)
				call write_bullet_and_variable_in_CASE_NOTE(("MFIP Cash portion for " & curr_cash_bene_mo & "/" & curr_cash_bene_yr), FormatCurrency(mfip_cash_amt))
				call write_bullet_and_variable_in_CASE_NOTE(("MFIP Food portion for " & curr_cash_bene_mo & "/" & curr_cash_bene_yr), FormatCurrency(mfip_food_amt))
			END IF
		


