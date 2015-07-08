'Created by Robert Kalb and Charles Potter from Anoka County and and Ilse Ferris from Hennepin County.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTES - APPROVED PROGRAMS.vbs"
start_time = timer

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
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog benefits_approved, 0, 0, 271, 260, "Benefits Approved"
  CheckBox 80, 5, 30, 10, "SNAP", snap_approved_check
  CheckBox 115, 5, 30, 10, "Cash", cash_approved_check
  CheckBox 150, 5, 50, 10, "Health Care", hc_approved_check
  CheckBox 210, 5, 50, 10, "Emergency", emer_approved_check
  EditBox 55, 20, 60, 15, case_number
  ComboBox 180, 20, 80, 15, "Initial"+chr(9)+"Renewal"+chr(9)+"Recertification"+chr(9)+"Change"+chr(9)+"Reinstate", type_of_approval
  EditBox 115, 45, 150, 15, benefit_breakdown
  CheckBox 5, 65, 255, 10, "Check here to have the script autofill the SNAP approval.", autofill_snap_check
  EditBox 155, 80, 15, 15, snap_start_mo
  EditBox 170, 80, 15, 15, snap_start_yr
  EditBox 230, 80, 15, 15, snap_end_mo
  EditBox 245, 80, 15, 15, snap_end_yr
  CheckBox 5, 105, 255, 10, "Check here to have the script autofill the CASH approval.", autofill_cash_check
  EditBox 155, 120, 15, 15, cash_start_mo
  EditBox 170, 120, 15, 15, cash_start_yr
  EditBox 230, 120, 15, 15, cash_end_mo
  EditBox 245, 120, 15, 15, cash_end_yr
  EditBox 55, 145, 210, 15, other_notes
  EditBox 75, 165, 190, 15, programs_pending
  EditBox 55, 185, 210, 15, docs_needed
  'CheckBox 10, 205, 235, 10, "Check here if child support disregard was applied to MFIP/DWP case", CASH_WCOM_checkbox
  CheckBox 10, 220, 125, 10, "Check here if the case was FIATed", FIAT_checkbox
  EditBox 75, 235, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 160, 235, 50, 15
    CancelButton 215, 235, 50, 15
  Text 5, 25, 50, 10, "Case Number:"
  Text 5, 40, 110, 20, "Benefit Breakdown (Issuance/Spenddown/Premium):"
  Text 10, 85, 130, 10, "Select SNAP approval range (MM YY)..."
  Text 195, 85, 25, 10, "through"
  Text 10, 125, 130, 10, "Select CASH approval range (MM YY)..."
  Text 195, 125, 25, 10, "through"
  Text 5, 150, 45, 10, "Other Notes:"
  Text 5, 170, 70, 10, "Pending Program(s):"
  Text 5, 190, 50, 10, "Verifs Needed:"
  Text 15, 240, 60, 10, "Worker Signature: "
  Text 120, 25, 60, 10, "Type of Approval:"
  Text 5, 5, 70, 10, "Approved Programs:"
EndDialog


'THE SCRIPT----------------------------------------------------------------------------------------------------

EMConnect ""

maxis_check_function

'Finds the case number
call find_variable("Case Nbr: ", case_number, 8)
case_number = trim(case_number)
case_number = replace(case_number, "_", "")
If IsNumeric(case_number) = False then case_number = ""

'Finds the benefit month
EMReadScreen on_SELF, 4, 2, 50
IF on_SELF = "SELF" THEN
	CALL find_variable("Benefit Period (MM YY): ", bene_month, 2)
	IF bene_month <> "" THEN CALL find_variable("Benefit Period (MM YY): " & bene_month & " ", bene_year, 2)
ELSE
	CALL find_variable("Month: ", bene_month, 2)
	IF bene_month <> "" THEN CALL find_variable("Month: " & bene_month & " ", bene_year, 2)
END IF

'Converts the variables in the dialog into the variables "bene_month" and "bene_year" to autofill the edit boxes.
snap_start_mo = bene_month
snap_start_yr = bene_year
snap_end_mo = bene_month
snap_end_yr = bene_year

cash_start_mo = bene_month
cash_start_yr = bene_year
cash_end_mo = bene_month
cash_end_yr = bene_year

'Displays the dialog and navigates to case note
Do
  Do
    Do
      Dialog benefits_approved
      If buttonpressed = cancel then stopscript
	IF snap_approved_check = 0 AND autofill_snap_check = checked THEN MsgBox "You checked to have the SNAP results autofilled but did not select that SNAP was approved. Please reconsider your selections and try again."
	IF cash_approved_check = 0 AND autofill_cash_check = checked THEN MsgBox "You checked to have the CASH results autofilled but did not select that CASH was approved. Please reconsider your selections and try again."
      If case_number = "" then MsgBox "You must have a case number to continue!"
	If worker_signature = "" then Msgbox "Please sign your case note"

	IF autofill_cash_check = checked AND cash_approved_check = checked THEN
		'Calculates the number of benefit months the worker is trying to case note.
		cash_start = cdate(cash_start_mo & "/01/" & cash_start_yr)
		cash_end = cdate(cash_end_mo & "/01/" & cash_end_yr)
		IF datediff("M", date, cash_start) > 1 THEN MsgBox "Your CASH start month is invalid. You cannot case note eligibility results from more than 1 month into the future. Please change your months."
		IF datediff("M", date, cash_end) > 1 THEN MsgBox "Your CASH end month is invalid. You cannot case note eligibility results from more than 1 month into the future. Please change your months."
		IF datediff("M", cash_start, cash_end) < 0 THEN MsgBox "Please double check your CASH date range. Your start month cannot be later than your end month."
	END IF

	IF autofill_snap_check = checked AND snap_approved_check = checked THEN 
		'Calculates the number of benefit months the worker is trying to case note.
		snap_start = cdate(snap_start_mo & "/01/" & snap_start_yr)
		snap_end = cdate(snap_end_mo & "/01/" & snap_end_yr)
		IF datediff("M", date, snap_start) > 1 THEN MsgBox "Your SNAP start month is invalid. You cannot case note eligibility results from more than 1 month into the future. Please change your months."
		IF datediff("M", date, snap_end) > 1 THEN MsgBox "Your SNAP end month is invalid. You cannot case note eligibility results from more than 1 month into the future. Please change your months."
		IF datediff("M", snap_start, snap_end) < 0 THEN MsgBox "Please double check your SNAP date range. Your start month cannot be later than your end month."
	END IF

    Loop until case_number <> "" AND _
	worker_signature <> "" AND _
	((snap_approved_check = checked AND autofill_snap_check = checked AND (datediff("M", snap_start, snap_end) >= 0) AND (datediff("M", date, snap_start) < 2) AND (datediff("M", date, snap_end) < 2)) OR (autofill_snap_check = 0)) AND _
	((cash_approved_check = checked AND autofill_cash_check = checked AND (datediff("M", cash_start, cash_end) >= 0) AND (datediff("M", date, cash_start) < 2) AND (datediff("M", date, cash_end) < 2)) OR (autofill_cash_check = 0))

    transmit
    EMReadScreen MAXIS_check, 5, 1, 39
    If MAXIS_check <> "MAXIS" and MAXIS_check <> "AXIS " then MsgBox "You appear to be locked out of MAXIS. Are you passworded out? Did you navigate away from MAXIS?"
  Loop until MAXIS_check = "MAXIS" or MAXIS_check = "AXIS "
  call navigate_to_screen("case", "note")
  PF9
  EMReadScreen mode_check, 7, 20, 3
  If mode_check <> "Mode: A" and mode_check <> "Mode: E" then MsgBox "For some reason, the script can't get to a case note. Did you start the script in inquiry by mistake? Navigate to MAXIS production, or shut down the script and try again."
Loop until mode_check = "Mode: A" or mode_check = "Mode: E"

total_snap_months = (datediff("m", snap_start, snap_end)) + 1
total_cash_months = (datediff("m", cash_start, cash_end)) + 1

'Navigates to the ELIG results for SNAP, if the worker desires to have the script autofill the case note with SNAP approval information.
IF autofill_snap_check = checked THEN
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

		IF dwp_elig_summ <> " " THEN 
			EMReadScreen date_of_last_DWP_version, 8, 7, 48
'			IF cdate(date_of_last_DWP_version) = date THEN 
			prog_to_check_array = prog_to_check_array & "DW" & cash_month & cash_year & "/"
		END IF
		IF mfip_elig_summ <> " " THEN
			EMReadScreen date_of_last_MFIP_version, 8, 8, 48
'			IF date_of_last_MFIP_version = "11/06/14" THEN
			prog_to_check_array = prog_to_check_array & "MF" & cash_month & cash_year & "/"
		END IF
		IF msa_elig_summ <> " " THEN
			EMReadScreen date_of_last_MSA_version, 8, 11, 48
'			IF cdate(date_of_last_MSA_version) = date THEN
			prog_to_check_array = prog_to_check_array & "MS" & cash_month & cash_year & "/"
		END IF
		IF ga_elig_summ <> " " THEN 
			EMReadScreen date_of_last_GA_version, 8, 12, 48
'			IF cdate(date_of_last_GA_version) = date THEN
			prog_to_check_array = prog_to_check_array & "GA" & cash_month & cash_year & "/"
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


'updates WCOM with notice requirements if MFIP or DWP child support income disregarded in the budget
read_row = 7

If CASH_WCOM_checkbox = checked THEN 
	Call navigate_to_MAXIS_screen ("SPEC", "WCOM")
	EMReadscreen CASH_check, 2, read_row, 26  'checking to make sure that notice is for MFIP or DWP
	EMReadScreen Print_status_check, 7, read_row, 71 'checking to see if notice is in 'waiting status'
	'checking program type and if it's a notice that is in waiting status (waiting status will make it editable)
	If(CASH_check = "MF" AND Print_status_check = "Waiting") OR (CASH_check = "DW" AND Print_status_check = "Waiting") THEN 
		EMSetcursor read_row, 13
		EMSendKey "x"
		Transmit
		PF9
		EMSetCursor 03, 15
		'WCOM required by workers upon approval of MFIP and DWP cases with child support FIAT'd out of the budget
		Call write_variable_in_SPEC_MEMO("************************************************************")
		Call write_variable_in_SPEC_MEMO("")
		Call write_variable_in_SPEC_MEMO("Starting July 1, 2015 a new law begins that allows us to not count some of the child support you get when determining your monthly MFIP/DWP benefit amount:")
		Call write_variable_in_SPEC_MEMO("")
		Call write_variable_in_SPEC_MEMO("* $100 for an assistance unit with one child")
		Call write_variable_in_SPEC_MEMO("* $200 for an assistance unit with two or more children")
		Call write_variable_in_SPEC_MEMO("")
		Call write_variable_in_SPEC_MEMO("Because of this change, you may see an increase in your benefit amount.")
		Call write_variable_in_SPEC_MEMO("************************************************************")
		PF4
		PF3
	ELSE 
		Msgbox "There is not a pending notice for this cash case. The script was unable to update your notice."
	END if
END If

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
IF autofill_cash_check = checked THEN
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
		ELSEIF left(cash_approval_result, 4) = "DWP_" THEN
			dwp_shel_amt = right(left(cash_approval_result, 12), 8)
			dwp_pers_amt = left(right(cash_approval_result, 12), 8)
			curr_cash_bene_mo = left(right(cash_approval_result, 4), 2)
			curr_cash_bene_yr = right(cash_approval_result, 2)
			call write_bullet_and_variable_in_CASE_NOTE(("DWP Shelter Benefit Amount for " & curr_cash_bene_mo & "/" & curr_cash_bene_yr), FormatCurrency(dwp_shel_amt))
			call write_bullet_and_variable_in_CASE_NOTE(("DWP Personal Needs Amount for " & curr_cash_bene_mo & "/" & curr_cash_bene_yr), FormatCurrency(dwp_pers_amt))
		ELSE
			cash_program = left(cash_approval_result, 4)
			cash_program = replace(cash_program, "_", "")
			cash_bene_amt = right(left(cash_approval_result, 12), 8)
			curr_cash_bene_mo = left(right(cash_approval_result, 4), 2)
			curr_cash_bene_yr = right(cash_approval_result, 2)
			cash_header = (cash_program & " Amount for " & curr_cash_bene_mo & "/" & curr_cash_bene_yr)
			call write_bullet_and_variable_in_CASE_NOTE(cash_header, FormatCurrency(cash_bene_amt))
		END IF
	NEXT
END IF
IF FIAT_checkbox = 1 THEN Call write_variable_in_CASE_NOTE ("* This case has been FIATed.")
If CASH_WCOM_checkbox = 1 THEN Call write_variable_in_CASE_NOTE ("* The child support disregard was applied to this case.")
IF other_notes <> "" THEN call write_bullet_and_variable_in_CASE_NOTE("Approval Notes", other_notes)
IF programs_pending <> "" THEN call write_bullet_and_variable_in_CASE_NOTE("Programs Pending", programs_pending)
If docs_needed <> "" then call write_bullet_and_variable_in_CASE_NOTE("Verifs needed", docs_needed) 
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)



'Runs denied progs if selected
If closed_progs_check = checked then run_from_github(script_repository & "NOTES/NOTES - CLOSED PROGRAMS.vbs")

'Runs denied progs if selected
If denied_progs_check = checked then run_script(script_repository & "NOTES/NOTES - DENIED PROGRAMS.vbs")

script_end_procedure("Success! Please remember to check the generated notice to make sure it reads correctly. If not please add WCOMs to make notice read correctly.")
