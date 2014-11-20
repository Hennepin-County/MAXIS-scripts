'Created by Robert Kalb and Charles Potter from Anoka County.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - approved programs"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-Maxis-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog benefits_approved, 0, 0, 271, 285, "Benefits Approved"
  CheckBox 15, 25, 35, 10, "SNAP", snap_approved_check
  CheckBox 75, 25, 55, 10, "Health Care", hc_approved_check
  CheckBox 155, 25, 35, 10, "Cash", cash_approved_check
  CheckBox 210, 25, 55, 10, "Emergency", emer_approved_check
  ComboBox 70, 40, 85, 15, ""+chr(9)+"Initial"+chr(9)+"Renewal"+chr(9)+"Recertification"+chr(9)+"Change"+chr(9)+"Reinstate", type_of_approval
  EditBox 65, 60, 70, 15, case_number
  EditBox 120, 85, 145, 15, benefit_breakdown
  CheckBox 5, 105, 255, 10, "Check here to have the script autofill the SNAP approval.", autofill_snap_check
  EditBox 155, 120, 15, 15, snap_start_mo
  EditBox 170, 120, 15, 15, snap_start_yr
  EditBox 230, 120, 15, 15, snap_end_mo
  EditBox 245, 120, 15, 15, snap_end_yr
  CheckBox 5, 145, 255, 10, "Check here to have the script autofill the CASH approval.", autofill_cash_check
  EditBox 155, 160, 15, 15, cash_start_mo
  EditBox 170, 160, 15, 15, cash_start_yr
  EditBox 230, 160, 15, 15, cash_end_mo
  EditBox 245, 160, 15, 15, cash_end_yr
  EditBox 55, 185, 210, 15, other_notes
  EditBox 85, 205, 180, 15, programs_pending
  EditBox 65, 225, 200, 15, docs_needed
  EditBox 65, 245, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 155, 260, 50, 15
    CancelButton 210, 260, 50, 15
  Text 5, 5, 70, 10, "Approved Programs:"
  Text 5, 45, 65, 10, "Type of Approval:"
  Text 5, 65, 55, 10, "Case Number:"
  Text 5, 80, 110, 20, "Benefit Breakdown (Issuance/Spenddown/Premium):"
  Text 10, 125, 130, 10, "Select SNAP approval range (MM YY)..."
  Text 195, 125, 25, 10, "through"
  Text 10, 165, 130, 10, "Select CASH approval range (MM YY)..."
  Text 195, 165, 25, 10, "through"
  Text 5, 190, 45, 10, "Other Notes:"
  Text 5, 210, 75, 10, "Pending Program(s):"
  Text 5, 230, 55, 10, "Verifs Needed:"
  Text 5, 250, 60, 10, "Worker Signature: "
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
      If buttonpressed = 0 then stopscript
	IF snap_approved_check = 0 AND autofill_snap_check = checked THEN MsgBox "You checked to have the SNAP results autofilled but did not select that SNAP was approved. Please reconsider your selections and try again."
	IF cash_approved_check = 0 AND autofill_cash_check = checked THEN MsgBox "You checked to have the CASH results autofilled but did not select that CASH was approved. Please reconsider your selections and try again."
      If case_number = "" then MsgBox "You must have a case number to continue!"
	If worker_signature = "" then Msgbox "Please sign your case note"

	IF autofill_cash_check = checked AND cash_approved_check = 1 THEN
		'Calculates the number of benefit months the worker is trying to case note.
		cash_start = cdate(cash_start_mo & "/01/" & cash_start_yr)
		cash_end = cdate(cash_end_mo & "/01/" & cash_end_yr)
		IF datediff("M", date, cash_start) > 1 THEN MsgBox "Your CASH start month is invalid. You cannot case note eligibility results from more than 1 month into the future. Please change your months."
		IF datediff("M", date, cash_end) > 1 THEN MsgBox "Your CASH end month is invalid. You cannot case note eligibility results from more than 1 month into the future. Please change your months."
		IF datediff("M", cash_start, cash_end) < 0 THEN MsgBox "Please double check your CASH date range. Your start month cannot be later than your end month."
	END IF

	IF autofill_snap_check = checked AND snap_approved_check = 1 THEN 
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
	((cash_approved_check = 1 AND autofill_cash_check = checked AND (datediff("M", cash_start, cash_end) >= 0) AND (datediff("M", date, cash_start) < 2) AND (datediff("M", date, cash_end) < 2)) OR (autofill_cash_check = 0))

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
IF autofill_snap_check = 1 THEN
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
				EMReadScreen snap_bene_amt, 3, 13, 75
				EMReadScreen current_snap_bene_mo, 2, 19, 54
				EMReadScreen current_snap_bene_yr, 2, 19, 57
				snap_bene_amt = replace(snap_bene_amt, " ", "0")
				snap_approval_array = snap_approval_array & snap_bene_amt & current_snap_bene_mo & current_snap_bene_yr & " "
			ELSE
				script_end_procedure("Your most recent SNAP approval for the benefit month chosen is not from today. The script cannot autofill this result. Process manually.")
			END IF
		ELSE
			EMReadScreen approval_versions, 2, 2, 18
			IF approval_versions = "1" THEN script_end_procedure("You do not have an approved version of SNAP in the selected benefit month. Please approve before running the script.")
			approval_versions = int(approval_versions)
			approval_to_check = approval_versions - 1
			EMWriteScreen approval_to_check, 19, 78
			transmit
			EMReadScreen approval_date, 8, 3, 14
			approval_date = cdate(approval_date)
			IF approval_date = date THEN
				EMReadScreen snap_bene_amt, 3, 13, 75
				EMReadScreen current_snap_bene_mo, 2, 19, 54
				EMReadScreen current_snap_bene_yr, 2, 19, 57
				snap_bene_amt = replace(snap_bene_amt, " ", "0")
				snap_approval_array = snap_approval_array & snap_bene_amt & current_snap_bene_mo & current_snap_bene_yr & " "
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

'----------This version only autofills MFIP. Later versions can be updated with other cash programs.----------
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

		EMReadScreen mfip_elig_summ, 1, 8, 40
			IF mfip_elig_summ <> " " THEN prog_to_check = "MFIP"
		EMReadScreen dwp_elig_summ, 1, 7, 40
			IF dwp_elig_summ <> " " THEN prog_to_check = "DWP"
		EMReadScreen msa_elig_summ, 1, 11, 40
			IF msa_elig_summ <> " " THEN prog_to_check = "MSA"
		EMReadScreen ga_elig_summ, 1, 12, 40
			IF ga_elig_summ <> " " THEN prog_to_check = "GA"

		IF prog_to_check = "MFIP" THEN
			'MFIP portion
			IF len(cash_month) = 1 THEN cash_month = "0" & cash_month
			call navigate_to_screen("ELIG", "MFIP")
			EMWriteScreen cash_month, 20, 56
			EMWriteScreen cash_year, 20, 59
			EMWRiteScreen "MFSM", 20, 71
			transmit
			EMReadScreen cash_approved_version, 8, 3, 3
			IF cash_approved_version = "APPROVED" THEN
				EMReadScreen cash_approval_date, 8, 3, 14
				cash_approval_date = cdate(cash_approval_date)
				IF cash_approval_date = date THEN
					EMReadScreen mfip_bene_cash_amt, 4, 15, 74
					EMReadScreen mfip_bene_food_amt, 4, 16, 74
					EMReadScreen current_cash_bene_mo, 2, 20, 55
					EMReadScreen current_cash_bene_yr, 2, 20, 58
					mfip_bene_cash_amt = replace(mfip_bene_cash_amt, " ", "0")
					mfip_bene_food_amt = replace(mfip_bene_food_amt, " ", "0")
					cash_approval_array = cash_approval_array & "MFIP" & mfip_bene_cash_amt & mfip_bene_food_amt & current_cash_bene_mo & current_cash_bene_yr & " "
				ELSE
					MsgBox "Your most recent CASH approval for the benefit month chosen is not from today. The script cannot autofill this result."
					'script_end_procedure("Your most recent CASH approval for the benefit month chosen is not from today. The script cannot autofill this result. Process manually.")
				END IF
			ELSE
				EMReadScreen cash_approval_versions, 1, 2, 18
				IF cash_approval_versions = "1" THEN script_end_procedure("You do not have an approved version of CASH in the selected benefit month. Please approve before running the script.")
				cash_approval_versions = int(cash_approval_versions)
				cash_approval_to_check = cash_approval_versions - 1
				EMWriteScreen cash_approval_to_check, 20, 79
				transmit
				EMReadScreen cash_approval_date, 8, 3, 14
				cash_approval_date = cdate(cash_approval_date)
				IF cash_approval_date = date THEN
					EMReadScreen mfip_bene_cash_amt, 4, 15, 74
					EMReadScreen mfip_bene_food_amt, 4, 16, 74
					EMReadScreen current_cash_bene_mo, 2, 20, 55
					EMReadScreen current_cash_bene_yr, 2, 20, 58
					mfip_bene_cash_amt = replace(mfip_bene_cash_amt, " ", "0")
					mfip_bene_food_amt = replace(mfip_bene_food_amt, " ", "0")
					cash_approval_array = cash_approval_array & mfip_bene_cash_amt & mfip_bene_food_amt & current_cash_bene_mo & current_cash_bene_yr & " "
				ELSE
					MsgBox "Your most recent CASH approval for the benefit month chosen is not from today. The script cannot autofill this result."
				END IF
			END IF	

		ELSEIF prog_to_check = "GA" THEN
			'GA portion
			IF len(cash_month) = 1 THEN cash_month = "0" & cash_month
			call navigate_to_screen("ELIG", "GA")
			EMWriteScreen cash_month, 20, 54
			EMWriteScreen cash_year, 20, 57
			EMWRiteScreen "GASM", 20, 70
			transmit
			EMReadScreen cash_approved_version, 8, 3, 3
			IF cash_approved_version = "APPROVED" THEN
				EMReadScreen cash_approval_date, 8, 3, 15
				cash_approval_date = cdate(cash_approval_date)
				IF cash_approval_date = date THEN
					EMReadScreen GA_bene_cash_amt, 4, 14, 73
					EMReadScreen current_cash_bene_mo, 2, 20, 54
					EMReadScreen current_cash_bene_yr, 2, 20, 57
					GA_bene_cash_amt = replace(GA_bene_cash_amt, " ", "0")
					cash_approval_array = cash_approval_array & "GA__" & GA_bene_cash_amt & current_cash_bene_mo & current_cash_bene_yr & " "
				ELSE
					MsgBox "Your most recent CASH approval for the benefit month chosen is not from today. The script cannot autofill this result."
				END IF
			ELSE
				EMReadScreen cash_approval_versions, 1, 2, 18
				IF cash_approval_versions = "1" THEN script_end_procedure("You do not have an approved version of CASH in the selected benefit month. Please approve before running the script.")
				cash_approval_versions = int(cash_approval_versions)
				cash_approval_to_check = cash_approval_versions - 1
				EMWriteScreen cash_approval_to_check, 20, 79
				transmit
				EMReadScreen cash_approval_date, 8, 3, 15
				cash_approval_date = cdate(cash_approval_date)
				IF cash_approval_date = date THEN
					EMReadScreen GA_bene_cash_amt, 4, 14, 73
					EMReadScreen current_cash_bene_mo, 2, 20, 54
					EMReadScreen current_cash_bene_yr, 2, 20, 57
					GA_bene_cash_amt = replace(GA_bene_cash_amt, " ", "0")
					cash_approval_array = cash_approval_array & "GA__" & GA_bene_cash_amt & current_cash_bene_mo & current_cash_bene_yr & " "
				ELSE
					MsgBox "Your most recent CASH approval for the benefit month chosen is not from today. The script cannot autofill this result."
				END IF
			END IF
	
		ELSEIF prog_to_check = "MSA" THEN
			'MSA portion
			IF len(cash_month) = 1 THEN cash_month = "0" & cash_month
			call navigate_to_screen("ELIG", "MSA")
			EMWriteScreen cash_month, 20, 56
			EMWriteScreen cash_year, 20, 59
			EMWRiteScreen "MSSM", 20, 71
			transmit
			EMReadScreen cash_approved_version, 8, 3, 3
			IF cash_approved_version = "APPROVED" THEN
				EMReadScreen cash_approval_date, 8, 3, 14
				cash_approval_date = cdate(cash_approval_date)
				IF cash_approval_date = date THEN
					EMReadScreen MSA_bene_cash_amt, 4, 17, 74
					EMReadScreen current_cash_bene_mo, 2, 20, 54
					EMReadScreen current_cash_bene_yr, 2, 20, 57
					MSA_bene_cash_amt = replace(MSA_bene_cash_amt, " ", "0")
					cash_approval_array = cash_approval_array & "MSA_" & MSA_bene_cash_amt & current_cash_bene_mo & current_cash_bene_yr & " "
				ELSE
					MsgBox "Your most recent CASH approval for the benefit month chosen is not from today. The script cannot autofill this result."
				END IF
			ELSE
				EMReadScreen cash_approval_versions, 1, 2, 18
				IF cash_approval_versions = "1" THEN script_end_procedure("You do not have an approved version of CASH in the selected benefit month. Please approve before running the script.")
				cash_approval_versions = int(cash_approval_versions)
				cash_approval_to_check = cash_approval_versions - 1
				EMWriteScreen cash_approval_to_check, 20, 79
				transmit
				EMReadScreen cash_approval_date, 8, 3, 14
				cash_approval_date = cdate(cash_approval_date)
				IF cash_approval_date = date THEN
					EMReadScreen MSA_bene_cash_amt, 4, 17, 74
					EMReadScreen current_cash_bene_mo, 2, 20, 54
					EMReadScreen current_cash_bene_yr, 2, 20, 57
					MSA_bene_cash_amt = replace(MSA_bene_cash_amt, " ", "0")
					cash_approval_array = cash_approval_array & "MSA_" & MSA_bene_cash_amt & current_cash_bene_mo & current_cash_bene_yr & " "
				ELSE
					MsgBox "Your most recent CASH approval for the benefit month chosen is not from today. The script cannot autofill this result."
				END IF
			END IF

		ELSEIF prog_to_check = "DWP" THEN
			'DWP portion
			IF len(cash_month) = 1 THEN cash_month = "0" & cash_month
			call navigate_to_screen("ELIG", "DWP")
			EMWriteScreen cash_month, 20, 56
			EMWriteScreen cash_year, 20, 59
			EMWRiteScreen "DWSM", 20, 71
			transmit
			EMReadScreen cash_approved_version, 8, 3, 3
			IF cash_approved_version = "APPROVED" THEN
				EMReadScreen cash_approval_date, 8, 3, 14
				cash_approval_date = cdate(cash_approval_date)
				IF cash_approval_date = date THEN
					EMReadScreen DWP_bene_shel_amt, 4, 13, 74
					EMReadScreen DWP_bene_pers_amt, 4, 14, 74
					EMReadScreen current_cash_bene_mo, 2, 20, 54
					EMReadScreen current_cash_bene_yr, 2, 20, 57
					DWP_bene_shel_amt = replace(DWP_bene_shel_amt, " ", "0")
					DWP_bene_pers_amt = replace(DWP_bene_pers_amt, " ", "0")
					cash_approval_array = cash_approval_array & "DWP_" & DWP_bene_shel_amt & DWP_bene_pers_amt & current_cash_bene_mo & current_cash_bene_yr & " "
				ELSE
					MsgBox "Your most recent CASH approval for the benefit month chosen is not from today. The script cannot autofill this result."
				END IF
			ELSE
				EMReadScreen cash_approval_versions, 1, 2, 18
				IF cash_approval_versions = "1" THEN script_end_procedure("You do not have an approved version of CASH in the selected benefit month. Please approve before running the script.")
				cash_approval_versions = int(cash_approval_versions)
				cash_approval_to_check = cash_approval_versions - 1
				EMWriteScreen cash_approval_to_check, 20, 79
				transmit
				EMReadScreen cash_approval_date, 8, 3, 14
				cash_approval_date = cdate(cash_approval_date)
				IF cash_approval_date = date THEN
					EMReadScreen DWP_bene_shel_amt, 4, 13, 74
					EMReadScreen DWP_bene_pers_amt, 4, 14, 74
					EMReadScreen current_cash_bene_mo, 2, 20, 54
					EMReadScreen current_cash_bene_yr, 2, 20, 57
					DWP_bene_shel_amt = replace(DWP_bene_shel_amt, " ", "0")
					DWP_bene_pers_amt = replace(DWP_bene_pers_amt, " ", "0")
					cash_approval_array = cash_approval_array & "DWP_" & DWP_bene_shel_amt & DWP_bene_pers_amt & current_cash_bene_mo & current_cash_bene_yr & " "
				ELSE
					MsgBox "Your most recent CASH approval for the benefit month chosen is not from today. The script cannot autofill this result."
				END IF
			END IF

		ELSE
			MsgBox("There are no CASH results in " & cash_month & "/" & cash_year & ".")

		END IF

		cash_month = cash_month + 1
		IF cash_month = 13 THEN
			cash_month = 1
			cash_year = cash_year + 1
		END IF
		cash_count = cash_count + 1
	LOOP UNTIL cash_count = total_cash_months
END IF

cash_approval_array = trim(cash_approval_array)
cash_approval_array = split(cash_approval_array)

'Case notes
call navigate_to_screen("CASE", "NOTE")
PF9
IF snap_approved_check = 1 THEN approved_programs = approved_programs & "SNAP/"
IF hc_approved_check = 1 THEN approved_programs = approved_programs & "HC/"
IF cash_approved_check = 1 THEN approved_programs = approved_programs & "CASH/"
IF emer_approved_check = 1 THEN approved_programs = approved_programs & "EMER/"
EMSendKey "---Approved " & approved_programs & "<backspace>" & " " & type_of_approval & "---" & "<newline>"
IF benefit_breakdown <> "" THEN call write_editbox_in_case_note("Benefit Breakdown", benefit_breakdown, 6)
IF autofill_snap_check = checked THEN
	FOR EACH snap_approval_result in snap_approval_array
		bene_amount = left(snap_approval_result, 3)
		benefit_month = left(right(snap_approval_result, 4), 2)
		benefit_year = right(snap_approval_result, 2)
		snap_header = ("SNAP for " & benefit_month & "/" & benefit_year)
		call write_editbox_in_case_note(snap_header, FormatCurrency(bene_amount), 6)
	NEXT
END IF
IF autofill_cash_check = checked THEN
	FOR EACH cash_approval_result IN cash_approval_array

		IF left(cash_approval_result, 4) = "MFIP" THEN
			mfip_cash_amt = right(left(cash_approval_result, 8), 4)
			mfip_food_amt = left(right(cash_approval_result, 8), 4)
			curr_cash_bene_mo = left(right(cash_approval_result, 4), 2)
			curr_cash_bene_yr = right(cash_approval_result, 2)
			call write_editbox_in_case_note(("MFIP Cash Amount for " & curr_cash_bene_mo & "/" & curr_cash_bene_yr), FormatCurrency(mfip_cash_amt), 6)
			call write_editbox_in_case_note(("MFIP Food Amount for " & curr_cash_bene_mo & "/" & curr_cash_bene_yr), FormatCurrency(mfip_food_amt), 6)
		ELSEIF left(cash_approval_result, 4) = "DWP_" THEN
			dwp_shel_amt = right(left(cash_approval_result, 8), 4)
			dwp_pers_amt = left(right(cash_approval_result, 8), 4)
			curr_cash_bene_mo = left(right(cash_approval_result, 4), 2)
			curr_cash_bene_yr = right(cash_approval_result, 2)
			call write_editbox_in_case_note(("DWP Shelter Benefit Amount for " & curr_cash_bene_mo & "/" & curr_cash_bene_yr), FormatCurrency(dwp_shel_amt), 6)
			call write_editbox_in_case_note(("DWP Personal Needs Amount for " & curr_cash_bene_mo & "/" & curr_cash_bene_yr), FormatCurrency(dwp_pers_amt), 6)
		ELSE
			cash_program = left(cash_approval_result, 4)
			cash_program = replace(cash_program, "_", "")
			cash_bene_amt = right(left(cash_approval_result, 8), 4)
			curr_cash_bene_mo = left(right(cash_approval_result, 4), 2)
			curr_cash_bene_yr = right(cash_approval_result, 2)
			cash_header = (cash_program & " Amount for " & curr_cash_bene_mo & "/" & curr_cash_bene_yr)
			call write_editbox_in_case_note(cash_header, FormatCurrency(cash_bene_amt), 6)
		END IF
	NEXT
END IF
IF other_notes <> "" THEN call write_editbox_in_case_note("Approval Notes", other_notes, 6)
IF programs_pending <> "" THEN call write_editbox_in_case_note("Programs Pending", programs_pending, 6)
If docs_needed <> "" then call write_editbox_in_case_note("Verifs needed", docs_needed, 6) 
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

'Runs denied progs if selected
If closed_progs_check = 1 then run_another_script("C:\DHS-Maxis-Scripts\Script Files\NOTE - closed progs.vbs")

'Runs denied progs if selected
If denied_progs_check = 1 then run_another_script("C:\DHS-Maxis-Scripts\Script Files\NOTE - denied progs.vbs")

script_end_procedure("")

