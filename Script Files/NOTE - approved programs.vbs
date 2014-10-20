'Created by Robert Kalb and Charles Potter from Anoka County.

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTE - approved programs"
start_time = timer

'LOADING ROUTINE FUNCTIONS----------------------------------------------------------------------------------------------------
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'DIALOGS----------------------------------------------------------------------------------------------------

BeginDialog benefits_approved, 0, 0, 271, 240, "Benefits Approved"
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
  EditBox 55, 140, 210, 15, other_notes
  EditBox 85, 160, 180, 15, programs_pending
  EditBox 65, 180, 200, 15, docs_needed
  EditBox 65, 200, 80, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 155, 220, 50, 15
    CancelButton 210, 220, 50, 15
  Text 10, 125, 130, 10, "Select SNAP approval range (MM YY)..."
  Text 5, 45, 65, 10, "Type of Approval:"
  Text 5, 185, 55, 10, "Verifs Needed:"
  Text 5, 205, 60, 10, "Worker Signature: "
  Text 5, 65, 55, 10, "Case Number:"
  Text 195, 125, 25, 10, "through"
  Text 5, 5, 70, 10, "Approved Programs:"
  Text 5, 80, 110, 20, "Benefit Breakdown (Issuance/Spenddown/Premium):"
  Text 5, 145, 45, 10, "Other Notes:"
  Text 5, 165, 75, 10, "Pending Program(s):"
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

'Displays the dialog and navigates to case note
Do
  Do
    Do
      Dialog benefits_approved
      If buttonpressed = 0 then stopscript
	IF snap_approved_check = 0 AND autofill_snap_check = 1 THEN MsgBox "You checked to have the SNAP results autofilled but did not select that SNAP was approved. Please reconsider your selections and try again."
	IF cash_approved_check = 0 AND autofill_cash_check = 1 THEN MsgBox "You checked to have the CASH results autofilled but did not select that CASH was approved. Please reconsider your selections and try again."
      If case_number = "" then MsgBox "You must have a case number to continue!"
	If worker_signature = "" then Msgbox "Please sign your case note"

	IF autofill_snap_check = 1 AND snap_approved_check = 1 THEN 
		'Calculates the number of benefit months the worker is trying to case note.
		snap_start = cdate(snap_start_mo & "/01/" & snap_start_yr)
		snap_end = cdate(snap_end_mo & "/01/" & snap_end_yr)
		IF datediff("M", date, snap_start) > 1 THEN MsgBox "Your start month is invalid. You cannot case note eligibility results from more than 1 month into the future. Please change your months."
		IF datediff("M", date, snap_end) > 1 THEN MsgBox "Your end month is invalid. You cannot case note eligibility results from more than 1 month into the future. Please change your months."
		IF datediff("M", snap_start, snap_end) < 0 THEN MsgBox "Please double check your date range. Your start month cannot be later than your end month."
	END IF

    Loop until case_number <> "" and worker_signature <> "" AND ((snap_approved_check = 1 AND autofill_snap_check = 1 AND (datediff("M", snap_start, snap_end) >= 0) AND (datediff("M", date, snap_start) < 2) AND (datediff("M", date, snap_end) < 2)) OR (autofill_snap_check = 0))
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
			EMReadScreen approval_versions, 1, 2, 19
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

'Case notes
call navigate_to_screen("CASE", "NOTE")
PF9
IF snap_approved_check = 1 THEN approved_programs = approved_programs & "SNAP/"
IF hc_approved_check = 1 THEN approved_programs = approved_programs & "HC/"
IF cash_approved_check = 1 THEN approved_programs = approved_programs & "CASH/"
IF emer_approved_check = 1 THEN approved_programs = approved_programs & "EMER/"
EMSendKey "---Approved " & approved_programs & "<backspace>" & " " & type_of_approval & "---" & "<newline>"
IF benefit_breakdown <> "" THEN call write_editbox_in_case_note("Benefit Breakdown", benefit_breakdown, 6)
IF autofill_snap_check = 1 THEN
	FOR EACH snap_approval_result in snap_approval_array
		bene_amount = left(snap_approval_result, 3)
		benefit_month = left(right(snap_approval_result, 4), 2)
		benefit_year = right(snap_approval_result, 2)
		snap_header = ("SNAP for " & benefit_month & "/" & benefit_year)
		call write_editbox_in_case_note(snap_header, FormatCurrency(bene_amount), 6)
	NEXT
END IF
IF other_notes <> "" THEN call write_editbox_in_case_note("Approval Notes", other_notes, 6)
IF programs_pending <> "" THEN call write_editbox_in_case_note("Programs Pending", programs_pending, 6)
If docs_needed <> "" then call write_editbox_in_case_note("Verifs needed", docs_needed, 6) 
call write_new_line_in_case_note("---")
call write_new_line_in_case_note(worker_signature)

'Runs denied progs if selected
If closed_progs_check = 1 then run_another_script("C:\DHS-MAXIS-Scripts\Script Files\NOTE - closed progs.vbs")

'Runs denied progs if selected
If denied_progs_check = 1 then run_another_script("C:\DHS-MAXIS-Scripts\Script Files\NOTE - denied progs.vbs")


script_end_procedure("")