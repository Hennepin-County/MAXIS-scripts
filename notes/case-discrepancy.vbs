'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CASE DISCREPANCY.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 72                      'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
'END OF stats block=========================================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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
		FuncLib_URL = "C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs"
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
call changelog_update("03/01/2020", "Updated TIKL functionality and TIKL text in the case note.", "Ilse Ferris")
call changelog_update("07/10/2018", "The ACTIONS TAKEN field is no longer a mandatory field.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'The script----------------------------------------------------------------------------------------------------
'Connecing to MAXIS, establishing the county code, and grabbing the case number
EMConnect ""
call MAXIS_case_number_finder(MAXIS_case_number)
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 336, 245, "Case discrepancy"
  EditBox 90, 10, 70, 15, MAXIS_case_number
  DropListBox 90, 35, 70, 15, "Select one..."+chr(9)+"found/pending"+chr(9)+"resolved ", discrepancy_status
  EditBox 90, 55, 70, 15, MNsure_case_number
  CheckBox 175, 20, 25, 10, "MA", MA_checkbox
  CheckBox 210, 20, 30, 10, "MSP", MSP_checkbox
  CheckBox 250, 20, 35, 10, "MNsure", MNsure_checkbox
  CheckBox 290, 20, 30, 10, "SNAP", SNAP_checkbox
  CheckBox 175, 35, 30, 10, "DWP", DWP_checkbox
  CheckBox 210, 35, 30, 10, "MFIP", MFIP_checkbox
  CheckBox 250, 35, 30, 10, "MSA", MSA_checkbox
  CheckBox 290, 35, 25, 10, "GA", GA_checkbox
  CheckBox 175, 50, 30, 10, "GRH", GRH_checkbox
  CheckBox 210, 50, 30, 10, "RCA", RCA_checkbox
  CheckBox 250, 50, 35, 10, "EMER", EMER_checkbox
  EditBox 90, 75, 240, 15, MEMB_PMI
  EditBox 90, 100, 240, 15, describe_discrepancy
  EditBox 90, 125, 240, 15, verifs_needed
  EditBox 90, 150, 240, 15, other_notes
  EditBox 90, 175, 240, 15, actions_taken
  CheckBox 25, 200, 60, 10, "MAXIS updated", MAXIS_updated
  CheckBox 95, 200, 60, 10, "MMIS updated", MMIS_updated
  CheckBox 160, 200, 170, 10, "Set TIKL for 10 day return of verifcations needed", TIKL_checkbox
  EditBox 90, 220, 130, 15, worker_signature
  CheckBox 350, 35, 25, 10, "GA", GA_checkbox
  CheckBox 235, 50, 30, 10, "GRH", GRH_checkbox
  CheckBox 270, 50, 30, 10, "RCA", RCA_checkbox
  CheckBox 310, 50, 35, 10, "EMER", EMER_checkbox
  EditBox 90, 85, 240, 15, describe_discrepancy
  EditBox 90, 110, 240, 15, MEMB_PMI
  EditBox 90, 135, 240, 15, verifs_needed
  EditBox 90, 160, 240, 15, other_notes
  EditBox 90, 185, 240, 15, actions_taken
  CheckBox 20, 215, 60, 10, "MAXIS updated", MAXIS_updated
  CheckBox 90, 215, 60, 10, "MMIS updated", MMIS_updated
  CheckBox 155, 215, 170, 10, "Set TIKL for 10 day return of verifcations needed", TIKL_checkbox
  EditBox 65, 235, 130, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 285, 235, 50, 15
    CancelButton 340, 235, 50, 15
  Text 40, 10, 45, 10, "Case number:"
  Text 35, 30, 55, 10, "MNsure case #:"
  Text 30, 50, 60, 10, "Discrepancy type:"
  Text 25, 70, 65, 10, "Discrepancy status:"
  GroupBox 225, 5, 160, 65, "Programs effected by the discrepancy:"
  Text 5, 90, 85, 10, "Describe the discrepancy:"
  Text 25, 115, 65, 10, "HH memb/PMI #(s):"
  Text 20, 140, 70, 10, "Verifications needed: "
  Text 50, 165, 40, 10, "Other notes:"
  Text 10, 190, 80, 10, "Resolution/Action taken:"
  Text 5, 240, 60, 10, "Worker signature:"
EndDialog

DO
	DO
	  err_msg = ""								'establishing value of varaible, this is necessary for the Do...LOOP
		dialog Dialog1				'initial dialog
		cancel_confirmation		'script ends if cancel is selected
    Call validate_MAXIS_case_number(err_msg, "*")
    If MNsure_case_number = "" and MNsure_checkbox = 1 then err_msg = err_msg & vbnewline & "* Enter the MNsure case number."
		If discrepancy_type = "Select or Type..." then err_msg = err_msg & vbnewline & "* Select a discrepancy type or enter your own type."
    If discrepancy_status = "Select one..." then err_msg = err_msg & vbnewline & "* Select a discrepancy status."
		If (MNsure_checkbox <> 1 and DWP_checkbox <> 1 and EMER_checkbox <> 1 and GA_checkbox <> 1 and GRH_checkbox <> 1 and MA_checkbox <> 1 and MFIP_checkbox <> 1 and MSA_checkbox <> 1 and MSP_checkbox <> 1 and RCA_checkbox <> 1 and SNAP_checkbox <> 1) then err_msg = err_msg & vbnewline & "* You must enter at least one program."
		If describe_discrepancy = "" then err_msg = err_msg & vbnewline & "* Describe the discrepancy."
    If MEMB_PMI = "" then err_msg = err_msg & vbnewline & "* Enter the HH member and/or PMI #'s the discrepancy effects."
		If worker_signature = "" then err_msg = err_msg & vbnewline & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


' Check if case is privileged and end script if it is privileged
Call navigate_to_MAXIS_screen_review_PRIV("CASE", "NOTE", is_this_priv)
If is_this_priv = True then script_end_procedure("This case is privileged and cannot be accessed. The script will now stop.")
' Confirm that the case is in county and end script if the case is out of county
EMReadScreen county_code, 4, 21, 14
If county_code <> worker_county_code then script_end_procedure("This case is out-of-county, and cannot access CASE:NOTE. The script will now stop.")


'Creating an incremantal variable based on the programs selected
If MA_checkbox = 1 then progs_effect = progs_effect & " MA,"
If MSP_checkbox = 1 then progs_effect = progs_effect & " Medicare savings program (MSP),"
If MNsure_checkbox = 1 then progs_effect = progs_effect & " MNsure,"
If SNAP_checkbox = 1 then progs_effect = progs_effect & " SNAP,"
If DWP_checkbox = 1 then progs_effect = progs_effect & " DWP,"
If MFIP_checkbox = 1 then progs_effect = progs_effect & " MFIP,"
IF MSA_checkbox = 1 then progs_effect = progs_effect & " MSA,"
If GA_checkbox = 1 then progs_effect = progs_effect & " GA,"
IF GRH_checkbox = 1 then progs_effect = progs_effect & " GRH,"
If RCA_checkbox = 1 then progs_effect = progs_effect & " RCA,"
If EMER_checkbox = 1 then progs_effect = progs_effect & " Emergency,"

'trims excess spaces of progs_effect
progs_effect = trim(progs_effect)
'takes the last comma off of progs_effect variable
If right(progs_effect, 1) = "," THEN progs_effect = left(progs_effect, len(progs_effect) - 1)

'TIKL coding
'Call create_TIKL(TIKL_text, num_of_days, date_to_start, ten_day_adjust, TIKL_note_text)
If TIKL_checkbox = 1 then Call create_TIKL("The following verifications were requested 10 days ago for a case discrepancy: " & verifs_needed, 10, date, True, TIKL_note_text)

'The case notes----------------------------------------------------------------------------------------------------
start_a_blank_case_note
Call write_variable_in_CASE_NOTE("~Case discrepancy " & discrepancy_status & "-" & discrepancy_type & "~")
Call write_bullet_and_variable_in_CASE_NOTE("Program(s) effected by discrepancy", progs_effect)
Call write_bullet_and_variable_in_CASE_NOTE("Description of the discrepancy", describe_discrepancy)
Call write_bullet_and_variable_in_CASE_NOTE("MNsure case #", MNsure_case_number)
Call write_bullet_and_variable_in_CASE_NOTE("HH member(s)/PMI#(s)", MEMB_PMI)
Call write_bullet_and_variable_in_CASE_NOTE("Verifications needed", verifs_needed)
If TIKL_checkbox = 1 then Call write_variable_in_CASE_NOTE("* TIKL'd out for 10 day return of requested verifications.")
Call write_variable_in_CASE_NOTE(TIKL_note_text)
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_bullet_and_variable_in_CASE_NOTE("Resolution/actions taken", actions_taken)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
