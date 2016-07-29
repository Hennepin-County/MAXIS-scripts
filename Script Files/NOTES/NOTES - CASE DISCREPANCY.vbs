'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - CASE DISCREPANCY.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 0                       'manual run time in seconds
STATS_denomination = "C"                   'C is for each CASE
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

'Dialogs----------------------------------------------------------------------------------------------------
BeginDialog case_discrepancy_dialog, 0, 0, 336, 225, "Case discrepancy"
  EditBox 70, 10, 65, 15, MAXIS_case_number
  DropListBox 70, 35, 65, 15, "Select one..."+chr(9)+"found/pending"+chr(9)+"resolved ", discrepancy_status
  CheckBox 150, 20, 30, 10, "DWP", DWP_checkbox
  CheckBox 185, 20, 35, 10, "EMER", EMER_checkbox
  CheckBox 225, 20, 25, 10, "GA", GA_checkbox
  CheckBox 260, 20, 30, 10, "GRH", GRH_checkbox
  CheckBox 295, 20, 25, 10, "MA", MA_checkbox
  CheckBox 150, 35, 30, 10, "MFIP", MFIP_checkbox
  CheckBox 185, 35, 30, 10, "MSP", MSP_checkbox
  CheckBox 225, 35, 30, 10, "MSA", MSA_checkbox
  CheckBox 260, 35, 30, 10, "RCA", RCA_checkbox
  CheckBox 295, 35, 30, 10, "SNAP", SNAP_checkbox
  EditBox 95, 60, 235, 15, MEMB_PMI
  EditBox 95, 85, 235, 15, describe_discrepancy
  EditBox 95, 110, 235, 15, verifs_needed
  EditBox 95, 135, 235, 15, other_notes
  EditBox 95, 160, 235, 15, actions_taken
  CheckBox 25, 185, 60, 10, "MAXIS updated", MAXIS_updated
  CheckBox 95, 185, 60, 10, "MMIS updated", MMIS_updated
  CheckBox 160, 185, 170, 10, "Set TIKL for 10 day return of verifcations needed", TIKL_checkbox
  EditBox 95, 205, 125, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 225, 205, 50, 15
    CancelButton 280, 205, 50, 15
  Text 10, 90, 85, 10, "Describe the discrepancy:"
  Text 50, 140, 40, 10, "Other notes:"
  Text 20, 15, 45, 10, "Case number:"
  Text 25, 115, 70, 10, "Verifications needed: "
  Text 20, 65, 70, 10, "HH memb #/PMI #(s):"
  Text 15, 165, 80, 10, "Resolution/Action taken:"
  Text 5, 40, 65, 10, "Discrepancy status:"
  GroupBox 140, 5, 190, 45, "Programs effected by the discrepancy:"
  Text 35, 210, 60, 10, "Worker signature:"
EndDialog

'The script----------------------------------------------------------------------------------------------------
'Connecing to MAXIS, establishing the county code, and grabbing the case number
EMConnect ""
get_county_code
call MAXIS_case_number_finder(MAXIS_case_number)
 										
DO										
	DO									
		err_msg = ""								'establishing value of varaible, this is necessary for the Do...LOOP	
		dialog case_discrepancy_dialog				'initial dialog			
		If buttonpressed = 0 THEN stopscript		'script ends if cancel is selected							
		IF len(MAXIS_case_number) > 8 or isnumeric(MAXIS_case_number) = false THEN err_msg = err_msg & vbNewline & "* Enter a valid case number."	'mandatory field		
		If discrepancy_status = "Select one..." then err_msg = err_msg & vbnewline & "* Select a discrepancy status."
		If (DWP_checkbox <> 1 and EMER_checkbox <> 1 and GA_checkbox <> 1 and GRH_checkbox <> 1 and MA_checkbox <> 1 and MFIP_checkbox <> 1 and MSA_checkbox <> 1 and MSP_checkbox <> 1 and RCA_checkbox <> 1 and SNAP_checkbox <> 1) then err_msg = err_msg & vbnewline & "* You must enter at least one program."	
		If MEMB_PMI = "" then err_msg = err_msg & vbnewline & "* Enter the HH member and/or PMI #'s the discrepancy effects."
		If describe_discrepancy = "" then err_msg = err_msg & vbnewline & "* Describe the discrepancy."
		If actions_taken = "" then err_msg = err_msg & vbnewline & "* Enter the resolution/actions taken."	
		If worker_signature = "" then err_msg = err_msg & vbnewline & "* Sign your case note."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect						
	LOOP UNTIL err_msg = ""									'loops until all errors are resolved
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					

'Creating an incremantal variable based on 
If DWP_checkbox = 1 then progs_effect = progs_effect & " DWP,"
If EMER_checkbox = 1 then progs_effect = progs_effect & " Emergency,"
If GA_checkbox = 1 then progs_effect = progs_effect & " GA,"
IF GRH_checkbox = 1 then progs_effect = progs_effect & " GRH,"
If MA_checkbox = 1 then progs_effect = progs_effect & " MA,"
If MFIP_checkbox = 1 then progs_effect = progs_effect & " MFIP,"
If MSP_checkbox = 1 then progs_effect = progs_effect & " Medicare savings program (MSP),"
IF MSA_checkbox = 1 then progs_effect = progs_effect & " MSA,"
If RCA_checkbox = 1 then progs_effect = progs_effect & " RCA," 
If SNAP_checkbox = 1 then progs_effect = progs_effect & " SNAP," 
'trims excess spaces of pending_progs
progs_effect = trim(progs_effect)
'takes the last comma off of pending_progs when autofilled into dialog if more more than one app date is found and additional app is selected
If right(progs_effect, 1) = "," THEN progs_effect = left(progs_effect, len(progs_effect) - 1) 

'The case notes----------------------------------------------------------------------------------------------------
start_a_blank_case_note
Call write_variable_in_CASE_NOTE("---Case discrepancy " & discrepancy_status & "---")
Call write_bullet_and_variable_in_CASE_NOTE("Programs effected by discrepancy", progs_effect)
Call write_bullet_and_variable_in_CASE_NOTE("HH member/PMI #'s", MEMB_PMI)
Call write_bullet_and_variable_in_CASE_NOTE("Description of the discrepancy", describe_discrepancy)
Call write_bullet_and_variable_in_CASE_NOTE("Verifications needed", verifs_needed)
If TIKL_checkbox = 1 then Call write_variable_in_CASE_NOTE("* TIKL'd out for 10 day return of requested verifications.")
Call write_bullet_and_variable_in_CASE_NOTE("Other notes", other_notes)
Call write_bullet_and_variable_in_CASE_NOTE("Resolution/actions taken", actions_taken)
Call write_variable_in_CASE_NOTE("---")
Call write_variable_in_CASE_NOTE(worker_signature)

if TIKL_checkbox = 1 then 
	call navigate_to_MAXIS_screen("dail", "writ")
	call create_MAXIS_friendly_date(date, 10, 5, 18) 
	Call write_variable_in_TIKL("The following verifications were requested 10 days ago for a case discrepancy: " & verifs_needed)
	transmit	
	PF3
	script_end_procedure("Success, your TIKL has been created. Please send a verification request to the client(s).")
Else 
	script_end_procedure("")
End if 