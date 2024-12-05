'STATS GATHERING=============================================================================================================
name_of_script = "ACTIONS - ENTER DATE OF DEATH.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 150                	'manual run time in seconds
STATS_denomination = "C"       		      'C is for each CASE
'END OF stats block==========================================================================================================

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

'CHANGELOG BLOCK===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
CALL changelog_update("11/20/24", "Initial version.", "Mark Riegel, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT==================================================================================================================
EMConnect "" 'Connects to BlueZone
get_county_code
Call check_for_MAXIS(False)
CALL MAXIS_case_number_finder(MAXIS_case_number)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)

'Setting initial variables
active_pending_CASH_SNAP = False
active_pending_HC = False
active_pending_case = False

'Initial Case Number Dialog 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 221, 110, "Enter Date of Death for Household Member"
  EditBox 75, 5, 50, 15, MAXIS_case_number
  EditBox 75, 25, 20, 15, MAXIS_footer_month
  EditBox 105, 25, 20, 15, MAXIS_footer_year
  EditBox 75, 45, 140, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 125, 90, 45, 15
    CancelButton 170, 90, 45, 15
    PushButton 150, 5, 65, 15, "Script Instructions", msg_show_instructions_btn
    PushButton 150, 25, 65, 15, "TE02.08.008", poli_temp_btn
  Text 20, 10, 50, 10, "Case Number:"
  Text 20, 30, 45, 10, "Footer month:"
  Text 10, 50, 60, 10, "Worker Signature:"
  Text 10, 65, 200, 20, "Script Purpose: Updates case based on date of death for household member in accordance with POLI/TEMP 02.08.008."
EndDialog

Do 
  Do
    err_msg = ""
    Dialog Dialog1
    cancel_without_confirmation()
    Call validate_MAXIS_case_number(err_msg, "*")
    Call validate_footer_month_entry(MAXIS_footer_month, MAXIS_footer_year, err_msg, "*")
    If trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Sign your case note."
    
    If ButtonPressed = msg_show_instructions_btn Then 
      err_msg = "LOOP"
      'Add in link to instructions once created
      run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/ACTIONS/ACTIONS%20-%20ENTER%20DATE%20OF%20DEATH.docx"
    End If
    If ButtonPressed = poli_temp_btn Then
      run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/sites/hs-es-poli-temp/Documents%203/Forms/AllItems.aspx?id=%2Fsites%2Fhs%2Des%2Dpoli%2Dtemp%2FDocuments%203%2FTE%2002%2E08%2E008%20CLOSING%20MAXIS%20AND%20MMIS%20DUE%20TO%20DEATH%2Epdf&parent=%2Fsites%2Fhs%2Des%2Dpoli%2Dtemp%2FDocuments%203"
    End If 
    IF err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
  Loop until err_msg = ""
  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Create list of household members
Call generate_client_list(list_of_household_members, "Select One ...")          'Using the client list functionality the script will read STAT for all the household members to populate droplist box

'Date of Death for Household Member Dialog 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 256, 75, "Enter Date of Death for Household Member"
  DropListBox 110, 5, 140, 20, list_of_household_members, household_member_that_died
  EditBox 110, 25, 45, 15, date_of_death
  ButtonGroup ButtonPressed
    OkButton 155, 55, 45, 15
    CancelButton 205, 55, 45, 15
  Text 5, 5, 100, 10, "Household Member that Died:"
  Text 55, 30, 50, 10, "Date of Death:"
  Text 160, 25, 50, 10, "(MM/DD/YYYY)"
EndDialog

Do 
  Do
    err_msg = ""
    Dialog Dialog1
    cancel_without_confirmation()
    If household_member_that_died = "Select One ..." THEN err_msg = err_msg & vbCr & "* Please select the household member that has died."
    If len(date_of_death) <> 10 or IsDate(date_of_death) = False THEN err_msg = err_msg & vbCr & "* Please enter the date of death in the format MM/DD/YYYY."
    If IsDate(date_of_death) Then
      If DateDiff("D", date_of_death, date) < 0 Then err_msg = err_msg & vbCr & "* The date of death cannot be in the future."
    End If
    IF err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
  Loop until err_msg = ""
  CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Convert date of death to MAXIS friendly date
date_of_death_maxis_format = replace(date_of_death, "/", " ")

'Determine which programs are active or pending
Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)

'Determine if case is active, if it is active then evaluate what programs are active or pending
If case_active = True or case_pending = True Then
  active_pending_case = True
  If instr(list_active_programs, "SNAP") OR _
    instr(list_active_programs, "MSA") OR _
    instr(list_active_programs, "MFIP") OR _
    instr(list_active_programs, "GA") OR _
    instr(list_active_programs, "DWP") OR _
    instr(list_active_programs, "EGA") OR _
    instr(list_active_programs, "EA") OR _
    instr(list_pending_programs, "GA") OR _
    instr(list_pending_programs, "MSA") OR _
    instr(list_pending_programs, "MFIP") OR _
    instr(list_pending_programs, "DWP") OR _
    instr(list_pending_programs, "FS") OR _
    instr(list_pending_programs, "MA") OR _
    instr(list_pending_programs, "MSP") OR _
    instr(list_pending_programs, "CASH") OR _
    instr(list_pending_programs, "SNAP") OR _
    instr(list_pending_programs, "EA") Then
      active_pending_CASH_SNAP = True
  ElseIf instr(list_active_programs, "HC") OR instr(list_pending_programs, "HC") Then
    active_pending_HC = True
  Else
    active_pending_case = False
  End If
Else
  active_pending_case = False
End If

'If case is not active or pending, then script will end
If active_pending_case = False Then script_end_procedure("The case is not active and no programs are pending. The script will now end.")

If active_pending_HC = True Then
  'Healthcare case
  Dialog1 = ""
  BeginDialog Dialog1, 0, 0, 266, 75, "Enter Date of Death for HH Member - Healthcare Case"
    DropListBox 100, 5, 160, 15, "Select One..."+chr(9)+"MDH Minnesota Death Search"+chr(9)+"Social Security Administration record (SOLQ-I)"+chr(9)+"Authorized Representative"+chr(9)+"Power of Attorney"+chr(9)+"Other Adult Family Member", death_verification
    EditBox 100, 25, 160, 15, who_reported_death
    ButtonGroup ButtonPressed
      OkButton 170, 50, 45, 15
      CancelButton 215, 50, 45, 15
      PushButton 5, 55, 65, 15, "TE02.08.008", poli_temp_btn
    Text 5, 10, 95, 10, "How was the death verified?"
    Text 5, 30, 85, 10, "Who reported the death?"
  EndDialog

  Do 
    Do
      err_msg = ""
      Dialog Dialog1
      cancel_without_confirmation()
      If ButtonPressed = poli_temp_btn Then
        run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/sites/hs-es-poli-temp/Documents%203/Forms/AllItems.aspx?id=%2Fsites%2Fhs%2Des%2Dpoli%2Dtemp%2FDocuments%203%2FTE%2002%2E08%2E008%20CLOSING%20MAXIS%20AND%20MMIS%20DUE%20TO%20DEATH%2Epdf&parent=%2Fsites%2Fhs%2Des%2Dpoli%2Dtemp%2FDocuments%203"
      End If 
      If trim(death_verification) = "Select One..." Then err_msg = err_msg & vbCr & "* Please indicate how you verified the death."
      If trim(who_reported_death) = "" Then err_msg = err_msg & vbCr & "* Please indicate who reported the death."
      IF err_msg <> "" and err_msg <> "LOOP" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
  Loop until are_we_passworded_out = false					'loops until user passwords back in

  'Navigate to SELF to update footer month to match the month of death
  back_to_SELF
  MAXIS_footer_month = left(date_of_death_maxis_format, 2)
  MAXIS_footer_year = right(date_of_death_maxis_format, 2)
  EMWriteScreen left(date_of_death_maxis_format, 2), 20, 43
  EMWriteScreen right(date_of_death_maxis_format, 2), 20, 46

  Call navigate_to_MAXIS_screen("CASE", "NOTE")

  'Now it navigates to a blank case note
  Call start_a_blank_case_note

  '...and enters a CASE/NOTE with the worker signature
  CALL write_variable_in_case_note("*** Date of Death Verified - HH Memb " & left(household_member_that_died, 2) & " ***")
  CALL write_bullet_and_variable_in_case_note( "Death verified by", death_verification)
  CALL write_bullet_and_variable_in_case_note( "Death reported by", who_reported_death)
  CALL write_variable_in_case_note("---")
  CALL write_variable_in_case_note(worker_signature)

  'Save CASE/NOTE and navigate to SELF
  back_to_SELF
End If

If left(household_member_that_died, 2) <> "01" Then
  'If HH Memb other than 01 has died, STAT/MEMB and STAT/REMO need to be updated
  Dialog1 = ""
  BeginDialog Dialog1, 0, 0, 266, 85, "Enter Date of Death for HH Member"
    ButtonGroup ButtonPressed
      OkButton 170, 65, 45, 15
      CancelButton 215, 65, 45, 15
    Text 5, 5, 255, 35, "The script will update STAT/MEMB and STAT/REMO with the entered date of death. If the date of death has been entered on a panel, it will overwrite the previously entered date."
    ButtonGroup ButtonPressed
      PushButton 5, 65, 65, 15, "TE02.08.008", poli_temp_btn
    Text 5, 45, 75, 10, "Click 'OK' to proceed."
  EndDialog

  Do
    Dialog Dialog1

    If ButtonPressed = poli_temp_btn Then
      run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/sites/hs-es-poli-temp/Documents%203/Forms/AllItems.aspx?id=%2Fsites%2Fhs%2Des%2Dpoli%2Dtemp%2FDocuments%203%2FTE%2002%2E08%2E008%20CLOSING%20MAXIS%20AND%20MMIS%20DUE%20TO%20DEATH%2Epdf&parent=%2Fsites%2Fhs%2Des%2Dpoli%2Dtemp%2FDocuments%203"
    End If 

    cancel_without_confirmation()
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
  Loop until are_we_passworded_out = false					'loops until user passwords back in

  'Navigate to SELF to update footer month to match the month of death
  back_to_SELF
  MAXIS_footer_month = left(date_of_death_maxis_format, 2)
  MAXIS_footer_year = right(date_of_death_maxis_format, 2)
  EMWriteScreen left(date_of_death_maxis_format, 2), 20, 43
  EMWriteScreen right(date_of_death_maxis_format, 2), 20, 46

  Call navigate_to_MAXIS_screen("STAT", "MEMB")
  'Navigate to HH Memb that died and put the panel in edit mode
  EMWriteScreen left(household_member_that_died, 2), 20, 76
  transmit
  PF9   'Put panel in edit mode
  EMWriteScreen left(date_of_death_maxis_format, 2), 19, 42
  EMWriteScreen mid(date_of_death_maxis_format, 4, 2), 19, 45
  EMWriteScreen right(date_of_death_maxis_format, 4), 19, 48
  transmit  'Save update
  transmit  'Bypass warning message
  
  'Navigate to STAT/REMO to update with death info
  Call navigate_to_MAXIS_screen("STAT", "REMO")
  EMWriteScreen left(household_member_that_died, 2), 20, 76
  EMWriteScreen "NN", 20, 79
  transmit
  PF9   'Put panel in edit mode
  'Write date of death to panel and enter reason code '01'
  EMWriteScreen left(date_of_death_maxis_format, 2), 8, 53
  EMWriteScreen mid(date_of_death_maxis_format, 4, 2), 8, 56
  EMWriteScreen right(date_of_death_maxis_format, 2), 8, 59
  EMWriteScreen "01", 8, 71
  transmit

ElseIf left(household_member_that_died, 2) = "01" Then
  'If HH Memb 01 has died, STAT/MEMB only needs to be updated

  Dialog1 = ""
  BeginDialog Dialog1, 0, 0, 266, 85, "Enter Date of Death for HH Member"
    ButtonGroup ButtonPressed
      OkButton 170, 65, 45, 15
      CancelButton 215, 65, 45, 15
    Text 5, 5, 255, 35, "The script will update STAT/MEMB with the entered date of death. If the date of death has been entered on a panel, it will overwrite the previously entered date."
    ButtonGroup ButtonPressed
      PushButton 5, 65, 65, 15, "TE02.08.008", poli_temp_btn
    Text 5, 45, 75, 10, "Click 'OK' to proceed."
  EndDialog

  Do
    Dialog Dialog1

    If ButtonPressed = poli_temp_btn Then
      run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/sites/hs-es-poli-temp/Documents%203/Forms/AllItems.aspx?id=%2Fsites%2Fhs%2Des%2Dpoli%2Dtemp%2FDocuments%203%2FTE%2002%2E08%2E008%20CLOSING%20MAXIS%20AND%20MMIS%20DUE%20TO%20DEATH%2Epdf&parent=%2Fsites%2Fhs%2Des%2Dpoli%2Dtemp%2FDocuments%203"
    End If 

    cancel_without_confirmation()
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
  Loop until are_we_passworded_out = false					'loops until user passwords back in

  'Navigate to SELF to update footer month to match the month of death
  back_to_SELF
  MAXIS_footer_month = left(date_of_death_maxis_format, 2)
  MAXIS_footer_year = right(date_of_death_maxis_format, 2)
  EMWriteScreen left(date_of_death_maxis_format, 2), 20, 43
  EMWriteScreen right(date_of_death_maxis_format, 2), 20, 46

  Call navigate_to_MAXIS_screen("STAT", "MEMB")
  'Navigate to HH Memb that died and put the panel in edit mode
  EMWriteScreen left(household_member_that_died, 2), 20, 76
  transmit
  PF9   'Put panel in edit mode
  EMWriteScreen left(date_of_death_maxis_format, 2), 19, 42
  EMWriteScreen mid(date_of_death_maxis_format, 4, 2), 19, 45
  EMWriteScreen right(date_of_death_maxis_format, 4), 19, 48
  transmit  'Save update
  transmit  'Bypass warning message

End If

script_end_procedure("Success! The script has updated the date of death for the selected HH member. Please review the case and approve eligibility results if appropriate.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------12/05/2024
'--Tab orders reviewed & confirmed----------------------------------------------12/05/2024
'--Mandatory fields all present & Reviewed--------------------------------------12/05/2024
'--All variables in dialog match mandatory fields-------------------------------12/05/2024
'Review dialog names for content and content fit in dialog----------------------12/05/2024
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------12/05/2024
'--CASE:NOTE Header doesn't look funky------------------------------------------12/05/2024
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------12/05/2024
'--write_variable_in_CASE_NOTE function: confirm proper punctuation is used-----12/05/2024
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------12/05/2024
'--MAXIS_background_check reviewed (if applicable)------------------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------N/A
'--Out-of-County handling reviewed----------------------------------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------12/05/2024
'--BULK - review output of statistics and run time/count (if applicable)--------N/A
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------12/05/2024
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------12/05/2024
'--Incrementors reviewed (if necessary)-----------------------------------------12/05/2024
'--Denomination reviewed -------------------------------------------------------12/05/2024
'--Script name reviewed---------------------------------------------------------12/05/2024
'--BULK - remove 1 incrementor at end of script reviewed------------------------N/A

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------12/05/2024
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------12/05/2024
'--Remove testing code/unnecessary code-----------------------------------------12/05/2024
'--Review/update SharePoint instructions----------------------------------------12/05/2024
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------

