'STATS GATHERING=============================================================================================================
name_of_script = "TYPE - NEW SCRIPT TEMPLATE.vbs"       'REPLACE TYPE with either ACTIONS, BULK, DAIL, NAV, NOTES, NOTICES, or UTILITIES. The name of the script should be all caps. The ".vbs" should be all lower case.
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 1            'manual run time in seconds  -----REPLACE STATS_MANUALTIME = 1 with the anctual manualtime based on time study
STATS_denomination = "C"        'C is for each case; I is for Instance, M is for member; REPLACE with the denomonation appliable to your script.
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
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County
CALL changelog_update("11/20/24", "Initial version.", "Mark Riegel, Hennepin County") 

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT==================================================================================================================
EMConnect "" 'Connects to BlueZone
get_county_code
Call check_for_MAXIS(False)
CALL MAXIS_case_number_finder(MAXIS_case_number)

'Initial Case Number Dialog 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 221, 110, "Enter Date of Death for Household Member"
  EditBox 75, 5, 50, 15, MAXIS_case_number
  EditBox 75, 25, 20, 15, MAXIS_footer_month
  EditBox 105, 25, 20, 15, MAXIS_footer_year
  EditBox 75, 45, 140, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 125, 90, 45, 15
    PushButton 150, 5, 65, 15, "Script Instructions", msg_show_instructions_btn
    CancelButton 170, 90, 45, 15
  Text 20, 10, 50, 10, "Case Number:"
  Text 20, 30, 45, 10, "Footer month:"
  Text 10, 50, 60, 10, "Worker Signature:"
  Text 10, 65, 200, 20, "Script Purpose: Updates case based on date of death for household member in accordance with POLI/TEMP 02.08.008."
  ButtonGroup ButtonPressed
    PushButton 150, 25, 65, 15, "TE02.08.008", poli_temp_btn
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
      ' run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/ACTIONS/ACTIONS%20-%20ENTER%20DATE%20OF%20DEATH.docx"
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

'Navigate to STAT/TYPE to determine applicable programs
Call navigate_to_MAXIS_screen("STAT", "TYPE")
'Read programs for HH Memb
stat_type_row = 6
Do
  EMReadScreen type_panel_HH_memb, 2, stat_type_row, 3
  If type_panel_HH_memb = left(household_member_that_died, 2) Then
    EMReadScreen stat_type_cash, 1, stat_type_row, 28
    EMReadScreen stat_type_hc, 1, stat_type_row, 37
    EMReadScreen stat_type_snap, 1, stat_type_row, 46
    EMReadScreen stat_type_emer, 1, stat_type_row, 55
    EMReadScreen stat_type_grh, 1, stat_type_row, 64
    EMReadScreen stat_type_iv-e, 1, stat_type_row, 73
    Exit Do
  Else
    stat_type_row = stat_type_row + 1
    If stat_type_row = 19 Then Exit Do
  End If
Loop

'Determine what programs apply to HH Memb
If (stat_type_cash = "Y" OR stat_type_snap = "Y") and stat_type_hc <> "Y" Then
  'If CASH or SNAP case, but not HC case
  cash_snap_only = true
ElseIf stat_type_hc = "Y" Then
  'If HC case, doesn't matter if CASH or SNAP is active because steps are the same
  hc_case = True
Else
  'To do - handling needed here?
End If


'Navigate to STAT/MEMB to determine if: 
  '->DOD entered already
  '->DOD and DOB are the same
Call navigate_to_MAXIS_screen("STAT", "MEMB")
'Navigate to HH Memb that died
EMWriteScreen left(household_member_that_died, 2), 20, 76
transmit
EMReadScreen DOB_memb_panel, 10, 8, 42
'Convert DOB to MM/DD/YYYY format
DOB_memb_panel = replace(DOB_memb_panel, " ", "/")
If DOB_memb_panel = date_of_death then death_same_day_as_birth = True
EMReadScreen DOD_memb_panel, 10, 19, 42
'Convert DOD to MM/DD/YYYY format
DOD_memb_panel = replace(DOD_memb_panel, " ", "/")
If DOD_memb_panel = "__ __ ____" Then 
  DOD_memb_panel_blank = True   
Else
  'If the panel is not blank, need to determine if DOD on panel matches DOD entered in dialog
  DOD_memb_panel_blank = False
  If DOD_memb_panel = date_of_death Then
    DOD_memb_matches = True
  Else
    DOD_memb_matches = False
  End If
End If

'Navigate to STAT/REMO to determine if HH Memb removed already
'STAT/REMO should only be updated if HH MEMB OTHER than 01 has died
If left(household_member_that_died, 2) <> "01" Then
  Call navigate_to_MAXIS_screen("STAT", "REMO")
  EMReadScreen DOD_remo_panel, 8, 8, 53
  'Convert DOD_remo_panel to MM/DD/YYYY
  DOD_remo_panel = left(replace(DOD_remo_panel, " ", "/"), 6) & "20" & right(DOD_remo_panel, 2)
  EMReadScreen reason_code_remo_panel, 2, 8, 71
  If DOD_remo_panel = "__ __ __" then 
    DOD_remo_panel_blank = True 
  Else
    DOD_remo_panel_blank = False
    If DOD_remo_panel = date_of_death Then
      DOD_remo_matches = True
    Else
      DOD_remo_matches = False
    End If
  End If
  If reason_code_remo_panel = "01" Then 
    DOD_remo_reason_01 = True 
  Else
    DOD_remo_reason_01 = False
  End If
End If

If cash_snap_only = true Then
  'Cash and/or SNAP only case
  If left(household_member_that_died, 2) <> "01" Then
    'If HH Memb other than 01 has died, STAT/MEMB and STAT/REMO have not been updated
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 266, 85, "Enter Date of Death for HH Member - SNAP and/or CASH"
      ButtonGroup ButtonPressed
        OkButton 170, 65, 45, 15
        CancelButton 215, 65, 45, 15
      Text 5, 5, 255, 35, "The script will update STAT/MEMB and STAT/REMO with the entered date of death. If the date of death has been entered on a panel, it will overwrite the previously entered date. The script will approve a version of ineligiblity for the month after the date of death."
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

    'To do - add handling to approve ineligibility results for next month

  ElseIf left(household_member_that_died, 2) = "01" Then
    'If HH Memb other than 01 has died, STAT/MEMB has not been updated, STAT/REMO not evaluated since Memb 01 died

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 266, 85, "Enter Date of Death for HH Member - SNAP and/or CASH"
      ButtonGroup ButtonPressed
        OkButton 170, 65, 45, 15
        CancelButton 215, 65, 45, 15
      Text 5, 5, 255, 35, "The script will update STAT/MEMB with the entered date of death. If the date of death has been entered on a panel, it will overwrite the previously entered date. The script will approve a version of ineligiblity for the month after the date of death."
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

    'To do - add handling to approve ineligibility results for next month
  End If

ElseIf hc_case = True Then
  'Healthcare case
  Dialog1 = ""
  BeginDialog Dialog1, 0, 0, 266, 75, "Enter Date of Death for HH Member - Healthcare Case"
    DropListBox 100, 5, 160, 15, "Select One..."+chr(9)+"MDH Minnesota Death Search"+chr(9)+"Social Security Administration record (SOLQ-I)"+chr(9)+"Authorized Representative"+chr(9)+"Power of Attorney"+chr(9)+"Other Adult Family Member", death_verification
    EditBox 100, 25, 160, 15, who_reported_death
    ButtonGroup ButtonPressed
      PushButton 5, 55, 65, 15, "TE02.08.008", poli_temp_btn
      OkButton 170, 50, 45, 15
      CancelButton 215, 50, 45, 15
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

  '...and enters a title (replace variables with your own content)...
  CALL write_variable_in_case_note("*** Date of Death Verified - HH Memb " & left(household_member_that_died, 2) & " ***")

  '...some editboxes or droplistboxes (replace variables with your own content)...
  CALL write_bullet_and_variable_in_case_note( "Death verified by", death_verification)
  CALL write_bullet_and_variable_in_case_note( "Death reported by", who_reported_death)

  '...and a worker signature.
  CALL write_variable_in_case_note("---")
  CALL write_variable_in_case_note(worker_signature)

  'Save CASE/NOTE and navigate to SELF
  back_to_SELF

  If left(household_member_that_died, 2) <> "01" Then
    'If HH Memb other than 01 has died, STAT/MEMB and STAT/REMO have not been updated
    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 266, 85, "Enter Date of Death for HH Member - SNAP and/or CASH"
      ButtonGroup ButtonPressed
        OkButton 170, 65, 45, 15
        CancelButton 215, 65, 45, 15
      Text 5, 5, 255, 35, "The script will update STAT/MEMB and STAT/REMO with the entered date of death. If the date of death has been entered on a panel, it will overwrite the previously entered date. The script will approve a version of ineligiblity for the month after the date of death."
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

    'To do - add handling to approve ineligibility results for next month

  ElseIf left(household_member_that_died, 2) = "01" Then
    'If HH Memb other than 01 has died, STAT/MEMB has not been updated, STAT/REMO not evaluated since Memb 01 died

    Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 266, 85, "Enter Date of Death for HH Member - SNAP and/or CASH"
      ButtonGroup ButtonPressed
        OkButton 170, 65, 45, 15
        CancelButton 215, 65, 45, 15
      Text 5, 5, 255, 35, "The script will update STAT/MEMB with the entered date of death. If the date of death has been entered on a panel, it will overwrite the previously entered date. The script will approve a version of ineligiblity for the month after the date of death."
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

    'To do - add handling to approve ineligibility results for next month
    
  End If
End If

'End the script. Put any success messages in between the quotes, *always* starting with the word "Success!"
script_end_procedure("The script has updated the date of death for the selected HH member.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation - Version date 01/12/2023
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------
'--Tab orders reviewed & confirmed----------------------------------------------
'--Mandatory fields all present & Reviewed--------------------------------------
'--All variables in dialog match mandatory fields-------------------------------
'Review dialog names for content and content fit in dialog----------------------
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------
'--CASE:NOTE Header doesn't look funky------------------------------------------
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------
'--write_variable_in_CASE_NOTE function: confirm that proper punctuation is used -----------------------------------
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------
'--MAXIS_background_check reviewed (if applicable)------------------------------
'--PRIV Case handling reviewed -------------------------------------------------
'--Out-of-County handling reviewed----------------------------------------------
'--script_end_procedures (w/ or w/o error messaging)----------------------------
'--BULK - review output of statistics and run time/count (if applicable)--------
'--All strings for MAXIS entry are uppercase vs. lower case (Ex: "X")-----------
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------
'--Incrementors reviewed (if necessary)-----------------------------------------
'--Denomination reviewed -------------------------------------------------------
'--Script name reviewed---------------------------------------------------------
'--BULK - remove 1 incrementor at end of script reviewed------------------------

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------
'--Remove testing message boxes-------------------------------------------------
'--Remove testing code/unnecessary code-----------------------------------------
'--Review/update SharePoint instructions----------------------------------------
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------
'--COMPLETE LIST OF SCRIPTS update policy references----------------------------
'--Complete misc. documentation (if applicable)---------------------------------
'--Update project team/issue contact (if applicable)----------------------------

