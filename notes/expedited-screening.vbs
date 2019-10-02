'Required for statistical purposes==========================================================================================
name_of_script = "NOTES - EXPEDITED SCREENING.vbs"
start_time = timer
STATS_counter = 1               'sets the stats counter at one
STATS_manualtime = 180          'manual run time in seconds
STATS_denomination = "C"        'C is for each case
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
CALL changelog_update("09/01/2019", "Updated Utility standards that go into effect for 10/01/2019. Added application date field for accurate expedited screening.", "Ilse Ferris, Hennepin County")
CALL changelog_update("09/01/2018", "Updated Utility standards that go into effect for 10/01/2018.", "Ilse Ferris, Hennepin County")
call changelog_update("09/25/2017", "Updated HEST standards for 10/17 standard changes.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog exp_screening_dialog, 0, 0, 186, 195, "Expedited Screening Dialog"
  EditBox 40, 10, 50, 15, MAXIS_case_number
  EditBox 130, 10, 50, 15, application_date
  EditBox 105, 35, 50, 15, income
  EditBox 105, 55, 50, 15, assets
  EditBox 105, 75, 50, 15, rent
  CheckBox 20, 105, 55, 10, "Heat (or AC)", heat_AC_check
  CheckBox 80, 105, 45, 10, "Electricity", electric_check
  CheckBox 135, 105, 35, 10, "Phone", phone_check
  EditBox 75, 125, 105, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 75, 145, 50, 15
    CancelButton 130, 145, 50, 15
  Text 10, 40, 95, 10, "Income received this month:"
  Text 10, 60, 95, 10, "Cash, checking, or savings: "
  Text 10, 80, 90, 10, "AMT paid for rent/mortgage:"
  GroupBox 10, 95, 170, 25, "Utilities claimed (check below):"
  Text 10, 130, 60, 10, "Worker signature:"
  Text 10, 15, 25, 10, "Case #: "
  GroupBox 5, 160, 175, 30, "**IMPORTANT**"
  Text 15, 170, 160, 15, "The income, assets and shelter costs fields will default to $0 if left blank. "
  Text 95, 15, 35, 10, "App Date:"
  GroupBox 5, 0, 180, 30, ""
EndDialog

'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""    'Connecting to BlueZone
call MAXIS_case_number_finder(MAXIS_case_number) 'It will search for a case number.
' application_date = date & ""
If MAXIS_case_number <> "" Then
    Call navigate_to_MAXIS_screen("STAT", "PROG")
    EMReadScreen snap_pend_check, 4, 10, 74
    If snap_pend_check = "PEND" Then
        EMReadScreen snap_app_date, 8, 10, 33
        application_date = replace(snap_app_date, " ", "/")
    End If
    transmit
    EMReadScreen check_for_hcre, 4, 2, 50
    If check_for_hcre = "HCRE" Then
        PF10
    End If
End If

'Shows the dialog
Do
	Do
        err_msg = ""							'establishing value of variable, this is necessary for the Do...LOOP
        dialog exp_screening_dialog				'main dialog
        If buttonpressed = 0 THEN stopscript	'script ends if cancel is selected
        IF len(MAXIS_case_number) > 8 or isnumeric(MAXIS_case_number) = false then err_msg = err_msg & vbCr & "* Enter a valid case number."		'mandatory field
        If isdate(application_date) = False then err_msg = err_msg & vbCr & "* Enter a valid applcation date."
        If (trim(income) <> "" and isnumeric(income) = false) then err_msg = err_msg & vbCr & "* The income fields must be numeric only. Do not put letters or symbols in these sections."
        If (trim(assets) <> "" and isnumeric(assets) = false) then err_msg = err_msg & vbCr & "* The assets fields must be numeric only. Do not put letters or symbols in these sections."
        If (trim(rent) <> "" and isnumeric(rent) = false) then err_msg = err_msg & vbCr & "* The rent fields must be numeric only. Do not put letters or symbols in these sections."
        If trim(worker_signature) = "" then err_msg = err_msg & vbCr & "* Enter your worker signature."
        If err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & VbCr & err_msg & VbCr		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call check_for_MAXIS(FALSE) 'checking for an active MAXIS session

'DATE BASED LOGIC FOR UTILITY AMOUNTS------------------------------------------------------------------------------------------
If application_date >= cdate("10/01/2019") then			'these variables need to change every October
    heat_AC_amt = 490
    electric_amt = 143
    phone_amt = 49
else
    heat_AC_amt = 493
    electric_amt = 126
    phone_amt = 47
End if

'LOGIC AND CALCULATIONS----------------------------------------------------------------------------------------------------
'Logic for figuring out utils. The highest priority for the if...then is heat/AC, followed by electric and phone, followed by phone and electric separately.
If heat_AC_check = checked then
	utilities = heat_AC_amt
ElseIf electric_check = checked and phone_check = checked then
	utilities = phone_amt + electric_amt					'Phone standard plus electric standard.
ElseIf phone_check = checked and electric_check = unchecked then
	utilities = phone_amt
ElseIf electric_check = checked and phone_check = unchecked then
	utilities = electric_amt
End if

'in case no options are clicked, utilities are set to zero.
If phone_check = unchecked and electric_check = unchecked and heat_AC_check = unchecked then utilities = 0

'If nothing is written for income/assets/rent info, we set to zero.
If trim(income) = "" then income = 0
If trim(assets) = "" then assets = 0
If trim(rent) = "" then rent = 0

'Calculates expedited status based on above numbers
If (int(income) < 150 and int(assets) <= 100) or ((int(income) + int(assets)) < (int(rent) + cint(utilities))) then expedited_status = "client appears expedited"
If (int(income) + int(assets) >= int(rent) + cint(utilities)) and (int(income) >= 150 or int(assets) > 100) then expedited_status = "client does not appear expedited"
'----------------------------------------------------------------------------------------------------

'Navigates to STAT/DISQ using current month as footer month. If it can't get in to the current month due to CAF received in a different month, it'll find that month and navigate to it.
Call navigate_to_MAXIS_screen("STAT", "DISQ")
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year) 'grabbing footer month and year

'Reads the DISQ info for the case note.
EMReadScreen DISQ_member_check, 34, 24, 2
If DISQ_member_check = "DISQ DOES NOT EXIST FOR ANY MEMBER" then
	has_DISQ = False
Else
	has_DISQ = True
End if

'Reads MONY/DISB to see if EBT account is open
IF expedited_status = "client appears expedited" THEN
	Call navigate_to_MAXIS_screen("MONY", "DISB")
	EMReadScreen EBT_account_status, 1, 14, 27
END IF

'THE CASE NOTE----------------------------------------------------------------------------------------------------
	call navigate_to_MAXIS_screen("case", "note")
	PF9

	EMReadScreen case_note_check, 17, 2, 33
	EMReadScreen mode_check, 1, 20, 09
	If case_note_check <> "Case Notes (NOTE)" or mode_check <> "A" then    'this will account for those cases when the script is run on an out of county case.
		msgbox "The script can't open a case note. You may be in inquiry or entered a case number that is in another county." &_
		vbNewLine & vbNewLine & "This result for this case is " & expedited_status & vbNewLine & vbNewLine & "Please run the script again if you were in inquiry to add a case note."
		script_end_procedure("")
	else
		'Body of the case note
        Call write_variable_in_CASE_NOTE("~ Received Application for SNAP, " & expedited_status & " ~")
		call write_variable_in_CASE_NOTE("---")
		call write_variable_in_CASE_NOTE("     CAF 1 income claimed this month: $" & income)
		call write_variable_in_CASE_NOTE("         CAF 1 liquid assets claimed: $" & assets)
		call write_variable_in_CASE_NOTE("         CAF 1 rent/mortgage claimed: $" & rent)
		call write_variable_in_CASE_NOTE("        Utilities (amt/HEST claimed): $" & utilities)
		call write_variable_in_CASE_NOTE("---")
		If has_DISQ = True then call write_variable_in_CASE_NOTE("A DISQ panel exists for someone on this case.")
		If has_DISQ = False then call write_variable_in_CASE_NOTE("No DISQ panels were found for this case.")
		If expedited_status = "client appears expedited" AND EBT_account_status = "Y" then call write_variable_in_CASE_NOTE("* EBT Account IS open. Recipient will NOT be able to get a replacement card in the agency.  Rapid Electronic Issuance (REI) with caution.")
		If expedited_status = "client appears expedited" AND EBT_account_status = "N" then call write_variable_in_CASE_NOTE("* EBT Account is NOT open. Recipient is able to get initial card in the agency. Rapid Electronic Issuance (REI) can be used, but only to avoid an emergency issuance or to meet EXP criteria.")
		call write_variable_in_CASE_NOTE("---")
		call write_variable_in_CASE_NOTE(worker_signature)
		If expedited_status = "client appears expedited" then
			MsgBox "This client appears expedited. A same day interview needs to be offered."
		End if
		If expedited_status = "client does not appear expedited" then
			MsgBox "This client does not appear expedited. A same day interview does not need to be offered."
		End if
	End if
script_end_procedure("")
