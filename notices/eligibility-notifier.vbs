'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "NOTICES - ELIGIBILITY NOTIFIER.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 195                               'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
'END OF stats block==============================================================================================

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
call changelog_update("03/18/2022", "Removed ApplyMN website, replaced it with MNbenefits website.", "Ilse Ferris, Hennepin County")
call changelog_update("05/28/2020", "Update to the notice wording, added virtual drop box information.", "MiKayla Handley, Hennepin County")
CALL changelog_update("06/25/2019", "Updated the verbiage in the notice to be more informative and clear.", "Casey Love, Hennepin County")
CALL changelog_update("12/29/2017", "Coordinates for sending MEMO's has changed in SPEC function. Updated script to support change.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""
'Searches for a case number
call MAXIS_case_number_finder(MAXIS_case_number)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 181, 120, "Potential Eligibility MEMO"
  EditBox 75, 5, 50, 15, MAXIS_case_number
  CheckBox 10, 30, 30, 10, "SNAP", SNAP_checkbox
  CheckBox 55, 30, 30, 10, "CASH", CASH_checkbox
  CheckBox 100, 30, 25, 10, "MA", MA_checkbox
  CheckBox 140, 30, 30, 10, "MSP", MSP_checkbox
  DropListBox 100, 55, 75, 10, ""+chr(9)+"Apply in MAXIS"+chr(9)+"Apply in MNSure", HC_apply_method
  EditBox 90, 75, 60, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 30, 95, 50, 15
    CancelButton 90, 95, 50, 15
  Text 25, 10, 50, 10, "Case number:"
  Text 20, 80, 65, 10, "Worker signature:"
  Text 10, 50, 85, 20, "If HC was checked please pick system to apply in:"
EndDialog

'This Do...loop shows the appointment letter dialog, and contains logic to require most fields.
DO
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
		If SNAP_checkbox <> checked AND CASH_checkbox <> checked AND MA_checkbox <> checked AND MSP_checkbox <> checked THEN err_msg = err_msg & "Please select a program." & vbNewLine
		If MSP_checkbox = checked AND HC_apply_method <> "Apply in MAXIS" THEN err_msg = err_msg & "You selected MSP, at this time you cannot apply in Mnsure if you have Medicare. Please review selections" & vbNewLine
		If (MSP_checkbox = checked or MA_checkbox = checked) AND HC_apply_method = "" THEN err_msg = err_msg & "You selected a HC program, please select a system to apply in." & vbNewLine
		If isnumeric(MAXIS_case_number) = False or len(MAXIS_case_number) > 8 then err_msg = err_msg & "You must fill in a valid case number." & vbNewLine
		If worker_signature = "" then err_msg = err_msg & "You must sign your case note." & vbNewLine
		IF err_msg <> "" THEN msgbox err_msg
	Loop until err_msg = ""
	call check_for_password(are_we_passworded_out)  'Adding functionality for MAXIS v.6 Passworded Out issue'
LOOP UNTIL are_we_passworded_out = false

Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)  ' start the memo writing process

'formatting variable
IF SNAP_checkbox = checked THEN progs_to_apply_in_maxis = "SNAP or "
IF CASH_checkbox = checked THEN progs_to_apply_in_maxis = progs_to_apply_in_maxis & "CASH or "
IF MA_checkbox = checked AND HC_apply_method = "Apply in MAXIS" THEN progs_to_apply_in_maxis = progs_to_apply_in_maxis & "MA(in MAXIS) or "
IF MA_checkbox = checked AND HC_apply_method = "Apply in MNSure" THEN progs_to_apply_in_maxis = progs_to_apply_in_maxis & "MA(in MNSure) or "
IF MSP_checkbox = checked THEN progs_to_apply_in_maxis = progs_to_apply_in_maxis & "MSP or "
progs_to_apply_in_maxis = left(progs_to_apply_in_maxis,(len(progs_to_apply_in_maxis) - "3"))

'Writes the MEMO.
'call write_variable_in_SPEC_MEMO("***********************************************************")
IF SNAP_checkbox = checked THEN call write_variable_in_SPEC_MEMO("You may be eligible for the Supplemental Nutritional Assistance Program (SNAP).")
IF CASH_checkbox = checked THEN call write_variable_in_SPEC_MEMO("You may be eligible for CASH assistance.")
IF MA_checkbox = checked THEN call write_variable_in_SPEC_MEMO("You may be eligible for Medical assistance(MA).")
IF MSP_checkbox = checked THEN call write_variable_in_SPEC_MEMO("You may be eligible for the Medicare Savings Program (MSP).")
call write_variable_in_SPEC_MEMO("")
IF HC_apply_method = "Apply in MNSure" THEN
	call write_variable_in_SPEC_MEMO("To apply for MA you can apply online at MNSURE.ORG.")
	call write_variable_in_SPEC_MEMO("")
END IF
IF SNAP_checkbox = checked or CASH_checkbox = checked or HC_apply_method = "Apply in MAXIS" THEN

    IF SNAP_checkbox = checked THEN these_progs = "SNAP or "
    IF CASH_checkbox = checked THEN these_progs = these_progs & "CASH or "
    IF MA_checkbox = checked AND HC_apply_method = "Apply in MAXIS" THEN these_progs = these_progs & "MA or "
    these_progs = left(these_progs,(len(these_progs) - "3"))

    call write_variable_in_SPEC_MEMO("To apply for " & these_progs & "apply online at mnbenefits.mn.gov")
    call write_variable_in_SPEC_MEMO("")
End If
call write_variable_in_SPEC_MEMO("You can always apply for any program by contacting Hennepin County at 612-596-1300 to request a paper application.") ', or complete an application at any of the Human Service Centers:"'

call write_variable_in_SPEC_MEMO("")
CALL digital_experience
'Call write_variable_in_SPEC_MEMO("- 7051 Brooklyn Blvd Brooklyn Center 55429")
'Call write_variable_in_SPEC_MEMO("- 1011 1st St S Hopkins 55343")
'Call write_variable_in_SPEC_MEMO("- 9600 Aldrich Ave S Bloomington 55420")
'Call write_variable_in_SPEC_MEMO("- 1001 Plymouth Ave N Minneapolis 55411")
'Call write_variable_in_SPEC_MEMO("- 525 Portland Ave S Minneapolis 55415")
'Call write_variable_in_SPEC_MEMO("- 2215 East Lake Street Minneapolis 55407")
call write_variable_in_SPEC_MEMO("")
IF SNAP_checkbox = checked or CASH_checkbox = checked THEN
	call write_variable_in_SPEC_MEMO("When applying for SNAP and/or CASH you can submit the first page of the paper application to set your date of application. Your first month's benefit amount will be prorated based on your application date.")
	call write_variable_in_SPEC_MEMO("")
END IF
' If SNAP_checkbox = checked Then
'     Call write_variable_in_SPEC_MEMO("If your income and assets are less than your monthly shelter expenses or your income and assets are very low, you may quality to have your SNAP application processing expedited, come in right away to complete your application and interview.")
'     Call write_variable_in_SPEC_MEMO("")
' End If

call write_variable_in_SPEC_MEMO("This is a notice to inform you of programs you might have eligibility for. An application must be submitted and elibility will be determined during the application process.")
'call write_variable_in_SPEC_MEMO("***********************************************************")
'Exits the MEMO
PF4

'Navigates to CASE/NOTE and starts a blank one
start_a_blank_CASE_NOTE

'Writes the case note--------------------------------------------
call write_variable_in_CASE_NOTE("**Potential Eligibility Notice Sent**")
call write_bullet_and_variable_in_CASE_NOTE("Programs client may be eligible for", progs_to_apply_in_maxis)
If forms_to_arep = "Y" then call write_variable_in_CASE_NOTE("* Copy of notice sent to AREP.")              'Defined above
If forms_to_swkr = "Y" then call write_variable_in_CASE_NOTE("* Copy of notice sent to Social Worker.")     'Defined above
call write_variable_in_CASE_NOTE("---")
call write_variable_in_CASE_NOTE(worker_signature)

script_end_procedure("")
