'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "bulk-applications.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 335                      'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

' 'Reading Locally held FuncLib in leiu of issues with connecting to GitHub
' Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
' Set fso_command = run_another_script_fso.OpenTextFile("C:\MAXIS-scripts\MASTER FUNCTIONS LIBRARY.vbs")
' text_from_the_other_script = fso_command.ReadAll
' fso_command.Close
' Execute text_from_the_other_script

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("02/05/2018", "Initial version.", "MiKayla Handley, Hennepin County")


'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------Custom function
function start_a_new_spec_memo_and_continue(success_var)
'--- This function navigates user to SPEC/MEMO and starts a new SPEC/MEMO, selecting client, AREP, and SWKR if appropriate
'===== Keywords: MAXIS, notice, navigate, edit
  success_var = True
	call navigate_to_MAXIS_screen("SPEC", "MEMO")				'Navigating to SPEC/MEMO

	PF5															'Creates a new MEMO. If it's unable the script will stop.
	EMReadScreen memo_display_check, 12, 2, 33
	If memo_display_check = "Memo Display" then success_var = False

	'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
	    arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
	    call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
	    EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
	    PF5                                                     'PF5s again to initiate the new memo process
	END IF
	'Checking for SWKR
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
	    swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
	    call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
	    EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
	    PF5                                           'PF5s again to initiate the new memo process
	END IF
	EMWriteScreen "x", 5, 12                                        'Initiates new memo to client
	IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	transmit                                                        'Transmits to start the memo writing process
end function

function ONLY_create_MAXIS_friendly_date(date_variable)
'--- This function creates a MM DD YY date.
'~~~~~ date_variable: the name of the variable to output
	var_month = datepart("m", date_variable)
	If len(var_month) = 1 then var_month = "0" & var_month
	var_day = datepart("d", date_variable)
	If len(var_day) = 1 then var_day = "0" & var_day
	var_year = datepart("yyyy", date_variable)
	var_year = right(var_year, 2)
	date_variable = var_month &"/" & var_day & "/" & var_year
end function

'defining this function here because it needs to not end the script if a MEMO fails.
function start_a_new_spec_memo()
'--- This function navigates user to SPEC/MEMO and starts a new SPEC/MEMO, selecting client, AREP, and SWKR if appropriate
'===== Keywords: MAXIS, notice, navigate, edit
	call navigate_to_MAXIS_screen("SPEC", "MEMO")				'Navigating to SPEC/MEMO

	PF5															'Creates a new MEMO. If it's unable the script will stop.
	EMReadScreen memo_display_check, 12, 2, 33
	If memo_display_check = "Memo Display" then memo_started = False

	'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
	    arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
	    call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
	    EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
	    PF5                                                     'PF5s again to initiate the new memo process
	END IF
	'Checking for SWKR
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
	    swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
	    call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
	    EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
	    PF5                                           'PF5s again to initiate the new memo process
	END IF
	EMWriteScreen "x", 5, 12                                        'Initiates new memo to client
	IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	transmit                                                        'Transmits to start the memo writing process
end function

Function check_pnd2_for_denial(coded_denial, SNAP_pnd2_code, cash_pnd2_code, emer_pnd2_code)
  Call navigate_to_MAXIS_screen("REPT", "PND2")
  row = 7
  col = 5
  EMSearch MAXIS_case_number, row, col      'finding correct case to check PND2 codes

  IF SNAP_check = checked Then
  	EMReadScreen SNAP_pnd2_code, 1, row, 62
  	IF SNAP_pnd2_code = "R" THEN coded_denial = coded_denial & " SNAP withdrawn on PND2."
  	IF SNAP_pnd2_code = "I" THEN coded_denial = coded_denial & " SNAP application incomplete, denied on PND2."
  	IF SNAP_pnd2_code = "_" THEN
  		'If SNAP is selected by the user but the SNAP column is empty on PND2, the script is going to look on the next row for ADDITIONAL APP...
  		EMReadScreen additional_maxis_application, 20, row + 1, 16
  		additional_maxis_application = trim(additional_maxis_application)
  		IF InStr(additional_maxis_application, "ADDITIONAL") <> 0 THEN
  			EMReadScreen SNAP_pnd2_code, 1, row + 1, 62
  			IF SNAP_pnd2_code = "R" THEN coded_denial = coded_denial & " SNAP withdrawn on PND2."
  			IF SNAP_pnd2_code = "I" THEN coded_denial = coded_denial & " SNAP application incomplete, denied on PND2."
  		END IF
  	END IF
  END IF
  IF cash_check = checked Then
  	EMReadScreen cash_pnd2_code, 1, row, 54
  	IF cash_pnd2_code = "R" THEN coded_denial = coded_denial & " CASH withdrawn on PND2."
  	IF cash_pnd2_code = "I" THEN coded_denial = coded_denial & " CASH application incomplete, denied on PND2."
  	IF cash_pnd2_code = "_" THEN
  		'If CASH is selected by the user but the CASH column is empty on PND2, the script is going to look on the next row for ADDITIONAL APP...
  		EMReadScreen additional_maxis_application, 20, row + 1, 16
  		additional_maxis_application = trim(additional_maxis_application)
  		IF InStr(additional_maxis_application, "ADDITIONAL") <> 0 THEN
  			EMReadScreen cash_pnd2_code, 1, row + 1, 54
  			IF cash_pnd2_code = "R" THEN coded_denial = coded_denial & " CASH withdrawn on PND2."
  			IF cash_pnd2_code = "I" THEN coded_denial = coded_denial & " CASH application incomplete, denied on PND2."
  		END IF
  	END IF
  END IF
End function


'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone
'Grabbing the worker's X number.
CALL find_variable("User: ", worker_number, 7)

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'if user is not Hennepin County - the script will end. Process is not approved for other counties
'------------------------------------------------------------------------------------------------------establishing date variables
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year = CM_plus_1_yr

CM_minus_1_mo = right("0" & DatePart("m", DateAdd("m", -1, date)), 2)
CM_minus_1_yr = right(DatePart("yyyy", DateAdd("m", -1, date)), 2)

current_date = date
Call ONLY_create_MAXIS_friendly_date(current_date)			'reformatting the dates to be MM/DD/YY format to measure against the panel dates

'dialog and dialog DO...Loop
Do
	Do
		'The dialog is defined in the loop as it can change as buttons are pressed
		BeginDialog file_select_dialog, 0, 0, 221, 50, "Select the source file"
  			ButtonGroup ButtonPressed
    		PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
    		OkButton 110, 30, 50, 15
    		CancelButton 165, 30, 50, 15
  			EditBox 5, 10, 165, 15, file_selection_path
		EndDialog
		err_msg = ""
		Dialog file_select_dialog
		If ButtonPressed = cancel then stopscript
		If ButtonPressed = select_a_file_button then
			If file_selection_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
				objExcel.Quit 'Closing the Excel file that was opened on the first push'
				objExcel = "" 	'Blanks out the previous file path'
			End If
			call file_selection_system_dialog(file_selection_path, ".xlsx") 'allows the user to select the file'
		End If
		If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
		If IsNumeric(MAXIS_footer_month) = False or len(MAXIS_footer_month) > 2 or len(MAXIS_footer_month) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer month."
  		If IsNumeric(MAXIS_footer_year) = False or len(MAXIS_footer_year) > 2 or len(MAXIS_footer_year) < 2 then err_msg = err_msg & vbNewLine & "* Enter a valid footer year."
		If err_msg <> "" Then MsgBox err_msg
	Loop until err_msg = ""
	If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
	If err_msg <> "" Then MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Sets constants for the array to make the script easier to read (and easier to code)'

'Each of the case numbers will be stored at this position'
const case_number           = 0
const excel_row             = 1
const client_name						= 2
const program_group_ID			= 3
const worker_ID		   				= 4
const program_status				= 5
const priv_case             = 6
const out_of_co             = 7
const written_lang          = 8
const SNAP_case             = 9
const CASH_case             = 10
const application_date      = 11
const interview_date    		= 12
const appt_notc_sent        = 13 'dates'
const nomi_sent             = 14 'dates'
const notc_confirm          = 15
const deny_day30						= 16
const need_appt_notc        = 17
const need_nomi             = 18
const appointment_date			= 19
const error_notes 					= 20

'Sets up the array to store all the information for each client'
Dim application_array()
ReDim application_array (error_notes, 0)

'Now the script adds all the clients on the excel list into an array
row = 2're-establishing the row to start checking the members for
case_entry = 0
'reading each line of the Excel file and adding case number information to the array
'this string will take care of the duplicate maxis case number, leaving only one entry in the array - the excel sheet should be sorted by case number'
all_casenumber_string = "*"
Do
	anything_number = trim(objExcel.Cells(row, 2).value)

	If instr(all_casenumber_string, "*" & anything_number & "*") = 0 then
        'MsgBox anything_number
		all_casenumber_string = all_casenumber_string & anything_number & "*"
		ReDim Preserve application_array(error_notes, case_entry)
		application_array(worker_ID, case_entry) = trim(objExcel.Cells(row, 1).value)
		application_array(case_number, case_entry) = trim(objExcel.Cells(row, 2).value)
		application_array(excel_row, case_entry) = row
		application_array(client_name, case_entry) = trim(objExcel.cells(row, 3).value) 'storing all of the excel information
		application_array(program_group_ID, case_entry) = trim(objExcel.cells(row, 4).value)
		application_array(program_status, case_entry) = trim(objExcel.cells(row, 5).value)
		application_array(application_date, case_entry) = trim(objExcel.cells(row, 6).value)
		application_array(interview_date, case_entry) = trim(objExcel.cells(row, 7).value)
		application_array(CASH_case, case_entry) = FALSE
		application_array(SNAP_case, case_entry) = FALSE
		application_array(need_appt_notc, case_entry) = TRUE
		application_array(need_nomi, case_entry) = TRUE
		case_entry = case_entry + 1
	End If
  row = row + 1
  next_case_number = trim(objExcel.Cells(row, 1).Value)
	stats_counter = stats_counter + 1
loop until next_case_number = ""

total_cases = case_entry

back_to_self
EMWriteScreen MAXIS_footer_month, 20, 43		'Writes in Current month plus one
EMWriteScreen MAXIS_footer_year, 20, 46		'Writes in Current month plus one's year

For case_entry = 0 to Ubound(application_array, 2) 'grabbing additional information form the case
	'Establishing values for each case in the array of cases
	MAXIS_case_number	= application_array(case_number, case_entry)
	'MsgBox application_array(case_number, case_entry)
	CALL navigate_to_MAXIS_screen("CASE", "CURR")
	'Checking for PRIV cases.
	EMReadScreen priv_check, 6, 24, 14 'If it can't get into the case needs to skip
	EMReadScreen county_check, 2, 21, 16    'Looking to see if case has Hennepin COunty worker
	IF priv_check = "PRIVIL" THEN
	  priv_case_list = priv_case_list & "|" & MAXIS_case_number
	 	application_array(priv_case, case_entry) = TRUE
	ELSEIF county_check <> "27" THEN
	 	application_array(out_of_co, case_entry) = "OUT OF COUNTY - " & county_check
	ELSE
	  application_array(priv_case, case_entry) = FALSE
	  'MEMB for written language
	  Call navigate_to_MAXIS_screen("STAT", "MEMB")
	  EMReadScreen language_code, 2, 13, 42
	  application_array(written_lang, case_entry) = language_code

	  'PROG to determine programs active
	  Call navigate_to_MAXIS_screen("STAT", "PROG")
	  EMReadScreen cash_prog_one, 2, 6, 67               'reading for active MFIP program - which has different requirements
	  EMReadScreen cash_stat_one, 4, 6, 74
	  EMReadScreen cash_prog_two, 2, 7, 67
	  EMReadScreen cash_stat_two, 4, 7, 74
		EMReadScreen cash_one_pend, 4, 6, 74
		EMReadScreen cash_two_pend, 4, 7, 74
		EMReadScreen fs_pend, 4, 10, 74
		EMReadScreen team_number, 7, 21, 13
		'handling for a case with nothing pending'
		IF fs_pend <> "PEND" and cash_one_pend <> "PEND" and cash_two_pend <> "PEND" Then
			application_array(need_appt_notc, case_entry) = FALSE
			application_array(need_nomi, case_entry) = FALSE
			application_array(notc_confirm, case_entry) = ", Not needed"
			application_array(error_notes, case_entry) = ", Acted on"
		END IF

		IF fs_pend = "PEND" THEN
			application_array(SNAP_case, case_entry) = TRUE
			EMReadScreen interview_date_done, 8, 10, 55
			interview_date_done = replace(interview_date_done, " ", "/")
			IF interview_date_done <> "__/__/__" THEN
				application_array(need_appt_notc, case_entry) = FALSE
				application_array(need_nomi, case_entry) = FALSE
				application_array(interview_date, case_entry) = interview_date_done
                application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", SNAP Interview entered on PROG"
			END IF
		END IF
		IF cash_one_pend = "PEND" THEN
			application_array(CASH_case, case_entry) = TRUE
			EMReadScreen interview_date_done, 8, 6, 55
			interview_date_done = replace(interview_date_done, " ", "/")
			IF interview_date_done <> "__/__/__" THEN
				application_array(need_appt_notc, case_entry) = FALSE
				application_array(need_nomi, case_entry) = FALSE
				application_array(interview_date, case_entry) = interview_date_done
        application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", CASH Interview entered on PROG"
			END IF
			IF FS_pend <> "PEND" THEN
				IF FS_pend <> "ACTV" THEN application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Cash requested SNAP is not"
			END IF
		END IF
		IF cash_two_pend = "PEND" THEN
			application_array(CASH_case, case_entry) = TRUE
			EMReadScreen interview_date_done, 8, 7, 55
			interview_date_done = replace(interview_date_done, " ", "/")
			IF interview_date_done <> "__/__/__" THEN
				application_array(need_appt_notc, case_entry) = FALSE
				application_array(need_nomi, case_entry) = FALSE
				application_array(interview_date, case_entry) = interview_date_done
                application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", CASH Interview entered on PROG."
			END IF
			IF FS_pend <> "PEND" THEN
				IF FS_pend <> "ACTV" THEN application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Cash requested SNAP is not"
			END IF
		END IF

		IF application_array(need_appt_notc, case_entry) = TRUE or application_array(need_nomi, case_entry) = TRUE Then
			'need to look in case note for ("~ Appointment letter sent in MEMO ~") need
			Call navigate_to_MAXIS_screen("CASE", "NOTE")
			note_row = 5
			day_before_app = DateAdd("d", -1, application_array(application_date, case_entry)) 'will set the date one day prior to app date'
			Do
				EMReadScreen note_date, 8, note_row, 6
				EMReadScreen note_title, 55, note_row, 25
				note_title = trim(note_title)
				'MsgBox note_title TODO Cash requested SNAP is active
				IF note_title = "~ Client missed application interview, NOMI sent via sc" then application_array(nomi_sent, case_entry) = note_date
				IF left(note_title, 35) = "~ Appointment letter sent in MEMO ~" then
                    application_array(appt_notc_sent, case_entry) = note_date
				ElseIF left(note_title, 42) = "~ Appointment letter sent in MEMO for SNAP" then
						        application_array(appt_notc_sent, case_entry) = note_date
				ElseIF left(note_title, 37) = "~ Appointment letter sent in MEMO for" then
					EMReadScreen appt_date, 10, note_row, 63
					appt_date = replace(appt_date, "~", "")
				 	appt_date = trim(appt_date)
					application_array(appointment_date, case_entry) = appt_date
          application_array(appt_notc_sent, case_entry) = note_date
                    'MsgBox application_array(appointment_date, case_entry)
				END IF
				IF left(note_title, 32) = "**Client missed SNAP interview**" then application_array(nomi_sent, case_entry) = note_date
				IF left(note_title, 32) = "**Client missed CASH interview**" then application_array(nomi_sent, case_entry) = note_date
				IF left(note_title, 37) = "**Client missed SNAP/CASH interview**" then application_array(nomi_sent, case_entry) = note_date
				IF note_title = "~ Client has not completed application interview, NOMI" then application_array(nomi_sent, case_entry) = note_date
				'other notes that may not have processed correctly' ***Add program: incomplete***
				IF left(note_title, 15) = "***Add program:" then
					application_array(interview_date, case_entry) = note_date
					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Interview indicated in case note - see interview date, add a program case note"
				END IF
				IF left(note_title, 33) = "***Intake Interview Completed ***" then
					application_array(interview_date, case_entry) = note_date
					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Interview indicated in case note - see interview date, Interview completed"
				END IF
				IF left(note_title, 55) = "~ Client has not completed CASH application interview ~" then
					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", CASH only interview needed"
				END IF
				IF left(note_title, 50) = "~ Client has not completed application interview ~" then
					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Review case"
				END IF
				 IF left(note_title, 40) = "***Reapplication Interview Completed ***" then
 					application_array(interview_date, case_entry) = note_date
 					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Interview indicated in case note - see interview date, Interview completed"
 				END IF
				IF left(note_title, 42) = "~ Interview Completed for SNAP ~" then
					application_array(interview_date, case_entry) = note_date
					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Interview indicated in case note - see interview date, Interview completed"
				END IF
				IF left(note_title, 42) = "*client interviewed* onboarding processing" then
					application_array(interview_date, case_entry) = note_date
					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Interview indicated in case note - see interview date, Interview completed"
				END IF
				IF left(note_title, 34) = "***Intake: pending mentor review**" then
					application_array(interview_date, case_entry) = note_date
					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Interview indicated in case note - see interview date"
				END IF
				IF left(note_title, 23) = "~ Interview Completed ~" then
					application_array(interview_date, case_entry) = note_date
					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Interview indicated in case note - see interview date, Interview completed"
				END IF
				IF left(note_title, 10) = "***Intake:" then
					application_array(interview_date, case_entry) = note_date
					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Interview indicated in case note - see interview date"
				END IF
				IF left(note_title, 24) = "~ Application interview ~" then
					application_array(interview_date, case_entry) = note_date
					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Review case interview may not be needed"
				END IF
				IF left(note_title, 33) = "***Intake Interview Completed ***" then
					application_array(interview_date, case_entry) = note_date
					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Interview indicated in case note - see interview date, Interview completed"
				END IF
				IF left(UCase(note_title), 51) = "Phone call from client re: Phone interview Complete" then
            application_array(interview_date, case_entry) = note_date
            application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Review case phone interview indicated"
        END IF
				IF left(UCase(note_title), 41) = "Phone call from client re: SNAP interview" then
						application_array(interview_date, case_entry) = note_date
						application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Review case phone interview indicated"
				END IF
				IF left(UCase(note_title), 19) = "----DENIED SNAP----" then
					application_array(deny_day30, case_entry) = "PREVIOUS"
					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", DENY SNAP case note"
				END IF
        IF left(UCase(note_title), 19) = "----DENIED CASH----" then
            application_array(deny_day30, case_entry) = "PREVIOUS"
            application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", DENY CASH case note"
        END IF
        IF left(UCase(note_title), 24) = "----DENIED SNAP/CASH----" then
            application_array(deny_day30, case_entry) = "PREVIOUS"
            application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", DENY SNAP/CASH"
        END IF
				IF left(note_title, 31) = "~ Denied CASH/SNAP via script ~" then
            application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", SCRIPT DENIAL ALREADY NOTED"
        END IF
        IF left(note_title, 31) = "~ Denied CASH via script ~" then
            application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", SCRIPT DENIAL ALREADY NOTED"
        END IF
        IF left(note_title, 26) = "~ Denied SNAP via script ~" then
            application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", SCRIPT DENIAL ALREADY NOTED"
        END IF
        IF left(note_title, 20) = "**Courtesy Interview" then
			 		application_array(appt_notc_sent, case_entry) = note_date
					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Interview completed out of county"
				END IF
				IF left(note_title, 18) = "**New CAF received" then
			 		application_array(appt_notc_sent, case_entry) = note_date
					application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Other appt notice used"
				END IF
					IF note_date = "        " then Exit Do
					note_row = note_row + 1
					IF note_row = 19 THEN
						PF8
						note_row = 5
					END IF
					EMReadScreen next_note_date, 8, note_row, 6
					IF next_note_date = "        " then Exit Do
			Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'
			IF application_array(worker_ID, case_entry) = "X127EF8" or application_array(worker_ID, case_entry) = "X127EJ1" THEN
				application_array(need_nomi, case_entry) = FALSE
				application_array(need_appt_notc, case_entry) = FALSE
				application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", IMD case load review case"
			END IF
			IF application_array(interview_date, case_entry) <> "" Then
				application_array(need_nomi, case_entry) = FALSE
				application_array(need_appt_notc, case_entry) = FALSE
			END IF
			IF application_array(appt_notc_sent, case_entry) = "" Then
				application_array(need_nomi, case_entry) = FALSE
			Else
				application_array(need_appt_notc, case_entry) = FALSE
				'MsgBox datediff("d", application_array(appt_notc_sent, case_entry), date)
				IF application_array(appointment_date, case_entry) = "" Then application_array(appointment_date, case_entry) = dateadd("d", 7, application_array(appt_notc_sent, case_entry)) 'this will add seven days to the date the appt letter was sent'
        'MsgBox application_array(appointment_date, case_entry)
				'IF application_array(appointment_date, case_entry) <= Date THEN
        If datediff("d", application_array(appointment_date, case_entry), date) >= 0 Then
          ' MsgBox "Case needs NOMI - appt date: " & application_array(appointment_date, case_entry)
				IF application_array(nomi_sent, case_entry) <> "" THEN application_array(need_nomi, case_entry) = FALSE
				Else
					application_array(need_nomi, case_entry) = FALSE
				END IF
			END IF
		END IF
		'MsgBox datediff("d", application_array(application_date, case_entry), date)
        'TODO add handling for if SNAP interview has been completed but NOT face to face and Cash needs a denial
		IF datediff("d", application_array(application_date, case_entry), date) >= 30 and application_array(interview_date, case_entry) = "" and application_array(notc_confirm, case_entry) <> "Not needed" THEN
		    IF application_array(need_appt_notc, case_entry) = FALSE and application_array(need_nomi, case_entry) = FALSE THEN
			'MsgBox "Both false notice"
			'MsgBox application_array(nomi_sent, case_entry)
            IF application_array(nomi_sent, case_entry) <> "" then
              last_contact_day = dateadd("d", 30, application_array(application_date, case_entry))
    				IF datediff("d", application_array(nomi_sent, case_entry), date) >= 10 or datediff("d", application_array(nomi_sent, case_entry), last_contact_day) > 0 THEN
    				'MsgBox datediff("d", application_array(nomi_sent, case_entry), date)
    					Call navigate_to_MAXIS_screen("REPT", "PND2")
    					Row = 1
    					Col = 1
    					EMSearch MAXIS_case_number, row, col
    					EMReadScreen nbr_days_pending, 3, row, 50
    		  		    nbr_days_pending = trim(nbr_days_pending)
    					nbr_days_pending = nbr_days_pending * 1
    					IF nbr_days_pending >= 30 THEN application_array(deny_day30, case_entry) = TRUE
    					'msgbox nbr_days_pending
    				END IF
    			END IF
            END IF
		END IF
	END IF
	'MsgBox "Need to send Appt Notice - " & application_array(need_appt_notc, case_entry) & vbNewLine & "Need to send NOMI - " & application_array(need_nomi, case_entry) & vbNewLine & vbNewLine & "Interview done on - " & application_array(interview_date, case_entry)
	'If application_array(deny_day30, case_entry) = TRUE THEN MsgBox "this case may be a denial"
CALL back_to_SELF
Stats_counter = stats_counter + 1
NEXT
'adding my headers to excel
objExcel.cells(1, 8).value 	  = "Notice Sent"
objExcel.cells(1, 9).value 		= "Nomi Sent"
objExcel.cells(1, 10).value 	= "Confirmation"
objExcel.cells(1, 11).value 	= "Denial"
objExcel.Cells(1, 12).value 	= "Case information"
objExcel.cells(1, 13).value 	= "Privileged"
objExcel.cells(1, 14).value 	= "Out of County"
'objExcel.Cells(1, 16).value 	= "other"
'doing all the things and reporting it to excel'
FOR case_entry = 0 to Ubound(application_array, 2)
    MAXIS_case_number = application_array(case_number, case_entry)        'setting this for using navigate functions
    forms_to_swkr = ""
    forms_to_arep = ""
    memo_started = TRUE
    if application_array(priv_case, case_entry) = FALSE and application_array(out_of_co, case_entry) = "" then                  'PRIV cases will not have a MEMO attempted
      if application_array(CASH_case, case_entry) = TRUE then           'setting the language for the notices - MFIP or SNAP
      	if application_array(SNAP_case, case_entry) = TRUE then
        	programs = "CASH/SNAP"
      	else
        	programs = "CASH"
      	end if
    	else
        programs = "SNAP"
    	end if
      'The function will return this as FALSE if PF5 does not move past MEMO DISPLAY
			'MsgBox dateadd("d", 7, application_array(application_date, case_entry))
			'MsgBox dateadd("d", 30, application_array(application_date, case_entry))
			'need to check and make sure this is not 10 days !!'
			missed_interview_date  = dateadd("d", 7, application_array(application_date, case_entry))
			IF missed_interview_date < application_array(appt_notc_sent, case_entry) THEN  missed_interview_date = dateadd("d", 7, application_array(appt_notc_sent, case_entry))

			need_intv_date = dateadd("d", 7, application_array(application_date, case_entry))    'NOTE - had to change this - it did not call the full array - dates were wrong.
			If need_intv_date <= date then need_intv_date = dateadd("d", 7, date)
			need_intv_date = need_intv_date & ""		'turns interview date into string for variable
		  last_contact_day = dateadd("d", 30, application_array(application_date, case_entry))
			nomi_last_contact_day = dateadd("d", 30, application_array(application_date, case_entry))
			'ensuring that we have given the client an additional10days fromt he day nomi sent'
			If DateDiff("d", need_intv_date, last_contact_day) < 1 then last_contact_day = need_intv_date
			IF DateDiff("d", application_array(nomi_sent, case_entry), nomi_last_contact_day) < 1 then nomi_last_contact_day = dateadd("d", 10, application_array(nomi_sent, case_entry))

			'msgbox "last contact " & last_contact_day & vbNewLine & "nomi last contact " & nomi_last_contact_day

		IF application_array(need_appt_notc, case_entry) = TRUE THEN
		  start_a_new_spec_memo_and_continue(memo_started)		'Writes the appt letter into the MEMO.
			IF memo_started = True THEN
				EMsendkey("************************************************************")
		    Call write_variable_in_SPEC_MEMO("You recently applied for assistance in Hennepin County on " & application_array(application_date, case_entry) & ".")
		    Call write_variable_in_SPEC_MEMO("An interview is required to process your application.")
				Call write_variable_in_SPEC_MEMO(" ")
		    Call write_variable_in_SPEC_MEMO("The interview must be completed by " & need_intv_date & ".")
		    Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at")
				Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday through Friday.")
				Call write_variable_in_SPEC_MEMO(" ")
				Call write_variable_in_SPEC_MEMO("If we do not hear from you by " & last_contact_day & " your application will be denied.") 'add 30 days
				Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to-face interview.")
		    Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
				Call write_variable_in_SPEC_MEMO("You can also request a paper copy.")
				Call write_variable_in_SPEC_MEMO("Auth: 7CFR 273.2(e)(3). ")
		    Call write_variable_in_SPEC_MEMO("************************************************************")
				application_array(appt_notc_sent, case_entry) = date
				PF4
			ELSE
				application_array(notc_confirm, case_entry) = "N" 'Setting this as N if the MEMO failed
				application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Memo failed"
				call back_to_SELF
			END IF
    ELSEIF application_array(need_nomi, case_entry) = TRUE THEN
			start_a_new_spec_memo_and_continue(memo_started)		'Writes the NOMI into the MEMO.
			IF memo_started = TRUE THEN
			  EMsendkey("************************************************************")
			  Call write_variable_in_SPEC_MEMO("You recently applied for assistance on " & application_array(application_date, case_entry) & ".")
				Call write_variable_in_SPEC_MEMO("Your interview should have been completed by " & application_array(appointment_date, case_entry) & ".")
				Call write_variable_in_SPEC_MEMO("An interview is required to process your application.")
			  Call write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at ")
				Call write_variable_in_SPEC_MEMO("612-596-1300 between 9:00am and 4:00pm Monday through Friday.")
				Call write_variable_in_SPEC_MEMO(" ")
				Call write_variable_in_SPEC_MEMO("If you do not complete the interview by " & nomi_last_contact_day & " your application will be denied.") 'add 30 days
				Call write_variable_in_SPEC_MEMO(" ")
				Call write_variable_in_SPEC_MEMO("If you are applying for a cash program for pregnant women or minor children, you may need a face-to- face interview.")
			  Call write_variable_in_SPEC_MEMO("Domestic violence brochures are available at https://edocs.dhs.state.mn.us/lfserver/Public/DHS-3477-ENG.")
				Call write_variable_in_SPEC_MEMO("You can also request a paper copy.")
				Call write_variable_in_SPEC_MEMO("Auth: 7CFR 273.2(e)(3). ")
				Call write_variable_in_SPEC_MEMO("************************************************************")
				application_array(nomi_sent, case_entry) = date
				PF4
			ELSE
				application_array(notc_confirm, case_entry) = "N"         'Setting this as N if the MEMO failed
				application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Memo failed"
				call back_to_SELF
			END IF
	 	ELSEIF application_array(deny_day30, case_entry) = TRUE THEN
			'MsgBox application_array(deny_day30, case_entry)
			start_a_new_spec_memo_and_continue(memo_started)		'Writes the denial into the MEMO.
			IF memo_started = True THEN
				EMsendkey("************************************************************")
				Call write_variable_in_SPEC_MEMO("We received your application on " & application_array(application_date, case_entry) & ".")
				Call write_variable_in_SPEC_MEMO("Your interview was not completed by " & nomi_last_contact_day & ".")
				call write_variable_in_spec_memo("Due to failing to complete the interview within 30 days of your application date your case has been denied.")
				Call write_variable_in_SPEC_MEMO("************************************************************")
				PF4
			ELSE
				application_array(notc_confirm, case_entry) = "N"         'Setting this as N if the MEMO failed
				application_array(error_notes, case_entry) = application_array(error_notes, case_entry) & ", Memo failed"
				call back_to_SELF
				'MsgBox "What memo was sent?"
			END IF
		ELSE
			application_array(notc_confirm, case_entry) = "Not needed"
		END IF
'checking to ensure memo was sent'
		IF application_array(notc_confirm, case_entry) <> "N" and application_array(notc_confirm, case_entry) <> ", Not needed"  Then
		memo_row = 7
		'creating a variable in the MM/DD/YY format to compare with date read from MAXIS
		today_mo = DatePart("m", date)
		today_mo = right("00" & today_mo, 2)

		today_day = DatePart("d", date)
		today_day = right("00" & today_day, 2)

		today_yr = DatePart("yyyy", date)
		today_yr = right(today_yr, 2)

		today_date = today_mo & "/" & today_day & "/" & today_yr
			Do
				EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
				EMReadScreen print_status, 7, memo_row, 67
				'MsgBox print_status
				If create_date = today_date AND print_status = "Waiting" Then   'MEMOs created today and still waiting is likely our MEMO.
					application_array(notc_confirm, case_entry) = "Y"             'If we've found this then no reason to keep looking.
					successful_notices = successful_notices + 1
					'MsgBox application_array(notc_confirm, case_entry)                 'For statistical purposes
					Exit Do
				End If
					memo_row = memo_row + 1           'Looking at next row'
			Loop Until create_date = "        "
			IF application_array(notc_confirm, case_entry) = "Y" and application_array(deny_day30, case_entry) = TRUE THEN
				write_casenote_confirm = MsgBox ("Do want to case note?", vbYesNo + vbQuestion, "Confirm case note")
				IF write_casenote_confirm = vbNo Then	application_array(notc_confirm, case_entry) = "N"
			END IF
		END IF
		IF application_array(notc_confirm, case_entry) = "Y" THEN
			Call start_a_blank_case_note
				IF application_array(need_appt_notc, case_entry) = TRUE THEN
			 		Call write_variable_in_CASE_NOTE("~ Appointment letter sent in MEMO for " & need_intv_date & "~")
					Call write_variable_in_CASE_NOTE("* A notice has been sent via SPEC/MEMO informing the client of needed interview.")
				END IF
				IF application_array(need_nomi, case_entry) = TRUE THEN
					Call write_variable_in_CASE_NOTE("~ Client has not completed application interview, NOMI sent via script ~ ")
					Call write_variable_in_CASE_NOTE("* A notice was previously sent to client with detail about completing an interview. ")
				END IF
				IF application_array(need_appt_notc, case_entry) = TRUE or application_array(need_nomi, case_entry) = TRUE THEN
					Call write_variable_in_CASE_NOTE("* Households failing to complete the interview within 30 days of the date they file an application will receive a denial notice")
					Call write_variable_in_CASE_NOTE("* A link to the domestic violence brochure sent to client in SPEC/MEMO as a part of interview notice.")
					Call write_variable_in_CASE_NOTE("---")
					Call write_variable_in_CASE_NOTE(worker_signature & " via bulk on demand waiver script")
					'MsgBox "What casenote was sent?"
					PF3
				END IF
				IF application_array(deny_day30, case_entry) = TRUE THEN
					Call write_variable_in_case_note("~ Denied " & programs & " via script ~")
					Call write_bullet_and_variable_in_case_note("Application date", application_array(application_date, case_entry))
					Call write_variable_in_case_note("* Reason for denial: interview was not completed timely.")
					Call write_variable_in_case_note("* Confirmed client was provided sufficient 10 day notice.")
					Call write_bullet_and_variable_in_case_note("NOMI sent to client on ", application_array(nomi_sent, case_entry))
					Call write_variable_in_case_note("---")
					Call write_variable_in_CASE_NOTE(worker_signature & " via bulk on demand waiver script")
					'MsgBox "What casenote was sent?"
					PF3
				END IF
			END IF
		END IF
'finally filling in the information'
	row = application_array(excel_row, case_entry)
    'MsgBox row
  objExcel.Cells(row, 7).Value = application_array(interview_date, case_entry)
  objExcel.Cells(row, 8).Value = application_array(appt_notc_sent, case_entry) '= "Appt Notice Success"
	objExcel.Cells(row, 9).Value = application_array(nomi_sent, case_entry) '= "NOMI Success"
	objExcel.Cells(row, 10).Value = application_array(notc_confirm, case_entry)
	objExcel.Cells(row, 11).Value = application_array(deny_day30, case_entry)'true or blank'
	IF application_array(error_notes, case_entry) <> "" THEN application_array(error_notes, case_entry) = right(application_array(error_notes, case_entry), len(application_array(error_notes, case_entry))- 2)
	objExcel.Cells(row, 12).Value = application_array(error_notes, case_entry)
  IF application_array(priv_case, case_entry) = TRUE THEN objExcel.Cells(row, 13).Value = "PRIV"
  objExcel.Cells(row, 14).Value = application_array(out_of_co, case_entry)
	IF application_array(worker_ID, case_entry) = "X127EF8" or application_array(worker_ID, case_entry) = "X127EJ1" THEN objExcel.Rows(row).font.colorindex = 10
	IF application_array(error_notes, case_entry) = "Interview indicated in case note - see interview date" THEN objExcel.Rows(row).font.colorindex = 21 'blue'
	IF application_array(error_notes, case_entry) = "Review case" THEN objExcel.Rows(row).font.colorindex = 46 'orange'
	IF application_array(deny_day30, case_entry) = TRUE THEN
		objExcel.Rows(row).font.colorindex = 3 'red'
		objExcel.Rows(row).font.bold = True
	END IF
	IF application_array(interview_date, case_entry) <> "" THEN objExcel.Rows(row).font.colorindex = 5 'orange'
	call back_to_SELF
NEXT
FOR case_entry = 1 to 15									'formatting the cells'
	objExcel.Cells(1, case_entry).Font.Bold = True		'bold font'
	objExcel.Columns(case_entry).AutoFit()				'sizing the columns'
NEXT'

entry_row = 1
stats_header_col = 17
stats_col = 18
thirty_days_ago = DateAdd("d", -30, date)

objExcel.Cells(entry_row, stats_header_col).Value       = "Appointment Notices run on:"     'Date and time the script was completed
objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
objExcel.Cells(entry_row, stats_col).Value              = now
entry_row = entry_row + 1

objExcel.Cells(entry_row, stats_header_col).Value       = "Runtime (in seconds)"            'Enters the amount of time it took the script to run
objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
objExcel.Cells(entry_row, stats_col).Value              = timer - query_start_time
entry_row = entry_row + 1

objExcel.Cells(entry_row, stats_header_col).Value       = "Total Cases assesed"             'All cases from the spreadsheet
objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
objExcel.Cells(entry_row, stats_col).Value              = total_cases
entry_row = entry_row + 1

objExcel.Cells(entry_row, stats_header_col).Value       = "Cases at 30 DAYS"        'number of notices that were successful
objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIFS(F:F," & Chr(34)  & thirty_days_ago & Chr(34) & ")"                'This was incremented on the For Next loop where the memos were written
entry_row = entry_row + 1

objExcel.Cells(entry_row, stats_header_col).Value       = "Cases OVER 30 DAYS"        'number of notices that were successful
objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIFS(F:F," & Chr(34) & "<" & thirty_days_ago & Chr(34) & ")"                'This was incremented on the For Next loop where the memos were written
entry_row = entry_row + 1

objExcel.Cells(entry_row, stats_header_col).Value       = "Cases at potential Denial"           'calculation of the percent of successful notices
objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIFS(K:K," & Chr(34) & "TRUE" & Chr(34) & ")"
entry_row = entry_row + 1

objExcel.Cells(entry_row, stats_header_col).Value       = "Appointment Notices Sent"           'calculation of the percent of successful notices
objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIFS(H:H," & Chr(34) & date & Chr(34) & ",J:J," & Chr(34) & "Y" & Chr(34) & ")"
entry_row = entry_row + 1

objExcel.Cells(entry_row, stats_header_col).Value       = "NOMIs Sent"           'calculation of the percent of successful notices
objExcel.Cells(entry_row, stats_header_col).Font.Bold 	= TRUE
objExcel.Cells(entry_row, stats_col).Value              = "=COUNTIFS(I:I," & Chr(34) & date & Chr(34) & ",J:J," & Chr(34) & "Y" & Chr(34) & ")"

entry_row = entry_row + 1

for row_to_change = 1 to entry_row
    objExcel.Cells(row_to_change, stats_header_col).font.colorindex = 1
    objExcel.Cells(row_to_change, stats_col).font.colorindex = 1
next

script_end_procedure("Success! The list is complete. Please review the cases that appear to be in error.")
