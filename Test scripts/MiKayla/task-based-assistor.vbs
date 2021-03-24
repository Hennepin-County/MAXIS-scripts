'STATS GATHERING--------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - TASK BASED ASSISTOR.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 100                      'manual run time in seconds
STATS_denomination = "C"       			   'M is for each CASE
'END OF stats block==============================================================================================

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY================================================================
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
CALL changelog_update("01/15/2021", "Initial version.", "MiKayla Handley, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
FUNCTION find_MAXIS_worker_number(x_number)
	EMReadScreen SELF_check, 4, 2, 50		'Does this to check to see if we're on SELF screen
	IF SELF_check = "SELF" THEN				'if on the self screen then x # is read from coordinates
		EMReadScreen x_number, 7, 22, 8
	ELSE
		Call find_variable("PW: ", x_number, 7)	'if not, then the PW: variable is searched to find the worker #
		If isnumeric(MAXIS_worker_number) = true then 	 'making sure that the worker # is a number
			MAXIS_worker_number = x_number				'delcares the MAXIS_worker_number to be the x_number
		End if
	END if
END FUNCTION
'THE SCRIPT-----------------------------------------------------------------------------------------------------------
EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

''----------------------------------------------------------------------------------------------------The current day's assignment
'report_date = replace(date, "/", "-")   'Changing the format of the date to use as file path selection default
previous_date = dateadd("d", -1, date)
Call change_date_to_soonest_working_day(previous_date)       'finds the most recent previous working day for the file names
file_date = replace(previous_date, "/", "-")   'Changing the format of the date to use as file path selection default
'file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & file_date & ".xlsx"

BeginDialog Dialog1, 0, 0, 266, 115, "TASK BASED REVIEW"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used for task based review on a list of pending SNAP and/or MFIP cases."
  Text 15, 70, 230, 15, "Select the Excel file that contains your information by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog

'dialog and dialog DO...Loop
Do
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation
        If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue."
        If err_msg <> "" Then MsgBox err_msg
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Opening today's list
Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file
'objExcel.worksheets("Report 1").Activate                                 'Activates the initial BOBI report

'Establishing array
DIM task_based_array()           'Declaring the array
ReDim task_based_array(18, 0)     'Resizing the array
'Creating constants to value the array elements
const date_assigned_const       = 0 '= "Date Assigned"
const SSR_name_const	        = 1 '= "SSR Name"
const maxis_case_number_const 	= 2 '= "Case Number"
const case_name_const       	= 3 '= "Case Name"
const basket_const  			= 4 '= "Basket"
const assigned_to_const    		= 5 '= "Assigned to"
const worker_number_const       = 6 '= "Assigned Worker X127#"
const do_this_const        		= 7 '= "Does worker log indicate they could work the case?"
const case_logged_const         = 8 '= "Case logged by assigned worker?"
const case_note_date_const      = 9  '= "Case Note Date"
const case_note_match_const     = 10 '= "Worker who made case note same as assigned worker"
const case_note_keyword_const   = 11 '= "Does case note title contain keyword"
const DAIL_count_const        	= 12 '= "DAIL Count"
const DAIL_type_const        	= 13 '= "DAIL Type"
const ECF_type_const            = 14 '= "EWS ECF Item Count"
const ECF_form_const            = 15 '= "ECF Form Types"
const oldest_APPL_date_const    = 16 '= "Oldest ECF APPL Date"
const prev_comments_const       = 17 '= "Comments"
const case_status_const 		= 18 '= "Pending over 30 days"
const interview_const           = 19 '= "Interview Completed"

'Now the script adds all the clients on the excel list into an array
excel_row = 5                   're-establishing the row to start based on when Report 1 starts
entry_record = 0                'incrementor for the array and count
all_case_numbers_array = "*"    'setting up string to find duplicate case numbers
Do
    'Reading information from the BOBI report in Excel
    worker_number = objExcel.cells(excel_row, 6).Value
    worker_number = trim(worker_number)

    MAXIS_case_number = objExcel.cells(excel_row, 2).Value
    MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do

    program_ID = objExcel.cells(excel_row, 17).Value
    program_ID = trim(program_ID)

    application_date = objExcel.cells(excel_row, 15).Value
    interview_date   = objExcel.cells(excel_row, 18).Value

    days_pending = datediff("D", application_date, date)

    'If the case number is found in the string of case numbers, it's not added again.
    If instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") then
        add_to_array = False
    Else
        'Adding client information to the array
        ReDim Preserve task_based_array(18, entry_record)	'This resizes the array based on the number of cases

		task_based_array(date_assigned_const,      entry_record) = ""
		'task_based_array(SSR_name_const, 		   entry_record) = ""
		task_based_array(maxis_case_number_const,  entry_record) = MAXIS_case_number
		task_based_array(case_name_const,          entry_record) = MAXIS_case_name
		task_based_array(basket_const,  		   entry_record) = ""
		task_based_array(assigned_to_const,        entry_record) = ""
		task_based_array(worker_number_const,      entry_record) = worker_number
		task_based_array(do_this_const,            entry_record) = ""
		task_based_array(case_logged_const,    	   entry_record) = ""
		task_based_array(case_note_date_const,     entry_record) = ""
		task_based_array(case_note_match_const,    entry_record) = ""
		task_based_array(case_note_keyword_const,  entry_record) = ""
		task_based_array(DAIL_count_const,         entry_record) = ""
		task_based_array(DAIL_type_const,          entry_record) = ""
		task_based_array(ECF_form_const,   		   entry_record) = ""
		task_based_array(ECF_type_const,           entry_record) = ""
		task_based_array(oldest_APPL_date_const,   entry_record) = trim(application_date)
		task_based_array(prev_comments_const,      entry_record) = program_ID
		task_based_array(case_status_const,        entry_record) = days_pending
		task_based_array(interview_const,          entry_record) = trim(interview_date)
		'making space in the array for these variables, but valuing them as "" for now

        entry_record = entry_record + 1			'This increments to the next entry in the array
        stats_counter = stats_counter + 1       'Increment for stats counter
    End if
    excel_row = excel_row + 1
Loop

back_to_self                            'resetting MAXIS back to self before getting started
Call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

'Loading of cases is complete. Reviewing the cases in the array.
For item = 0 to UBound(task_based_array, 2)
    worker_number       = task_based_array(worker_number_const,    item)     're-valuing array variables
    MAXIS_case_number   = task_based_array(case_number_const,      item)
    program_ID          = task_based_array(program_ID_const,       item)
    days_pending        = task_based_array(days_pending_const,     item)
    application_date    = task_based_array(application_date_const, item)
	MAXIS_case_name     = task_based_array(application_date_const, item)

	'setting the footer month to make the updates in'
	CALL convert_date_into_MAXIS_footer_month(date_received, MAXIS_footer_month, MAXIS_footer_year)
	MAXIS_footer_month_confirmation

	If left(worker_number, 4) <> "X127" then                                    'Out of county cases from initial upload
        task_based_array(case_status_const, item) = "OUT OF COUNTY CASE"
    Else
        Call navigate_to_MAXIS_screen("STAT", "PROG")
        EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip - checking in PROD and INQUIRY
        IF priv_check = "PRIV" then                                             'PRIV cases
            EmReadscreen priv_worker, 26, 24, 46
            task_based_array(case_status_const, item) = trim(priv_worker)
            task_based_array(do_this_const, item) = "Privileged Cases"
        ELSE
            EMReadScreen county_code, 4, 21, 21                                 'Out of county cases from STAT
            If county_code <> "X127" then
                task_based_array(case_status_const, item) = "OUT OF COUNTY CASE"
            End if
		ELSE
			EMReadScreen case_invalid_error, 72, 24, 2 'if a person enters an invalid footer month for the case the script will attempt to navigate'
			task_based_array(case_status_const, item) = trim(case_invalid_error)
			task_based_array(do_this_const, item) = "Error Message"
			PF10
        End if

    End if

	'Reading the app date from PROG need to compare for over 30 days and the interview stuffs
	EMReadScreen cash1_app_date, 8, 6, 33
	cash1_app_date = replace(cash1_app_date, " ", "/")
	EMReadScreen cash2_app_date, 8, 7, 33
	cash2_app_date = replace(cash2_app_date, " ", "/")
	EMReadScreen emer_app_date, 8, 8, 33
	emer_app_date = replace(emer_app_date, " ", "/")
	EMReadScreen grh_app_date, 8, 9, 33
	grh_app_date = replace(grh_app_date, " ", "/")
	EMReadScreen snap_app_date, 8, 10, 33
	snap_app_date = replace(snap_app_date, " ", "/")
	EMReadScreen ive_app_date, 8, 11, 33
	ive_app_date = replace(ive_app_date, " ", "/")
	EMReadScreen hc_app_date, 8, 12, 33
	hc_app_date = replace(hc_app_date, " ", "/")
	EMReadScreen cca_app_date, 8, 14, 33
	cca_app_date = replace(cca_app_date, " ", "/")

	'Reading the program status
	EMReadScreen cash1_status_check, 4, 6, 74
	EMReadScreen cash2_status_check, 4, 7, 74
	EMReadScreen emer_status_check, 4, 8, 74
	EMReadScreen grh_status_check, 4, 9, 74
	EMReadScreen snap_status_check, 4, 10, 74
	EMReadScreen ive_status_check, 4, 11, 74
	EMReadScreen hc_status_check, 4, 12, 74
	EMReadScreen cca_status_check, 4, 14, 74
	'----------------------------------------------------------------------------------------------------ACTIVE program coding
	EMReadScreen cash1_prog_check, 2, 6, 67     'Reading cash 1
	EMReadScreen cash2_prog_check, 2, 7, 67     'Reading cash 2
	EMReadScreen emer_prog_check, 2, 8, 67      'EMER Program

	'Logic to determine if MFIP is active
	IF cash1_prog_check = "MF" or cash1_prog_check = "GA" or cash1_prog_check = "DW" or cash1_prog_check = "MS" THEN
		IF cash1_status_check = "ACTV" THEN cash_active = TRUE
	END IF
	IF cash2_prog_check = "MF" or cash2_prog_check = "GA" or cash2_prog_check = "DW" or cash2_prog_check = "MS" THEN
		IF cash2_status_check = "ACTV" THEN cash2_active = TRUE
	END IF
	IF emer_prog_check = "EG" and emer_status_check = "ACTV" THEN emer_active = TRUE
	IF emer_prog_check = "EA" and emer_status_check = "ACTV" THEN emer_active = TRUE

	IF cash1_status_check = "ACTV" THEN cash_active  = TRUE
	IF cash2_status_check = "ACTV" THEN cash2_active = TRUE
	IF snap_status_check  = "ACTV" THEN SNAP_active  = TRUE
	IF grh_status_check   = "ACTV" THEN grh_active   = TRUE
	IF ive_status_check   = "ACTV" THEN IVE_active   = TRUE
	IF hc_status_check    = "ACTV" THEN hc_active    = TRUE
	IF cca_status_check   = "ACTV" THEN cca_active   = TRUE

	active_programs = ""        'Creates a variable that lists all the active.
	IF cash_active = TRUE or cash2_active = TRUE THEN active_programs = active_programs & "CASH, "
	IF emer_active = TRUE THEN active_programs = active_programs & "Emergency, "
	IF grh_active  = TRUE THEN active_programs = active_programs & "GRH, "
	IF snap_active = TRUE THEN active_programs = active_programs & "SNAP, "
	IF ive_active  = TRUE THEN active_programs = active_programs & "IV-E, "
	IF hc_active   = TRUE THEN active_programs = active_programs & "HC, "
	IF cca_active  = TRUE THEN active_programs = active_programs & "CCA"

	active_programs = trim(active_programs)  'trims excess spaces of active_programs
	If right(active_programs, 1) = "," THEN active_programs = left(active_programs, len(active_programs) - 1)

	'----------------------------------------------------------------------------------------------------Pending programs
	programs_applied_for = ""   'Creates a variable that lists all pending cases.
	additional_programs_applied_for = ""
	'cash I
	IF cash1_status_check = "PEND" then
	    If cash1_app_date = application_date THEN
	        cash_pends = TRUE
	        programs_applied_for = programs_applied_for & "CASH, "
	    Else
	        additional_programs_applied_for = additional_programs_applied_for & "CASH, "
	    End if
	End if
	'cash II
	IF cash2_status_check = "PEND" then
	    if cash2_app_date = application_date THEN
	        cash2_pends = TRUE
	        programs_applied_for = programs_applied_for & "CASH, "
	    Else
	        additional_programs_applied_for = additional_programs_applied_for & "CASH, "
	    End if
	End if
	'SNAP
	IF snap_status_check  = "PEND" then
	    If snap_app_date  = application_date THEN
	        SNAP_pends = TRUE
	        programs_applied_for = programs_applied_for & "SNAP, "
	    else
	        additional_programs_applied_for = additional_programs_applied_for & "SNAP, "
	    end if
	End if
	'GRH
	IF grh_status_check = "PEND" then
	    If grh_app_date = application_date THEN
	        grh_pends = TRUE
	        programs_applied_for = programs_applied_for & "GRH, "
	    else
	        additional_programs_applied_for = additional_programs_applied_for & "GRH, "
	    End if
	End if
	'I-VE
	IF ive_status_check = "PEND" then
	    if ive_app_date = application_date THEN
	        IVE_pends = TRUE
	        programs_applied_for = programs_applied_for & "IV-E, "
	    else
	        additional_programs_applied_for = additional_programs_applied_for & "IV-E, "
	    End if
	End if
	'HC
	IF hc_status_check = "PEND" then
	    If hc_app_date = application_date THEN
	        hc_pends = TRUE
	        programs_applied_for = programs_applied_for & "HC, "
	    else
	        additional_programs_applied_for = additional_programs_applied_for & "HC, "
	    End if
	End if
	'CCA
	IF cca_status_check = "PEND" then
	    If cca_app_date = application_date THEN
	        cca_pends = TRUE
	        programs_applied_for = programs_applied_for & "CCA, "
	    else
	        additional_programs_applied_for = additional_programs_applied_for & "CCA, "
	    End if
	End if
	'EMER
	If emer_status_check = "PEND" then
	    If emer_app_date = application_date then
	        emer_pends = TRUE
	        IF emer_prog_check = "EG" THEN programs_applied_for = programs_applied_for & "EGA, "
	        IF emer_prog_check = "EA" THEN programs_applied_for = programs_applied_for & "EA, "
	    else
	        IF emer_prog_check = "EG" THEN additional_programs_applied_for = additional_programs_applied_for & "EGA, "
	        IF emer_prog_check = "EA" THEN additional_programs_applied_for = additional_programs_applied_for & "EA, "
	    End if
	End if

	programs_applied_for = trim(programs_applied_for)       'trims excess spaces of programs_applied_for
	If right(programs_applied_for, 1) = "," THEN programs_applied_for = left(programs_applied_for, len(programs_applied_for) - 1)

	additional_programs_applied_for = trim(additional_programs_applied_for)       'trims excess spaces of programs_applied_for
	If right(additional_programs_applied_for, 1) = "," THEN additional_programs_applied_for = left(additional_programs_applied_for, len(additional_programs_applied_for) - 1)

	CALL Navigate_to_MAXIS_screen("STAT", "MEMB")   'navigating to stat memb to gather the ref number and name.
    IF access_denied_check = "ACCESS DENIED" Then
        PF10
        last_name = "UNABLE TO FIND"
        first_name = " - Access Denied"
        mid_initial = ""
    ELSE
        EMReadscreen last_name, 25, 6, 30
        EMReadscreen first_name, 12, 6, 63
        last_name = trim(replace(last_name, "_", ""))
        first_name = trim(replace(first_name, "_", ""))
    	MAXIS_case_name = first_name & " "  & last_name
		task_based_array(MAXIS_case_name_const, item) = MAXIS_case_name
	END IF

        If check_case_note = True then
            Call navigate_to_MAXIS_screen("CASE", "NOTE")
            'starting at the 1st case note, checking the headers for the NOTES - EXPEDITED SCREENING text or the NOTES - EXPEDITED DETERMINATION text
            MAXIS_row = 5
            Do
                EMReadScreen first_case_note_date, 8, 5, 6 'static reading of the case note date to determine if no case notes acutually exist.
                If trim(first_case_note_date) = "" then
                    case_note_found = True
                    task_based_array(case_status_const, item) = "Case Notes Do Not Exist"
                    task_based_array(do_this_const, item) = "Exp Screening Req"
                    screening_count = screening_count + 1
                    exit do
                Else
                    EMReadScreen case_note_date, 8, MAXIS_row, 6    'incremented row - reading the case note date
                    EMReadScreen case_note_header, 55, MAXIS_row, 25
                    case_note_header = lcase(trim(case_note_header))

                    If trim(case_note_date) = "" then
                        case_note_found = False             'The end of the case notes has been found
                        exit do
                    ElseIf instr(case_note_header, "appears expedited") or instr(case_note_header, "appears expedit") then
                        case_note_found = True
                        task_based_array(case_status_const, item) = "Appears Expedited"
                        task_based_array(do_this_const, item) = "Req Exp Processing"
                        task_count = task_count + 1
                        exit do
                    Elseif instr(case_note_header, "does not appear") or instr(case_note_header, "appears not expedited") then
                        case_note_found = True
                        task_based_array(case_status_const, item) = "Screened, Not EXP"
                        task_based_array(do_this_const, item) = "Not Expedited"
                        exit do
                    Else
                        case_note_found = False         'defaulting to false if not able to find an expedited case note
                        MAXIS_row = MAXIS_row + 1
                        IF MAXIS_row = 19 then
                            PF8                         'moving to next case note page if at the end of the page
                            MAXIS_row = 5
                        End if
                    END IF
                END IF
            LOOP until cdate(case_note_date) < cdate(application_date)                        'repeats until the case note date is less than the application date
            If case_note_found = False then
                task_based_array(do_this_const, item) = "review this"
                screening_count = screening_count + 1
            End if

'this is where I go to dail dail '
			CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
			DO
				EMReadScreen dail_check, 4, 2, 48
				If next_dail_check <> "DAIL" then
					MAXIS_case_number = ""
					CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
				End if
			Loop until dail_check = "DAIL"

				EMWriteScreen worker, 21, 6
				transmit
				transmit 'transmit past 'not your dail message'

			    Call dail_type_selection
			    EMReadScreen number_of_dails, 1, 3, 67		'Reads where the count of DAILs is listed

				DO
					If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped

					dail_row = 6			'Because the script brings each new case to the top of the page, dail_row starts at 6.
					DO
						dail_type = ""
						dail_msg = ""

					    'Determining if there is a new case number...
					    EMReadScreen new_case, 8, dail_row, 63
					    new_case = trim(new_case)
					    IF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
							Call write_value_and_transmit("T", dail_row, 3)
							dail_row = 6
						ELSEIF new_case = "CASE NBR" THEN
						    '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
						    Call write_value_and_transmit("T", dail_row + 1, 3)
							dail_row = 6
						End if

			            'Reading the DAIL Information
						EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
			            MAXIS_case_number = trim(MAXIS_case_number)
			            MAXIS_case_number = right("00000000" & MAXIS_case_number, 8) 'outputs in 8 digits format

			            EMReadScreen dail_type, 4, dail_row, 6

			            EMReadScreen dail_msg, 61, dail_row, 20
						dail_msg = trim(dail_msg)

			            EMReadScreen dail_month, 8, dail_row, 11
			            dail_month = trim(dail_month)

			            stats_counter = stats_counter + 1   'I increment thee
			            Call non_actionable_dails   'Function to evaluate the DAIL messages

			            IF add_to_excel = True then
							'--------------------------------------------------------------------...and put that in Excel.
							objExcel.Cells(excel_row, ).Value = worker
							objExcel.Cells(excel_row, .Value = MAXIS_case_number
								'objExcel.Cells(excel_row, ).Value = dail_month
							objExcel.Cells(excel_row, ).Value = dail_month
							'objExcel.Cells(excel_row, '5').Value = dail_msg
							excel_row = excel_row + 1

							Call write_value_and_transmit("D", dail_row, 3)' why dont we just keep this in and delete the unactionable dails @ILSE?'
							EMReadScreen other_worker_error, 13, 24, 2
							If other_worker_error = "** WARNING **" then transmit
							deleted_dails = deleted_dails + 1
						else
							add_to_excel = False
							dail_row = dail_row + 1
			                If len(dail_month) = 5 then
			                    output_year = ("20" & right(dail_month, 2))
			                    output_month = left(dail_month, 2)
			                    output_day = "01"
			                    dail_month = output_year & "-" & output_month & "-" & output_day
			                elseif trim(dail_month) <> "" then
			                    'Adjusting data for output to SQL
			                    output_year     = DatePart("yyyy",dail_month)   'YYYY-MM-DD format
			                    output_month    = right("0" & DatePart("m", dail_month), 2)
			                    output_day      = DatePart("d", dail_month)
			                    dail_month = output_year & "-" & output_month & "-" & output_day
			                End if

			                dail_string = worker & " " & MAXIS_case_number & " " & dail_type & " " & dail_month & " " & dail_msg
			                'If the case number is found in the string of case numbers, it's not added again.
			                If instr(all_dail_array, "*" & dail_string & "*") then
			                    If dail_type = "HIRE" then
			                        add_to_array = True
			                    Else
			                        add_to_array = False
			                    End if
			                    'msgbox "Duplicate Found: " & dail_string & vbcr & add_to_array
			                else
			                    add_to_array = True
			                End if

			                If add_to_array = True then
			                    ReDim Preserve DAIL_array(4, DAIL_count)	'This resizes the array based on the number of rows in the Excel File'
			            	    DAIL_array(worker_const,	           DAIL_count) = worker
			            	    DAIL_array(maxis_case_number_const,    DAIL_count) = MAXIS_case_number
			            	    DAIL_array(dail_type_const, 	       DAIL_count) = dail_type
			            	    DAIL_array(dail_month_const, 		   DAIL_count) = dail_month 'if they have the same constants and variables will they freak out '
			            	    DAIL_array(dail_msg_const, 		       DAIL_count) = dail_msg
			                    Dail_count = DAIL_count + 1
			                    all_dail_array = trim(all_dail_array & dail_string & "*") 'Adding MAXIS case number to case number string
			                    dail_string = ""
			                else
			                    false_count = false_count + 1
			                End if
						End if

						EMReadScreen message_error, 11, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.
						If message_error = "NO MESSAGES" then
							CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
							Call write_value_and_transmit(worker, 21, 6)
							transmit   'transmit past 'not your dail message'
			                Call dail_type_selection
							exit do
						End if

						'...going to the next page if necessary
						EMReadScreen next_dail_check, 4, dail_row, 4
						If trim(next_dail_check) = "" then
							PF8
							EMReadScreen last_page_check, 21, 24, 2
							If last_page_check = "THIS IS THE LAST PAGE" then
								all_done = true
								exit do
							Else
								dail_row = 6
							End if
						End if
					LOOP
					IF all_done = true THEN exit do
				LOOP
			Next

        End if
    End if
Next

'Excel output of cases and information in their applicable categories - PRIV, Req EXP Processing, Exp Screening Required, Not Expedited
Msgbox "Output to Excel Starting."      'warning message to whomever is running the script

'time line of actual runs
'todo save as copy and see how long it takes to run their actual list'

    ObjExcel.Worksheets.Add().Name = task_status
	ObjExcel.Cells(1, 1).Value = "Date Assigned"
	ObjExcel.Cells(1, 2).Value = "SSR Name"
	ObjExcel.Cells(1, 3).Value = "Case Number"
	ObjExcel.Cells(1, 4).Value = "Case Name"
	ObjExcel.Cells(1, 5).Value = "Basket"
	ObjExcel.Cells(1, 6).Value = "Assigned to"
	ObjExcel.Cells(1, 7).Value = "Assigned Worker X127#"
	ObjExcel.Cells(1, 8).Value = "Does worker log indicate they could work the case?"
	ObjExcel.Cells(1, 9).Value = "Case logged by assigned worker?"
	ObjExcel.Cells(1, 10).Value = "Case Note Date"
	ObjExcel.Cells(1, 11).Value = "Worker who made case note same as assigned worker"
	ObjExcel.Cells(1, 12).Value = "Does case note title contain keyword"
	ObjExcel.Cells(1, 13).Value = "DAIL Count"
	ObjExcel.Cells(1, 14).Value = "DAIL Type"
	ObjExcel.Cells(1, 15).Value = "EWS ECF Item Count"
	ObjExcel.Cells(1, 16).Value = "ECF Form Types"
	ObjExcel.Cells(1, 17).Value = "Oldest ECF APPL Date"
	ObjExcel.Cells(1, 18).Value = "Comments"
	ObjExcel.Cells(1, 19).Value = "Pending over 30 days"
	ObjExcel.Cells(1, 20).Value = "Interview Completed"

	objExcel.Columns(1).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
    objExcel.Columns(10).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
    objExcel.Columns(17).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY

    Excel_row = 2

    For item = 0 to UBound(task_based_array, 2)
	    objExcel.Cells(excel_row, 1).Value = task_based_array(date_assigned_const,      item)
	    objExcel.Cells(excel_row, 2).Value = task_based_array(SSR_name_const, 		    item)
	    objExcel.Cells(excel_row, 3).Value = task_based_array(maxis_case_number_const,  item) = MAXIS_case_number
	    objExcel.Cells(excel_row, 4).Value = task_based_array(case_name_const,          item) = MAXIS_case_name
	    objExcel.Cells(excel_row, 5).Value = task_based_array(basket_const,  		    item) = ""
	    objExcel.Cells(excel_row, 6).Value = task_based_array(assigned_to_const,        item) = ""
	    objExcel.Cells(excel_row, 7).Value = task_based_array(worker_number_const,      item) = worker_number
	    objExcel.Cells(excel_row, 8).Value = task_based_array(do_this_const,            item)
	    objExcel.Cells(excel_row, 9).Value = task_based_array(case_logged_const,    	item)
	    objExcel.Cells(excel_row, 10).Value = task_based_array(case_note_date_const,    item)
	    objExcel.Cells(excel_row, 11).Value = task_based_array(case_note_match_const,   item)
	    objExcel.Cells(excel_row, 12).Value = task_based_array(case_note_keyword_const, item)
	    objExcel.Cells(excel_row, 13).Value = task_based_array(DAIL_count_const,        item)
	    objExcel.Cells(excel_row, 14).Value = task_based_array(DAIL_type_const,         item)
	    objExcel.Cells(excel_row, 15).Value = task_based_array(ECF_form_const,   		item)
	    objExcel.Cells(excel_row, 16).Value = task_based_array(ECF_type_const,          item)
	    objExcel.Cells(excel_row, 17).Value = task_based_array(oldest_APPL_date_const,  item) = trim(application_date)
	    objExcel.Cells(excel_row, 18).Value = task_based_array(prev_comments_const,     item) = program_ID
	    objExcel.Cells(excel_row, 19).Value = task_based_array(case_status_const,       item) = days_pending
	    objExcel.Cells(excel_row, 20).Value = task_based_array(interview_const,         item) = trim(interview_date)
	    'making space in the array for these variables, but valuing them as "" for now
        excel_row = excel_row + 1
    Next

    FOR i = 1 to 8		'formatting the cells
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT


objWorkbook.Save()  'saves existing workbook as same name
objExcel.Quit

'logging usage stats
STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success, the run is complete. The workbook has been saved.")
