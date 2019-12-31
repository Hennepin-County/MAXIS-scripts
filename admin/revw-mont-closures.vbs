'Required for statistical purposes===============================================================================
name_of_script = "BULK - REVW-MONT CLOSURES.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 145                     'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
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
call changelog_update("05/17/2019", "Moved to BZ folder, Added all agency option and cleaned up old coding.", "Ilse Ferris, Hennepin County")
call changelog_update("06/27/2018", "Added/updated closing message.", "Ilse Ferris, Hennepin County")
call changelog_update("03/12/2018", "Removing information from CASE/NOTE regarding HC application.", "Casey Love, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'USER DETERMINATION-------------------------------------------------
'Getting network ID info for use by the next part of the script.
Set objNet = CreateObject("WScript.NetWork")

'Determines user to enable debugging features. Add individuals to this if...then to include them with developer mode.
If ucase(objNet.UserName) = "ILFE001" then
    user_name = "Ilse"
    user_found = true
Elseif ucase(objNet.UserName) = "CALO001" then
    user_name = "Casey"
    user_found = true
Elseif ucase(objNet.UserName) = "WFS395" then
    user_name = "MiKayla"
    user_found = true
Else
    user_found = false
End if

inquiry_testing = MsgBox("Master Developer detected. Hello, " & user_name & "." & vbcr & "Would you like to enable inquiry testing and bypass date restrictions?", vbYesNoCancel)
If inquiry_testing = vbCancel then cancel_without_confirmation   'If cancelled...

'There's a date restriction on this script: it should only run the last week of the month.
'If inquiry_testing = vbYes, it should bypass this restriction. Otherwise, it should filter through.
'Because of that, I've included "if inquiry_testing <> vbYes" at the beginning of my date restriction.
If inquiry_testing <> vbYes and datepart("m", dateadd("d", 8, date)) = datepart("m", date) then script_end_procedure("This script cannot be run until the last week of the month.")

'----------------------CONNECTING TO BLUEZONE, RUNNING THE DIALOG, AND NAVIGATING TO REPT/REVW
EMConnect ""

MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year  = CM_plus_1_yr

next_month = CM_plus_1_mo & "/" & CM_plus_1_yr

all_workers_check = 1
MONT_check = 1
REVW_check = 1

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 206, 90, "REVW/MONT closures"
  EditBox 160, 5, 40, 15, worker_number
  CheckBox 5, 25, 150, 10, "Check here to run this query agency-wide.", all_workers_check
  CheckBox 15, 55, 120, 10, "REPT/MONT? (HRFs)", MONT_check
  CheckBox 15, 70, 120, 10, "REPT/REVW? (CSRs and ARs)", REVW_check
  ButtonGroup ButtonPressed
    OkButton 160, 50, 40, 15
    CancelButton 160, 70, 40, 15
  GroupBox 5, 40, 145, 45, "Case note closing/incomplete cases from:"
  Text 5, 10, 145, 10, "Enter 7-digit x number(s), comma seperated:"
EndDialog
Do
	Do
        err_msg = ""
		Dialog Dialog1
		cancel_without_confirmation
        If trim(worker_number) = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases."
  		If trim(worker_number) <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases, not both options."
  	  	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
  	LOOP until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ", ")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & "," & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next

	worker_array = split(worker_array, ",")
End if

'THIS PART DOES THE REPT REVW----------------------------------------------------------------------------------------------------
priv_case_list = ""
Call MAXIS_footer_month_confirmation

If REVW_check = 1 then
    For each worker in worker_array
        'msgbox worker
	    Call navigate_to_MAXIS_screen("rept", "revw")
        Call write_value_and_transmit(worker, 21, 6)

	    row = 7
	    Do
	    	EMReadScreen MAXIS_case_number, 8, row, 6																'Gets case number
	    	EMReadScreen cash_status, 1, row, 35															'Checks for cash status
	    	If cash_status = "N" or cash_status = "I" then are_programs_closing = True						'If "N" or "I", adds to the array
	    	EMReadScreen FS_status, 1, row, 45																'Checks for FS status
	    	If FS_status = "N" or FS_status = "I" then are_programs_closing = True							'If "N" or "I", adds to the array
	    	EMReadScreen HC_status, 1, row, 49																'Checks for FS status
	    	If HC_status = "N" or HC_status = "I" then 														'If "N" or "I", checks additional info before adding to the array
	    		EMReadScreen exempt_IR_check, 1, row, 51													'Checks for exempt IRs (starred IRs)
	    		If exempt_IR_check <> "*" then are_programs_closing = True									'Only adds cases to array if they are not exempt from an IR
	    	End if

	    	'If the above found the case is closing, it adds to the array.
	    	If are_programs_closing = True then case_number_list = trim(case_number_list & " " & trim(MAXIS_case_number))
	    	are_programs_closing = ""		'Clears out variable

	    	row = row + 1
	    	If row = 19 then
	    		PF8
	    		EMReadScreen last_check, 4, 24, 14
	    		row = 7
	    	End if
	    Loop until trim(MAXIS_case_number) = "" or last_check = "LAST"

	    case_number_array = split(case_number_list)
	      '-----------------------NAVIGATING TO EACH CASE AND CASE NOTING THE ONES THAT ARE CLOSING
	    For each MAXIS_case_number in case_number_array
	    	CALL navigate_to_MAXIS_screen("rept", "revw")  'Reads MAGI code for each case.
	    	EMReadScreen MAGI_code, 4, 7, 54
	    	EMReadScreen priv_check, 4, 24, 14 'Checking if we can get into stat (need to bypass Privileged cases)
	    	IF priv_check <> "PRIV" THEN 'Not privileged, we can go ahead and do everything
	    		call navigate_to_MAXIS_screen("stat", "revw") 'In case of error prone cases
	    		EMReadScreen cash_review_code, 1, 7, 40
	    		EMReadScreen FS_review_code, 1, 7, 60
	    		EMReadScreen HC_review_code, 1, 7, 73
	    		If cash_review_code = "N" then cash_review_status = "closing for no renewal CAF."
	    		If cash_review_code = "I" then cash_review_status = "closing for incomplete review. See previous case notes for details on what's needed."
	    		If FS_review_code = "N" then
	    			EMWriteScreen "x", 5, 58
	    			transmit
	    			EMReadScreen recertification_date, 8, 9, 64
	    			recertification_date = cdate(replace(recertification_date, " ", "/"))
	    			If datepart("m", recertification_date) = datepart("m", dateadd("m", 1, now)) then
	    				FS_review_document = "renewal CAF"
	    			Else
	    				FS_review_document = "CSR"
	    			End if
	    			FS_review_status = "closing for no " & FS_review_document & "."
	    			transmit
	    		End if
	    		If FS_review_code = "I" then FS_review_status = "closing for incomplete review. See previous case notes for details on what's needed."
	    		If HC_review_code = "N" then
	    			EMWriteScreen "x", 5, 71
	    			transmit
	    			EMReadScreen recertification_date, 8, 9, 27
	    			recertification_date = cdate(replace(recertification_date, " ", "/"))
	    			If datepart("m", recertification_date) = datepart("m", dateadd("m", 1, now)) then
	    				HC_review_document = "renewal document"
	    			Else
	    				HC_review_document = "CSR"
	    			End if
	    			HC_review_status = "closing for no " & HC_review_document & "."
	    			transmit
	    		End if
	    		If HC_review_code = "I" then HC_review_status = "closing for incomplete review. See previous case notes for details on what's needed."

	    		'Checking for the active CASH program. If the case is GRH, MSA, GA, MFIP, or DWP, the client is eligible for an additional 30 day reinstatement period.
	    		'If the case is RCA, the client is not eligible for an additional 30 day reinstatement period for no-or-incomplete review.
	    		'For policy on the matter, see Bulletin #14-69-05 (http://www.dhs.state.mn.us/main/groups/publications/documents/pub/dhs16_185044.pdf)
	    		IF cash_review_code = "N" OR cash_review_code = "I" THEN
	    			EMWriteScreen "PROG", 20, 71
	    			transmit

	    			EMReadScreen cash_status, 4, 9, 74
	    			IF cash_status = "ACTV" THEN
	    				cash_prog = "GR"
	    			ELSE
	    				EMReadScreen cash_status, 4, 6, 74
	    				IF cash_status = "ACTV" THEN
	    					EMReadScreen cash_prog, 2, 6, 67
	    				ELSE
	    					EMReadScreen cash_status, 4, 7, 74
	    					EMReadScreen cash_prog, 2, 7, 67
	    				END IF
	    			END IF

	    			IF cash_prog = "GR" OR cash_prog = "GA" OR cash_prog = "MS" OR cash_prog = "DW" OR cash_prog = "MF" THEN
	    				elig_for_cash_rein = True
	    			ELSEIF cash_prog = "RC" THEN
	    				elig_for_cash_rein = False
	    			END IF
	    			EMWriteScreen "REVW", 20, 71
	    		END IF

	    		'---------------THIS SECTION FIGURES OUT WHEN PROGRAMS CAN TURN IN NEW RENEWALS AND WHEN THEY BECOME INTAKES AGAIN
	    		EMReadScreen first_of_working_month, 5, 20, 55		'Used by the following logic to determine the first date
	    		first_of_working_month = cdate(replace(first_of_working_month, " ", "/01/"))	'Added "/01/" to make it a date

	    		If HC_review_status <> "" then	'Added additional logic as currently MAGI clients get an additonal 4 months to turn in renewal paperwork.
	    			last_day_to_turn_in_HC_docs = dateadd("d", -1, (dateadd("m", 4, first_of_working_month)))
	    			HC_intake_date = dateadd("d", 1, last_day_to_turn_in_HC_docs)
	    		End If

	    		If FS_review_status <> "" then
	    			If FS_review_code = "I" or FS_review_document = "CSR" then
	    				last_day_to_turn_in_SNAP_docs = dateadd("d", -1, (dateadd("m", 1, first_of_working_month)))
	    				SNAP_intake_date = dateadd("m", 1, first_of_working_month)
	    			Else
	    				last_day_to_turn_in_SNAP_docs = dateadd("d", -1, first_of_working_month)
	    				SNAP_intake_date = first_of_working_month
	    			End if
	    		End if

	    		If cash_review_status <> "" then
	    			IF elig_for_cash_rein = True THEN
	    				last_day_to_turn_in_cash_docs = dateadd("d", -1, dateadd("M", 1, first_of_working_month))
	    				cash_intake_date = dateadd("M", 1, first_of_working_month)
	    			ELSEIF elig_for_cash_rein = False THEN
	    				last_day_to_turn_in_cash_docs = dateadd("d", -1, first_of_working_month)
	    				cash_intake_date = first_of_working_month
	    			END IF
	    		End if

	    		'---------------NOW IT CASE NOTES
	    		If inquiry_testing <> vbYes then
	    			Call start_a_blank_CASE_NOTE
	    			If HC_review_code = "I" or FS_review_code = "I" or cash_review_code = "I" then
	    				call write_variable_in_case_note("-Programs closing for " & next_month & ": Incomplete review-")
	    			Else
	    				call write_variable_in_case_note("-Programs closing for " & next_month & ": No review-")
	    			End if
	    			call write_bullet_and_variable_in_case_note("Cash", cash_review_status)
	    			call write_bullet_and_variable_in_case_note("SNAP", FS_review_status)
	    			call write_bullet_and_variable_in_case_note("HC", HC_review_status)

	    			'last_day_to_turn_in_cash_docs = trim(last_day_to_turn_in_cash_docs)
	    			'IF last_day_to_turn_in_cash_docs <> "" THEN call write_variable_in_case_note("* Client has until " & last_day_to_turn_in_cash_docs & " to turn in CAF/CSR and/or proofs for cash.")
	    			'
                    'last_day_to_turn_in_SNAP_docs = trim(last_day_to_turn_in_SNAP_docs)
	    			'IF last_day_to_turn_in_SNAP_docs <> "" THEN call write_variable_in_case_note("* Client has until " & last_day_to_turn_in_SNAP_docs & " to turn in CAF/CSR and/or proofs for SNAP.")

                    last_day_to_turn_in_HC_docs = trim(last_day_to_turn_in_HC_docs)
	    			IF last_day_to_turn_in_HC_docs <> "" THEN call write_variable_in_case_note("* Client has until " & last_day_to_turn_in_HC_docs & " to turn in HC review doc and/or proofs.")

                    If cash_review_status <> "" and cash_intake_date <> "" then call write_variable_in_case_note("* Client needs to turn in new application for cash on " & cash_intake_date & ".")
	    			If FS_review_status <> "" and SNAP_intake_date <> "" then call write_variable_in_case_note("* Client needs to turn in new application for SNAP on " & SNAP_intake_date & ".")
	    		Else	'special handling for inquiry_testing (developers testing scenarios)
	    			string_for_msgbox = 	"Cash: " & cash_review_status & chr(10) & _
	    									"SNAP: " & FS_review_status & chr(10) & _
	    									"HC: " & HC_review_status & chr(10) & _
	    									"Last CASH doc date: " & last_day_to_turn_in_cash_docs & chr(10) & _
	    									"New CASH application would be needed on or after this date: " & cash_intake_date & chr(10) & _
	    									"Last SNAP doc date: " & last_day_to_turn_in_SNAP_docs & chr(10) & _
	    									"New SNAP application would be needed on or after this date: : " & SNAP_intake_date & chr(10) & _
	    									"Last HC doc date: " & last_day_to_turn_in_HC_docs & chr(10) & _
	    									"New HC application would be needed on or after this date: " & HC_intake_date
	    			debugging_MsgBox = MsgBox(string_for_msgbox, vbOKCancel)
	    			If debugging_MsgBox = vbCancel then cancel_without_confirmation
	    		End if
	    	ELSE 'This is a privileged case, we need to skip to the next one, so we won't do anything with it
	    		priv_case_list = priv_case_list & " " & MAXIS_case_number 'saving a list of priv cases for later.
	    	END IF
	    		'----------------NOW IT RESETS THE VARIABLES FOR THE REVIEW CODES, STATUS, AND DATES
	    	first_of_working_month = ""
	    	cash_review_code = ""
	    	FS_review_code = ""
	    	HC_review_code = ""
	    	cash_review_status = ""
	    	FS_review_status = ""
	    	HC_review_status = ""
	    	last_day_to_turn_in_cash_docs = ""
	    	last_day_to_turn_in_SNAP_docs = ""
	    	last_day_to_turn_in_HC_docs = ""
	    	cash_intake_date = ""
	    	SNAP_intake_date = ""
	    	HC_intake_date = ""
	    	cash_prog = ""
	    	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
	    Next

	    'Call navigate_to_MAXIS_screen("rept", "revw")
	    'Call write_value_and_transmit(worker_number, 21, 6)
    next
End If

'Resetting the case number array
case_number_list = ""

'THIS PART DOES THE REPT MONT----------------------------------------------------------------------------------------------------
If MONT_check = 1 then
    For each worker in worker_array
        'msgbox worker
        Call navigate_to_MAXIS_screen("rept", "mont")           'Navigating to MONT
        Call write_value_and_transmit(worker, 21, 6)

        row = 7     'Setting the variable for the following do...loop
        Do  'This reads the case number and program status. If an "N" or "I" is detected it will add to the case_number_list variable.
            EMReadScreen MAXIS_case_number, 8, row, 6
            EMReadScreen program_status, 9, row, 45
            are_programs_closing = instr(program_status, "N") <> 0 or instr(program_status, "I") <> 0
            If are_programs_closing = True then case_number_list = trim(case_number_list & " " & trim(MAXIS_case_number))
            row = row + 1
            If row = 19 then
                PF8
                EMReadScreen last_check, 4, 24, 14
                row = 7
            End if
        Loop until trim(MAXIS_case_number) = "" or last_check = "LAST"

        'Creating an array out of the case number array
        case_number_array = split(case_number_list)

        'Navigating to each case, and case noting the ones that are closing.
        For each MAXIS_case_number in case_number_array
            'Going to the case, checking for error prone
            call navigate_to_MAXIS_screen("stat", "mont")
        	EMReadScreen priv_check, 4, 24, 14 'Checking if we can get into stat (need to bypass Privileged cases)
        	IF priv_check <> "PRIV" THEN 'Not privileged, we can go ahead and do everything
        		call navigate_to_MAXIS_screen("stat", "mont") 'In case of error prone cases

        		'Reading the review codes, converting them to a status update for the case note
        		EMReadScreen cash_review_code, 1, 11, 43
        		EMReadScreen FS_review_code, 1, 11, 53
        	  '---------------NOW IT CASE NOTES
        		PF4
        		PF9

        		If inquiry_testing <> vbYes then
                    If FS_review_code = "I" and cash_review_code = "I" then
                    elseif FS_review_code = "I" then
                        header = "--SNAP Closing for " & next_month
                    elseif cash_review_code = "I" then
                        header = "--Cash Closing for " & next_month
                    End if
                    If FS_review_code = "I" or cash_review_code = "I" then
        		      call write_variable_in_CASE_NOTE("-" & header & ": Incomplete HRF--")
        		    Else
        		      call write_variable_in_CASE_NOTE("-" & header & ": HRF not provided--")
        		    End if
        		else 'special handling for inquiry_testing (developers testing scenarios)
	    			string_for_msgbox = 	"Cash: " & cash_review_status & chr(10) & _
	    									"SNAP: " & FS_review_status & chr(10) & _
                                            "Last CASH doc date: " & last_day_to_turn_in_cash_docs & chr(10) & _
	    									"New CASH application would be needed on or after this date: " & cash_intake_date & chr(10) & _
	    									"Last SNAP doc date: " & last_day_to_turn_in_SNAP_docs & chr(10) & _
	    									"New SNAP application would be needed on or after this date: : " & SNAP_intake_date & chr(10) & _
	    			debugging_MsgBox = MsgBox(string_for_msgbox, vbOKCancel)
	    			If debugging_MsgBox = vbCancel then cancel_without_confirmation
                End if
        	ELSE 'Prived case, add the case number to the list
        		priv_case_list = priv_case_list & " " & MAXIS_case_number
        	END IF

          '----------------NOW IT RESETS THE VARIABLES FOR THE REVIEW CODES, STATUS, AND DATES
            cash_review_code = ""
            GRH_review_code = ""
            FS_review_code = ""
            HC_review_code = ""
            cash_review_status = ""
            GRH_review_status = ""
            FS_review_status = ""
            HC_review_status = ""
            first_of_working_month = ""
            last_day_to_turn_in_docs = ""
            intake_date = ""
        	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
        Next

        'Call navigate_to_MAXIS_screen("rept", "mont")
        'Call write_value_and_transmit(worker_number, 21, 6)
    Next
End If

MsgBox "Success! All cases that are coded in REPT/REVW and/or REPT/MONT as either an ''N'' or an ''I'' have been case noted for why they're closing, and what documents need to get turned in."
IF trim(priv_case_list) <> "" THEN MsgBox "Please note the following case numbers that are PRIVILEGED and could not be updated by the script.  They must be case noted manually:" & VbCr & priv_case_list

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created.")
