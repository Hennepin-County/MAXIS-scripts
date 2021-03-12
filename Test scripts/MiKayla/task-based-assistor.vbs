'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - TASK BASED ASSISTOR.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 60                      'manual run time in seconds
STATS_denomination = "M"       			   'M is for each CASE
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

call changelog_update("12/29/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Custom function for this script----------------------------------------------------------------------------------------------------
FUNCTION get_case_stuff
	back_to_self
	EMWriteScreen MAXIS_case_number, 18, 43
	'just in case we want to know later '
	Call navigate_to_MAXIS_screen("CASE", "CURR")
	EMReadScreen CURR_panel_check, 4, 2, 55
	If CURR_panel_check <> "CURR" then ObjExcel.Cells(excel_row, 2).Value = ""

	EMReadScreen case_status, 8, 8, 9
	case_status = trim(case_status)
	ObjExcel.Cells(excel_row, 2).Value = case_status
	MAXIS_case_number = ""
	excel_row = excel_row + 1
	'using new variable count to calculate percentages
	IF case_status = "ACTIVE" then active_status = active_status + 1
	IF case_status = "APP OPEN" then active_status = active_status + 1

	IF case_status = "APP CLOS" then inactive_status = inactive_status + 1
	IF case_status = "INACTIVE" then inactive_status = inactive_status + 1

	If case_status = "CAF2 PEN" then pending_status = pending_status + 1
	If case_status = "CAF1 PEN" then pending_status = pending_status + 1

	IF case_status = "REIN" then rein_status = rein_status + 1
	STATS_counter = STATS_counter + 1

	'to gather case name'
	Call navigate_to_MAXIS_screen ("STAT", "MEMB")
	EMReadScreen panel_check, 4, 2, 55
	If panel_check <> "MEMB" then ObjExcel.Cells(excel_row, 2).Value = "ERROR"
	EMReadScreen first_name, 12, 6, 63
	EMReadScreen last_name, 25, 6, 30
	client_info = client_info & replace(first_name, "_", "") & " " & replace(last_name, "_", "")
	client_info = right(client_info, len(client_info) - 1)

	Call navigate_to_MAXIS_screen("CASE", "NOTE")       'First to case note to find what has ahppened'
	EMReadScreen panel_check, 4, 2, 55
	If panel_check <> "NOTE" then ObjExcel.Cells(excel_row, 2).Value = "ERROR"
' count
'please read expedited review
	day_before_app = DateAdd("d", -1, ALL_PENDING_CASES_ARRAY(application_date, case_entry)) 'will set the date one day prior to app date'

		note_row = 5        'these always need to be reset when looking at Case note
		note_date = ""
		note_title = ""
		appt_date = ""
		Do                  'this do-loop moves down the list of case notes - looking at each row in MAXIS
			EMReadScreen note_date, 8, note_row, 6      'reading the date of the row
			EMReadScreen note_title, 55, note_row, 25   'reading the header of the note
			note_title = trim(note_title)               'trim it down

			IF note_date = "        " then Exit Do      'if the case is new, we will hit blank note dates and we don't need to read any further
			note_row = note_row + 1                     'going to the next row to look at the next notws
			IF note_row = 19 THEN                       'if we have reached the end of the list of case notes then we will go to the enxt page of notes
				PF8
				note_row = 5
			END IF
			EMReadScreen next_note_date, 8, note_row, 6 'looking at the next note date
			IF next_note_date = "        " then Exit Do
		Loop until datevalue(next_note_date) < day_before_app 'looking ahead at the next case note kicking out the dates before app'

		go_to_top_of_notes      'this is a function defined above so that if we need to read for different notes we don't miss ones on the first pages if we went to PF8
'SET A STARTING COLUMN'

END FUNCTION


'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""

this_month = CM_mo & " " & CM_yr
next_month = CM_plus_1_mo & " " & CM_plus_1_yr
CM_minus_2_mo =  right("0" & DatePart("m", DateAdd("m", -2, date)), 2)

'Finding the right folder to automatically save the file
month_folder = "DAIL " & CM_mo & "-" & DatePart("yyyy", date) & ""
task_folder = replace(this_month, " ", "-") & " Task Based"
report_date = replace(date, "/", "-")

'dialog and dialog DO...Loop
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 301, 100, "ADMIN - TASK BASED ASSISTOR"
  ButtonGroup ButtonPressed
    PushButton 15, 40, 60, 15, "Browse...", select_a_file_button
  EditBox 80, 40, 205, 15, file_selection_path
  ButtonGroup ButtonPressed
    OkButton 190, 80, 50, 15
    CancelButton 245, 80, 50, 15
  Text 40, 60, 230, 10, "This script should be used with the task based assignment."
  Text 15, 15, 275, 20, "Select the Excel file that contains your information by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 285, 70, "Using this script:"
EndDialog

Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog Dialog1
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 126, 50, "Select the excel row to start"
  EditBox 75, 5, 40, 15, excel_row_to_restart
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog

DO
    dialog Dialog1
    If buttonpressed = 0 then stopscript								'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in


'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True
excel_row = excel_row_to_restart

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X NUMBER"
objExcel.Cells(1, 2).Value = "CASE #"
objExcel.Cells(1, 3).Value = "DAIL TYPE"
objExcel.Cells(1, 4).Value = "DAIL MO."
objExcel.Cells(1, 5).Value = "DAIL MESSAGE"

FOR i = 1 to 5		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

DIM DAIL_array()
ReDim DAIL_array(4, 0)
Dail_count = 0              'Incremental for the array
all_dail_array = "*"    'setting up string to find duplicate DAIL messages. At times there is a glitch in the DAIL, and messages are reviewed a second time.
false_count = 0

'constants for array
const worker_const	            = 0
const maxis_case_number_const   = 1
const dail_type_const 	        = 2
const dail_month_const		    = 3
const dail_msg_const		    = 4

'Sets variable for all of the Excel stuff
excel_row = 2
deleted_dails = 0	'establishing the value of the count for deleted deleted_dails

CALL navigate_to_MAXIS_screen("DAIL", "DAIL")

'This for...next contains each worker indicated above
For each worker in worker_array
    MAXIS_case_number = ""
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
				objExcel.Cells(excel_row, 1).Value = worker
				objExcel.Cells(excel_row, 2).Value = MAXIS_case_number
				objExcel.Cells(excel_row, 3).Value = dail_type
				objExcel.Cells(excel_row, 4).Value = dail_month
				objExcel.Cells(excel_row, 5).Value = dail_msg
				excel_row = excel_row + 1

				Call write_value_and_transmit("D", dail_row, 3)
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
            	    DAIL_array(dail_month_const, 		   DAIL_count) = dail_month
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


script_end_procedure("All done.")
