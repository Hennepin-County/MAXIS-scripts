'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - SEND Missed NOMIs.vbs"
start_time = timer
STATS_counter = 0			 'sets the stats counter at one
STATS_manualtime = 304			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
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
'defining this function here because it needs to not end the script if a MEMO fails.
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
'DIALOGS ===================================================================================================================

'Initial Dialog which requests a file path for the excel file
BeginDialog recert_list_dlg, 0, 0, 361, 105, "On Demand Recertifications"
  EditBox 130, 60, 175, 15, recertification_cases_excel_file_path
  ButtonGroup ButtonPressed
    PushButton 310, 60, 45, 15, "Browse...", select_a_file_button
  EditBox 75, 85, 140, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 250, 85, 50, 15
    CancelButton 305, 85, 50, 15
  Text 10, 10, 170, 10, "Welcome to the On Demand Recertification Notifier."
  Text 10, 25, 340, 30, "This script will send an Appointment Notice or NOMI for recertification for a list of cases in a county that currently has an On Demand Waiver in effect for interviews. If your county does not have this waiver, this script should not be used."
  Text 10, 65, 120, 10, "Select an Excel file for recert cases:"
  Text 10, 90, 60, 10, "Worker Signature"
EndDialog

Do
	Dialog recert_list_dlg
	If ButtonPressed = cancel then stopscript
	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(recertification_cases_excel_file_path, ".xlsx")
Loop until ButtonPressed = OK and recertification_cases_excel_file_path <> "" and worker_signature <> ""

'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
call excel_open(recertification_cases_excel_file_path, True, True, ObjExcel, objWorkbook)
'Activates worksheet based on user selection
objExcel.worksheets("All cases").Activate

MAXIS_footer_month = "05"
MAXIS_footer_year = "18"
last_day_of_recert = "04/30/18"

excel_row = 2
case_number_col = 2
nomi_success_col = 10
interview_col = 12
date_app_col = 13
revw_stat_col = 11

'creating a variable in the MM/DD/YY format to compare with date read from MAXIS
today_mo = DatePart("m", date)
today_mo = right("00" & today_mo, 2)

today_day = DatePart("d", date)
today_day = right("00" & today_day, 2)

today_yr = DatePart("yyyy", date)
today_yr = right(today_yr, 2)

today_date = today_mo & "/" & today_day & "/" & today_yr

Do
    MAXIS_case_number = objExcel.Cells(excel_row, case_number_col).Value
    MAXIS_case_number = trim(MAXIS_case_number)
    nomi_info = objExcel.Cells(excel_row, nomi_success_col).Value
    nomi_info = trim(nomi_info)
    interview_date =  objExcel.Cells(excel_row, interview_col).Value
    interview_date = trim(interview_date)

    If nomi_info = "PRIV" Then
        Call navigate_to_MAXIS_screen("STAT", "REVW")

        app_recvd = TRUE
        interview_complete = TRUE

        EMReadScreen recvd_date, 8, 13, 37
        recvd_date = replace(recvd_date, " ", "/")
        if recvd_date = "__/__/__" then
            app_recvd = FALSE
            recvd_date = ""
        end if

        EMReadScreen interview_date, 8, 15, 37
        interview_date = replace(interview_date, " ", "/")
        if interview_date = "__/__/__" then
            interview_complete = FALSE
            interview_date = ""
        end if

        EMReadScreen review_status, 1, 7, 60
        if review_status = "_" then review_status = ""

        If interview_complete = FALSE Then
            'Writing the SPEC MEMO - dates will be input from the determination made earlier.
            Call start_a_new_spec_memo_and_continue(memo_started)

            IF memo_started = True THEN         'The function will return this as FALSE if PF5 does not move past MEMO DISPLAY

                if app_recvd = TRUE then CALL write_variable_in_SPEC_MEMO("We received your Recertification on " & recvd_date & ".")
                if app_recvd = FALSE then CALL write_variable_in_SPEC_MEMO("Your Recertification has not been received yet.")
                CALL write_variable_in_SPEC_MEMO("")
                CALL write_variable_in_SPEC_MEMO("You must have an interview by " & last_day_of_recert & " or your benefits will end. ")
                CALL write_variable_in_SPEC_MEMO("")
                CALL write_variable_in_SPEC_MEMO("To complete a phone interview, call the EZ Info Line at 612-596-1300 between 9:00am and 4:00pm Monday through Friday.")
                CALL write_variable_in_SPEC_MEMO("")
                CALL write_variable_in_SPEC_MEMO("You may also come in to the office to complete an interview between 8:00 am and 4:30pm Monday through Friday.")
                CALL write_variable_in_SPEC_MEMO("")
                CALL write_variable_in_SPEC_MEMO("If we do not hear from you by " & last_day_of_recert & ", your benefits will end on " & last_day_of_recert & ".")


                PF4         'Submit the MEMO

                memo_row = 7                                            'Setting the row for the loop to read MEMOs
                notc_confirm = "N"         'Defaulting this to 'N'
                Do
                    EMReadScreen create_date, 8, memo_row, 19                 'Reading the date of each memo and the status
                    EMReadScreen print_status, 7, memo_row, 67
                    If create_date = today_date AND print_status = "Waiting" Then   'MEMOs created today and still waiting is likely our MEMO.
                        notc_confirm = "Y"             'If we've found this then no reason to keep looking.
                        Exit Do
                    End If

                    memo_row = memo_row + 1           'Looking at next row'
                Loop Until create_date = "        "

            ELSE
                notc_confirm = "N"         'Setting this as N if the MEMO failed
            END IF

            'writing a case note only if the notice was successfully sent.
            if notc_confirm = "Y" then

                start_a_blank_case_note

                EMSendKey("*** NOMI Sent for SNAP Recertification***")
                if app_recvd = TRUE then CALL write_variable_in_CASE_NOTE("* Recertification app received on " & recvd_date)
                if app_recvd = FALSE then CALL write_variable_in_CASE_NOTE("* Recertification app has NOT been received. Client must submit paperwork.")
                CALL write_variable_in_CASE_NOTE("* A notice was previously sent to client with detail about how to call in for an interview.")
                CALL write_variable_in_CASE_NOTE("* Client must call 612-596-1300 to complete interview.")
                If forms_to_arep = "Y" then CALL write_variable_in_CASE_NOTE("* Copy of notice sent to AREP.")
                If forms_to_swkr = "Y" then CALL write_variable_in_CASE_NOTE("* Copy of notice sent to Social Worker.")
                call write_variable_in_case_note("---")
                call write_variable_in_case_note(worker_signature)
            end if
        End If

        objExcel.Cells(excel_row, nomi_success_col).Value   = notc_confirm
        objExcel.Cells(excel_row, revw_stat_col).Value      = review_status
        objExcel.Cells(excel_row, date_app_col).Value       = recvd_date
        objExcel.Cells(excel_row, interview_col).Value      = interview_date
    End If

    excel_row = excel_row + 1
    'MsgBox excel_row
    back_to_self
Loop until MAXIS_case_number = ""
