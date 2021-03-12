'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - MONY-DISB CASHOUT.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = "450"                'manual run time in seconds
STATS_denomination = "C"       			'C is for each CASE
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
'TODO CALL ONLY_create_MAXIS_friendly_date'
call changelog_update("11/25/2020", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_password(false)
'Connecting to BlueZone, grabbing case number
EMConnect ""

'-------------------------------------------------------------------------------------------------DIALOG
BeginDialog Dialog1, 0, 0, 241, 90, "CASHOUT - MONY/DISB"
  ButtonGroup ButtonPressed
    PushButton 5, 30, 50, 15, "Browse...", select_a_file_button
  EditBox 65, 30, 170, 15, file_selection_path
  DropListBox 65, 55, 45, 15, "initial"+chr(9)+"revert", action_taken
  EditBox 200, 50, 15, 15, MAXIS_footer_month
  EditBox 220, 50, 15, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    OkButton 130, 70, 50, 15
    CancelButton 185, 70, 50, 15
  Text 10, 5, 220, 15, "Select the Excel file that contains the information by selecting the 'Browse' button, and finding the file."
  Text 150, 55, 50, 10, "Footer MM/YY:"
  Text 10, 60, 50, 10, "Action to take:"
EndDialog


'----------------------------------------------------------------------------------------------------THE SCRIPT
Do
	Do
		err_msg = ""
		Dialog Dialog1
		cancel_confirmation
		If ButtonPressed = select_a_file_button THEN
			If file_selection_path <> "" THEN 'This is handling for if the BROWSE button is pushed more than once'
				objExcel.Quit 'Closing the Excel file that was opened on the first push'
				objExcel = "" 	'Blanks out the previous file path'
			End If
			call file_selection_system_dialog(file_selection_path,".xlsx") 'allows the user to select the file'
		End If
		If err_msg <> "" THEN MsgBox err_msg
	Loop until err_msg = ""
	If objExcel = "" THEN call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
	If MAXIS_footer_month = "" THEN err_msg = err_msg & vbNewLine & "Please advise the footer month which you want this script to run."
	If MAXIS_footer_year = "" THEN err_msg = err_msg & vbNewLine & "Please advise the footer year which you want this script to run."
	If action_taken = "" THEN err_msg = err_msg & vbNewLine & "Please select which option you are taking with this script run."
	If file_selection_path = "" THEN err_msg = err_msg & vbNewLine & "Use the Browse button to select the file that has your data"
	If err_msg <> "" THEN MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

CALL check_for_MAXIS(False)
back_to_SELF
ObjExcel.Cells(1, 1).Value = "CASE NUMBER"
ObjExcel.Cells(1, 2).Value = "AMOUNT"
ObjExcel.Cells(1, 3).Value = "FS STATUS"
ObjExcel.Cells(1, 4).Value = "UPDATE MADE"
ObjExcel.Cells(1, 5).Value = "METHOD"
ObjExcel.Cells(1, 6).Value = "NOTES"
ObjExcel.Cells(1, 7).Value = "SPEC/WCOM CANCELED"
ObjExcel.Cells(1, 8).Value = "PRINT STATUS"
ObjExcel.Cells(1, 9).Value = "REVERTED"

excel_row = 2           'establishing the row to start

DO
	back_to_SELF
	EMWriteScreen MAXIS_footer_month, 20, 43			'goes back to self and enters the date that the user selcted'
	EMWriteScreen MAXIS_footer_year, 20, 46
	'Assign case number from Excel
	MAXIS_case_number = ObjExcel.Cells(excel_row, 1).Value
	MAXIS_case_number = trim(MAXIS_case_number)
	'Exiting if the case number is blank
	If MAXIS_case_number = "" then exit do
	EMWriteScreen MAXIS_case_number, 18, 43
	IF action_taken = "initial" THEN 'checking for case status and AREP information'
	    Call navigate_to_MAXIS_screen("CASE", "CURR")
	    row = 1                                                 'look for SNAP
        col = 1
        EMSearch "FS:", row, col
        IF row <> 0 Then
            EMReadScreen fs_status, 9, row, col + 4
            fs_status = trim(fs_status)
        END IF
    	If fs_status = "ACTIVE" or fs_status = "APP CLOSE" or fs_status = "APP OPEN" or fs_status = "PENDING" THEN
    		update_case = TRUE
    		case_active = TRUE
    	End If
    	If fs_status = "PENDING" Then
    		update_case = FALSE
    		case_active = FALSE
    	END If

		IF case_active = TRUE THEN
        	Call navigate_to_MAXIS_screen("STAT", "ADDR")
    		EMReadScreen priv_check, 4, 2, 50
    		IF priv_check = "SELF" THEN
    			action_note = "Privileged"
    		ELSE
    			EMReadScreen addr_line_01, 22, 6, 43 'this doesnt make sense come back to this'
    	    	Call navigate_to_MAXIS_screen("STAT", "ALTP")
    	    	EMReadScreen altp_addr_line_01, 22, 12, 37
				Call navigate_to_MAXIS_screen("STAT", "AREP")
				EMReadScreen arep_addr_line, 22, 05, 332 'if panel does nto exist it will not match'
				IF trim(addr_line_01) = trim(altp_addr_line_01) THEN
    	    		update_case = FALSE
    	       		action_note = "ADDR same ALTP"
				ELSEIF trim(addr_line_01) = trim(arep_addr_line) THEN
	    	    		update_case = FALSE
	    	       		action_note = "ADDR same AREP"
    	    	ELSE
				 	update_case = TRUE
				END IF
			END IF
		END IF
	END IF
		IF update_case = TRUE or action_taken = "revert" THEN
			Call navigate_to_MAXIS_screen("MONY", "DISB")
	    	EMReadscreen payment_method, 2, 5, 35
			EMReadscreen worker_mail_preference, 2, 9, 35
			EMReadScreen updated_mony_disb_date, 8, 9, 40
			IF payment_method = "DD" or payment_method = "EB" THEN
				update_case = FALSE
				action_note = "payment method"
			END IF
			IF worker_mail_preference = "RG" and action_taken = "initial" THEN
				PF9
			   	EMWriteScreen "IC", 9, 35 'Worker Mail Preference'
				EMWriteScreen "27", 10, 35 'Pick Up County '
				EMWriteScreen "02", 10, 47 'Pick Up Office'
                TRANSMIT
				EMReadScreen warning_error_message, 8, 24, 2   'checking the bottom for an error message
				warning_error_message = trim(warning_error_message)
				IF warning_error_message = "WARNING:" THEN 'we can transmit past warning messages and then look again
					TRANSMIT
					EMReadScreen error_message, 75, 24, 2   'checking the bottom for an error message
					error_message = trim(error_message)
				ELSEIF error_message <> "" THEN      'if there is anything here - assume an error
					update_case = FALSE
					action_note = error_message
					PF10
				ELSE
					action_note = "update complete"
				END IF
			ELSEIF worker_mail_preference = "IC" and action_taken = "initial" THEN
					update_case = FALSE
			   		action_note = "already updated to IC" & replace(updated_mony_disb_date, " ", "/")
					'is there another action needed on the panel?'
			ELSEIF worker_mail_preference = "IC" and action_taken = "revert" THEN
					PF9
				    EMWriteScreen "RG", 9, 35
					TRANSMIT
					EMReadScreen warning_error_message, 8, 24, 2   'checking the bottom for an error message
					warning_error_message = trim(warning_error_message)
					IF warning_error_message = "WARNING:" THEN 'we can transmit past warning messages and then look again
						TRANSMIT
						EMReadScreen error_message, 75, 24, 2   'checking the bottom for an error message
						error_message = trim(error_message)
					ELSEIF error_message <> "" THEN      'if there is anything here - assume an error
						update_case = FALSE
						action_note = error_message
						PF10
					ELSE
						action_note = "revert complete"
					END IF
				    revert_complete = TRUE
			ELSEIF worker_mail_preference = "RG" and action_taken = "revert" THEN
				    revert_complete = "N/A"
				    action_note = "already reverted " & replace(updated_mony_disb_date, " ", "/")
			END IF
			IF action_note = "update complete" THEN
				start_a_blank_CASE_NOTE
                CALL write_variable_in_CASE_NOTE("MONY/DISB UPDATED " & MAXIS_footer_month &"/"& MAXIS_footer_year)
                CALL write_variable_in_CASE_NOTE("To allow FS cash out cases to be issued PEBT benefits. These   benefits will be issued by DHS in the form of a check and sent to a county office. The county office will then mail checks to the client's payee. After all PEBT benefits are issued, MONY/DISB will be changed back to regular mail. Clients do not need to pick up their benefit check, they should contact their payee for distribution.")
		        CALL write_variable_in_CASE_NOTE("VIA BULK SCRIPT")
     	   	    PF3 'saving the case note
         	    action_note = "update complete & case/note"

				Call navigate_to_MAXIS_screen("SPEC", "WCOM")
				row = 7                             'Defining row and col for the search feature.
				col = 1
				EMSearch "SEND", row, col
				Do 'IF datediff("D", date, todays_date) = 0 THEN ....... = True trying to get the date to readthe date as a dates
					EMReadscreen todays_date, 8, row, 16
					EMReadscreen print_status, 8, row, 71
			 		'If row <> 0 Then PF8
					IF todays_date = "" THEN
				  		print_status = "no notice"
					    EXIT DO
					END IF
					IF todays_date = replace(date, " ", "/") THEN '"02/09/21" change to the current date andthis works perfect'
						EMWriteScreen "C", row, 13
					   	TRANSMIT
					    EMReadscreen second_check, 8, row, 71
					    IF second_check <>  "Canceled"  THEN print_status "REVIEW"
						IF second_check =  "Canceled"  THEN exit do
					ELSE
					   	row = row + 1
				    END IF
				Loop until row = 10
			END IF
		END IF
	amount_cashout = objExcel.cells(excel_row, 2).Value
	objExcel.Cells(excel_row,  3).Value = trim(case_active) 'true/false based on case status
	objExcel.Cells(excel_row,  4).Value = trim(update_case) 	'if case meets criteria to cashout
	objExcel.Cells(excel_row,  5).Value = trim(payment_method) 'payment method
	objExcel.Cells(excel_row,  6).Value = trim(action_note) 'notes or error reason
	objExcel.Cells(excel_row,  7).Value = trim(spec_wcom_canceled) 'spec/wcom status
	objExcel.Cells(excel_row,  8).Value = trim(print_status) 'notes or error reason
	objExcel.Cells(excel_row,  9).Value = trim(revert_complete) 'spec/wcom statusc
	excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1
	back_to_SELF
	action_note = ""
	payment_method = ""
	update_case = ""
	case_active = ""
	spec_wcom_canceled = ""
	revert_complete = ""
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'Loops until there are no more cases in the Excel list


FOR i = 1 to 8		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

STATS_counter = STATS_counter - 1
script_end_procedure("Success! Please review the list generated.")
