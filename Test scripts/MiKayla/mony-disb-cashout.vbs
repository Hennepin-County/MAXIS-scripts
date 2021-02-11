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
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 251, 105, ""
  DropListBox 70, 20, 45, 15, "initial"+chr(9)+"revert", action_taken
  EditBox 210, 20, 15, 15, MAXIS_footer_month
  EditBox 225, 20, 15, 15, MAXIS_footer_year
  ButtonGroup ButtonPressed
    PushButton 15, 40, 50, 15, "Browse...", select_a_file_button
  EditBox 70, 40, 170, 15, file_selection_path
  ButtonGroup ButtonPressed
    OkButton 140, 85, 50, 15
    CancelButton 195, 85, 50, 15
  Text 25, 60, 215, 15, "Select the Excel file that contains the information by selecting the 'Browse' button, and finding the file."
  Text 160, 25, 50, 10, "Footer MMYY:"
  GroupBox 10, 5, 235, 75, "MONY/DISB CASHOUT"
  Text 20, 25, 50, 10, "Action to take:"
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
		If MAXIS_footer_month = "" THEN err_msg = err_msg & vbNewLine & "Please advise the footer year which  you want this script to run."
		If MAXIS_footer_year = "" THEN err_msg = err_msg & vbNewLine & "Please advise the footer year which  you want this script to run."
		If action_taken = "" THEN err_msg = err_msg & vbNewLine & "Please select which option you are taking with this script run."
		If file_selection_path = "" THEN err_msg = err_msg & vbNewLine & "Use the Browse button to select the file that has your data"
		If err_msg <> "" THEN MsgBox err_msg
	Loop until err_msg = ""
	If objExcel = "" THEN call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
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
ObjExcel.Cells(1, 8).Value = "REVERTED"

update_case = FALSE
excel_row = 2           're-establishing the row to start checking the members for

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
	Call navigate_to_MAXIS_screen("CASE", "CURR")
	row = 1                                                 'look for SNAP
    col = 1
    EMSearch "FS:", row, col
    If row <> 0 Then
        EMReadScreen fs_status, 9, row, col + 4
        fs_status = trim(fs_status)
        'fs_status = "ACTIVE" or fs_status = "APP CLOSE" or fs_status = "APP OPEN"  fs_status = "PENDING"
    End If
	If fs_status = "ACTIVE" or fs_status = "APP CLOSE" or fs_status = "APP OPEN" Then
		update_case = TRUE
		case_active = TRUE
	End If
	If fs_status = "PENDING" Then
		update_case = FALSE
		case_active = FALSE
	END If

    Call navigate_to_MAXIS_screen("STAT", "ADDR")
	EMReadScreen priv_check, 4, 2, 50
	IF priv_check = "SELF" THEN
		error_reason = "Privileged"
	ELSE
		EMReadScreen addr_line_01, 22, 6, 43
	    Call navigate_to_MAXIS_screen("STAT", "ALTP")
	    EMReadScreen altp_addr_line_01, 22, 12, 37
	    IF trim(addr_line_01) = trim(altp_addr_line_01) THEN
	    	update_case = FALSE
	       	error_reason = "ADDR same ALTP"
	    ELSE
			update_case = TRUE
			Call navigate_to_MAXIS_screen("MONY", "DISB")
	    	EMReadscreen payment_method, 2, 5, 35
			IF payment_method = "DD" or payment_method = "EB" THEN
				update_case = FALSE
				error_reason = "payment method"
			END IF
			IF update_case = TRUE THEN
				EMReadscreen worker_mail_preference, 2, 9, 35
	    	    IF worker_mail_preference = "RG" and action_taken = "initial" THEN
			    	PF9
			    	EMWriteScreen "IC", 9, 35 'Worker Mail Preference'
					EMWriteScreen "27", 10, 35 'Pick Up County '
					EMWriteScreen "02", 10, 47 'Pick Up Office'
                	TRANSMIT
					EMReadScreen look_for_error, 8, 24, 2   'checking the bottom for an error message
					look_for_error = trim(look_for_error)
					If look_for_error = "WARNING:" Then     'we can transmit past warning messages and then look again
						TRANSMIT
						EMReadScreen really_look_for_error,72, 24, 2   'checking the bottom for an error message
						really_look_for_error = trim(really_look_for_error)
					End If
					If really_look_for_error <> "" Then        'if there is anything here - assume an error
						MsgBox really_look_for_error
					END IF
					error_reason = "transfer back to RG"
                ELSEIF worker_mail_preference = "IC" THEN
						EMReadScreen updated_mony_disb_date, 8, 9, 40
			    		error_reason = "already updated " & replace(updated_mony_disb_date, " ", "/")
				ELSEIF action_taken = "revert" THEN
					   	EMReadscreen worker_mail_preference, 2, 9, 35
				    	IF worker_mail_preference = "IC" THEN
				    		PF9
				    		EMWriteScreen "RG", 9, 35
				    		TRANSMIT
				    		EMReadScreen look_for_error, 8, 24, 2   'checking the bottom for an error message
				    		look_for_error = trim(look_for_error)
				    		If look_for_error = "WARNING:" THEN'we can transmit past warning messages and then look     gain
				    			TRANSMIT
				    			revert_complete = TRUE
				    			EMReadScreen really_look_for_error,72, 24, 2   'checking the bottom for an error message
				    			really_look_for_error = trim(really_look_for_error)
				    		End If
				    		If really_look_for_error <> "" Then        'if there is anything here - assume an error
				    			MsgBox really_look_for_error
				    			revert_complete = FALSE
				    			'exit do
				    		END IF
				    		revert_complete = TRUE
				    	ELSE
				    		IF worker_mail_preference = "IC" THEN
				    			EMReadScreen updated_mony_disb_date, 8, 9, 40
				    			error_reason = "already updated " & replace(updated_mony_disb_date, " ", "/")
				    		END IF
							IF worker_mail_preference = "RG" THEN
								error_reason = "COMPLETE"
				    			revert_complete = TRUE
							END IF
				    	END IF
				    END IF
				end if
				IF error_reason = "transfer back to RG" THEN
					start_a_blank_CASE_NOTE
                    CALL write_variable_in_CASE_NOTE("MONY/DISB UPDATED " & MAXIS_footer_month &"/"& MAXIS_footer_year)
                    CALL write_variable_in_CASE_NOTE("To allow FS cash out cases to be issued PEBT benefits. These benefits will be issued by DHS in the form 'of a check and sent to a county office. The county office will then mail checks to the clients payee. After all PEBT benefits are issued, 'MONY/DISB will be changed back to regular mail. Clients do not need to pick up their benefit check, they should contact their payee for 'distribution.")
				    CALL write_variable_in_CASE_NOTE("VIA BULK SCRIPT")
     	   	        PF3 'saving the case note
         	        error_reason = "Case/note updated"
				END IF

			    IF error_reason = "Case/note updated" or revert_complete = TRUE THEN
			    	Call navigate_to_MAXIS_screen("SPEC", "WCOM")
			    	MAXIS_row = 7
			    	Do
			    		EMReadscreen todays_date, 8, MAXIS_row, 16
			    		EMReadscreen print_status, 8, MAXIS_row, 71
			    		EmReadscreen doc_description, 4, MAXIS_row, 30
			    		EmReadscreen prog_type, 2, MAXIS_row, 26
			    		IF todays_date = "" THEN
			    			EMWaitReady
			    			TRANSMIT
			    			spec_wcom_canceled = FALSE & "NO DATE"
			    		END IF
			    		IF todays_date = "12/17/20" THEN 'if i use 12/16/20 it works but even getting it to recognize it is a date failed
			    			IF print_status = "Canceled" THEN spec_wcom_canceled = FALSE
			    		  	IF print_status = "Waiting" THEN
			    				If doc_description	= "SEND" THEN
			    					IF prog_type = "FS" THEN
			    						EMWriteScreen "C", MAXIS_row, 13
			    						TRANSMIT
			    						spec_wcom_canceled = TRUE
			    					END IF
								END IF
							END IF
			    		ELSE
			    			MAXIS_row = MAXIS_row + 1
			    		END IF
			    	Loop until MAXIS_row = 10
				END IF
			END IF
		END IF
	'END IF
'END IF
	amount_cashout = objExcel.cells(excel_row, 2).Value
	objExcel.Cells(excel_row,  3).Value = trim(case_active) 'true/false based on case status
	objExcel.Cells(excel_row,  4).Value = trim(update_case) 	'if case meets criteria to cashout
	objExcel.Cells(excel_row,  5).Value = trim(payment_method) 'payment method
	objExcel.Cells(excel_row,  6).Value = trim(error_reason) 'notes or error reason
	objExcel.Cells(excel_row,  7).Value = trim(spec_wcom_canceled) 'spec/wcom status
	'objExcel.Cells(excel_row,  8).Value = trim(revert_complete) 'spec/wcom statusc
	excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1
	back_to_SELF
	error_reason = ""
	payment_method = ""
	update_case = ""
	case_active = ""
	spec_wcom_canceled = ""
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'Loops until there are no more cases in the Excel list


FOR i = 1 to 8		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

STATS_counter = STATS_counter - 1
script_end_procedure("Success! Please review the list generated.")
