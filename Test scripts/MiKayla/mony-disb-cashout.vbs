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
BeginDialog Dialog1, 0, 0, 266, 95, "MONY/DISB CASHOUT"
ButtonGroup ButtonPressed
	PushButton 200, 25, 50, 15, "Browse...", select_a_file_button
	OkButton 145, 75, 50, 15
CancelButton 200, 75, 50, 15
EditBox 15, 25, 180, 15, file_selection_path
GroupBox 10, 5, 250, 65, "MONY/DISB CASHOUT"
Text 20, 45, 170, 20, "Select the Excel file that contains the information by selecting the 'Browse' button, and finding the file."
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
		If file_selection_path = "" THEN err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
		If err_msg <> "" THEN MsgBox err_msg
	Loop until err_msg = ""
	If objExcel = "" THEN call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
	If err_msg <> "" THEN MsgBox err_msg
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in


CALL check_for_MAXIS(False)
back_to_SELF
ObjExcel.Cells(1, 1).Value = "CASE NUMBER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "AMOUNT"
objExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(1, 3).Value = "FS STATUS"
ObjExcel.Cells(1, 3).Font.Bold = TRUE
ObjExcel.Cells(1, 4).Value = "UPDATE MADE"
ObjExcel.Cells(1, 4).Font.Bold = TRUE
ObjExcel.Cells(1, 5).Value = "METHOD"
ObjExcel.Cells(1, 5).Font.Bold = TRUE
ObjExcel.Cells(1, 6).Value = "NOTES"
ObjExcel.Cells(1, 6).Font.Bold = TRUE

update_case = FALSE
excel_row = 2           're-establishing the row to start checking the members for

Do
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
	If priv_check = "SELF" then
		error_reason = "Privileged"
	Else
	    EMReadScreen addr_line_01, 22, 6, 43
	    Call navigate_to_MAXIS_screen("STAT", "ALTP")
	    EMReadScreen altp_addr_line_01, 22, 12, 37

	    IF trim(addr_line_01) = trim(altp_addr_line_01) THEN
	    	update_case = FALSE
	    	'MsgBox addr_line_01 & " " & altp_addr_line_01
	    	error_reason = "ADDR same ALTP"
	    ELSE
	    	Call navigate_to_MAXIS_screen("MONY", "DISB")
	    	EMReadscreen payment_method, 2, 5, 35
			IF payment_method = "DD" or payment_method = "EB" THEN update_case = FALSE

			EMReadscreen worker_mail_preference, 2, 9, 35
	    	IF worker_mail_preference = "IC" THEN
				PF9
				EMWriteScreen "RG", 9, 35
            	TRANSMIT
            	error_reason = "transferred back to RG"
            ELSE
				error_reason = "none"
			END IF

	    END IF
	    	'start_a_blank_CASE_NOTE
            'CALL write_variable_in_CASE_NOTE("MONY/DISB UPDATED")
            'CALL write_variable_in_CASE_NOTE("To allow FS cash out cases to be issued PEBT benefits.  These benefits will be     'issued  by DHS in the form of a check and sent to a county office.  The county office will then mail checks to     the 'clients payee.  After all PEBT benefits are issued, MONY/DISB will be changed back to regular mail.  Clients     do 'not need to pick up their benefit check, they should contact their payee for distribution")
	    	'CALL write_variable_in_CASE_NOTE("VIA BULK SCRIPT")
	    	'PF3 'saving the case note
        	'error_reason = "Case/note updated"

	    	amount_cashout = objExcel.cells(excel_row, 2).Value
	    	objExcel.Cells(excel_row,  3).Value = trim(case_active) 'true/false based on case status
	    	objExcel.Cells(excel_row,  4).Value = trim(update_case) 	'if case meets criteria to cashout
	    	objExcel.Cells(excel_row,  5).Value = trim(payment_method) 'payment method
            objExcel.Cells(excel_row,  6).Value = trim(error_reason) 'notes or error reason
            excel_row = excel_row + 1
            STATS_counter = STATS_counter + 1
            back_to_SELF
			error_reason = ""
			payment_method = ""
			update_case = ""
			case_active = ""
	END IF
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""	'Loops until there are no more cases in the Excel list

FOR i = 1 to 6							'making the columns stretch to fit the widest cell
	objExcel.Columns(i).AutoFit()
NEXT

STATS_counter = STATS_counter - 1
script_end_procedure("Success! Please review the list generated.")
