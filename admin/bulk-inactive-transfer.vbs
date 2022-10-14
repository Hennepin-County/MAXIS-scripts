'Required for statistical purposes==========================================================================================
name_of_script = "BULK - INACTIVE TRANSFER.vbs"
start_time = timer
STATS_counter = 1                     	'sets the stats counter at one
STATS_manualtime = 229                	'manual run time in seconds
STATS_denomination = "C"       		'C is for each CASE
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
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK=================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
CALL changelog_update("10/14/2022", "Update to merge BULK/INAC script and update file handling.", "MiKayla Handley, Hennepin County") '#916
CALL changelog_update("07/01/2022", "Update to ensure run is complete with error handling.", "MiKayla Handley, Hennepin County") '#868'
CALL changelog_update("02/14/2019", "Initial version.", "MiKayla Handley, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""
Call check_for_MAXIS(end_script)
Call MAXIS_footer_finder(MAXIS_footer_month, MAXIS_footer_year)
MAXIS_footer_month = right("0" &             DatePart("m",           DateAdd("m", -4, date)            ), 4)
MAXIS_footer_year =  right(                  DatePart("yyyy",        DateAdd("m", -4, date)            ), 4)

Do
	Do
		'The dialog is defined in the loop as it can change as buttons are pressed
	    '-------------------------------------------------------------------------------------------------DIALOG
	    Dialog1 = "" 'Blanking out previous dialog detail
        BeginDialog Dialog1, 0, 0, 271, 55, "BULK INAC TRANSFER"
          ButtonGroup ButtonPressed
            PushButton 5, 35, 50, 15, "Instructions", hsr_instruction_button
            OkButton 155, 35, 50, 15
            CancelButton 210, 35, 50, 15
          Text 5, 5, 260, 25, "The script will add cases to an excel sheet (using BULK INAC/REPT) that have been closed for 4 months or more and review the caseload to determine the appropriate transfer actions."
        EndDialog
	  	err_msg = ""
		DIALOG Dialog1
		cancel_confirmation
		If ButtonPressed = select_a_file_button then call file_selection_system_dialog(BULK_INAC_transfer_path, ".xlsx")
 		If ButtonPressed = hsr_instruction_button Then
			err_msg = "LOOP"
			run "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe https://hennepin.sharepoint.com/:w:/r/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/ADMIN/ADMIN%20%E2%80%93%20BULK%20%E2%80%93%20INAC%20Transfer.docx?d=wa6d3e0f66e8940dfa57272414bfd1f76&csf=1&web=1&e=QOcxxB"
		End If
		If err_msg <> "" and err_msg <> "LOOP" Then MsgBox err_msg
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = FALSE					'loops until user passwords back in
'
Call excel_open(BULK_INAC_transfer_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file

'setting the footer month to make the updates in'
back_to_self 'resetting MAXIS back to self before getting started

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_MAXIS(false) 'one more time just in case '

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'This bit freezes the top row of the Excel sheet for better use ability when there is a lot of information
ObjExcel.ActiveSheet.Range("A2").Select
objExcel.ActiveWindow.FreezePanes = True

'Setting the first 3 col as worker, case number, and name
ObjExcel.Cells(1, 1).Value = "WORKER"
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "CASE NAME"
ObjExcel.Cells(1, 4).Value = "APPL DATE"
ObjExcel.Cells(1, 5).Value = "INAC DATE"
ObjExcel.Cells(1, 6).Value = "TRANSFERED"
ObjExcel.Cells(1, 7).Value = "CONFRIM"

'script will go to REPT/USER, and load all of the workers into an array.
CALL create_array_of_all_active_x_numbers_by_supervisor

'Setting the variable for what's to come
excel_row = 2
all_case_numbers_array = "*"

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("rept", "inac")
	EMWriteScreen worker, 21, 16
	EMWriteScreen MAXIS_footer_month, 20, 54 'these dates have been established as current month minus 4'
	EMWriteScreen MAXIS_footer_year, 20, 57
	TRANSMIT

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 10
	If has_content_check <> " " then
		'Grabbing each case number on screen
		Do
			'Set variable for next do...loop
			MAXIS_row = 7
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 3		'Reading case number
				EMReadScreen client_name, 25, MAXIS_row, 14		'Reading client name
				EMReadScreen appl_date, 8, MAXIS_row, 39		'Reading appl date
				EMReadScreen inac_date, 8, MAXIS_row, 49		'Reading inactive date
				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				MAXIS_case_number = trim(MAXIS_case_number)
				If MAXIS_case_number <> "" and instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*")
				If MAXIS_case_number = "" then exit do			'Exits do if we reach the end
				'Adding the case to Excel
				If case_numer <> "        " then
					ObjExcel.Cells(excel_row, 1).Value = worker
					ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
					ObjExcel.Cells(excel_row, 3).Value = client_name
					ObjExcel.Cells(excel_row, 4).Value = appl_date
					ObjExcel.Cells(excel_row, 5).Value = inac_date
					excel_row = excel_row + 1
				End if
				MAXIS_row = MAXIS_row + 1
				add_case_info_to_Excel = ""	'Blanking out variable
				MAXIS_case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
		 	PF8
			EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
next
' this is the end ofcreating the BULK INAC list'


'Now the script adds all the clients on the excel list into an array
transfer_to_worker = "X127CCL" 'setting the worker to the closed basket'
transfer_case_action = TRUE
Do
    previous_worker_number = objExcel.cells(excel_row_to_restart, 1).Value          're-establishing the worker number for functions to use
    If trim(previous_worker_number) = "" then exit do ' this will need to exit'
	IF previous_worker_number = "X127CCL" OR previous_worker_number = "X1274EC" or previous_worker_number = "X127966" or previous_worker_number = "X127AP7" or previous_worker_number = "X127CSS" or previous_worker_number = "X127EF8" or previous_worker_number = "X127EF9" or previous_worker_number = "X127EH9" or previous_worker_number = "X127EJ1" or previous_worker_number = "X127EM2" or previous_worker_number = "X127EM3" or previous_worker_number = "X127EM4" or previous_worker_number = "X127EN6" or previous_worker_number = "X127EN8" or previous_worker_number = "X127EN9" or previous_worker_number = "X127EP1" or previous_worker_number = "X127EP2" or previous_worker_number = "X127EQ6" or previous_worker_number = "X127EQ7" or previous_worker_number = "X127EW4" or previous_worker_number = "X127EW6" or previous_worker_number = "X127EW7" or previous_worker_number = "X127EW8" or previous_worker_number = "X127EX4" or previous_worker_number = "X127EX5" or previous_worker_number = "X127EZ2" or previous_worker_number = "X127F3E" or previous_worker_number = "X127F3F" or previous_worker_number = "X127F3J" or previous_worker_number = "X127F3K" or previous_worker_number = "X127F3N" or previous_worker_number = "X127F3P" or previous_worker_number = "X127F4A" or previous_worker_number = "X127F4B" or previous_worker_number = "X127FE2" or previous_worker_number = "X127FE3" or previous_worker_number = "X127FE6" or previous_worker_number = "X127FF1" or previous_worker_number = "X127FF2" or previous_worker_number = "X127FF4" or previous_worker_number = "X127FF5" or previous_worker_number = "X127FG1" or previous_worker_number = "X127FG2" or previous_worker_number = "X127FG5" or previous_worker_number = "X127FG6" or previous_worker_number = "X127FG7" or previous_worker_number = "X127FG9" or previous_worker_number = "X127FH3" or previous_worker_number = "X127FI1" or previous_worker_number = "X127FI3" or previous_worker_number = "X127FI6" or previous_worker_number = "X127FJ2" or previous_worker_number = "X127GF5" or previous_worker_number = "X127Q95" or previous_worker_number = "X127Y86" or previous_worker_number = "X127EP8" or previous_worker_number = "X127EN5" THEN
		transfer_case_action  = FALSE
		action_completed = "Excluded"
	ELSE
		transfer_case_action = True
		action_completed = "Confirmed"
	END IF

	MAXIS_case_number 	 = objExcel.cells(excel_row_to_restart, 2).Value          're-establishing the case numbers for functions to use
    IF trim(MAXIS_case_number) = "" THEN EXIT DO 'this should end the script'

    IF transfer_case_action = TRUE THEN
	    'go to SPEC/XFER
		CALL navigate_to_MAXIS_screen_review_PRIV("SPEC", "XFER", is_this_priv) ' need discovery on priv cases for xfer handling'
		IF is_this_priv = TRUE THEN
			transfer_case_action = FALSE
			action_completed = "PRIV"
		ELSE
		    EMWriteScreen "X", 7, 16                               'transfer within county option
	        TRANSMIT
	        PF9                                                    'putting the transfer in edit mode
	        EMreadscreen primary_worker, 7, 21, 16                 'how does PW act differently than SW?'
	        EMreadscreen servicing_worker, 7, 18, 65               'checking to see if the transfer_to_worker is the same as the primary_worker (because then it won't transfer)
	        EMreadscreen second_servicing_worker, 7, 18, 74        'checking to see if the transfer_to_worker is the same as the second_servicing_worker (because then it won't transfer)
	        IF second_servicing_worker <> "_______" THEN CALL clear_line_of_text(18, 74)

	        'going for the transfer
	        EMWriteScreen transfer_to_worker, 18, 61           'entering the worker information
	        TRANSMIT                                           'saving - this should then take us to the transfer menu
	        EMReadScreen panel_check, 4, 2, 55                 'reading to see if we made it to the right place
	        If panel_check = "XWKR" THEN                       'this is not the right place
	        	action_completed = "Transfer failed " & panel_check
	        	PF10 'backout
	        	PF3 'SPEC menu
	        	PF3 'SELF Menu'
	        Else                                                 'if we are in the right place - read to see if the new worker is the transfer_to_worker
	        	EMReadScreen primary_worker, 7, 21, 16
	        	If primary_worker <> transfer_to_worker THEN     'if it is not the transfer_to_worker - the transfer failed.
					EMReadScreen MISC_error_check,  74, 24, 02
					action_completed = trim(MISC_error_check)
	        	ELSE
					action_completed = "already in worker " & primary_worker
				END IF
	        END IF
		END IF
	END IF
		'Export data to Excel
		ObjExcel.Cells(excel_row_to_restart, 6).Value = trim(transfer_case_action)
		objExcel.cells(excel_row_to_restart, 7).Value = trim(action_completed)
		STATS_counter = STATS_counter + 1    'adds one instance to the stats counter
		excel_row_to_restart = excel_row_to_restart + 1	'increments the excel row so we don't overwrite our data
LOOP UNTIL previous_worker_number = ""
FOR i = 1 to 7		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Query date/time/runtime info

ObjExcel.Cells(1, 10).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, 10).Value = now
ObjExcel.Cells(2, 10).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, 10).Value = timer - query_start_time

'Autofitting columns
For col_to_autofit = 1 to 10
	ObjExcel.columns(col_to_autofit).AutoFit()
Next
'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! The list is complete. Please review the cases that appear to be in error.")


''----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------07/01/2022
'--Tab orders reviewed & confirmed----------------------------------------------07/01/2022
'--Mandatory fields all present & Reviewed--------------------------------------07/01/2022
'--All variables in dialog match mandatory fields-------------------------------07/01/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------07/01/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------07/01/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------07/01/2022
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------07/01/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------07/01/2022------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------07/01/2022
'--Out-of-County handling reviewed----------------------------------------------07/01/2022------------------N/A
'--script_end_procedures (w/ or w/o error messaging)----------------------------07/01/2022
'--BULK - review output of statistics and run time/count (if applicable)--------07/01/2022------------------N/A
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---07/01/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------07/01/2022------------------N/A
'--Incrementors reviewed (if necessary)-----------------------------------------07/01/2022------------------N/A
'--Denomination reviewed -------------------------------------------------------07/01/2022
'--Script name reviewed---------------------------------------------------------07/01/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------07/01/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------07/01/2022
'--comment Code-----------------------------------------------------------------07/01/2022
'--Update Changelog for release/update------------------------------------------07/01/2022
'--Remove testing message boxes-------------------------------------------------07/01/2022
'--Remove testing code/unnecessary code-----------------------------------------07/01/2022
'--Review/update SharePoint instructions----------------------------------------07/01/2022
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------07/01/2022
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------07/01/2022
'--Complete misc. documentation (if applicable)---------------------------------07/01/2022
'--Update project team/issue contact (if applicable)----------------------------07/01/2022
