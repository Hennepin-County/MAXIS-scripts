'Gathering stats
name_of_script = "BULK - MAGI NON MAGI REPORT.vbs"
start_time = timer
stats_counter = 1
stats_manualtime = 67
STATS_denomination = "C"

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'Connecting
EMConnect ""
CALL check_for_MAXIS(true)


'The dialog
BeginDialog Dialog1, 0, 0, 261, 135, "MAGI/Non-MAGI Check"
  Text 10, 15, 115, 10, "Please enter X1 numbers to check"
  EditBox 135, 10, 120, 15, x1_numbers
  Text 10, 30, 245, 10, "To check multiple caseloads, enter the X1 numbers separated by commas."
  CheckBox 10, 50, 240, 10, "Check here to include info about other active programs.", other_progs_check
  CheckBox 10, 65, 245, 10, "Check here to check all cases in your agencies.", all_workers_check
  CheckBox 10, 80, 245, 10, "Check here to add the worker's supervisor added to the report.", supervisor_check
  ButtonGroup ButtonPressed
  	OkButton 155, 115, 50, 15
  	CancelButton 205, 115, 50, 15
EndDialog

DIALOG
cancel_confirmation

'Double checking that MAXIS is not timed out
CALL check_for_MAXIS(false)

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the column headers
objExcel.Cells(1, 1).Value = "X Number"
objExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 3).Value = "NAME"
objExcel.Cells(1, 4).Value = "MAGI Persons"
objExcel.Cells(1, 5).Value = "Non-MAGI Persons"
objExcel.Cells(1, 6).Value = "# of MAGI"
objExcel.Cells(1, 7).Value = "# of Non-MAGI"
objExcel.Cells(1, 8).Value = "MAGI Household"
objExcel.Cells(1, 9).Value = "Mixed Household"
objExcel.Cells(1, 10).Value = "Non-MAGI Household"
objExcel.Cells(1, 11).Value = "MAGI Review aligned?"
objExcel.Cells(1, 12).Value = "HC ER MONTH"
'Here's your option to add other active programs
IF other_progs_check = 1 THEN
	objExcel.Cells(1, 13).Value = "CASH PROG"
	objExcel.Cells(1, 13).Font.Bold = TRUE
	objExcel.Cells(1, 14).Value = "SNAP ACTV"
	objExcel.Cells(1, 14).Font.Bold = TRUE
END IF

'And now BOLD because format
FOR i = 1 TO 12
	objExcel.Cells(1, i).Font.Bold = TRUE
NEXT

'Determining the list of workers to check
IF all_workers_check = 1 THEN
	CALL create_array_of_all_active_x_numbers_in_county(all_workers_array, right(worker_county_code, 2))
ELSE
	'If the user entered multiple users, the script will look for commas and separate
	IF InStr(x1_numbers, ",") = 0 THEN
		all_workers_array = split(x1_numbers)
	ELSE
		x1_numbers = replace(x1_numbers, " ", "")
		all_workers_array = split(x1_numbers, ",")
	END IF
END IF

'Grabbing case numbers
excel_row = 2
FOR EACH worker IN all_workers_array
	CALL navigate_to_MAXIS_screen("REPT", "ACTV")
	CALL write_value_and_transmit(worker, 21, 13)
	'Returning to first page of REPT/ACTV
	PF7

	'Top of REPT/ACTV
	rept_actv_row = 7
	DO
		'Reading for the end of the page
		EMReadScreen this_is_the_last_page, 21, 24, 2
		'Reading for hc status
		EMReadScreen hc_status, 1, rept_actv_row, 64
		'If the client's HC is active...
		IF hc_status = "A" THEN
			'Writing the worker X1 number...
			objExcel.Cells(excel_row, 1).Value = UCASE(worker)
			'Reading and writing case number
			EMReadScreen MAXIS_case_number, 8, rept_actv_row, 12
			objExcel.Cells(excel_row, 2).Value = trim(MAXIS_case_number)
			'Reading and writing client name
			EMReadScreen client_name, 20, rept_actv_row, 21
			objExcel.Cells(excel_row, 3).Value = trim(client_name)
			'If the user requested to see info about CASH and SNAP on the report...
			IF other_progs_check = 1 THEN
				EMReadScreen cash_one_status, 1, rept_actv_row, 54
				EMReadScreen cash_one_prog, 2, rept_actv_row, 51
				EMReadScreen cash_two_status, 1, rept_actv_row, 59
				EMReadScreen cash_two_prog, 2, rept_actv_row, 56
				EMReadScreen snap_status, 1, rept_actv_row, 61
				IF cash_one_status = "A" THEN objExcel.Cells(excel_row, 13).Value = cash_one_prog
				IF cash_two_status = "A" THEN objExcel.Cells(excel_row, 13).Value = cash_two_prog
				IF snap_status = "A" THEN objExcel.Cells(excel_row, 14).Value = "Y"
			END IF
			'Going to the next Excel row
			excel_row = excel_row + 1
			stats_counter = stats_counter + 1
		END IF
		'Going to the next REPT/ACTV row
		rept_actv_row = rept_actv_row + 1
		'If the script gets to row 19...
		IF rept_actv_row = 19 THEN
			'...go to the next page...
			PF8
			'...and reset REPT/ACTV row...
			rept_actv_row = 7
			'...and if the script is on the last page on REPT/ACTV, it exits the do/loop
			'   and starts on the next worker.
			IF this_is_the_last_page = "THIS IS THE LAST PAGE" THEN EXIT DO
		END IF
	LOOP
NEXT

'Going back through and determining MAGI & Non-MAGI
excel_row = 2
Do
	'Assigning a value to MAXIS_case_number
	MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
	'When the script gets to the end of the list, it exits the do/loop
	If MAXIS_case_number = "" then exit do

	'Reseting critical values
	MAGI_count = 0
	nonMAGI_count = 0
	magi_clients = ""
	non_magi_clients = ""

	'setting the row to read on ELIG/HC
	hhmm_row = 8
	DO
		'navigating to ELIG/HC
		CALL navigate_to_MAXIS_screen("ELIG", "HC")
		'reading the hc reference number
		EMReadScreen hc_ref_num, 2, hhmm_row, 3
		'looking to see that information is found for that client
		EMReadScreen hc_information_found, 70, hhmm_row, 3
		hc_information_found = trim(hc_information_found)
		EMReadScreen elig_result, 4, hhmm_row, 41
		EMReadScreen elig_status, 6, hhmm_row, 50
		'...and if information is found for that row...
		IF hc_information_found <> "" THEN
			'...if the client is eligible and active...
			IF elig_result = "ELIG" AND elig_status = "ACTIVE" THEN
				'looking for the first character on hc request...
				EMReadScreen hc_requested, 1, hhmm_row, 28
				'...if the client is active on a medicare savings program...
				IF hc_requested = "S" OR hc_requested = "Q" OR hc_requested = "I" THEN 			'IF the HH MEMB is MSP ONLY then they are automatically Budg Mthd B
					IF hc_ref_num = "  " THEN
						temp_hhmm_row = hhmm_row
						DO
							EMReadScreen hc_ref_num, 2, temp_hhmm_row, 3
							IF hc_ref_num = "  " THEN
								temp_hhmm_row = temp_hhmm_row - 1
							ELSE
								EXIT DO
							END IF
						LOOP
					END IF
					IF InStr(non_magi_clients, hc_ref_num & ";") = 0 THEN non_magi_clients = non_magi_clients & hc_ref_num & ";"
					hhmm_row = hhmm_row + 1
				'...otherwise, if the client is active on Medicaid or EMA...
				ELSEIF hc_requested = "M" or hc_requested = "E" THEN
					'...going in to grab the budget method...
					EMWriteScreen "X", hhmm_row, 26
					transmit
					EMReadScreen budg_mthd, 1, 13, 76
					IF budg_mthd = "A" THEN
						IF hc_ref_num = "  " THEN
							temp_hhmm_row = hhmm_row
							DO
								EMReadScreen hc_ref_num, 2, temp_hhmm_row, 3
								IF hc_ref_num = "  " THEN
									temp_hhmm_row = temp_hhmm_row - 1
								ELSE
									EXIT DO
								END IF
							LOOP
						END IF
						IF InStr(magi_clients, hc_ref_num & ";") = 0 THEN magi_clients = magi_clients & hc_ref_num & ";"
					ELSE
						IF hc_ref_num = "  " THEN
							temp_hhmm_row = hhmm_row
							DO
								EMReadScreen hc_ref_num, 2, temp_hhmm_row, 3
								IF hc_ref_num = "  " THEN
									temp_hhmm_row = temp_hhmm_row - 1
								ELSE
									EXIT DO
								END IF
							LOOP
						END IF
						IF InStr(non_magi_clients, hc_ref_num & ";") = 0 THEN non_magi_clients = non_magi_clients & hc_ref_num & ";"
					END IF
					PF3
					hhmm_row = hhmm_row + 1
				ELSEIF hc_requested = "N" THEN
					hhmm_row = hhmm_row + 1
				END IF
			ELSE
				hhmm_row = hhmm_row + 1
			END IF
		ELSE
			EXIT DO
		END IF
	LOOP UNTIL hhmm_row = 20 OR hc_ref_num = "  "

	'Going back to determine if the individual is still MAGI...SSI, Medicare, and MA-EPD disq the person as Non-MAGI
	IF magi_clients <> "" THEN
		magi_peeps = replace(magi_clients & "~~~", ";~~~", "")
		magi_peeps = split(magi_peeps, ";")
		FOR EACH client IN magi_peeps
			CALL navigate_to_MAXIS_screen("STAT", "MEMB")
			CALL write_value_and_transmit(client, 20, 76)
			EMReadScreen client_age, 2, 8, 76
			IF client_age = "  " THEN client_age = 0
			'Removing that client from the non-magi list
			IF client_age >= 65 THEN
				magi_clients = replace(magi_clients, client & ";", "")
				IF InStr(non_magi_clients, client & ";") = 0 THEN non_magi_clients = non_magi_clients & client & ";"
			ELSE
				'Checking DISA for a MA-EPD & SSI
				CALL navigate_to_MAXIS_screen("STAT", "DISA")
				CALL write_value_and_transmit(client, 20, 76)
				EMReadScreen hc_disa_status, 2, 13, 59
				IF hc_disa_status = "03" OR hc_disa_status = "04" OR hc_disa_status = "22" THEN
					magi_clients = replace(magi_clients, client & ";", "")
					IF InStr(non_magi_clients, client & ";") = 0 THEN non_magi_clients = non_magi_clients & client & ";"
				ELSE
					CALL navigate_to_MAXIS_screen("STAT", "MEDI")
					CALL write_value_and_transmit(client, 20, 76)
					EMReadScreen medi_ref_num, 15, 6, 44
					medi_ref_num = trim(replace(medi_ref_num, "_", ""))
					IF medi_ref_num <> "" THEN
						magi_clients = replace(magi_clients, client & ";", "")
						IF InStr(non_magi_clients, client & ";") = 0 THEN non_magi_clients = non_magi_clients & client & ";"
					END IF
				END IF
			END IF
		NEXT
	END IF

	'Going back to determine if the individual is still Non-MAGI...SSI, Medicare, and MA-EPD disq the person as Non-MAGI
	IF non_magi_clients <> "" THEN
		non_magi_peeps = replace(non_magi_clients & "~~~", ";~~~", "")
		non_magi_peeps = split(non_magi_peeps, ";")
		FOR EACH client IN non_magi_peeps
			'Checking for SSI, MA-EPD, and Medicare
			non_magi = ""
			CALL navigate_to_MAXIS_screen("STAT", "MEDI")
			CALL write_value_and_transmit(client, 20, 76)
			EMReadScreen medi_case_number, 15, 6, 44
			medi_case_number = trim(replace(medi_case_number, "_", ""))
			IF medi_case_number <> "" THEN non_magi = TRUE
			IF non_magi <> TRUE THEN
				CALL navigate_to_MAXIS_screen("STAT", "DISA")
				CALL write_value_and_transmit(client, 20, 76)
				EMReadScreen hc_disa_status, 2, 13, 59
				IF hc_disa_status = "03" OR hc_disa_status = "04" OR hc_disa_status = "22" THEN non_magi = TRUE
			END IF
			IF non_magi <> TRUE THEN
				non_magi_clients = replace(non_magi_clients, client & ";", "")
				IF InStr(magi_clients, client & ";") = 0 THEN magi_clients = magi_clients & client & ";"
			END IF
		NEXT
	END IF

	'Writing all these ding-dang values to Excel
	objExcel.Cells(excel_row, 4).Value = magi_clients
	IF magi_clients <> "" THEN
		MAGI_count = UBound(split(magi_clients, ";"))
		objExcel.Cells(excel_row, 6).Value = MAGI_count
	ELSE
		objExcel.Cells(excel_row, 6).Value = 0
	END IF

	objExcel.Cells(excel_row, 5).Value = non_magi_clients
	IF non_magi_clients <> "" THEN
		nonMAGI_count = UBound(split(non_magi_clients, ";"))
		objExcel.Cells(excel_row, 7).Value = nonMAGI_count
	ELSE
		objExcel.Cells(excel_row, 7).Value = 0
	END IF

	CALL navigate_to_MAXIS_screen("STAT", "REVW")
	EMReadScreen revw_does_not_exist, 19, 24, 2
	IF revw_does_not_exist <> "REVW DOES NOT EXIST" THEN
		EMwritescreen "X", 5, 71
		Transmit
		'Checking to make sure pop up opened
		DO
			EMReadScreen revw_pop_up_check, 8, 4, 44
			EMWaitReady 1, 1
		LOOP until revw_pop_up_check = "RENEWALS"
		'Reading HC reviews to compare them
		EMReadScreen hc_income_renewal, 8, 8, 27
		EMReadScreen hc_IA_renewal, 8, 8, 71
		EMReadScreen hc_annual_renewal, 8, 9, 27
		objExcel.Cells(excel_row, 12).Value = replace(hc_annual_renewal, " ", "/")
		IF MAGI_count <> 0 THEN
			IF hc_income_renewal = "__ 01 __" THEN hc_compare_renewal = hc_IA_renewal
			IF hc_IA_renewal = "__ 01 __" THEN hc_compare_renewal = hc_income_renewal

			IF hc_annual_renewal = hc_compare_renewal THEN
				objExcel.Cells(excel_row, 11).Value = "Y"
			ELSE
				objExcel.Cells(excel_row, 11).Value = "Y"
			END IF
		END IF
	ELSE
		objExcel.Cells(excel_row, 12).Value = "NO REVIEW DATE"
	END IF

	IF MAGI_count <> 0 AND nonMAGI_count = 0 THEN
		objExcel.Cells(excel_row, 8).Value = "Y"
	ELSEIF MAGI_count <> 0 AND nonMAGI_count <> 0 THEN
		objExcel.Cells(excel_row, 9).Value = "Y"
	ELSEIF MAGI_count = 0 AND nonMAGI_count <> 0 THEN
		objExcel.Cells(excel_row, 10).Value = "Y"
	END IF

	excel_row = excel_row + 1
LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""

FOR i = 1 TO 14
	objExcel.Columns(i).AutoFit()
NEXT

IF supervisor_check = 1 THEN
  STATS_manualtime = STATS_manualtime + 25

	'Adding a column to the left of the data
	SET objSheet = objWorkbook.Sheets("Sheet1")
	objSheet.Columns("A:A").Insert -4161
	objExcel.Cells(1, 1).Value = "SUPERVISOR NAME"

	'Going to REPT/USER
	CALL navigate_to_MAXIS_screen("REPT", "USER")

	'Starting back at the top of the page
	excel_row = 2
	DO
		worker_id = objExcel.Cells(excel_row, 2).Value
		prev_worker_id = objExcel.Cells(excel_row - 1, 2).Value
		IF worker_id <> prev_worker_id THEN
			'Entering the worker number into REPT/USER
			CALL write_value_and_transmit(worker_id, 21, 12)
			CALL write_value_and_transmit("X", 7, 3)
			'Grabbing the supervisor X1 number
			EMReadScreen supervisor_id, 7, 14, 61
			transmit
			CALL write_value_and_transmit(supervisor_id, 21, 12)
			EMReadScreen supervisor_name, 18, 7, 14
			supervisor_name = trim(supervisor_name)
			objExcel.Cells(excel_row, 1).Value = supervisor_name
		ELSE
			'Adding the supervisor name from the previous row if the X1 number on this row matches the X1 number on the previous row
			objExcel.Cells(excel_row, 1).Value = objExcel.Cells(excel_row - 1, 1).Value
		END IF
		excel_row = excel_row + 1
	LOOP UNTIL objExcel.Cells(excel_row, 2).Value = ""

	objExcel.Columns(1).AutoFit()
END IF

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success!" & vbCr & vbCr & "The script has finished running.")
