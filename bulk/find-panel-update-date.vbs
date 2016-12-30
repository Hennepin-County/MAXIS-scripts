'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - FIND PANEL UPDATE DATE.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 28                      'manual run time in seconds
STATS_denomination = "I"       						 'I is for each ITEM
'END OF stats block==============================================================================================

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

'=====FUNCTIONS=====
FUNCTION build_hh_array(hh_array)
	hh_array = ""
	panel_row = 5
	DO
		EMReadScreen hh_member, 2, panel_row, 3
		IF hh_member <> "  " THEN
			hh_array = hh_array & hh_member & ","
			panel_row = panel_row + 1
		END IF
	LOOP UNTIL hh_member = "  "
	hh_array = trim(hh_array)
	hh_array = split(hh_array, ",")
END FUNCTION

'=====DIALOG=====
BeginDialog panel_update_check_dlg, 0, 0, 226, 255, "Panels to Check"
  EditBox 75, 10, 145, 15, workers_list
  CheckBox 20, 75, 30, 10, "JOBS", JOBS_checkbox
  CheckBox 60, 75, 35, 10, "UNEA", UNEA_checkbox
  CheckBox 100, 75, 30, 10, "BUSI", BUSI_checkbox
  CheckBox 140, 75, 35, 10, "SPON", SPON_checkbox
  CheckBox 180, 75, 30, 10, "RBIC", RBIC_checkbox
  CheckBox 20, 115, 35, 10, "COEX", COEX_checkbox
  CheckBox 60, 115, 30, 10, "DCEX", DCEX_checkbox
  CheckBox 100, 115, 30, 10, "HEST", HEST_checkbox
  CheckBox 140, 115, 30, 10, "SHEL", SHEL_checkbox
  CheckBox 180, 115, 35, 10, "WKEX", WKEX_checkbox
  CheckBox 20, 155, 35, 10, "PACT", PACT_checkbox
  CheckBox 60, 155, 30, 10, "PARE", PARE_checkbox
  CheckBox 100, 155, 30, 10, "PBEN", PBEN_checkbox
  CheckBox 140, 155, 35, 10, "STWK", STWK_checkbox
  CheckBox 180, 155, 35, 10, "WREG", WREG_checkbox
  DropListBox 105, 180, 115, 15, "Select one..."+chr(9)+"Updated in prev. 30 days"+chr(9)+"Updated in prev. 6 mos"+chr(9)+"Not updated more than 12 mos"+chr(9)+"Not updated more than 24 mos", time_period
  CheckBox 10, 205, 210, 10, "Check here to add the supervisor's name to the spreadsheet.", supervisor_check
  ButtonGroup ButtonPressed
    OkButton 115, 225, 50, 15
    CancelButton 170, 225, 50, 15
  Text 10, 185, 90, 10, "Select time period to check:"
  Text 15, 30, 150, 10, "* Please enter only 7-digit worker numbers."
  GroupBox 10, 100, 210, 30, "Expense Panels to Check"
  Text 15, 40, 205, 10, "* For multiple workers, separate worker numbers by a comma."
  GroupBox 10, 60, 210, 30, "Income Panels to Check"
  Text 10, 15, 65, 10, "Worker Number(s):"
  GroupBox 10, 140, 210, 30, "Other panels to check"
EndDialog

'>>>>> THE SCRIPT <<<<<
EMConnect ""

'>>>>> LOADING THE DIALOG <<<<<
DO
	DO
		err_msg = ""
		DIALOG panel_update_check_dlg
		cancel_confirmation
		IF time_period = "Select one..." THEN err_msg = err_msg & vbCr & "* Please select a date range for the script to analyze."

		'Breaking down the workers_list to determine if the user entered multiple workers or if the script is going to be run for just one worker.
		IF InStr(workers_list, ",") <> 0 THEN
		workers_list = replace(workers_list, " ", "")
		workers_list = split(workers_list, ",")
		ELSEIF InStr(workers_list, ",") = 0 THEN
		'multiple_workers = split(workers_list)
		workers_list = split(workers_list)
		END IF

		'>>>>> ADDING TO err_msg IF THE USER SELECTS NO STAT PANELS. <<<<<
		IF JOBS_checkbox = 0 AND _
		UNEA_checkbox = 0 AND _
		BUSI_checkbox = 0 AND _
		RBIC_checkbox = 0 AND _
		SPON_checkbox = 0 AND _
		COEX_checkbox = 0 AND _
		DCEX_checkbox = 0 AND _
		HEST_checkbox = 0 AND _
		SHEL_checkbox = 0 AND _
		WKEX_checkbox = 0 AND _
		PACT_checkbox = 0 AND _
	  	PARE_checkbox = 0 AND _
	  	PBEN_checkbox = 0 AND _
	  	STWK_checkbox = 0 AND _
	  	WREG_checkbox = 0 THEN err_msg = err_msg & vbCr & "* You must select at least one STAT panel to check."

		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	LOOP UNTIL err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until re_we_passworded_out = false					'loops until user passwords back in

'>>>>> EXECTING THE PANEL UPDATE SEARCH FOR EACH WORKER <<<<<
FOR EACH maxis_worker IN workers_list
	IF maxis_worker <> "" THEN
		'>>>>> CREATING A UNIQUE EXCEL FILE <<<<<
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = True
		Set objWorkbook = objExcel.Workbooks.Add()
		objExcel.DisplayAlerts = True

		'>>>>> SETTING EXCEL HEADERS
		objExcel.Cells(1, 1).Value = "X NUMBER"
		objExcel.Cells(1, 2).Value = "CASE NUMBER"
		objExcel.Cells(1, 3).Value = "CLIENT NAME"
		col_to_use = 4
		IF JOBS_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "JOBS"
			JOBS_col = col_to_use
			col_to_use = col_to_use + 1
		END IF
		IF UNEA_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "UNEA"
			UNEA_col = col_to_use
			col_to_use = col_to_use + 1
		END IF
		IF BUSI_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "BUSI"
			BUSI_col = col_to_use
			col_to_use = col_to_use + 1
		END IF
		IF RBIC_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "RBIC"
			RBIC_col = col_to_use
			col_to_use = col_to_use + 1
		END IF
		IF SPON_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "SPON"
			SPON_col = col_to_use
			col_to_use = col_to_use + 1
		END IF
		IF COEX_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "COEX"
			COEX_col = col_to_use
			col_to_use = col_to_use + 1
		END IF
		IF DCEX_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "DCEX"
			DCEX_col = col_to_use
			col_to_use = col_to_use + 1
		END IF
		IF HEST_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "HEST"
			HEST_col = col_to_use
			col_to_use = col_to_use + 1
		END IF
		IF SHEL_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "SHEL"
			SHEL_col = col_to_use
			col_to_use = col_to_use + 1
		END IF
		IF WKEX_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "WKEX"
			WKEX_col = col_to_use
			col_to_use = col_to_use + 1
		END IF
		IF PACT_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "PACT"
			PACT_col = col_to_use
			col_to_use = col_to_use + 1
		END IF
		IF PARE_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "PARE"
			PARE_col = col_to_use
			col_to_use = col_to_use + 1
		END IF
		IF PBEN_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "PBEN"
			PBEN_col = col_to_use
			col_to_use = col_to_use + 1
		END IF
		IF STWK_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "STWK"
			STWK_col = col_to_use
			col_to_use = col_to_use + 1
		END IF
		IF WREG_checkbox = 1 THEN
			objExcel.Cells(1, col_to_use).Value = "WREG"
			WREG_col = col_to_use
			col_to_use = col_to_use + 1
		END IF

		FOR i = 1 TO col_to_use
			objExcel.Cells(1, i).Font.Bold = True
		NEXT

		objExcel.Columns(col_to_use).ColumnWidth = 1
		objExcel.Columns(col_to_use + 1).ColumnWidth = 1
		objExcel.Cells(1, col_to_use + 2).Value = "Time Criteria: "
		objExcel.Cells(1, col_to_use + 3).Value = time_period
		objExcel.Cells(1, col_to_use + 2).Font.Bold = True
		objExcel.Cells(1, col_to_use + 3).Font.Bold = True
		objExcel.Columns(col_to_use + 2).AutoFit()
		objExcel.Columns(col_to_use + 3).AutoFit()

		'>>>>> BUILDING A LIST OF CASE NUMBERS AND CLIENTS <<<<<
		CALL navigate_to_MAXIS_screen("REPT", "ACTV")
		EMReadScreen ACTV_Xnumber, 7, 21, 13
		IF UCase(maxis_worker) <> UCase(ACTV_Xnumber) THEN CALL write_value_and_transmit(maxis_worker, 21, 13)  'if the script transmits on the current worker with their own x# it will skip first page.
		excel_row = 2
		DO
			rept_row = 7
			EMReadScreen last_page, 21, 24, 2
			DO
				EMReadScreen MAXIS_case_number, 8, rept_row, 12
				MAXIS_case_number = trim(MAXIS_case_number)
				EMReadScreen client_name, 20, rept_row, 21
				client_name = trim(client_name)
				IF MAXIS_case_number <> "" THEN
					'>>>>> ADDING WORKER NUMBER, CASE NUMBER, AND CLIENT NAME TO EXCEL <<<<<
					objExcel.Cells(excel_row, 1).Value = maxis_worker
					objExcel.Cells(excel_row, 2).Value = MAXIS_case_number
					objExcel.Cells(excel_row, 3).Value = client_name
					excel_row = excel_row + 1
				END IF
				rept_row = rept_row + 1
			LOOP UNTIL rept_row = 19
			PF8
		LOOP UNTIL last_page = "THIS IS THE LAST PAGE"

		'>>>>> GOING BACK THROUGH THE EXCEL LIST TO SEARCH FOR PANEL UPDATE DATE <<<<<
		excel_row = 2
		DO
			back_to_SELF
			MAXIS_case_number = objExcel.Cells(excel_row, 2).Value
			EMWriteScreen "STAT", 16, 43
			EMWriteScreen "________", 18, 43
			EMWriteScreen MAXIS_case_number, 18, 43
			transmit

			'>>>>> PRIVILEGED CHECK <<<<<
			row = 1
			col = 1
			EMSearch "PRIVILEGED", row, col
			'SELF check protecting against background cases.
			DO
				EMWriteScreen "STAT", 16, 43
				EMWriteScreen "________", 18, 43
				EMWriteScreen MAXIS_case_number, 18, 43
				transmit
				EMReadScreen self_check, 4, 2, 50
				IF row = 24 THEN EXIT DO
			LOOP until self_check <> "SELF"

			IF row <> 24 THEN
				IF JOBS_checkbox = 1 THEN
				STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("JOBS", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "JOBS")
					END IF
					CALL build_hh_array(JOBS_array)
					FOR EACH person IN JOBS_array
						IF person <> "" THEN
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
							CALL write_value_and_transmit(person, 20, 76)
							EMReadScreen updated_date, 8, 21, 55
							updated_date = replace(updated_date, " ", "/")
							IF updated_date <> "////////" THEN
								IF time_period = "Updated in prev. 30 days" THEN
									IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, JOBS_col).Value = objExcel.Cells(excel_row, JOBS_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Updated in prev. 6 mos" THEN
									IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, JOBS_col).Value = objExcel.Cells(excel_row, JOBS_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 12 mos" THEN
									IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, JOBS_col).Value = objExcel.Cells(excel_row, JOBS_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 24 mos" THEN
									IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, JOBS_col).Value = objExcel.Cells(excel_row, JOBS_col).Value & person & ", " & updated_date & "; "
								END IF
							END IF
						END IF
					NEXT
				END IF
				IF UNEA_checkbox = 1 THEN
				STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("UNEA", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "UNEA")
					END IF
					CALL build_hh_array(UNEA_array)
					FOR EACH person IN UNEA_array
						IF person <> "" THEN
						STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
							CALL write_value_and_transmit(person, 20, 76)
							EMReadScreen updated_date, 8, 21, 55
							updated_date = replace(updated_date, " ", "/")
							IF updated_date <> "////////" THEN
								IF time_period = "Updated in prev. 30 days" THEN
									IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, UNEA_col).Value = objExcel.Cells(excel_row, UNEA_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Updated in prev. 6 mos" THEN
									IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, UNEA_col).Value = objExcel.Cells(excel_row, UNEA_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 12 mos" THEN
									IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, UNEA_col).Value = objExcel.Cells(excel_row, UNEA_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 24 mos" THEN
									IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, UNEA_col).Value = objExcel.Cells(excel_row, UNEA_col).Value & person & ", " & updated_date & "; "
								END IF
							END IF
						END IF
					NEXT
				END IF
				IF BUSI_checkbox = 1 THEN
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("BUSI", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "BUSI")
					END IF
					CALL build_hh_array(BUSI_array)
					FOR EACH person IN BUSI_array
						IF person <> "" THEN
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
							CALL write_value_and_transmit(person, 20, 76)
							EMReadScreen updated_date, 8, 21, 55
							updated_date = replace(updated_date, " ", "/")
							IF updated_date <> "////////" THEN
								IF time_period = "Updated in prev. 30 days" THEN
									IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, BUSI_col).Value = objExcel.Cells(excel_row, BUSI_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Updated in prev. 6 mos" THEN
									IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, BUSI_col).Value = objExcel.Cells(excel_row, BUSI_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 12 mos" THEN
									IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, BUSI_col).Value = objExcel.Cells(excel_row, BUSI_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 24 mos" THEN
									IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, BUSI_col).Value = objExcel.Cells(excel_row, BUSI_col).Value & person & ", " & updated_date & "; "
								END IF
							END IF
						END IF
					NEXT
				END IF
				IF RBIC_checkbox = 1 THEN
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("RBIC", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "RBIC")
					END IF
					CALL build_hh_array(RBIC_array)
					FOR EACH person IN RBIC_array
						IF person <> "" THEN
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
							CALL write_value_and_transmit(person, 20, 76)
							EMReadScreen updated_date, 8, 21, 55
							updated_date = replace(updated_date, " ", "/")
							IF updated_date <> "////////" THEN
								IF time_period = "Updated in prev. 30 days" THEN
									IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, RBIC_col).Value = objExcel.Cells(excel_row, RBIC_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Updated in prev. 6 mos" THEN
									IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, RBIC_col).Value = objExcel.Cells(excel_row, RBIC_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 12 mos" THEN
									IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, RBIC_col).Value = objExcel.Cells(excel_row, RBIC_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 24 mos" THEN
									IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, RBIC_col).Value = objExcel.Cells(excel_row, RBIC_col).Value & person & ", " & updated_date & "; "
								END IF
							END IF
						END IF
					NEXT
				END IF
				IF SPON_checkbox = 1 THEN
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("SPON", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "SPON")
					END IF
					CALL build_hh_array(SPON_array)
					FOR EACH person IN SPON_array
						IF person <> "" THEN
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
							CALL write_value_and_transmit(person, 20, 76)
							EMReadScreen updated_date, 8, 21, 55
							updated_date = replace(updated_date, " ", "/")
							IF updated_date <> "////////" THEN
								IF time_period = "Updated in prev. 30 days" THEN
									IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, SPON_col).Value = objExcel.Cells(excel_row, SPON_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Updated in prev. 6 mos" THEN
									IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, SPON_col).Value = objExcel.Cells(excel_row, SPON_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 12 mos" THEN
									IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, SPON_col).Value = objExcel.Cells(excel_row, SPON_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 24 mos" THEN
									IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, SPON_col).Value = objExcel.Cells(excel_row, SPON_col).Value & person & ", " & updated_date & "; "
								END IF
							END IF
						END IF
					NEXT
				END IF
				IF COEX_checkbox = 1 THEN
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("COEX", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "COEX")
					END IF
					CALL build_hh_array(COEX_array)
					FOR EACH person IN COEX_array
						IF person <> "" THEN
						 	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
							CALL write_value_and_transmit(person, 20, 76)
							EMReadScreen updated_date, 8, 21, 55
							updated_date = replace(updated_date, " ", "/")
							IF updated_date <> "////////" THEN
								IF time_period = "Updated in prev. 30 days" THEN
									IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, COEX_col).Value = objExcel.Cells(excel_row, COEX_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Updated in prev. 6 mos" THEN
									IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, COEX_col).Value = objExcel.Cells(excel_row, COEX_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 12 mos" THEN
									IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, COEX_col).Value = objExcel.Cells(excel_row, COEX_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 24 mos" THEN
									IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, COEX_col).Value = objExcel.Cells(excel_row, COEX_col).Value & person & ", " & updated_date & "; "
								END IF
							END IF
						END IF
					NEXT
				END IF
				IF DCEX_checkbox = 1 THEN
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("DCEX", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "DCEX")
					END IF
					CALL build_hh_array(DCEX_array)
					FOR EACH person IN DCEX_array
						IF person <> "" THEN
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
							CALL write_value_and_transmit(person, 20, 76)
							EMReadScreen updated_date, 8, 21, 55
							updated_date = replace(updated_date, " ", "/")
							IF updated_date <> "////////" THEN
								IF time_period = "Updated in prev. 30 days" THEN
									IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, DCEX_col).Value = objExcel.Cells(excel_row, DCEX_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Updated in prev. 6 mos" THEN
									IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, DCEX_col).Value = objExcel.Cells(excel_row, DCEX_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 12 mos" THEN
									IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, DCEX_col).Value = objExcel.Cells(excel_row, DCEX_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 24 mos" THEN
									IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, DCEX_col).Value = objExcel.Cells(excel_row, DCEX_col).Value & person & ", " & updated_date & "; "
								END IF
							END IF
						END IF
					NEXT
				END IF
				IF HEST_checkbox = 1 THEN
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("HEST", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "HEST")
					END IF
					EMReadScreen updated_date, 8, 21, 55
					updated_date = replace(updated_date, " ", "/")
					IF updated_date <> "////////" THEN
						IF time_period = "Updated in prev. 30 days" THEN
							IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, HEST_col).Value = updated_date & "; "
						ELSEIF time_period = "Updated in prev. 6 mos" THEN
							IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, HEST_col).Value = updated_date & "; "
						ELSEIF time_period = "Not updated more than 12 mos" THEN
							IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, HEST_col).Value = updated_date & "; "
						ELSEIF time_period = "Not updated more than 24 mos" THEN
							IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, HEST_col).Value = updated_date & "; "
						END IF
					END IF
				END IF
				IF SHEL_checkbox = 1 THEN
				 	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("SHEL", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "SHEL")
					END IF
					CALL build_hh_array(SHEL_array)
					FOR EACH person IN SHEL_array
						IF person <> "" THEN
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
							CALL write_value_and_transmit(person, 20, 76)
							EMReadScreen updated_date, 8, 21, 55
							updated_date = replace(updated_date, " ", "/")
							IF updated_date <> "////////" THEN
								IF time_period = "Updated in prev. 30 days" THEN
									IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, SHEL_col).Value = objExcel.Cells(excel_row, SHEL_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Updated in prev. 6 mos" THEN
									IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, SHEL_col).Value = objExcel.Cells(excel_row, SHEL_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 12 mos" THEN
									IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, SHEL_col).Value = objExcel.Cells(excel_row, SHEL_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 24 mos" THEN
									IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, SHEL_col).Value = objExcel.Cells(excel_row, SHEL_col).Value & person & ", " & updated_date & "; "
								END IF
							END IF
						END IF
					NEXT
				END IF
				IF WKEX_checkbox = 1 THEN
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("WKEX", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "WKEX")
					END IF
					CALL build_hh_array(WKEX_array)
					FOR EACH person IN WKEX_array
						IF person <> "" THEN
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
							CALL write_value_and_transmit(person, 20, 76)
							EMReadScreen updated_date, 8, 21, 55
							updated_date = replace(updated_date, " ", "/")
							IF updated_date <> "////////" THEN
								IF time_period = "Updated in prev. 30 days" THEN
									IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, WKEX_col).Value = objExcel.Cells(excel_row, WKEX_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Updated in prev. 6 mos" THEN
									IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, WKEX_col).Value = objExcel.Cells(excel_row, WKEX_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 12 mos" THEN
									IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, WKEX_col).Value = objExcel.Cells(excel_row, WKEX_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 24 mos" THEN
									IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, WKEX_col).Value = objExcel.Cells(excel_row, WKEX_col).Value & person & ", " & updated_date & "; "
								END IF
							END IF
						END IF
					NEXT
				END IF
				'PACT is the only case-based panel so the person array has been removed
				IF PACT_checkbox = 1 THEN
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("PACT", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "PACT")
					END IF
					EMReadScreen updated_date, 8, 21, 55
					updated_date = replace(updated_date, " ", "/")
					IF updated_date <> "////////" THEN
						IF time_period = "Updated in prev. 30 days" THEN
							IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, PACT_col).Value = updated_date
						ELSEIF time_period = "Updated in prev. 6 mos" THEN
							IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, PACT_col).Value = updated_date
						ELSEIF time_period = "Not updated more than 12 mos" THEN
							IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, PACT_col).Value = updated_date
						ELSEIF time_period = "Not updated more than 24 mos" THEN
							IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, PACT_col).Value = updated_date
						END IF
					END IF
				END IF
				IF PARE_checkbox = 1 THEN
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("PARE", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "PARE")
					END IF
					CALL build_hh_array(PARE_array)
					FOR EACH person IN PARE_array
						IF person <> "" THEN
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
							CALL write_value_and_transmit(person, 20, 76)
							EMReadScreen updated_date, 8, 21, 55
							updated_date = replace(updated_date, " ", "/")
							IF updated_date <> "////////" THEN
								IF time_period = "Updated in prev. 30 days" THEN
									IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, PARE_col).Value = objExcel.Cells(excel_row, PARE_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Updated in prev. 6 mos" THEN
									IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, PARE_col).Value = objExcel.Cells(excel_row, PARE_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 12 mos" THEN
									IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, PARE_col).Value = objExcel.Cells(excel_row, PARE_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 24 mos" THEN
									IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, PARE_col).Value = objExcel.Cells(excel_row, PARE_col).Value & person & ", " & updated_date & "; "
								END IF
							END IF
						END IF
					NEXT
				END IF
				IF PBEN_checkbox = 1 THEN
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("PBEN", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "PBEN")
					END IF
					CALL build_hh_array(PBEN_array)
					FOR EACH person IN PBEN_array
						IF person <> "" THEN
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
							CALL write_value_and_transmit(person, 20, 76)
							EMReadScreen updated_date, 8, 21, 55
							updated_date = replace(updated_date, " ", "/")
							IF updated_date <> "////////" THEN
								IF time_period = "Updated in prev. 30 days" THEN
									IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, PBEN_col).Value = objExcel.Cells(excel_row, PBEN_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Updated in prev. 6 mos" THEN
									IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, PBEN_col).Value = objExcel.Cells(excel_row, PBEN_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 12 mos" THEN
									IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, PBEN_col).Value = objExcel.Cells(excel_row, PBEN_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 24 mos" THEN
									IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, PBEN_col).Value = objExcel.Cells(excel_row, PBEN_col).Value & person & ", " & updated_date & "; "
								END IF
							END IF
						END IF
					NEXT
				END IF
				IF STWK_checkbox = 1 THEN
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("STWK", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "STWK")
					END IF
					CALL build_hh_array(STWK_array)
					FOR EACH person IN STWK_array
						IF person <> "" THEN
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
							CALL write_value_and_transmit(person, 20, 76)
							EMReadScreen updated_date, 8, 21, 55
							updated_date = replace(updated_date, " ", "/")
							IF updated_date <> "////////" THEN
								IF time_period = "Updated in prev. 30 days" THEN
									IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, STWK_col).Value = objExcel.Cells(excel_row, STWK_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Updated in prev. 6 mos" THEN
									IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, STWK_col).Value = objExcel.Cells(excel_row, STWK_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 12 mos" THEN
									IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, STWK_col).Value = objExcel.Cells(excel_row, STWK_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 24 mos" THEN
									IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, STWK_col).Value = objExcel.Cells(excel_row, STWK_col).Value & person & ", " & updated_date & "; "
								END IF
							END IF
						END IF
					NEXT
				END IF
				IF WREG_checkbox = 1 THEN
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
					EMReadScreen in_stat, 4, 20, 21
					IF in_stat = "STAT" THEN     'prevents error where navigate_to_MAXIS_screen jumps back out for each read
						CALL write_value_and_transmit("WREG", 20, 71)
					ELSE
						CALL navigate_to_MAXIS_screen("STAT", "WREG")
					END IF
					CALL build_hh_array(WREG_array)
					FOR EACH person IN WREG_array
						IF person <> "" THEN
							STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
							CALL write_value_and_transmit(person, 20, 76)
							EMReadScreen updated_date, 8, 21, 55
							updated_date = replace(updated_date, " ", "/")
							IF updated_date <> "////////" THEN
								IF time_period = "Updated in prev. 30 days" THEN
									IF DateDiff("D", updated_date, date) <= 30 THEN objExcel.Cells(excel_row, WREG_col).Value = objExcel.Cells(excel_row, WREG_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Updated in prev. 6 mos" THEN
									IF DateDiff("D", updated_date, date) <= 180 THEN objExcel.Cells(excel_row, WREG_col).Value = objExcel.Cells(excel_row, WREG_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 12 mos" THEN
									IF DateDiff("D", updated_date, date) > 365 THEN objExcel.Cells(excel_row, WREG_col).Value = objExcel.Cells(excel_row, WREG_col).Value & person & ", " & updated_date & "; "
								ELSEIF time_period = "Not updated more than 24 mos" THEN
									IF DateDiff("D", updated_date, date) > 730 THEN objExcel.Cells(excel_row, WREG_col).Value = objExcel.Cells(excel_row, WREG_col).Value & person & ", " & updated_date & "; "
								END IF
							END IF
						END IF
					NEXT
				END IF
			END IF
			excel_row = excel_row + 1
		LOOP UNTIL objExcel.Cells(excel_row, 1).Value = ""
		FOR i = 1 to col_to_use
			objExcel.Columns(i).AutoFit()
		NEXT
	END IF

	IF supervisor_check = 1 THEN
	'Adding additional manual time to the stats counter. I have timed this out to be about 25 seconds per case.
	STATS_manualtime = STATS_manualtime + 25

	'Adding a column to the left of the data
	SET objSheet = objWorkbook.Sheets("Sheet1")
	objSheet.Columns("A:A").Insert -4161
	objExcel.Cells(1, 1).Value = "SUPERVISOR NAME"
	objExcel.Cells(1, 1).Font.Bold = True

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

NEXT
back_to_SELF

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success!")
