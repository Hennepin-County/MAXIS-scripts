'Hard coding that needs to be updated each year: MAXIS_footer_year, counted_date_year 

'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - BANKED MONTHS REPORT.vbs"
start_time = timer
STATS_counter = 1              'sets the stats counter at one
STATS_manualtime = 219         'manual run time in seconds
STATS_denomination = "C"       'C is for each CASE
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

EMConnect ""		'connecting to MAXIS
Call get_county_code	'gets county name to input into the 1st col of the spreadsheet
'developer_mode = TRUE 	'defauting the person note option to NOT person note

county_name = left(county_name, len(county_name)-7)

'Runs the dialog'
Do
	Do
		Do
			'The dialog is defined in the loop as it can change as buttons are pressed (populating the dropdown)'
			BeginDialog dfln_selection_dialog, 0, 0, 266, 75, "Banked Month Report"
			  EditBox 15, 25, 190, 15, dfln_list_excel_file_path
			  ButtonGroup ButtonPressed
			    PushButton 215, 25, 45, 15, "Browse...", select_a_file_button
			  DropListBox 15, 55, 140, 15, "select one..." & sheet_list, worksheet_dropdown
			  ButtonGroup ButtonPressed
			    OkButton 175, 55, 40, 15
			    CancelButton 220, 55, 40, 15
			  Text 10, 10, 255, 10, "Select the Excel File that DHS provided with the list of Convicted Drug Felons."
			  Text 10, 45, 150, 10, "Select the correct worksheet in the Excel file:"
			EndDialog
			err_msg = ""
			Dialog dfln_selection_dialog
			cancel_confirmation
			If ButtonPressed = select_a_file_button then 
				If dfln_list_excel_file_path <> "" then 'This is handling for if the BROWSE button is pushed more than once'
					objExcel.Quit 'Closing the Excel file that was opened on the first push'
					objExcel = "" 	'Blanks out the previous file path'
					sheet_list = ""	'Blanks the Month list out so that the previous worksheets are not still included'
				End If 
				call file_selection_system_dialog(dfln_list_excel_file_path, ".xlsx") 'allows the user to select the file'
			End If 
			If dfln_list_excel_file_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
			If err_msg <> "" Then MsgBox err_msg
		Loop until err_msg = ""
		If objExcel = "" Then call excel_open(dfln_list_excel_file_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
		sheet_list = ""
		For Each objWorkSheet In objWorkbook.Worksheets
			sheet_list = sheet_list & chr(9) & objWorkSheet.Name
		Next
		If worksheet_dropdown = "select one..." then err_msg = err_msg & vbNewLine & "You must select a month that you are running this script for."
		If err_msg <> "" Then MsgBox err_msg
	Loop until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in					

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

objExcel.worksheets(worksheet_dropdown).Activate			'Activates the selected worksheet'

Const Col_Name = 0
Const Col_Numb = 1

excel_col = 1
array_counter = 0
Dim col_name_array
ReDim col_name_array(1, 0)
Do
	ReDim Preserve col_name_array(1, array_counter)
	col_name_array(Col_Name, array_counter) = ucase(replace(objExcel.cells(1, excel_col).Value, " ", ""))
	col_name_array(Col_Numb, array_counter) = excel_col
	excel_col = excel_col + 1
	array_counter = array_counter + 1
	end_of_list = objExcel.cells(1, excel_col).Value
Loop until end_of_list = ""

Const row_numb     = 0
Const month_reptd  = 1
Const case_numb    = 2
Const pers_id      = 3
Const cty_court_01 = 4
Const sent_dt_01   = 5
Const addr1_01     = 6
Const addr2_01     = 7
Const city_01      = 8
Const state_01     = 9
Const zip_01       = 10

	Const cty_court_02 = 11
	Const sent_dt_02 = 12
	Const addr1_02 = 13
	Const addr2_02 = 14
	Const city_02 = 15
	Const state_02 = 16
	Const zip_02 = 17

	Const cty_court_03 = 18
	Const sent_dt_03 = 19
	Const addr1_03 = 20
	Const addr2_03 = 21
	Const city_03 = 22
	Const state_03 = 23
	Const zip_03 = 24

	Const cty_court_04 = 25
	Const sent_dt_04 = 26
	Const addr1_04 = 27
	Const addr2_04 = 28
	Const city_04 = 29
	Const state_04 = 30
	Const zip_04 = 31

	Const cty_court_05 = 32
	Const sent_dt_05 = 33
	Const addr1_05 = 34
	Const addr2_05 = 35
	Const city_05 = 36
	Const state_05 = 37
	Const zip_05 = 38

	Const cty_court_06 = 39
	Const sent_dt_06 = 40
	Const addr1_06 = 41
	Const addr2_06 = 42
	Const city_06 = 43
	Const state_06 = 44
	Const zip_06 = 45

Const clt_name = 46
Const ref_numb = 47
Const case_pop = 48
Const actv_cty = 49
Const cash_prog = 50
Const worker_nbr = 51
Const superv_nbr = 52

Const stat_addr = 53

Dim dfln_to_process_array
ReDim dfln_to_process_array(53, 0)

array_counter = 0
excel_row = 2

Do 
	For column = 0 to UBound(col_name_array, 2)
		If col_name_array(Col_Name, column) = "COUNTIES" Then 
			If UCase(objExcel.cells(excel_row, col_name_array(Col_Numb, column)).Value) = UCase(county_name) Then
				ReDim Preserve dfln_to_process_array(53, array_counter)
				dfln_to_process_array(row_numb, array_counter) = excel_row
				array_counter = array_counter + 1
			End If 
			Exit For
		End If 
	Next
	excel_row = excel_row + 1
	end_of_list = objExcel.cells(excel_row, 2).Value
Loop until end_of_list = ""

For person = 0 to Ubound(dfln_to_process_array, 2)
	For column = 0 to UBound(col_name_array, 2)
		If col_name_array(Col_Name, column) = "REPORT_MONTH"             Then dfln_to_process_array(month_reptd, person)  = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)	
		If col_name_array(Col_Name, column) = "CASENUMBER"               Then dfln_to_process_array(case_numb, person)    = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "PERSONID"                 Then dfln_to_process_array(pers_id, person)      = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)

		If col_name_array(Col_Name, column) = "COUNTYCOURTDESCRIPTION01" Then dfln_to_process_array(cty_court_01, person) = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "SENTENCEDATE01"           Then dfln_to_process_array(sent_dt_01, person)   = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ADDRESS1_01"              Then dfln_to_process_array(addr1_01, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ADDRESS2_01"              Then dfln_to_process_array(addr2_01, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "CITY01"                   Then dfln_to_process_array(city_01, person)      = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "STATE01"                  Then dfln_to_process_array(state_01, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ZIP01"                    Then dfln_to_process_array(zip_01, person)       = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	
		If col_name_array(Col_Name, column) = "COUNTYCOURTDESCRIPTION02" Then dfln_to_process_array(cty_court_02, person) = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "SENTENCEDATE02"           Then dfln_to_process_array(sent_dt_02, person)   = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ADDRESS1_02"              Then dfln_to_process_array(addr1_02, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ADDRESS2_02"              Then dfln_to_process_array(addr2_02, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "CITY02"                   Then dfln_to_process_array(city_02, person)      = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "STATE02"                  Then dfln_to_process_array(state_02, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ZIP02"                    Then dfln_to_process_array(zip_02, person)       = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)

		If col_name_array(Col_Name, column) = "COUNTYCOURTDESCRIPTION03" Then dfln_to_process_array(cty_court_03, person) = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "SENTENCEDATE03"           Then dfln_to_process_array(sent_dt_03, person)   = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ADDRESS1_03"              Then dfln_to_process_array(addr1_03, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ADDRESS2_03"              Then dfln_to_process_array(addr2_03, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "CITY03"                   Then dfln_to_process_array(city_03, person)      = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "STATE03"                  Then dfln_to_process_array(state_03, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ZIP03"                    Then dfln_to_process_array(zip_03, person)       = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)

		If col_name_array(Col_Name, column) = "COUNTYCOURTDESCRIPTION04" Then dfln_to_process_array(cty_court_04, person) = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "SENTENCEDATE04"           Then dfln_to_process_array(sent_dt_04, person)   = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ADDRESS1_04"              Then dfln_to_process_array(addr1_04, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ADDRESS2_04"              Then dfln_to_process_array(addr2_04, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "CITY04"                   Then dfln_to_process_array(city_04, person)      = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "STATE04"                  Then dfln_to_process_array(state_04, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ZIP04"                    Then dfln_to_process_array(zip_04, person)       = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)

		If col_name_array(Col_Name, column) = "COUNTYCOURTDESCRIPTION05" Then dfln_to_process_array(cty_court_05, person) = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "SENTENCEDATE05"           Then dfln_to_process_array(sent_dt_05, person)   = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ADDRESS1_05"              Then dfln_to_process_array(addr1_05, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ADDRESS2_05"              Then dfln_to_process_array(addr2_05, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "CITY05"                   Then dfln_to_process_array(city_05, person)      = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "STATE05"                  Then dfln_to_process_array(state_05, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ZIP05"                    Then dfln_to_process_array(zip_05, person)       = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)

		If col_name_array(Col_Name, column) = "COUNTYCOURTDESCRIPTION06" Then dfln_to_process_array(cty_court_06, person) = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "SENTENCEDATE06"           Then dfln_to_process_array(sent_dt_06, person)   = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ADDRESS1_06"              Then dfln_to_process_array(addr1_06, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ADDRESS2_06"              Then dfln_to_process_array(addr2_06, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "CITY06"                   Then dfln_to_process_array(city_06, person)      = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "STATE06"                  Then dfln_to_process_array(state_06, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
		If col_name_array(Col_Name, column) = "ZIP06"                    Then dfln_to_process_array(zip_06, person)       = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)

	Next
	
	If developer_mode = TRUE Then MsgBox "In " & dfln_to_process_array(month_reptd, person) & " case # " & dfln_to_process_array(case_numb, person) & " for person: " &_
	   dfln_to_process_array(pers_id, person) & " was reported as a drug felon by " & dfln_to_process_array(cty_court_01, person) & " and sentanced on " &_
	   dfln_to_process_array(sent_dt_01, person) & " listed at address: " & vbNewLine & vbNewLine & dfln_to_process_array(addr1_01, person) &_
	   dfln_to_process_array(addr2_01, person) & vbNewLine & dfln_to_process_array(city_01, person) & ", " & dfln_to_process_array(state_01, person) &_
	   " " & dfln_to_process_array(zip_01, person)
Next

objExcel.Quit 

For person = 0 to Ubound(dfln_to_process_array, 2)
	MAXIS_case_number = dfln_to_process_array(case_numb, person)
	Call navigate_to_MAXIS_screen ("STAT", "MEMB")
	EMReadScreen memb_confirm, 4, 2, 48
	If memb_confirm = "MEMB" Then 
		memb_row = 5
		Do 
			EMReadScreen ref_nbr, 2, memb_row, 3
			If ref_nbr = "  " Then Exit Do 
			EMWriteScreen ref_nbr, 20, 76
			transmit
			EMReadScreen PMI_Number, 8, 4, 46
			PMI_Number = replace(PMI_Number, "_", "")
			IF PMI_Number = dfln_to_process_array(pers_id, person) Then 
				EMReadScreen first_name, 12, 6, 63
				EMReadScreen last_nmae, 25, 6, 30
				EMReadScreen middile_i, 1, 6, 79
				first_name = replace(first_name, "_", "")
				last_nmae = replace(last_nmae, "_", "")
				If middile_i <> "_" Then 
					middile_i = middile_i & "."
				Else 
					middile_i = ""
				End If 
				dfln_to_process_array(clt_name, person) = last_nmae & ", " & first_name & " " & middile_i
				dfln_to_process_array(ref_numb, person) = ref_nbr
				Exit Do
			End If 
			memb_row = memb_row + 1
		Loop until memb_row = 20

		cash_active = FALSE
		snap_active = FALSE 
		hc_active = FALSE 

		Call navigate_to_MAXIS_screen ("STAT", "PROG")
		
		EMReadScreen x_numb, 7, 21, 21
		dfln_to_process_array(worker_nbr, person) = x_numb
		
		EMReadScreen cash_1_status, 4, 6, 74
		EMReadScreen cash_2_status, 4, 7, 74
		EMReadScreen snap_status,   4, 10, 74
		EMReadScreen hc_status,     4, 12, 74
		
		If cash_1_status = "ACTV" OR cash_1_status = "PEND" Then EMReadScreen cash_1_prog, 2, 6, 67
		If cash_2_status = "ACTV" OR cash_1_status = "PEND" Then EMReadScreen cash_2_prog, 2, 6, 67	
		IF cash_1_status = "ACTV" OR cash_2_status = "ACTV" Then cash_active = TRUE
		IF snap_status = "ACTV" Then snap_active = TRUE
		IF hc_status = "ACTV" Then hc_active = TRUE 
		
		IF cash_1_prog = "GA" OR cash_2_prog = "GA" Then dfln_to_process_array(cash_prog, person) = "GA"
		IF cash_1_prog = "MS" OR cash_2_prog = "MS" Then dfln_to_process_array(cash_prog, person) = "MSA"
		IF cash_1_prog = "RC" OR cash_2_prog = "RC" Then dfln_to_process_array(cash_prog, person) = "RCA"
		IF cash_1_prog = "MF" OR cash_2_prog = "MF" Then dfln_to_process_array(cash_prog, person) = "MFIP"
		IF cash_1_prog = "DW" OR cash_2_prog = "DW" Then dfln_to_process_array(cash_prog, person) = "DWP"
		IF cash_active = FALSE Then dfln_to_process_array(cash_prog, person) = "NONE"
		
		IF cash_active = FALSE AND snap_active = FALSE AND hc_active = FALSE Then 
			dfln_to_process_array(actv_cty, person) = "Closed"
		Else 
			EMReadScreen cty_code, 2, 21, 23
			dfln_to_process_array(actv_cty, person) = cty_code
		End If 
		
		If dfln_to_process_array(cash_prog, person) = "GA" OR dfln_to_process_array(cash_prog, person) = "MSA" OR dfln_to_process_array(cash_prog, person) = "RCA" Then dfln_to_process_array(case_pop, person) = "Adult"
		If dfln_to_process_array(cash_prog, person) = "MFIP" OR dfln_to_process_array(cash_prog, person) = "DWP" Then dfln_to_process_array(case_pop, person) = "Family"
		
		If dfln_to_process_array(case_pop, person) = "" then 
			dfln_to_process_array(case_pop, person) = "UNKNOWN"
			Call navigate_to_MAXIS_screen ("STAT", "PREG") 
			EMReadScreen due_dt, 8, 10, 53
			If due_dt <> "__ __ __" Then 
				due_dt = replace(due_dt, " ", "/")
				If DateDiff("d", date, due_dt) > 0 Then dfln_to_process_array(case_pop, person) = "Family"
			End If 
		End If 
		
		Call navigate_to_MAXIS_screen ("STAT", "ADDR")
		EMReadScreen stat_addr_1, 22, 6, 43
		EMReadScreen stat_addr_2, 22, 7, 43
		EMReadScreen stat_addr_C, 15, 8, 43
		EMReadScreen stat_addr_S, 2,  8, 66
		EMReadScreen stat_addr_Z, 7,  9, 43
		stat_addr_1 = replace(stat_addr_1, "_", "")
		stat_addr_2 = replace(stat_addr_2, "_", "")
		stat_addr_C = replace(stat_addr_C, "_", "")
		stat_addr_S = replace(stat_addr_S, "_", "")
		stat_addr_Z = replace(stat_addr_Z, "_", "")
		
		dfln_to_process_array(stat_addr, person) = stat_addr_1 & "~" & stat_addr_2 & "~" & stat_addr_C & "~" & stat_addr_S & "~" & stat_addr_Z
		
	End If 
Next

Call navigate_to_MAXIS_screen("REPT", "USER")
For person = 0 to Ubound(dfln_to_process_array, 2)
	If right(dfln_to_process_array(worker_nbr, person) , 3) <> "CLS" Then 
		EMWriteScreen dfln_to_process_array(worker_nbr, person), 21, 12
		transmit
		EMWriteScreen "X", 7, 3
		transmit
		EMReadScreen x_numb, 7, 14, 61
		dfln_to_process_array(superv_nbr, person) = x_numb
		transmit
	End If 
	x_numb = ""
Next

back_to_self

For person = 0 to Ubound(dfln_to_process_array, 2)
	If dfln_to_process_array(case_pop, person) = "UNKNOWN" then
		MAXIS_case_number = dfln_to_process_array(case_numb, person)
		Call navigate_to_MAXIS_screen ("CASE", "PERS")
		EMReadScreen second_pers, 2, 13, 3
		If second_pers = "  " Then 
			dfln_to_process_array(case_pop, person) = "Adult"
		Else
			EMReadScreen relationship, 20, 14, 18
			relationship = trim(relationship)
			If relationship = "Child" Then 
				dfln_to_process_array(case_pop, person) = "Family"
				other_pers = "kid"
			ElseIf relationship = "Spouse" Then 
				EMReadscreen other_pers, 2, 16, 3
				If other_pers = "  " Then dfln_to_process_array(case_pop, person) = "Adult"
			End If 
		End If 
	End If 
Next

back_to_self

For person = 0 to Ubound(dfln_to_process_array, 2)
	Call navigate_to_MAXIS_screen ("STAT", "DFLN")
	mx_row = 6
	If dfln_to_process_array(cty_court_01, person) <> "?" Then 
		EMWriteScreen dfln_to_process_array(cty_court_01, person), 6, 41
		sent_day = right("00" & DatePart("d", dfln_to_process_array(sent_dt_01, person)), 2)  
		sent_mth = right("00" & DatePart("m", dfln_to_process_array(sent_dt_01, person)), 2)
		sent_year = right(DatePart("yyyy", dfln_to_process_array(sent_dt_01, person)), 2)    
		EMWriteScreen sent_day, 6, 27
		EMWriteScreen sent_mth, 6, 30
		EMWriteScreen sent_year, 6, 33
		EMWriteScreen dfln_to_process_array(state_01, person), 6, 75  
		mx_row = mx_row + 1
	End If 	
	     
	If dfln_to_process_array(cty_court_02, person) <> "?" Then 
		EMWriteScreen dfln_to_process_array(cty_court_02, person), 6, 41
		sent_day = right("00" & DatePart("d", dfln_to_process_array(sent_dt_02, person)), 2)  
		sent_mth = right("00" & DatePart("m", dfln_to_process_array(sent_dt_02, person)), 2)
		sent_year = right(DatePart("yyyy", dfln_to_process_array(sent_dt_02, person)), 2)    
		EMWriteScreen sent_day, 6, 27
		EMWriteScreen sent_mth, 6, 30
		EMWriteScreen sent_year, 6, 33
		EMWriteScreen dfln_to_process_array(state_02, person), 6, 75  
		mx_row = mx_row + 1
	End If 	
	
	If dfln_to_process_array(cty_court_01, person) <> "?" Then 
		EMWriteScreen dfln_to_process_array(cty_court_01, person), 6, 41
		sent_day = right("00" & DatePart("d", dfln_to_process_array(sent_dt_01, person)), 2)  
		sent_mth = right("00" & DatePart("m", dfln_to_process_array(sent_dt_01, person)), 2)
		sent_year = right(DatePart("yyyy", dfln_to_process_array(sent_dt_01, person)), 2)    
		EMWriteScreen sent_day, 6, 27
		EMWriteScreen sent_mth, 6, 30
		EMWriteScreen sent_year, 6, 33
		EMWriteScreen dfln_to_process_array(state_01, person), 6, 75  
		mx_row = mx_row + 1
	End If 	

	If col_name_array(Col_Name, column) = "COUNTYCOURTDESCRIPTION02" Then dfln_to_process_array(cty_court_02, person) = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "SENTENCEDATE02"           Then dfln_to_process_array(sent_dt_02, person)   = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ADDRESS1_02"              Then dfln_to_process_array(addr1_02, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ADDRESS2_02"              Then dfln_to_process_array(addr2_02, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "CITY02"                   Then dfln_to_process_array(city_02, person)      = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "STATE02"                  Then dfln_to_process_array(state_02, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ZIP02"                    Then dfln_to_process_array(zip_02, person)       = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)

	If col_name_array(Col_Name, column) = "COUNTYCOURTDESCRIPTION03" Then dfln_to_process_array(cty_court_03, person) = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "SENTENCEDATE03"           Then dfln_to_process_array(sent_dt_03, person)   = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ADDRESS1_03"              Then dfln_to_process_array(addr1_03, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ADDRESS2_03"              Then dfln_to_process_array(addr2_03, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "CITY03"                   Then dfln_to_process_array(city_03, person)      = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "STATE03"                  Then dfln_to_process_array(state_03, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ZIP03"                    Then dfln_to_process_array(zip_03, person)       = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)

	If col_name_array(Col_Name, column) = "COUNTYCOURTDESCRIPTION04" Then dfln_to_process_array(cty_court_04, person) = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "SENTENCEDATE04"           Then dfln_to_process_array(sent_dt_04, person)   = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ADDRESS1_04"              Then dfln_to_process_array(addr1_04, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ADDRESS2_04"              Then dfln_to_process_array(addr2_04, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "CITY04"                   Then dfln_to_process_array(city_04, person)      = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "STATE04"                  Then dfln_to_process_array(state_04, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ZIP04"                    Then dfln_to_process_array(zip_04, person)       = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)

	If col_name_array(Col_Name, column) = "COUNTYCOURTDESCRIPTION05" Then dfln_to_process_array(cty_court_05, person) = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "SENTENCEDATE05"           Then dfln_to_process_array(sent_dt_05, person)   = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ADDRESS1_05"              Then dfln_to_process_array(addr1_05, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ADDRESS2_05"              Then dfln_to_process_array(addr2_05, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "CITY05"                   Then dfln_to_process_array(city_05, person)      = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "STATE05"                  Then dfln_to_process_array(state_05, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ZIP05"                    Then dfln_to_process_array(zip_05, person)       = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)

	If col_name_array(Col_Name, column) = "COUNTYCOURTDESCRIPTION06" Then dfln_to_process_array(cty_court_06, person) = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "SENTENCEDATE06"           Then dfln_to_process_array(sent_dt_06, person)   = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ADDRESS1_06"              Then dfln_to_process_array(addr1_06, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ADDRESS2_06"              Then dfln_to_process_array(addr2_06, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "CITY06"                   Then dfln_to_process_array(city_06, person)      = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "STATE06"                  Then dfln_to_process_array(state_06, person)     = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)
	If col_name_array(Col_Name, column) = "ZIP06"                    Then dfln_to_process_array(zip_06, person)       = trim(objExcel.cells(dfln_to_process_array(row_numb, person), col_name_array(Col_Numb, column)).Value)


Next

Set objNewExcel = CreateObject("Excel.Application")
Set objWorkbook = objNewExcel.Workbooks.Add()

objNewExcel.Visible = True


county_col = 1
objNewExcel.Cells(1, county_col).Value = "County"
objNewExcel.Cells(1, county_col).Font.Bold = True
month_col = 2
objNewExcel.Cells(1, month_col).Value = "Report Month"
objNewExcel.Cells(1, month_col).Font.Bold = True
county_code_col = 3
objNewExcel.Cells(1, county_code_col).Value = "County Number"
objNewExcel.Cells(1, county_code_col).Font.Bold = True
supervisor_col = 4
objNewExcel.Cells(1, supervisor_col).Value = "Supervisor"
objNewExcel.Cells(1, supervisor_col).Font.Bold = True
worker_col = 5
objNewExcel.Cells(1, worker_col).Value = "Worker"
objNewExcel.Cells(1, worker_col).Font.Bold = True

case_nbr_col = 6
objNewExcel.Cells(1, case_nbr_col).Value = "Case Number"
objNewExcel.Cells(1, case_nbr_col).Font.Bold = True
pers_id_col = 7
objNewExcel.Cells(1, pers_id_col).Value = "Person ID"
objNewExcel.Cells(1, pers_id_col).Font.Bold = True
name_col = 8
objNewExcel.Cells(1, name_col).Value = "Name"
objNewExcel.Cells(1, name_col).Font.Bold = True
pop_col = 9
objNewExcel.Cells(1, pop_col).Value = "Population"
objNewExcel.Cells(1, pop_col).Font.Bold = True
cash_prog_col = 10 
objNewExcel.Cells(1, cash_prog_col).Value = "Cash Program"
objNewExcel.Cells(1, cash_prog_col).Font.Bold = True
actv_cty_col = 11
objNewExcel.Cells(1, actv_cty_col).Value = "Active County"
objNewExcel.Cells(1, actv_cty_col).Font.Bold = True

court1_col = 12
objNewExcel.Cells(1, court1_col).Value = "County Court 1"
objNewExcel.Cells(1, court1_col).Font.Bold = True
sntc_dt_1_col = 13
objNewExcel.Cells(1, sntc_dt_1_col).Value = "Sentence Date 1"
objNewExcel.Cells(1, sntc_dt_1_col).Font.Bold = True
addr1_01_col = 14
objNewExcel.Cells(1, addr1_01_col).Value = "Address1_01"
objNewExcel.Cells(1, addr1_01_col).Font.Bold = True
addr2_01_col = 15
objNewExcel.Cells(1, addr2_01_col).Value = "Address2_01"
objNewExcel.Cells(1, addr2_01_col).Font.Bold = True
city1_col = 16
objNewExcel.Cells(1, city1_col).Value = "City 1"
objNewExcel.Cells(1, city1_col).Font.Bold = True
state1_col = 17
objNewExcel.Cells(1, state1_col).Value = "State 1"
objNewExcel.Cells(1, state1_col).Font.Bold = True
zip1_col = 18
objNewExcel.Cells(1, zip1_col).Value = "Zip 1"
objNewExcel.Cells(1, zip1_col).Font.Bold = True

stat_addr_col = 19
objNewExcel.Cells(1, stat_addr_col).Value = "ADDR on STAT"
objNewExcel.Cells(1, stat_addr_col).Font.Bold = True

excel_row = 2

For person = 0 to Ubound(dfln_to_process_array, 2)

	objNewExcel.Cells(excel_row, county_col).Value = county_name
	objNewExcel.Cells(excel_row, month_col).Value  = dfln_to_process_array(month_reptd, person) 
	objNewExcel.Cells(excel_row, county_code_col).Value = right(worker_county_code, 2)
	objNewExcel.Cells(excel_row, supervisor_col).Value = dfln_to_process_array(superv_nbr, person)
	objNewExcel.Cells(excel_row, worker_col).Value = dfln_to_process_array(worker_nbr, person)
	objNewExcel.Cells(excel_row, case_nbr_col).Value = dfln_to_process_array(case_numb, person)   
	objNewExcel.Cells(excel_row, pers_id_col).Value = dfln_to_process_array(pers_id, person) 
	objNewExcel.Cells(excel_row, name_col).Value = dfln_to_process_array(clt_name, person) 
	
	objNewExcel.Cells(excel_row, pop_col).Value = dfln_to_process_array(case_pop, person)   
	objNewExcel.Cells(excel_row, cash_prog_col).Value = dfln_to_process_array(cash_prog, person) 
	objNewExcel.Cells(excel_row, actv_cty_col).Value = dfln_to_process_array(actv_cty, person) 
	
	objNewExcel.Cells(excel_row, court1_col).Value = dfln_to_process_array(cty_court_01, person) 
	objNewExcel.Cells(excel_row, sntc_dt_1_col).Value = dfln_to_process_array(sent_dt_01, person)   
	objNewExcel.Cells(excel_row, addr1_01_col).Value = dfln_to_process_array(addr1_01, person)  
	objNewExcel.Cells(excel_row, addr2_01_col).Value = dfln_to_process_array(addr2_01, person)    
	objNewExcel.Cells(excel_row, city1_col).Value = dfln_to_process_array(city_01, person)   
	objNewExcel.Cells(excel_row, state1_col).Value = dfln_to_process_array(state_01, person)   
	objNewExcel.Cells(excel_row, zip1_col).Value = dfln_to_process_array(zip_01, person)     

	If dfln_to_process_array(stat_addr, person) <> "" Then 
		stat_addr_array = split(dfln_to_process_array(stat_addr, person), "~")
		address_entry = stat_addr_array(0) & " " & stat_addr_array(1) & " " & stat_addr_array(2) & ", " & stat_addr_array(3) & " " & stat_addr_array(4)
		objNewExcel.Cells(excel_row, stat_addr_col).Value = address_entry
	End If 
	excel_row = excel_row + 1
Next
For col_to_autofit = 1 to 19
	ObjNewExcel.columns(col_to_autofit).AutoFit()
Next
'objNewExcel.columns(4).columnwidth = 850
