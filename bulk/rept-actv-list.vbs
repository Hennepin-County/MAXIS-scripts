'PLEASE NOTE: this script was designed to run off of the BULK - pull data into Excel script.
'As such, it might not work if ran separately from that.

'Required for statistical purposes==========================================================================================
name_of_script = "BULK - REPT-ACTV LIST.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 13                      'manual run time in seconds
STATS_denomination = "C"       							'C is for each CASE
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
Call changelog_update("12/10/2016", "Added IV-E, Child Care and FIATed case statuses to script. Also added closing message informing user that script has ended sucessfully.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS----------------------------------------------------------------------
BeginDialog pull_REPT_data_into_excel_dialog, 0, 0, 286, 150, "Pull REPT data into Excel dialog"
  EditBox 135, 25, 145, 15, worker_number
  CheckBox 70, 70, 150, 10, "Check here to run this query county-wide.", all_workers_check
  CheckBox 70, 85, 150, 10, "Identity FIATed cases on the spreadsheet", FIAT_check
  CheckBox 10, 25, 40, 10, "SNAP?", SNAP_check
  CheckBox 10, 40, 40, 10, "Cash?", cash_check
  CheckBox 10, 55, 40, 10, "HC?", HC_check
  CheckBox 10, 70, 40, 10, "EA?", EA_check
  CheckBox 10, 85, 40, 10, "GRH?", GRH_check
  CheckBox 10, 100, 40, 10, "IV-E?", IVE_check
  CheckBox 10, 115, 50, 10, "Child Care?", CC_check
  ButtonGroup ButtonPressed
    OkButton 175, 125, 50, 15
    CancelButton 230, 125, 50, 15
  Text 95, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  Text 70, 45, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  GroupBox 5, 10, 60, 120, "Progs to scan"
  Text 70, 30, 65, 10, "Worker(s) to check:"
  Text 70, 100, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
EndDialog

'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
get_county_code

'Connects to BlueZone
EMConnect ""

'Shows dialog
Dialog pull_rept_data_into_Excel_dialog
If buttonpressed = cancel then stopscript

'Asks to grab COLA related stats (will occur below main info collection)
COLA_stats = MsgBox("Seek COLA income-related info from ACTV cases?", vbYesNo)
If COLA_stats = vbCancel then StopScript				'Cancel button from MsgBox
If COLA_stats = vbYes then collect_COLA_stats = True	'Will use this variable below

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'Checking for MAXIS
Call check_for_MAXIS(True)

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first 4 col as worker, case number, name, and APPL date
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
objExcel.Cells(1, 2).Font.Bold = TRUE
ObjExcel.Cells(1, 3).Value = "NAME"
objExcel.Cells(1, 3).Font.Bold = TRUE
ObjExcel.Cells(1, 4).Value = "NEXT REVW DATE"
objExcel.Cells(1, 4).Font.Bold = TRUE

'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'	Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 5 'Starting with 5 because cols 1-4 are already used

If SNAP_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "SNAP?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	snap_actv_col = col_to_use
	col_to_use = col_to_use + 1
	SNAP_letter_col = convert_digit_to_excel_column(snap_actv_col)
End if
If cash_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "CASH?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	cash_actv_col = col_to_use
	col_to_use = col_to_use + 1
	cash_letter_col = convert_digit_to_excel_column(cash_actv_col)
End if
If HC_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "HC?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	HC_actv_col = col_to_use
	col_to_use = col_to_use + 1
	HC_letter_col = convert_digit_to_excel_column(HC_actv_col)
End if
If EA_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "EA?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	EA_actv_col = col_to_use
	col_to_use = col_to_use + 1
	EA_letter_col = convert_digit_to_excel_column(EA_actv_col)
End if
If GRH_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "GRH?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	GRH_actv_col = col_to_use
	col_to_use = col_to_use + 1
	GRH_letter_col = convert_digit_to_excel_column(GRH_actv_col)
End if
If IVE_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "IV-E?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	IVE_actv_col = col_to_use
	col_to_use = col_to_use + 1
	IVE_letter_col = convert_digit_to_excel_column(IVE_actv_col)
End if
If CC_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "CC?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	CC_actv_col = col_to_use
	col_to_use = col_to_use + 1
	CC_letter_col = convert_digit_to_excel_column(CC_actv_col)
End if
If FIAT_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "FIAT?"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	FIAT_actv_col = col_to_use
	col_to_use = col_to_use + 1
	FIAT_letter_col = convert_digit_to_excel_column(FIAT_actv_col)
End if
If collect_COLA_stats = true then
	ObjExcel.Cells(1, col_to_use).Value = "COLA income types to verify"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	COLA_income_to_verify_col = col_to_use
	col_to_use = col_to_use + 1
End if

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'formatting array
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(x1_number)		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next

	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'Setting the variable for what's to come
excel_row = 2

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("rept", "actv")
	EMWriteScreen worker, 21, 13
	transmit
	EMReadScreen user_worker, 7, 21, 71		'
	EMReadScreen p_worker, 7, 21, 13
	IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 8
	If has_content_check <> " " then

		'Grabbing each case number on screen
		Do
			'Set variable for next do...loop
			MAXIS_row = 7

			'Checking for the last page of cases.
			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12	'Reading case number
				EMReadScreen client_name, 21, MAXIS_row, 21			'Reading client name
				EMReadScreen next_revw_date, 8, MAXIS_row, 42		'Reading application date
				EMReadScreen cash_status, 9, MAXIS_row, 51			'Reading cash status
				EMReadScreen SNAP_status, 1, MAXIS_row, 61			'Reading SNAP status
				EMReadScreen HC_status, 1, MAXIS_row, 64			'Reading HC status
				EMReadScreen EA_status, 1, MAXIS_row, 67			'Reading EA status
				EMReadScreen GRH_status, 1, MAXIS_row, 70			'Reading GRH status
				EMReadScreen IVE_status, 1, MAXIS_row, 73			'Reading IV-E status
				EMReadScreen FIAT_status, 1, MAXIS_row, 77			'Reading the FIAT status of a case
				EMReadScreen CC_status, 1, MAXIS_row, 80			'Reading CC status

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				If trim(MAXIS_case_number) <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & " " & MAXIS_case_number)

				If MAXIS_case_number = "        " then exit do			'Exits do if we reach the end

				'Using if...thens to decide if a case should be added (status isn't blank or inactive and respective box is checked)
				If SNAP_status <> " " and SNAP_status <> "I" and SNAP_check = checked then add_case_info_to_Excel = True
				If HC_status <> " " and HC_status <> "I" and HC_check = checked then add_case_info_to_Excel = True
				If EA_status <> " " and EA_status <> "I" and EA_check = checked then add_case_info_to_Excel = True
				If GRH_status <> " " and GRH_status <> "I" and GRH_check = checked then add_case_info_to_Excel = True
				If IVE_status <> " " and IVE_status <> "I" and IVE_check = checked then add_case_info_to_Excel = True
				If CC_status <> " " and CC_status <> "I" and CC_check = checked then add_case_info_to_Excel = True
				If FIAT_status <> " " and FIAT_status <> "I" and FIAT_check = checked then add_case_info_to_Excel = True

				'Cash requires different handling due to containing multiple program types in one column
				If (instr(cash_status, " A ") <> 0 or instr(cash_status, " P ") <> 0) and cash_check = checked then add_case_info_to_Excel = True

				If add_case_info_to_Excel = True then
					ObjExcel.Cells(excel_row, 1).Value = worker
					ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
					ObjExcel.Cells(excel_row, 3).Value = client_name
					ObjExcel.Cells(excel_row, 4).Value = replace(next_revw_date, " ", "/")
					ObjExcel.Cells(excel_row, 5).Value = abs(days_pending)
					If SNAP_check = checked then ObjExcel.Cells(excel_row, snap_actv_col).Value = SNAP_status
					If cash_check = checked then ObjExcel.Cells(excel_row, cash_actv_col).Value = cash_status
					If HC_check = checked then ObjExcel.Cells(excel_row, HC_actv_col).Value = HC_status
					If EA_check = checked then ObjExcel.Cells(excel_row, EA_actv_col).Value = EA_status
					If GRH_check = checked then ObjExcel.Cells(excel_row, GRH_actv_col).Value = GRH_status
					If IVE_check = checked then ObjExcel.Cells(excel_row, IVE_actv_col).Value = IVE_status
					If CC_check = checked then ObjExcel.Cells(excel_row, CC_actv_col).Value = CC_status
					If FIAT_check = checked then ObjExcel.Cells(excel_row, FIAT_actv_col).Value = FIAT_status
					excel_row = excel_row + 1
				End if
				MAXIS_row = MAXIS_row + 1
				add_case_info_to_Excel = ""	'Blanking out variable
				MAXIS_case_number = ""			'Blanking out variable
				STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
			Loop until MAXIS_row = 19
			PF8

		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

If collect_COLA_stats = True then
	'Reset Excel Row
	excel_row = 2

	'This loop will navigate to UNEA and check each case for the specified types of income
	Do
		'Assign case number from Excel
		MAXIS_case_number = ObjExcel.Cells(excel_row, 2)

		'Exiting if the case number is blank
		If MAXIS_case_number = "" then exit do

		'Navigate to STAT/UNEA for said case number
		call navigate_to_MAXIS_screen("STAT", "UNEA")

		'Reading list of household members, dumping into array
		MAXIS_row = 6		'Second row with a HH member number, first row is always "01"
		HH_member_array = "01"	'Setting this now as the loop won't check the first row
		Do	'reading each one and adding to the variable
			EMReadScreen HH_member_from_list, 2, MAXIS_row, 3
			If HH_member_from_list = "  " then exit do
			HH_member_array = HH_member_array & "|" & HH_member_from_list
			MAXIS_row = MAXIS_row + 1
		Loop until HH_member_from_list = "  "
		HH_member_array = split(HH_member_array, "|")	'Splitting array

		'Will navigate to each one and read the income type. If the income type is one of the COLA-specific incomes, it will add to a variable to be dumped in spreadsheet
		For each HH_member in HH_member_array
			Do
				If HH_member <> "01" Then 'This prevents skipping the first unea panel for memb01.
					EMWriteScreen HH_member, 20, 76	'Writing member number
					transmit					'Transmitting to panel
				End if
				EMReadScreen income_type, 2, 5, 37	'Reading income type
				If income_type = "06" or income_type = "11" or income_type = "12" or income_type = "13" or income_type = "83" or _
				income_type = "17" or income_type = "18" or income_type = "29" or income_type = "08" or income_type = "35" then	'Only runs for certain income types
					If COLA_income_types = "" then 'If blank, it just adds the income. If not, it adds a comma and the income.
						COLA_income_types = "MEMB " & HH_member & ": " & income_type
					Else
						COLA_income_types = COLA_income_types & ", " & "MEMB " & HH_member & ": " & income_type
					End if
				End if
				EMReadScreen current_panel, 1, 2, 73	'reads current and total, to see if we're at the end of the UNEA panels
				EMReadScreen total_panels, 1, 2, 78
				transmit	'goes to the next panel
			Loop until current_panel = total_panels		'End this loop when we've reached the end of all panels
		Next

		'Writes the variable to Excel
		ObjExcel.Cells(excel_row, COLA_income_to_verify_col).Value = COLA_income_types

		'Clears old variables
		HH_member_array = ""
		COLA_income_types = ""

		excel_row = excel_row + 1	'Advances to look at the next row
	Loop until MAXIS_case_number = ""
End if

col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns

'Query date/time/runtime info
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time

'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your REPT/ACTV list has been created.")
