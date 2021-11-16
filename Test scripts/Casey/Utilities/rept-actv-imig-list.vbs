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
CALL changelog_update("01/27/2020", "Updated the handling for CCL.", "MiKayla Handley, Hennepin County")
CALL changelog_update("05/07/2018", "Updated the characters to pull for the client's name.", "MiKayla Handley, Hennepin County")
CALL changelog_update("01/12/2018", "Entering a supervisor X-Number in the Workers to Check will pull all X-Numbers listed under that supervisor in MAXIS. Addiional bug fix where script was missing cases.", "Casey Love, Hennepin County")
Call changelog_update("12/10/2016", "Added IV-E, Child Care and FIATed case statuses to script. Also added closing message informing user that script has ended sucessfully.", "Ilse Ferris, Hennepin County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'This function is used to grab all active X numbers according to the supervisor X number(s) inputted
FUNCTION create_array_of_all_active_x_numbers_by_supervisor(array_name, supervisor_array)
	'Getting to REPT/USER
	CALL navigate_to_MAXIS_screen("REPT", "USER")
	'Sorting by supervisor
	PF5
	PF5
	'Reseting array_name
	array_name = ""
	'Splitting the list of inputted supervisors...
	supervisor_array = replace(supervisor_array, " ", "")
	supervisor_array = split(supervisor_array, ",")
	FOR EACH unit_supervisor IN supervisor_array
		IF unit_supervisor <> "" THEN
			'Entering the supervisor number and sending a transmit
			CALL write_value_and_transmit(unit_supervisor, 21, 12)
			MAXIS_row = 7
			DO
				EMReadScreen worker_ID, 8, MAXIS_row, 5
				worker_ID = trim(worker_ID)
				IF worker_ID = "" THEN EXIT DO
				array_name = trim(array_name & " " & worker_ID)
				MAXIS_row = MAXIS_row + 1
				IF MAXIS_row = 19 THEN
					PF8
					EMReadScreen end_check, 9, 24,14
					If end_check = "LAST PAGE" Then Exit Do
					MAXIS_row = 7
				END IF
			LOOP
		END IF
	NEXT
	'Preparing array_name for use...
	array_name = split(array_name)
END FUNCTION

'THE SCRIPT-------------------------------------------------------------------------
'Determining specific county for multicounty agencies...
'get_county_code
'Connects to BlueZone
EMConnect ""
get_imig_details = true
'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 286, 130, "Pull REPT data into Excel dialog"
  EditBox 135, 20, 145, 15, worker_number
  CheckBox 70, 60, 150, 10, "Check here to run this query county-wide.", all_workers_check
  CheckBox 70, 70, 150, 10, "Identity FIATed cases on the spreadsheet", FIAT_check
  CheckBox 10, 15, 40, 10, "All Active", all_programs
  CheckBox 10, 30, 40, 10, "SNAP", SNAP_check
  CheckBox 10, 40, 40, 10, "CASH", cash_check
  CheckBox 10, 50, 40, 10, "HC", HC_check
  CheckBox 10, 60, 40, 10, "EA", EA_check
  CheckBox 10, 70, 40, 10, "GRH", GRH_check
  CheckBox 10, 80, 40, 10, "IV-E", IVE_check
  CheckBox 10, 90, 50, 10, "CCA", CCA_check
  ButtonGroup ButtonPressed
    OkButton 185, 110, 45, 15
    CancelButton 235, 110, 45, 15
  Text 70, 40, 210, 20, "Enter 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  Text 70, 85, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
  Text 110, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
  GroupBox 5, 5, 55, 100, "Progs to scan"
  Text 70, 25, 65, 10, "Worker(s) to check:"
EndDialog

Do
	Do
		Dialog dialog1
		cancel_without_confirmation
		If (all_workers_check = 0 AND worker_number = "") then MsgBox "Please enter at least one worker number." 'allows user to select the all workers check, and not have worker number be ""
	LOOP until all_workers_check = 1 or worker_number <> ""
	Call check_for_password(are_we_passworded_out)
Loop until check_for_password(are_we_passworded_out) = False		'loops until user is password-ed out

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
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "NAME"
ObjExcel.Cells(1, 4).Value = "NEXT REVW DATE"
FOR i = 1 to 4		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
NEXT

'Figuring out what to put in each Excel col. To add future variables to this, add the checkbox variables below and copy/paste the same code!
'Below, use the "[blank]_col" variable to recall which col you set for which option.
col_to_use = 5 'Starting with 5 because cols 1-4 are already used
If all_programs = checked then
	SNAP_check = checked
	cash_check = checked
	HC_check = checked
	EA_check = checked
	GRH_check = checked
	IVE_check = checked
	CCA_check = checked
END IF
If SNAP_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "SNAP"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	snap_actv_col = col_to_use
	col_to_use = col_to_use + 1
	SNAP_letter_col = convert_digit_to_excel_column(snap_actv_col)
End if
If cash_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "CASH"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	cash_actv_col = col_to_use
	col_to_use = col_to_use + 1
	cash_letter_col = convert_digit_to_excel_column(cash_actv_col)
End if
If HC_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "HC"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	HC_actv_col = col_to_use
	col_to_use = col_to_use + 1
	HC_letter_col = convert_digit_to_excel_column(HC_actv_col)
End if
If EA_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "EA"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	EA_actv_col = col_to_use
	col_to_use = col_to_use + 1
	EA_letter_col = convert_digit_to_excel_column(EA_actv_col)
End if
If GRH_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "GRH"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	GRH_actv_col = col_to_use
	col_to_use = col_to_use + 1
	GRH_letter_col = convert_digit_to_excel_column(GRH_actv_col)
End if
If IVE_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "IV-E"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	IVE_actv_col = col_to_use
	col_to_use = col_to_use + 1
	IVE_letter_col = convert_digit_to_excel_column(IVE_actv_col)
End if
If CCA_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "CC"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	CC_actv_col = col_to_use
	col_to_use = col_to_use + 1
	CC_letter_col = convert_digit_to_excel_column(CC_actv_col)
End if
If FIAT_check = checked then
	ObjExcel.Cells(1, col_to_use).Value = "FIAT"
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
If get_imig_details = true then
	ObjExcel.Cells(1, col_to_use).Value = "IMIG Exist"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	imig_exists_col = col_to_use
	col_to_use = col_to_use + 1

	ObjExcel.Cells(1, col_to_use).Value = "Asylee/Refugee"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	imig_code_col = col_to_use
	col_to_use = col_to_use + 1

	ObjExcel.Cells(1, col_to_use).Value = "SPON Exists"
	objExcel.Cells(1, col_to_use).Font.Bold = TRUE
	spon_exists_col = col_to_use
	col_to_use = col_to_use + 1
End if

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else		'If worker numbers are litsted - this will create an array of workers to check
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'formatting array
	For each x1_number in x1s_from_dialog
		x1_number = trim(ucase(x1_number))					'Formatting the x numbers so there are no errors
		Call navigate_to_MAXIS_screen ("REPT", "USER")		'This part will check to see if the x number entered is a supervisor of anyone
		PF5
		PF5
		EMWriteScreen x1_number, 21, 12
		transmit
		EMReadScreen sup_id_check, 7, 7, 5					'This is the spot where the first person is listed under this supervisor
		IF sup_id_check <> "       " Then 					'If this frist one is not blank then this person is a supervisor
			supervisor_array = trim(supervisor_array & " " & x1_number)		'The script will add this x number to a list of supervisors
		Else
			If worker_array = "" then						'Otherwise this x number is added to a list of workers to run the script on
				worker_array = trim(x1_number)
			Else
				worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
			End if
		End If
		PF3
	Next

	If supervisor_array <> "" Then 				'If there are any x numbers identified as a supervisor, the script will run the function above
		Call create_array_of_all_active_x_numbers_by_supervisor (more_workers_array, supervisor_array)
		workers_to_add = join(more_workers_array, ", ")
		If worker_array = "" then				'Adding all x numbers listed under the supervisor to the worker array
			worker_array = workers_to_add
		Else
			worker_array = worker_array & ", " & trim(ucase(workers_to_add))
		End if
	End If

	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'Setting the variable for what's to come
excel_row = 2
all_case_numbers_array = "*"

For each worker in worker_array
	worker = trim(ucase(worker))					'Formatting the worker so there are no errors
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("rept", "actv")
	EMWriteScreen worker, 21, 13
	TRANSMIT
	EMReadScreen user_worker, 7, 21, 71		'
	EMReadScreen p_worker, 7, 21, 13
	IF user_worker = p_worker THEN PF7		'If the user is checking their own REPT/ACTV, the script will back up to page 1 of the REPT/ACTV

	'msgbox "worker " & worker

	IF worker_number = "X127CCL" or worker = "127CCL" THEN
		DO
			EmReadScreen worker_confirmation, 20, 3, 11 'looking for CENTURY PLAZA CLOSED
			EMWaitReady 0, 0
			'MsgBox "Are we waiting?"
		LOOP UNTIL worker_confirmation = "CENTURY PLAZA CLOSED"
	END IF

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 8
	If has_content_check <> " " then
		'Grabbing each case number on screen
		Do
		    'Set variable for next do...loop
			MAXIS_row = 7
			'Checking for the last page of cases.
			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
			EMReadscreen number_of_pages, 4, 3, 76 'getting page number because to ensure it doesnt fail'
			number_of_pages = trim(number_of_pages)
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
				MAXIS_case_number = trim(MAXIS_case_number)
				If MAXIS_case_number <> "" and instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") <> 0 then exit do
				all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*")

				If MAXIS_case_number = "" Then Exit Do			'Exits do if we reach the end

				'Using if...thens to decide if a case should be added (status isn't blank or inactive and respective box is checked)
				If SNAP_status <> " " and SNAP_status <> "I" and SNAP_check = checked then add_case_info_to_Excel = True
				If HC_status <> " " and HC_status <> "I" and HC_check = checked then add_case_info_to_Excel = True
				If EA_status <> " " and EA_status <> "I" and EA_check = checked then add_case_info_to_Excel = True
				If GRH_status <> " " and GRH_status <> "I" and GRH_check = checked then add_case_info_to_Excel = True
				If IVE_status <> " " and IVE_status <> "I" and IVE_check = checked then add_case_info_to_Excel = True
				If CC_status <> " " and CC_status <> "I" and CCA_check = checked then add_case_info_to_Excel = True
				If FIAT_status <> " " and FIAT_status <> "I" and FIAT_check = checked then add_case_info_to_Excel = True

				'Cash requires different handling due to containing multiple program types in one column
				If (instr(cash_status, " A ") <> 0 or instr(cash_status, " P ") <> 0) and cash_check = checked then add_case_info_to_Excel = True

				If add_case_info_to_Excel = True then
					ObjExcel.Cells(excel_row, 1).Value = worker
					ObjExcel.Cells(excel_row, 2).Value = MAXIS_case_number
					ObjExcel.Cells(excel_row, 3).Value = client_name
					IF next_revw_date <> "        " THEN ObjExcel.Cells(excel_row, 4).Value = replace(next_revw_date, " ", "/")
					ObjExcel.Cells(excel_row, 5).Value = abs(days_pending)
					If SNAP_check = checked then ObjExcel.Cells(excel_row, snap_actv_col).Value = SNAP_status
					If cash_check = checked then ObjExcel.Cells(excel_row, cash_actv_col).Value = cash_status
					If HC_check = checked then ObjExcel.Cells(excel_row, HC_actv_col).Value = HC_status
					If EA_check = checked then ObjExcel.Cells(excel_row, EA_actv_col).Value = EA_status
					If GRH_check = checked then ObjExcel.Cells(excel_row, GRH_actv_col).Value = GRH_status
					If IVE_check = checked then ObjExcel.Cells(excel_row, IVE_actv_col).Value = IVE_status
					If CCA_check = checked then ObjExcel.Cells(excel_row, CC_actv_col).Value = CC_status
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
	END IF
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

If get_imig_details = true then
	'Reset Excel Row
	excel_row = 2

	'This loop will navigate to UNEA and check each case for the specified types of income
	Do
		'Assign case number from Excel
		MAXIS_case_number = ObjExcel.Cells(excel_row, 2)
		'Exiting if the case number is blank
		If MAXIS_case_number = "" then exit do
		'Navigate to STAT/UNEA for said case number
		call navigate_to_MAXIS_screen("STAT", "IMIG")
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

		imig_exists = False
		memb_is_asylee_or_refugee = False
		spon_exists = False
		imig_code = ""
		'Will navigate to each one and read the income type. If the income type is one of the COLA-specific incomes, it will add to a variable to be dumped in spreadsheet
		For each HH_member in HH_member_array
			If HH_member <> "01" Then 'This prevents skipping the first unea panel for memb01.
				EMWriteScreen HH_member, 20, 76	'Writing member number
				transmit					'Transmitting to panel
			End if
			EMReadScreen imig_instance, 1, 2, 73
			If imig_instance = "1" Then
				imig_exists = True
				EMReadScreen imig_code, 2, 6, 45
				EMReadScreen imig_adj, 2, 9, 45

				If imig_code = "21" Then memb_is_asylee_or_refugee = True
				If imig_code = "22" Then memb_is_asylee_or_refugee = True
				If imig_adj = "21" Then memb_is_asylee_or_refugee = True
				If imig_adj = "22" Then memb_is_asylee_or_refugee = True
			End If
		Next

		call navigate_to_MAXIS_screen("STAT", "SPON")

		For each HH_member in HH_member_array
			If HH_member <> "01" Then 'This prevents skipping the first unea panel for memb01.
				EMWriteScreen HH_member, 20, 76	'Writing member number
				transmit					'Transmitting to panel
			End if
			EMReadScreen spon_instance, 1, 2, 73
			If spon_instance = "1" Then spon_exists = True
		Next
		'Writes the variable to Excel
		ObjExcel.Cells(excel_row, imig_exists_col).Value = imig_exists
		ObjExcel.Cells(excel_row, imig_code_col).Value = memb_is_asylee_or_refugee
		ObjExcel.Cells(excel_row, spon_exists_col).Value = spon_exists

		'Clears old variables
		HH_member_array = ""

		excel_row = excel_row + 1	'Advances to look at the next row
	Loop until MAXIS_case_number = ""
End If

col_to_use = col_to_use + 2	'Doing two because the wrap-up is two columns

'Query date/time/runtime info
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time
ObjExcel.Cells(3, col_to_use - 1).Value = "Number of pages found"	'Goes back one, as this is on the next row
ObjExcel.Cells(3, col_to_use).Value = number_of_pages
'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)

script_end_procedure("Success! Your REPT/ACTV list has been created.")
