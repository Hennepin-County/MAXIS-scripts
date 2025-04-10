'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - Health Care Pending Report.vbs"
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
call changelog_update("02/12/2025", "Initial version.", "Casey Love, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'DECLARATIONS ==============================================================================================================

'MAIN LIST columns
Const Caseload_col					= 1
Const Case_Number_col				= 2
Const Case_Name_col					= 3
Const APPL_Date_col					= 4
Const Days_Pending_col				= 5
Const Pended_Date_col				= 6
Const Date_Added_to_List_col		= 7
Const MIPPA_col						= 8
Const METS_Transition_col			= 9
Const EMA_col						= 10
Const SMRT_Application_col			= 11
Const HC_Eval_Date_col				= 12
Const Verifs_Requested_Date_col		= 13
Const Initial_Assignment_Worker_col	= 14
Const Initial_Assignment_Date_col	= 15
Const Day_20_col					= 16
Const Day_20_Assignment_Worker_col	= 17
Const Day_20_Assignment_Date_col	= 18
Const Day_45_col					= 19
Const Day_45_Assignment_Worker_col	= 20
Const Day_45_Assignment_Date_col	= 21
Const Day_55_col					= 22
Const Day_55_Assignment_Worker_col	= 23
Const Day_55_Assignment_Date_col	= 24
Const Day_60_col					= 25
Const Day_60_Assignment_Worker_col	= 26
Const Day_60_Assignment_Date_col	= 27
Const Overdue_col					= 28
Const Needs_Assignment_col			= 29
Const Priority_col					= 30
Const Most_Recent_Assignment_Worker_col		= 31
Const Most_Recent_Assignment_Date_col		= 32
Const Currently_Assigned_col 		= 33


'File Paths
controller_hc_pending_excel = t_drive & "\Eligibility Support\Assignments\ADS Health Care\Functional Data\Current Pending Health Care Cases.xlsx"
snapshot_hc_pending_excel = t_drive & "\Eligibility Support\Assignments\ADS Health Care\Functional Data\HC Pending Snapshot Data.xlsx"
pending_hc_update_cookie = t_drive & "\Eligibility Support\Assignments\ADS Health Care\Functional Data\PendingHCUpdate.txt"

exclude_list = ""
exclude_list = exclude_list & "X127EN6" 			' TEFRA"
exclude_list = exclude_list & "~X127FG1" 			' Foster Care / IV-E"
exclude_list = exclude_list & "~X127EW6" 			' Foster Care / IV-E"
exclude_list = exclude_list & "~X1274EC" 			' Foster Care / IV-E"
exclude_list = exclude_list & "~X127FG2" 			' Foster Care / IV-E"
exclude_list = exclude_list & "~X127EW4" 			' Foster Care / IV-E"
exclude_list = exclude_list & "~X127F3D" 			' MA - BC"
exclude_list = exclude_list & "~X127EK4" 			' LTC+ - General"
exclude_list = exclude_list & "~X127EK9" 			' LTC+ - General"
exclude_list = exclude_list & "~X127EF8" 			' 1800 - Team 160"
exclude_list = exclude_list & "~X127EF9" 			' 1800 - Team 160"
exclude_list = exclude_list & "~X1275H5" 			' Privileged Cases"
exclude_list = exclude_list & "~X127FAT" 			' Privileged Cases"
exclude_list = exclude_list & "~X127F3H" 			' Privileged Cases"
exclude_list = exclude_list & "~X127FF5" 			' Contracted - North Ridge Facilities"
exclude_list = exclude_list & "~X127FG7" 			' Contracted - Monarch Facilities Contract"
exclude_list = exclude_list & "~X127EM4" 			' Contracted - A Villa Facilities Contract"
exclude_list = exclude_list & "~X127EW8" 			' Contracted - Ebenezer Care Center/ Martin Luther Care Center"
exclude_list = exclude_list & "~X127FF8" 			' Contracted - North Memorial"
exclude_list = exclude_list & "~X127FF6" 			' Contracted - HCMC"
exclude_list = exclude_list & "~X127FF7" 			' Contracted - HCMC"
exclude_list = exclude_list & "~X127FI1" 			' METS Retro Request"

exclude_list = exclude_list & "~X127EH3"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127EJ8"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127EK1"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127EK2"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127EK6"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127EK4"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127EK9"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127EM9"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127EN6"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127EP5"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127EP9"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127F3F"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127FE5"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127FH4"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127FH5"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127FI2"			' From BT - these are LTC cases from the spreadsheet
exclude_list = exclude_list & "~X127FI7"			' From BT - these are LTC cases from the spreadsheet

exclude_array = split(UCase(exclude_list), "~")

worker_array = worker_array & "X127HJ5"
worker_array = worker_array & " X127EH8"
worker_array = worker_array & " X127FE6"
worker_array = worker_array & " X127F3F"
worker_array = worker_array & " X127F3P"
worker_array = worker_array & " X127F3K"
worker_array = worker_array & " X127EQ2"
worker_array = worker_array & " X127EH2"
worker_array = worker_array & " X127EJ4"
worker_array = worker_array & " X127EK5"
worker_array = worker_array & " X127EM8"
worker_array = worker_array & " X127EZ2"
worker_array = worker_array & " X127EE4"
worker_array = worker_array & " X127EE5"
worker_array = worker_array & " X127EL1"
worker_array = worker_array & " X127EJ1"
worker_array = worker_array & " X127EE1"
worker_array = worker_array & " X127EE2"
worker_array = worker_array & " X127EE3"
worker_array = worker_array & " X127EE6"
worker_array = worker_array & " X127EE7"
worker_array = worker_array & " X127EG4"
worker_array = worker_array & " X127EL2"
worker_array = worker_array & " X127EL3"
worker_array = worker_array & " X127EL4"
worker_array = worker_array & " X127EL5"
worker_array = worker_array & " X127EL6"
worker_array = worker_array & " X127EL7"
worker_array = worker_array & " X127EL8"
worker_array = worker_array & " X127EL9"
worker_array = worker_array & " X127EN1"
worker_array = worker_array & " X127EN2"
worker_array = worker_array & " X127EN3"
worker_array = worker_array & " X127EN4"
worker_array = worker_array & " X127EN8"
worker_array = worker_array & " X127EN9"
worker_array = worker_array & " X127EQ3"
worker_array = worker_array & " X127EQ4"
worker_array = worker_array & " X127EQ5"
worker_array = worker_array & " X127EQ6"
worker_array = worker_array & " X127EQ7"
worker_array = worker_array & " X127EQ9"
worker_array = worker_array & " X127EX1"
worker_array = worker_array & " X127EX2"
worker_array = worker_array & " X127EX3"
worker_array = worker_array & " X127EX4"
worker_array = worker_array & " X127EX5"
worker_array = worker_array & " X127EX7"
worker_array = worker_array & " X127EX8"
worker_array = worker_array & " X127EX9"
worker_array = worker_array & " X127EH9"
worker_array = worker_array & " X127EM2"
worker_array = worker_array & " X127EH3"
worker_array = worker_array & " X127EJ8"
worker_array = worker_array & " X127EK1"
worker_array = worker_array & " X127EK2"
worker_array = worker_array & " X127EK6"
worker_array = worker_array & " X127EP5"
worker_array = worker_array & " X127EP9"
worker_array = worker_array & " X127FE5"
worker_array = worker_array & " X127FH4"
worker_array = worker_array & " X127FH5"
worker_array = worker_array & " X127FI7"
worker_array = worker_array & " X127EN5"
worker_array = worker_array & " X127EN7"
worker_array = worker_array & " X127EQ1"
worker_array = worker_array & " X127EQ8"
worker_array = worker_array & " X127F4E"
worker_array = worker_array & " X127ES1"
worker_array = worker_array & " X127ES2"
worker_array = worker_array & " X127ES3"
worker_array = worker_array & " X127ES4"
worker_array = worker_array & " X127ES5"
worker_array = worker_array & " X127ES6"
worker_array = worker_array & " X127ES7"
worker_array = worker_array & " X127ES8"
worker_array = worker_array & " X127ES9"
worker_array = worker_array & " X127ET2"
worker_array = worker_array & " X127ET3"
worker_array = worker_array & " X127ET4"
worker_array = worker_array & " X127ET5"
worker_array = worker_array & " X127ET6"
worker_array = worker_array & " X127ET7"
worker_array = worker_array & " X127ET8"
worker_array = worker_array & " X127ET9"
worker_array = worker_array & " X127EZ1"
worker_array = worker_array & " X127EZ5"
worker_array = worker_array & " X127EZ7"
worker_array = worker_array & " X127FA5"
worker_array = worker_array & " X127ET1"
worker_array = worker_array & " X127J8C"
worker_array = worker_array & " X127PB6"
worker_array = worker_array & " X127GS2"

'FOR REVIEW COUNTS
total_cases_count			= 0
mippa_cases_count			= 0
ema_cases_count				= 0
mets_trans_cases_count		= 0
reg_cases_count				= 0
days_1_10_cases_count		= 0
days_11_20_cases_count		= 0
days_21_30_cases_count		= 0
days_31_40_cases_count		= 0
days_41_50_cases_count		= 0
days_51_60_cases_count		= 0
days_60_cases_count			= 0
smrt_days_60_cases_count	= 0
hc_eval_done_cases_count	= 0
verifs_sent_cases_count		= 0
smrt_app_cases_count		= 0
assign_avail_cases_count	= 0
priority_1_cases_count		= 0
priority_2_cases_count		= 0
priority_3_cases_count		= 0
priority_4_cases_count		= 0
priority_5_cases_count		= 0
priority_6_cases_count		= 0
curr_assign_cases_count		= 0

const hsr_name_const = 0
const hsr_count_const = 1

Dim ASSIGNED_ARRAY()
ReDim ASSIGNED_ARRAY(hsr_count_const, 0)
each_hsr = 0

function random_selection(out_of_number, rand_selected)
	'The out_of_number variable is the chance of selection. For a one in three chance, the out_of_number should be set to 3
	'The selected variable is a boolean of if the option queried should be selected. It will return a one in out_of_number chance of TRUE
	rand_selected = False
	Randomize      		 				'Before calling Rnd, use the Randomize statement without an argument to initialize the random-number generator.
	rnd_nbr = rnd						'Create a random number between 0 and 1
	size_up = rnd_nbr * out_of_number	'Multiply by the out-of-number to create a number that is between 0 and the out-of-number (exclusive) - this is a float (decimal number)
	chance_selection = int(size_up)		'Take only the integer of the float from above
	If chance_selection = 0 Then rand_selected = True		'If the integer is 0, (which is a one in out_of_number chance) then the selection is TRUE - we use 0 because there is ALWAYS a 0
end function

lock_main_list = t_drive & "\Eligibility Support\Assignments\ADS Health Care\Functional Data\DataLockCookie.txt"
hold_main_list = t_drive & "\Eligibility Support\Assignments\ADS Health Care\Functional Data\DataHoldCookie.txt"

function create_data_lock(lock_type)
	If lock_type = "MAIN" Then lock_file = lock_main_list
	If lock_type = "HOLD" Then lock_file = hold_main_list

	With (CreateObject("Scripting.FileSystemObject"))
		If .FileExists(lock_file) = False then
			Set objTextStream = .OpenTextFile(lock_file, 2, true)

			'Write the contents of the text file
			If lock_type = "MAIN" Then objTextStream.WriteLine "The Current HC Pending Excel is unavailable as it is being updated."
			If lock_type = "HOLD" Then objTextStream.WriteLine "The Current HC Pending Excel is unavailable as an assignment is happening."
			objTextStream.WriteLine "Lock date: " & date
			objTextStream.WriteLine "Lock time: " & time
			objTextStream.WriteLine "Locked by: " & windows_user_ID

			objTextStream.Close
		End If
	End With
end function

function release_data_lock(lock_type)
	If lock_type = "MAIN" Then lock_file = t_drive & "\Eligibility Support\Assignments\ADS Health Care\Functional Data\DataLockCookie.txt"
	If lock_type = "HOLD" Then lock_file = t_drive & "\Eligibility Support\Assignments\ADS Health Care\Functional Data\DataHoldCookie.txt"

	With (CreateObject("Scripting.FileSystemObject"))
		If .FileExists(lock_file) = True then
			.DeleteFile(lock_file)
		End If
	End With
end function

function review_hc_pending_counts()
	If total_cases_count = 0 Then
		Call create_data_lock("HOLD")

		Call excel_open(controller_hc_pending_excel, True, False, ObjExcel, objWorkbook)
		objExcel.worksheets("Cases").Activate			'Activates the selected worksheet'

		excel_row = 2
		Do
			mippa_case = ""
			mets_trans_case = ""
			ema_case = ""
			smrt_case = ""
			case_overdue = ""
			needs_assignment = ""
			on_assignment = ""


			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, MIPPA_col), mippa_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, METS_Transition_col), mets_trans_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, EMA_col), ema_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, SMRT_Application_col), smrt_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, Overdue_col), case_overdue)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, Needs_Assignment_col), needs_assignment)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, Currently_Assigned_col), on_assignment)
			HC_Eval_date = 				ObjExcel.Cells(excel_row, HC_Eval_Date_col).Value
			Verif_requested_date = 		ObjExcel.Cells(excel_row, Verifs_Requested_Date_col).Value
			case_priority = 			ObjExcel.Cells(excel_row, Priority_col)
			days_pending =				ObjExcel.Cells(excel_row, Days_Pending_col).Value
			worker_name = trim(ObjExcel.Cells(excel_row, Most_Recent_Assignment_Worker_col).Value)

			total_cases_count = total_cases_count + 1
			If mippa_case Then mippa_cases_count = mippa_cases_count + 1
			' If ema_case Then ema_cases_count = ema_cases_count + 1
			If mets_trans_case Then mets_trans_cases_count = mets_trans_cases_count + 1
			If smrt_case Then smrt_app_cases_count = smrt_app_cases_count + 1
			' If not mippa_case and not ema_case and not mets_trans_case Then reg_cases_count = reg_cases_count + 1
			If not mippa_case and not mets_trans_case Then reg_cases_count = reg_cases_count + 1
			If IsDate(HC_Eval_date) = True Then hc_eval_done_cases_count = hc_eval_done_cases_count + 1
			If IsDate(Verif_requested_date) = True Then verifs_sent_cases_count = verifs_sent_cases_count + 1
			If days_pending < 11 Then days_1_10_cases_count = days_1_10_cases_count + 1
			If days_pending > 10 and days_pending < 21 Then days_11_20_cases_count = days_11_20_cases_count + 1
			If days_pending > 20 and days_pending < 31 Then days_21_30_cases_count = days_21_30_cases_count + 1
			If days_pending > 30 and days_pending < 41 Then days_31_40_cases_count = days_31_40_cases_count + 1
			If days_pending > 40 and days_pending < 51 Then days_41_50_cases_count = days_41_50_cases_count + 1
			If days_pending > 50 and days_pending < 61 Then days_51_60_cases_count = days_51_60_cases_count + 1
			If days_pending > 60 Then days_60_cases_count = days_60_cases_count + 1
			If days_pending > 60 and smrt_case Then smrt_days_60_cases_count = smrt_days_60_cases_count + 1
			If needs_assignment = True and on_assignment <> True Then
				assign_avail_cases_count = assign_avail_cases_count + 1
				If case_priority = "1" Then priority_1_cases_count = priority_1_cases_count + 1
				If case_priority = "2" Then priority_2_cases_count = priority_2_cases_count + 1
				If case_priority = "3" Then priority_3_cases_count = priority_3_cases_count + 1
				If case_priority = "4" Then priority_4_cases_count = priority_4_cases_count + 1
				If case_priority = "5" Then priority_5_cases_count = priority_5_cases_count + 1
				If case_priority = "6" Then priority_6_cases_count = priority_6_cases_count + 1
			End If
			If on_assignment = True Then
				curr_assign_cases_count = curr_assign_cases_count + 1
				If ASSIGNED_ARRAY(hsr_name_const, 0) = "" Then
					ASSIGNED_ARRAY(hsr_name_const, 0) = worker_name
					ASSIGNED_ARRAY(hsr_count_const, 0) = 1
				Else
					worker_found = False
					For known_wrkr = 0 to UBound(ASSIGNED_ARRAY, 2)
						If ASSIGNED_ARRAY(hsr_name_const, known_wrkr) = worker_name Then
							worker_found = True
							ASSIGNED_ARRAY(hsr_count_const, known_wrkr) = ASSIGNED_ARRAY(hsr_count_const, known_wrkr) + 1
						End If
					Next

					If worker_found = False Then
						ReDim preserve ASSIGNED_ARRAY(hsr_count_const, each_hsr)
						ASSIGNED_ARRAY(hsr_name_const, each_hsr) = worker_name
						ASSIGNED_ARRAY(hsr_count_const, each_hsr) = 1
						each_hsr = each_hsr + 1
					End If
				End If
			End If

			' caseload_number = 			ObjExcel.Cells(excel_row, Caseload_col).Value
			' MAXIS_case_number = 		ObjExcel.Cells(excel_row, Case_Number_col).Value
			' case_name = 				ObjExcel.Cells(excel_row, Case_Name_col).Value
			' APPL_date = 				ObjExcel.Cells(excel_row, APPL_Date_col).Value
			' Days_pending = 				ObjExcel.Cells(excel_row, Days_Pending_col).Value
			' pended_date = 				ObjExcel.Cells(excel_row, Pended_Date_col).Value
			' date_added_to_list = 		ObjExcel.Cells(excel_row, Date_Added_to_List_col).Value
			' initial_assignment_worker = ObjExcel.Cells(excel_row, Initial_Assignment_Worker_col).Value
			' initial_assignment_date = 	ObjExcel.Cells(excel_row, Initial_Assignment_Date_col).Value
			' Day_20_date = 				ObjExcel.Cells(excel_row, Day_20_col).Value
			' Day_20_assignment_worker = 	ObjExcel.Cells(excel_row, Day_20_Assignment_Worker_col).Value
			' Day_20_assignment_date = 	ObjExcel.Cells(excel_row, Day_20_Assignment_Date_col).Value
			' Day_45_date = 				ObjExcel.Cells(excel_row, Day_45_col).Value
			' Day_45_assignment_worker = 	ObjExcel.Cells(excel_row, Day_45_Assignment_Worker_col).Value
			' Day_45_assignment_date = 	ObjExcel.Cells(excel_row, Day_45_Assignment_Date_col).Value
			' Day_55_date = 				ObjExcel.Cells(excel_row, Day_55_col).Value
			' Day_55_assignment_worker = 	ObjExcel.Cells(excel_row, Day_55_Assignment_Worker_col).Value
			' Day_55_assignment_date = 	ObjExcel.Cells(excel_row, Day_55_Assignment_Date_col).Value
			' Day_60_date = 				ObjExcel.Cells(excel_row, Day_60_col).Value
			' Day_60_assignment_worker = 	ObjExcel.Cells(excel_row, Day_60_Assignment_Worker_col).Value
			' Day_60_assignment_date = 	ObjExcel.Cells(excel_row, Day_60_Assignment_Date_col).Value
			' Last_Assingment_Date = 		ObjExcel.Cells(excel_row, Most_Recent_Assignment_Date_col).Value
			' trim(ObjExcel.Cells(excel_row, Most_Recent_Assignment_Worker_col).Value)
			' ' case_priority = 			ObjExcel.Cells(excel_row, Priority_col).Value


			excel_row = excel_row + 1
			next_MAXIS_case_number = trim(ObjExcel.Cells(excel_row, Case_Number_col).Value)
		LOOP until next_MAXIS_case_number = ""

		ObjExcel.ActiveWorkbook.Close
		ObjExcel.Application.Quit
		ObjExcel.Quit
		Call release_data_lock("HOLD")
	End If

	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 506, 260, "Pending Health Care Counts"
		ButtonGroup ButtonPressed
			OkButton 450, 240, 50, 15
			'CancelButton 450, 240, 50, 15
		Text 15, 15, 50, 10, "Total Cases:"
		Text 75, 15, 35, 10, total_cases_count
		Text 190, 15, 30, 10, " MIPPAs:"
		Text 30, 90, 20, 10, mippa_cases_count
		Text 200, 25, 20, 10, " EMA:"
		Text 230, 25, 35, 10, ema_cases_count
		Text 160, 35, 65, 10, "METS Transitions:"
		Text 230, 35, 35, 10, mets_trans_cases_count
		Text 185, 45, 40, 10, "  Standard: "
		Text 230, 45, 35, 10, reg_cases_count
		GroupBox 15, 60, 485, 175, "Standard Cases: " & reg_cases_count
		Text 25, 75, 60, 10, "Pending Days"
		Text 230, 15, 35, 10, days_1_10_cases_count
		Text 55, 90, 75, 10, "1 - 10 Days"
		Text 30, 100, 20, 10, days_11_20_cases_count
		Text 55, 100, 80, 10, "11 - 20 Days"
		Text 30, 110, 20, 10, days_21_30_cases_count
		Text 55, 110, 85, 10, "21 - 30 Days"
		Text 30, 120, 20, 10, days_31_40_cases_count
		Text 55, 120, 85, 10, "31 - 40 Days"
		Text 30, 130, 20, 10, days_41_50_cases_count
		Text 55, 130, 85, 10, "41 - 50 Days"
		Text 30, 140, 20, 10, days_51_60_cases_count
		Text 55, 140, 85, 10, "51 - 60 Days"
		Text 30, 150, 20, 10, days_60_cases_count
		Text 55, 150, 85, 10, "Over 60 Days"
		Text 20, 175, 70, 10, "Work Process"
		Text 30, 190, 20, 10, hc_eval_done_cases_count
		Text 55, 190, 75, 10, "HC Eval Done"
		Text 30, 200, 20, 10, verifs_sent_cases_count
		Text 55, 200, 65, 10, "Verifs Sent"
		Text 30, 210, 20, 10, smrt_app_cases_count
		Text 55, 210, 65, 10, "SMRT App"
		Text 30, 220, 20, 10, smrt_days_60_cases_count
		Text 55, 220, 100, 10, "SMRT App over Day 60"
		Text 145, 75, 50, 10, "Assignments"
		Text 155, 90, 25, 10, assign_avail_cases_count
		Text 180, 90, 100, 10, "Available for Assignment"
		Text 160, 100, 25, 10, priority_1_cases_count
		Text 185, 100, 135, 10, "Priority 1 - Overdue and Verifs Due"
		Text 160, 110, 25, 10, priority_2_cases_count
		Text 185, 110, 135, 10, "Priority 2 - HC Eval Not Complete"
		Text 160, 120, 25, 10, priority_3_cases_count
		Text 185, 120, 135, 10, "Priority 3 - Case at Day 20"
		Text 160, 130, 25, 10, priority_4_cases_count
		Text 185, 130, 135, 10, "Priority 4 - Case at Day 45"
		Text 160, 140, 25, 10, priority_5_cases_count
		Text 185, 140, 135, 10, "Priority 5 - Case at Day 55"
		Text 160, 150, 25, 10, priority_6_cases_count
		Text 185, 150, 135, 10, "Priority 6 - Case at Day 60"
		Text 330, 90, 25, 10, curr_assign_cases_count
		Text 355, 90, 100, 10, "Currently Assigned"
		y_pos = 100
		' If ASSIGNED_ARRAY(hsr_name_const, 0) = "" Then
			For known_wrkr = 0 to UBound(ASSIGNED_ARRAY, 2)
				Text 335, y_pos, 25, 10, ASSIGNED_ARRAY(hsr_count_const, known_wrkr)
				Text 360, y_pos, 135, 10, " - " & ASSIGNED_ARRAY(hsr_name_const, known_wrkr)
				y_pos = y_pos + 1
			Next
		' End If
	EndDialog

	Dialog Dialog1

end function

'===========================================================================================================================
'Creating an object for the stream of text which we'll use frequently
Dim objTextStream

'THE SCRIPT-------------------------------------------------------------------------
'Gathering county code for multi-county...
get_county_code

'Connects to BlueZone
EMConnect ""

update_date = ""
update_time = ""
With (CreateObject("Scripting.FileSystemObject"))
	If .FileExists(pending_hc_update_cookie) = True then
		Set objTextStream = .OpenTextFile(pending_hc_update_cookie, ForReading)

		'Reading the entire text file into a string
		every_line_in_text_file = objTextStream.ReadAll

		'Splitting the text file contents into an array which will be sorted
		pending_hc_run_details = split(every_line_in_text_file, vbNewLine)

		For Each text_line in pending_hc_run_details
			text_line = trim(text_line)
			If text_line <> "" Then
				line_info = split(text_line, "&*^&*^")
				If line_info(0) = "update_date" Then update_date = line_info(1)
				If line_info(0) = "update_time" Then update_time = line_info(1)
			End If
		Next
		objTextStream.Close
	End If
End With

update_info_line = ""
If update_date <> "" Then update_info_line = "HC Pending Detail Last Updated: " & update_date
If update_time <> "" Then update_info_line = update_info_line & " at " & update_time

admin_run = False
If windows_user_ID = "CALO001" Then admin_run = true
If windows_user_ID = "DACO003" Then admin_run = true
If windows_user_ID = "MARI001" Then admin_run = true
If windows_user_ID = "BETE001" Then admin_run = true
If windows_user_ID = "YEYA001" Then admin_run = true
If windows_user_ID = "JAAR001" Then admin_run = true
script_run_options = "Individual Worker Assignment Creation"+chr(9)+"Open My Worklist"+chr(9)+"Complete Individual Worklist"
If admin_run = true Then script_run_options = "Individual Worker Assignment Creation"+chr(9)+"List Management"+chr(9)+"Open My Worklist"+chr(9)+"Complete Individual Worklist"+chr(9)+"Review Counts"

BeginDialog Dialog1, 0, 0, 216, 105, "Health Care Pending Assignments"
  DropListBox 10, 60, 195, 45, script_run_options, operation_selection
  ButtonGroup ButtonPressed
    OkButton 100, 80, 50, 15
    CancelButton 155, 80, 50, 15
  Text 10, 10, 195, 20, "This script helps facilitate the pending Health Care cases and assignments."
  If admin_run = true Then Text 15, 30, 190, 10, update_info_line
  Text 10, 50, 110, 10, "Select the operation needed:"
EndDialog

'Dialog asks what stats are being pulled
Dialog Dialog1
cancel_without_confirmation

run_assignment_selection = False
run_list_management = False
If operation_selection = "Individual Worker Assignment Creation" Then run_assignment_selection = True
If operation_selection = "Complete Individual Worklist" Then run_assignment_selection = True
If operation_selection = "List Management" Then run_list_management = True

If operation_selection = "Open My Worklist" Then

	'Finding the user name - we aren't using the function because we need the comma in place
	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	SQL_table = "SELECT * from ES.V_ESAllStaff WHERE EmpLogOnID = '" & windows_user_ID & "'"				'identifying the table that stores the ES Staff user information

	'This is the file path the data tables
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open SQL_table, objConnection							'Here we connect to the data tables

	Do While NOT objRecordSet.Eof										'now we will loop through each item listed in the table of ES Staff
		table_user_id = objRecordSet("EmpLogOnID")						'setting the user ID from table data
		If table_user_id = windows_user_ID Then							'If the ID on thils loop of the data information matches the ID of the person running the script, we have found the staff person
			worker_name = objRecordSet("EmpFullName")				'Save the user name
			Exit Do														'if we have found the person, we stop looping
		End If
		objRecordSet.MoveNext											'Going to the next row in the table
	Loop

	'Now we disconnect from the table and close the connections
	objRecordSet.Close
	objConnection.Close
	Set objRecordSet=nothing
	Set objConnection=nothing

	worker_name = trim(worker_name)
	name_array = split(worker_name, ",")
	last_name = trim(name_array(0))
	first_name =  trim(name_array(1))
	If InStr(first_name, " ") Then
		first_name_array = split(first_name)
		first_name = first_name_array(0)
	End If
	indv_worklist_file_name = first_name & " " & left(last_name, 1) & " Assignment.xlsx"
	indv_worklist_file_path = t_drive & "\Eligibility Support\Assignments\ADS Health Care\" & indv_worklist_file_name

	If objFSO.FileExists(indv_worklist_file_path) Then
		Call excel_open(indv_worklist_file_path, True, False, ObjWrkrExcel, objWrkrWorkbook)
	Else
		no_worklist_msg = MsgBox("It appears there is no existing worklist for " & first_name & " " & left(last_name, 1) & "." & vbCr & vbCr & "Would you like to create one now?", vbQuestion + vbYesNo, "No Worklist Found")
		If no_worklist_msg = vbYes Then
			operation_selection = "Individual Worker Assignment Creation"
			run_assignment_selection = True
		End If
		If no_worklist_msg = vbNo Then end_msg = "No worklist found." & vbCr & vbCr & "Script ended as there is no request for a new worklist to be created."
	End If
End If

If operation_selection = "Review Counts" Then review_hc_pending_counts

If run_list_management = True Then
	'Dialog asks what stats are being pulled
	Do
		Do
			err_msg = ""

			Dialog1 = ""
			' DropListBox 80, 30, 130, 45, "Check All non-LTC"+chr(9)+"Limited or Quick", worker_selection_option
			BeginDialog Dialog1, 0, 0, 231, 175, "ADMIN Pending HC Work"
				DropListBox 80, 30, 130, 45, "Check All non-LTC", worker_selection_option
				CheckBox 15, 50, 175, 10, "Check here to add new cases to the report.", add_new_cases_checkbox
				CheckBox 15, 65, 205, 10, "Check here to remove cases from the report if not on PND2.", remove_from_list_checkbox
				CheckBox 15, 80, 205, 10, "Check here to evaluate case details.", evaluate_case_details_checkbox
				CheckBox 15, 95, 205, 10, "Check here to evaluate assignments.", evaluate_assignments_checkbox
				CheckBox 15, 110, 205, 10, "Check here to capture a Snapshot of Pending data", data_snapshot_checkbox
				ButtonGroup ButtonPressed
					OkButton 115, 150, 50, 15
					CancelButton 170, 150, 50, 15
					PushButton 150, 10, 70, 15, "Review Counts", review_counts_btn
				Text 10, 10, 125, 10, "Ongoing Pending Health Care Report"
				Text 15, 35, 65, 10, "Workers to Check:"
				Text 10, 125, 210, 20, "NOTE: running queries county-wide can take a significant amount of time and resources. This should be done after hours."
			EndDialog

			Dialog Dialog1
			cancel_without_confirmation

			If ButtonPressed = review_counts_btn Then
				err_msg = "LOOP"
				review_hc_pending_counts
			End If
		Loop until err_msg = ""
		CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
	Loop until are_we_passworded_out = false					'loops until user passwords back in
	'Checking for MAXIS
	Call check_for_MAXIS(False)

	worklist_start_time = timer

	If add_new_cases_checkbox = checked Then add_pending_cases_to_report = True
	If add_new_cases_checkbox = unchecked Then add_pending_cases_to_report = False
	If remove_from_list_checkbox = checked Then remove_no_longer_pending_cases = True
	If remove_from_list_checkbox = unchecked Then remove_no_longer_pending_cases = False
	If evaluate_case_details_checkbox = checked Then evaluate_cases_in_MAXIS = True
	If evaluate_case_details_checkbox = unchecked Then evaluate_cases_in_MAXIS = False
	If evaluate_assignments_checkbox = checked Then evaluate_assignments = True
	If evaluate_assignments_checkbox = unchecked Then evaluate_assignments = False
	If worker_selection_option = "Check All non-LTC" Then search_all = True
	If worker_selection_option = "Limited or Quick" Then search_all = False

	If search_all = False Then remove_no_longer_pending_cases = False 		'Cannot remove cases if we don't look at them all

	capture_pnd2 = False
	If add_pending_cases_to_report = True or remove_no_longer_pending_cases = True Then capture_pnd2 = True

	If capture_pnd2 = True Then
		call navigate_to_MAXIS_screen("REPT", "USER")

		'Hitting PF5 to force sorting, which allows directly selecting a county
		PF5

		'Inserting county
		EMWriteScreen county_code, 21, 6
		transmit

		'Declaring the MAXIS row
		MAXIS_row = 7
		If search_all = True Then
			Do
				Do
					'Reading MAXIS information for this row, adding to spreadsheet
					EMReadScreen worker_ID, 8, MAXIS_row, 5					'worker ID
					worker_ID = UCase(trim(worker_ID))
					If worker_ID = "        " then exit do					'exiting before writing to array, in the event this is a blank (end of list)
					exclude_x_number = False
					for each basket in exclude_array
						If basket = worker_ID Then
							exclude_x_number = True
							Exit For
						End If
					next
					If exclude_x_number = False  Then worker_array = trim(worker_array & " " & worker_ID)				'writing to variable
					MAXIS_row = MAXIS_row + 1
				Loop until MAXIS_row = 19

				'Seeing if there are more pages. If so it'll grab from the next page and loop around, doing so until there's no more pages.
				EMReadScreen more_pages_check, 7, 19, 3
				If more_pages_check = "More: +" then
					PF8			'getting to next screen
					MAXIS_row = 7	'redeclaring MAXIS row so as to start reading from the top of the list again
				End if
			Loop until more_pages_check = "More:  " or more_pages_check = "       "	'The or works because for one-page only counties, this will be blank
		End If
		worker_array = split(worker_array)
	End If

	MAXIS_footer_month = CM_mo
	MAXIS_footer_year = CM_yr

	'Starting the query start time (for the query runtime at the end)
	query_start_time = timer

	const worker_numb_const		= 0
	const case_num_const 		= 1
	const case_name_const 		= 2
	const appl_date_const 		= 3
	const on_controller_const	= 4
	const on_PND2_const			= 5
	const excel_row_const 		= 6
	const on_assign_const		= 7
	const final_hc_const		= 10

	Dim HC_REPT_PND2()
	ReDim HC_REPT_PND2(final_hc_const, 0)
	case_count = 0


	If capture_pnd2 = True Then
		'Setting the variable for what's to come
		all_case_numbers_array = "*"

		For each worker in worker_array
			back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
			Call navigate_to_MAXIS_screen("REPT", "PND2")       'looking at PND2 to confirm day 30 AND look for MSA cases - which get 60 days
			EMWriteScreen worker, 21, 13
			transmit
			'This code is for bypassing a warning box if the basket has too many cases
			EMWaitReady 0, 0
			row = 1
			col = 1
			EMSearch "The REPT:PND2 Display Limit Has Been Reached.", row, col
			If row <> 0 THEN transmit

			'TODO add handling to read for an additional app line so that we are sure we are reading the correct line for days pending and cash program
			'Skips workers with no info
			EMReadScreen has_content_check, 6, 3, 74
			If has_content_check <> "0 Of 0" then
				'Grabbing each case number on screen
				Do
					MAXIS_row = 7
					Do
						EMReadScreen MAXIS_case_number, 8, MAXIS_row, 5	'Reading case number
						EMReadScreen client_name, 22, MAXIS_row, 16		'Reading client name
						EMReadScreen APPL_date, 8, MAXIS_row, 38		'Reading application date
						EMReadScreen days_pending, 4, MAXIS_row, 49		'Reading days pending
						EMReadScreen HC_status, 1, MAXIS_row, 65		'Reading HC status

						'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
						client_name = trim(client_name)
						MAXIS_case_number = trim(MAXIS_case_number)
						If client_name <> "ADDITIONAL APP" Then			'When there is an additional app on this rept, the script actually reads a case number even though one is not visible to the worker on the screen - so we are skipping this ghosting issue because it will ALWAYS find the previous case number.
							If MAXIS_case_number <> "" and instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") <> 0 then exit do
							all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*")
						End If

						If MAXIS_case_number = "" AND client_name = "" Then Exit Do			'Exits do if we reach the end

						'If additional application is rec'd then the excel output is the client's name, not ADDITIONAL APP
						if client_name = "ADDITIONAL APP" then
							EMReadScreen alt_client_name, 22, MAXIS_row - 1, 16
							client_name = "* " & trim(alt_client_name)                    'replaces alt name as the client name
						Else
							EMReadScreen next_client, 22, MAXIS_row + 1, 16
							next_client = trim(next_client)
							If next_client = "ADDITIONAL APP" Then client_name = "* " & client_name
						END IF

						'Cleaning up each program's status
						HC_status = trim(replace(HC_status, "_", ""))

						If HC_status <> "" then add_case_info_to_Excel = True

						If add_case_info_to_Excel = True then
							ReDim preserve HC_REPT_PND2(final_hc_const, case_count)
							HC_REPT_PND2(worker_numb_const, case_count) = worker
							HC_REPT_PND2(case_num_const, case_count) = MAXIS_case_number
							HC_REPT_PND2(case_name_const, case_count) = client_name
							HC_REPT_PND2(appl_date_const, case_count) = replace(APPL_date, " ", "/")
							HC_REPT_PND2(on_controller_const, case_count) = False

							case_count = case_count + 1
						End if
						MAXIS_row = MAXIS_row + 1
						add_case_info_to_Excel = ""	'Blanking out variable
						MAXIS_case_number = ""			'Blanking out variable
					Loop until MAXIS_row = 19
					PF8
					EMReadScreen last_page_check, 21, 24, 2
				Loop until last_page_check = "THIS IS THE LAST PAGE"
			End if
			STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
		next
	End If

	With (CreateObject("Scripting.FileSystemObject"))
		If .FileExists(lock_main_list) = True then script_end_procedure("HC Pending details are being updated by somneone else. Try again in a little while.")

		list_being_viewed = .FileExists(hold_main_list)
		If list_being_viewed = True Then MsgBox "Another worker is pulling an assignment. The script will pause while this completes. It usually takes less than a minute to become available. Please wait."
		Do while list_being_viewed = True
			' WScript.Sleep 200
			EMWaitReady 0, 1000
			list_being_viewed = .FileExists(hold_main_list)
		Loop
	End With
	Call create_data_lock("MAIN")

	' interview_tracking_excel = t_drive & "\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage\Added to Work List\Interview Tracking.xlsx"
	Call excel_open(controller_hc_pending_excel, True, False, ObjExcel, objWorkbook)
	objExcel.worksheets("Cases").Activate			'Activates the selected worksheet'

	Dim known_case_number_array()
	ReDim known_case_number_array(final_hc_const, 0)
	known_case_count = 0


	cases_available_for_assignment = 0
	bottom_threshold = 10
	excel_row = 2
	Do
		MAXIS_case_number = ObjExcel.Cells(excel_row, Case_Number_col).Value		'establishing what the case number is for each case
		avail_for_assign = ""
		ReDim Preserve known_case_number_array(final_hc_const, known_case_count)
		known_case_number_array(case_num_const, known_case_count) = trim(MAXIS_case_number)
		Call read_boolean_from_excel(ObjExcel.Cells(excel_row, Currently_Assigned_col), known_case_number_array(on_assign_const, known_case_count))
		Call read_boolean_from_excel(ObjExcel.Cells(excel_row, Needs_Assignment_col), avail_for_assign)
		If avail_for_assign = True and known_case_number_array(on_assign_const, known_case_count) <> True Then cases_available_for_assignment = cases_available_for_assignment + 1
		known_case_number_array(on_PND2_const, known_case_count) = False
		known_case_number_array(excel_row_const, known_case_count) = excel_row
		known_case_count = known_case_count + 1
		excel_row = excel_row + 1									'moves Excel to next row
		next_MAXIS_case_number = ObjExcel.Cells(excel_row, Case_Number_col).Value		'establishing what the case number is for each case
	LOOP until next_MAXIS_case_number = ""								'Loops until all the case have been noted

	excel_row_new_cases_start = excel_row
	If capture_pnd2 = True Then
		current_pnd2_count = 0
		For pnd2_case = 0 to UBound(HC_REPT_PND2, 2)
			For controller_case = 0 to UBound(known_case_number_array, 2)
				' MsgBox "controller_case - " & controller_case & vbCr & "HC_REPT_PND2(case_num_const, pnd2_case) - " & HC_REPT_PND2(case_num_const, pnd2_case)
				If known_case_number_array(case_num_const, controller_case) = HC_REPT_PND2(case_num_const, pnd2_case) Then
					HC_REPT_PND2(on_controller_const, pnd2_case) = True
					known_case_number_array(on_PND2_const, controller_case) = True
					Exit For
				End If
			Next
			If add_pending_cases_to_report = True Then
				If HC_REPT_PND2(on_controller_const, pnd2_case) = False Then
					ObjExcel.Cells(excel_row, Caseload_col).Value		= HC_REPT_PND2(worker_numb_const, pnd2_case)
					ObjExcel.Cells(excel_row, Case_Number_col).Value	= HC_REPT_PND2(case_num_const, pnd2_case)
					ObjExcel.Cells(excel_row, Case_Name_col).Value		= HC_REPT_PND2(case_name_const, pnd2_case)
					ObjExcel.Cells(excel_row, APPL_Date_col).Value		= HC_REPT_PND2(appl_date_const, pnd2_case)
					ObjExcel.Cells(excel_row, Date_Added_to_List_col).Value		= Date

					excel_row = excel_row + 1
				End If
			End If
		Next
		objWorkbook.Save()		'saving the excel
		current_pnd2_count = UBound(HC_REPT_PND2, 2)+1
	End If

	If remove_no_longer_pending_cases = True Then

		hc_cases_acted_on = user_myDocs_folder & "hc_closed_list_at_" & replace(replace(replace(now, "/", "_"),":", "_")," ", "_") & ".txt"
		With (CreateObject("Scripting.FileSystemObject"))
			If .FileExists(hc_cases_acted_on) = True then
				.DeleteFile(hc_cases_acted_on)
			End If

			'If the file doesn't exist, it needs to create it here and initialize it here! After this, it can just exit as the file will now be initialized

			If .FileExists(hc_cases_acted_on) = False then
				'Setting the object to open the text file for appending the new data
				Set objTextStream = .OpenTextFile(hc_cases_acted_on, ForWriting, true)

				For excel_case = UBound(known_case_number_array, 2) to 0 Step -1
					If known_case_number_array(on_PND2_const, excel_case) = False Then
						objTextStream.WriteLine known_case_number_array(case_num_const, excel_case) & " -- " & known_case_number_array(excel_row_const, excel_case)
					End If
				Next

				objTextStream.Close
			End If
		End With

		removed_cases_count = 0
		For excel_case = UBound(known_case_number_array, 2) to 0 Step -1
			If known_case_number_array(on_PND2_const, excel_case) = False and known_case_number_array(on_assign_const, excel_case) <> True Then
				MAXIS_case_number = known_case_number_array(case_num_const, excel_case)
				excel_row = known_case_number_array(excel_row_const, excel_case)

				Set xmlTracDoc = CreateObject("Microsoft.XMLDOM")
				' xmlTracPath = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Assignments\Script Testing Logs\Interview Team Usage\interview_details_" & MAXIS_case_number & "_at_" & replace(replace(replace(now, "/", "_"),":", "_")," ", "_") & ".xml"
				xmlTracPath = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Assignments\ADS Health Care\Functional Data\Pending Ended Tracking\hc_pending_ended_" & MAXIS_case_number & "_at_" & replace(replace(replace(now, "/", "_"),":", "_")," ", "_") & ".xml"

				xmlTracDoc.async = False

				Set root = xmlTracDoc.createElement("HCPendSummary")
				xmlTracDoc.appendChild root

				Set element = xmlTracDoc.createElement("RemovedFromPendList")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(date)
				element.appendChild info

				Set element = xmlTracDoc.createElement("CaseloadNumber")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Caseload_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("CaseNumber")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Case_Number_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("CaseName")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Case_Name_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("APPLDate")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, APPL_Date_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("DaysPending")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Days_Pending_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("PendedDate")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Pended_Date_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("DateAddedToList")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Date_Added_to_List_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("MIPPA")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, MIPPA_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("METSTransition")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, METS_Transition_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("EMA")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, EMA_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("HCEvalDate")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, HC_Eval_Date_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("VerifsRequestedDate")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Verifs_Requested_Date_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("InitialAssignmentWorker")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Initial_Assignment_Worker_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("InitialAssignmentDate")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Initial_Assignment_Date_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("Day20")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Day_20_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("Day20Worker")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Day_20_Assignment_Worker_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("Day20Assignment")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Day_20_Assignment_Date_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("Day45")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Day_45_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("Day45Worker")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Day_45_Assignment_Worker_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("Day45Assignemnt")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Day_45_Assignment_Date_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("Day55")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Day_55_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("Day55Worker")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Day_55_Assignment_Worker_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("Day55Assignment")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Day_55_Assignment_Date_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("Day60")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Day_60_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("Day60Worker")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Day_60_Assignment_Worker_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("Day60Assignment")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Day_60_Assignment_Date_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("Overdue")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Overdue_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("NeedsAssignment")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Needs_Assignment_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("Priority")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Priority_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("MostRecentAssignmentWorker")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Most_Recent_Assignment_Worker_col).Value))
				element.appendChild info

				Set element = xmlTracDoc.createElement("MostRecentAssignmentDate")
				root.appendChild element
				Set info = xmlTracDoc.createTextNode(trim(ObjExcel.Cells(excel_row, Most_Recent_Assignment_Date_col).Value))
				element.appendChild info

				xmlTracDoc.save(xmlTracPath)

				Set xml = CreateObject("Msxml2.DOMDocument")
				Set xsl = CreateObject("Msxml2.DOMDocument")

				Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
				txt = Replace(fso.OpenTextFile(xmlTracPath).ReadAll, "><", ">" & vbCrLf & "<")
				stylesheet = "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & _
				"<xsl:output method=""xml"" indent=""yes""/>" & _
				"<xsl:template match=""/"">" & _
				"<xsl:copy-of select="".""/>" & _
				"</xsl:template>" & _
				"</xsl:stylesheet>"

				xsl.loadXML stylesheet
				xml.loadXML txt

				xml.transformNode xsl

				xml.Save xmlTracPath

				ObjExcel.Rows(excel_row).EntireRow.Delete
				removed_cases_count = removed_cases_count + 1
				excel_row_new_cases_start = excel_row_new_cases_start - 1
			End If
		Next

		objWorkbook.Save()		'saving the excel
	End If

	'This section adds the most recent case note information (date, x number and case note to the Excel list. The user will need to select this option in the checkbox on the dialog.)
	If evaluate_cases_in_MAXIS = True or capture_pnd2 = True Then
		excel_row = 2		'starting with row 2 (1st cell with case information)
		If evaluate_cases_in_MAXIS = False Then excel_row = excel_row_new_cases_start
		Do
			MAXIS_case_number = ObjExcel.Cells(excel_row, Case_Number_col).Value		'establishing what the case number is for each case
			appl_date = ObjExcel.Cells(excel_row, APPL_Date_col).Value		'establishing what the case number is for each case
			pended_date = trim(ObjExcel.Cells(excel_row, Pended_Date_col).Value)		'establishing what the case number is for each case
			If pended_date <> "" and IsDate(pended_date) = False Then
				pended_date = ""
				ObjExcel.Cells(excel_row, Days_Pending_col).Value = ""
			End If

			mippa_case = ""
			mets_trans_case = ""
			ema_case = ""
			smrt_case = ""
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, MIPPA_col), mippa_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, METS_Transition_col), mets_trans_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, EMA_col), ema_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, SMRT_Application_col), smrt_case)
			curr_pended_date		= ObjExcel.Cells(excel_row, Pended_Date_col).Value
			curr_hc_eval_note_date 	= ObjExcel.Cells(excel_row, HC_Eval_Date_col).Value
			curr_verif_note_date 	= ObjExcel.Cells(excel_row, Verifs_Requested_Date_col).Value

			If mippa_case = "" 		Then mippa_case = False
			If mets_trans_case = "" Then mets_trans_case = False
			If ema_case = "" 		Then ema_case = False
			If smrt_case = "" 		Then smrt_case = False

			too_old_date = DateAdd("D", -1, appl_date)              'We don't need to read notes from before the CAF date

			If MAXIS_case_number = "" then exit do						'leaves do if no case number is on the next Excel row

			Call navigate_to_MAXIS_screen("CASE", "NOTE")				'headin' over to CASE/NOTE
			EMWaitReady 0, 0
			EMReadScreen case_note_check, 17, 2, 33
			If case_note_check = "Case Notes (NOTE)" Then
				hc_eval_note_date = ""
				verif_note_date = ""

				' mippa_case = False
				' mets_trans_case = False
				smrt_case = False
				smrt_started = False
				smrt_ended = False

				note_row = 5
				Do
					EMReadScreen note_date, 	8, note_row, 6                  'reading the note date
					EMReadScreen note_worker, 	7, note_row, 16
					EMReadScreen note_title, 	55, note_row, 25               'reading the note header
					note_title = trim(note_title)

					If left(note_title, 32) = "~ HC Pended from a METS case for" Then
						pended_date = note_date
						mets_trans_case = True
					End If
					If left(note_title, 52) = "~ MIPPA/Extra Help request received via REPT/MLAR on" Then
						pended_date = note_date
						mippa_case = True
					End If
					If left(note_title, 31) = "---Initial SMRT referral reques" Then smrt_started = True
					If left(note_title, 31) = "---ISDS referral completed for " Then smrt_started = True
					If left(note_title, 31) = "---SMRT NOT submitted to ISDS--" Then smrt_ended = True
					If left(note_title, 31) = "---SMRT determination received:" Then smrt_ended = True
					If left(note_title, 31) = "---SMRT Determination Request W" Then smrt_ended = True
					If left(note_worker, 2) = "PW" and InStr(note_title, "SMRT") Then
						If InStr(UCASE(note_title), "APPROVED") Then smrt_started = True
						If InStr(UCASE(note_title), "DENIED") Then smrt_ended = True
						If InStr(UCASE(note_title), "DENY") Then smrt_ended = True
					End If

					If left(note_title, 36) = "~ Application Received (MNsure HCAPP" Then pended_date = note_date
					If left(note_title, 48) = "~ Application Received (HC - Certain Populations" Then pended_date = note_date
					If left(note_title, 33) = "~ Application Received (LTC HCAPP" Then pended_date = note_date
					If left(note_title, 44) = "~ Application Received (HCAPP for B/C Cancer" Then pended_date = note_date
					If left(note_title, 34) = "Subsequent Application Requesting:" and InStr(note_title, "HC") <> 0 Then pended_date = note_date

					ucase_note_title = UCase(note_title)
					If InStr(ucase_note_title, "VERIFICATIONS REQUESTED") <> 0 Then
						If verif_note_date = "" Then verif_note_date = note_date
					End If

					hc_eval_note_found = False
					If InStr(note_title, "HC Certain Pops App:") <> 0 			Then hc_eval_note_found = True
					If InStr(note_title, "MNSure HC App:") <> 0 				Then hc_eval_note_found = True
					If InStr(note_title, "HC Renewal Form:") <> 0 				Then hc_eval_note_found = True
					If InStr(note_title, "Combined AR:") <> 0 					Then hc_eval_note_found = True
					If InStr(note_title, "LTC HC App:") <> 0 					Then hc_eval_note_found = True
					If InStr(note_title, "HC Renewal Form for Families:") <> 0 	Then hc_eval_note_found = True
					If InStr(note_title, "LTC Renewal:") <> 0 					Then hc_eval_note_found = True
					If InStr(note_title, "MN Family Planning App:") <> 0 		Then hc_eval_note_found = True
					If hc_eval_note_found = True and hc_eval_note_date = "" 	Then hc_eval_note_date = note_date

					note_row = note_row + 1
					If note_row = 19 Then
						note_row = 5
						PF8
						EMReadScreen check_for_last_page, 9, 24, 14
						If check_for_last_page = "LAST PAGE" Then Exit Do
					End If
					EMReadScreen next_note_date, 8, note_row, 6
					If next_note_date = "        " Then Exit Do
				Loop until DateDiff("d", too_old_date, next_note_date) <= 0

				If smrt_ended = True Then smrt_case = False
				If smrt_started = True and smrt_ended = False Then smrt_case = True

				ObjExcel.Cells(excel_row, MIPPA_col).Value 					= mippa_case
				ObjExcel.Cells(excel_row, METS_Transition_col).Value 		= mets_trans_case
				ObjExcel.Cells(excel_row, SMRT_Application_col).Value 		= smrt_case

				If curr_pended_date = "" and pended_date <> "" 	Then ObjExcel.Cells(excel_row, Pended_Date_col).Value 			= pended_date
				If curr_hc_eval_note_date = "" 					Then ObjExcel.Cells(excel_row, HC_Eval_Date_col).Value 			= hc_eval_note_date
				If curr_verif_note_date = "" 					Then ObjExcel.Cells(excel_row, Verifs_Requested_Date_col).Value = verif_note_date

			End If
			' EMReadScreen case_note_info, 74 , 5, 6						'reads the most recent case note
			' If trim(case_note_info) <> "" then ObjExcel.Cells(excel_row, col_to_use).Value = case_note_info	'If it's not blank, then it writes the information into Excel

			pended_date = ObjExcel.Cells(excel_row, Pended_Date_col).Value
			If trim(pended_date) = "" Then
				Call back_to_SELF
				Call navigate_to_MAXIS_screen("STAT", "HCRE")
				EMReadScreen hcre_updated_date, 8, 21, 55
				hcre_updated_date = trim(hcre_updated_date)
				If hcre_updated_date <> "" Then ObjExcel.Cells(excel_row, Pended_Date_col).Value = replace(hcre_updated_date, " ", "/")
			End If

			excel_row = excel_row + 1									'moves Excel to next row
			Call back_to_SELF
		LOOP until MAXIS_case_number = ""								'Loops until all the case have been noted

		objWorkbook.Save()		'saving the excel
	End If


	case_to_assign_count = 0
	pri_1_case_count = 0
	pri_2_case_count = 0
	pri_3_case_count = 0
	pri_4_case_count = 0
	pri_5_case_count = 0
	pri_6_case_count = 0
	case_on_assign_count = 0
	If evaluate_assignments = True or capture_pnd2 = True Then
		excel_row = 2		'starting with row 2 (1st cell with case information)
		If evaluate_assignments = False Then excel_row = excel_row_new_cases_start

		' If cases_available_for_assignment > bottom_threshold Then excel_row = excel_row_new_cases_start
		Do
			caseload_number = 			ObjExcel.Cells(excel_row, Caseload_col).Value
			MAXIS_case_number = 		ObjExcel.Cells(excel_row, Case_Number_col).Value
			case_name = 				ObjExcel.Cells(excel_row, Case_Name_col).Value
			APPL_date = 				ObjExcel.Cells(excel_row, APPL_Date_col).Value
			Days_pending = 				ObjExcel.Cells(excel_row, Days_Pending_col).Value
			pended_date = 				ObjExcel.Cells(excel_row, Pended_Date_col).Value
			date_added_to_list = 		ObjExcel.Cells(excel_row, Date_Added_to_List_col).Value
			HC_Eval_date = 				ObjExcel.Cells(excel_row, HC_Eval_Date_col).Value
			Verif_requested_date = 		ObjExcel.Cells(excel_row, Verifs_Requested_Date_col).Value
			initial_assignment_worker = ObjExcel.Cells(excel_row, Initial_Assignment_Worker_col).Value
			initial_assignment_date = 	ObjExcel.Cells(excel_row, Initial_Assignment_Date_col).Value
			Day_20_date = 				ObjExcel.Cells(excel_row, Day_20_col).Value
			Day_20_assignment_worker = 	ObjExcel.Cells(excel_row, Day_20_Assignment_Worker_col).Value
			Day_20_assignment_date = 	ObjExcel.Cells(excel_row, Day_20_Assignment_Date_col).Value
			Day_45_date = 				ObjExcel.Cells(excel_row, Day_45_col).Value
			Day_45_assignment_worker = 	ObjExcel.Cells(excel_row, Day_45_Assignment_Worker_col).Value
			Day_45_assignment_date = 	ObjExcel.Cells(excel_row, Day_45_Assignment_Date_col).Value
			Day_55_date = 				ObjExcel.Cells(excel_row, Day_55_col).Value
			Day_55_assignment_worker = 	ObjExcel.Cells(excel_row, Day_55_Assignment_Worker_col).Value
			Day_55_assignment_date = 	ObjExcel.Cells(excel_row, Day_55_Assignment_Date_col).Value
			Day_60_date = 				ObjExcel.Cells(excel_row, Day_60_col).Value
			Day_60_assignment_worker = 	ObjExcel.Cells(excel_row, Day_60_Assignment_Worker_col).Value
			Day_60_assignment_date = 	ObjExcel.Cells(excel_row, Day_60_Assignment_Date_col).Value
			Last_Assingment_Date = 		ObjExcel.Cells(excel_row, Most_Recent_Assignment_Date_col).Value
			' case_priority = 			ObjExcel.Cells(excel_row, Priority_col).Value
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, MIPPA_col), mippa_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, METS_Transition_col), mets_trans_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, EMA_col), ema_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, SMRT_Application_col), smrt_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, Overdue_col), case_overdue)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, Currently_Assigned_col), on_assignment)


			worked_in_past_week = False
			If IsDate(Last_Assingment_Date) = True Then
				Last_Assingment_Date = DateAdd("d", 0, Last_Assingment_Date)
				days_since_last_work = DateDiff("d", Last_Assingment_Date, date)
				If days_since_last_work < 8 Then worked_in_past_week = True
			End If
			If on_assignment = True Then case_on_assign_count = case_on_assign_count + 1

			If on_assignment <> True and worked_in_past_week = False Then
				If pended_date <> "" and IsDate(pended_date) = False Then
					pended_date = ""
					ObjExcel.Cells(excel_row, Days_Pending_col).Value = ""
				End If

				If APPL_date <> "" 					Then APPL_date = 				DateAdd("d", 0, APPL_date)
				If pended_date <> "" 				Then pended_date = 				DateAdd("d", 0, pended_date)
				If date_added_to_list <> "" 		Then date_added_to_list = 		DateAdd("d", 0, date_added_to_list)
				If HC_Eval_date <> "" 				Then HC_Eval_date = 			DateAdd("d", 0, HC_Eval_date)
				If Verif_requested_date <> "" 		Then Verif_requested_date = 	DateAdd("d", 0, Verif_requested_date)
				If initial_assignment_date <> "" 	Then initial_assignment_date = 	DateAdd("d", 0, initial_assignment_date)
				If Day_20_date <> "" 				Then Day_20_date = 				DateAdd("d", 0, Day_20_date)
				If Day_20_assignment_date <> "" 	Then Day_20_assignment_date = 	DateAdd("d", 0, Day_20_assignment_date)
				If Day_45_date <> "" 				Then Day_45_date = 				DateAdd("d", 0, Day_45_date)
				If Day_45_assignment_date <> "" 	Then Day_45_assignment_date = 	DateAdd("d", 0, Day_45_assignment_date)
				If Day_55_date <> "" 				Then Day_55_date = 				DateAdd("d", 0, Day_55_date)
				If Day_55_assignment_date <> "" 	Then Day_55_assignment_date = 	DateAdd("d", 0, Day_55_assignment_date)
				If Day_60_date <> "" 				Then Day_60_date = 				DateAdd("d", 0, Day_60_date)
				If Day_60_assignment_date <> "" 	Then Day_60_assignment_date = 	DateAdd("d", 0, Day_60_assignment_date)
				If Last_Assingment_Date <> "" 		Then Last_Assingment_Date = 	DateAdd("d", 0, Last_Assingment_Date)

				Days_pending = Days_pending * 1
				' MsgBox "MIPPA_col - " & MIPPA_col & vbCr & "METS_Transition_col - " & METS_Transition_col & vbCr & "excel_row - " & excel_row
				'TEST on FRI didn't work
				If mippa_case = False and mets_trans_case = False Then 'and ema_case = False Then
					days_since_last_assignment = 5000
					If IsDate(Last_Assingment_Date) = True Then
						days_since_last_assignment = DateDiff("d", Last_Assingment_Date, date)
					End If

					case_needs_assignment = False
					case_priority = ""
					verifs_are_due = False
					If IsDate(Verif_requested_date) = True Then
						If DateDiff("d", Verif_requested_date, date) >= 10 Then verifs_are_due = True
					End If
					If case_overdue = True and verifs_are_due = True and smrt_case = False Then case_priority = 1
					If case_overdue = True and IsDate(Verif_requested_date) = False and smrt_case = False Then case_priority = 1
					If HC_Eval_date = "" and case_priority = "" Then case_priority = 2
					If case_priority = "" and days_since_last_assignment > 10 Then
						If IsDate(Day_20_date) = True then
							diff_day_20 = ABS(DateDiff("d", date, Day_20_date))
							If diff_day_20 < 4 and Day_20_assignment_worker = "" Then case_priority = 3
						End If
						If IsDate(Day_45_date) = True then
							diff_day_45 = ABS(DateDiff("d", date, Day_45_date))
							If diff_day_45 < 4 and Day_45_assignment_worker = "" Then case_priority = 4
						End If
						If IsDate(Day_55_date) = True then
							diff_day_55 = ABS(DateDiff("d", date, Day_55_date))
							If diff_day_55 < 4 and Day_55_assignment_worker = "" Then case_priority = 5
						End If
					End If
					If case_priority = "" Then
						If IsDate(Day_60_date) = True then
							diff_day_60 = ABS(DateDiff("d", date, Day_60_date))
							If diff_day_60 < 4 and Day_60_assignment_worker = "" Then case_priority = 6
						End If
					End If

					If IsNumeric(case_priority) = True Then
						case_needs_assignment = True
						case_to_assign_count = case_to_assign_count + 1
						If case_priority = 1 Then pri_1_case_count = pri_1_case_count + 1
						If case_priority = 2 Then pri_2_case_count = pri_2_case_count + 1
						If case_priority = 3 Then pri_3_case_count = pri_3_case_count + 1
						If case_priority = 4 Then pri_4_case_count = pri_4_case_count + 1
						If case_priority = 5 Then pri_5_case_count = pri_5_case_count + 1
						If case_priority = 6 Then pri_6_case_count = pri_6_case_count + 1
					End If
					ObjExcel.Cells(excel_row, Priority_col) = case_priority
					ObjExcel.Cells(excel_row, Needs_Assignment_col) = case_needs_assignment
					ObjExcel.Cells(excel_row, Currently_Assigned_col).Value = False


				End If

				' MsgBox "excel_row - " & excel_row & vbCr & "case_priority - " & case_priority & vbCr & "case_needs_assignment - " & case_needs_assignment & vbCr & "verifs_are_due - " & verifs_are_due
			End If
			excel_row = excel_row + 1									'moves Excel to next row
			next_MAXIS_case_number = trim(ObjExcel.Cells(excel_row, Case_Number_col).Value)
		LOOP until next_MAXIS_case_number = ""								'Loops until all the case have been noted
		objWorkbook.Save()		'saving the excel
	End If

	capture_snapshot = False
	If data_snapshot_checkbox = checked Then capture_snapshot = True
	total_pnd_case_count 	= 0
	mippa_case_count 		= 0
	mets_trans_case_count 	= 0
	ema_case_count 			= 0
	smrt_case_count 		= 0
	smrt_60_plus_count 		= 0
	day_1_10_count 			= 0
	day_11_20_count 		= 0
	day_21_30_count 		= 0
	day_31_40_count 		= 0
	day_41_50_count 		= 0
	day_51_60_count 		= 0
	day_61_90_count 		= 0
	day_91_plus_count 		= 0
	day_1_20_count 			= 0
	day_21_45_count 		= 0
	day_46_55_count 		= 0
	day_56_60_count 		= 0
	overdue_count 			= 0

	If capture_snapshot = True Then
		excel_row = 2
		Do
			mippa_case = ""
			mets_trans_case = ""
			ema_case = ""
			smrt_case = ""
			case_overdue = ""
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, MIPPA_col), mippa_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, METS_Transition_col), mets_trans_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, EMA_col), ema_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, SMRT_Application_col), smrt_case)
			Call read_boolean_from_excel(ObjExcel.Cells(excel_row, Overdue_col), case_overdue)

			total_pnd_case_count 	= total_pnd_case_count + 1
			If mippa_case = True Then mippa_case_count 				= mippa_case_count + 1
			If mets_trans_case = True Then mets_trans_case_count 	= mets_trans_case_count + 1
			If ema_case = True Then ema_case_count 					= ema_case_count + 1
			If case_overdue = True Then overdue_count 				= overdue_count + 1
			If smrt_case = True Then smrt_case_count 				= smrt_case_count + 1
			If ObjExcel.Cells(excel_row, Days_Pending_col).Value > 60 and smrt_case = True Then smrt_60_plus_count 		= smrt_60_plus_count + 1
			If ObjExcel.Cells(excel_row, Days_Pending_col).Value < 11 Then day_1_10_count 			= day_1_10_count + 1
			If ObjExcel.Cells(excel_row, Days_Pending_col).Value < 21 Then day_1_20_count 			= day_1_20_count + 1
			If ObjExcel.Cells(excel_row, Days_Pending_col).Value > 90 Then day_91_plus_count 		= day_91_plus_count + 1
			If ObjExcel.Cells(excel_row, Days_Pending_col).Value > 10 and ObjExcel.Cells(excel_row, Days_Pending_col).Value < 21 Then day_11_20_count 		= day_11_20_count + 1
			If ObjExcel.Cells(excel_row, Days_Pending_col).Value > 20 and ObjExcel.Cells(excel_row, Days_Pending_col).Value < 31 Then day_21_30_count 		= day_21_30_count + 1
			If ObjExcel.Cells(excel_row, Days_Pending_col).Value > 30 and ObjExcel.Cells(excel_row, Days_Pending_col).Value < 41 Then day_31_40_count 		= day_31_40_count + 1
			If ObjExcel.Cells(excel_row, Days_Pending_col).Value > 40 and ObjExcel.Cells(excel_row, Days_Pending_col).Value < 51 Then day_41_50_count 		= day_41_50_count + 1
			If ObjExcel.Cells(excel_row, Days_Pending_col).Value > 50 and ObjExcel.Cells(excel_row, Days_Pending_col).Value < 61 Then day_51_60_count 		= day_51_60_count + 1
			If ObjExcel.Cells(excel_row, Days_Pending_col).Value > 60 and ObjExcel.Cells(excel_row, Days_Pending_col).Value < 91 Then day_61_90_count 		= day_61_90_count + 1
			If ObjExcel.Cells(excel_row, Days_Pending_col).Value > 20 and ObjExcel.Cells(excel_row, Days_Pending_col).Value < 46 Then day_21_45_count 		= day_21_45_count + 1
			If ObjExcel.Cells(excel_row, Days_Pending_col).Value > 45 and ObjExcel.Cells(excel_row, Days_Pending_col).Value < 56 Then day_46_55_count 		= day_46_55_count + 1
			If ObjExcel.Cells(excel_row, Days_Pending_col).Value > 55 and ObjExcel.Cells(excel_row, Days_Pending_col).Value < 61 Then day_56_60_count 		= day_56_60_count + 1
			excel_row = excel_row + 1									'moves Excel to next row
			next_MAXIS_case_number = ObjExcel.Cells(excel_row, Case_Number_col).Value		'establishing what the case number is for each case
		LOOP until next_MAXIS_case_number = ""								'Loops until all the case have been noted
	End If

	'TESTING - Commented out for testing
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit
	objExcel.Quit

	Call release_data_lock("MAIN")

	'TXT cookie that lets us know when the last update happened.
	With (CreateObject("Scripting.FileSystemObject"))
		If .FileExists(pending_hc_update_cookie) = True then
			.DeleteFile(pending_hc_update_cookie)
		End If

		'If the file doesn't exist, it needs to create it here and initialize it here! After this, it can just exit as the file will now be initialized
		If .FileExists(pending_hc_update_cookie) = False then
			'Setting the object to open the text file for appending the new data
			Set objTextStream = .OpenTextFile(pending_hc_update_cookie, ForWriting, true)
			objTextStream.WriteLine "update_date&*^&*^" & date
			objTextStream.WriteLine "update_time&*^&*^" & FormatDateTime(time,4) ' time

			objTextStream.Close
		End If
	End With

	' RECORD SNAPSHOT OF DATA
	const SNPSHT_Data_Collected_Date_COL 	= 01
	const SNPSHT_Total_Pending_Cases_COL 	= 02
	const SNPSHT_MIPPA_Count_COL 			= 03
	const SNPSHT_METS_Transition_Count_COL 	= 04
	const SNPSHT_EMA_Count_COL 				= 05
	const SNPSHT_SMRT_Pending_Count_COL 	= 06
	const SNPSHT_SMRT_over_60_Days_COL 		= 07
	const SNPSHT_Days_1_10_COL 				= 08
	const SNPSHT_Days_11_20_COL 			= 09
	const SNPSHT_Days_21_30_COL 			= 10
	const SNPSHT_Days_31_40_COL 			= 11
	const SNPSHT_Days_41_50_COL 			= 12
	const SNPSHT_Days_51_60_COL 			= 13
	const SNPSHT_Days_61_90_COL 			= 14
	const SNPSHT_Days_90_COL 				= 15
	const SNPSHT_Days_1_20_COL 				= 16
	const SNPSHT_Days_1_20_Percent_COL 		= 17
	const SNPSHT_Days_21_45_COL 			= 18
	const SNPSHT_Days_21_45_Percent_COL 	= 19
	const SNPSHT_Days_46_55_COL 			= 20
	const SNPSHT_Days_46_55_Percent_COL 	= 21
	const SNPSHT_Days_56_60_COL 			= 22
	const SNPSHT_Days_56_60_Percent_COL 	= 23
	const SNPSHT_Overdue_COL 				= 24
	const SNPSHT_Overdue_Percent_COL 		= 25


	If capture_snapshot = True Then
		Call excel_open(snapshot_hc_pending_excel, True, False, ObjExcel, objWorkbook)
		excel_row = 1
		Do
			excel_row = excel_row + 1
			listed_date = trim(ObjExcel.Cells(excel_row, SNPSHT_Data_Collected_Date_COL).Value)
		Loop until listed_date = ""

		ObjExcel.Cells(excel_row, SNPSHT_Data_Collected_Date_COL).Value	 	= date
		ObjExcel.Cells(excel_row, SNPSHT_Total_Pending_Cases_COL).Value	 	= total_pnd_case_count
		ObjExcel.Cells(excel_row, SNPSHT_MIPPA_Count_COL).Value	 			= mippa_case_count
		ObjExcel.Cells(excel_row, SNPSHT_METS_Transition_Count_COL).Value	= mets_trans_case_count
		ObjExcel.Cells(excel_row, SNPSHT_EMA_Count_COL).Value	 			= ema_case_count
		ObjExcel.Cells(excel_row, SNPSHT_SMRT_Pending_Count_COL).Value	 	= smrt_case_count
		ObjExcel.Cells(excel_row, SNPSHT_SMRT_over_60_Days_COL).Value	 	= smrt_60_plus_count
		ObjExcel.Cells(excel_row, SNPSHT_Days_1_10_COL).Value	 			= day_1_10_count
		ObjExcel.Cells(excel_row, SNPSHT_Days_11_20_COL).Value	 			= day_11_20_count
		ObjExcel.Cells(excel_row, SNPSHT_Days_21_30_COL).Value	 			= day_21_30_count
		ObjExcel.Cells(excel_row, SNPSHT_Days_31_40_COL).Value	 			= day_31_40_count
		ObjExcel.Cells(excel_row, SNPSHT_Days_41_50_COL).Value	 			= day_41_50_count
		ObjExcel.Cells(excel_row, SNPSHT_Days_51_60_COL).Value	 			= day_51_60_count
		ObjExcel.Cells(excel_row, SNPSHT_Days_61_90_COL).Value	 			= day_61_90_count
		ObjExcel.Cells(excel_row, SNPSHT_Days_90_COL).Value	 				= day_91_plus_count
		ObjExcel.Cells(excel_row, SNPSHT_Days_1_20_COL).Value	 			= day_1_20_count
		' ObjExcel.Cells(excel_row, SNPSHT_Days_1_20_Percent_COL).Value	 	=
		ObjExcel.Cells(excel_row, SNPSHT_Days_21_45_COL).Value	 			= day_21_45_count
		' ObjExcel.Cells(excel_row, SNPSHT_Days_21_45_Percent_COL).Value	 	=
		ObjExcel.Cells(excel_row, SNPSHT_Days_46_55_COL).Value	 			= day_46_55_count
		' ObjExcel.Cells(excel_row, SNPSHT_Days_46_55_Percent_COL).Value	 	=
		ObjExcel.Cells(excel_row, SNPSHT_Days_56_60_COL).Value	 			= day_56_60_count
		' ObjExcel.Cells(excel_row, SNPSHT_Days_56_60_Percent_COL).Value	 	=
		ObjExcel.Cells(excel_row, SNPSHT_Overdue_COL).Value	 				= overdue_count
		' ObjExcel.Cells(excel_row, SNPSHT_Overdue_Percent_COL).Value	 		=
		objWorkbook.Save()		'saving the excel
		objExcel.ActiveWorkbook.Close
		objExcel.Application.Quit
		objExcel.Quit
	End If

	hc_pend_worklist_run_time = timer-worklist_start_time
	hc_pend_worklist_run_min = int(hc_pend_worklist_run_time/60)
	hc_pend_worklist_run_sec = hc_pend_worklist_run_time MOD 60
	end_msg = "Health Care pending Report updated with current information." & vbCr & vbCr

	If capture_pnd2 = True Then end_msg = end_msg & "Total number of HC Pending Cases: " & current_pnd2_count & "." & vbCr
	If remove_no_longer_pending_cases = True Then end_msg = end_msg & "Number of cases no longer pending and removed from the list: " & removed_cases_count & "." & vbCr
	If evaluate_assignments = True Then
		end_msg = end_msg & "Total cases available for assignment:" & case_to_assign_count & "." & vbCr
		end_msg = end_msg & "Cases at priority 1: " & pri_1_case_count & "." & vbCr
		end_msg = end_msg & "Cases at priority 2: " & pri_2_case_count & "." & vbCr
		end_msg = end_msg & "Cases at priority 3: " & pri_3_case_count & "." & vbCr
		end_msg = end_msg & "Cases at priority 4: " & pri_4_case_count & "." & vbCr
		end_msg = end_msg & "Cases at priority 5: " & pri_5_case_count & "." & vbCr
		end_msg = end_msg & "Cases at priority 6: " & pri_6_case_count & "." & vbCr & vbCr
		end_msg = end_msg & "Cases currently assigned: " & case_on_assign_count & vbCr
	End If
	end_msg = end_msg & "Script run time: " & hc_pend_worklist_run_min & " minutes " & hc_pend_worklist_run_sec & " seconds."
End If


'PULL CASES TO WORK ----------------------------------------------------------------

If run_assignment_selection = True Then
	requested_case_count = "10"
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 221, 95, "HC Cases Worklist Creation"
		EditBox 170, 45, 40, 15, requested_case_count
		ButtonGroup ButtonPressed
			OkButton 105, 70, 50, 15
			CancelButton 160, 70, 50, 15
		Text 10, 10, 90, 10, "The script will:"
		Text 20, 20, 150, 10, "- Record your previous worklist information."
		Text 20, 30, 150, 10, "- Create a new worklist."
		Text 10, 50, 160, 10, "How many cases do you want on your worklist?"
	EndDialog

	'Dialog asks what stats are being pulled
	Do
		err_msg = ""

		Dialog Dialog1
		cancel_without_confirmation

		If IsNumeric(requested_case_count) = False Then err_msg = err_msg & vbCr & "* Enter a valid number for the number of cases to put on the worklist"
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
	Loop until err_mag = ""
	requested_case_count = requested_case_count * 1

	'Finding the user name - we aren't using the function because we need the comma in place
	'Creating objects for Access
	Set objConnection = CreateObject("ADODB.Connection")
	Set objRecordSet = CreateObject("ADODB.Recordset")

	SQL_table = "SELECT * from ES.V_ESAllStaff WHERE EmpLogOnID = '" & windows_user_ID & "'"				'identifying the table that stores the ES Staff user information

	'This is the file path the data tables
	objConnection.Open "Provider = SQLOLEDB.1;Data Source= " & "" &  "hssqlpw139;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;" & ""
	objRecordSet.Open SQL_table, objConnection							'Here we connect to the data tables

	Do While NOT objRecordSet.Eof										'now we will loop through each item listed in the table of ES Staff
		table_user_id = objRecordSet("EmpLogOnID")						'setting the user ID from table data
		If table_user_id = windows_user_ID Then							'If the ID on thils loop of the data information matches the ID of the person running the script, we have found the staff person
			worker_name = objRecordSet("EmpFullName")				'Save the user name
			Exit Do														'if we have found the person, we stop looping
		End If
		objRecordSet.MoveNext											'Going to the next row in the table
	Loop

	'Now we disconnect from the table and close the connections
	objRecordSet.Close
	objConnection.Close
	Set objRecordSet=nothing
	Set objConnection=nothing

	worker_name = trim(worker_name)
	name_array = split(worker_name, ",")
	last_name = trim(name_array(0))
	first_name =  trim(name_array(1))
	If InStr(first_name, " ") Then
		first_name_array = split(first_name)
		first_name = first_name_array(0)
	End If
	indv_worklist_file_name = first_name & " " & left(last_name, 1) & " Assignment.xlsx"
	indv_worklist_file_path = t_drive & "\Eligibility Support\Assignments\ADS Health Care\" & indv_worklist_file_name
	indv_worklist_template_file_path = t_drive & "\Eligibility Support\Assignments\ADS Health Care\Functional Data\Worker Assignment Template.xlsx"

	'check to see if the current pending list is 'locked' (maybe use a locking cookie) - if locked - end the script and tell the worker they need to wait.
	With (CreateObject("Scripting.FileSystemObject"))
		If .FileExists(lock_main_list) = True then script_end_procedure("HC Pending details are being updated. Try again in a little while.")

		list_being_viewed = .FileExists(hold_main_list)
		If list_being_viewed = True Then MsgBox "Another worker is pulling an assignment. The script will pause while this completes. It usually takes less than a minute to become available. Please wait."
		Do while list_being_viewed = True
			' WScript.Sleep 200
			EMWaitReady 0, 1000
			list_being_viewed = .FileExists(hold_main_list)
		Loop
	End With
	Call create_data_lock("HOLD")

	'TODO - maybe have BT create a sharepoint file

	'WORKER LISTS column definitions
	const wrkr_assign_hsr_col 	= 01 		'HSR Assignment
	const wrkr_assign_date_col 	= 02 		'Assignment Date
	const wrkr_case_numb_col 	= 03 		'Case Number
	const wrkr_case_name_col 	= 04 		'Name
	const wrkr_appl_date_col 	= 05 		'APPL Date
	const wrkr_days_pend_col 	= 06 		'Days Pending
	const wrkr_population_col 	= 07 		'Population
	const wrkr_hc_eval_date_col = 08 		'HC Evaluation Run
	const wrkr_verifs_date_col 	= 09 		'Verifs Requested
	const wrkr_day_60_col 		= 10 		'Day 60
	const wrkr_assign_compl_col = 11 		'Assignment Completed
	const wrkr_case_stat_col 	= 12 		'Approved, Denied, Pending
	const wrkr_specialty_col 	= 13 		'SMRT / MIPPA / MAEPD / EMA / METS Transition / Etc
	const wrkr_deny_date_col 	= 14 		'Date denial can be acted on
	const wrkr_notes_col 		= 15 		'Notes

	const end_const = 20

	Dim COMP_ASSIGN_ARRAY()
	ReDim COMP_ASSIGN_ARRAY(end_const, 0)
	work_counter = 0

	'Update the current pending cases log with assignment information and release the case for reassignment if needed
	Call excel_open(controller_hc_pending_excel, False, False, ObjExcel, objWorkbook)
	objExcel.worksheets("Cases").Activate			'Activates the selected worksheet'

	'If yes:
	If objFSO.FileExists(indv_worklist_file_path) Then
		'Open the sheet
		Call excel_open(indv_worklist_file_path, False, False, ObjWrkrExcel, objWrkrWorkbook)

		'Read information from the sheet
		excel_row = 2
		Do
			'Add to an array
			ReDim preserve COMP_ASSIGN_ARRAY(end_const, work_counter)
			COMP_ASSIGN_ARRAY(wrkr_assign_hsr_col, 		work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_assign_hsr_col).Value) 		 		'HSR Assignment
			COMP_ASSIGN_ARRAY(wrkr_assign_date_col, 	work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_assign_date_col).Value) 		 		'Assignment Date
			COMP_ASSIGN_ARRAY(wrkr_case_numb_col, 		work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_case_numb_col).Value) 		 		'Case Number
			COMP_ASSIGN_ARRAY(wrkr_case_name_col, 		work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_case_name_col).Value) 		 		'Name
			COMP_ASSIGN_ARRAY(wrkr_appl_date_col, 		work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_appl_date_col).Value) 		 		'APPL Date
			COMP_ASSIGN_ARRAY(wrkr_days_pend_col, 		work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_days_pend_col).Value) 		 		'Days Pending
			COMP_ASSIGN_ARRAY(wrkr_population_col, 		work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_population_col).Value) 		 		'Population
			COMP_ASSIGN_ARRAY(wrkr_hc_eval_date_col, 	work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_hc_eval_date_col).Value) 		 		'HC Evaluation Run
			COMP_ASSIGN_ARRAY(wrkr_verifs_date_col, 	work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_verifs_date_col).Value) 		 		'Verifs Requested
			COMP_ASSIGN_ARRAY(wrkr_day_60_col, 			work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_day_60_col).Value) 		 		'Day 60
			COMP_ASSIGN_ARRAY(wrkr_assign_compl_col, 	work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_assign_compl_col).Value) 		 		'Assignment Completed
			COMP_ASSIGN_ARRAY(wrkr_case_stat_col, 		work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_case_stat_col).Value) 		 		'Approved, Denied, Pending
			COMP_ASSIGN_ARRAY(wrkr_specialty_col, 		work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_specialty_col).Value) 		 		'SMRT / MIPPA / MAEPD / EMA / METS Transition / Etc
			COMP_ASSIGN_ARRAY(wrkr_deny_date_col, 		work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_deny_date_col).Value) 		 		'Date denial can be acted on
			COMP_ASSIGN_ARRAY(wrkr_notes_col, 			work_counter) = trim(ObjWrkrExcel.Cells(excel_row, wrkr_notes_col).Value) 		 		'Notes


			'Save The Assignment Details for record keeping
			Set xmlAssignDoc = CreateObject("Microsoft.XMLDOM")
			xmlAssignPath = "\\hcgg.fr.co.hennepin.mn.us\lobroot\hsph\team\Eligibility Support\Assignments\ADS Health Care\Functional Data\Completed Reviews\hc_pending_assignment_" & COMP_ASSIGN_ARRAY(wrkr_case_numb_col, work_counter) & "_on_" & replace(replace(replace(date, "/", "_"),":", "_")," ", "_") & ".xml"

			xmlAssignDoc.async = False

			Set root = xmlAssignDoc.createElement("HCPendAssignment")
			xmlAssignDoc.appendChild root


			Set element = xmlAssignDoc.createElement("AssignedHSR")
			root.appendChild element
			Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_assign_hsr_col, work_counter))
			element.appendChild info

			Set element = xmlAssignDoc.createElement("AssignedDate")
			root.appendChild element
			Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_assign_date_col, work_counter))
			element.appendChild info

			Set element = xmlAssignDoc.createElement("CaseNumber")
			root.appendChild element
			Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_case_numb_col, work_counter))
			element.appendChild info

			Set element = xmlAssignDoc.createElement("CaseName")
			root.appendChild element
			Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_case_name_col, work_counter))
			element.appendChild info

			Set element = xmlAssignDoc.createElement("APPLDate")
			root.appendChild element
			Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_appl_date_col, work_counter))
			element.appendChild info

			Set element = xmlAssignDoc.createElement("DaysPending")
			root.appendChild element
			Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_days_pend_col, work_counter))
			element.appendChild info

			Set element = xmlAssignDoc.createElement("Population")
			root.appendChild element
			Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_population_col, work_counter))
			element.appendChild info

			Set element = xmlAssignDoc.createElement("HCEvalRunDate")
			root.appendChild element
			Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_hc_eval_date_col, work_counter))
			element.appendChild info

			Set element = xmlAssignDoc.createElement("VerifsDate")
			root.appendChild element
			Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_verifs_date_col, work_counter))
			element.appendChild info

			Set element = xmlAssignDoc.createElement("DaySixty")
			root.appendChild element
			Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_day_60_col, work_counter))
			element.appendChild info

			Set element = xmlAssignDoc.createElement("AssignmentCompleted")
			root.appendChild element
			If IsDate(COMP_ASSIGN_ARRAY(wrkr_assign_compl_col, work_counter)) = True Then
				Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_assign_compl_col, work_counter))
			Else
				If COMP_ASSIGN_ARRAY(wrkr_assign_compl_col, work_counter) <> "" Then Set info = xmlAssignDoc.createTextNode("INCOMPELTE??? - " & COMP_ASSIGN_ARRAY(wrkr_assign_compl_col, work_counter))
				If COMP_ASSIGN_ARRAY(wrkr_assign_compl_col, work_counter) = "" Then Set info = xmlAssignDoc.createTextNode("INCOMPELTE???")
			End If
			element.appendChild info

			Set element = xmlAssignDoc.createElement("CaseStatus")
			root.appendChild element
			Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_case_stat_col, work_counter))
			element.appendChild info

			Set element = xmlAssignDoc.createElement("Specialty")
			root.appendChild element
			Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_specialty_col, work_counter))
			element.appendChild info

			Set element = xmlAssignDoc.createElement("DateToDeny")
			root.appendChild element
			Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_deny_date_col, work_counter))
			element.appendChild info

			Set element = xmlAssignDoc.createElement("Notes")
			root.appendChild element
			Set info = xmlAssignDoc.createTextNode(COMP_ASSIGN_ARRAY(wrkr_notes_col, work_counter))
			element.appendChild info

			xmlAssignDoc.save(xmlAssignPath)

			Set xml = CreateObject("Msxml2.DOMDocument")
			Set xsl = CreateObject("Msxml2.DOMDocument")

			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			txt = Replace(fso.OpenTextFile(xmlAssignPath).ReadAll, "><", ">" & vbCrLf & "<")
			stylesheet = "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & _
			"<xsl:output method=""xml"" indent=""yes""/>" & _
			"<xsl:template match=""/"">" & _
			"<xsl:copy-of select="".""/>" & _
			"</xsl:template>" & _
			"</xsl:stylesheet>"

			xsl.loadXML stylesheet
			xml.loadXML txt

			xml.transformNode xsl

			xml.Save xmlAssignPath

			excel_row = excel_row + 1
			work_counter = work_counter + 1
			next_case_numb = trim(ObjWrkrExcel.Cells(excel_row, wrkr_assign_hsr_col).Value)
		Loop until next_case_numb = ""

		rows_to_skip_string = "~"
		For wrkr_cases = 0 to UBound(COMP_ASSIGN_ARRAY, 2)
			excel_row = 2
			Do
				If trim(ObjExcel.Cells(excel_row, Case_Number_col).Value) = COMP_ASSIGN_ARRAY(wrkr_case_numb_col, wrkr_cases) Then
					rows_to_skip_string = rows_to_skip_string & excel_row & "~"
					ObjExcel.Cells(excel_row, Currently_Assigned_col).Value 	= False
					ObjExcel.Cells(excel_row, HC_Eval_Date_col).Value 			= COMP_ASSIGN_ARRAY(wrkr_hc_eval_date_col, wrkr_cases)
					ObjExcel.Cells(excel_row, Verifs_Requested_Date_col).Value 	= COMP_ASSIGN_ARRAY(wrkr_verifs_date_col, wrkr_cases)
					If IsDate(COMP_ASSIGN_ARRAY(wrkr_assign_compl_col, wrkr_cases)) = True Then
						ObjExcel.Cells(excel_row, Needs_Assignment_col).Value 		= False
						' ObjExcel.Cells(excel_row, ).Value = COMP_ASSIGN_ARRAY(wrkr_case_numb_col, wrkr_cases)

						pnd_case_priority = trim(ObjExcel.Cells(excel_row, Priority_col).Value)
						If pnd_case_priority <> "" Then pnd_case_priority = pnd_case_priority * 1
						If pnd_case_priority = 1 or pnd_case_priority = 6 Then
							ObjExcel.Cells(excel_row, Day_60_Assignment_Worker_col).Value 			= COMP_ASSIGN_ARRAY(wrkr_assign_hsr_col, wrkr_cases)
							ObjExcel.Cells(excel_row, Day_60_Assignment_Date_col).Value 			= COMP_ASSIGN_ARRAY(wrkr_assign_date_col, wrkr_cases)
						ElseIf pnd_case_priority = 2 Then
							ObjExcel.Cells(excel_row, Initial_Assignment_Worker_col).Value 			= COMP_ASSIGN_ARRAY(wrkr_assign_hsr_col, wrkr_cases)
							ObjExcel.Cells(excel_row, Initial_Assignment_Date_col).Value 			= COMP_ASSIGN_ARRAY(wrkr_assign_date_col, wrkr_cases)
						ElseIf pnd_case_priority = 3 Then
							ObjExcel.Cells(excel_row, Day_20_Assignment_Worker_col).Value 			= COMP_ASSIGN_ARRAY(wrkr_assign_hsr_col, wrkr_cases)
							ObjExcel.Cells(excel_row, Day_20_Assignment_Date_col).Value 			= COMP_ASSIGN_ARRAY(wrkr_assign_date_col, wrkr_cases)
						ElseIf pnd_case_priority = 4 Then
							ObjExcel.Cells(excel_row, Day_45_Assignment_Worker_col).Value 			= COMP_ASSIGN_ARRAY(wrkr_assign_hsr_col, wrkr_cases)
							ObjExcel.Cells(excel_row, Day_45_Assignment_Date_col).Value 			= COMP_ASSIGN_ARRAY(wrkr_assign_date_col, wrkr_cases)
						ElseIf pnd_case_priority = 5 Then
							ObjExcel.Cells(excel_row, Day_55_Assignment_Worker_col).Value 			= COMP_ASSIGN_ARRAY(wrkr_assign_hsr_col, wrkr_cases)
							ObjExcel.Cells(excel_row, Day_55_Assignment_Date_col).Value 			= COMP_ASSIGN_ARRAY(wrkr_assign_date_col, wrkr_cases)
						End If
					Else
						ObjExcel.Cells(excel_row, Needs_Assignment_col).Value 		= True
					End If
					Exit Do
				End If
				excel_row = excel_row + 1
				next_case_numb = trim(ObjExcel.Cells(excel_row, Case_Number_col).Value)
			Loop Until next_case_numb = ""
		Next

		ObjWrkrExcel.ActiveWorkbook.Close
		ObjWrkrExcel.Application.Quit
		ObjWrkrExcel.Quit
	End If

	If operation_selection = "Complete Individual Worklist" Then

		With (CreateObject("Scripting.FileSystemObject"))
			If .FileExists(indv_worklist_file_path) = True then
				.DeleteFile(indv_worklist_file_path)
			End If
		End With

		end_msg = "Worklist details logged and worklist deleted." & vbCr & vbCr & "No new worklist created, run the script again to create a worklist when you are ready for more pending work."

	Else
		'Open a template and save as the worker's worklist
		Call excel_open(indv_worklist_template_file_path, False, False, ObjWrkrExcel, objWrkrWorkbook)

		'Select the number of cases requested based on priority and count
		'add the worker info to the HC Pending to 'assign' the case
		'Add the case information to the worker's worklist
		wrkr_excel_row = 2
		selected_case_count = 0
		For priority_select = 1 to 6
			full_excel_row = 2
			Do
				pnd_case_priority = trim(ObjExcel.Cells(full_excel_row, Priority_col).Value)
				If pnd_case_priority <> "" Then pnd_case_priority = pnd_case_priority * 1
				call read_boolean_from_excel(ObjExcel.Cells(full_excel_row, Needs_Assignment_col).Value, pnd_case_need_assign)
				call read_boolean_from_excel(ObjExcel.Cells(full_excel_row, Currently_Assigned_col).Value, pnd_case_curr_assign)
				If pnd_case_need_assign = "" Then pnd_case_need_assign = False
				If pnd_case_curr_assign = "" Then pnd_case_curr_assign = False
				assign_this_case = False
				If pnd_case_priority = priority_select and pnd_case_need_assign = True and pnd_case_curr_assign = False Then
					' Call random_selection(3, assign_this_case)
					assign_this_case = True
				End If
				If InStr(rows_to_skip_string, "~" & full_excel_row & "~") <> 0 Then assign_this_case = False
				If assign_this_case Then

					ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_assign_hsr_col).Value 	= first_name & " " & left(last_name, 1)
					ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_assign_date_col).Value 	= date
					ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_case_numb_col).Value 	= ObjExcel.Cells(full_excel_row, Case_Number_col).Value
					ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_case_name_col).Value 	= ObjExcel.Cells(full_excel_row, Case_Name_col).Value
					ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_appl_date_col).Value 	= ObjExcel.Cells(full_excel_row, APPL_Date_col).Value
					ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_hc_eval_date_col).Value = ObjExcel.Cells(full_excel_row, HC_Eval_Date_col).Value
					ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_verifs_date_col).Value 	= ObjExcel.Cells(full_excel_row, Verifs_Requested_Date_col).Value

					' ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_days_pend_col).Value 	= ObjExcel.Cells(full_excel_row, ).Value
					' ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_population_col).Value 	= ObjExcel.Cells(full_excel_row, ).Value
					' ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_day_60_col).Value 		= ObjExcel.Cells(full_excel_row, ).Value
					' ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_assign_compl_col).Value = ObjExcel.Cells(full_excel_row, ).Value
					' ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_case_stat_col).Value 	= ObjExcel.Cells(full_excel_row, ).Value
					' ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_specialty_col).Value 	= ObjExcel.Cells(full_excel_row, ).Value
					' ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_deny_date_col).Value 	= ObjExcel.Cells(full_excel_row, ).Value
					' ObjWrkrExcel.Cells(wrkr_excel_row, wrkr_notes_col).Value 		= ObjExcel.Cells(full_excel_row, ).Value

					ObjExcel.Cells(full_excel_row, Most_Recent_Assignment_Worker_col).Value = first_name & " " & left(last_name, 1)
					ObjExcel.Cells(full_excel_row, Most_Recent_Assignment_Date_col).Value = date
					ObjExcel.Cells(full_excel_row, Currently_Assigned_col).Value = True
					wrkr_excel_row = wrkr_excel_row + 1
					selected_case_count = selected_case_count + 1
					STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
				End If
				If selected_case_count >= requested_case_count Then Exit Do
				full_excel_row = full_excel_row + 1
				next_case_numb = trim(ObjExcel.Cells(full_excel_row, Case_Number_col).Value)
			Loop until next_case_numb = ""
			If selected_case_count = requested_case_count Then Exit For
		Next

		'report to the worker that the cases are ready - leave the worklist open.
		ObjWrkrExcel.ActiveWorkbook.SaveAs indv_worklist_file_path
		ObjWrkrExcel.ActiveWorkbook.Close
		ObjWrkrExcel.Application.Quit
		ObjWrkrExcel.Quit
	End If

	objWorkbook.Save()		'saving the excel
	ObjExcel.ActiveWorkbook.Close
	ObjExcel.Application.Quit
	ObjExcel.Quit

	Call release_data_lock("HOLD")

	If operation_selection <> "Complete Individual Worklist" Then
		Call excel_open(indv_worklist_file_path, True, True, ObjWrkrExcel, objWrkrWorkbook)

		end_msg = "Worklist created of HC Pending Cases." & vbCr
		If COMP_ASSIGN_ARRAY(wrkr_case_numb_col, 0) <> "" Then end_msg = end_msg & UBound(COMP_ASSIGN_ARRAY, 2)+1 & " cases from previous worklist have been recorded." & vbCr
		end_msg = end_msg & requested_case_count & " cases added to a new worklist."
	End If
End If

'Logging usage stats
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure(end_msg)



' BeginDialog Dialog1, 0, 0, 506, 260, "Pending Health Care Counts"
'   ButtonGroup ButtonPressed
'     OkButton 395, 240, 50, 15
'     CancelButton 450, 240, 50, 15
'   Text 15, 15, 50, 10, "Total Cases:"
'   Text 75, 15, 35, 10, "XXX (total cases)"
'   Text 190, 15, 30, 10, " MIPPAs:"
'   Text 30, 90, 20, 10, "XXX (1-10)"
'   Text 200, 25, 20, 10, " EMA:"
'   Text 230, 25, 35, 10, "XXX (ema)"
'   Text 160, 35, 65, 10, "METS Transitions:"
'   Text 230, 35, 35, 10, "XXX (mets trans)"
'   Text 185, 45, 40, 10, "  Standard: "
'   Text 230, 45, 35, 10, "XXX (reg cases)"
'   GroupBox 15, 60, 485, 175, "Standard Cases"
'   Text 25, 75, 60, 10, "Pending Days"
'   Text 55, 90, 25, 10, " - 1 - 10"
'   Text 230, 15, 35, 10, "XXX (mippa)"
'   Text 30, 100, 20, 10, "XXX (11-20)"
'   Text 55, 100, 30, 10, " - 11 - 20"
'   Text 30, 110, 20, 10, "XXX (21-30)"
'   Text 55, 110, 35, 10, " - 21 - 30"
'   Text 30, 120, 20, 10, "XXX (31-40)"
'   Text 55, 120, 35, 10, " - 31 - 40"
'   Text 30, 130, 20, 10, "XXX (41-50)"
'   Text 55, 130, 35, 10, " - 41 - 50"
'   Text 30, 140, 20, 10, "XXX (51-60)"
'   Text 55, 140, 35, 10, " - 51 - 60"
'   Text 30, 150, 20, 10, "XXX (60+)"
'   Text 55, 150, 35, 10, " - Over 60"
'   Text 20, 175, 70, 10, "Work Process"
'   Text 30, 190, 20, 10, "XXX (HC Evan)"
'   Text 55, 190, 75, 10, " - HC Eval Done"
'   Text 30, 200, 20, 10, "XXX (verifs)"
'   Text 55, 200, 65, 10, " - Verifs Sent"
'   Text 30, 210, 20, 10, "XXX (smrt)"
'   Text 55, 210, 65, 10, " - SMRT App"
'   Text 145, 75, 50, 10, "Assignments"
'   Text 155, 90, 25, 10, "XXX (mippa)"
'   Text 180, 90, 100, 10, " - Available for Assignment"
'   Text 160, 100, 25, 10, "XXX (pri 1)"
'   Text 185, 100, 135, 10, " - Priority 1 - Overdue and Verifs Due"
'   Text 160, 110, 25, 10, "XXX (pri 2)"
'   Text 185, 110, 135, 10, " - Priority 2 - HC Eval Not Complete"
'   Text 160, 120, 25, 10, "XXX (pri 3)"
'   Text 185, 120, 135, 10, " - Priority 3 - Case at Day 20"
'   Text 160, 130, 25, 10, "XXX (pri 4)"
'   Text 185, 130, 135, 10, " - Priority 4 - Case at Day 45"
'   Text 160, 140, 25, 10, "XXX (pri 5)"
'   Text 185, 140, 135, 10, " - Priority 5 - Case at Day 55"
'   Text 160, 150, 25, 10, "XXX (pri 6)"
'   Text 185, 150, 135, 10, " - Priority 6 - Case at Day 60"
'   Text 330, 90, 25, 10, "XXX (assigned)"
'   Text 355, 90, 100, 10, " - Currently Assigned"
'   Text 335, 100, 25, 10, "XXX (hsr 1)"
'   Text 360, 100, 135, 10, " - HSR NAME 1"
'   Text 335, 110, 25, 10, "XXX (hsr 2)"
'   Text 360, 110, 135, 10, " - HSR NAME 2"
' EndDialog
