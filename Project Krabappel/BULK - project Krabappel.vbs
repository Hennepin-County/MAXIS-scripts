'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "BULK - project Krabappel"
start_time = timer

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Script Files\FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Project Krabappel\KRABAPPEL FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'VARIABLES TO DECLARE-----------------------------------------------------------------------
excel_file_path = "C:\DHS-MAXIS-Scripts\Project Krabappel\Krabappel template.xlsx"		'Might want to predeclare with a default, and allow users to change it.
how_many_cases_to_make = "1"		'Defaults to 1, but users can modify this.

'Opens Excel file here, as it needs to populate the dialog with the details from the spreadsheet.
call excel_open(excel_file_path, True, True, ObjExcel, objWorkbook)

'Set objWorkSheet = objWorkbook.Worksheet
For Each objWorkSheet In objWorkbook.Worksheets
    If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "controls" then scenario_list = scenario_list & chr(9) & objWorkSheet.Name
Next

'DIALOGS-----------------------------------------------------------------------------------------------------------
BeginDialog training_case_creator_dialog, 0, 0, 371, 130, "Training case creator dialog"
  EditBox 85, 5, 220, 15, excel_file_path
  DropListBox 65, 40, 105, 15, "select one..." & scenario_list, scenario_dropdown
  DropListBox 250, 40, 115, 15, "yes, approve all cases"+chr(9)+"no, leave cases in PND2 status"+chr(9)+"no, leave cases in PND1 status", approve_case_dropdown
  EditBox 125, 60, 40, 15, how_many_cases_to_make
  EditBox 130, 80, 235, 15, workers_to_XFER_cases_to
  ButtonGroup ButtonPressed
    OkButton 265, 110, 50, 15
    CancelButton 320, 110, 50, 15
    PushButton 310, 5, 55, 15, "Reload details", reload_excel_file_button
  Text 5, 10, 75, 10, "File path of Excel file:"
  Text 60, 25, 310, 10, "Note: if you're using the DHS-provided spreadsheet, you should not have to change this value."
  Text 5, 45, 55, 10, "Scenario to run:"
  Text 190, 45, 60, 10, "Approve cases?:"
  Text 5, 65, 120, 10, "How many cases are you creating?:"
  Text 5, 85, 125, 10, "Workers to XFER cases to (x1#####):"
  Text 5, 100, 250, 25, "Please note: if you just wrote a scenario on the spreadsheet, it is recommended that you ''test'' it first by running a single case through. DHS staff cannot triage issues with agency-written scenarios."
EndDialog

'--------- Project Krabappel --------------
'Connects to BlueZone
EMConnect ""



Do
	Do
		Dialog training_case_creator_dialog
		If buttonpressed = cancel then stopscript
		If scenario_dropdown = "select one..." then MsgBox ("You must select a scenario from the dropdown!")
	Loop until scenario_dropdown <> "select one..."
	final_check_before_running = MsgBox("Here's what the scenario will try to create. Please review before proceeding:" & Chr(10) & Chr(10) & _
										"Scenario selection: " & scenario_dropdown & Chr(10) & _
										"Approving cases: " & approve_case_dropdown & Chr(10) & _
										"Amt of cases to make: " & how_many_cases_to_make & Chr(10) & _
										"Workers to XFER cases to: " & workers_to_XFER_cases_to & Chr(10) & Chr(10) & _
										"It is VERY IMPORTANT to review these details before proceeding. It is also highly recommended that if you've created your own scenarios, " & _
										"test them first creating a single case. This is to check to see if any details were missed on the spreadsheet. DHS CANNOT TRIAGE ISSUES WITH " & _
										"COUNTY/AGENCY CUSTOMIZED SCENARIOS." & Chr(10) & Chr(10) & _
										"Please also note that creating training cases can take a very long time. If you are creating hundreds of cases, you may want to run this " & _
										"overnight, or on a secondary machine." & Chr(10) & Chr(10) & _
										"If you are ready to continue, press ''Yes''. Otherwise, press ''no'' to return to the previous screen.", vbYesNo)
Loop until final_check_before_running = vbYes

'<<<<<<<<<<<DIALOG SHOULD GO HERE, FOR NOW IT WILL SELECT THE ONLY CASE ON THE LIST
	'DIALOG SHOULD ASK FOR WORKER NUMBERS IN AN EDITBOX (TO TURN TO AN ARRAY)
	'DIALOG SHOULD ASK IF EACH PIECE NEEDS TO HAPPEN (SO, PROVIDE EARLY TERMINATION FOR INSTANCES WHERE WE JUST WANT TO LEAVE A CASE IN PND1 OR PND2 STATUS)
	'DIALOG SHOULD POP UP A MSGBOX CONFIRMING DETAILS AND WARNING THAT THIS COULD TAKE A WHILE
	



'Determines how many HH members there are, as this script can run for multiple-member households.
excel_col = 3																		'Col 3 is always the primary applicant's col
Do																					'Loops through each col looking for more HH members. If found, it adds one to the counter.
	If ObjExcel.Cells(2, excel_col).Value <> "" then excel_col = excel_col + 1		'Adds one so that the loop will check again
Loop until ObjExcel.Cells(2, excel_col).Value = ""									'Exits loop when we have no number in the MEMB col
total_membs = excel_col - 3															'minus 3 because we started on column 3

'========================================================================APPL PANELS========================================================================
For cases_to_make = 1 to how_many_cases_to_make

	'Navigates to SELF, checks for MAXIS training, stops if not on MAXIS training
	back_to_self
	EMReadScreen training_region_check, 8, 22, 48
	If training_region_check <> "TRAINING" then script_end_procedure("You must be in the training region to use this script. It will now stop.")

	'Assigning the Excel info to variables for appl, and enters into MAXIS. It does this by first declaring a "starting row" variable for each section, and then
	'	each variable will be that row plus however far down it may be on the spreadsheet. This will enable future variable addition without having to modify
	'	hundreds of variable entries here.

	'Grabs APPL screen variables (APPL date, primary applicant name (memb 01))
	APPL_starting_excel_row = 4		'Starting row for APPL function pieces
	APPL_date = ObjExcel.Cells(APPL_starting_excel_row, 3).Value
	APPL_last_name = ObjExcel.Cells(APPL_starting_excel_row + 1, 3).Value
	APPL_first_name = ObjExcel.Cells(APPL_starting_excel_row + 2, 3).Value
	APPL_middle_initial = ObjExcel.Cells(APPL_starting_excel_row + 3, 3).Value

	'Gets the footer month and year of the application off of the spreadsheet, enters into SELF and transmits (can only enter an application on APPL in the footer month of app)
	footer_month = left(APPL_date, 2)
	If right(footer_month, 1) = "/" then footer_month = "0" & left(footer_month, 1)		'Does this to account for single digit months
	footer_year = right(APPL_date, 2)
	EMWriteScreen footer_month, 20, 43
	EMWriteScreen footer_year, 20, 46
	transmit

	'Goes to APPL function
	call navigate_to_screen("APPL", "____")

	'Enters info in APPL and transmits
	call create_MAXIS_friendly_date(APPL_date, 0, 4, 63)
	EMWriteScreen APPL_last_name, 7, 30
	EMWriteScreen APPL_first_name, 7, 63
	EMWriteScreen APPL_middle_initial, 7, 79
	transmit

	'Uses a for...next to enter each HH member's info
	For current_memb = 1 to total_membs
		current_excel_col = current_memb + 2							'There's two columns before the first HH member, so we have to add 2 to get the current excel col
		reference_number = ObjExcel.Cells(2, current_excel_col).Value	'Always in the second row. This is the HH member number

		'Gets MEMB info for the current household member using the current_excel_col field. Starts by declaring the MEMB starting row
		MEMB_starting_excel_row = 5
		MEMB_last_name = ObjExcel.Cells(MEMB_starting_excel_row, current_excel_col).Value
		MEMB_first_name = ObjExcel.Cells(MEMB_starting_excel_row + 1, current_excel_col).Value
		MEMB_mid_init = ObjExcel.Cells(MEMB_starting_excel_row + 2, current_excel_col).Value
		MEMB_age = ObjExcel.Cells(MEMB_starting_excel_row + 3, current_excel_col).Value
		MEMB_DOB_verif = ObjExcel.Cells(MEMB_starting_excel_row + 4, current_excel_col).Value
		MEMB_gender = ObjExcel.Cells(MEMB_starting_excel_row + 5, current_excel_col).Value
		MEMB_ID_verif = ObjExcel.Cells(MEMB_starting_excel_row + 6, current_excel_col).Value
		MEMB_rel_to_appl = ObjExcel.Cells(MEMB_starting_excel_row + 7, current_excel_col).Value
		MEMB_spoken_lang = ObjExcel.Cells(MEMB_starting_excel_row + 8, current_excel_col).Value
		MEMB_interpreter_yn = ObjExcel.Cells(MEMB_starting_excel_row + 9, current_excel_col).Value
		MEMB_alias_yn = ObjExcel.Cells(MEMB_starting_excel_row + 10, current_excel_col).Value
		MEMB_hisp_lat_yn = ObjExcel.Cells(MEMB_starting_excel_row + 11, current_excel_col).Value

		DO  'This DO-LOOP is to check that the CL's SSN created via random number generation is unique. If the SSN matches an SSN on file, the script creates a new SSN and re-enters the CL's information on MEMB. The checking for duplicates part is on the bottom, as that occurs when the worker presses transmit.
			DO
				Randomize
				ssn_first = Rnd
				ssn_first = 1000000000 * ssn_first
				ssn_first = left(ssn_first, 3)
			LOOP UNTIL left(ssn_first, 1) <> "9"	'starting with a 9 is invalid
			Randomize
			ssn_mid = Rnd
			ssn_mid = 100000000 * ssn_mid
			ssn_mid = left(ssn_mid, 2)
			Randomize
			ssn_end = Rnd 
			ssn_end = 100000000 * ssn_end
			ssn_end = left(ssn_end, 4)
		
			'Entering info on MEMB
			EMWriteScreen reference_number, 4, 33
			EMWriteScreen MEMB_last_name, 6, 30
			EMWriteScreen MEMB_first_name, 6, 63
			EMWriteScreen MEMB_mid_init, 6, 79
			EMWriteScreen ssn_first, 7, 42		'Determined above
			EMWriteScreen ssn_mid, 7, 46
			EMWriteScreen ssn_end, 7, 49
			EMWriteScreen "P", 7, 68			'All SSNs should pend in the training region
			EMWriteScreen "01", 8, 42			'At this time, everyone will have a January 1st birthday. The year will be determined by the age on the spreadsheet
			EMWriteScreen "01", 8, 45
			EMWriteScreen datepart("yyyy", date) - abs(MEMB_age), 8, 48
			EMWriteScreen MEMB_DOB_verif, 8, 68
			EMWriteScreen MEMB_gender, 9, 42
			EMWriteScreen MEMB_ID_verif, 9, 68
			EMWriteScreen MEMB_rel_to_appl, 10, 42
			EMWriteScreen MEMB_spoken_lang, 12, 42
			EMWriteScreen MEMB_spoken_lang, 13, 42
			EMWriteScreen MEMB_interpreter_yn, 14, 68
			EMWriteScreen MEMB_alias_yn, 15, 42
			EMWriteScreen MEMB_alien_ID, 15, 68
			EMWriteScreen MEMB_hisp_lat_yn, 16, 68
			EMWriteScreen "X", 17, 34			'Enters race as unknown at this time
			transmit
			DO				'Does this as a loop based on Robert's suggestion that there may be issues in loading without one. It's a small popup window.
				EMReadScreen race_mini_box, 18, 5, 12
				IF race_mini_box = "X AS MANY AS APPLY" THEN
					EMWriteScreen "X", 15, 12
					transmit
					transmit
				END IF
			LOOP UNTIL race_mini_box = "X AS MANY AS APPLY"
			cl_ssn = ssn_first & "-" & ssn_mid & "-" & ssn_end
			EMReadScreen ssn_match, 11, 8, 7
			IF cl_ssn <> ssn_match THEN
				PF8
				PF8
				PF5
			ELSE
				PF3
			END IF
		LOOP UNTIL cl_ssn <> ssn_match
		EMWaitReady 0, 0
		EMWriteScreen "Y", 6, 67
		transmit

		'Gets MEMI info from spreadsheet
		MEMI_starting_excel_row = 17
		MEMI_marital_status = ObjExcel.Cells(MEMI_starting_excel_row, current_excel_col).Value
		MEMI_spouse = ObjExcel.Cells(MEMI_starting_excel_row + 1, current_excel_col).Value
		MEMI_last_grade_completed = ObjExcel.Cells(MEMI_starting_excel_row + 2, current_excel_col).Value
		MEMI_cit_yn = ObjExcel.Cells(MEMI_starting_excel_row + 3, current_excel_col).Value

		'Updates MEMI with the info
		EMWriteScreen MEMI_marital_status, 7, 49
		EMWriteScreen MEMI_spouse, 8, 49
		EMWriteScreen MEMI_last_grade_completed, 9, 49
		EMWriteScreen MEMI_cit_yn, 10, 49
		EMWriteScreen "NO", 10, 78		'Always defaulting to none for cit/ID proof right now
		EMWriteScreen "Y", 13, 49		'Always defualting to yes for been in MN > 12 months
		EMWriteScreen "N", 13, 78		'Always defualting to no for residence verification
		transmit
		
		
	Next

	'This next transmit gets to the ADDR screen
	transmit

	'Gets ADDR info from spreadsheet, gets from column 3 because it's case based
	ADDR_starting_excel_row = 21
	ADDR_line_one = ObjExcel.Cells(ADDR_starting_excel_row, 3).Value
	ADDR_line_two = ObjExcel.Cells(ADDR_starting_excel_row + 1, 3).Value
	ADDR_city = ObjExcel.Cells(ADDR_starting_excel_row + 2, 3).Value
	ADDR_zip = ObjExcel.Cells(ADDR_starting_excel_row + 3, 3).Value
	ADDR_county = ObjExcel.Cells(ADDR_starting_excel_row + 4, 3).Value
	ADDR_addr_verif = ObjExcel.Cells(ADDR_starting_excel_row + 5, 3).Value
	ADDR_homeless = ObjExcel.Cells(ADDR_starting_excel_row + 6, 3).Value
	ADDR_reservation = ObjExcel.Cells(ADDR_starting_excel_row + 7, 3).Value
	ADDR_mailing_addr_line_one = ObjExcel.Cells(ADDR_starting_excel_row + 8, 3).Value
	ADDR_mailing_addr_line_two = ObjExcel.Cells(ADDR_starting_excel_row + 9, 3).Value
	ADDR_mailing_addr_city = ObjExcel.Cells(ADDR_starting_excel_row + 10, 3).Value
	ADDR_mailing_addr_zip = ObjExcel.Cells(ADDR_starting_excel_row + 11, 3).Value
	ADDR_phone_1 = ObjExcel.Cells(ADDR_starting_excel_row + 12, 3).Value
	ADDR_phone_2 = ObjExcel.Cells(ADDR_starting_excel_row + 13, 3).Value
	ADDR_phone_3 = ObjExcel.Cells(ADDR_starting_excel_row + 14, 3).Value

	'Writes spreadsheet info to ADDR
	EMWriteScreen ADDR_line_one, 6, 43
	EMWriteScreen ADDR_line_two, 7, 43
	EMWriteScreen ADDR_city, 8, 43
	EMWriteScreen "MN", 8, 66		'Defaults to MN for all cases at this time
	EMWriteScreen ADDR_zip, 9, 43
	EMWriteScreen ADDR_county, 9, 66
	EMWriteScreen ADDR_addr_verif, 9, 74
	EMWriteScreen ADDR_homeless, 10, 43
	EMWriteScreen ADDR_reservation, 10, 74
	EMWriteScreen ADDR_mailing_addr_line_one, 13, 43
	EMWriteScreen ADDR_mailing_addr_line_two, 14, 43
	EMWriteScreen ADDR_mailing_addr_city, 15, 43
	If ADDR_mailing_addr_line_one <> "" then EMWriteScreen "MN", 16, 43	'Only writes if the user indicated a mailing address. Defaults to MN at this time.
	EMWriteScreen ADDR_mailing_addr_zip, 16, 52
	EMWriteScreen left(ADDR_phone_1, 3), 17, 45						'Has to split phone numbers up into three parts each
	EMWriteScreen mid(ADDR_phone_1, 5, 3), 17, 51
	EMWriteScreen right(ADDR_phone_1, 4), 17, 55
	EMWriteScreen left(ADDR_phone_2, 3), 18, 45
	EMWriteScreen mid(ADDR_phone_2, 5, 3), 18, 51
	EMWriteScreen right(ADDR_phone_2, 4), 18, 55
	EMWriteScreen left(ADDR_phone_3, 3), 19, 45
	EMWriteScreen mid(ADDR_phone_3, 5, 3), 19, 51
	EMWriteScreen right(ADDR_phone_3, 4), 19, 55

	'Reads the case number and adds to an array before exiting
	EMReadScreen current_case_number, 8, 20, 37
	case_number_array = case_number_array & replace(current_case_number, "_", "") & "|"
	
	transmit
	EMReadScreen addr_warning, 7, 3, 6
	IF addr_warning = "Warning" THEN transmit
	transmit
	PF3
Next

'Removing the last "|" from the case_number_array so as to avoid it trying to work a blank case number through PND1
case_number_array = left(case_number_array, len(case_number_array) - 1)

'Splitting the case numbers into an array
case_number_array = split(case_number_array, "|")

'========================================================================PND1 PANELS========================================================================

For each case_number in case_number_array
	'Navigates into STAT. For PND1 cases, this will trigger workflow for adding the right panels.
	call navigate_to_screen ("STAT", "____")
	
	'Transmits, to get to TYPE panel
	transmit
	
	'At this time, it will always mark GRH and IV-E as "N"
	EMWriteScreen "N", 6, 64	'GRH
	EMWriteScreen "N", 6, 73	'IV-E
	
	'Reading and writing info for the TYPE panel
	'Uses a for...next to enter each HH member's info
	For current_memb = 1 to total_membs
		current_excel_col = current_memb + 2							'There's two columns before the first HH member, so we have to add 2 to get the current excel col
		current_MAXIS_row = current_memb + 5							'MEMB 01 always gets entered on row 6, which each subsequent added to the following row. Adding 5 to current_memb simplifies this.
		'reference_number = ObjExcel.Cells(2, current_excel_col).Value	'Always in the second row. This is the HH member number
		
		'Reading the info
		TYPE_starting_excel_row = 36
		TYPE_cash_yn = objExcel.Cells(TYPE_starting_excel_row, current_excel_col).Value
		TYPE_hc_yn = objExcel.Cells(TYPE_starting_excel_row + 1, current_excel_col).Value
		TYPE_fs_yn = objExcel.Cells(TYPE_starting_excel_row + 2, current_excel_col).Value
		
		'Writing the info
		EMWriteScreen TYPE_cash_yn, current_MAXIS_row, 28
		EMWriteScreen TYPE_hc_yn, current_MAXIS_row, 37
		EMWriteScreen TYPE_fs_yn, current_MAXIS_row, 46
		EMWriteScreen "N", current_MAXIS_row, 55			'At this time, it will always mark EMER as "N"
		
		'If any TYPE options are selected, we need to track this to know which items to type on PROG. If any are "Y", it'll update these variables.
		If ucase(TYPE_cash_yn) = "Y" then cash_application = True
		If ucase(TYPE_hc_yn) = "Y" then hc_application = True
		If ucase(TYPE_fs_yn) = "Y" then SNAP_application = True
	Next
	
	'Transmits to get to PROG
	transmit
	
	'Gathers the mig worker variable from Excel. Since it's the only one, we won't use a PROG starting row variable. And since it's case based, we'll only look in col 3
	PROG_mig_worker = objExcel.Cells(39, 3).Value
	
	'Enters in the APPL date on PROG for any programs applied for, and the interview date will always be the APPL date at this time.
	If cash_application = True then
		call create_MAXIS_friendly_date(APPL_date, 0, 6, 33)
		call create_MAXIS_friendly_date(APPL_date, 0, 6, 44)
		call create_MAXIS_friendly_date(APPL_date, 0, 6, 55)
	End if
	If SNAP_application = True then
		call create_MAXIS_friendly_date(APPL_date, 0, 10, 33)
		call create_MAXIS_friendly_date(APPL_date, 0, 10, 44)
		call create_MAXIS_friendly_date(APPL_date, 0, 10, 55)
	End if
	If HC_application = True then call create_MAXIS_friendly_date(APPL_date, 0, 12, 33)		'No interview or elig begin dt for HC
	
	'Enters migrant worker info
	EMWriteScreen PROG_mig_worker, 18, 67
	
	'If the case is HC, it needs to transmit one more time, to get off of the HCRE screen (we'll add it later)
	If HC_application = True then transmit
	
	'Transmits (gets to REVW)
	transmit
	
	'Now we're on REVW and it needs to take different actions for each program. We need to know 6 month and 12 month dates though, for the sake of figuring out review months.
	'Scanning info from REPT section of spreadsheet
	REVW_starting_excel_row = 40
	REVW_ar_or_ir = objExcel.Cells(REVW_starting_excel_row, 3).Value	'Will return either a blank, an "IR", or an "AR"
	REVW_exempt = objExcel.Cells(REVW_starting_excel_row + 1, 3).Value	'Case based, so we'll only look at col 3
	
	'Determining those dates
	six_month_recert_date = dateadd("m", 6, APPL_date)							'Determines info for the six month recert
	six_month_month = datepart("m", six_month_recert_date)
	If len(six_month_month) = 1 then six_month_month = "0" & six_month_month 
	six_month_year = right(six_month_recert_date, 2)
	one_year_recert_date = dateadd("m", 12, APPL_date)							'Determines info for the annual recert
	one_year_month = datepart("m", one_year_recert_date)
	If len(one_year_month) = 1 then one_year_month = "0" & one_year_month 
	one_year_year = right(one_year_recert_date, 2)

	'Adds cash dates
	If cash_application = true then
		EMWriteScreen one_year_month, 9, 37
		EMWriteScreen one_year_year, 9, 43
	End if
	
	'Adds SNAP dates and info
	If SNAP_application = true then
		EMWriteScreen "N", 15, 75		'Phone interview field
		EMWriteScreen "x", 5, 58		
		transmit
		EMWriteScreen six_month_month, 9, 26
		EMWriteScreen six_month_year, 9, 32
		EMWriteScreen one_year_month, 9, 64
		EMWriteScreen one_year_year, 9, 70
		transmit
		transmit
	End if
	
	'Adds HC dates and info
	If HC_application = true then
		EMWriteScreen "x", 5, 71
		transmit
		If REVW_ar_or_ir = "IR" then
			EMWriteScreen six_month_month, 8, 27
			EMWriteScreen six_month_year, 8, 33
		ElseIf REVW_ar_or_ir = "AR" then
			EMWriteScreen six_month_month, 8, 71
			EMWriteScreen six_month_year, 8, 77
		End if
		EMWriteScreen one_year_month, 9, 27
		EMWriteScreen one_year_year, 9, 33
		EMWriteScreen REVW_exempt, 9, 71
		transmit
		transmit
	End if

	transmit
	transmit	
	
Next

'========================================================================PND2 PANELS========================================================================
For each case_number in case_number_array
	'Navigates to STAT/SUMM for each case
	call navigate_to_screen("STAT", "SUMM")
	ERRR_screen_check
		
	'Uses a for...next to enter each HH member's info
	For current_memb = 1 to total_membs
		current_excel_col = current_memb + 2							'There's two columns before the first HH member, so we have to add 2 to get the current excel col
		reference_number = ObjExcel.Cells(2, current_excel_col).Value	'Always in the second row. This is the HH member number
		
		'Goes to STAT/MEMB to associate a SSN to each member, this will be useful for UNEA/MEDI panels
		call navigate_to_screen("STAT", "MEMB")
		EMWriteScreen reference_number, 20, 76
		transmit
		EMReadScreen SSN_first, 3, 7, 42
		EMReadScreen SSN_mid, 2, 7, 46
		EMReadScreen SSN_last, 4, 7, 49
		
		
	Next


Next


'========================================================================APPROVAL========================================================================
'<<<<<Needs to wait for cases to come out of background


MsgBox "EXIT"
stopscript

'call script_end_procedure("Success! Your cases have been made and transferred to the workers indicated in the dialog.")