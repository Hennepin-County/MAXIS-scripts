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
excel_file_path = "C:\DHS-MAXIS-Scripts\Project Krabappel\Krabappel template.xlsx"

'--------- Project Krabappel --------------
'Connects to BlueZone
EMConnect ""

'Opens Excel file
call excel_open(excel_file_path, True, True, ObjExcel, objWorkbook)

'<<<<<<<<<<<DIALOG SHOULD GO HERE, FOR NOW IT WILL SELECT THE ONLY CASE ON THE LIST

'Determines how many HH members there are, as this script can run for multiple-member households.
excel_col = 3																		'Col 3 is always the primary applicant's col
Do																					'Loops through each col looking for more HH members. If found, it adds one to the counter.
	If ObjExcel.Cells(2, excel_col).Value <> "" then excel_col = excel_col + 1		'Adds one so that the loop will check again
Loop until ObjExcel.Cells(2, excel_col).Value = ""									'Exits loop when we have no number in the MEMB col
total_membs = excel_col - 3															'minus 3 because we started on column 3

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

transmit
EMReadScreen addr_warning, 7, 3, 6
IF addr_warning = "Warning" THEN transmit
transmit
PF3

stopscript

'PND1 function variables
'TYPE_cash_yn
'TYPE_hc_yn
'TYPE_fs_yn
'PROG_mig_worker
'REVW_ar_or_ir
'REVW_exempt

'VARIABLES THAT NEED TO BE COLLECTED PER EACH MEMB (IN FOR NEXT)
'SSN_first
'SSN_mid
'SSN_last

'Do all STAT panels
'STORE ALL CASE NUMBERS AS AN ARRAY!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'Do approval

call script_end_procedure("")