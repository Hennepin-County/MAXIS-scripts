'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "BULK - project Krabappel"
start_time = timer

'Hello MsgBox
hello_MsgBox = MsgBox("Hello! Thanks for clicking the training case creator!" & Chr(10) & Chr(10) & _
				"This script is very new, and was put together with a lot of hard work from five scriptwriters, using bits and pieces from various scripts written over the last few years." & Chr(10) & Chr(10) & _
				"We're very proud of the work we've done, but it's a very new concept, and there's bound to be issues here and there. Please keep this in mind as you use this exciting new tool!" & Chr(10) & Chr(10) & _
				"If you run into any issues, or have any questions, please join the discussion in SIR's BlueZone Scripts page. One of the scriptwriters involved will be more than happy to assist you." & Chr(10) & Chr(10) & _
				"Have fun!" & Chr(10) & Chr(10) & _
				"Veronica Cary, DHS" & Chr(10) & _
				"Robert Fewins-Kalb, Anoka County" & Chr(10) & _
				"Charles Potter, Anoka County" & Chr(10) & _
				"David Courtright, St. Louis County" & Chr(10) & _
				"Lucas Shanley, St. Louis County", vbOKCancel)
If hello_MsgBox = vbCancel then stopscript

'LOADING ROUTINE FUNCTIONS---------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER FUNCTIONS LIBRARY.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			script_end_procedure("Script ended due to error connecting to GitHub.")
END IF

'LOADING ROUTINE FUNCTIONS
Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
Set fso_command = run_another_script_fso.OpenTextFile("C:\DHS-MAXIS-Scripts\Project Krabappel\KRABAPPEL FUNCTIONS FILE.vbs")
text_from_the_other_script = fso_command.ReadAll
fso_command.Close
Execute text_from_the_other_script

'========================================================================TRANSFER CASES========================================================================
Function transfer_cases(workers_to_XFER_cases_to, case_number_array)
	'Creates an array of the workers selected in the dialog
	workers_to_XFER_cases_to = split(replace(workers_to_XFER_cases_to, " ", ""), ",")

	'Creates a new two-dimensional array for assigning a worker to each case_number, and collecting the county code for each worker
	Dim transfer_array()
	ReDim transfer_array(ubound(case_number_array), 2)

	'Assigns a case_number to each row in the first column of the array
	For x = 0 to ubound(case_number_array)
		transfer_array(x, 0) = case_number_array(x)
	Next

	'Reassigning x as a 0 for the following do...loop
	x = 0

	'Assigning y as 0, to be used by the following do...loop for deciding which worker gets which case
	y = 0

	'Now, it'll assign a worker to each case number in the transfer_array. Does this on a loop so that a worker can get multiple cases if that is indicated.
	Do
		transfer_array(x, 1) = workers_to_XFER_cases_to(y)	'Assigns column 2 of the array to a worker in the workers_to_XFER_cases_to array
		x = x + 1											'Adds +1 to X
		y = y + 1											'Adds +1 to Y
		If y > ubound(workers_to_XFER_cases_to) then y = 0	'Resets to allow the first worker in the array to get anonther one
	Loop until x > ubound(case_number_array)

	'--------Now, the array is two columns (case_number, worker_assigned)!

'	'Script must figure out who the current worker is, and what agency they are with. This is vital because transferring within an agency uses different screens than inter-agency.
'		'To do this, the script will start by analysing the current worker in REPT/ACTV.

	'Now, the array will figure out the current worker, by looking it up in SELF
	back_to_SELF
	EMReadScreen current_user, 7, 22, 8
	
	'Now, it will go to REPT/USER, and look up the county code for this individual.
	call navigate_to_screen("REPT", "USER")
	EMWriteScreen current_user, 21, 12
	transmit
	
	'Now, it will read the county code for the current user
	EMReadScreen user_county_code, 2, 7, 38
	
	'Now, we enter each worker number into this screen, and return their county codes inside the array.
	For x = 0 to ubound(case_number_array)
		EMWriteScreen transfer_array(x, 1), 21, 12		'Writes worker
		transmit										'Gets to next screen
		EMReadScreen array_county_code, 2, 7, 38		'Reads the county code for this worker
		transfer_array(x, 2) = array_county_code		'Adds the array_county_code to x, 2 of the transfer array
	Next

	'Resetting "x" to be a zero placeholder for the following for...next
	x = 0

	'Now we actually transfer the cases. This for...next does the work (details in comments below)
	For x = 0 to ubound(case_number_array)		'case_number_array is the same as the first col of the transfer_array
		'Assigns the number from the array to the case_number variable
		case_number = transfer_array(x, 0)
		
		'Checks to make sure case isn't in background
		MAXIS_background_check
		
		'Determines interagency transfers by comparing the current active user (gathered above) to the user in the transfer array.
		If user_county_code = transfer_array(x, 2) then
			county_to_county_XFER = False
		Else
			county_to_county_XFER = True
		End if

		'Getting to SPEC/XFER manually
		back_to_SELF
		EMWriteScreen "SPEC", 16, 43
		EMWriteScreen "________", 18, 43
		EMWriteScreen case_number, 18, 43
		EMWriteScreen "XFER", 21, 70
		transmit
		
		'Now to transfer the cases.
		If county_to_county_XFER = False then
			EMWriteScreen "x", 7, 16
			transmit
			PF9
			EMWriteScreen transfer_array(x, 1), 18, 61
			transmit
			transmit
		Else
			EMWriteScreen "x", 9, 16
			transmit
			PF9
			call create_MAXIS_friendly_date(date, 0, 4, 28)
			call create_MAXIS_friendly_date(date, 0, 4, 61)
			EMWriteScreen "N", 5, 28
			call create_MAXIS_friendly_date(date, 0, 5, 61)
			EMWriteScreen transfer_array(x, 1), 18, 61
			transmit
			transmit
		End if
	Next
End function

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
'NOTE: droplistbox for scenario list must be: ["select one..." & scenario_list] in order to be dynamic
BeginDialog training_case_creator_dialog, 0, 0, 446, 125, "Training case creator dialog"
  EditBox 85, 5, 295, 15, excel_file_path
  DropListBox 60, 40, 140, 15, "select one..." & scenario_list, scenario_dropdown
  DropListBox 275, 40, 165, 15, "yes, approve all cases"+chr(9)+"no, but enter all STAT panels needed to approve"+chr(9)+"no, but do TYPE/PROG/REVW"+chr(9)+"no, just APPL all cases", approve_case_dropdown
  EditBox 125, 60, 40, 15, how_many_cases_to_make
  CheckBox 205, 65, 210, 10, "Check here to XFER cases, and enter worker numbers below.", XFER_check
  EditBox 130, 80, 310, 15, workers_to_XFER_cases_to
  ButtonGroup ButtonPressed
    OkButton 335, 105, 50, 15
    CancelButton 390, 105, 50, 15
    PushButton 385, 5, 55, 15, "Reload details", reload_excel_file_button
  Text 5, 10, 75, 10, "File path of Excel file:"
  Text 130, 25, 310, 10, "Note: if you're using the DHS-provided spreadsheet, you should not have to change this value."
  Text 5, 45, 55, 10, "Scenario to run:"
  Text 210, 45, 65, 10, "App/XFER cases?:"
  Text 5, 65, 120, 10, "How many cases are you creating?:"
  Text 5, 85, 125, 10, "Workers to XFER cases to (x1#####):"
  Text 5, 100, 325, 20, "Please note: if you just wrote a scenario on the spreadsheet, it is recommended that you ''test'' it first by running a single case through. DHS staff cannot triage issues with agency-written scenarios."
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
	

'Activates worksheet based on user selection
objExcel.worksheets(scenario_dropdown).Activate

'Determines how many HH members there are, as this script can run for multiple-member households.
excel_col = 3																		'Col 3 is always the primary applicant's col
Do																					'Loops through each col looking for more HH members. If found, it adds one to the counter.
	If ObjExcel.Cells(2, excel_col).Value <> "" then excel_col = excel_col + 1		'Adds one so that the loop will check again
Loop until ObjExcel.Cells(2, excel_col).Value = ""									'Exits loop when we have no number in the MEMB col
total_membs = excel_col - 3															'minus 3 because we started on column 3

'Focuses BlueZone so that everyone can see what it's doing
EMFocus

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

		DO	'This DO-LOOP is to check that the CL's SSN created via random number generation is unique. If the SSN matches an SSN on file, the script creates a new SSN and re-enters the CL's information on MEMB. The checking for duplicates part is on the bottom, as that occurs when the worker presses transmit.
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

'Ends here if the user selected to just APPL all cases
If approve_case_dropdown = "no, just APPL all cases" then 
	If XFER_check = checked then call transfer_cases(workers_to_XFER_cases_to, case_number_array)
	script_end_procedure("Success! Cases made and appl'd, per your request.")
End if
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


'Ends here if the user selected to just do TYPE/PROG/REVW for all cases
If approve_case_dropdown = "no, but do TYPE/PROG/REVW" then 
	If XFER_check = checked then call transfer_cases(workers_to_XFER_cases_to, case_number_array)
	script_end_procedure("Success! Cases made and appl'd, and TYPE/PROG/REVW updated, per your request.")
End if
'========================================================================PND2 PANELS========================================================================

For each case_number in case_number_array
	
	
	'Navigates to STAT/SUMM for each case
	call navigate_to_screen("STAT", "SUMM")
	MAXIS_background_check
	ERRR_screen_check
		
	'Uses a for...next to enter each HH member's info (person based panels only
	For current_memb = 1 to total_membs
		current_excel_col = current_memb + 2							'There's two columns before the first HH member, so we have to add 2 to get the current excel col
		reference_number = ObjExcel.Cells(2, current_excel_col).Value	'Always in the second row. This is the HH member number
		
		'--------------READS ENTIRE EXCEL SHEET FOR THIS HH MEMB
		ABPS_starting_excel_row = 42
		ABPS_supp_coop = ObjExcel.Cells(ABPS_starting_excel_row, current_excel_col).Value
		ABPS_gc_status = ObjExcel.Cells(ABPS_starting_excel_row + 1, current_excel_col).Value

		ACCT_starting_excel_row = 44
		ACCT_type = ObjExcel.Cells(ACCT_starting_excel_row, current_excel_col).Value
		ACCT_numb = ObjExcel.Cells(ACCT_starting_excel_row + 1, current_excel_col).Value
		ACCT_location = ObjExcel.Cells(ACCT_starting_excel_row + 2, current_excel_col).Value
		ACCT_balance = ObjExcel.Cells(ACCT_starting_excel_row + 3, current_excel_col).Value
		ACCT_bal_ver = ObjExcel.Cells(ACCT_starting_excel_row + 4, current_excel_col).Value
		ACCT_date = ObjExcel.Cells(ACCT_starting_excel_row + 5, current_excel_col).Value
		ACCT_withdraw = ObjExcel.Cells(ACCT_starting_excel_row + 6, current_excel_col).Value
		ACCT_cash_count = ObjExcel.Cells(ACCT_starting_excel_row + 7, current_excel_col).Value
		ACCT_snap_count = ObjExcel.Cells(ACCT_starting_excel_row + 8, current_excel_col).Value
		ACCT_HC_count = ObjExcel.Cells(ACCT_starting_excel_row + 9, current_excel_col).Value
		ACCT_GRH_count = ObjExcel.Cells(ACCT_starting_excel_row + 10, current_excel_col).Value
		ACCT_IV_count = ObjExcel.Cells(ACCT_starting_excel_row + 11, current_excel_col).Value
		ACCT_joint_owner = ObjExcel.Cells(ACCT_starting_excel_row + 12, current_excel_col).Value
		ACCT_share_ratio = ObjExcel.Cells(ACCT_starting_excel_row + 13, current_excel_col).Value
		ACCT_interest_date_mo = ObjExcel.Cells(ACCT_starting_excel_row + 14, current_excel_col).Value
		ACCT_interest_date_yr = ObjExcel.Cells(ACCT_starting_excel_row + 15, current_excel_col).Value

		ACUT_starting_excel_row = 60
		ACUT_shared = ObjExcel.Cells(ACUT_starting_excel_row, current_excel_col).Value
		ACUT_heat = ObjExcel.Cells(ACUT_starting_excel_row + 1, current_excel_col).Value
		ACUT_heat_verif = ObjExcel.Cells(ACUT_starting_excel_row + 2, current_excel_col).Value
		ACUT_air = ObjExcel.Cells(ACUT_starting_excel_row + 3, current_excel_col).Value
		ACUT_air_verif = ObjExcel.Cells(ACUT_starting_excel_row + 4, current_excel_col).Value
		ACUT_electric = ObjExcel.Cells(ACUT_starting_excel_row + 5, current_excel_col).Value
		ACUT_electric_verif = ObjExcel.Cells(ACUT_starting_excel_row + 6, current_excel_col).Value
		ACUT_fuel = ObjExcel.Cells(ACUT_starting_excel_row + 7, current_excel_col).Value
		ACUT_fuel_verif = ObjExcel.Cells(ACUT_starting_excel_row + 8, current_excel_col).Value
		ACUT_garbage = ObjExcel.Cells(ACUT_starting_excel_row + 9, current_excel_col).Value
		ACUT_garbage_verif = ObjExcel.Cells(ACUT_starting_excel_row + 10, current_excel_col).Value
		ACUT_water = ObjExcel.Cells(ACUT_starting_excel_row + 11, current_excel_col).Value
		ACUT_water_verif = ObjExcel.Cells(ACUT_starting_excel_row + 12, current_excel_col).Value
		ACUT_sewer = ObjExcel.Cells(ACUT_starting_excel_row + 13, current_excel_col).Value
		ACUT_sewer_verif = ObjExcel.Cells(ACUT_starting_excel_row + 14, current_excel_col).Value
		ACUT_other = ObjExcel.Cells(ACUT_starting_excel_row + 15, current_excel_col).Value
		ACUT_other_verif = ObjExcel.Cells(ACUT_starting_excel_row + 16, current_excel_col).Value
		ACUT_phone = ObjExcel.Cells(ACUT_starting_excel_row + 17, current_excel_col).Value

		BUSI_starting_excel_row = 78
		BUSI_type = ObjExcel.Cells(BUSI_starting_excel_row, current_excel_col).Value
		BUSI_start_date = ObjExcel.Cells(BUSI_starting_excel_row + 1, current_excel_col).Value
		BUSI_end_date = ObjExcel.Cells(BUSI_starting_excel_row + 2, current_excel_col).Value
		BUSI_cash_total_retro = ObjExcel.Cells(BUSI_starting_excel_row + 3, current_excel_col).Value
		BUSI_cash_total_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 4, current_excel_col).Value
		BUSI_cash_total_ver = ObjExcel.Cells(BUSI_starting_excel_row + 5, current_excel_col).Value
		BUSI_IV_total_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 6, current_excel_col).Value
		BUSI_IV_total_ver = ObjExcel.Cells(BUSI_starting_excel_row + 7, current_excel_col).Value
		BUSI_snap_total_retro = ObjExcel.Cells(BUSI_starting_excel_row + 8, current_excel_col).Value
		BUSI_snap_total_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 9, current_excel_col).Value
		BUSI_snap_total_ver = ObjExcel.Cells(BUSI_starting_excel_row + 10, current_excel_col).Value
		BUSI_hc_total_prosp_a = ObjExcel.Cells(BUSI_starting_excel_row + 11, current_excel_col).Value
		BUSI_hc_total_ver_a = ObjExcel.Cells(BUSI_starting_excel_row + 12, current_excel_col).Value
		BUSI_hc_total_prosp_b = ObjExcel.Cells(BUSI_starting_excel_row + 13, current_excel_col).Value
		BUSI_hc_total_ver_b = ObjExcel.Cells(BUSI_starting_excel_row + 14, current_excel_col).Value
		BUSI_cash_exp_retro = ObjExcel.Cells(BUSI_starting_excel_row + 15, current_excel_col).Value
		BUSI_cash_exp_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 16, current_excel_col).Value
		BUSI_cash_exp_ver = ObjExcel.Cells(BUSI_starting_excel_row + 17, current_excel_col).Value
		BUSI_IV_exp_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 18, current_excel_col).Value
		BUSI_IV_exp_ver = ObjExcel.Cells(BUSI_starting_excel_row + 19, current_excel_col).Value
		BUSI_snap_exp_retro = ObjExcel.Cells(BUSI_starting_excel_row + 20, current_excel_col).Value
		BUSI_snap_exp_prosp = ObjExcel.Cells(BUSI_starting_excel_row + 21, current_excel_col).Value
		BUSI_snap_exp_ver = ObjExcel.Cells(BUSI_starting_excel_row + 22, current_excel_col).Value
		BUSI_hc_exp_prosp_a = ObjExcel.Cells(BUSI_starting_excel_row + 23, current_excel_col).Value
		BUSI_hc_exp_ver_a = ObjExcel.Cells(BUSI_starting_excel_row + 24, current_excel_col).Value
		BUSI_hc_exp_prosp_b = ObjExcel.Cells(BUSI_starting_excel_row + 25, current_excel_col).Value
		BUSI_hc_exp_ver_b = ObjExcel.Cells(BUSI_starting_excel_row + 26, current_excel_col).Value
		BUSI_retro_hours = ObjExcel.Cells(BUSI_starting_excel_row + 27, current_excel_col).Value
		BUSI_prosp_hours = ObjExcel.Cells(BUSI_starting_excel_row + 28, current_excel_col).Value
		BUSI_hc_total_est_a = ObjExcel.Cells(BUSI_starting_excel_row + 29, current_excel_col).Value
		BUSI_hc_total_est_b = ObjExcel.Cells(BUSI_starting_excel_row + 30, current_excel_col).Value
		BUSI_hc_exp_est_a = ObjExcel.Cells(BUSI_starting_excel_row + 31, current_excel_col).Value
		BUSI_hc_exp_est_b = ObjExcel.Cells(BUSI_starting_excel_row + 32, current_excel_col).Value
		BUSI_hc_hours_est = ObjExcel.Cells(BUSI_starting_excel_row + 33, current_excel_col).Value

		CARS_starting_excel_row = 112
		CARS_type = ObjExcel.Cells(CARS_starting_excel_row, current_excel_col).Value
		CARS_year = ObjExcel.Cells(CARS_starting_excel_row + 1, current_excel_col).Value
		CARS_make = ObjExcel.Cells(CARS_starting_excel_row + 2, current_excel_col).Value
		CARS_model = ObjExcel.Cells(CARS_starting_excel_row + 3, current_excel_col).Value
		CARS_trade_in = ObjExcel.Cells(CARS_starting_excel_row + 4, current_excel_col).Value
		CARS_loan = ObjExcel.Cells(CARS_starting_excel_row + 5, current_excel_col).Value
		CARS_value_source = ObjExcel.Cells(CARS_starting_excel_row + 6, current_excel_col).Value
		CARS_ownership_ver = ObjExcel.Cells(CARS_starting_excel_row + 7, current_excel_col).Value
		CARS_amount_owed = ObjExcel.Cells(CARS_starting_excel_row + 8, current_excel_col).Value
		CARS_amount_owed_ver = ObjExcel.Cells(CARS_starting_excel_row + 9, current_excel_col).Value
		CARS_date = ObjExcel.Cells(CARS_starting_excel_row + 10, current_excel_col).Value
		CARS_owed_as_of = ObjExcel.Cells(CARS_starting_excel_row + 11, current_excel_col).Value
		CARS_use = ObjExcel.Cells(CARS_starting_excel_row + 12, current_excel_col).Value
		CARS_HC_benefit = ObjExcel.Cells(CARS_starting_excel_row + 13, current_excel_col).Value
		CARS_joint_owner = ObjExcel.Cells(CARS_starting_excel_row + 14, current_excel_col).Value
		CARS_share_ratio = ObjExcel.Cells(CARS_starting_excel_row + 15, current_excel_col).Value

		CASH_starting_excel_row = 127
		CASH_amount = ObjExcel.Cells(CASH_starting_excel_row, current_excel_col).Value

		DCEX_starting_excel_row = 128
		DCEX_provider = ObjExcel.Cells(DCEX_starting_excel_row, current_excel_col).Value
		DCEX_reason = ObjExcel.Cells(DCEX_starting_excel_row + 1, current_excel_col).Value
		DCEX_subsidy = ObjExcel.Cells(DCEX_starting_excel_row + 2, current_excel_col).Value
		DCEX_child_number1 = ObjExcel.Cells(DCEX_starting_excel_row + 3, current_excel_col).Value
		DCEX_child_number1_retro = ObjExcel.Cells(DCEX_starting_excel_row + 4, current_excel_col).Value
		DCEX_child_number1_pro = ObjExcel.Cells(DCEX_starting_excel_row + 5, current_excel_col).Value
		DCEX_child_number1_ver = ObjExcel.Cells(DCEX_starting_excel_row + 6, current_excel_col).Value
		DCEX_child_number2 = ObjExcel.Cells(DCEX_starting_excel_row + 7, current_excel_col).Value
		DCEX_child_number2_retro = ObjExcel.Cells(DCEX_starting_excel_row + 8, current_excel_col).Value
		DCEX_child_number2_pro = ObjExcel.Cells(DCEX_starting_excel_row + 9, current_excel_col).Value
		DCEX_child_number2_ver = ObjExcel.Cells(DCEX_starting_excel_row + 10, current_excel_col).Value
		DCEX_child_number3 = ObjExcel.Cells(DCEX_starting_excel_row + 11, current_excel_col).Value
		DCEX_child_number3_retro = ObjExcel.Cells(DCEX_starting_excel_row + 12, current_excel_col).Value
		DCEX_child_number3_pro = ObjExcel.Cells(DCEX_starting_excel_row + 13, current_excel_col).Value
		DCEX_child_number3_ver = ObjExcel.Cells(DCEX_starting_excel_row + 14, current_excel_col).Value
		DCEX_child_number4 = ObjExcel.Cells(DCEX_starting_excel_row + 15, current_excel_col).Value
		DCEX_child_number4_retro = ObjExcel.Cells(DCEX_starting_excel_row + 16, current_excel_col).Value
		DCEX_child_number4_pro = ObjExcel.Cells(DCEX_starting_excel_row + 17, current_excel_col).Value
		DCEX_child_number4_ver = ObjExcel.Cells(DCEX_starting_excel_row + 18, current_excel_col).Value
		DCEX_child_number5 = ObjExcel.Cells(DCEX_starting_excel_row + 19, current_excel_col).Value
		DCEX_child_number5_retro = ObjExcel.Cells(DCEX_starting_excel_row + 20, current_excel_col).Value
		DCEX_child_number5_pro = ObjExcel.Cells(DCEX_starting_excel_row + 21, current_excel_col).Value
		DCEX_child_number5_ver = ObjExcel.Cells(DCEX_starting_excel_row + 22, current_excel_col).Value
		DCEX_child_number6 = ObjExcel.Cells(DCEX_starting_excel_row + 23, current_excel_col).Value
		DCEX_child_number6_retro = ObjExcel.Cells(DCEX_starting_excel_row + 24, current_excel_col).Value
		DCEX_child_number6_pro = ObjExcel.Cells(DCEX_starting_excel_row + 25, current_excel_col).Value
		DCEX_child_number6_ver = ObjExcel.Cells(DCEX_starting_excel_row + 26, current_excel_col).Value

		DIET_starting_excel_row = 155
		DIET_mfip_1 = ObjExcel.Cells(DIET_starting_excel_row, current_excel_col).Value
		DIET_mfip_1_ver = ObjExcel.Cells(DIET_starting_excel_row + 1, current_excel_col).Value
		DIET_mfip_2 = ObjExcel.Cells(DIET_starting_excel_row + 2, current_excel_col).Value
		DIET_mfip_2_ver = ObjExcel.Cells(DIET_starting_excel_row + 3, current_excel_col).Value
		DIET_msa_1 = ObjExcel.Cells(DIET_starting_excel_row + 4, current_excel_col).Value
		DIET_msa_1_ver = ObjExcel.Cells(DIET_starting_excel_row + 5, current_excel_col).Value
		DIET_msa_2 = ObjExcel.Cells(DIET_starting_excel_row + 6, current_excel_col).Value
		DIET_msa_2_ver = ObjExcel.Cells(DIET_starting_excel_row + 7, current_excel_col).Value
		DIET_msa_3 = ObjExcel.Cells(DIET_starting_excel_row + 8, current_excel_col).Value
		DIET_msa_3_ver = ObjExcel.Cells(DIET_starting_excel_row + 9, current_excel_col).Value
		DIET_msa_4 = ObjExcel.Cells(DIET_starting_excel_row + 10, current_excel_col).Value
		DIET_msa_4_ver = ObjExcel.Cells(DIET_starting_excel_row + 11, current_excel_col).Value

		DISA_starting_excel_row = 167
		DISA_begin_date = ObjExcel.Cells(DISA_starting_excel_row, current_excel_col).Value
		DISA_end_date = ObjExcel.Cells(DISA_starting_excel_row + 1, current_excel_col).Value
		DISA_cert_begin = ObjExcel.Cells(DISA_starting_excel_row + 2, current_excel_col).Value
		DISA_cert_end = ObjExcel.Cells(DISA_starting_excel_row + 3, current_excel_col).Value
		DISA_wavr_begin = ObjExcel.Cells(DISA_starting_excel_row + 4, current_excel_col).Value
		DISA_wavr_end = ObjExcel.Cells(DISA_starting_excel_row + 5, current_excel_col).Value
		DISA_grh_begin = ObjExcel.Cells(DISA_starting_excel_row + 6, current_excel_col).Value
		DISA_grh_end = ObjExcel.Cells(DISA_starting_excel_row + 7, current_excel_col).Value
		DISA_cash_status = ObjExcel.Cells(DISA_starting_excel_row + 8, current_excel_col).Value
		DISA_cash_status_ver = ObjExcel.Cells(DISA_starting_excel_row + 9, current_excel_col).Value
		DISA_snap_status = ObjExcel.Cells(DISA_starting_excel_row + 10, current_excel_col).Value
		DISA_snap_status_ver = ObjExcel.Cells(DISA_starting_excel_row + 11, current_excel_col).Value
		DISA_hc_status = ObjExcel.Cells(DISA_starting_excel_row + 12, current_excel_col).Value
		DISA_hc_status_ver = ObjExcel.Cells(DISA_starting_excel_row + 13, current_excel_col).Value
		DISA_waiver = ObjExcel.Cells(DISA_starting_excel_row + 14, current_excel_col).Value
		DISA_drug_alcohol = ObjExcel.Cells(DISA_starting_excel_row + 15, current_excel_col).Value

		DSTT_starting_excel_row = 183
		DSTT_ongoing_income = ObjExcel.Cells(DSTT_starting_excel_row, current_excel_col).Value
		DSTT_HH_income_stop_date = ObjExcel.Cells(DSTT_starting_excel_row + 1, current_excel_col).Value
		DSTT_income_expected_amt = ObjExcel.Cells(DSTT_starting_excel_row + 2, current_excel_col).Value

		EATS_starting_excel_row = 186
		EATS_together = ObjExcel.Cells(EATS_starting_excel_row, current_excel_col).Value
		EATS_boarder = ObjExcel.Cells(EATS_starting_excel_row + 1, current_excel_col).Value
		EATS_group_one = ObjExcel.Cells(EATS_starting_excel_row + 2, current_excel_col).Value
		EATS_group_two = ObjExcel.Cells(EATS_starting_excel_row + 3, current_excel_col).Value
		EATS_group_three = ObjExcel.Cells(EATS_starting_excel_row + 4, current_excel_col).Value

		EMMA_starting_excel_row = 191
		EMMA_medical_emergency = ObjExcel.Cells(EMMA_starting_excel_row, current_excel_col).Value
		EMMA_health_consequence = ObjExcel.Cells(EMMA_starting_excel_row + 1, current_excel_col).Value
		EMMA_verification = ObjExcel.Cells(EMMA_starting_excel_row + 2, current_excel_col).Value
		EMMA_begin_date = ObjExcel.Cells(EMMA_starting_excel_row + 3, current_excel_col).Value
		EMMA_end_date = ObjExcel.Cells(EMMA_starting_excel_row + 4, current_excel_col).Value

		EMPS_starting_excel_row = 196
		EMPS_orientation_date = ObjExcel.Cells(EMPS_starting_excel_row, current_excel_col).Value
		EMPS_orientation_attended = ObjExcel.Cells(EMPS_starting_excel_row + 1, current_excel_col).Value
		EMPS_good_cause = ObjExcel.Cells(EMPS_starting_excel_row + 2, current_excel_col).Value
		EMPS_sanc_begin = ObjExcel.Cells(EMPS_starting_excel_row + 3, current_excel_col).Value
		EMPS_sanc_end = ObjExcel.Cells(EMPS_starting_excel_row + 4, current_excel_col).Value
		EMPS_memb_at_home = ObjExcel.Cells(EMPS_starting_excel_row + 5, current_excel_col).Value
		EMPS_care_family = ObjExcel.Cells(EMPS_starting_excel_row + 6, current_excel_col).Value
		EMPS_crisis = ObjExcel.Cells(EMPS_starting_excel_row + 7, current_excel_col).Value
		EMPS_hard_employ = ObjExcel.Cells(EMPS_starting_excel_row + 8, current_excel_col).Value
		EMPS_under1 = ObjExcel.Cells(EMPS_starting_excel_row + 9, current_excel_col).Value
		EMPS_DWP_date = ObjExcel.Cells(EMPS_starting_excel_row + 10, current_excel_col).Value

		FACI_starting_excel_row = 207
		FACI_vendor_number = ObjExcel.Cells(FACI_starting_excel_row, current_excel_col).Value
		FACI_name = ObjExcel.Cells(FACI_starting_excel_row + 1, current_excel_col).Value
		FACI_type = ObjExcel.Cells(FACI_starting_excel_row + 2, current_excel_col).Value
		FACI_FS_eligible = ObjExcel.Cells(FACI_starting_excel_row + 3, current_excel_col).Value
		FACI_FS_facility_type = ObjExcel.Cells(FACI_starting_excel_row + 4, current_excel_col).Value
		FACI_date_in = ObjExcel.Cells(FACI_starting_excel_row + 5, current_excel_col).Value
		FACI_date_out = ObjExcel.Cells(FACI_starting_excel_row + 6, current_excel_col).Value

		HCRE_starting_excel_row = 218
		HCRE_appl_addnd_date_input = ObjExcel.Cells(HCRE_starting_excel_row, current_excel_col).Value
		HCRE_retro_months_input = ObjExcel.Cells(HCRE_starting_excel_row + 1, current_excel_col).Value
		HCRE_recvd_by_service_date_input = ObjExcel.Cells(HCRE_starting_excel_row + 2, current_excel_col).Value

		HEST_starting_excel_row = 221
		HEST_FS_choice_date = ObjExcel.Cells(HEST_starting_excel_row, current_excel_col).Value
		HEST_first_month = ObjExcel.Cells(HEST_starting_excel_row + 1, current_excel_col).Value
		HEST_heat_air_retro = ObjExcel.Cells(HEST_starting_excel_row + 2, current_excel_col).Value
		HEST_heat_air_pro = ObjExcel.Cells(HEST_starting_excel_row + 3, current_excel_col).Value
		HEST_electric_retro = ObjExcel.Cells(HEST_starting_excel_row + 4, current_excel_col).Value
		HEST_electric_pro = ObjExcel.Cells(HEST_starting_excel_row + 5, current_excel_col).Value
		HEST_phone_retro = ObjExcel.Cells(HEST_starting_excel_row + 6, current_excel_col).Value
		HEST_phone_pro = ObjExcel.Cells(HEST_starting_excel_row + 7, current_excel_col).Value

		IMIG_starting_excel_row = 229
		IMIG_imigration_status = ObjExcel.Cells(IMIG_starting_excel_row, current_excel_col).Value
		IMIG_entry_date = ObjExcel.Cells(IMIG_starting_excel_row + 1, current_excel_col).Value
		IMIG_status_date = ObjExcel.Cells(IMIG_starting_excel_row + 2, current_excel_col).Value
		IMIG_status_ver = ObjExcel.Cells(IMIG_starting_excel_row + 3, current_excel_col).Value
		IMIG_status_LPR_adj_from = ObjExcel.Cells(IMIG_starting_excel_row + 4, current_excel_col).Value
		IMIG_nationality = ObjExcel.Cells(IMIG_starting_excel_row + 5, current_excel_col).Value

		INSA_starting_excel_row = 235
		INSA_pers_coop_ohi = ObjExcel.Cells(INSA_starting_excel_row, current_excel_col).Value
		INSA_good_cause_status = ObjExcel.Cells(INSA_starting_excel_row + 1, current_excel_col).Value
		INSA_good_cause_cliam_date = ObjExcel.Cells(INSA_starting_excel_row + 2, current_excel_col).Value
		INSA_good_cause_evidence = ObjExcel.Cells(INSA_starting_excel_row + 3, current_excel_col).Value
		INSA_coop_cost_effect = ObjExcel.Cells(INSA_starting_excel_row + 4, current_excel_col).Value
		INSA_insur_name = ObjExcel.Cells(INSA_starting_excel_row + 5, current_excel_col).Value
		INSA_prescrip_drug_cover = ObjExcel.Cells(INSA_starting_excel_row + 6, current_excel_col).Value
		INSA_prescrip_end_date = ObjExcel.Cells(INSA_starting_excel_row + 7, current_excel_col).Value

		JOBS_1_starting_excel_row = 243
		JOBS_1_inc_type = ObjExcel.Cells(JOBS_1_starting_excel_row, current_excel_col).Value
		JOBS_1_inc_verif = ObjExcel.Cells(JOBS_1_starting_excel_row + 1, current_excel_col).Value
		JOBS_1_employer_name = ObjExcel.Cells(JOBS_1_starting_excel_row + 2, current_excel_col).Value
		JOBS_1_inc_start = ObjExcel.Cells(JOBS_1_starting_excel_row + 3, current_excel_col).Value
		JOBS_1_pay_freq = ObjExcel.Cells(JOBS_1_starting_excel_row + 4, current_excel_col).Value
		JOBS_1_wkly_hrs = ObjExcel.Cells(JOBS_1_starting_excel_row + 5, current_excel_col).Value
		JOBS_1_hrly_wage = ObjExcel.Cells(JOBS_1_starting_excel_row + 6, current_excel_col).Value

		JOBS_2_starting_excel_row = 250
		JOBS_2_inc_type = ObjExcel.Cells(JOBS_2_starting_excel_row, current_excel_col).Value
		JOBS_2_inc_verif = ObjExcel.Cells(JOBS_2_starting_excel_row + 1, current_excel_col).Value
		JOBS_2_employer_name = ObjExcel.Cells(JOBS_2_starting_excel_row + 2, current_excel_col).Value
		JOBS_2_inc_start = ObjExcel.Cells(JOBS_2_starting_excel_row + 3, current_excel_col).Value
		JOBS_2_pay_freq = ObjExcel.Cells(JOBS_2_starting_excel_row + 4, current_excel_col).Value
		JOBS_2_wkly_hrs = ObjExcel.Cells(JOBS_2_starting_excel_row + 5, current_excel_col).Value
		JOBS_2_hrly_wage = ObjExcel.Cells(JOBS_2_starting_excel_row + 6, current_excel_col).Value

		JOBS_3_starting_excel_row = 257
		JOBS_3_inc_type = ObjExcel.Cells(JOBS_3_starting_excel_row, current_excel_col).Value
		JOBS_3_inc_verif = ObjExcel.Cells(JOBS_3_starting_excel_row + 1, current_excel_col).Value
		JOBS_3_employer_name = ObjExcel.Cells(JOBS_3_starting_excel_row + 2, current_excel_col).Value
		JOBS_3_inc_start = ObjExcel.Cells(JOBS_3_starting_excel_row + 3, current_excel_col).Value
		JOBS_3_pay_freq = ObjExcel.Cells(JOBS_3_starting_excel_row + 4, current_excel_col).Value
		JOBS_3_wkly_hrs = ObjExcel.Cells(JOBS_3_starting_excel_row + 5, current_excel_col).Value
		JOBS_3_hrly_wage = ObjExcel.Cells(JOBS_3_starting_excel_row + 6, current_excel_col).Value

		MEDI_starting_excel_row = 264
		MEDI_claim_number_suffix = ObjExcel.Cells(MEDI_starting_excel_row, current_excel_col).Value
		MEDI_part_A_premium = ObjExcel.Cells(MEDI_starting_excel_row + 1, current_excel_col).Value
		MEDI_part_B_premium = ObjExcel.Cells(MEDI_starting_excel_row + 2, current_excel_col).Value
		MEDI_part_A_begin_date = ObjExcel.Cells(MEDI_starting_excel_row + 3, current_excel_col).Value
		MEDI_part_B_begin_date = ObjExcel.Cells(MEDI_starting_excel_row + 4, current_excel_col).Value

		MMSA_starting_excel_row = 269
		MMSA_liv_arr = ObjExcel.Cells(MMSA_starting_excel_row, current_excel_col).Value
		MMSA_cont_elig = ObjExcel.Cells(MMSA_starting_excel_row + 1, current_excel_col).Value
		MMSA_spous_inc = ObjExcel.Cells(MMSA_starting_excel_row + 2, current_excel_col).Value
		MMSA_shared_hous = ObjExcel.Cells(MMSA_starting_excel_row + 3, current_excel_col).Value

		MSUR_starting_excel_row = 273
		MSUR_begin_date = ObjExcel.Cells(MSUR_starting_excel_row, current_excel_col).Value

		OTHR_starting_excel_row = 274
		OTHR_type = ObjExcel.Cells(OTHR_starting_excel_row, current_excel_col).Value
		OTHR_cash_value = ObjExcel.Cells(OTHR_starting_excel_row + 1, current_excel_col).Value
		OTHR_cash_value_ver = ObjExcel.Cells(OTHR_starting_excel_row + 2, current_excel_col).Value
		OTHR_owed = ObjExcel.Cells(OTHR_starting_excel_row + 3, current_excel_col).Value
		OTHR_owed_ver = ObjExcel.Cells(OTHR_starting_excel_row + 4, current_excel_col).Value
		OTHR_date = ObjExcel.Cells(OTHR_starting_excel_row + 5, current_excel_col).Value
		OTHR_cash_count = ObjExcel.Cells(OTHR_starting_excel_row + 6, current_excel_col).Value
		OTHR_SNAP_count = ObjExcel.Cells(OTHR_starting_excel_row + 7, current_excel_col).Value
		OTHR_HC_count = ObjExcel.Cells(OTHR_starting_excel_row + 8, current_excel_col).Value
		OTHR_IV_count = ObjExcel.Cells(OTHR_starting_excel_row + 9, current_excel_col).Value
		OTHR_joint = ObjExcel.Cells(OTHR_starting_excel_row + 10, current_excel_col).Value
		OTHR_share_ratio = ObjExcel.Cells(OTHR_starting_excel_row + 11, current_excel_col).Value

		PARE_starting_excel_row = 286
		PARE_child_1 = ObjExcel.Cells(PARE_starting_excel_row, current_excel_col).Value
		PARE_child_1_relation = ObjExcel.Cells(PARE_starting_excel_row + 1, current_excel_col).Value
		PARE_child_1_verif = ObjExcel.Cells(PARE_starting_excel_row + 2, current_excel_col).Value
		PARE_child_2 = ObjExcel.Cells(PARE_starting_excel_row + 3, current_excel_col).Value
		PARE_child_2_relation = ObjExcel.Cells(PARE_starting_excel_row + 4, current_excel_col).Value
		PARE_child_2_verif = ObjExcel.Cells(PARE_starting_excel_row + 5, current_excel_col).Value
		PARE_child_3 = ObjExcel.Cells(PARE_starting_excel_row + 6, current_excel_col).Value
		PARE_child_3_relation = ObjExcel.Cells(PARE_starting_excel_row + 7, current_excel_col).Value
		PARE_child_3_verif = ObjExcel.Cells(PARE_starting_excel_row + 8, current_excel_col).Value
		PARE_child_4 = ObjExcel.Cells(PARE_starting_excel_row + 9, current_excel_col).Value
		PARE_child_4_relation = ObjExcel.Cells(PARE_starting_excel_row + 10, current_excel_col).Value
		PARE_child_4_verif = ObjExcel.Cells(PARE_starting_excel_row + 11, current_excel_col).Value
		PARE_child_5 = ObjExcel.Cells(PARE_starting_excel_row + 12, current_excel_col).Value
		PARE_child_5_relation = ObjExcel.Cells(PARE_starting_excel_row + 13, current_excel_col).Value
		PARE_child_5_verif = ObjExcel.Cells(PARE_starting_excel_row + 14, current_excel_col).Value
		PARE_child_6 = ObjExcel.Cells(PARE_starting_excel_row + 15, current_excel_col).Value
		PARE_child_6_relation = ObjExcel.Cells(PARE_starting_excel_row + 16, current_excel_col).Value
		PARE_child_6_verif = ObjExcel.Cells(PARE_starting_excel_row + 17, current_excel_col).Value

		PBEN_1_starting_excel_row = 304
		PBEN_1_referal_date = ObjExcel.Cells(PBEN_1_starting_excel_row, current_excel_col).Value
		PBEN_1_type = ObjExcel.Cells(PBEN_1_starting_excel_row + 1, current_excel_col).Value
		PBEN_1_appl_date = ObjExcel.Cells(PBEN_1_starting_excel_row + 2, current_excel_col).Value
		PBEN_1_appl_ver = ObjExcel.Cells(PBEN_1_starting_excel_row + 3, current_excel_col).Value
		PBEN_1_IAA_date = ObjExcel.Cells(PBEN_1_starting_excel_row + 4, current_excel_col).Value
		PBEN_1_disp = ObjExcel.Cells(PBEN_1_starting_excel_row + 5, current_excel_col).Value

		PBEN_2_starting_excel_row = 310
		PBEN_2_referal_date = ObjExcel.Cells(PBEN_2_starting_excel_row, current_excel_col).Value
		PBEN_2_type = ObjExcel.Cells(PBEN_2_starting_excel_row + 1, current_excel_col).Value
		PBEN_2_appl_date = ObjExcel.Cells(PBEN_2_starting_excel_row + 2, current_excel_col).Value
		PBEN_2_appl_ver = ObjExcel.Cells(PBEN_2_starting_excel_row + 3, current_excel_col).Value
		PBEN_2_IAA_date = ObjExcel.Cells(PBEN_2_starting_excel_row + 4, current_excel_col).Value
		PBEN_2_disp = ObjExcel.Cells(PBEN_2_starting_excel_row + 5, current_excel_col).Value

		PBEN_3_starting_excel_row = 316
		PBEN_3_referal_date = ObjExcel.Cells(PBEN_3_starting_excel_row, current_excel_col).Value
		PBEN_3_type = ObjExcel.Cells(PBEN_3_starting_excel_row + 1, current_excel_col).Value
		PBEN_3_appl_date = ObjExcel.Cells(PBEN_3_starting_excel_row + 2, current_excel_col).Value
		PBEN_3_appl_ver = ObjExcel.Cells(PBEN_3_starting_excel_row + 3, current_excel_col).Value
		PBEN_3_IAA_date = ObjExcel.Cells(PBEN_3_starting_excel_row + 4, current_excel_col).Value
		PBEN_3_disp = ObjExcel.Cells(PBEN_3_starting_excel_row + 5, current_excel_col).Value

		PDED_starting_excel_row = 322
		PDED_wid_deduction = ObjExcel.Cells(PDED_starting_excel_row, current_excel_col).Value
		PDED_adult_child_disregard = ObjExcel.Cells(PDED_starting_excel_row + 1, current_excel_col).Value
		PDED_wid_disregard = ObjExcel.Cells(PDED_starting_excel_row + 2, current_excel_col).Value
		PDED_unea_income_deduction_reason = ObjExcel.Cells(PDED_starting_excel_row + 3, current_excel_col).Value
		PDED_unea_income_deduction_value = ObjExcel.Cells(PDED_starting_excel_row + 4, current_excel_col).Value
		PDED_earned_income_deduction_reason = ObjExcel.Cells(PDED_starting_excel_row + 5, current_excel_col).Value
		PDED_earned_income_deduction_value = ObjExcel.Cells(PDED_starting_excel_row + 6, current_excel_col).Value
		PDED_ma_epd_inc_asset_limit = ObjExcel.Cells(PDED_starting_excel_row + 7, current_excel_col).Value
		PDED_guard_fee = ObjExcel.Cells(PDED_starting_excel_row + 8, current_excel_col).Value
		PDED_rep_payee_fee = ObjExcel.Cells(PDED_starting_excel_row + 9, current_excel_col).Value
		PDED_other_expense = ObjExcel.Cells(PDED_starting_excel_row + 10, current_excel_col).Value
		PDED_shel_spcl_needs = ObjExcel.Cells(PDED_starting_excel_row + 11, current_excel_col).Value
		PDED_excess_need = ObjExcel.Cells(PDED_starting_excel_row + 12, current_excel_col).Value
		PDED_restaurant_meals = ObjExcel.Cells(PDED_starting_excel_row + 13, current_excel_col).Value

		PREG_starting_excel_row = 336
		PREG_conception_date = ObjExcel.Cells(PREG_starting_excel_row, current_excel_col).Value
		PREG_conception_date_ver = ObjExcel.Cells(PREG_starting_excel_row + 1, current_excel_col).Value
		PREG_third_trimester_ver = ObjExcel.Cells(PREG_starting_excel_row + 2, current_excel_col).Value
		PREG_due_date = ObjExcel.Cells(PREG_starting_excel_row + 3, current_excel_col).Value
		PREG_multiple_birth = ObjExcel.Cells(PREG_starting_excel_row + 4, current_excel_col).Value

		RBIC_starting_excel_row = 341
		RBIC_type = ObjExcel.Cells(RBIC_starting_excel_row, current_excel_col).Value
		RBIC_start_date = ObjExcel.Cells(RBIC_starting_excel_row + 1, current_excel_col).Value
		RBIC_end_date = ObjExcel.Cells(RBIC_starting_excel_row + 2, current_excel_col).Value
		RBIC_group_1 = ObjExcel.Cells(RBIC_starting_excel_row + 3, current_excel_col).Value
		RBIC_retro_income_group_1 = ObjExcel.Cells(RBIC_starting_excel_row + 4, current_excel_col).Value
		RBIC_prosp_income_group_1 = ObjExcel.Cells(RBIC_starting_excel_row + 5, current_excel_col).Value
		RBIC_ver_income_group_1 = ObjExcel.Cells(RBIC_starting_excel_row + 6, current_excel_col).Value
		RBIC_group_2 = ObjExcel.Cells(RBIC_starting_excel_row + 7, current_excel_col).Value
		RBIC_retro_income_group_2 = ObjExcel.Cells(RBIC_starting_excel_row + 8, current_excel_col).Value
		RBIC_prosp_income_group_2 = ObjExcel.Cells(RBIC_starting_excel_row + 9, current_excel_col).Value
		RBIC_ver_income_group_2 = ObjExcel.Cells(RBIC_starting_excel_row + 10, current_excel_col).Value
		RBIC_group_3 = ObjExcel.Cells(RBIC_starting_excel_row + 11, current_excel_col).Value
		RBIC_retro_income_group_3 = ObjExcel.Cells(RBIC_starting_excel_row + 12, current_excel_col).Value
		RBIC_prosp_income_group_3 = ObjExcel.Cells(RBIC_starting_excel_row + 13, current_excel_col).Value
		RBIC_ver_income_group_3 = ObjExcel.Cells(RBIC_starting_excel_row + 14, current_excel_col).Value
		RBIC_retro_hours = ObjExcel.Cells(RBIC_starting_excel_row + 15, current_excel_col).Value
		RBIC_prosp_hours = ObjExcel.Cells(RBIC_starting_excel_row + 16, current_excel_col).Value
		RBIC_exp_type_1 = ObjExcel.Cells(RBIC_starting_excel_row + 17, current_excel_col).Value
		RBIC_exp_retro_1 = ObjExcel.Cells(RBIC_starting_excel_row + 18, current_excel_col).Value
		RBIC_exp_prosp_1 = ObjExcel.Cells(RBIC_starting_excel_row + 19, current_excel_col).Value
		RBIC_exp_ver_1 = ObjExcel.Cells(RBIC_starting_excel_row + 20, current_excel_col).Value
		RBIC_exp_type_2 = ObjExcel.Cells(RBIC_starting_excel_row + 21, current_excel_col).Value
		RBIC_exp_retro_2 = ObjExcel.Cells(RBIC_starting_excel_row + 22, current_excel_col).Value
		RBIC_exp_prosp_2 = ObjExcel.Cells(RBIC_starting_excel_row + 23, current_excel_col).Value
		RBIC_exp_ver_2 = ObjExcel.Cells(RBIC_starting_excel_row + 24, current_excel_col).Value

		REST_starting_excel_row = 366
		REST_type = ObjExcel.Cells(REST_starting_excel_row, current_excel_col).Value
		REST_type_ver = ObjExcel.Cells(REST_starting_excel_row + 1, current_excel_col).Value
		REST_market = ObjExcel.Cells(REST_starting_excel_row + 2, current_excel_col).Value
		REST_market_ver = ObjExcel.Cells(REST_starting_excel_row + 3, current_excel_col).Value
		REST_owed = ObjExcel.Cells(REST_starting_excel_row + 4, current_excel_col).Value
		REST_owed_ver = ObjExcel.Cells(REST_starting_excel_row + 5, current_excel_col).Value
		REST_date = ObjExcel.Cells(REST_starting_excel_row + 6, current_excel_col).Value
		REST_status = ObjExcel.Cells(REST_starting_excel_row + 7, current_excel_col).Value
		REST_joint = ObjExcel.Cells(REST_starting_excel_row + 8, current_excel_col).Value
		REST_share_ratio = ObjExcel.Cells(REST_starting_excel_row + 9, current_excel_col).Value
		REST_agreement_date = ObjExcel.Cells(REST_starting_excel_row + 10, current_excel_col).Value

		SCHL_starting_excel_row = 377
		SCHL_status = ObjExcel.Cells(SCHL_starting_excel_row, current_excel_col).Value
		SCHL_ver = ObjExcel.Cells(SCHL_starting_excel_row + 1, current_excel_col).Value
		SCHL_type = ObjExcel.Cells(SCHL_starting_excel_row + 2, current_excel_col).Value
		SCHL_district_nbr = ObjExcel.Cells(SCHL_starting_excel_row + 3, current_excel_col).Value
		SCHL_kindergarten_start_date = ObjExcel.Cells(SCHL_starting_excel_row + 4, current_excel_col).Value
		SCHL_grad_date = ObjExcel.Cells(SCHL_starting_excel_row + 5, current_excel_col).Value
		SCHL_grad_date_ver = ObjExcel.Cells(SCHL_starting_excel_row + 6, current_excel_col).Value
		SCHL_primary_secondary_funding = ObjExcel.Cells(SCHL_starting_excel_row + 7, current_excel_col).Value
		SCHL_FS_eligibility_status = ObjExcel.Cells(SCHL_starting_excel_row + 8, current_excel_col).Value
		SCHL_higher_ed = ObjExcel.Cells(SCHL_starting_excel_row + 9, current_excel_col).Value

		SECU_starting_excel_row = 387
		SECU_type = ObjExcel.Cells(SECU_starting_excel_row, current_excel_col).Value
		SECU_pol_numb = ObjExcel.Cells(SECU_starting_excel_row + 1, current_excel_col).Value
		SECU_name = ObjExcel.Cells(SECU_starting_excel_row + 2, current_excel_col).Value
		SECU_cash_val = ObjExcel.Cells(SECU_starting_excel_row + 3, current_excel_col).Value
		SECU_date = ObjExcel.Cells(SECU_starting_excel_row + 4, current_excel_col).Value
		SECU_cash_ver = ObjExcel.Cells(SECU_starting_excel_row + 5, current_excel_col).Value
		SECU_face_val = ObjExcel.Cells(SECU_starting_excel_row + 6, current_excel_col).Value
		SECU_withdraw = ObjExcel.Cells(SECU_starting_excel_row + 7, current_excel_col).Value
		SECU_cash_count = ObjExcel.Cells(SECU_starting_excel_row + 8, current_excel_col).Value
		SECU_SNAP_count = ObjExcel.Cells(SECU_starting_excel_row + 9, current_excel_col).Value
		SECU_HC_count = ObjExcel.Cells(SECU_starting_excel_row + 10, current_excel_col).Value
		SECU_GRH_count = ObjExcel.Cells(SECU_starting_excel_row + 11, current_excel_col).Value
		SECU_IV_count = ObjExcel.Cells(SECU_starting_excel_row + 12, current_excel_col).Value
		SECU_joint = ObjExcel.Cells(SECU_starting_excel_row + 13, current_excel_col).Value
		SECU_share_ratio = ObjExcel.Cells(SECU_starting_excel_row + 14, current_excel_col).Value

		SHEL_starting_excel_row = 402
		SHEL_subsidized = ObjExcel.Cells(SHEL_starting_excel_row, current_excel_col).Value
		SHEL_shared = ObjExcel.Cells(SHEL_starting_excel_row + 1, current_excel_col).Value
		SHEL_paid_to = ObjExcel.Cells(SHEL_starting_excel_row + 2, current_excel_col).Value
		SHEL_rent_retro = ObjExcel.Cells(SHEL_starting_excel_row + 3, current_excel_col).Value
		SHEL_rent_retro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 4, current_excel_col).Value
		SHEL_rent_pro = ObjExcel.Cells(SHEL_starting_excel_row + 5, current_excel_col).Value
		SHEL_rent_pro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 6, current_excel_col).Value
		SHEL_lot_rent_retro = ObjExcel.Cells(SHEL_starting_excel_row + 7, current_excel_col).Value
		SHEL_lot_rent_retro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 8, current_excel_col).Value
		SHEL_lot_rent_pro = ObjExcel.Cells(SHEL_starting_excel_row + 9, current_excel_col).Value
		SHEL_lot_rent_pro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 10, current_excel_col).Value
		SHEL_mortgage_retro = ObjExcel.Cells(SHEL_starting_excel_row + 11, current_excel_col).Value
		SHEL_mortgage_retro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 12, current_excel_col).Value
		SHEL_mortgage_pro = ObjExcel.Cells(SHEL_starting_excel_row + 13, current_excel_col).Value
		SHEL_mortgage_pro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 14, current_excel_col).Value
		SHEL_insur_retro = ObjExcel.Cells(SHEL_starting_excel_row + 15, current_excel_col).Value
		SHEL_insur_retro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 16, current_excel_col).Value
		SHEL_insur_pro = ObjExcel.Cells(SHEL_starting_excel_row + 17, current_excel_col).Value
		SHEL_insur_pro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 18, current_excel_col).Value
		SHEL_taxes_retro = ObjExcel.Cells(SHEL_starting_excel_row + 19, current_excel_col).Value
		SHEL_taxes_retro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 20, current_excel_col).Value
		SHEL_taxes_pro = ObjExcel.Cells(SHEL_starting_excel_row + 21, current_excel_col).Value
		SHEL_taxes_pro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 22, current_excel_col).Value
		SHEL_room_retro = ObjExcel.Cells(SHEL_starting_excel_row + 23, current_excel_col).Value
		SHEL_room_retro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 24, current_excel_col).Value
		SHEL_room_pro = ObjExcel.Cells(SHEL_starting_excel_row + 25, current_excel_col).Value
		SHEL_room_pro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 26, current_excel_col).Value
		SHEL_garage_retro = ObjExcel.Cells(SHEL_starting_excel_row + 27, current_excel_col).Value
		SHEL_garage_retro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 28, current_excel_col).Value
		SHEL_garage_pro = ObjExcel.Cells(SHEL_starting_excel_row + 29, current_excel_col).Value
		SHEL_garage_pro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 30, current_excel_col).Value
		SHEL_subsidy_retro = ObjExcel.Cells(SHEL_starting_excel_row + 31, current_excel_col).Value
		SHEL_subsidy_retro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 32, current_excel_col).Value
		SHEL_subsidy_pro = ObjExcel.Cells(SHEL_starting_excel_row + 33, current_excel_col).Value
		SHEL_subsidy_pro_ver = ObjExcel.Cells(SHEL_starting_excel_row + 34, current_excel_col).Value

		SIBL_starting_excel_row = 437
		SIBL_group_1 = ObjExcel.Cells(SIBL_starting_excel_row, current_excel_col).Value
		SIBL_group_2 = ObjExcel.Cells(SIBL_starting_excel_row + 1, current_excel_col).Value
		SIBL_group_3 = ObjExcel.Cells(SIBL_starting_excel_row + 2, current_excel_col).Value

		SPON_starting_excel_row = 440
		SPON_type = ObjExcel.Cells(SPON_starting_excel_row, current_excel_col).Value
		SPON_ver = ObjExcel.Cells(SPON_starting_excel_row + 1, current_excel_col).Value
		SPON_name = ObjExcel.Cells(SPON_starting_excel_row + 2, current_excel_col).Value
		SPON_state = ObjExcel.Cells(SPON_starting_excel_row + 3, current_excel_col).Value

		STEC_starting_excel_row = 444
		STEC_type_1 = ObjExcel.Cells(STEC_starting_excel_row, current_excel_col).Value
		STEC_amt_1 = ObjExcel.Cells(STEC_starting_excel_row + 1, current_excel_col).Value
		STEC_actual_from_thru_months_1 = ObjExcel.Cells(STEC_starting_excel_row + 2, current_excel_col).Value
		STEC_ver_1 = ObjExcel.Cells(STEC_starting_excel_row + 3, current_excel_col).Value
		STEC_earmarked_amt_1 = ObjExcel.Cells(STEC_starting_excel_row + 4, current_excel_col).Value
		STEC_earmarked_from_thru_months_1 = ObjExcel.Cells(STEC_starting_excel_row + 5, current_excel_col).Value
		STEC_type_2 = ObjExcel.Cells(STEC_starting_excel_row + 6, current_excel_col).Value
		STEC_amt_2 = ObjExcel.Cells(STEC_starting_excel_row + 7, current_excel_col).Value
		STEC_actual_from_thru_months_2 = ObjExcel.Cells(STEC_starting_excel_row + 8, current_excel_col).Value
		STEC_ver_2 = ObjExcel.Cells(STEC_starting_excel_row + 9, current_excel_col).Value
		STEC_earmarked_amt_2 = ObjExcel.Cells(STEC_starting_excel_row + 10, current_excel_col).Value
		STEC_earmarked_from_thru_months_2 = ObjExcel.Cells(STEC_starting_excel_row + 11, current_excel_col).Value

		STIN_starting_excel_row = 456
		STIN_type_1 = ObjExcel.Cells(STIN_starting_excel_row, current_excel_col).Value
		STIN_amt_1 = ObjExcel.Cells(STIN_starting_excel_row + 1, current_excel_col).Value
		STIN_avail_date_1 = ObjExcel.Cells(STIN_starting_excel_row + 2, current_excel_col).Value
		STIN_months_covered_1 = ObjExcel.Cells(STIN_starting_excel_row + 3, current_excel_col).Value
		STIN_ver_1 = ObjExcel.Cells(STIN_starting_excel_row + 4, current_excel_col).Value
		STIN_type_2 = ObjExcel.Cells(STIN_starting_excel_row + 5, current_excel_col).Value
		STIN_amt_2 = ObjExcel.Cells(STIN_starting_excel_row + 6, current_excel_col).Value
		STIN_avail_date_2 = ObjExcel.Cells(STIN_starting_excel_row + 7, current_excel_col).Value
		STIN_months_covered_2 = ObjExcel.Cells(STIN_starting_excel_row + 8, current_excel_col).Value
		STIN_ver_2 = ObjExcel.Cells(STIN_starting_excel_row + 9, current_excel_col).Value

		STWK_starting_excel_row = 466
		STWK_empl_name = ObjExcel.Cells(STWK_starting_excel_row, current_excel_col).Value
		STWK_wrk_stop_date = ObjExcel.Cells(STWK_starting_excel_row + 1, current_excel_col).Value
		STWK_wrk_stop_date_verif = ObjExcel.Cells(STWK_starting_excel_row + 2, current_excel_col).Value
		STWK_inc_stop_date = ObjExcel.Cells(STWK_starting_excel_row + 3, current_excel_col).Value
		STWK_refused_empl_yn = ObjExcel.Cells(STWK_starting_excel_row + 4, current_excel_col).Value
		STWK_vol_quit = ObjExcel.Cells(STWK_starting_excel_row + 5, current_excel_col).Value
		STWK_ref_empl_date = ObjExcel.Cells(STWK_starting_excel_row + 6, current_excel_col).Value
		STWK_gc_cash = ObjExcel.Cells(STWK_starting_excel_row + 7, current_excel_col).Value
		STWK_gc_grh = ObjExcel.Cells(STWK_starting_excel_row + 8, current_excel_col).Value
		STWK_gc_fs = ObjExcel.Cells(STWK_starting_excel_row + 9, current_excel_col).Value
		STWK_fs_pwe = ObjExcel.Cells(STWK_starting_excel_row + 10, current_excel_col).Value
		STWK_maepd_ext = ObjExcel.Cells(STWK_starting_excel_row + 11, current_excel_col).Value

		UNEA_1_starting_excel_row = 478
		UNEA_1_inc_type = ObjExcel.Cells(UNEA_1_starting_excel_row, current_excel_col).Value
		UNEA_1_inc_verif = ObjExcel.Cells(UNEA_1_starting_excel_row + 1, current_excel_col).Value
		UNEA_1_claim_suffix = ObjExcel.Cells(UNEA_1_starting_excel_row + 2, current_excel_col).Value
		UNEA_1_start_date = ObjExcel.Cells(UNEA_1_starting_excel_row + 3, current_excel_col).Value
		UNEA_1_pay_freq = ObjExcel.Cells(UNEA_1_starting_excel_row + 4, current_excel_col).Value
		UNEA_1_inc_amount = ObjExcel.Cells(UNEA_1_starting_excel_row + 5, current_excel_col).Value

		UNEA_2_starting_excel_row = 484
		UNEA_2_inc_type = ObjExcel.Cells(UNEA_2_starting_excel_row, current_excel_col).Value
		UNEA_2_inc_verif = ObjExcel.Cells(UNEA_2_starting_excel_row + 1, current_excel_col).Value
		UNEA_2_claim_suffix = ObjExcel.Cells(UNEA_2_starting_excel_row + 2, current_excel_col).Value
		UNEA_2_start_date = ObjExcel.Cells(UNEA_2_starting_excel_row + 3, current_excel_col).Value
		UNEA_2_pay_freq = ObjExcel.Cells(UNEA_2_starting_excel_row + 4, current_excel_col).Value
		UNEA_2_inc_amount = ObjExcel.Cells(UNEA_2_starting_excel_row + 5, current_excel_col).Value

		UNEA_3_starting_excel_row = 490
		UNEA_3_inc_type = ObjExcel.Cells(UNEA_3_starting_excel_row, current_excel_col).Value
		UNEA_3_inc_verif = ObjExcel.Cells(UNEA_3_starting_excel_row + 1, current_excel_col).Value
		UNEA_3_claim_suffix = ObjExcel.Cells(UNEA_3_starting_excel_row + 2, current_excel_col).Value
		UNEA_3_start_date = ObjExcel.Cells(UNEA_3_starting_excel_row + 3, current_excel_col).Value
		UNEA_3_pay_freq = ObjExcel.Cells(UNEA_3_starting_excel_row + 4, current_excel_col).Value
		UNEA_3_inc_amount = ObjExcel.Cells(UNEA_3_starting_excel_row + 5, current_excel_col).Value

		WREG_starting_excel_row = 496
		WREG_fs_pwe = ObjExcel.Cells(WREG_starting_excel_row, current_excel_col).Value
		WREG_fset_status = ObjExcel.Cells(WREG_starting_excel_row + 1, current_excel_col).Value
		WREG_defer_fs = ObjExcel.Cells(WREG_starting_excel_row + 2, current_excel_col).Value
		WREG_fset_orientation_date = ObjExcel.Cells(WREG_starting_excel_row + 3, current_excel_col).Value
		WREG_fset_sanction_date = ObjExcel.Cells(WREG_starting_excel_row + 4, current_excel_col).Value
		WREG_num_sanctions = ObjExcel.Cells(WREG_starting_excel_row + 5, current_excel_col).Value
		WREG_abawd_status = ObjExcel.Cells(WREG_starting_excel_row + 6, current_excel_col).Value
		WREG_ga_basis = ObjExcel.Cells(WREG_starting_excel_row + 7, current_excel_col).Value

		'-------------------------------ACTUALLY FILLING OUT MAXIS
		
		'Goes to STAT/MEMB to associate a SSN to each member, this will be useful for UNEA/MEDI panels
		call navigate_to_screen("STAT", "MEMB")
		EMWriteScreen reference_number, 20, 76
		transmit
		EMReadScreen SSN_first, 3, 7, 42
		EMReadScreen SSN_mid, 2, 7, 46
		EMReadScreen SSN_last, 4, 7, 49
		
		'ACCT
		If ACCT_type <> "" then call write_panel_to_MAXIS_ACCT(ACCT_type, ACCT_numb, ACCT_location, ACCT_balance, ACCT_bal_ver, ACCT_date, ACCT_withdraw, ACCT_cash_count, ACCT_snap_count, ACCT_HC_count, ACCT_GRH_count, ACCT_IV_count, ACCT_joint_owner, ACCT_share_ratio, ACCT_interest_date_mo, ACCT_interest_date_yr)
		
		'ACUT
		If ACUT_shared <> "" then call write_panel_to_MAXIS_ACUT(ACUT_shared, ACUT_heat, ACUT_air, ACUT_electric, ACUT_fuel, ACUT_garbage, ACUT_water, ACUT_sewer, ACUT_other, ACUT_phone, ACUT_heat_verif, ACUT_air_verif, ACUT_electric_verif, ACUT_fuel_verif, ACUT_garbage_verif, ACUT_water_verif, ACUT_sewer_verif, ACUT_other_verif)
		
		'BUSI
		If BUSI_type <> "" then call write_panel_to_MAXIS_BUSI(busi_type, busi_start_date, busi_end_date, busi_cash_total_retro, busi_cash_total_prosp, busi_cash_total_ver, busi_IV_total_prosp, busi_IV_total_ver, busi_snap_total_retro, busi_snap_total_prosp, busi_snap_total_ver, busi_hc_total_prosp_a, busi_hc_total_ver_a, busi_hc_total_prosp_b, busi_hc_total_ver_b, busi_cash_exp_retro, busi_cash_exp_prosp, busi_cash_exp_ver, busi_IV_exp_prosp, busi_IV_exp_ver, busi_snap_exp_retro, busi_snap_exp_prosp, busi_snap_exp_ver, busi_hc_exp_prosp_a, busi_hc_exp_ver_a, busi_hc_exp_prosp_b, busi_hc_exp_ver_b, busi_retro_hours, busi_prosp_hours, busi_hc_total_est_a, busi_hc_total_est_b, busi_hc_exp_est_a, busi_hc_exp_est_b, busi_hc_hours_est)
		
		'CARS
		If CARS_type <> "" then call write_panel_to_MAXIS_CARS(cars_type, cars_year, cars_make, cars_model, cars_trade_in, cars_loan, cars_value_source, cars_ownership_ver, cars_amount_owed, cars_amount_owed_ver, cars_date, cars_owed_as_of, cars_use, cars_HC_benefit, cars_joint_owner, cars_share_ratio)
		
		'CASH
		If CASH_amount <> "" then call write_panel_to_MAXIS_CASH(cash_amount)
		
		'DCEX
		If DCEX_provider <> "" then call write_panel_to_MAXIS_DCEX(DCEX_provider, DCEX_reason, DCEX_subsidy, DCEX_child_number1, DCEX_child_number1_ver, DCEX_child_number1_retro, DCEX_child_number1_pro, DCEX_child_number2, DCEX_child_number2_ver, DCEX_child_number2_retro, DCEX_child_number2_pro, DCEX_child_number3, DCEX_child_number3_ver, DCEX_child_number3_retro, DCEX_child_number3_pro, DCEX_child_number4, DCEX_child_number4_ver, DCEX_child_number4_retro, DCEX_child_number4_pro, DCEX_child_number5, DCEX_child_number5_ver, DCEX_child_number5_retro, DCEX_child_number5_pro, DCEX_child_number6, DCEX_child_number6_ver, DCEX_child_number6_retro, DCEX_child_number6_pro)
		
		'DIET
		If DIET_mfip_1 <> "" or DIET_MSA_1 <> "" then call write_panel_to_MAXIS_DIET(DIET_mfip_1, DIET_mfip_1_ver, DIET_mfip_2, DIET_mfip_2_ver, DIET_msa_1, DIET_msa_1_ver, DIET_msa_2, DIET_msa_2_ver, DIET_msa_3, DIET_msa_3_ver, DIET_msa_4, DIET_msa_4_ver)
		
		'DISA
		If DISA_begin_date <> "" then call write_panel_to_MAXIS_DISA(disa_begin_date, disa_end_date, disa_cert_begin, disa_cert_end, disa_wavr_begin, disa_wavr_end, disa_grh_begin, disa_grh_end, disa_cash_status, disa_cash_status_ver, disa_snap_status, disa_snap_status_ver, disa_hc_status, disa_hc_status_ver, disa_waiver, disa_drug_alcohol)
		
		'DSTT
		If DSTT_ongoing_income <> "" then call write_panel_to_MAXIS_DSTT(DSTT_ongoing_income, DSTT_HH_income_stop_date, DSTT_income_expected_amt)
		
		'EATS
		If EATS_together <> "" then call write_panel_to_MAXIS_EATS(eats_together, eats_boarder, eats_group_one, eats_group_two, eats_group_three)
		
		
		'EMMA
		If EMMA_medical_emergency <> "" then call write_panel_to_MAXIS_EMMA(EMMA_medical_emergency, EMMA_health_consequence, EMMA_verification, EMMA_begin_date, EMMA_end_date)
		
		'EMPS
		If EMPS_memb_at_home <> "" then call write_panel_to_MAXIS_EMPS(EMPS_orientation_date, EMPS_orientation_attended, EMPS_good_cause, EMPS_sanc_begin, EMPS_sanc_end, EMPS_memb_at_home, EMPS_care_family, EMPS_crisis, EMPS_hard_employ, EMPS_under1, EMPS_DWP_date)
		
		'FACI
		If FACI_name <> "" then call write_panel_to_MAXIS_FACI(FACI_vendor_number, FACI_name, FACI_type, FACI_FS_eligible, FACI_FS_facility_type, FACI_date_in, FACI_date_out)
		
		'HCRE
		If HCRE_appl_addnd_date_input <> "" then call write_panel_to_MAXIS_HCRE(HCRE_appl_addnd_date_input,HCRE_retro_months_input,HCRE_recvd_by_service_date_input)
		
		'HEST
		If HEST_FS_choice_date <> "" then call write_panel_to_MAXIS_HEST(HEST_FS_choice_date, HEST_first_month, HEST_heat_air_retro, HEST_electric_retro, HEST_phone_retro, HEST_heat_air_pro, HEST_electric_pro, HEST_phone_pro)
		
		'IMIG
		If IMIG_imigration_status <> "" then call write_panel_to_MAXIS_IMIG(IMIG_imigration_status, IMIG_entry_date, IMIG_status_date, IMIG_status_ver, IMIG_status_LPR_adj_from, IMIG_nationality)
		
		'INSA
		If INSA_pers_coop_ohi <> "" then call write_panel_to_MAXIS_INSA(INSA_pers_coop_ohi,INSA_good_cause_status,INSA_good_cause_cliam_date,INSA_good_cause_evidence,INSA_coop_cost_effect,INSA_insur_name,INSA_prescrip_drug_cover,INSA_prescrip_end_date)
		
		'JOBS1
		If JOBS_1_inc_type <> "" then call write_panel_to_MAXIS_JOBS(JOBS_1_inc_type, JOBS_1_inc_verif, JOBS_1_employer_name, JOBS_1_inc_start, JOBS_1_wkly_hrs, JOBS_1_hrly_wage, JOBS_1_pay_freq)
		
		'JOBS2
		If JOBS_2_inc_type <> "" then call write_panel_to_MAXIS_JOBS(JOBS_2_inc_type, JOBS_2_inc_verif, JOBS_2_employer_name, JOBS_2_inc_start, JOBS_2_wkly_hrs, JOBS_2_hrly_wage, JOBS_2_pay_freq)
		
		'JOBS3
		If JOBS_3_inc_type <> "" then call write_panel_to_MAXIS_JOBS(JOBS_3_inc_type, JOBS_3_inc_verif, JOBS_3_employer_name, JOBS_3_inc_start, JOBS_3_wkly_hrs, JOBS_3_hrly_wage, JOBS_3_pay_freq)
		
		'MEDI
		If MEDI_claim_number_suffix <> "" then call write_panel_to_MAXIS_MEDI(SSN_first, SSN_mid, SSN_last, MEDI_claim_number_suffix, MEDI_part_A_premium, MEDI_part_B_premium, MEDI_part_A_begin_date, MEDI_part_B_begin_date)
		
		'MMSA
		If MMSA_liv_arr <> "" then call write_panel_to_MAXIS_MMSA(MMSA_liv_arr, MMSA_cont_elig, MMSA_spous_inc, MMSA_shared_hous)
		
		'MSUR
		If MSUR_begin_date <> "" then call write_panel_to_MAXIS_MSUR(MSUR_begin_date)
		
		'OTHR
		If OTHR_type <> "" then call write_panel_to_MAXIS_OTHR(OTHR_type, OTHR_cash_value, OTHR_cash_value_ver, OTHR_owed, OTHR_owed_ver, OTHR_date, OTHR_cash_count, OTHR_SNAP_count, OTHR_HC_count, OTHR_IV_count, OTHR_joint, OTHR_share_ratio)
			
		'PARE
		If PARE_child_1 <> "" then call write_panel_to_MAXIS_PARE(PARE_child_1, PARE_child_1_relation, PARE_child_1_verif, PARE_child_2, PARE_child_2_relation, PARE_child_2_verif, PARE_child_3, PARE_child_3_relation, PARE_child_3_verif, PARE_child_4, PARE_child_4_relation, PARE_child_4_verif, PARE_child_5, PARE_child_5_relation, PARE_child_5_verif, PARE_child_6, PARE_child_6_relation, PARE_child_6_verif)
		
		'PBEN 1
		If PBEN_1_IAA_date <> "" then call write_panel_to_MAXIS_PBEN(PBEN_1_referal_date, PBEN_1_type, PBEN_1_appl_date, PBEN_1_appl_ver, PBEN_1_IAA_date, PBEN_1_disp)
		
		'PBEN 2
		If PBEN_2_IAA_date <> "" then call write_panel_to_MAXIS_PBEN(PBEN_2_referal_date, PBEN_2_type, PBEN_2_appl_date, PBEN_2_appl_ver, PBEN_2_IAA_date, PBEN_2_disp)
		
		'PBEN 3
		If PBEN_3_IAA_date <> "" then call write_panel_to_MAXIS_PBEN(PBEN_3_referal_date, PBEN_3_type, PBEN_3_appl_date, PBEN_3_appl_ver, PBEN_3_IAA_date, PBEN_3_disp)
		
		'PDED
		If PDED_ <> "" then call write_panel_to_MAXIS_PDED(PDED_wid_deduction, PDED_adult_child_disregard, PDED_wid_disregard, PDED_unea_income_deduction_reason, PDED_unea_income_deduction_value, PDED_earned_income_deduction_reason, PDED_earned_income_deduction_value, PDED_ma_epd_inc_asset_limit, PDED_guard_fee, PDED_rep_payee_fee, PDED_other_expense, PDED_shel_spcl_needs, PDED_excess_need, PDED_restaurant_meals)
		
		'PREG
		If PREG_conception_date <> "" then call write_panel_to_MAXIS_PREG(PREG_conception_date, PREG_conception_date_ver, PREG_third_trimester_ver,PREG_due_date, PREG_multiple_birth)
		
		'RBIC
		If RBIC_ <> "" then call write_panel_to_MAXIS_RBIC(rbic_type, rbic_start_date, rbic_end_date, rbic_group_1, rbic_retro_income_group_1, rbic_prosp_income_group_1, rbic_ver_income_group_1, rbic_group_2, rbic_retro_income_group_2, rbic_prosp_income_group_2, rbic_ver_income_group_2, rbic_group_3, rbic_retro_income_group_3, rbic_prosp_income_group_3, rbic_ver_income_group_3, rbic_retro_hours, rbic_prosp_hours, rbic_exp_type_1, rbic_exp_retro_1, rbic_exp_prosp_1, rbic_exp_ver_1, rbic_exp_type_2, rbic_exp_retro_2, rbic_exp_prosp_2, rbic_exp_ver_2)
		
		'REST
		If REST_ <> "" then call write_panel_to_MAXIS_REST(rest_type, rest_type_ver, rest_market, rest_market_ver, rest_owed, rest_owed_ver, rest_date, rest_status, rest_joint, rest_share_ratio, rest_agreement_date)
		
		'SCHL
		If SCHL_ <> "" then call write_panel_to_MAXIS_SCHL(SCHL_status, SCHL_ver, SCHL_type, SCHL_district_nbr, SCHL_kindergarten_start_date, SCHL_grad_date, SCHL_grad_date_ver, SCHL_primary_secondary_funding, SCHL_FS_eligibility_status, SCHL_higher_ed)
		
		'SECU
		If SECU_ <> "" then call write_panel_to_MAXIS_SECU(secu_type, secu_pol_numb, secu_name, secu_cash_val, secu_date, secu_cash_ver, secu_face_val, secu_withdraw, secu_cash_count, secu_SNAP_count, secu_HC_count, secu_GRH_count, secu_IV_count, secu_joint, secu_share_ratio)
		
		'SHEL
		If SHEL_subsidized <> "" then call write_panel_to_MAXIS_SHEL(SHEL_subsidized, SHEL_shared, SHEL_paid_to, SHEL_rent_retro, SHEL_rent_retro_ver, SHEL_rent_pro, SHEL_rent_pro_ver, SHEL_lot_rent_retro, SHEL_lot_rent_retro_ver, SHEL_lot_rent_pro, SHEL_lot_rent_pro_ver, SHEL_mortgage_retro, SHEL_mortgage_retro_ver, SHEL_mortgage_pro, SHEL_mortgage_pro_ver, SHEL_insur_retro, SHEL_insur_retro_ver, SHEL_insur_pro, SHEL_insur_pro_ver, SHEL_taxes_retro, SHEL_taxes_retro_ver, SHEL_taxes_pro, SHEL_taxes_pro_ver, SHEL_room_retro, SHEL_room_retro_ver, SHEL_room_pro, SHEL_room_pro_ver, SHEL_garage_retro, SHEL_garage_retro_ver, SHEL_garage_pro, SHEL_garage_pro_ver, SHEL_subsidy_retro, SHEL_subsidy_retro_ver, SHEL_subsidy_pro, SHEL_subsidy_pro_ver)

		'SIBL
		If SIBL_group_1 <> "" then call write_panel_to_MAXIS_SIBL(SIBL_group_1, SIBL_group_2, SIBL_group_3)
		
		'SPON
		If SPON_ <> "" then call write_panel_to_MAXIS_SPON(SPON_type, SPON_ver, SPON_name, SPON_state)
		
		'STEC
		If STEC_ <> "" then call write_panel_to_MAXIS_STEC(STEC_type_1, STEC_amt_1, STEC_actual_from_thru_months_1, STEC_ver_1, STEC_earmarked_amt_1, STEC_earmarked_from_thru_months_1, STEC_type_2, STEC_amt_2, STEC_actual_from_thru_months_2, STEC_ver_2, STEC_earmarked_amt_2, STEC_earmarked_from_thru_months_2)
		
		'STIN
		If STIN_ <> "" then call write_panel_to_MAXIS_STIN(STIN_type_1, STIN_amt_1, STIN_avail_date_1, STIN_months_covered_1, STIN_ver_1, STIN_type_2, STIN_amt_2, STIN_avail_date_2, STIN_months_covered_2, STIN_ver_2)
		
		'STWK
		If STWK_ <> "" then call write_panel_to_MAXIS_STWK(STWK_empl_name, STWK_wrk_stop_date, STWK_wrk_stop_date_verif, STWK_inc_stop_date, STWK_refused_empl_yn, STWK_vol_quit, STWK_ref_empl_date, STWK_gc_cash, STWK_gc_grh, STWK_gc_fs, STWK_fs_pwe, STWK_maepd_ext)
		
		'UNEA 1
		If UNEA_1_inc_type <> "" then call write_panel_to_MAXIS_UNEA(UNEA_1_inc_type, UNEA_1_inc_verif, UNEA_1_claim_suffix, UNEA_1_start_date, UNEA_1_pay_freq, UNEA_1_inc_amount, SSN_first, SSN_mid, SSN_last)
		
		'UNEA 2
		If UNEA_2_inc_type <> "" then call write_panel_to_MAXIS_UNEA(UNEA_2_inc_type, UNEA_2_inc_verif, UNEA_2_claim_suffix, UNEA_2_start_date, UNEA_2_pay_freq, UNEA_2_inc_amount, SSN_first, SSN_mid, SSN_last)
		
		'UNEA 3
		If UNEA_3_inc_type <> "" then call write_panel_to_MAXIS_UNEA(UNEA_3_inc_type, UNEA_3_inc_verif, UNEA_3_claim_suffix, UNEA_3_start_date, UNEA_3_pay_freq, UNEA_3_inc_amount, SSN_first, SSN_mid, SSN_last)
				
		'WREG
		If WREG_fs_pwe <> "" then call write_panel_to_MAXIS_WREG(WREG_fs_pwe, WREG_fset_status, WREG_defer_fs, WREG_fset_orientation_date, WREG_fset_sanction_date, WREG_num_sanctions, WREG_abawd_status, WREG_ga_basis)
	
		'ABPS (must do after PARE, because the ABPS function checks PARE for a child list)
		If abps_supp_coop <> "" then call write_panel_to_MAXIS_ABPS(abps_supp_coop,abps_gc_status)
	
	Next

	'Gets back to self
	back_to_self

Next

'Ends here if the user selected to leave cases in PND2 status
If approve_case_dropdown = "no, but enter all STAT panels needed to approve" then 
	If XFER_check = checked then call transfer_cases(workers_to_XFER_cases_to, case_number_array)
	script_end_procedure("Success! Cases made, STAT panels added, and left in PND2 status, per your request.")
End if
	
'========================================================================APPROVAL========================================================================


FOR EACH case_number IN case_number_array
	If SNAP_application = True then 
		DO
			back_to_SELF
			EMWriteScreen "ELIG", 16, 43
			EMWriteScreen case_number, 18, 43
			EMWriteScreen appl_date_month, 20, 43
			EMWriteScreen appl_date_year, 20, 46
			EMWriteScreen "FS", 21, 70
			'========== This TRANSMIT sends the case to the FSPR screen ==========
			transmit
			EMReadScreen no_version, 10, 24, 2
		LOOP UNTIL no_version <> "NO VERSION"
		EMReadScreen is_case_approved, 10, 3, 3
		IF is_case_approved <> "UNAPPROVED" THEN
			back_to_SELF
		ELSE
		'========== This TRANSMIT sends the case to the FSCR screen ==========
			transmit
		'========== Reading for EXPEDITED STATUS ==========
			EMReadScreen is_case_expedited, 9, 4, 3
		'========== This TRANSMIT sends the case to the FSB1 screen ==========
			transmit
		'========== This TRANSMIT sends the case to the FSB2 screen ==========
			transmit
		'========== This TRANSMIT sends the case to the FSSM screen ==========
			transmit
			IF is_case_expedited <> "EXPEDITED" THEN
				DO
					EMWriteScreen "APP", 19, 70
					transmit
					EMReadScreen not_allowed, 11, 24, 18
					EMReadScreen locked_by_background, 6, 24, 19
					EMReadScreen what_is_next, 5, 16, 44
				LOOP UNTIL not_allowed <> "NOT ALLOWED" AND locked_by_background <> "LOCKED" OR what_is_next = "(Y/N)"
				DO
					EMReadScreen please_examine, 14, 4, 25
				LOOP UNTIL please_examine = "PLEASE EXAMINE"
				EMWriteScreen "Y", 16, 51
				transmit
				transmit
			ELSE
				DO
					EMWriteScreen "APP", 19, 70
					transmit
					EMReadScreen not_allowed, 11, 24, 18
					EMReadScreen locked_by_background, 6, 24, 19
					EMReadScreen what_is_next, 5, 16, 44
				LOOP UNTIL not_allowed <> "NOT ALLOWED" AND locked_by_background <> "LOCKED" OR what_is_next = "(Y/N)"
				DO
					EMReadScreen rei_benefit, 3, 15, 33
				LOOP UNTIL rei_benefit = "REI"
				EMWriteScreen "Y", 15, 60
				transmit
				DO
					EMReadScreen rei_confirm, 3, 14, 30
				LOOP UNTIL rei_confirm = "REI"
				EMWriteScreen "Y", 14, 62
				transmit
				DO
					EMReadScreen continue_with_approval, 5, 16, 44
				LOOP UNTIL continue_with_approval = "(Y/N)"
				EMWriteScreen "Y", 16, 51
				transmit
				transmit
			END IF
		END IF
		
		If HC_application = True then
		'<<<<<<<<<INSERT LOGIC HERE
		End if
		
		If cash_application = True then
		'<<<<<<<<<INSERT LOGIC HERE
		End if
		
	End if
	
	'Checks for WORK panel (Workforce One Referral), makes one with a week from now as the appointment date as a default (we can add a specific date/location checker as an enhancement
	EMReadScreen WORK_check, 4, 2, 51
	If WORK_check = "WORK" then
		call create_MAXIS_friendly_date(date, 7, 7, 59)
		EMWriteScreen "X", 7, 47
		transmit
		EMWriteScreen "X", 5, 9
		transmit
		transmit
		transmit
		'Special error handling for DHS and possibly multicounty agencies (don't have WF1 sites)
		EMReadScreen ES_provider_check, 2, 2, 37		'Looks for the ES in ES provider, indicating we're stuck on a screen
		If worker_county_code = "MULTICOUNTY" and ES_provider_check = "ES" then 
			'Clear out the X and get back to the SELF menu
			EMWriteScreen "_", 5, 9	
			transmit
			back_to_SELF
		End if
	End if
NEXT


If XFER_check = checked then call transfer_cases(workers_to_XFER_cases_to, case_number_array)
script_end_procedure("Success! Cases made and approved, per your request.")