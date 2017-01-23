'Required for statistical purposes===============================================================================
name_of_script = "BULK - CHECK SNAP FOR GA RCA.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 69                               'manual run time in seconds
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

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("11/28/2016", "Added safety functionality for if MAXIS is passworded out.", "Casey Love, Ramsey County")
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


BeginDialog check_snap_dlg, 0, 0, 166, 100, "Check SNAP for GA/RCA"
  EditBox 100, 10, 55, 15, worker_number
  CheckBox 10, 35, 145, 10, "Or check here to run this on all workers.", all_worker_check
  CheckBox 10, 60, 145, 10, "Chere here to add supervisor name to list.", supervisor_check
  ButtonGroup ButtonPressed
    OkButton 35, 80, 50, 15
    CancelButton 85, 80, 50, 15
  Text 10, 10, 85, 20, "Enter worker X number(s) (7 digit format)"
EndDialog


EMConnect ""

Call check_for_MAXIS(True)

benefit_month = datepart("M", dateadd("M", 1, date))
IF len(benefit_month) <> 2 THEN benefit_month = "0" & benefit_month
benefit_year = datepart("YYYY", dateadd("M", 1, date))
benefit_year = right(benefit_year, 2)

back_to_SELF
EMWriteScreen benefit_month, 20, 43
EMWriteScreen benefit_year, 20, 46

DO
	DO
		DIALOG check_snap_dlg
		IF ButtonPressed = 0 THEN stopscript
	LOOP UNTIL (worker_number = "" AND all_worker_check = 1) OR (all_worker_check = 0 AND worker_number <> "")
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in 

IF all_worker_check = 1 THEN
	CALL navigate_to_MAXIS_screen("REPT", "USER")
	PF5

	rept_user_row = 7
	DO
		EMReadScreen worker_number, 7, rept_user_row, 5
		worker_number = trim(worker_number)
		IF worker_number <> "" THEN worker_array = worker_array & worker_number & " "
		rept_user_row = rept_user_row + 1
		IF rept_user_row = 19 THEN
			rept_user_row = 7
			PF8
		END IF
		EMReadScreen last_page, 21, 24, 2
	LOOP UNTIL worker_number = "" OR last_page = "THIS IS THE LAST PAGE"

	worker_array = trim(worker_array)
	worker_array = split(worker_array)
ELSE
	worker_array = split(worker_number, ", ")
END IF

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the first 3 col as worker, case number, and name
ObjExcel.Cells(1, 1).Value = "X Number"
ObjExcel.Cells(1, 2).Value = "CASE NUMBER"
ObjExcel.Cells(1, 3).Value = "NAME"
ObjExcel.Cells(1, 4).Value = "RCA Discrepancy?"
ObjExcel.Cells(1, 5).Value = "GA Discrepancy?"
objExcel.Cells(1, 6).Value = "GA in SNAP Budget"
objExcel.Cells(1, 7).Value = "GA Monthly Grant"
objExcel.Cells(1, 8).Value = "GA Issuance Amt"

FOR i = 1 TO 8
	objExcel.Cells(1, i).Font.Bold = TRUE
NEXT

excel_row = 2
FOR EACH worker IN worker_array
	IF worker = "" THEN EXIT FOR
	CALL navigate_to_MAXIS_screen("REPT", "ACTV")
	EMWriteScreen worker, 21, 13
	transmit

	CALL find_variable("User: ", current_user, 7)
	IF ucase(worker) = ucase(current_user) THEN PF7

	rept_actv_row = 7
	DO
		DO
			EMReadScreen last_page, 21, 24, 2
			EMReadScreen MAXIS_case_number, 8, rept_actv_row, 12
			MAXIS_case_number = trim(MAXIS_case_number)
			EMReadScreen snap_status, 1, rept_actv_row, 61
			EMReadScreen cash_status, 1, rept_actv_row, 54
			EMReadScreen cash_prog, 2, rept_actv_row, 51
			EMReadScreen client_name, 20, rept_actv_row, 21
			IF snap_status = "A" AND cash_status = "A" AND (cash_prog = "RC" OR cash_prog = "GA") THEN
				case_array = case_array & MAXIS_case_number & " "
				objExcel.Cells(excel_row, 1).Value = worker
				objExcel.Cells(excel_row, 2).Value = MAXIS_case_number
				objExcel.Cells(excel_row, 3).Value = client_name
				excel_row = excel_row + 1
				STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
			END IF
			rept_actv_row = rept_actv_row + 1
		LOOP UNTIL rept_actv_row = 19
			PF8
			rept_actv_row = 7
	LOOP UNTIL MAXIS_case_number = "" OR last_page = "THIS IS THE LAST PAGE"
NEXT

case_array = trim(case_array)
case_array = split(case_array)

excel_row = 2
'navigates to ELIG to determine if RCA or GA has been correctly fiated into SNAP budget.
FOR EACH MAXIS_case_number IN case_array
	ga_status = ""
	ga_amount = ""
	rca_status = ""
	rca_amount = ""
	cash_prog = ""
	pa_amount = ""
	CALL navigate_to_MAXIS_screen("ELIG", "FS")
	EMReadScreen approved, 8, 3, 3
	EMReadScreen version, 2, 2, 12
	version = trim(version)
	version = version - 1
	IF len(version) <> 2 THEN version = "0" & version
	IF approved <> "APPROVED" THEN
		EMWriteScreen version, 19, 78
		transmit
	END IF

	EMWriteScreen "FSB1", 19, 70
	transmit

	CALL find_variable("PA Grants..............$", pa_amount, 10)
	pa_amount = replace(pa_amount, "_", "")
	pa_amount = trim(pa_amount)
	IF pa_amount = "" THEN pa_amount = "0.00"
	CALL navigate_to_MAXIS_screen("CASE", "CURR")
	CALL find_variable("GA: ", ga_status, 6)
	IF ga_status = "ACTIVE" OR ga_status = "APP CL" THEN
		cash_prog = "GA"
	ELSE
		CALL find_variable("RCA: ", rca_status, 6)
		IF rca_status = "ACTIVE" OR rca_status = "APP CL" THEN cash_prog = "RCA"
	END IF
	IF cash_prog = "GA" THEN
		CALL navigate_to_MAXIS_screen("ELIG", "GA")
		EMReadScreen approved, 8, 3, 3
		EMReadScreen version, 2, 2, 12
		version = trim(version)
		version = version - 1
		IF len(version) <> 2 THEN version = "0" & version
		IF approved <> "APPROVED" THEN
			EMWriteScreen version, 20, 78
			transmit
		END IF
		EMWriteScreen "GASM", 20, 70
		transmit
			CALL find_variable("Monthly Grant............$", ga_amount, 9)
			CALL find_variable("Amount To Be Paid........$", ga_to_be_paid, 9)
		ga_amount = trim(ga_amount)
		ga_to_be_paid = trim(ga_to_be_paid)
		IF pa_amount <> ga_amount OR pa_amount <> ga_to_be_paid THEN
			CALL navigate_to_MAXIS_screen("STAT", "REVW")
			EMReadScreen cash_revw_date, 8, 9, 37
			EMReadScreen snap_revw_date, 8, 9, 57
			bene_date = benefit_month & "/" & benefit_year
			cash_revw_date = replace(cash_revw_date, " 01 ", "/")
			snap_revw_date = replace(snap_revw_date, " 01 ", "/")
			IF bene_date = cash_revw_date OR bene_date = snap_revw_date THEN
				objExcel.Cells(excel_row, 5).Value = "REVW MONTH"
			ELSEIF bene_date <> cash_revw_date AND bene_date <> snap_revw_date THEN
				objExcel.Cells(excel_row, 5).Value = ("Yes")
				objExcel.Cells(excel_row, 6).Value = ("SNAP Budg = " & pa_amount)
				objExcel.Cells(excel_row, 7).Value = ("Mo Grant = " & ga_amount)
				objExcel.Cells(excel_row, 8).Value = ("Amt Paid = " & ga_to_be_paid)
			END IF
		ELSEIF pa_amount = ga_amount AND pa_amount = ga_to_be_paid THEN
			objExcel.Cells(excel_row, 5).Value = ("No")
			objExcel.Cells(excel_row, 6).Value = ("SNAP Budg = " & pa_amount)
			objExcel.Cells(excel_row, 7).Value = ("Mo Grant = " & ga_amount)
			objExcel.Cells(excel_row, 8).Value = ("Amt Paid = " & ga_to_be_paid)
		END IF
	ELSEIF cash_prog = "RCA" THEN
		CALL navigate_to_MAXIS_screen("ELIG", "RCA")
		EMReadScreen approved, 8, 3, 3
		EMReadScreen version, 2, 2, 12
		version = trim(version)
		version = version - 1
		IF len(version) <> 2 THEN version = "0" & version
		IF approved <> "APPROVED" THEN
			EMWriteScreen version, 19, 78
			transmit
		END IF
			EMWriteScreen "RCSM", 19, 70
		transmit

		CALL find_variable("Grant Amount..............$", rca_amount, 10)
		rca_amount = trim(rca_amount)
		IF pa_amount <> rca_amount THEN
			CALL navigate_to_MAXIS_screen("STAT", "REVW")
			EMReadScreen cash_revw_date, 8, 9, 37
			EMReadScreen snap_revw_date, 8, 9, 57
			bene_date = benefit_month & "/" & benefit_year
			cash_revw_date = replace(cash_revw_date, " 01 ", "/")
			snap_revw_date = replace(snap_revw_date, " 01 ", "/")
			IF bene_date = cash_revw_date OR bene_date = snap_revw_date THEN
				objExcel.Cells(excel_row, 4).Value = "REVW MONTH"
			ELSEIF bene_date <> cash_revw_date AND bene_date <> snap_revw_date THEN
				objExcel.Cells(excel_row, 4).Value = ("Yes, RCA. SNAP Budg = " & pa_amount & "; RCA Amount = " & rca_amount)
			END IF
		ELSEIF pa_amount = rca_amount THEN
			objExcel.Cells(excel_row, 4).Value = ("Budgetted for SNAP: " & pa_amount & "; RCA Amount: " & rca_amount)
		END IF
	ELSEIF cash_prog = "SET TO CLOSE" THEN
		objExcel.Cells(excel_row, 4).Value = ("CASH set to close")
	END IF

	excel_row = excel_row + 1

NEXT

FOR i = 1 to 8
	objExcel.Columns(i).AutoFit()
NEXT

IF supervisor_check = 1 THEN
	'Adding additional manual time to the stats counter. I have timed this out to be about 25 seconds per case.
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
script_end_procedure("Done")