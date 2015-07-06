'Gathering stats
name_of_script = "BULK - CHECK SNAP FOR GA RCA.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN		'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF default_directory = "C:\DHS-MAXIS-Scripts\Script Files\" THEN			'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		ELSEIF beta_agency = "" or beta_agency = True then							'If you're a beta agency, you should probably use the beta branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/BETA/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else																		'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
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
					"URL: " & FuncLib_URL
					script_end_procedure("Script ended due to error connecting to GitHub.")
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

BeginDialog check_snap_dlg, 0, 0, 166, 80, "Check SNAP for GA/RCA"
  EditBox 105, 10, 50, 15, worker_number
  CheckBox 15, 35, 145, 10, "Or check here to run this on all workers.", all_worker_check
  ButtonGroup ButtonPressed
    OkButton 35, 60, 50, 15
    CancelButton 85, 60, 50, 15
  Text 15, 15, 85, 10, "Enter worker X number(s)"
EndDialog

EMConnect ""

benefit_month = datepart("M", dateadd("M", 1, date))
IF len(benefit_month) <> 2 THEN benefit_month = "0" & benefit_month
benefit_year = datepart("YYYY", dateadd("M", 1, date))
benefit_year = right(benefit_year, 2)

back_to_SELF
EMWriteScreen benefit_month, 20, 43
EMWriteScreen benefit_year, 20, 46

DO
	DIALOG check_snap_dlg
		IF ButtonPressed = 0 THEN stopscript
LOOP UNTIL (worker_number = "" AND all_worker_check = 1) OR (all_worker_check = 0 AND worker_number <> "")

IF all_worker_check = 1 THEN
	CALL navigate_to_screen("REPT", "USER")
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
ObjExcel.Cells(1, 4).Value = "Error on SNAP?"

excel_row = 2
FOR EACH worker IN worker_array
	IF worker = "" THEN EXIT FOR
	CALL navigate_to_screen("REPT", "ACTV")
	EMWriteScreen worker, 21, 13
	transmit

	CALL find_variable("User: ", current_user, 7)
	IF ucase(worker) = ucase(current_user) THEN PF7

	rept_actv_row = 7
	DO
		DO
			EMReadScreen last_page, 21, 24, 2
			EMReadScreen case_number, 8, rept_actv_row, 12
			case_number = trim(case_number)
			EMReadScreen snap_status, 1, rept_actv_row, 61
			EMReadScreen cash_status, 1, rept_actv_row, 54
			EMReadScreen cash_prog, 2, rept_actv_row, 51
			EMReadScreen client_name, 20, rept_actv_row, 21
			IF snap_status = "A" AND cash_status = "A" AND (cash_prog = "RC" OR cash_prog = "GA") THEN 
				case_array = case_array & case_number & " "
				objExcel.Cells(excel_row, 1).Value = worker
				objExcel.Cells(excel_row, 2).Value = case_number
				objExcel.Cells(excel_row, 3).Value = client_name
				excel_row = excel_row + 1
			END IF
			rept_actv_row = rept_actv_row + 1
		LOOP UNTIL rept_actv_row = 19
			PF8
			rept_actv_row = 7
	LOOP UNTIL case_number = "" OR last_page = "THIS IS THE LAST PAGE"
NEXT

case_array = trim(case_array)
case_array = split(case_array)

excel_row = 2
'navigates to ELIG to determine if RCA or GA has been correctly fiated into SNAP budget.
FOR EACH case_number IN case_array
	ga_status = ""
	ga_amount = ""
	rca_status = ""
	rca_amount = ""
	cash_prog = ""
	pa_amount = ""
	CALL navigate_to_screen("ELIG", "FS")
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
	CALL navigate_to_screen("CASE", "CURR")
	CALL find_variable("GA: ", ga_status, 6)
	IF ga_status = "ACTIVE" OR ga_status = "APP CL" THEN
		cash_prog = "GA"
	ELSE
		CALL find_variable("RCA: ", rca_status, 6)
		IF rca_status = "ACTIVE" OR rca_status = "APP CL" THEN cash_prog = "RCA"
	END IF
	IF cash_prog = "GA" THEN
		CALL navigate_to_screen("ELIG", "GA")
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
		ga_amount = trim(ga_amount)
		IF pa_amount <> ga_amount THEN
			CALL navigate_to_screen("STAT", "REVW")
			ERRR_screen_check
			EMReadScreen cash_revw_date, 8, 9, 37
			EMReadScreen snap_revw_date, 8, 9, 57
			bene_date = benefit_month & "/" & benefit_year
			cash_revw_date = replace(cash_revw_date, " 01 ", "/")
			snap_revw_date = replace(snap_revw_date, " 01 ", "/")
			IF bene_date = cash_revw_date OR bene_date = snap_revw_date THEN
				objExcel.Cells(excel_row, 4).Value = "REVW MONTH"
			ELSEIF bene_date <> cash_revw_date AND bene_date <> snap_revw_date THEN
				objExcel.Cells(excel_row, 4).Value = ("Yes, GA. SNAP Budg = " & pa_amount & "; GA Amount = " & ga_amount)
			END IF
		ELSEIF pa_amount = ga_amount THEN
			objExcel.Cells(excel_row, 4).Value = ("Budgetted for SNAP: " & pa_amount & "; GA Amount: " & ga_amount)
		END IF
	ELSEIF cash_prog = "RCA" THEN
		CALL navigate_to_screen("ELIG", "RCA")
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
			CALL navigate_to_screen("STAT", "REVW")
			ERRR_screen_check
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

script_end_procedure("Done")
