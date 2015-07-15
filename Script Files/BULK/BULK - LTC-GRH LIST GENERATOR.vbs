'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - LTC-GRH LIST GENERATOR.vbs"
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

'CONNECTS TO BlueZone
EMConnect ""

'DIALOG TO DETERMINE WHERE TO GO IN MAXIS TO GET THE INFO
BeginDialog LTC_GRH_list_generator_dialog, 0, 0, 156, 115, "LTC-GRH list generator dialog"
  DropListBox 65, 5, 85, 15, "REPT/ACTV"+chr(9)+"REPT/REVS"+chr(9)+"REPT/REVW", REPT_panel
  EditBox 55, 25, 20, 15, footer_month
  EditBox 130, 25, 20, 15, footer_year
  EditBox 75, 45, 75, 15, worker_number
  ButtonGroup ButtonPressed
    OkButton 20, 95, 50, 15
    CancelButton 85, 95, 50, 15
  Text 5, 10, 55, 10, "Create list from:"
  Text 5, 30, 45, 10, "Footer month:"
  Text 85, 30, 40, 10, "Footer year:"
  Text 5, 50, 65, 10, "Worker number(s):"
  Text 5, 65, 145, 25, "Enter last three digits of each, (ex: x100###). If entering multiple workers, separate each with a comma."
EndDialog


'DISPLAYS DIALOG
Dialog LTC_GRH_list_generator_dialog
If buttonpressed = cancel then stopscript

'CHECKS FOR PASSWORD PROMPT/MAXIS STATUS
transmit
MAXIS_check_function

'NAVIGATES BACK TO SELF TO FORCE THE FOOTER MONTH, THEN NAVIGATES TO THE SELECTED SCREEN
back_to_self
EMWriteScreen "________", 18, 43
call navigate_to_screen("rept", right(REPT_panel, 4))
If right(REPT_panel, 4) = "REVS" then
	current_month_plus_one = datepart("m", dateadd("m", 1, date))
	If len(current_month_plus_one) = 1 then current_month_plus_one = "0" & current_month_plus_one
	current_month_plus_one_year = datepart("yyyy", dateadd("m", 1, date))
	current_month_plus_one_year = right(current_month_plus_one_year, 2)
	EMWriteScreen current_month_plus_one, 20, 43
	EMWriteScreen current_month_plus_one_year, 20, 46
	transmit
	EMWriteScreen footer_month, 20, 55
	EMWriteScreen footer_year, 20, 58
	transmit
	footer_month = current_month_plus_one
	footer_year = current_month_plus_one_year
End if

'CHECKS TO MAKE SURE WE'VE MOVED PAST SELF MENU. IF WE HAVEN'T, THE SCRIPT WILL STOP. AN ERROR MESSAGE SHOULD DISPLAY ON THE BOTTOM OF THE MENU.
EMReadScreen SELF_check, 4, 2, 50
If SELF_check = "SELF" then script_end_procedure("Can't get past SELF menu. Check error message and try again!")

'DEFINES THE EXCEL_ROW VARIABLE FOR WORKING WITH THE SPREADSHEET
excel_row = 2

'OPENS A NEW EXCEL SPREADSHEET
Set objExcel = CreateObject("Excel.Application") 
objExcel.Visible = True  
Set objWorkbook = objExcel.Workbooks.Add() 

'FORMATS THE EXCEL SPREADSHEET WITH THE HEADERS, AND SETS THE COLUMN WIDTH
ObjExcel.Cells(1, 1).Value = "WORKER"
objExcel.Cells(1, 1).Font.Bold = TRUE
ObjExcel.Cells(1, 2).Value = "M#"
objExcel.Cells(1, 2).Font.Bold = TRUE
objExcel.Cells(1, 2).ColumnWidth = 9
ObjExcel.Cells(1, 3).Value = "Name"
objExcel.Cells(1, 3).Font.Bold = TRUE
objExcel.Cells(1, 3).ColumnWidth = 27
ObjExcel.Cells(1, 4).Value = "FACI name"
objExcel.Cells(1, 4).Font.Bold = TRUE
objExcel.Cells(1, 4).ColumnWidth = 47
ObjExcel.Cells(1, 5).Value = "AREP name"
objExcel.Cells(1, 5).Font.Bold = TRUE
objExcel.Cells(1, 5).ColumnWidth = 35
ObjExcel.Cells(1, 6).Value = "Waiver type on first DISA panel found"
objExcel.Cells(1, 6).Font.Bold = TRUE
objExcel.Cells(1, 6).ColumnWidth = 35

'Splitting array for use by the for...next statement
worker_number_array = split(worker_number, ",")

For each worker in worker_number_array

	If trim(worker) = "" then exit for

	worker_ID = worker_county_code & trim(worker)

	If REPT_panel = "REPT/ACTV" then 'THE REPT PANEL HAS THE worker NUMBER IN DIFFERENT COLUMNS. THIS WILL DETERMINE THE CORRECT COLUMN FOR THE worker NUMBER TO GO
		worker_ID_col = 13
	Else
		worker_ID_col = 6
	End if
	EMReadScreen default_worker_number, 7, 21, worker_ID_col 'CHECKING THE CURRENT worker NUMBER. IF IT DOESN'T NEED TO CHANGE IT WON'T. OTHERWISE, THE SCRIPT WILL INPUT THE CORRECT NUMBER.
	If ucase(worker_ID) <> default_worker_number then
		EMWriteScreen worker_ID, 21, worker_ID_col
		transmit
	End if


	'THIS DO...LOOP DUMPS THE CASE NUMBER AND NAME OF EACH CLIENT INTO A SPREADSHEET
	Do



		'This Do...loop checks for the password prompt.
		Do
			EMReadScreen password_prompt, 38, 2, 23
			IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then try again."
		Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"

		EMReadScreen last_page_check, 21, 24, 02
		row = 7 'defining the row to look at
		Do
			If REPT_panel = "REPT/ACTV" then
				EMReadScreen case_number, 8, row, 12 'grabbing case number
				EMReadScreen client_name, 18, row, 21 'grabbing client name
			Else
				EMReadScreen case_number, 8, row, 6 'grabbing case number
				EMReadScreen client_name, 15, row, 16 'grabbing client name
			End if
			If trim(case_number) = "" then exit do	'quits if we're out of cases
			ObjExcel.Cells(excel_row, 1).Value = worker_ID
			ObjExcel.Cells(excel_row, 2).Value = trim(case_number)
			ObjExcel.Cells(excel_row, 3).Value = trim(client_name)
			excel_row = excel_row + 1
			row = row + 1
		Loop until row = 19 or trim(case_number) = ""

		PF8 'going to the next screen

	Loop until last_page_check = "THIS IS THE LAST PAGE"

next

'NOW THE SCRIPT IS CHECKING STAT/FACI FOR EACH CASE.----------------------------------------------------------------------------------------------------

excel_row = 2 'Resetting the case row to investigate.

do until ObjExcel.Cells(excel_row, 1).Value = "" 'shuts down when there's no more case numbers
	FACI_name = "" 'Resetting this variable in case a FACI cannot be found.
	case_number = ObjExcel.Cells(excel_row, 2).Value 
	If case_number = "" then exit do

	'This Do...loop gets back to SELF
	back_to_self

	'NAVIGATES TO STAT/FACI for the correct footer month
	EMWriteScreen footer_month, 20, 43
	EMWriteScreen footer_year, 20, 46
	transmit
	call navigate_to_screen("STAT", "FACI")

	'CHECKS FOR ERROR PRONE CASES
	ERRR_screen_check


	'LOOKS FOR MULTIPLE STAT/FACI PANELS, GOES TO THE MOST RECENT ONE
	Do
		EMReadScreen FACI_current_panel, 1, 2, 73
		EMReadScreen FACI_total_check, 1, 2, 78
		EMReadScreen in_year_check_01, 4, 14, 53
		EMReadScreen in_year_check_02, 4, 15, 53
		EMReadScreen in_year_check_03, 4, 16, 53
		EMReadScreen in_year_check_04, 4, 17, 53
		EMReadScreen in_year_check_05, 4, 18, 53
		EMReadScreen out_year_check_01, 4, 14, 77
		EMReadScreen out_year_check_02, 4, 15, 77
		EMReadScreen out_year_check_03, 4, 16, 77
		EMReadScreen out_year_check_04, 4, 17, 77
		EMReadScreen out_year_check_05, 4, 18, 77
		If (in_year_check_01 <> "____" and out_year_check_01 = "____") or (in_year_check_02 <> "____" and out_year_check_02 = "____") or (in_year_check_03 <> "____" and out_year_check_03 = "____") or (in_year_check_04 <> "____" and out_year_check_04 = "____") or (in_year_check_05 <> "____" and out_year_check_05 = "____") then
			currently_in_FACI = True
			exit do
		Elseif FACI_current_panel = FACI_total_check then 
			currently_in_FACI = False
			exit do
		Else
			transmit
		End if
	Loop until FACI_current_panel = FACI_total_check

	'GETS FACI NAME AND PUTS IT IN SPREADSHEET, IF CLIENT IS IN FACI.
	If currently_in_FACI = True then
		EMReadScreen FACI_name, 30, 6, 43
		ObjExcel.Cells(excel_row, 4).Value = trim(replace(FACI_name, "_", ""))
	End if

	'NAVIGATES TO AREP, READS THE NAME, AND ADDS TO SPREADSHEET
	EMWriteScreen "AREP", 20, 71
	transmit
	EMReadScreen AREP_name, 37, 4, 32
	AREP_name = replace(AREP_name, "_", "")
	ObjExcel.Cells(excel_row, 5).Value = AREP_name

	'Navigates to DISA and checks the waiver type
	EMWriteScreen "DISA", 20, 71
	transmit
	EMReadScreen DISA_waiver_type, 1, 14, 59
	If DISA_waiver_type = "_" then DISA_waiver_type = ""
	ObjExcel.Cells(excel_row, 6).Value = DISA_waiver_type

	excel_row = excel_row + 1 'setting up the script to check the next row.
loop

MsgBox "Success! Your list has been created."

script_end_procedure("")
