'Required for statistical purposes===============================================================================
name_of_script = "BULK - SWKR LIST GENERATOR.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 45                      'manual run time in seconds
STATS_denomination = "C"       						 'C is for each CASE
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
call changelog_update("11/28/2016", "Initial version.", "Charles Potter, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CONNECTS TO MAXIS
EMConnect ""

'DIALOG TO DETERMINE WHERE TO GO IN MAXIS TO GET THE INFO
BeginDialog SWKR_list_generator_dialog, 0, 0, 156, 115, "SWKR list generator dialog"
  DropListBox 65, 5, 85, 15, "REPT/ACTV"+chr(9)+"REPT/REVS"+chr(9)+"REPT/REVW", REPT_panel
  EditBox 55, 25, 20, 15, MAXIS_footer_month
  EditBox 130, 25, 20, 15, MAXIS_footer_year
  EditBox 75, 45, 75, 15, worker_number
  ButtonGroup ButtonPressed
    OkButton 20, 95, 50, 15
    CancelButton 85, 95, 50, 15
  Text 5, 10, 55, 10, "Create list from:"
  Text 5, 30, 45, 10, "Footer month:"
  Text 85, 30, 40, 10, "Footer year:"
  Text 5, 50, 65, 10, "Worker number(s):"
  Text 5, 65, 145, 25, "Enter all 7 digits of each, (ex: x######). If entering multiple workers, separate each with a comma."
EndDialog

'DISPLAYS DIALOG
Dialog SWKR_list_generator_dialog
If buttonpressed = cancel then stopscript

'CHECKS FOR PASSWORD PROMPT/MAXIS STATUS
transmit
Call check_for_MAXIS(True)

'NAVIGATES BACK TO SELF TO FORCE THE FOOTER MONTH, THEN NAVIGATES TO THE SELECTED SCREEN
back_to_self
EMWriteScreen "________", 18, 43
call navigate_to_MAXIS_screen("rept", right(REPT_panel, 4))
If right(REPT_panel, 4) = "REVS" then
	current_month_plus_one = datepart("m", dateadd("m", 1, date))
	If len(current_month_plus_one) = 1 then current_month_plus_one = "0" & current_month_plus_one
	current_month_plus_one_year = datepart("yyyy", dateadd("m", 1, date))
	current_month_plus_one_year = right(current_month_plus_one_year, 2)
	EMWriteScreen current_month_plus_one, 20, 43
	EMWriteScreen current_month_plus_one_year, 20, 46
	transmit
	EMWriteScreen MAXIS_footer_month, 20, 55
	EMWriteScreen MAXIS_footer_year, 20, 58
	transmit
	MAXIS_footer_month = current_month_plus_one
	MAXIS_footer_year = current_month_plus_one_year
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
ObjExcel.Cells(1, 4).Value = "SWKR name"
objExcel.Cells(1, 4).Font.Bold = TRUE
objExcel.Cells(1, 4).ColumnWidth = 35
ObjExcel.Cells(1, 5).Value = "Copy of Notice?"
objExcel.Cells(1, 5).Font.Bold = TRUE
objExcel.Cells(1, 5).ColumnWidth = 20

'Splitting array for use by the for...next statement
worker_number_array = split(worker_number, ",")

For each worker in worker_number_array

	If trim(worker) = "" then exit for
	worker_ID = trim(worker)
	If REPT_panel = "REPT/ACTV" then 'THE REPT PANEL HAS THE worker NUMBER IN DIFFERENT COLUMNS. THIS WILL DETERMINE THE CORRECT COLUMN FOR THE worker NUMBER TO GO
		worker_ID_col = 13
	Else
		worker_ID_col = 6
	End if
	EMReadScreen default_worker_number, 3, 21, worker_ID_col 'CHECKING THE CURRENT worker NUMBER. IF IT DOESN'T NEED TO CHANGE IT WON'T. OTHERWISE, THE SCRIPT WILL INPUT THE CORRECT NUMBER.
	If ucase(worker_ID) <> default_worker_number then
		EMWriteScreen worker_ID, 21, worker_ID_col
		transmit
	End if

	'THIS DO...LOOP DUMPS THE CASE NUMBER AND NAME OF EACH CLIENT INTO A SPREADSHEET
	Do
		EMReadScreen last_page_check, 21, 24, 02
		'This Do...loop checks for the password prompt.
		Do
			EMReadScreen password_prompt, 38, 2, 23
			IF password_prompt = "ACF2/CICS PASSWORD VERIFICATION PROMPT" then MsgBox "You are locked out of your case. Type your password then try again."
		Loop until password_prompt <> "ACF2/CICS PASSWORD VERIFICATION PROMPT"

		row = 7 'defining the row to look at
		Do
			If REPT_panel = "REPT/ACTV" then
				EMReadScreen MAXIS_case_number, 8, row, 12 'grabbing case number
				EMReadScreen client_name, 18, row, 21 'grabbing client name
			Else
				EMReadScreen MAXIS_case_number, 8, row, 6 'grabbing case number
				EMReadScreen client_name, 15, row, 16 'grabbing client name
			End if
			ObjExcel.Cells(excel_row, 1).Value = worker_ID
			ObjExcel.Cells(excel_row, 2).Value = trim(MAXIS_case_number)
			ObjExcel.Cells(excel_row, 3).Value = trim(client_name)
			excel_row = excel_row + 1
			row = row + 1
		Loop until row = 19 or trim(MAXIS_case_number) = ""
		PF8 'going to the next screen
	Loop until last_page_check = "THIS IS THE LAST PAGE"
	STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
Next

'NOW THE SCRIPT IS CHECKING STAT/AREP FOR EACH CASE.----------------------------------------------------------------------------------------------------
excel_row = 2 'Resetting the case row to investigate.

do until ObjExcel.Cells(excel_row, 2).Value = "" 'shuts down when there's no more case numbers
	SWKR_name = "" 'Resetting this variable in case a SWKR cannot be found.
	MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
	If MAXIS_case_number = "" then exit do

	'This Do...loop gets back to SELF
	back_to_self
	'NAVIGATES TO STAT/SWKR
	call navigate_to_MAXIS_screen("STAT", "SWKR")

	'NAVIGATES TO SWKR, READS THE NAME AND NOTICE Y/N, AND ADDS TO SPREADSHEET
	EMReadScreen SWKR_name, 34, 6, 32
	swkr_name = replace(swkr_name, "_", "")
	ObjExcel.Cells(excel_row, 4).Value = swkr_name
	EMReadScreen NOTC_Y_N, 1, 15, 63
	If NOTC_Y_N = "_" then NOTC_Y_N = ""
	ObjExcel.Cells(excel_row, 5).Value = NOTC_Y_N

	excel_row = excel_row + 1 'setting up the script to check the next row.
loop

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created.")
