'Required for statistical purposes===============================================================================
name_of_script = "BULK - GRH PROFESSIONAL NEED.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 51                      'manual run time in seconds
STATS_denomination = "C"       				'C is for each CASE
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
call changelog_update("01/08/2018", "Script updated to also gather waiver types from STAT/DISA.", "Ilse Ferris, Hennepin County")
call changelog_update("03/31/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CONNECTS TO BlueZone
EMConnect ""
all_workers_check = 1        'autochecking as this is the default setting

'DIALOG TO DETERMINE WHERE TO GO IN MAXIS TO GET THE INFO
BeginDialog GRH_Prof_Need_dialog, 0, 0, 266, 80, "GRH Professional Need Dialog"
  EditBox 70, 25, 190, 15, worker_number
  CheckBox 10, 65, 135, 10, "Click here to run for the entire agency.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 155, 60, 50, 15
    CancelButton 210, 60, 50, 15
  Text 5, 30, 60, 10, "Worker number(s):"
  Text 5, 45, 250, 10, "Enter7 digits of each worker number, (ex: x######), seperated by a comma."
  Text 10, 10, 250, 10, "This script will gather Professional Need Information for GRH active cases."
EndDialog

'DISPLAYS DIALOG
Do 
	Do 	
		err_msg = ""
		Dialog GRH_Prof_Need_dialog
		If buttonpressed = cancel then stopscript
		If trim(worker_number) = "" and all_workers_check = 0 then err_msg = err_msg & vbnewline & "* Enter at least one worker number."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine		
	Loop until err_msg = ""	
Call check_for_password(are_we_passworded_out)
LOOP UNTIL check_for_password(are_we_passworded_out) = False

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next
	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting up the Excel spreadsheet
ObjExcel.Cells(1, 1).Value = "Worker"
ObjExcel.Cells(1, 2).Value = "MAXIS Case #"
ObjExcel.Cells(1, 3).Value = "Facility name"
ObjExcel.Cells(1, 4).Value = "Facility type"
ObjExcel.Cells(1, 5).Value = "GRH plan required"
ObjExcel.Cells(1, 6).Value = "Plan verified"
ObjExcel.Cells(1, 7).Value = "Cty app placement"
ObjExcel.Cells(1, 8).Value = "Approval cty"
ObjExcel.Cells(1, 9).Value = "GRH DOC amount"
ObjExcel.Cells(1,10).Value = "GRH rate"
ObjExcel.Cells(1,11).Value = "GRH plan dates"
ObjExcel.Cells(1,11).Value = "Waiver type"

excel_row = 2 
 

'formatting the cells
FOR i = 1 to 11
	objExcel.Cells(1, i).Font.Bold = True		'bold font
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

For each worker in worker_array
    back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("REPT", "ACTV")
    EMWriteScreen worker, 21, 13
    transmit

	'THIS DO...LOOP DUMPS THE CASE NUMBER OF EACH CLIENT INTO A SPREADSHEET THAT IS ACTIVE ON GRH
	Do
		EMReadScreen last_page_check, 21, 24, 02
		row = 7 'defining the row to look at
		Do
			EMReadScreen GRH_prog, 1, row, 70
			If GRH_prog = "A" then 
				EMReadScreen MAXIS_case_number, 8, row, 12 'grabbing case number
				If trim(MAXIS_case_number) = "" then exit do	'quits if we're out of cases
				ObjExcel.Cells(excel_row, 1).Value = worker
				ObjExcel.Cells(excel_row, 2).Value = trim(MAXIS_case_number)
				excel_row = excel_row + 1
			End if 
			row = row + 1
			STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
		Loop until row = 19 or trim(MAXIS_case_number) = ""
		PF8 'going to the next screen
	Loop until last_page_check = "THIS IS THE LAST PAGE"
next

'NOW THE SCRIPT IS CHECKING STAT/FACI FOR EACH CASE.----------------------------------------------------------------------------------------------------
excel_row = 2 'Resetting the case row to investigate.

Do
	MAXIS_case_number= objExcel.cells(excel_row, 2).Value	're-establishing the case number to use for the case
    If trim(MAXIS_case_number) = "" then exit do
	
	'This Do...loop gets back to SELF
	back_to_self
	call navigate_to_MAXIS_screen("STAT", "FACI")
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
	If PRIV_check = "PRIV" then
		priv_case_list = priv_case_list & "|" & MAXIS_case_number
		SET objRange = objExcel.Cells(excel_row, 1).EntireRow
		objRange.Delete				'row gets deleted since it will get added to the priv case list at end of script 
		IF excel_row = 3 then 
			excel_row = excel_row
		Else 
			excel_row = excel_row - 1
		End if
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
    End if 
    
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
		If (in_year_check_01 <> "____" and out_year_check_01 = "____") or (in_year_check_02 <> "____" and out_year_check_02 = "____") or _
		(in_year_check_03 <> "____" and out_year_check_03 = "____") or (in_year_check_04 <> "____" and out_year_check_04 = "____") or (in_year_check_05 <> "____" and out_year_check_05 = "____") then
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
		EMReadScreen FACI_type, 2, 7, 43
		
		'List of FACI types
		IF FACI_type = "41" then FACI_type = "41: NF-I"
		IF FACI_type = "42" then FACI_type = "42: NF-II"
		IF FACI_type = "43" then FACI_type = "43: ICF-DD"
		IF FACI_type = "44" then FACI_type = "44: Short stay in NF-I"
		IF FACI_type = "45" then FACI_type = "45: Short stay in NF-II"
		IF FACI_type = "46" then FACI_type = "46: Short stay in ICF-DD"
		IF FACI_type = "47" then FACI_type = "47: RTC - Not IMD"
		IF FACI_type = "48" then FACI_type = "48: Medical Hospital"
		IF FACI_type = "49" then FACI_type = "49: MSOP"
		IF FACI_type = "50" then FACI_type = "50: IMD/RTC"
		IF FACI_type = "51" then FACI_type = "51: Rule 31 CD_IMD"
		IF FACI_type = "52" then FACI_type = "52: Rule 36 MI-IMD"
		IF FACI_type = "53" then FACI_type = "53: IMD Hospitals"
		IF FACI_type = "55" then FACI_type = "55: Adult Foster Care/Rule 203"
		IF FACI_type = "56" then FACI_type = "56: GRH (Not FC or Rule 36)"
		IF FACI_type = "57" then FACI_type = "57: Rule 36 MI - Non-IMD"
		IF FACI_type = "60" then FACI_type = "60: Non-GRH"
		IF FACI_type = "61" then FACI_type = "61: Rule 31 CD - Non-IMD"
		IF FACI_type = "67" then FACI_type = "67: Family Violence Shelter"
		IF FACI_type = "68" then FACI_type = "68: County Correctional Facility"
		IF FACI_type = "69" then FACI_type = "69: Non-Cty Adult Correctional"
		
		EMReadScreen GRH_plan_req, 1, 11, 52
		EMReadScreen GRH_plan_verif, 1, 11, 71
		EMReadScreen county_placement, 1, 12, 52
		EMReadScreen approval_county, 2, 12, 71
		EMReadScreen GRH_DOC_amt, 8, 13, 45
		EMReadScreen GRH_rate, 1, 14, 34
		
		ObjExcel.Cells(excel_row, 3).Value = trim(replace(FACI_name, "_", ""))
		ObjExcel.Cells(excel_row, 4).Value = trim(replace(FACI_type, "_", ""))
		ObjExcel.Cells(excel_row, 5).Value = trim(replace(GRH_plan_req, "_", ""))
		ObjExcel.Cells(excel_row, 6).Value = trim(replace(GRH_plan_verif, "_", ""))
		ObjExcel.Cells(excel_row, 7).Value = trim(replace(county_placement, "_", ""))
		ObjExcel.Cells(excel_row, 8).Value = trim(replace(approval_county, "_", ""))
		ObjExcel.Cells(excel_row, 9).Value = trim(replace(GRH_DOC_amt, "_", ""))
		ObjExcel.Cells(excel_row, 10).Value = trim(replace(GRH_rate, "_", ""))
	End if
 	
	Call navigate_to_MAXIS_screen("STAT", "DISA")
	EMReadScreen GRH_begin_date, 10, 9, 47
	EMReadScreen GRH_end_date, 10, 9, 69
	
	'begin dates on DISA for GRH plan
	If GRH_begin_date = "__ __ ____" then 
		GRH_begin_date = ""
	Else 
		GRH_begin_date = replace(GRH_begin_date, " ", "/")
	End if 
	'end dates on DISA for GRH plan
	If GRH_end_date = "__ __ ____" then 
		GRH_end_date = ""
	Else 
		GRH_end_date = replace(GRH_end_date, " ", "/")
	End if
	
 	GRH_plan_date = GRH_begin_date & " - " & GRH_end_date
	If trim(GRH_plan_date) = "-" then GRH_plan_date = ""
	
 	ObjExcel.Cells(excel_row, 11).Value = GRH_plan_date
	
	'checks the waiver type
	EMReadScreen DISA_waiver_type, 1, 14, 59
	If DISA_waiver_type = "_" then DISA_waiver_type = ""
	ObjExcel.Cells(excel_row, 11).Value = DISA_waiver_type
	
	excel_row = excel_row + 1 'setting up the script to check the next row.
LOOP UNTIL objExcel.Cells(excel_row, 2).Value = ""	'Loops until there are no more cases in the Excel list

'Query date/time/runtime info
objExcel.Cells(1, 13).Font.Bold = TRUE
objExcel.Cells(2, 13).Font.Bold = TRUE
ObjExcel.Cells(1, 13).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, 13).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, 14).Value = now
ObjExcel.Cells(2, 14).Value = timer - query_start_time

'formatting the cells
FOR i = 1 to 14
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created.")