'Required for statistical purposes===============================================================================
name_of_script = "BULK - HOMELESS DISCREPANCY.vbs"
start_time = timer
STATS_counter = 1         'sets the stats counter at one
STATS_manualtime = 90      'manual run time in seconds
STATS_denomination = "C"  'i is for each Item
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
call changelog_update("04/25/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================


'THE SCRIPT----------------------------------------------------------------------------------------------------
EMConnect ""

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 286, 110, "Homeless discrepancy dialog"
  EditBox 70, 50, 210, 15, worker_number
  CheckBox 5, 90, 155, 10, "Select all active workers in the agency", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 175, 85, 50, 15
    CancelButton 230, 85, 50, 15
  Text 5, 70, 275, 10, "Enter the fulll 7-digit worker number, separate each with a comma if more than one."
  Text 5, 55, 60, 10, "Worker number(s):"
  Text 15, 20, 255, 20, "To find discrepancies for cases that are identified as homeless, but have shelter and/or utility costs or an address that is not general delivery."
  GroupBox 5, 5, 275, 40, "Purpose of the script"
EndDialog
DO
	DO
		err_msg = ""
		Dialog Dialog1
		Cancel_without_confirmation
		If worker_number = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Enter a valid worker number."
		if worker_number <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Enter a worker number OR select the entire agency, not both."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Starting the query start time (for the query runtime at the end)
query_start_time = timer

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas
	'formatting array
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(x1_number)		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & ", " & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next
	'Split worker_array
	worker_array = split(worker_array, ", ")
End if

'start of the FOR...next loop
For each worker in worker_array
    back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("REPT", "ACTV")
    EMWriteScreen worker, 21, 13
    transmit

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 8
	If has_content_check <> " " then
		Do						'Grabbing each case number on screen
			row = 7		'Set variable for next do...loop
			'Checking for the last page of cases.
			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
			Do
				EMReadScreen MAXIS_case_number, 8, row, 12	'Reading case number
				MAXIS_case_number = trim(MAXIS_case_number)
				If MAXIS_case_number = "" then exit do			'Exits do if we reach the end

				'Cash requires different handling due to containing multiple program types in one column
				EMReadScreen cash_status, 9, row, 51
				cash_status = trim(cash_status)

				EMReadScreen snap_status, 1, row, 61

				If snap_status = "A" then
					add_to_array = True
				Elseif instr(cash_status, "MF A") then
					add_to_array = True
				Else
					add_to_array = False
				End if

				If add_to_array = True then
					'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and 	stops if we've seen this one before.
					If MAXIS_case_number <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
					all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & ",")
				End if
				row = row + 1
				MAXIS_case_number = ""			'Blanking out variable
			Loop until row = 19
			PF8
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

all_case_numbers_array = trim(all_case_numbers_array)
If right(all_case_numbers_array, 1) = "," then all_case_numbers_array = left(all_case_numbers_array, len(all_case_numbers_array) - 1)
case_number_array = split(all_case_numbers_array, ",")

discrep_amt = 0
DIM cases_array
ReDim cases_array(10, 0)

'constants for the array
const work_num 	 	= 0
const case_num 	 	= 1
const homeless 	 	= 2
const live_sit	 	= 3
const HEST		 	= 4
const SHEL 			= 5
const addr_one 	 	= 6
const addr_two		= 7
const addr_city		= 8
const addr_STATE	= 9
const addr_ZIP		= 10

'msgbox case_number_list
For each MAXIS_case_number in case_number_array
    IF MAXIS_case_number = "" then exit for
    Call navigate_to_MAXIS_screen ("STAT", "ADDR")
    EMReadScreen priv_check, 4, 2, 50
    If priv_check = "SELF" then
    	add_to_array = False
    Else
    	'Reading and cleaning up Residence address
		EMReadScreen worker, 7, 21, 21
    	EMReadScreen addr_line_1, 22, 6, 43
    	EMReadScreen addr_line_2, 22, 7, 43
    	EMReadScreen city, 15, 8, 43
    	EMReadScreen state, 2, 8, 66
    	EMReadScreen zip_code, 5, 9, 43
    	addr_line_1 = replace(addr_line_1, "_", "")
    	addr_line_2 = replace(addr_line_2, "_", "")
    	city = replace(city, "_", "")
    	State = replace(State, "_", "")
    	Zip_code = replace(Zip_code, "_", "")
    	'Reading homeless code
    	EMReadScreen homeless_code, 1, 10, 43
		homeless_code = replace(homeless_code, "_", "")
    	EMReadScreen living_sit, 2, 11, 43
		living_sit = replace(living_sit, "_", "")
    End if

	'Addding applicable cases to the array
	If homeless_code = "Y" then
		add_to_array = True
	elseif living_sit = "07" or living_sit = "08" or living_sit = "09" or living_sit = "10" then
 		add_to_array = True
	Elseif instr(addr_line_1, "GENERAL DELIVERY") or instr(addr_line_1, "GEN DELIVERY") or instr(addr_line_1, "330 12TH AVENUE S") then
		add_to_array = True
	Else
		add_to_array = False
	End if

	IF add_to_array = True then
		Redim Preserve cases_array(10,  discrep_amt)
		cases_array (work_num, 			discrep_amt) = worker
		cases_array (case_num, 			discrep_amt) = MAXIS_case_number
		cases_array (homeless, 			discrep_amt) = homeless_code
		cases_array (live_sit, 			discrep_amt) = living_sit
		cases_array (addr_one, 			discrep_amt) = addr_line_1
		cases_array (addr_two, 			discrep_amt) = addr_line_2
		cases_array (addr_city, 		discrep_amt) = city
		cases_array (addr_STATE, 		discrep_amt) = state
		cases_array (addr_ZIP, 			discrep_amt) = zip_code

		Call navigate_to_MAXIS_screen("STAT", "HEST")		'<<<<< Navigates to STAT/HEST
		EMReadScreen HEST_heat, 6, 13, 75 					'<<<<< Pulls information from the prospective side of HEAT/AC standard allowance

		IF trim(HEST_heat) <> "" then						'<<<<< If there is an amount on the hest line then the electric and phone allowances are not used
			Hest_costs = "Heat $" & HEST_heat & ","
		Else
			EMReadScreen HEST_elect, 6, 14, 75				'<<<<< Pulls information from prospective side of Electric standard if HEAT/AC is not used
			EMReadScreen HEST_phone, 6, 15, 75				'<<<<< Pulls information from prospective side of Phone standard if HEAT/AC is not used
			If trim(HEST_elect) <> "" then Hest_costs = Hest_costs & "Elec $" & HEST_elect & ","
			If trim(HEST_phone) <> "" then Hest_costs = Hest_costs & "Phone $" & HEST_phone & ","
		End If
		'takes the last comma off of Hest_costs when autofilled into dialog if more more than one app date is found and additional app is selected
		If right(Hest_costs, 1) = "," THEN Hest_costs = left(Hest_costs, len(Hest_costs) - 1)
		cases_array (HEST, discrep_amt) = HEST_costs

		Call navigate_to_MAXIS_screen("STAT", "SHEL")		'<<<<< Goes to SHEL for this person
		EMReadScreen rent_verif, 2, 11, 67
		If rent_verif <> "__" and rent_verif <> "NO" and rent_verif <> "?_" then EMReadScreen rent, 8, 11, 56
		If rent_verif = "__" or rent_verif = "NO" or rent_verif = "?_" then rent = "0"		'<<<<< Gets rent amount
		EMReadScreen lot_rent_verif, 2, 12, 67

		If lot_rent_verif <> "__" and lot_rent_verif <> "NO" and lot_rent_verif <> "?_" then EMReadScreen lot_rent, 8, 12, 56
		If lot_rent_verif = "__" or lot_rent_verif = "NO" or lot_rent_verif = "?_" then lot_rent = "0"		'<<<<< gets Lot Rent amount
		EMReadScreen mortgage_verif, 2, 13, 67

		If mortgage_verif <> "__" and mortgage_verif <> "NO" and mortgage_verif <> "?_" then EMReadScreen mortgage, 8, 13, 56
		If mortgage_verif = "__" or mortgage_verif = "NO" or mortgage_verif = "?_" then mortgage = "0"		'<<<<<< gets Mortgage amount
		EMReadScreen insurance_verif, 2, 14, 67

		If insurance_verif <> "__" and insurance_verif <> "NO" and insurance_verif <> "?_" then EMReadScreen insurance, 8, 14, 56
		If insurance_verif = "__" or insurance_verif = "NO" or insurance_verif = "?_" then insurance = "0"	'<<<<<< gets insurance amount and adds it to the class property
		EMReadScreen taxes_verif, 2, 15, 67

		If taxes_verif <> "__" and taxes_verif <> "NO" and taxes_verif <> "?_" then EMReadScreen taxes, 8, 15, 56
		If taxes_verif = "__" or taxes_verif = "NO" or taxes_verif = "?_" then taxes = "0"				'<<<<<<< gets taxes amount and adds it to the class property
		EMReadScreen room_verif, 2, 16, 67

		If room_verif <> "__" and room_verif <> "NO" and room_verif <> "?_" then EMReadScreen room, 8, 16, 56
		If room_verif = "__" or room_verif = "NO" or room_verif = "?_" then room = "0"						'<<<<<<< gets room/board amount
		EMReadScreen garage_verif, 2, 17, 67

		If garage_verif <> "__" and garage_verif <> "NO" and garage_verif <> "?_" then EMReadScreen garage, 8, 17, 56
		If garage_verif = "__" or garage_verif = "NO" or garage_verif = "?_" then garage = "0"				'<<<<<<< gets garage amount

		shel_costs = abs(rent) + abs(lot_rent) + abs(mortgage) + abs(insurance) + abs(taxes) + abs(room) + abs(garage)
		shel_costs = "$" & shel_costs
		'adding shel costs to the array
		cases_array (SHEL, discrep_amt) = shel_costs

		discrep_amt = discrep_amt + 1
		STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
		'blanking out variables
		HEST_costs = ""
		shel_costs = ""
		total_rent = ""
		rent = ""
		lot_rent = ""
		mortgage = ""
		insurance = ""
		taxes = ""
		room = ""
		garage = ""
	END If
next

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Creating columns
objExcel.Cells(1, 1).Value  = "X Number"
objExcel.Cells(1, 2).Value  = "Case Number"
objExcel.Cells(1, 3).Value  = "Homeless?"
objExcel.Cells(1, 4).Value  = "Living Situation"
objExcel.Cells(1, 5).Value  = "Utilities"
objExcel.Cells(1, 6).Value  = "Shelter costs"
objExcel.Cells(1, 7).Value  = "ADDRESS LINE 1"
objExcel.Cells(1, 8).Value  = "ADDRESS LINE 2"
objExcel.Cells(1, 9).Value  = "CITY"
objExcel.Cells(1, 10).Value = "STATE"
objExcel.Cells(1, 11).Value = "ZIP CODE"

FOR i = 1 to 11									'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

excel_row = 2

For i = 0 to Ubound(cases_array, 2)
	objExcel.Cells(excel_row,  1).Value = cases_array(work_num,		i)
	objExcel.Cells(excel_row,  2).Value = cases_array(case_num, 	i)
	objExcel.Cells(excel_row,  3).Value = cases_array(homeless, 	i)
	objExcel.Cells(excel_row,  4).Value = cases_array(live_sit,		i)
	objExcel.Cells(excel_row,  5).Value = cases_array(HEST,			i)
	objExcel.Cells(excel_row,  6).Value = cases_array(SHEL, 		i)
	objExcel.Cells(excel_row,  7).Value = cases_array(addr_one,		i)
	objExcel.Cells(excel_row,  8).Value = cases_array(addr_two,		i)
	objExcel.Cells(excel_row,  9).Value = cases_array(addr_city,	i)
	objExcel.Cells(excel_row, 10).Value = cases_array(addr_STATE,	i)
	objExcel.Cells(excel_row, 11).Value = cases_array(addr_ZIP,		i)
	excel_row = excel_row + 1
NEXT

'Query date/time/runtime info
objExcel.Cells(1, 12).Font.Bold = TRUE
ObjExcel.Cells(1, 12).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, 13).Value = now
objExcel.Cells(2, 12).Font.Bold = TRUE
ObjExcel.Cells(2, 12).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, 13).Value = timer - query_start_time

FOR i = 1 to 13									'formatting the cells'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)

script_end_procedure("Success!!")
