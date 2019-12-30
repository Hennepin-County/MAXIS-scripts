'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - GA ADVANCED AGE IDENTIFIER.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = "100"                'manual run time in seconds
STATS_denomination = "C"       			'C is for each CASE
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
call changelog_update("05/25/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT----------------------------------------------------------------------------------------------------
'Connects to BlueZone
EMConnect ""

footer_month = CM_plus_2_mo 
footer_year  = CM_plus_2_yr
all_workers_check = checked 

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 216, 115, "GA advanced age identifier"
  EditBox 75, 20, 135, 15, worker_number
  EditBox 170, 60, 20, 15, footer_month
  EditBox 190, 60, 20, 15, footer_year
  CheckBox 5, 80, 150, 10, "Check here to run this query county-wide.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 105, 95, 50, 15
    CancelButton 160, 95, 50, 15
  Text 5, 40, 210, 20, "Enter all 7 digits of your workers' x1 numbers (ex: x######), separated by a comma."
  Text 5, 25, 65, 10, "Worker(s) to check:"
  Text 70, 65, 100, 10, "Upcoming review month/year:"
  Text 10, 5, 125, 10, "***PULL REPT DATA INTO EXCEL***"
EndDialog
'Shows dialog
Do
	Do
		err_msg = ""
		Dialog Dialog1
		Cancel_without_confirmation
		If (all_workers_check = 0 AND worker_number = "") then err_msg = err_msg & vbNewLine & "* Please enter at least one worker number." 'allows user to select the all workers check, and not have worker number be ""					
		If IsNumeric(footer_month) = False or len(footer_month) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a 2-digit valid footer month."									
		If IsNumeric(footer_year) = False or len(footer_year) <> 2 then err_msg = err_msg & vbNewLine & "* Enter a 2-digit valid footer year."	
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine										
  	LOOP until err_msg = ""
	Call check_for_password(are_we_passworded_out)
Loop until check_for_password(are_we_passworded_out) = False		'loops until user is password-ed out

MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
'establishing the renewal date
renewal_date = footer_month & "/" & footer_year

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

actv_cases = 0

back_to_self
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46

For each worker in worker_array
    back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("REPT", "ACTV")
    EMWriteScreen worker, 21, 13
    transmit

	'Skips workers with no info
	EMReadScreen has_content_check, 1, 7, 8
	If has_content_check <> " " then
		Do						'Grabbing each case number on screen
			MAXIS_row = 7		'Set variable for next do...loop
			'Checking for the last page of cases.
			EMReadScreen last_page_check, 21, 24, 2	'because on REPT/ACTV it displays right away, instead of when the second F8 is sent
			Do
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 12	'Reading case number
				MAXIS_case_number = trim(MAXIS_case_number)
				If MAXIS_case_number = "" then exit do			'Exits do if we reach the end

				'Cash requires different handling due to containing multiple program types in one column
				EMReadScreen cash_status, 9, MAXIS_row, 51
				cash_status = trim(cash_status)
				If instr(cash_status, "GA A") then 
					EMReadScreen review_month, 2, MAXIS_row, 42		         'Reading review month
					EMReadScreen review_year, 2, MAXIS_row, 48		         'Reading review year
					review_date = review_month & "/" & review_year
					If renewal_date = review_date then 
						'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and 	stops if we've seen this one before.
						If MAXIS_case_number <> "" and instr(all_case_numbers_array, MAXIS_case_number) <> 0 then exit do
						all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & ",")
						actv_cases = actv_cases + 1
					End if 
				End if 
				MAXIS_row = MAXIS_row + 1
				MAXIS_case_number = ""			'Blanking out variable
			Loop until MAXIS_row = 19
			PF8
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

all_case_numbers_array = trim(all_case_numbers_array)
If right(all_case_numbers_array, 1) = "," then all_case_numbers_array = left(all_case_numbers_array, len(all_case_numbers_array) - 1)
case_number_array = split(all_case_numbers_array, ",")

entry_number = 0
Dim GA_array()
Redim GA_array(8, 0)

'Constants for the array
const basket_numb 	= 1
const case_numb 	= 2
const cl_name	 	= 3
const memb_numb		= 4
const cl_age		= 5
const wreg_panel	= 6
const dis_date		= 7
const dis_verifs	= 8

'msgbox "Active cases: " & actv_cases

For each MAXIS_case_number in case_number_array
	IF trim(MAXIS_case_number) = "" then exit for 
	'msgbox MAXIS_case_number
	back_to_self
	EMWriteScreen CM_mo, 20, 43
	EMWriteScreen CM_yr, 20, 46
	Call navigate_to_MAXIS_screen("CASE", "PERS")
    transmit
    EMReadScreen priv_check, 4, 2, 50
    If priv_check = "SELF" then
    	advanced_age = False 
	Else 
	    MAXIS_row = 10
	    Do
	    	EMReadScreen cash_code, 1, MAXIS_row, 48
			If trim(cash_code) = "" then exit do 
	    	If cash_code = "A" then 
				'msgbox cash_code
	    		EMReadScreen pers_ref_number, 2, MAXIS_row, 3
	    		member_number_list = pers_ref_number & ", "
	    	End if
	    	MAXIS_row = MAXIS_row + 3			'information is 3 rows apart
	    	If MAXIS_row = 19 then
	    		PF8
	    		MAXIS_row = 10					'changes MAXIS row if more than one page exists
	    	END if
	    	EMReadScreen last_PERS_page, 21, 24, 2
	    LOOP until last_PERS_page = "THIS IS THE LAST PAGE"
	    
	    member_number_list = trim(member_number_list)
	    If right(member_number_list, 1) = "," then member_number_list = left(member_number_list, len(member_number_list) - 1)
	    member_array= split(member_number_list, ",")
	    
	    Call navigate_to_MAXIS_screen("STAT", "MEMB")
	    For each memb_number in member_array 
	    	'msgbox memb_number
	    	EMWriteScreen memb_number, 20, 76				'enters member number
	    	transmit
	    	EMReadScreen client_age, 3, 8, 76
	    	client_age = trim(client_age)
			'msgbox client_age
	    	If client_age => 55 then 
	    		over_55 = true
	    		'STAT WREG PORTION
	    		Call navigate_to_MAXIS_screen("STAT", "WREG")
	    		EMWriteScreen memb_number, 20, 76				'enters member number
	    		transmit
				EMReadScreen client_name, 40, 4, 37
				client_name = trim(client_name)
				EMReadScreen fset_code, 2, 8, 50
	    		EMReadScreen abawd_code, 2, 13, 50			
	    		EMReadScreen ga_code, 2, 15, 50							''Reading the WREG coding for GA advance age 
				wreg_code = fset_code & "-" & abawd_code & "  GA: " & ga_coding 
				
	    		'STAT DISA PORTION
	    		Call navigate_to_MAXIS_screen("STAT", "DISA")
	    		EMWriteScreen memb_number, 20, 76				'enters member number
	    		transmit
				EMReadScreen worker, 7, 21, 21
	    		'Reading the disa dates
	    		EMReadScreen disa_start_date, 10, 6, 47			'reading DISA dates
	    		EMReadScreen disa_end_date, 10, 6, 69
	    		disa_start_date = Replace(disa_start_date," ","/")		'cleans up DISA dates
	    		disa_end_date = Replace(disa_end_date," ","/")
	    		disa_dates = trim(disa_start_date) & " - " & trim(disa_end_date)
	    		If disa_dates = "__/__/____ - __/__/____" then disa_dates = "NO DISA INFO"
	    		
	    		EMReadScreen disa_status, 2, 11, 59
	    		EMReadScreen verif_code, 1, 11, 69
	    		If trim(disa_status) <> "" then disa_status = replace(disa_status, "_", "")
	    		If trim(verif_code) <> "" then verif_code = replace(verif_code, "_", "")
	    		disa_info = disa_status & "-" & verif_code 
				'msgbox disa_info
	    		
	    		Redim Preserve GA_array(8, entry_number)
	    		GA_array(basket_numb, 	entry_number) = worker
	    		GA_array(case_numb, 	entry_number) = MAXIS_case_number
	    		GA_array(cl_name, 		entry_number) = client_name
	    		GA_array(memb_numb, 	entry_number) = memb_number
	    		GA_array(cl_age, 		entry_number) = client_age
				GA_array(wreg_panel, 	entry_number) = wreg_code
	    		GA_array(disa_date, 	entry_number) = disa_dates
	    		GA_array(dis_verifs, 	entry_number) = disa_info
	    		entry_number = entry_number + 1
				STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter 
	    	End if 
		Next 
	End if
Next 
STATS_counter = STATS_counter - 1

'Post MAXIS input coding---------------------------------------------------------------------------------------------------- 
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Setting the Excel rows with variables
ObjExcel.Cells(1, 1).Value = "Worker"
ObjExcel.Cells(1, 2).Value = "Case Number"
ObjExcel.Cells(1, 3).Value = "Client Name"
ObjExcel.Cells(1, 4).Value = "MEMB #"
ObjExcel.Cells(1, 5).Value = "Age"
ObjExcel.Cells(1, 6).Value = "WREG codes"
ObjExcel.Cells(1, 7).Value = "Disa dates"
ObjExcel.Cells(1, 8).Value = "Disa status"

FOR i = 1 to 8		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'End of setting up the Excel sheet----------------------------------------------------------------------------------------------------
excel_row = 2
For i = 0 to Ubound(GA_array, 2)
	ObjExcel.Cells(excel_row, 1).Value = ga_array (basket_numb, i)
	ObjExcel.Cells(excel_row, 2).Value = ga_array (case_numb, 	i)
	ObjExcel.Cells(excel_row, 3).Value = ga_array (cl_name,		i)
	ObjExcel.Cells(excel_row, 4).Value = ga_array (memb_numb, 	i)
	ObjExcel.Cells(excel_row, 5).Value = ga_array (cl_age, 		i)
	ObjExcel.Cells(excel_row, 6).Value = ga_array (wreg_panel, 	i)
	ObjExcel.Cells(excel_row, 7).Value = ga_array (disa_date, 	i)
	ObjExcel.Cells(excel_row, 8).Value = ga_array (dis_verifs, 	i)
	excel_row = excel_row + 1
next

'Query date/time/runtime info
ObjExcel.Cells(1, 9).Value = "Query date and time:"	'Goes back one, as this is on the next row
objExcel.Cells(1, 9).Font.Bold = TRUE
ObjExcel.Cells(2, 9).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
objExcel.Cells(2, 9).Font.Bold = TRUE
ObjExcel.Cells(3, 9).Value = "Case count:"	'Goes back one, as this is on the next row
objExcel.Cells(3, 9).Font.Bold = TRUE
ObjExcel.Cells(1, 10).Value = now
ObjExcel.Cells(2, 10).Value = timer - query_start_time
ObjExcel.Cells(3, 10).Value = STATS_counter

FOR i = 1 to 10		'formatting the cells'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

script_end_procedure("Success! Please review the list generated.")		