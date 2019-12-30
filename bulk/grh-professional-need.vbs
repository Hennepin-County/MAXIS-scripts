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
call changelog_update("01/25/2018", "Script updated to gather DISA and Cert dates from STAT/DISA. Removed several fields from STAT/FACI. Also organized how data is presented on Excel spreadsheet.", "Ilse Ferris, Hennepin County")
call changelog_update("01/12/2018", "Script updated to also gather FACI in dates, next revw date. Also organized how data is presented on Excel spreadsheet.", "Ilse Ferris, Hennepin County")
call changelog_update("01/08/2018", "Script updated to also gather waiver types from STAT/DISA.", "Ilse Ferris, Hennepin County")
call changelog_update("03/31/2017", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CONNECTS TO BlueZone
EMConnect ""
all_workers_check = 1        'autochecking as this is the default setting

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 80, "GRH Professional Need Dialog"
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
		Dialog Dialog1
		cancel_without_confirmation
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
ObjExcel.Cells(1, 2).Value = "Case #"
ObjExcel.Cells(1, 3).Value = "Next REVW"
ObjExcel.Cells(1, 4).Value = "Facility Name"
ObjExcel.Cells(1, 5).Value = "GRH Rate"
ObjExcel.Cells(1, 6).Value = "DISA Dates"
ObjExcel.Cells(1, 7).Value = "Certification Dates"
ObjExcel.Cells(1, 8).Value = "GRH Plan Dates"
ObjExcel.Cells(1, 9).Value = "Waiver Type"

FOR i = 1 to 9		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

excel_row = 2 

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
				EMReadScreen next_revw_date, 8, row, 42
				ObjExcel.Cells(excel_row, 1).Value = worker
				ObjExcel.Cells(excel_row, 2).Value = trim(MAXIS_case_number)
				ObjExcel.Cells(excel_row, 3).Value = replace(next_revw_date, " ", "/")
				excel_row = excel_row + 1
			End if 
			row = row + 1
			STATS_counter = STATS_counter + 1                      'adds one instance to the stats counter
		Loop until row = 19
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
		ObjExcel.Cells(excel_row, 4).Value = "PRIV cases"
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
    Else 
	    EMReadScreen member_number, 2, 4, 33
	    If member_number <> "01" then 
	    	EmWriteScreen "01", 20, 76						'For member 01 - All GRH cases should be for member 01. 
	    	Call write_value_and_transmit ("01", 20, 79)	'1st version of FACI 
	    End if 
	 
	    EMReadScreen FACI_total_check, 1, 2, 78
	    If FACI_total_check = "0" then 
	    	current_faci = False 
			ObjExcel.Cells(excel_row, 4).Value = "Case does not have a FACI panel."	
	    	case_status = ""
	    Else 
	    	row = 14
	    	Do 
	    		EMReadScreen date_out, 10, row, 71
	    		'msgbox "date out: " & date_out 
	    		If date_out = "__ __ ____" then 
	 				EMReadScreen date_in, 10, row, 47
					If date_in <> "__ __ ____" then 
						current_faci = TRUE
	    				exit do
	    			ELSE
	    				current_faci = False 
	    				row = row + 1
	    			End if 
	    		Else 
	    			row = row + 1
	    			'msgbox row
	    			current_faci = False	
	    		End if 	
	    		If row = 19 then 
	    			transmit
	    			row = 14
	    		End if 
	    		EMReadScreen last_panel, 5, 24, 2
	    	Loop until last_panel = "ENTER"	'This means that there are no other faci panels
	    End if 
		
	    'GETS FACI NAME AND PUTS IT IN SPREADSHEET, IF CLIENT IS IN FACI.
	    If current_faci = True then
	    	EMReadScreen FACI_name, 30, 6, 43
			EMReadScreen GRH_rate, 1, row, 34	
	    	ObjExcel.Cells(excel_row, 4).Value = trim(replace(FACI_name, "_", ""))
			ObjExcel.Cells(excel_row, 5).Value = trim(replace(GRH_rate, "_", ""))
	    End if 
		
	    Call navigate_to_MAXIS_screen("STAT", "DISA")
		'Reading the disa dates
		EMReadScreen disa_start_date, 10, 6, 47			'reading DISA dates
		EMReadScreen disa_end_date, 10, 6, 69
		disa_start_date = Replace(disa_start_date," ","/")		'cleans up DISA dates
		disa_end_date = Replace(disa_end_date," ","/")
		disa_dates = trim(disa_start_date) & " - " & trim(disa_end_date)
		If disa_dates = "__/__/____ - __/__/____" then disa_dates = ""
		ObjExcel.Cells(excel_row, 6).Value = disa_dates
		
		EMReadScreen cert_start_date, 10, 9, 47			'reading cert dates
		EMReadScreen cert_end_date, 10, 9, 69
		cert_start_date = Replace(cert_start_date," ","/")		'cleans up cert dates
		cert_end_date = Replace(cert_end_date," ","/")
		cert_dates = trim(cert_start_date) & " - " & trim(cert_end_date)
		If cert_dates = "__/__/____ - __/__/____" then cert_dates = ""
		ObjExcel.Cells(excel_row, 7).Value = cert_dates
		
		EMReadScreen GRH_start_date, 10, 9, 47			'reading GRH dates
		EMReadScreen GRH_end_date, 10, 9, 69
		GRH_start_date = Replace(GRH_start_date," ","/")		'cleans up GRH dates
		GRH_end_date = Replace(GRH_end_date," ","/")
		GRH_dates = trim(GRH_start_date) & " - " & trim(GRH_end_date)
		If GRH_dates = "__/__/____ - __/__/____" then GRH_dates = ""
		ObjExcel.Cells(excel_row, 8).Value = GRH_dates
	    
	    'checks the waiver type
	    EMReadScreen DISA_waiver_type, 1, 14, 59
	    If DISA_waiver_type = "_" then DISA_waiver_type = ""
	    ObjExcel.Cells(excel_row, 9).Value = DISA_waiver_type
	End if 
	
	excel_row = excel_row + 1 'setting up the script to check the next row.
LOOP UNTIL objExcel.Cells(excel_row, 2).Value = ""	'Loops until there are no more cases in the Excel list

'Query date/time/runtime info
objExcel.Cells(1, 10).Font.Bold = TRUE
objExcel.Cells(2, 10).Font.Bold = TRUE
ObjExcel.Cells(1, 10).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, 10).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, 11).Value = now
ObjExcel.Cells(2, 11).Value = timer - query_start_time

'formatting the cells
FOR i = 1 to 11
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created.")