'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - DAIL DECIMATOR - TASK BASED EDITION.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = 20
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
call changelog_update("08/10/2020", "Added email confirmation at the end of the script run.", "Ilse Ferris, Hennepin County")
call changelog_update("04/29/2020", "Added new ADULTS pending baskets X127EP6-X127EP8", "Ilse Ferris, Hennepin County")
call changelog_update("04/01/2020", "Added DWP baskets to FAD baskets.", "Ilse Ferris, Hennepin County")
call changelog_update("03/30/2020", "Update to exclude certain ADAD baskets.", "Ilse Ferris, Hennepin County")
call changelog_update("03/27/2020", "Added newly added ADAD and FAD baskets.", "Ilse Ferris, Hennepin County")
call changelog_update("03/19/2020", "Added functionality to remove duplicate messages from the remaining DAIL count.", "Ilse Ferris, Hennepin County")
call changelog_update("03/16/2020", "Added FAD option for task-based assginment.", "Ilse Ferris, Hennepin County")
call changelog_update("12/17/2019", "Added function to evaluate DAIL messages.", "Ilse Ferris, Hennepin County")
call changelog_update("12/12/2019", "Added DAIL type selection to dialog box.", "Ilse Ferris, Hennepin County")
call changelog_update("12/02/2019", "Updated Adults basket numbers per Faughn's request.", "Ilse Ferris, Hennepin County")
call changelog_update("10/23/2019", "Added Adults & ADS population option.", "Ilse Ferris, Hennepin County")
call changelog_update("10/17/2019", "Added ADS as population option.", "Ilse Ferris, Hennepin County")
call changelog_update("10/09/2019", "Update DAIL server name to switch from Dev to Production.", "Ilse Ferris, Hennepin County")
call changelog_update("09/26/2019", "Update DAIL options from types of DAIL messages to population-based options.", "Ilse Ferris, Hennepin County")
call changelog_update("09/12/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================

Function dail_type_selection
	'selecting the type of DAIl message
    If all_check = 0 then
	    EMWriteScreen "x", 4, 12		'transmits to the PICK screen
	    transmit
	    EMWriteScreen "_", 7, 39		'clears the all selection

	    If cola_check = 1 then EMWriteScreen "x", 8, 39
	    If clms_check = 1 then EMWriteScreen "x", 9, 39
	    If cses_check = 1 then EMWriteScreen "x", 10, 39
	    If elig_check = 1 then EMWriteScreen "x", 11, 39
	    If ievs_check = 1 then EMWriteScreen "x", 12, 39
	    If info_check = 1 then EMWriteScreen "x", 13, 39
	    If ive_check = 1 then EMWriteScreen "x", 14, 39
	    If ma_check = 1 then EMWriteScreen "x", 15, 39
 	    If mec2_check = 1 then EMWriteScreen "x", 16, 39
	    If pari_chck = 1 then EMWriteScreen "x", 17, 39
	    If pepr_check = 1 then EMWriteScreen "x", 18, 39
	    If tikl_check = 1 then EMWriteScreen "x", 19, 39
	    If wf1_check = 1 then EMWriteScreen "x", 20, 39
	    transmit
    End if
End Function

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""

this_month = CM_mo & " " & CM_yr
next_month = CM_plus_1_mo & " " & CM_plus_1_yr
CM_minus_2_mo =  right("0" & DatePart("m", DateAdd("m", -2, date)), 2)

'Finding the right folder to automatically save the file
month_folder = "DAIL " & CM_mo & "-" & DatePart("yyyy", date) & ""
decimator_folder = replace(this_month, " ", "-") & " DAIL Decimator"
report_date = replace(date, "/", "-")

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 251, 210, "DAIL Decimation Main Dialog"
  CheckBox 55, 105, 85, 10, "All Population Baskets", all_baskets_checkbox
  CheckBox 140, 105, 30, 10, "ADAD", ADAD_checkbox
  CheckBox 175, 105, 30, 10, "ADS", ADS_checkbox
  CheckBox 205, 105, 30, 10, "FAD", FAD_checkbox
  CheckBox 10, 120, 135, 10, "OR check here to process for all workers.", all_workers_check
  CheckBox 10, 155, 25, 10, "ALL",   all_check
  CheckBox 40, 155, 30, 10, "COLA",  cola_check
  CheckBox 75, 155, 30, 10, "CLMS",  clms_check
  CheckBox 110, 155, 30, 10, "CSES", cses_check
  CheckBox 145, 155, 30, 10, "ELIG", elig_check
  CheckBox 180, 155, 30, 10, "IEVS", ievs_check
  CheckBox 210, 155, 30, 10, "INFO", info_check
  CheckBox 10, 170, 25, 10, "IV-E",  ive_check
  CheckBox 40, 170, 25, 10, "MA",    ma_check
  CheckBox 75, 170, 30, 10, "MEC2",  mec2_check
  CheckBox 110, 170, 35, 10, "PARI", pari_chck
  CheckBox 145, 170, 30, 10, "PEPR", pepr_check
  CheckBox 180, 170, 30, 10, "TIKL", tikl_check
  CheckBox 210, 170, 30, 10, "WF1",  wf1_check
  ButtonGroup ButtonPressed
    OkButton 155, 190, 40, 15
    CancelButton 200, 190, 40, 15
  Text 10, 105, 40, 10, "Population:"
  GroupBox 5, 85, 240, 50, "Step 1. Select the population"
  Text 65, 5, 135, 10, "---DAIL Decimator:Task-Based Edition---"
  Text 10, 35, 220, 40, "This script will delete DAIL messages from the DAIL type and population selected below, and capture the actionable DAIL messages. The final step once the DAIL messages have been evaluated will be to out put all actionable DAIL messages to a SQL Database which feeds the Big Scoop Report."
  GroupBox 5, 20, 240, 60, "Using the DAIL Decimator script"
  GroupBox 5, 140, 240, 45, "Step 2. Select the type(s) of DAIL message to add to the report:"
EndDialog

Do
    Do 
        err_msg = ""
  	    dialog Dialog1
  	    cancel_without_confirmation 
        If all_baskets_checkbox = 1 then 
            If ADAD_checkbox = 1 or ADS_checkbox = 1 or FAD_checkbox = 1 then err_msg = err_msg & vbcr & "* You cannot select a population and all populations."
        End if 
        If all_workers_check = 1 then 
            If ADAD_checkbox = 1 or ADS_checkbox = 1 or FAD_checkbox = 1 or all_baskets_checkbox = 1 then err_msg = err_msg & vbcr & "* You cannot select a population(s) and all workers."
        End if 
        If (all_baskets_checkbox = 0 and all_workers_check = 0 and ADAD_checkbox = 0 and ADS_checkbox = 0 and FAD_checkbox = 0 and all_baskets_checkbox = 0) then err_msg = err_msg & vbcr & "* You must select at least one population option."
        If (all_check = 0 and cola_check = 0 and clms_check = 0 and cses_check  = 0 and elig_check  = 0 and ievs_check = 0 and info_check = 0 and ive_check = 0 and ma_check = 0 and mec2_check = 0 and pari_chck = 0 and pepr_check = 0 and tikl_check = 0 and wf1_check = 0) then err_msg = err_msg & vbcr & "* You must select at least one DAIL type."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
    Loop Until err_msg = ""     
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

back_to_SELF 'navigates back to self in case the worker is working within the DAIL. All messages for a single number may not be captured otherwise.

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
    dail_to_decimate = "ALL"
Else
    ADAD_baskets = "X127EE1,X127EE2,X127EE3,X127EE4,X127EE5,X127EE6,X127EE7,X127EL2,X127EL3,X127EL4,X127EL5,X127EL6,X127EL7,X127EL8,X127EN1,X127EN2,X127EN3,X127EN4,X127EN5,X127EQ1,X127EQ4,X127EQ5,X127EQ8,X127EQ9,X127EL9,X127ED8,X127EH8,X127EG4,X127EQ3,X127EQ2,X127EP6,X127EP7,X127EP8,"
    ADS_baskets = "X127EH1,X127EH2,X127EH3,X127EH6,X127EJ4,X127EJ6,X127EJ7,X127EJ8,X127EK1,X127EK2,X127EK4,X127EK5,X127EK6,X127EK9,X127EM1,X127EM7,X127EM8,X127EM9,X127EN6,X127EP3,X127EP4,X127EP5,X127EP9,X127F3F,X127FE5,X127FG3,X127FH4,X127FH5,X127FI2,X127FI7,X127EJ5,"
    FAD_baskets = "X127ES1,X127ES2,X127ES3,X127ES4,X127ES5,X127ES6,X127ES7,X127ES8,X127ES9,X127ET1,X127ET2,X127ET3,X127ET4,X127ET5,X127ET6,X127ET7,X127ET8,X127ET9,X127FE7,X127FE8,X127FE9"
    
    worker_numbers = ""     'Creating and valuing incrementor variables
    dail_to_decimate = ""
    
    If ADAD_checkbox = 1 then 
        worker_numbers = worker_numbers & ADAD_baskets
        dail_to_decimate = dail_to_decimate & "ADAD,"
    End if 
    
    If ADS_checkbox = 1 then 
        worker_numbers = worker_numbers & ADS_baskets
        dail_to_decimate = dail_to_decimate & "ADS,"
    End if 
    
    If FAD_checkbox = 1 then 
        worker_numbers = worker_numbers & FAD_baskets
        dail_to_decimate = dail_to_decimate & "FAD"
    End if 
    
    If all_baskets_checkbox = 1 then 
        worker_numbers = ADAD_baskets & "," & ADS_baskets & "," & FAD_baskets  'conditional logic in do loop doesn't allow for populations and baskets to be selcted. Not incremented variable.
        dail_to_decimate = "All Baskets"
    End if 
    
    dail_to_decimate = trim(dail_to_decimate)  'trims excess spaces of dail_to_decimate
    If right(dail_to_decimate, 1) = "," THEN dail_to_decimate = left(dail_to_decimate, len(dail_to_decimate) - 1)
    dail_to_decimate = dail_to_decimate & " TB"
        
    x1s_from_dialog = split(worker_numbers, ",")	'Splits the worker array based on commas

	'Need to add the worker_county_code to each one
	For each x1_number in x1s_from_dialog
		If worker_array = "" then
			worker_array = trim(ucase(x1_number))		'replaces worker_county_code if found in the typed x1 number
		Else
			worker_array = worker_array & "," & trim(ucase(x1_number)) 'replaces worker_county_code if found in the typed x1 number
		End if
	Next
	'Split worker_array
	worker_array = split(worker_array, ",")
End if

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to "DAIL List"
ObjExcel.ActiveSheet.Name = "Deleted DAILS - " & dail_to_decimate

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X NUMBER"
objExcel.Cells(1, 2).Value = "CASE #"
objExcel.Cells(1, 3).Value = "DAIL TYPE"
objExcel.Cells(1, 4).Value = "DAIL MO."
objExcel.Cells(1, 5).Value = "DAIL MESSAGE"

FOR i = 1 to 5		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

DIM DAIL_array()
ReDim DAIL_array(4, 0)
Dail_count = 0              'Incremental for the array
all_dail_array = "*"    'setting up string to find duplicate DAIL messages. At times there is a glitch in the DAIL, and messages are reviewed a second time.
false_count = 0

'constants for array
const worker_const	            = 0
const maxis_case_number_const   = 1
const dail_type_const 	        = 2
const dail_month_const		    = 3
const dail_msg_const		    = 4

'Sets variable for all of the Excel stuff
excel_row = 2
deleted_dails = 0	'establishing the value of the count for deleted deleted_dails

CALL navigate_to_MAXIS_screen("DAIL", "DAIL")

'This for...next contains each worker indicated above
For each worker in worker_array
    MAXIS_case_number = ""
	DO
		EMReadScreen dail_check, 4, 2, 48
		If next_dail_check <> "DAIL" then
			MAXIS_case_number = ""
			CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
		End if
	Loop until dail_check = "DAIL"

	EMWriteScreen worker, 21, 6
	transmit
	transmit 'transmit past 'not your dail message'

    Call dail_type_selection
    EMReadScreen number_of_dails, 1, 3, 67		'Reads where the count of DAILs is listed

	DO
		If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped

		dail_row = 6			'Because the script brings each new case to the top of the page, dail_row starts at 6.
		DO
			dail_type = ""
			dail_msg = ""

		    'Determining if there is a new case number...
		    EMReadScreen new_case, 8, dail_row, 63
		    new_case = trim(new_case)
		    IF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
				Call write_value_and_transmit("T", dail_row, 3)
				dail_row = 6
			ELSEIF new_case = "CASE NBR" THEN
			    '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
			    Call write_value_and_transmit("T", dail_row + 1, 3)
				dail_row = 6
			End if

            'Reading the DAIL Information
			EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
            MAXIS_case_number = trim(MAXIS_case_number)
            MAXIS_case_number = right("00000000" & MAXIS_case_number, 8) 'outputs in 8 digits format

            EMReadScreen dail_type, 4, dail_row, 6

            EMReadScreen dail_msg, 61, dail_row, 20
			dail_msg = trim(dail_msg)

            EMReadScreen dail_month, 8, dail_row, 11
            dail_month = trim(dail_month)

            stats_counter = stats_counter + 1   'I increment thee
            Call non_actionable_dails   'Function to evaluate the DAIL messages

            IF add_to_excel = True then
				'--------------------------------------------------------------------...and put that in Excel.
				objExcel.Cells(excel_row, 1).Value = worker
				objExcel.Cells(excel_row, 2).Value = MAXIS_case_number
				objExcel.Cells(excel_row, 3).Value = dail_type
				objExcel.Cells(excel_row, 4).Value = dail_month
				objExcel.Cells(excel_row, 5).Value = dail_msg
				excel_row = excel_row + 1

				Call write_value_and_transmit("D", dail_row, 3)
				EMReadScreen other_worker_error, 13, 24, 2
				If other_worker_error = "** WARNING **" then transmit
				deleted_dails = deleted_dails + 1
			else
				add_to_excel = False
				dail_row = dail_row + 1
                If len(dail_month) = 5 then
                    output_year = ("20" & right(dail_month, 2))
                    output_month = left(dail_month, 2)
                    output_day = "01"
                    dail_month = output_year & "-" & output_month & "-" & output_day
                elseif trim(dail_month) <> "" then
                    'Adjusting data for output to SQL
                    output_year     = DatePart("yyyy",dail_month)   'YYYY-MM-DD format
                    output_month    = right("0" & DatePart("m", dail_month), 2)
                    output_day      = DatePart("d", dail_month)
                    dail_month = output_year & "-" & output_month & "-" & output_day
                End if
                
                dail_string = worker & " " & MAXIS_case_number & " " & dail_type & " " & dail_month & " " & dail_msg
                'If the case number is found in the string of case numbers, it's not added again. 
                If instr(all_dail_array, "*" & dail_string & "*") then
                    If dail_type = "HIRE" then
                        add_to_array = True 
                    Else 
                        add_to_array = False
                    End if 
                    'msgbox "Duplicate Found: " & dail_string & vbcr & add_to_array
                else 
                    add_to_array = True 
                End if 
                
                If add_to_array = True then          
                    ReDim Preserve DAIL_array(4, DAIL_count)	'This resizes the array based on the number of rows in the Excel File'
            	    DAIL_array(worker_const,	           DAIL_count) = worker
            	    DAIL_array(maxis_case_number_const,    DAIL_count) = MAXIS_case_number
            	    DAIL_array(dail_type_const, 	       DAIL_count) = dail_type
            	    DAIL_array(dail_month_const, 		   DAIL_count) = dail_month
            	    DAIL_array(dail_msg_const, 		       DAIL_count) = dail_msg
                    Dail_count = DAIL_count + 1
                    all_dail_array = trim(all_dail_array & dail_string & "*") 'Adding MAXIS case number to case number string
                    dail_string = ""
                else
                    false_count = false_count + 1
                End if 
			End if

			EMReadScreen message_error, 11, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.
			If message_error = "NO MESSAGES" then
				CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
				Call write_value_and_transmit(worker, 21, 6)
				transmit   'transmit past 'not your dail message'
                Call dail_type_selection
				exit do
			End if

			'...going to the next page if necessary
			EMReadScreen next_dail_check, 4, dail_row, 4
			If trim(next_dail_check) = "" then
				PF8
				EMReadScreen last_page_check, 21, 24, 2
				If last_page_check = "THIS IS THE LAST PAGE" then
					all_done = true
					exit do
				Else
					dail_row = 6
				End if
			End if
		LOOP
		IF all_done = true THEN exit do
	LOOP
Next

STATS_counter = STATS_counter - 1
'Enters info about runtime for the benefit of folks using the script
objExcel.Cells(2, 7).Value = "Number of DAILs deleted:"
objExcel.Cells(3, 7).Value = "Average time to find/select/copy/paste one line (in seconds):"
objExcel.Cells(4, 7).Value = "Estimated manual processing time (lines x average):"
objExcel.Cells(5, 7).Value = "Script run time (in seconds):"
objExcel.Cells(6, 7).Value = "Estimated time savings by using script (in minutes):"
objExcel.Cells(7, 7).Value = "Number of messages reviewed/DAIL messages remaining:"
objExcel.Columns(7).Font.Bold = true
objExcel.Cells(2, 8).Value = deleted_dails
objExcel.Cells(3, 8).Value = STATS_manualtime
objExcel.Cells(4, 8).Value = STATS_counter * STATS_manualtime
objExcel.Cells(5, 8).Value = timer - start_time
objExcel.Cells(6, 8).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60
objExcel.Cells(7, 8).Value = STATS_counter

'Formatting the column width.
FOR i = 1 to 8
	objExcel.Columns(i).AutoFit()
NEXT

'Adding another sheet
ObjExcel.Worksheets.Add().Name = "Remaining DAIL messages"

excel_row = 2
'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X NUMBER"
objExcel.Cells(1, 2).Value = "CASE #"
objExcel.Cells(1, 3).Value = "DAIL TYPE"
objExcel.Cells(1, 4).Value = "DAIL MO."
objExcel.Cells(1, 5).Value = "DAIL MESSAGE"

FOR i = 1 to 5		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Export informaiton to Excel re: case status
For item = 0 to UBound(DAIL_array, 2)
	objExcel.Cells(excel_row, 1).Value = DAIL_array(worker_const, item)
	objExcel.Cells(excel_row, 2).Value = DAIL_array(maxis_case_number_const, item)
    objExcel.Cells(excel_row, 3).Value = DAIL_array(dail_type_const, item)
	objExcel.Cells(excel_row, 4).Value = DAIL_array(dail_month_const, item)
    objExcel.Cells(excel_row, 5).Value = DAIL_array(dail_msg_const, item)
	excel_row = excel_row + 1
Next

objExcel.Cells(1, 7).Value = "Remaning DAIL messages:"
objExcel.Columns(7).Font.Bold = true
objExcel.Cells(1, 8).Value = DAIL_count

'formatting the cells
FOR i = 1 to 8
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

'saving the Excel file
file_info = month_folder & "\" & decimator_folder & "\" & report_date & " " & dail_to_decimate & " " & deleted_dails

'Saves and closes the most recent Excel workbook with the Task based cases to process.
objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\DAIL list\" & file_info & ".xlsx"

'Setting constants
Const adOpenStatic = 3
Const adLockOptimistic = 3

''Creating objects for Database
Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

'How to connect to the database
'Provider: the type of connection you are establishing, in this case SQL Server.
'Data Source: The server you are connecting to.
'Initial Catalog: The name of the database.
'user id: your username.
'password: um, your password. ;)

objConnection.Open "Provider = SQLOLEDB.1;Data Source= HSSQLPW017;Initial Catalog= BlueZone_Statistics; Integrated Security=SSPI;Auto Translate=False;"

'Deleting ALL data fom DAIL table prior to loading new DAIL messages.
objRecordSet.Open "DELETE FROM EWS.DAILDecimator",objConnection, adOpenStatic, adLockOptimistic

'Export informaiton to Excel re: case status
For item = 0 to UBound(DAIL_array, 2)
    worker             = DAIL_array(worker_const, item)
    MAXIS_case_number  = DAIL_array(maxis_case_number_const, item)
    dail_type          = DAIL_array(dail_type_const, item)
    dail_month         = DAIL_array(dail_month_const, item)
    dail_msg           = DAIL_array(dail_msg_const, item)

    If instr(dail_msg, "'") then dail_msg = replace(dail_msg, "'", " ") 'SQL will not allow for an apostrophe
    If instr(dail_msg, "*") then dail_msg = replace(dail_msg, "*", " ") 'SQL will not allow for an apostrophe
    dail_msg = trim(dail_msg)
    'Opening Database and adding a record
    objRecordSet.Open "INSERT INTO EWS.DAILDecimator(EmpStateLogOnID, MaxisCaseNumber, DAILType, DAILMessage, DAILMonth)" & _
    "VALUES ('" & worker & "', '" & MAXIS_case_number & "', '" & dail_type & "', '" & dail_msg & "', '" & dail_month & "')", objConnection, adOpenStatic, adLockOptimistic
Next

'Closing the connection
objConnection.Close

'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
Call create_outlook_email("Laurie.Hennen@hennepin.us;Todd.Bennington@hennepin.us", "Ilse.Ferris@hennepin.us", "DAIL Decimator: Task-Based Edition complete. EOM.", "", "", True)

script_end_procedure("Success! Please review the list created for accuracy. False count is: " & false_count)
