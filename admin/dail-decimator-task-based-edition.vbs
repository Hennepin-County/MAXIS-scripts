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
call changelog_update("10/23/2019", "Added Adults & ADS population option.", "Ilse Ferris, Hennepin County")
call changelog_update("10/17/2019", "Added ADS as population option.", "Ilse Ferris, Hennepin County")
call changelog_update("10/09/2019", "Update DAIL server name to switch from Dev to Production.", "Ilse Ferris, Hennepin County")
call changelog_update("09/26/2019", "Update DAIL options from types of DAIL messages to population-based options.", "Ilse Ferris, Hennepin County")
call changelog_update("09/12/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

'END CHANGELOG BLOCK =======================================================================================================

BeginDialog dail_dialog, 0, 0, 266, 85, "Dail Decimator dialog"
  DropListBox 75, 50, 80, 15, "Select one..."+chr(9)+"ADS"+chr(9)+"Adults"+chr(9)+"Adults & ADS", dail_to_decimate
  CheckBox 15, 70, 145, 10, "OR Check here to process for all workers.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 165, 65, 45, 15
    CancelButton 215, 65, 45, 15
  GroupBox 10, 5, 250, 40, "Using the DAIL Decimator script"
  Text 20, 20, 235, 20, "This script should be used to remove All DAIL messages that have been determined by Quality Improvement staff do not require action."
  Text 15, 55, 60, 10, "Select population:"
EndDialog

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""

this_month = CM_mo & " " & CM_yr
next_month = CM_plus_1_mo & " " & CM_plus_1_yr
CM_minus_2_mo =  right("0" & DatePart("m", DateAdd("m", -2, date)), 2)

'Finding the right folder to automatically save the file
month_folder = "DAIL " & CM_mo & "-" & DatePart("yyyy", date) & ""
decimator_folder = replace(this_month, " ", "-") & " DAIL Decimator"
report_date = replace(date, "/", "-")

Do
	Do
  		err_msg = ""
  		dialog dail_dialog
  		cancel_without_confirmation
  		If trim(dail_to_decimate) = "Select one..." and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases."
  		If trim(dail_to_decimate) <> "Select one..." and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases, not both options."
  	  	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
  	LOOP until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

back_to_SELF 'navigates back to self in case the worker is working within the DAIL. All messages for a single number may not be captured otherwise.

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
    If dail_to_decimate = "Adults" then
        worker_number = "X127EE1,X127EE2,X127EE3,X127EE4,X127EE5,X127EE6,X127EE7,X127EL2,X127EL3,X127EL4,X127EL5,X127EL6,X127EL7,X127EL8,X127EN1,X127EN2,X127EN3,X127EN4,X127EN5,X127EQ1,X127EQ2,X127EQ4,X127EQ5,X127EQ8,X127EQ9,X127EG5,X127EJ1,X127EH9,X127EM2,X127FE6,X127F3D,X127EL8,X127EL9,X127ED8,X127EH8,X127EG4,X127F3P"
    Elseif dail_to_decimate = "ADS" then 
        worker_number = "X127EH1,X127EH2,X127EH3,X127EH6,X127EJ4,X127EJ6,X127EJ7,X127EJ8,X127EK1,X127EK2,X127EK4,X127EK5,X127EK6,X127EK9,X127EM1,X127EM7,X127EM8,X127EM9,X127EN6,X127EP3,X127EP4,X127EP5,X127EP9,X127F3F,X127FE5,X127FG3,X127FH4,X127FH5,X127FI2,X127FI7,X127EJ5"
    Elseif dail_to_decimate = "Adults & ADS" then 
        worker_number = "X127EE1,X127EE2,X127EE3,X127EE4,X127EE5,X127EE6,X127EE7,X127EL2,X127EL3,X127EL4,X127EL5,X127EL6,X127EL7,X127EL8,X127EN1,X127EN2,X127EN3,X127EN4,X127EN5,X127EQ1,X127EQ2,X127EQ4,X127EQ5,X127EQ8,X127EQ9,X127EG5,X127EJ1,X127EH9,X127EM2,X127FE6,X127F3D,X127EL8,X127EL9,X127ED8,X127EH8,X127EG4,X127F3P" & _
        ",X127EH1,X127EH2,X127EH3,X127EH6,X127EJ4,X127EJ6,X127EJ7,X127EJ8,X127EK1,X127EK2,X127EK4,X127EK5,X127EK6,X127EK9,X127EM1,X127EM7,X127EM8,X127EM9,X127EN6,X127EP3,X127EP4,X127EP5,X127EP9,X127F3F,X127FE5,X127FG3,X127FH4,X127FH5,X127FI2,X127FI7,X127EJ5"
    End if
    
    x1s_from_dialog = split(worker_number, ",")	'Splits the worker array based on commas

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

            '----------------------------------------------------------------------------------------------------CSES Messages
            If instr(dail_msg, "AMT CHILD SUPP MOD/ORD") OR _
                instr(dail_msg, "AP OF CHILD REF NBR:") OR _
                instr(dail_msg, "ADDRESS DIFFERS W/ CS RECORDS:") OR _
                instr(dail_msg, "CHILD SUPP PAYMT FREQUENCY IS MONTHLY FOR CHILD REF NBR") OR _
                instr(dail_msg, "CHILD SUPP PAYMT FREQUENCY IS MONTHLY FOR CHILD REF NBR") OR _
                instr(dail_msg, "CHILD SUPP PAYMTS PD THRU THE COURT/AGENCY FOR CHILD") OR _
                instr(dail_msg, "CS REPORTED: NEW EMPLOYER FOR CAREGIVER REF NBR") OR _
                instr(dail_msg, "IS LIVING W/CAREGIVER") OR _
                instr(dail_msg, "NAME DIFFERS W/ CS RECORDS:") OR _
                instr(dail_msg, "PATERNITY ON CHILD REF NBR") OR _
                instr(dail_msg, "REPORTED NAME CHG TO:") OR _
                instr(dail_msg, "BENEFITS RETURNED, IF IOC HAS NEW ADDRESS") OR _
    		    instr(dail_msg, "CASE IS CATEGORICALLY ELIGIBLE") OR _
                instr(dail_msg, "CASE NOT AUTO-APPROVED - HRF/SR/RECERT DUE") OR _
                instr(dail_msg, "CHANGE IN BUDGET CYCLE") OR _
                instr(dail_msg, "COMPLETE ELIG IN FIAT") OR _
    		    instr(dail_msg, "COUNTED IN LBUD AS UNEARNED INCOME") OR _
                instr(dail_msg, "COUNTED IN SBUD AS UNEARNED INCOME") OR _
                instr(dail_msg, "HAS EARNED INCOME IN 6 MONTH BUDGET BUT NO DCEX PANEL") OR _
                instr(dail_msg, "NEW DENIAL ELIG RESULTS EXIST") OR _
    		    instr(dail_msg, "NEW ELIG RESULTS EXIST") OR _
    		    instr(dail_msg, "POTENTIALLY CATEGORICALLY ELIGIBLE") OR _
                instr(dail_msg, "WARNING MESSAGES EXIST") OR _
                instr(dail_msg, "ADDR CHG*CHK SHEL") OR _
                instr(dail_msg, "APPLCT ID CHNGD") OR _
                instr(dail_msg, "CASE AUTOMATICALLY DENIED") OR _
			    instr(dail_msg, "CASE FILE INFORMATION WAS SENT ON") OR _
			    instr(dail_msg, "CASE NOTE ENTERED BY") OR _
			    instr(dail_msg, "CASE NOTE TRANSFER FROM") OR _
			    instr(dail_msg, "CASE VOLUNTARY WITHDRAWN") OR _
			    instr(dail_msg, "CASE XFER") OR _
                instr(dail_msg, "CHANGE REPORT FORM SENT ON") OR _
			    instr(dail_msg, "DIRECT DEPOSIT STATUS") OR _
                instr(dail_msg, "EFUNDS HAS NOTIFIED DHS THAT THIS CLIENT'S EBT CARD") OR _
			    instr(dail_msg, "MEMB:NEEDS INTERPRETER HAS BEEN CHANGED") OR _
			    instr(dail_msg, "MEMB:SPOKEN LANGUAGE HAS BEEN CHANGED") OR _
			    instr(dail_msg, "MEMB:RACE CODE HAS BEEN CHANGED FROM UNABLE") OR _
			    instr(dail_msg, "MEMB:SSN HAS BEEN CHANGED FROM") OR _
			    instr(dail_msg, "MEMB:SSN VER HAS BEEN CHANGED FROM") OR _
			    instr(dail_msg, "MEMB:WRITTEN LANGUAGE HAS BEEN CHANGED FROM") OR _
                instr(dail_msg, "MEMI: HAS BEEN DELETED BY THE PMI MERGE PROCESS") OR _
                instr(dail_msg, "NOT ACCESSED FOR 300 DAYS,SPEC NOT") OR _
			    instr(dail_msg, "PMI MERGED") OR _
			    instr(dail_msg, "THIS APPLICATION WILL BE AUTOMATICALLY DENIED") OR _
			    instr(dail_msg, "THIS CASE IS ERROR PRONE") OR _
                instr(dail_msg, "EMPL SERV REF DATE IS > 60 DAYS; CHECK ES PROVIDER RESPONSE") OR _
                instr(dail_msg, "MEMBER HAS TURNED 60 - FSET:WORK REG HAS BEEN UPDATED") OR _
                instr(dail_msg, "LAST GRADE COMPLETED") OR _
                instr(dail_msg, "~*~*~CLIENT WAS SENT AN APPT LETTER") OR _
                instr(dail_msg, "IF CLIENT HAS NOT COMPLETED RECERT, APPL CAF FOR") OR _
                instr(dail_msg, "UPDATE PND2 FOR CLIENT DELAY IF APPROPRIATE") then
    		        add_to_excel = True
                '----------------------------------------------------------------------------------------------------CORRECT STAT EDITS over 5 days old
            Elseif instr(dail_msg, "CORRECT STAT EDITS") then
                EmReadscreen stat_date, 8, dail_row, 39
                ten_days_ago = DateAdd("d", -5, date)
                If cdate(ten_days_ago) => cdate(stat_date) then
                    add_to_excel = True
                Else
                    add_to_excel = False
                End if
            '----------------------------------------------------------------------------------------------------REMOVING PEPR messages not CM or CM + 1
            Elseif dail_type = "PEPR" then
                if dail_month = this_month or dail_month = next_month then
                    add_to_excel = False
                Else
                    add_to_excel = True ' delete the old messages
                End if
            '----------------------------------------------------------------------------------------------------clearing elig messages older than CM
            Elseif instr(dail_msg, "OVERPAYMENT POSSIBLE") or InStr(dail_msg, "DISBURSE EXPEDITED SERVICE") then
                if dail_month = this_month or dail_month = next_month then
                    add_to_excel = False
                Else
                    add_to_excel = True ' delete the old messages
                End if
            '----------------------------------------------------------------------------------------------------clearing Exempt IR TIKL's over 2 months old.
            Elseif instr(dail_msg, "%^% SENT THROUGH") then
                TIKL_date = cdate(TIKL_date)
                TIKL_date = right("0" & DatePart("m",dail_month), 2)
                if TIKL_date = CM_minus_2_mo then
                    add_to_excel = True   ' delete the exempt IR message older than last month.
                Else
                    add_to_excel = False
                End if
                '----------------------------------------------------------------------------------------------------MEC2
            Elseif dail_type = "MEC2" then
                if  instr(dail_msg, "RSDI END DATE") OR _
                    instr(dail_msg, "SELF EMPLOYMENT REPORTED TO MEC²") OR _
                    instr(dail_msg, "SSI REPORTED TO MEC²") OR _
                    instr(dail_msg, "UNEMPLOYMENT INS") then
                    add_to_excel = FALSE            'Income based MEC2 messages will not be removed
                Else
                    add_to_excel =  True    'All other MEC2 messages can be deleted.
                End if
                '----------------------------------------------------------------------------------------------------TIKL
            Elseif dail_type = "TIKL" then
                if instr(dail_msg, "VENDOR") OR instr(dail_msg, "VND") then
                    add_to_excel = FALSE        'Will not delete TIKL's with vendor information
                Else
                    six_months = DateAdd("M", -6, date)
                    If cdate(six_months) => cdate(dail_month) then
                        add_to_excel = True     'Will delete any TIKL over 6 months old
                    Else
                        add_to_excel = False
                    End if
                End if
            Else
                add_to_excel = False
            End if

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
                 
                ReDim Preserve DAIL_array(4, DAIL_count)	'This resizes the array based on the number of rows in the Excel File'
            	DAIL_array(worker_const,	           DAIL_count) = worker
            	DAIL_array(maxis_case_number_const,    DAIL_count) = MAXIS_case_number
            	DAIL_array(dail_type_const, 	       DAIL_count) = dail_type
            	DAIL_array(dail_month_const, 		   DAIL_count) = dail_month
            	DAIL_array(dail_msg_const, 		       DAIL_count) = dail_msg
                Dail_count = DAIL_count + 1
			End if

			EMReadScreen message_error, 11, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.
			If message_error = "NO MESSAGES" then
				CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
				Call write_value_and_transmit(worker, 21, 6)
				transmit   'transmit past 'not your dail message'
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

script_end_procedure("Success! Please review the list created for accuracy.")