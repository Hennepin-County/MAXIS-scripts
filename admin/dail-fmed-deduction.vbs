'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - DAIL FMED DEDUCTION.vbs"
start_time = timer
STATS_counter = 1                    'sets the stats counter at one
STATS_manualtime = 180
STATS_denomination = "I"       			'I is for each item
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
call changelog_update("06/26/2023", "Resolved issue with exit functionality when a case load doesn't have a specific DAIL message type.", "Ilse Ferris, Hennepin County")
call changelog_update("08/04/2022", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""
all_workers_check = 1   'checked

'Finding the right folder to automatically save the file
month_folder = "DAIL " & CM_mo & "-" & DatePart("yyyy", date) & ""
decimator_folder = CM_mo & "-" & CM_yr & " DAIL Decimator"
report_date = replace(date, "/", "-")

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "DAIL FMED Deductions"
  GroupBox 10, 5, 250, 45, "Using the DAIL FMED Deductions Script"
  Text 20, 20, 235, 25, "This script should be used to send FMED Deductions notifications to residents. It will remove the TIKL messages and will send a SPEC/MEMO and CASE/NOTE actions taken."
  Text 15, 60, 60, 10, "Worker number(s):"
  EditBox 80, 55, 180, 15, worker_number
  CheckBox 80, 75, 135, 10, "Check here to process for all workers.", all_workers_check
  Text 5, 95, 60, 10, "Worker Signature:"
  EditBox 65, 90, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 180, 90, 40, 15
    CancelButton 220, 90, 40, 15
EndDialog

'the dialog
Do
	Do
  		err_msg = ""
  		dialog Dialog1
  		cancel_without_confirmation
  		If trim(worker_number) = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases."
  		If trim(worker_number) <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases, not both options."
        If trim(worker_signature) = "" then err_msg = err_msg & vbNewLine & "* Enter your worker signature."
  	  	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
  	LOOP until err_msg = ""
  	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call check_for_MAXIS(False)

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
ObjExcel.ActiveSheet.Name = "FMED Deduction PEPR's"

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X Number"
objExcel.Cells(1, 2).Value = "Case #"
objExcel.Cells(1, 3).Value = "DAIL Type"
objExcel.Cells(1, 4).Value = "DAIL Mo."
objExcel.Cells(1, 5).Value = "DAIL Message"
objExcel.Cells(1, 6).Value = "Case Status"

FOR i = 1 to 6		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

DIM DAIL_array()
ReDim DAIL_array(case_status_const, 0)
Dail_count = 0              'Incrementor for the array

'constants for array
const worker_const	                = 0
const maxis_case_number_const       = 1
const dail_type_const               = 2
const dail_month_const		        = 3
const dail_msg_const		        = 4
const excel_row_const               = 5
const case_status_const             = 6

'Sets variable for all of the Excel stuff
excel_row = 2
deleted_dails = 0	'establishing the value of the count for deleted deleted_dails
memo_count = 0      'counting number of memo's sent with this process.
all_case_numbers_array = "*"    'setting up string to find duplicate case numbers

MAXIS_case_number = ""
CALL navigate_to_MAXIS_screen("DAIL", "PICK")
EMWriteScreen "_", 7, 39    'blank out ALL selection
Call write_value_and_transmit("X", 18, 39)   'Select PEPR DAIL type

For each worker in worker_array
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
            IF new_case = "CASE NBR" THEN
			    '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
			    Call write_value_and_transmit("T", dail_row + 1, 3)
				dail_row = 6
			End if

            EMReadScreen maxis_case_number, 8, 5, 73
            EMReadScreen dail_month, 8, dail_row, 11
            EmReadscreen dail_type, 4, dail_row, 6
			EMReadScreen dail_msg, 61, dail_row, 20
            dail_msg = trim(dail_msg)
			stats_counter = stats_counter + 1

			If instr(dail_msg, "MEMBER HAS TURNED 60 - NOTIFY ABOUT POSSIBLE FMED DEDUCTION") then
                'If the case number is found in the string of case numbers, it's not added again.
                If instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") then
                    add_to_array = False
                Else
				    '--------------------------------------------------------------------...and add to the array/put that in Excel.
                    ReDim Preserve DAIL_array(case_status_const, DAIL_count)	'This resizes the array based on the number of rows in the Excel File'
            	    DAIL_array(worker_const,	           DAIL_count) = trim(worker)
            	    DAIL_array(maxis_case_number_const,    DAIL_count) = trim(maxis_case_number)
            	    DAIL_array(dail_type_const, 	       DAIL_count) = dail_type
            	    DAIL_array(dail_month_const, 		   DAIL_count) = trim(dail_month)
            	    DAIL_array(dail_msg_const, 		       DAIL_count) = trim(dail_msg)
                    DAIL_array(excel_row_const, 		   DAIL_count) = excel_row
                    DAIL_count = DAIL_count + 1

                    objExcel.Cells(excel_row, 1).Value = trim(worker)
                    objExcel.Cells(excel_row, 2).Value = trim(maxis_case_number)
                    objExcel.Cells(excel_row, 3).Value = dail_type
                    objExcel.Cells(excel_row, 4).Value = trim(dail_month)
                    objExcel.Cells(excel_row, 5).Value = trim(dail_msg)
                    excel_row = excel_row + 1
                    all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*") 'Adding MAXIS case number to case number string
                End if

				Call write_value_and_transmit("D", dail_row, 3)
				EMReadScreen other_worker_error, 13, 24, 2
				If other_worker_error = "** WARNING **" then transmit
				deleted_dails = deleted_dails + 1
			Else
                dail_row = dail_row + 1
            End if

			'...going to the next page if necessary
			EMReadScreen next_dail_check, 4, dail_row, 4
			If trim(next_dail_check) = "" then
				PF8
				EMReadScreen last_page_check, 21, 24, 2
                'DAIL/PICK when searching for specific DAIL types has message check of NO MESSAGES TYPE vs. NO MESSAGES WORK (for ALL DAIL/PICK selection).
                If last_page_check = "THIS IS THE LAST PAGE" or last_page_check = "NO MESSAGES TYPE" then
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

For item = 0 to Ubound(DAIL_array, 2)
    MAXIS_case_number = DAIL_array(maxis_case_number_const, item)
    Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
    EmReadscreen worker_county, 4, 21, 14
    If is_this_priv = True then
        DAIL_array(case_status_const, item) = "Privilged Case."
    Elseif worker_county <> worker_county_code then
        DAIL_array(case_status_const, item) = "Out-of-County Case."
    Else
        MAXIS_background_check
        Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, False)	'navigates to spec/memo and opens into edit mode
        'Writes the info into the MEMO.
        Call write_variable_in_SPEC_MEMO("************************************************************")
        Call write_variable_in_SPEC_MEMO("You are turning 60 next month, so you may be eligible for a new deduction for SNAP. Residents who are over 60 years old may receive increased SNAP benefits if they have recurring medical bills over $35 each month.")

        Call write_variable_in_SPEC_MEMO("If you have medical bills over $35 each month, please contact your team to discuss adjusting your benefits. You will need to send in proof of the medical bills, such as pharmacy receipts, an explanation of benefits, or premium notices.")
		CALL write_variable_in_SPEC_MEMO("")
		Call digital_experience
        Call write_variable_in_SPEC_MEMO("************************************************************")

        PF4
        EmReadscreen memo_confirmation, 26, 24, 2
        If memo_confirmation <> "NEW MEMO CREATE SUCCESSFUL" then
            DAIL_array(case_status_const, item) = "Unable to send MEMO. Process Manually."
        Else
            'THE CASE NOTE
            Call start_a_blank_CASE_NOTE
            Call write_variable_in_case_note("MEMB HAS TURNED 60-NOTIFY ABOUT POSSIBLE FMED DEDUCTION")
            Call write_variable_in_case_note("* Sent MEMO to client about FMED deductions.")
            Call write_variable_in_case_note("---")
            Call write_variable_in_case_note(worker_signature)

            PF3 'save CASE:NOTE
            DAIL_array(case_status_const, item) = "Success! MEMO sent."
            memo_count = memo_count + 1
        End if
    End if

    objExcel.Cells(DAIL_array(excel_row_const, item), 6).Value = DAIL_array(case_status_const, item)
Next

STATS_counter = STATS_counter - 1
'Enters info about runtime for the benefit of folks using the script
objExcel.Cells(2, 8).Value = "Number of DAILs processed:"
objExcel.Cells(3, 8).Value = "Number of Memo's sent to residents:"
objExcel.Cells(4, 8).Value = "Average time to find/select/copy/paste one line (in seconds):"
objExcel.Cells(5, 8).Value = "Estimated manual processing time (lines x average):"
objExcel.Cells(6, 8).Value = "Script run time (in seconds):"
objExcel.Cells(7, 8).Value = "Estimated time savings by using script (in minutes):"
objExcel.Cells(8, 8).Value = "Number of TIKL messages reviewed"
objExcel.Columns(8).Font.Bold = true
objExcel.Cells(2, 9).Value = deleted_dails
objExcel.Cells(3, 9).Value = memo_count
objExcel.Cells(4, 9).Value = STATS_manualtime
objExcel.Cells(5, 9).Value = STATS_counter * STATS_manualtime
objExcel.Cells(6, 9).Value = timer - start_time
objExcel.Cells(7, 9).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60
objExcel.Cells(8, 9).Value = STATS_counter

'Formatting the column width.
FOR i = 1 to 9
	objExcel.Columns(i).AutoFit()
NEXT

'saving the Excel file
file_info = month_folder & "\" & decimator_folder & "\" & report_date & " FMED Deduction PEPR's " & deleted_dails

'Saves and closes the most recent Excel workbook with the Task based cases to process.
objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\DAIL list\" & file_info & ".xlsx"

script_end_procedure_with_error_report("Success! Please review the list created for accuracy.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------08/10/2022
'--Tab orders reviewed & confirmed----------------------------------------------08/10/2022
'--Mandatory fields all present & Reviewed--------------------------------------08/10/2022
'--All variables in dialog match mandatory fields-------------------------------08/10/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------08/10/2022------------------N/A
'--CASE:NOTE Header doesn't look funky------------------------------------------08/10/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------08/10/2022------------------N/A
'
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------08/10/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------08/10/2022
'--PRIV Case handling reviewed -------------------------------------------------08/10/2022
'--Out-of-County handling reviewed----------------------------------------------08/10/2022
'--script_end_procedures (w/ or w/o error messaging)----------------------------08/10/2022
'--BULK - review output of statistics and run time/count (if applicable)--------08/10/2022
'--All strings for MAXIS entry are uppercase letters vs. lower case (Ex: "X")---08/10/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------08/10/2022
'--Incrementors reviewed (if necessary)-----------------------------------------08/10/2022
'--Denomination reviewed -------------------------------------------------------08/10/2022
'--Script name reviewed---------------------------------------------------------08/10/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------08/10/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub tasks are complete----------------------------------------08/10/2022
'--comment Code-----------------------------------------------------------------08/10/2022
'--Update Changelog for release/update------------------------------------------08/10/2022
'--Remove testing message boxes-------------------------------------------------08/10/2022
'--Remove testing code/unnecessary code-----------------------------------------08/10/2022
'--Review/update SharePoint instructions----------------------------------------08/10/2022: Instructions held locally at: T:\Eligibility Support\Mnsure\SSR Folder\SHELTER PROJECT\REPORTS Instructions\Laurie H Directions for Running Reports
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------08/10/2022------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------08/10/2022------------------N/A Will direct commmit Complete List of scripts
'--Complete misc. documentation (if applicable)---------------------------------08/10/2022
'--Update project team/issue contact (if applicable)----------------------------08/10/2022
