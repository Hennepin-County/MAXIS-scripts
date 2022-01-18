'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - DAIL CCD.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = 30
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
call changelog_update("01/18/2022", "Added out-of-county handling.", "Ilse Ferris, Hennepin County")
call changelog_update("04/06/2020", "Added autosave functionality.", "Ilse Ferris, Hennepin County")
call changelog_update("05/20/2019", "Removed output of all actionable DAIL messages to end of script run. Default all workers checkbox to checked.", "Ilse Ferris, Hennepin County")
call changelog_update("03/16/2019", "Added output of all actionable DAIL messages to end of script run.", "Ilse Ferris, Hennepin County")
call changelog_update("12/14/2018", "Updated DAIL selection to INFO only to reduce run time.", "Ilse Ferris, Hennepin County")
call changelog_update("10/31/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display

Function dail_selection
	'selecting the type of DAIl message
	EMWriteScreen "x", 4, 12		'transmits to the PICK screen
	transmit
	EMWriteScreen "_", 7, 39		'clears the all selection

    IF dail_to_decimate = "ALL" then selection_row = 7
    IF dail_to_decimate = "CSES" then selection_row = 10
	IF dail_to_decimate = "COLA" then selection_row = 8
	IF dail_to_decimate = "ELIG" then selection_row = 11
	IF dail_to_decimate = "INFO" then selection_row = 13
    IF dail_to_decimate = "PEPR" then selection_row = 18

	Call write_value_and_transmit("x", selection_row, 39)
End Function

'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""
dail_to_decimate = "INFO"    'defaults to all. Some x-numbers don't select the DAIL hence the default.
all_workers_check = 1   'checked

this_month = CM_mo & " " & CM_yr
next_month = CM_plus_1_mo & " " & CM_plus_1_yr
CM_minus_2_mo =  right("0" & DatePart("m", DateAdd("m", -2, date)), 2)

'Finding the right folder to automatically save the file
month_folder = "DAIL " & CM_mo & "-" & DatePart("yyyy", date) & ""
decimator_folder = replace(this_month, " ", "-") & " DAIL Decimator"
report_date = replace(date, "/", "-")

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 95, "DAIL CAPTURE, CASE NOTE, DELETE (CCD)"
  EditBox 80, 55, 180, 15, worker_number
  CheckBox 15, 80, 135, 10, "Check here to process for all workers.", all_workers_check
  ButtonGroup ButtonPressed
    OkButton 155, 75, 50, 15
    CancelButton 210, 75, 50, 15
  Text 15, 60, 60, 10, "Worker number(s):"
  GroupBox 10, 5, 250, 45, "Using the DAIL Decimator script"
  Text 20, 20, 235, 25, "This script should be used to remove and case note messages that have been determined by Quality Improvement staff do not require action besides a case note."
EndDialog
'the dialog
Do
	Do
  		err_msg = ""
  		dialog Dialog1
  		cancel_without_confirmation
  		If trim(worker_number) = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases."
  		If trim(worker_number) <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Select a worker number(s) or all cases, not both options."
  	  	IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
  	LOOP until err_msg = ""
  	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call check_for_MAXIS(False)

'If all workers are selected, the script will go to REPT/USER, and load all of the workers into an array. Otherwise it'll create a single-object "array" just for simplicity of code.
If all_workers_check = checked then
	call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
Else
	x1s_from_dialog = split(worker_number, ", ")	'Splits the worker array based on commas

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
ObjExcel.ActiveSheet.Name = "Deleted DAILS - INFO"

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X NUMBER"
objExcel.Cells(1, 2).Value = "CASE #"
objExcel.Cells(1, 3).Value = "DAIL TYPE"
objExcel.Cells(1, 4).Value = "DAIL MO."
objExcel.Cells(1, 5).Value = "DAIL MESSAGE"
objExcel.Cells(1, 6).Value = "DAIL NOTES"

FOR i = 1 to 6		'formatting the cells'
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

MAXIS_case_number = ""
CALL navigate_to_MAXIS_screen("DAIL", "DAIL")

'This for...next contains each worker indicated above
For each worker in worker_array
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

	Call dail_selection

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

            EMReadScreen maxis_case_number, 8, dail_row - 1, 73
            EMReadScreen dail_month, 8, dail_row, 11
			EMReadScreen dail_type, 4, dail_row, 6
			EMReadScreen dail_msg, 61, dail_row, 20
            dail_msg = trim(dail_msg)
            If right(dail_msg, 1) = "*" THEN dail_msg = left(dail_msg, len(dail_msg) - 1)
            dail_msg = trim(dail_msg)
			stats_counter = stats_counter + 1

			If instr(dail_msg, "SDX MATCH - PBEN UPDATED - MAXIS INTERFACED IAA DATE TO SSA") OR _
               instr(dail_msg, "MEMBER HAS TURNED 60 - FSET:WORK REG HAS BEEN UPDATED") OR _
   			   instr(dail_msg, "SDX MATCH - MAXIS INTERFACED IAA DATE TO SSA") then
               add_to_excel = TRUE
            elseif instr(dail_msg, "CANCELLED DUE TO AGING") then
                if left(dail_msg, 2) = "$0" then
                    add_to_excel = TRUE     'issuance is under $1.00 - Delete and case note
                Else
                    add_to_excel = False    'over $1.00
                End if
            elseif instr(dail_msg, "MEMI:CITIZENSHIP HAS BEEN VERIFIED THROUGH SSA") then
                EmReadscreen case_name, 56, dail_row - 1, 5
                dail_msg = dail_msg & " " & replace(case_name,"-", "")
                'msgbox dail_msg
                add_to_excel = true
            else
			    add_to_excel = False
			End if

            'instr(dail_msg, "HC AUTOCLOSED FOR RENEWAL AND ELIGIBILITY FOR METS") OR _ ---removed for now. Unable to find current messages 12/14/2018

			IF add_to_excel = True then
				'--------------------------------------------------------------------...and put that in Excel.
				objExcel.Cells(excel_row, 1).Value = worker
				objExcel.Cells(excel_row, 2).Value = trim(maxis_case_number)
				objExcel.Cells(excel_row, 3).Value = trim(dail_type)
				objExcel.Cells(excel_row, 4).Value = trim(dail_month)
				objExcel.Cells(excel_row, 5).Value = trim(dail_msg)
				excel_row = excel_row + 1

				Call write_value_and_transmit("D", dail_row, 3)
				EMReadScreen other_worker_error, 13, 24, 2
				If other_worker_error = "** WARNING **" then transmit
				deleted_dails = deleted_dails + 1
			else
				add_to_excel = False
				dail_row = dail_row + 1
                ReDim Preserve DAIL_array(4, DAIL_count)	'This resizes the array based on the number of rows in the Excel File'
            	DAIL_array(worker_const,	           DAIL_count) = worker
            	DAIL_array(maxis_case_number_const,    DAIL_count) = MAXIS_case_number
            	DAIL_array(dail_type_const, 	       DAIL_count) = dail_type
                If len(dail_month) = 5 then dail_month = replace(dail_month, " ", "/1/")
            	DAIL_array(dail_month_const, 		   DAIL_count) = dail_month
            	DAIL_array(dail_msg_const, 		       DAIL_count) = dail_msg
                Dail_count = DAIL_count + 1
			End if

			EMReadScreen message_error, 11, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.
			If message_error = "NO MESSAGES" then
				CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
				Call write_value_and_transmit(worker, 21, 6)
				transmit   'transmit past 'not your dail message'
				Call dail_selection
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

dail_msg = ""
excel_row = 2

Do
    MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
    MAXIS_case_number = trim(MAXIS_case_number)

    dail_msg = ObjExcel.Cells(excel_row, 5).Value
    dail_msg = trim(dail_msg)
    'Cleaning up the DAIL messages for the case note
    If right(dail_msg, 9) = "-SEE PF12" THEN dail_msg = left(dail_msg, len(dail_msg) - 9)
    If right(dail_msg, 1) = "*" THEN dail_msg = left(dail_msg, len(dail_msg) - 1)
    dail_msg = trim(dail_msg)

    Call navigate_to_MAXIS_screen("CASE", "NOTE")
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list

    If PRIV_check = "PRIV" then
        objExcel.Cells(excel_row, 6).Value = "PRIV, unable to case note."
        'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
    	Do
    		back_to_self
    		EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
    		If SELF_screen_check <> "SELF" then PF3
    	LOOP until SELF_screen_check = "SELF"
    	EMWriteScreen "________", 18, 43		'clears the MAXIS case number
    	transmit
    Else
        EmReadscreen county_check, 2, 21, 16
        If county_check <> "27" then
            objExcel.Cells(excel_row, 6).Value = "Out of county case."
        Else
            PF9
            CALL write_variable_in_case_note(dail_msg)
            CALL write_variable_in_case_note("")
            PF3 ' save message
            objExcel.Cells(excel_row, 6).Value = "Case note created."
        End if
    End If
    excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 2).Value = ""

STATS_counter = STATS_counter - 1
'Enters info about runtime for the benefit of folks using the script
objExcel.Cells(2, 7).Value = "Number of DAILs deleted:"
objExcel.Cells(3, 7).Value = "Average time to find/select/copy/paste one line (in seconds):"
objExcel.Cells(4, 7).Value = "Estimated manual processing time (lines x average):"
objExcel.Cells(5, 7).Value = "Script run time (in seconds):"
objExcel.Cells(6, 7).Value = "Estimated time savings by using script (in minutes):"
objExcel.Cells(7, 7).Value = "Number of " & dail_to_decimate & " messages reviewed"
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

'saving the Excel file
file_info = month_folder & "\" & decimator_folder & "\" & report_date & " CCD " & deleted_dails

'Saves and closes the most recent Excel workbook with the Task based cases to process.
objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\DAIL list\" & file_info & ".xlsx"

script_end_procedure("Success! Please review the list created for accuracy.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------01/18/2022
'--Tab orders reviewed & confirmed----------------------------------------------01/18/2022
'--Mandatory fields all present & Reviewed--------------------------------------01/18/2022
'--All variables in dialog match mandatory fields-------------------------------01/18/2022
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------01/18/2022
'--CASE:NOTE Header doesn't look funky------------------------------------------01/18/2022
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------01/18/2022--------------------N/A
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------01/18/2022
'--MAXIS_background_check reviewed (if applicable)------------------------------01/18/2022---------------------N/A
'--PRIV Case handling reviewed -------------------------------------------------01/18/2022
'--Out-of-County handling reviewed----------------------------------------------01/18/2022
'--script_end_procedures (w/ or w/o error messaging)----------------------------01/18/2022
'--BULK - review output of statistics and run time/count (if applicable)--------01/18/2022
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------01/18/2022
'--Incrementors reviewed (if necessary)-----------------------------------------01/18/2022
'--Denomination reviewed -------------------------------------------------------01/18/2022
'--Script name reviewed---------------------------------------------------------01/18/2022
'--BULK - remove 1 incrementor at end of script reviewed------------------------01/18/2022

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------01/18/2022
'--comment Code-----------------------------------------------------------------01/18/2022
'--Update Changelog for release/update------------------------------------------01/18/2022
'--Remove testing message boxes-------------------------------------------------01/18/2022
'--Remove testing code/unnecessary code-----------------------------------------01/18/2022
'--Review/update SharePoint instructions----------------------------------------01/18/2022---------------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------01/18/2022---------------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------01/18/2022
'--Complete misc. documentation (if applicable)---------------------------------01/18/2022
'--Update project team/issue contact (if applicable)----------------------------01/18/2022
