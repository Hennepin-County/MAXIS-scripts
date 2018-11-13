'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - DAIL CLEAN UP.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = 60
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
call changelog_update("11/05/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------DIALOG
BeginDialog dail_dialog, 0, 0, 266, 115, "DAIL INACTIVE CLEAN UP"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used once a month when cleaning up DAIL messages for inactive cases. The list requires a previously run list of cases, preferrably from the DAIL."
  Text 15, 70, 235, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""
MAXIS_footer_month = CM_mo 
MAXIS_footer_year = CM_yr 

'dialog and dialog DO...Loop	
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    Do
        err_msg = ""
    	Dialog dail_dialog
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If file_selection_path = "" then err_msg = "* Enter the file selection path."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    
'Setting a starting value for a list of cases so that every case is bracketed by * on both sides.
dail_cases = "*"

'Now the script adds all the clients on the excel list into an array
excel_row = 2 're-establishing the row to start checking the members for
entry_record = 0
Do   
    'Loops until there are no more cases in the Excel list
    MAXIS_case_number = objExcel.cells(excel_row, 2).Value          're-establishing the case numbers for functions to use
    MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do
    
    If instr(dail_cases, "*" & MAXIS_case_number & "*") = 0 then       'This indicates that the case number was not already found on the excel list 
        'msgbox MAXIS_case_number
        dail_cases = dail_cases & MAXIS_case_number & "*"       'adding the case number on the current row to the list of all the case numbers found.
    	entry_record = entry_record + 1			'This increments to the next entry in the array'
    	stats_counter = stats_counter + 1
    End if 
    excel_row = excel_row + 1
Loop

'msgbox entry_record
If left(dail_cases, 1) = "*" THEN dail_cases = right(dail_cases, len(dail_cases) - 1)
If right(dail_cases, 1) = "*" THEN dail_cases = left(dail_cases, len(dail_cases) - 1)
'msgbox "dail_cases " & dail_cases 
dail_array = split(dail_cases, "*")     

back_to_self
call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

'----------------------------------------------------------------------------------------------------Gathering the case status
inactive_cases = 0
entry = 0
DIM Inactive_array()
ReDim Inactive_array(2,0)

Const work_number = 0
Const case_number = 1
'Const case_status = 2

For each MAXIS_case_number in dail_array 
    
	MAXIS_case_number = replace(MAXIS_case_number, "*", "")
	call navigate_to_MAXIS_screen("CASE", "CURR")
    'msgbox MAXIS_case_number
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
	If PRIV_check = "PRIV" then	
		priv_list = priv_list & "," & MAXIS_case_number    
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the MAXIS case number
		transmit
    Else 
        EmReadscreen case_curr_status, 8, 8, 9
        IF case_curr_status = "INACTIVE" then 
            EmReadscreen worker_number, 7, 21, 14
            'Adding case information to the array
            ReDim Preserve Inactive_array(2, entry)	'This resizes the array based on the number of rows in the Excel File'
            Inactive_array(work_number,	entry) = worker_number
            Inactive_array(case_number, entry) = MAXIS_case_number	'The client information is added to the array'
            'Inactive_array(case_status, item) = case_curr_status
            entry = entry + 1
            inactive_cases = inactive_cases + 1
        End if 
    End if 
Next     

'msgbox inactive_cases 
'----------------------------------------------------------------------------------------------------'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to "DAIL List"
ObjExcel.ActiveSheet.Name = "Inactive Cases DAILs"

'Excel headers and formatting the columns
objExcel.Cells(1, 1).Value = "X NUMBER"
objExcel.Cells(1, 2).Value = "CASE #"
objExcel.Cells(1, 3).Value = "DAIL TYPE"
objExcel.Cells(1, 4).Value = "DAIL MO."
objExcel.Cells(1, 5).Value = "DAIL MESSAGE"
ObjExcel.Cells(1, 6).Value = case_status

FOR i = 1 to 6		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

excel_row = 2
MAXIS_case_number = ""
CALL navigate_to_MAXIS_screen("DAIL", "DAIL")

For item = 0 to UBound(Inactive_array, 2)
    worker_number      = Inactive_array(work_number, item)
	MAXIS_case_number  = Inactive_array(case_number, item)	
    
	'msgbox worker_number
	DO 
		EMReadScreen dail_check, 4, 2, 48
		If next_dail_check <> "DAIL" then 
			'MAXIS_case_number = ""
			CALL navigate_to_MAXIS_screen("DAIL", "DAIL")
		End if 
	Loop until dail_check = "DAIL"
	
	EMReadscreen dail_worker_number, 7, 21, 6
    If dail_worker_number <> worker_number then 
        EMWriteScreen worker_number, 21, 6
	    transmit
	    transmit 'transmit past 'not your dail message'
    End if 
	
    EMReadScreen number_of_dails, 1, 3, 67		'Reads where the count of DAILs is listed
	
	DO
		If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped
		dail_row = 6			'Because the script brings each new case to the top of the page, dail_row starts at 6.
		DO
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
            
            EmReadscreen DAIL_case_number, 8, dail_row - 1, 73
            msgbox DAIL_case_number
            If trim(DAIL_case_number) <> MAXIS_case_number then 
                case_found = false 
                dail_row = dail_row + 1
            else 
                case_found = true 
                msgbox case_found
                EMReadScreen dail_type, 4, dail_row, 6
    		    EMReadScreen dail_msg, 61, dail_row, 20
    		    dail_msg = trim(dail_msg)
                EMReadScreen dail_month, 8, dail_row, 11
    			'--------------------------------------------------------------------...and put that in Excel.
    			objExcel.Cells(excel_row, 1).Value = worker_number 
    			objExcel.Cells(excel_row, 2).Value = trim(maxis_case_number)
    			objExcel.Cells(excel_row, 3).Value = trim(dail_type)
    			objExcel.Cells(excel_row, 4).Value = trim(dail_month)
    			objExcel.Cells(excel_row, 5).Value = trim(dail_msg)
    		
    			Call write_value_and_transmit("D", dail_row, 3)	
    			EMReadScreen other_worker_error, 13, 24, 2
    			If other_worker_error = "** WARNING **" then transmit
                If trim(other_worker_error) <> "" then 
                    objExcel.Cells(excel_row, 6).Value = "Unable to delete."
                    msgbox other_worker_error
                    EmWriteScreen "_", dail_row, 3
                End if 
                dail_row = dail_row + 1
                excel_row = excel_row + 1
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

STATS_counter = STATS_counter - 1
'Enters info about runtime for the benefit of folks using the script
objExcel.Cells(2, 7).Value = "Number of DAILs deleted:"
objExcel.Cells(3, 7).Value = "Average time to find/select/copy/paste one line (in seconds):"
objExcel.Cells(4, 7).Value = "Estimated manual processing time (lines x average):"
objExcel.Cells(5, 7).Value = "Script run time (in seconds):"
objExcel.Cells(6, 7).Value = "Estimated time savings by using script (in minutes):"
objExcel.Cells(7, 7).Value = "Number of Dail messages reviewed"
objExcel.Columns(7).Font.Bold = true
objExcel.Cells(2, 8).Value = deleted_dails
objExcel.Cells(3, 8).Value = STATS_manualtime
objExcel.Cells(4, 8).Value = STATS_counter * STATS_manualtime
objExcel.Cells(5, 8).Value = timer - start_time
objExcel.Cells(6, 8).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60
objExcel.Cells(7, 8).Value = STATS_counter
objExcel.Cells(8, 8).Value = "Priv cases: " & priv_list 

'Formatting the column width.
FOR i = 1 to 8
	objExcel.Columns(i).AutoFit()
NEXT

script_end_procedure("Success! Please review the list created for accuracy.")