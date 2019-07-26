'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - EXPEDITED REVIEW.vbs"
start_time = timer
STATS_counter = 1                           'sets the stats counter at one
STATS_manualtime = 29                       'manual run time in seconds
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
CALL changelog_update("06/26/2019", "Initial version.", "Ilse Ferris, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOGS-------------------------------------------------------------------------------------------------------------
BeginDialog info_dialog, 0, 0, 266, 115, "ADMIN - EXPEDITED REVIEW"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used to review a BOBI list of pending SNAP and/or MFIP cases to ensure expedited screening and determinations are being made to ensure expedited timeliness rules are being followed."
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog


'Custom function for this script only: navigates to and checks case note for EXP screening case note--appears_exp will be input into the pending array, pending_array will
FUNCTION EXP_case_note_determination(appears_exp, pending_array)
	Call navigate_to_MAXIS_screen("CASE", "NOTE")

	'Checking for PRIV cases
	EMReadScreen priv_check, 6, 24, 14 			'If it can't get into the case needs to skip
	IF priv_check = "PRIVIL" or instr(priv_check, "NAT") THEN
		EMWriteScreen "________", 18, 43		'clears the case number
		transmit
		PF3
		pending_array(appears_exp, item) = true   'If the case is PRIV, then case is added to the excel spreadsheet to reviewed manually for EXP SNAP processing standards.
	ELse
		'starting at the 1st case note, checking the headers for the NOTES - EXPEDITED SCREENING text or the NOTES - EXPEDITED DETERMINATION text
		MAXIS_row = 5
		Do
			EMReadScreen case_note_date, 8, MAXIS_row, 6
			If trim(case_note_date) = "" then
				pending_array(appears_exp, item) = true 'if no case note exists, the case is added to the Excel list
				pending_array(case_notes, item) = "Expedited SNAP screening required"		'adds case notes to Excel re: screening is needed
				exit do
			else 
				EMReadScreen case_note_header, 55, MAXIS_row, 25
				case_note_header = lcase(trim(case_note_header))
				IF instr(case_note_header, "appears expedited") or instr(case_note_header, "appears expedit") then
					pending_array(appears_exp, item) = true            'if client appears exp is found, then case added to the Excel list
					pending_array(case_notes, item) = "EXP SNAP screened - appears expedited"		'adds case notes to Excel re: screening was completed
					exit do
				Elseif instr(case_note_header, "does not appear") then
                    pending_array(appears_exp, item) = false            'if client does not appear exp is found, then case will not be added to the Excel list
					exit do
				Else
					pending_array(appears_exp, item) = true			'defaults all other cases to true, to be addded to the Excel list
					pending_array(case_notes, item) = "Expedited SNAP screening required"		'adds case notes to Excel re: screening is needed
				END IF
			END IF
			MAXIS_row = MAXIS_row + 1
		LOOP until cdate(case_note_date) < cdate(appl_date)                        'repeats until the case note date is less than the application date
	END If
END FUNCTION

'THE SCRIPT-----------------------------------------------------------------------------------------------------------
EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr 

'dialog and dialog DO...Loop	
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog info_dialog
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

back_to_self
call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

DIM expedited_array()
ReDim expedited_array(5, 0)

'constants for array
const worker_number_const       = 0
const case_number_const	        = 1
const program_ID_const 	        = 2
const days_pending_const        = 3
const application_date_const    = 4
const case_status_const         = 5
const appears_exp_const         = 6

'Now the script adds all the clients on the excel list into an array
excel_row = 5 're-establishing the row to start checking the members for
entry_record = 0
Do   
    'Reading information from the BOBI report in Excel 
    worker_number = objExcel.cells(excel_row, 2).Value
    
    MAXIS_case_number = objExcel.cells(excel_row, 3).Value          're-establishing the case numbers for functions to use
    MAXIS_case_number = trim(MAXIS_case_number)
    
    program_ID = objExcel.cells(excel_row, 4).Value   
    program_ID = trim(program_ID)
    
    days_pending = objExcel.cells(excel_row, 7).Value
    days_pending = trim(days_pending) + 1   'This accounts for the data being a day behind 
    
    application_date = dateadd("D", days_pending, date) 
    
    If MAXIS_case_number = "" then exit do
    
    'Adding client information to the array'
    ReDim Preserve expedited_array(5, entry_record)	'This resizes the array based on the number of rows in the Excel File'
    expedited_array(worker_number_const, entry_record) = worker_number
    expedited_array(case_number_const, entry_record) = MAXIS_case_number		
    expedited_array(program_ID_const, entry_record) = program_ID        
    expedited_array(days_pending_const, entry_record) = days_pending         
    expedited_array(application_date_const, entry_record) = application_date           
    expedited_array(case_status_const, entry_record) = case_status
    expedited_array(appears_exp_const, entry_record) = ""

    entry_record = entry_record + 1			'This increments to the next entry in the array'
    stats_counter = stats_counter + 1
    excel_row = excel_row + 1
Loop



For item = 0 to UBound(expedited_array, 2)
    worker_number = expedited_array()
    
    expedited_array(worker_number_const, item) = worker_number
    
    
    
    Update_MMIS_array(clt_PMI, item) = client_pmi
    
    If instr(worker_number, "X127") then 
        case_status = "" 
        If 
        
        
    else 
        case_status = "Out-of-county case."
    End if 
    
    
'Loops until there are no more cases in the Excel list
    
	MAXIS_case_number = expedited_array(case_number ,item)	'Case number is set for each loop as it is used in the FuncLib functions'
	call navigate_to_MAXIS_screen("STAT", "PROG")
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
	If PRIV_check = "PRIV" then
		expedited_array(rate_two, item) = False  	
		expedited_array(case_status, item) = "PRIV case, cannot access/update." 
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the MAXIS case number
		transmit
    Else 
        EMReadscreen current_county, 4, 21, 21
        If lcase(current_county) <> worker_county_code then 
            expedited_array(rate_two, item) = False 
            expedited_array(case_status, item) = "Out-of-county case."
        Else 
            expedited_array(rate_two, item) = True  
        End if 
    End if 
    
    
    
	Call HCRE_panel_bypass			'Function to bypass a janky HCRE panel. If the HCRE panel has fields not completed/'reds up' this gets us out of there. 




''Sets constants for the array to make the script easier to read (and easier to code)
'Const work_num     = 1
'Const case_num     = 2		'Each of the case numbers will be stored at this position
'Const clt_name     = 3
'Const app_date     = 4
'Const days_pending = 5
'Const appears_exp  = 6      'appears_exp will be carried through to determine if the cases make it to the Excel list or not
'Const case_notes   = 7


''PND2 cases>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>PND2 information
'Adding another sheet
ObjExcel.Worksheets.Add().Name = "PND2 cases"

'Sets up the array to store all the information for each client'
Dim PND2_array ()
ReDim PND2_array (7, 0)
entry_record = 0

For each worker in worker_array
	back_to_self	'Does this to prevent "ghosting" where the old info shows up on the new screen for some reason
	Call navigate_to_MAXIS_screen("REPT", "PND2")
	EMWriteScreen worker, 21, 13
	transmit

	CALL find_variable("User: ", current_user, 7)
	IF ucase(worker) = ucase(current_user) THEN PF7

	'For PND2 cases, we can find SNAP pending cases and CASH pending case specifically. Adding SNAP and CASH pending cases to PND2 array.
	EMReadScreen has_content_check, 8, 7, 3  'Skips workers with no info
	If has_content_check <> "        " then
		'Grabbing each case number on screen
		Do
			'Set variable for next do...loop
			MAXIS_row = 7
			Do
				EMReadScreen SNAP_pending_status, 1, MAXIS_row, 62
				IF SNAP_pending_status <> "_" then add_to_PND2_array = true         'Adding pending SNAP cases to the PND2 array
				If SNAP_pending_status = "_" then
					EMReadScreen CASH_pending_status, 1, MAXIS_row, 54             'Adding pending CASH cases to the PND2 array
					If CASH_pending_status <> "_" then
						EMReadScreen CASH_program, 2, MAXIS_row, 56               'checking for specifc cash programs
                        If CASH_program = "MF" or CASH_program = "CA" then
						    add_to_PND2_array = true                              'if MF or undetermined cash program ("CA"), then these cases are added to the PND2 array
					    Else
						    add_to_PND2_array = false                             'all other cash cases are not added
                        End if
                    Else
                        add_to_PND2_array = false                               'all other cases are not added
					End if
				END IF

				EMReadScreen worker_basket, 7, 21, 13
				EMReadScreen MAXIS_case_number, 8, MAXIS_row, 5		 'Reading case number
				MAXIS_case_number = trim(MAXIS_case_number)
                EMReadScreen appl_date, 8, MAXIS_row, 38		     'Reading application date
				appl_date = replace(appl_date, " ", "/")
				EMReadScreen nbr_days_pending, 4, MAXIS_row, 49		 'Reading nbr days pending
				EMReadScreen client_name, 22, MAXIS_row, 16			 'Reading client name

				'Doing this because sometimes BlueZone registers a "ghost" of previous data when the script runs. This checks against an array and stops if we've seen this one before.
				client_name = trim(client_name)
				MAXIS_case_number = trim(MAXIS_case_number)
				If client_name <> "ADDITIONAL APP" Then			'When there is an additional app on this rept, the script actually reads a case number even though one is not visible to the worker on the screen - so we are skipping this ghosting issue because it will ALWAYS find the previous case number.
					If MAXIS_case_number <> "" and instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") <> 0 then exit do
					all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*")
				End If

				If MAXIS_case_number = "" AND client_name = "" Then Exit Do			'Exits do if we reach the end

				'If additional application is rec'd then the excel output is the client's name, not ADDITIONAL APP
				if client_name = "ADDITIONAL APP" then
					EMReadScreen alt_client_name, 22, MAXIS_row - 1, 16
					client_name = "* " & trim(alt_client_name)                    'replaces alt name as the client name
				Else
					EMReadScreen next_client, 22, MAXIS_row + 1, 16
					next_client = trim(next_client)
					If next_client = "ADDITIONAL APP" Then client_name = "* " & client_name
				END IF

				'Adding client information to the array'
				If add_to_PND2_array = true and MAXIS_case_number <> "" then
					ReDim Preserve PND2_array(7, entry_record)	'This resizes the array based on the number of rows in the Excel File'
					'The client information is added to the array'
					PND2_array (work_num,     entry_record) = worker_basket
					PND2_array (case_num,	  entry_record) = MAXIS_case_number
					PND2_array (clt_name,  	  entry_record) = client_name
					PND2_array (app_date, 	  entry_record) = appl_date
					PND2_array (days_pending, entry_record) = nbr_days_pending
					PND2_array (appears_exp,  entry_record) = true             'defaults appears_exp as true
					PND2_array (case_notes,   entry_record) = ""

					entry_record = entry_record + 1			'This increments to the next entry in the array
					STATS_counter = STATS_counter + 1
				END IF
				MAXIS_row = MAXIS_row + 1
			Loop until MAXIS_row = 19
			PF8
			EMReadScreen last_page_check, 21, 24, 2
		Loop until last_page_check = "THIS IS THE LAST PAGE"
	End if
next

'Now the script goes into STAT/PROG to determine if CASENOTE needs to be reviewed for the expedited screening casenote
If entry_record <> 0 then 
    For item = 0 to UBound(PND2_array, 2)
    	MAXIS_case_number = PND2_array(case_num, item)	'Case number for each loop from the array
    	appl_date = PND2_array(app_date, item)			'appl date for each loop from the array
    
    	back_to_self
    	EMWriteScreen MAXIS_case_number, 18, 43
        Call navigate_to_MAXIS_screen("STAT", "PROG")
    
    	'Checking for PRIV cases
    	EMReadScreen priv_check, 6, 24, 14 			'If it can't get into the case needs to skip
    	IF priv_check = "PRIVIL" THEN
    		PND2_array(appears_exp, item) = true    'If the case is PRIV, then case is added to the excel spreadsheet to reviewed manually for EXP SNAP processing standards.
    		EMWriteScreen "________", 18, 43		'clears the case number
    		transmit
    		PF3
    	ELse
    		'checking for ACTIVE SNAP
    		EMReadScreen SNAP_status, 4, 10, 74
    		If SNAP_status = "ACTV" then
    			check_case_note = false                  'if SNAP is active, the casenote is not searched as EXP does not need to be determined
    			PND2_array(appears_exp, item) = false    'case is not added to the Excel list
            elseIF SNAP_status = "PEND" then
    			check_case_note = true                   'if SNAP is pending, the casenote is searched to see if NOTES - EXPEDITED SCREENING has been completed
    		Else                                  'If SNAP is not active or pending.....
    			'Checking for ACTIVE MFIP
    			MAXIS_row = 6
    			Do
    				EMReadScreen cash_status, 2, MAXIS_row, 67
    				EMReadScreen program_status, 4, MAXIS_row, 74
    				If (cash_status = "MF" and program_status = "PEND") then
    						check_case_note = true                                'If MFIP is pending then the casenote is searched to see if NOTES - EXPEDITED SCREENING has been completed
    						exit do
    				Elseif (cash_status = "  " and program_status = "PEND") then
    					check_case_note = true                                    'If cash program is pending and undetermined by program then the casenote is searched to see if NOTES - EXPEDITED SCREENING has been completed
    					exit do
    				ELSE
                        check_case_note = false
    					PND2_array(appears_exp, item) = false                      'All other cases default to false, and the casenote is not searched as EXP does not need to be determined
    				END IF
    				MAXIS_row = MAXIS_row + 1
    			LOOP until MAXIS_row = 	8
    		END IF
    		'Because some cases don't have HCRE dates listed, so when you try to go past PROG the script gets caught up. Do...loop handles this instance.
    		PF3		'exits PROG to prompt HCRE if HCRE isn't complete
    		Do
    			EMReadscreen HCRE_panel_check, 4, 2, 50
    			If HCRE_panel_check = "HCRE" then
    				PF10	'exists edit mode in cases where HCRE isn't complete for a member
    				PF3
    			END IF
    		Loop until HCRE_panel_check <> "HCRE"		'repeats until case is not in the HCRE panel
    
            'If the case note needs to be reviewd for the NOTES - EXPEDITED SCREENING case note, then the
    		If check_case_note = true then Call EXP_case_note_determination(appears_exp, PND2_array) 'searching case notes for PND2 cases--See FUNCTION notes at top of script for detailed action as to what is happening here
        END IF
    NEXT
End if 

'adding information to the Excel list from PND2
ObjExcel.Cells(1, 1).Value = "Worker"
ObjExcel.Cells(1, 2).Value = "Case number"
ObjExcel.Cells(1, 3).Value = "Client name"
ObjExcel.Cells(1, 4).Value = "APPL date"
objExcel.Columns(4).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
ObjExcel.Cells(1, 5).Value = "# day pending"
ObjExcel.Cells(1, 6).Value = "NOTES"

FOR i = 1 to 6		'formatting the cells
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Addded the potentially EXP SNAP cases to the PND2 worksheet
excel_row = 2		'Setting the excel_row to start writing data on

If entry_record = 0 then 
    objExcel.Cells(excel_row, 1).Value = "No Pending SNAP/MFIP PND2 cases."
Else  
    For item = 0 to UBound(PND2_array, 2)
    	If PND2_array(appears_exp, item) = true then
    		objExcel.Cells(excel_row, 1).Value = PND2_array (work_num,   	item)	'Adding worker number
    		objExcel.Cells(excel_row, 2).Value = PND2_array (case_num,	 	item)	'Adding case number
    		objExcel.Cells(excel_row, 3).Value = PND2_array (clt_name, 	   	item)	'Addubg client name
    		objExcel.Cells(excel_row, 4).Value = PND2_array (app_date, 	   	item)	'Adding application date
    		objExcel.Cells(excel_row, 5).Value = PND2_array (days_pending, 	item)	'Adding number of days
    		objExcel.Cells(excel_row, 6).Value = PND2_array (case_notes, 	item)	'Adding notes re: what was found/not found in case notes
    		excel_row = excel_row + 1
    	End If
    Next
End if 

FOR i = 1 to 6		'formatting the cells
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'setting col to use to start writing run time information into to Excel
col_to_use = 8

'Query date/time/runtime info
objExcel.Cells(1, col_to_use - 1).Font.Bold = TRUE
objExcel.Cells(2, col_to_use - 1).Font.Bold = TRUE
ObjExcel.Cells(1, col_to_use - 1).Value = "Query date and time:"	'Goes back one, as this is on the next row
ObjExcel.Cells(1, col_to_use).Value = now
ObjExcel.Cells(2, col_to_use - 1).Value = "Query runtime (in seconds):"	'Goes back one, as this is on the next row
ObjExcel.Cells(2, col_to_use).Value = timer - query_start_time

ObjExcel.Cells(4, col_to_use - 1).Value = "Asterisks (*) indicates an ADDITIONAL APP exists."	'Row header
objExcel.Cells(4, col_to_use - 1).Font.Bold = TRUE						'Row header should be bold

'Autofitting columns
For col_to_autofit = 1 to col_to_use
	ObjExcel.columns(col_to_autofit).AutoFit()
Next

'logging usage stats
STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Please review the PND1 and PND2 lists for potential EXP SNAP processing.")
