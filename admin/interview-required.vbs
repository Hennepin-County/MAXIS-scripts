'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - INTERVIEW REQUIRED.vbs" 'BULK script that creates a list of cases that require an interview, and the contact phone numbers'
start_time = timer
STATS_counter = 1			     'sets the stats counter at one
STATS_manualtime = 	70			 'manual run time in seconds
STATS_denomination = "C"		 'C is for each case
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
call changelog_update("07/13/2019", "Added support for cases that have a ER and CSR in the same report month.", "Ilse Ferris, Hennepin County")
call changelog_update("01/16/2019", "Updated conditional handling and output of MFIP only cases.", "Ilse Ferris, Hennepin County")
call changelog_update("12/18/2018", "Updated to output two worksheets. One with ER case info, one with CSR case info.", "Ilse Ferris, Hennepin County")
call changelog_update("11/09/2018", "Added handling to export information about CSR's.", "Ilse Ferris, Hennepin County")
call changelog_update("09/18/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'DIALOG----------------------------------------------------------------------------------------------------
BeginDialog appointment_required_dialog, 0, 0, 286, 70, "Appointment required dialog"
  EditBox 70, 5, 210, 15, worker_number
  CheckBox 5, 40, 140, 10, "Select all active workers in the agency.", all_workers_check
  CheckBox 5, 55, 90, 10, "Current month plus two?", CM_plus_two_checkbox
  ButtonGroup ButtonPressed
    OkButton 175, 45, 50, 15
    CancelButton 230, 45, 50, 15
  Text 5, 10, 60, 10, "Worker number(s):"
  Text 5, 25, 275, 10, "Enter the fulll 7-digit worker number, separate each with a comma if more than one."
EndDialog

Function HCRE_panel_bypass() 
	'handling for cases that do not have a completed HCRE panel
	PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	Do
		EMReadscreen HCRE_panel_check, 4, 2, 50
		If HCRE_panel_check = "HCRE" then
			PF10	'exists edit mode in cases where HCRE isn't complete for a member
			PF3
		END IF
	Loop until HCRE_panel_check <> "HCRE"
End Function
		
'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone
all_workers_check = 1		'defaulting the check box to checked
CM_plus_two_checkbox = 1    'defaulting the check box to checked

'DISPLAYS DIALOG
DO
	DO
		err_msg = ""
		Dialog appointment_required_dialog
		If ButtonPressed = 0 then StopScript
		If worker_number = "" and all_workers_check = 0 then err_msg = err_msg & vbNewLine & "* Enter a valid worker number."
		if worker_number <> "" and all_workers_check = 1 then err_msg = err_msg & vbNewLine & "* Enter a worker number OR select the entire agency, not both." 
		If (CM_plus_two_checkbox = 1 and datePart("d", date) < 16) then err_msg = err_msg & VbNewLine & "* This is not a valid time period for REPT/REVS until the 16th of the month. Please select a new time period."
		IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
	LOOP until err_msg = ""
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS						
Loop until are_we_passworded_out = false					'loops until user passwords back in		

If CM_plus_two_checkbox = 1 then 
    REPT_month = CM_plus_2_mo
    REPT_year  = CM_plus_2_yr
Else 
    REPT_month = CM_plus_1_mo
    REPT_year  = CM_plus_1_yr
End if 

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

'We need to get back to SELF and manually update the footer month
back_to_self
'clears all data from the SELF screen
EMWriteScreen "____", 16, 43
EMWriteScreen "________", 18, 43
EMWriteScreen"____", 21, 70
transmit
transmit

Call navigate_to_MAXIS_screen("REPT", "REVS")
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46
transmit
EMWriteScreen REPT_month, 20, 55
EMWriteScreen REPT_year, 20, 58
transmit

'establishes counts and declaring arrays for recert cases with interview (SNAP/MFIP)
reviews_total = 0
total_cases_review = 0
DIM REVS_array()
REDim REVS_array(2, 0)

const case_number_const = 0
const snap_const = 1
const cash_const = 2

'start of the FOR...next loop
For each worker in worker_array
	If trim(worker) = "" then exit for
	worker_number = trim(worker)
	'writing in the worker number in the correct col
	EMWriteScreen worker, 21, 6
	transmit

    'Grabbing case numbers from REVS for requested worker
	DO	'All of this loops until last_page_check = "THIS IS THE LAST PAGE"
		MAXIS_row = 7	'Setting or resetting this to look at the top of the list
		DO		'All of this loops until MAXIS_row = 19
			'Reading case information (case number, SNAP status, and cash status)
			EMReadScreen MAXIS_case_number, 8, MAXIS_row, 6
			MAXIS_case_number = trim(MAXIS_case_number)
			EMReadScreen SNAP_status, 1, MAXIS_row, 45
			EMReadScreen cash_status, 1, MAXIS_row, 39
            
			'Navigates though until it runs out of case numbers to read
			IF MAXIS_case_number = "" then exit do

			'For some goofy reason the dash key shows up instead of the space key. No clue why. This will turn them into null variables.
			If cash_status = "-" 	then cash_status = ""
			If SNAP_status = "-" 	then SNAP_status = ""
			If HC_status = "-" 		then HC_status = ""
            
            SNAP_revw = ""
            CASH_revw = ""

			'Using if...thens to decide if a case should be added (status isn't blank and respective box is checked)
			If trim(SNAP_status) = "N" then 
                SNAP_revw = True 
                add_case = True 
            End if 
            
            If trim(cash_status) = "N" then 
                CASH_revw = True 
                add_case = true 
            End if 
            
            If (cash_status <> "N" AND snap_status <> "N") then 
                SNAP_revw = False 
                CASH_revw = False
                add_case = False  
            End if
             
            If add_case = True then 
                'msgbox MAXIS_case_number & vbcr & reviews_total + 1
				ReDim Preserve REVS_array(2, reviews_total)				        'This resizes the array based on the number of members being added to the array
				REVS_array(case_number_const, reviews_total) = MAXIS_case_number
                If SNAP_revw = True then 
                    snap_program = "SNAP"
                Else 
                    snap_program = ""
                End if 
            
                If CASH_revw = True then 
                    cash_program = "CASH"
                Else 
                    cash_program = ""
                End if 
                
                'msgbox MAXIS_case_number & vbcr & SNAP_program & vbcr & cash_program
                
                REVS_array(snap_const, reviews_total) = snap_program
                REVS_array(cash_const, reviews_total) = cash_program
				reviews_total = reviews_total + 1
				total_cases_review = total_cases_review + 1
			End if
			'On the next loop it must look to the next row
			MAXIS_row = MAXIS_row + 1

			'Clearing variables before next loop
			add_to_array = ""
			MAXIS_case_number = ""
		Loop until MAXIS_row = 19		'Last row in REPT/REVS
		'Because we were on the last row, or exited the do...loop because the case number is blank, it PF8s, then reads for the "THIS IS THE LAST PAGE" message (if found, it exits the larger loop)
		PF8
		EMReadScreen last_page_check, 21, 24, 2	'checking to see if we're at the end
        'if max reviews are reached, the goes to next worker is applicable
	Loop until last_page_check = "THIS IS THE LAST PAGE"
next
'msgbox reviews_total

recert_cases = 0	'value for the array
DIM Required_appt_array()
ReDim Required_appt_array(7, 0)

const basket_number = 0
const case_number	= 1
const active_progs  = 2
const case_interp	= 3
const case_lang		= 4
const phone_one     = 5
const phone_two     = 6
const phone_three   = 7

'establishes counts and declaring arrays for CSR cases. 
CSR_count = 0
DIM CSR_array()
REDim CSR_array(5, 0)

'constants for array
const basket_number_csr = 0
const case_number_csr	= 1
const active_progs_csr  = 2
const case_interp_csr	= 3
const case_lang_csr		= 4
const dup_exists_csr    = 5

MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr 
Call MAXIS_footer_month_confirmation

worker_number = ""
back_to_SELF

'DO 'Loops until there are no more cases in the Excel list
For item = 0 to uBound(REVS_array, 2) 
	MAXIS_case_number = REVS_array(case_number_const, item)
    SNAP_program = REVS_array(snap_const, item)
    cash_program = REVS_array(cash_const, item)
    
    'msgbox MAXIS_case_number & vbcr & SNAP_program & vbcr & cash_program
	CALL navigate_to_MAXIS_screen("STAT", "PROG")		'Goes to STAT/PROG
	EMReadScreen wrkr_numb, 7, 21, 21
	
    'Checking for PRIV cases.
	EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip
	IF priv_check = "PRIV" or instr(priv_check, "NAT") THEN 
		priv_case_list = priv_case_list & "|" & MAXIS_case_number
	ELSE						'For all of the cases that aren't privileged...
		MFIP_ACTIVE = FALSE		'Setting some variables for the loop
		SNAP_ACTIVE = FALSE	
        GRH_ACTIVE  = FALSE 
        
		SNAP_status_check = ""
        MFIP_prog_1_check = ""
		MFIP_status_1_check = ""
		MFIP_prog_2_check = ""
		MFIP_status_2_check = ""
        GRH_status_check = ""
        
        If SNAP_program = "SNAP" then 
            SNAP_ACTIVE = TRUE
        Else 
            SNAP_ACTIVE = FALSE 
        End if 
        
        If cash_program = "CASH" then 
		    'Reading the status and program
		    EMReadScreen MFIP_prog_1_check, 2, 6, 67		'checking for an active MFIP case
		    EMReadScreen MFIP_status_1_check, 4, 6, 74
		    EMReadScreen MFIP_prog_2_check, 2, 6, 67		'checking for an active MFIP case
		    EMReadScreen MFIP_status_2_check, 4, 6, 74
            EmReadscreen GRH_status_check, 4, 9, 74          'GRH cases for CSR array
            
            'Logic to determine if MFIP is active
    		If MFIP_prog_1_check = "MF" Then
    			If MFIP_status_1_check = "ACTV" Then MFIP_ACTIVE = TRUE
    		ElseIf MFIP_prog_2_check = "MF" Then
    			If MFIP_status_2_check = "ACTV" Then MFIP_ACTIVE = TRUE
            Else 
                MFIP_ACTVIE = FALSE
    		End If
    		
            If GRH_status_check = "ACTV" then 
                GRH_ACTIVE = TRUE
            Else 
                GRH_ACTIVE = FALSE
            END IF 
            
        End if     
		
        'msgbox MAXIS_case_number & vbcr & "SNAP active: " & snap_active & vbcr & " MFIP active: " & MFIP_active  
        'msgbox MFIP_active & vbcr & SNAP_active & vbcr & GRH_active
		    
		HCRE_panel_bypass	'function I created to ensure that we don't get trapped in the HCRE panel
        
        recert_status = "NO"	'Defaulting this to no because if SNAP or MFIP are not active - no recert will be scheduled
		program_list = ""
        CALL navigate_to_MAXIS_screen("STAT", "REVW") 'Going to STAT/REVW to to check for ER vs CSR for SNAP cases
        
        If SNAP_ACTIVE = TRUE Then
			EMReadScreen SNAP_review_check, 8, 9, 57
			If SNAP_review_check = "__ 01 __" then 		'If this is blank there are big issues
				recert_status = "NO"
			Else
				EMwritescreen "x", 5, 58		'Opening the SNAP popup
				Transmit
				DO
				    EMReadScreen SNAP_popup_check, 7, 5, 43
				LOOP until SNAP_popup_check = "Reports"

				'The script will now read the CSR MO/YR and the Recert MO/YR
				EMReadScreen CSR_mo, 2, 9, 26
				EMReadScreen CSR_yr, 2, 9, 32
				EMReadScreen recert_mo, 2, 9, 64
				EMReadScreen recert_yr, 2, 9, 70

				'Comparing CSR and ER daates to the month of REVS review
				IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN 
                    CSR_month = True 
                    If CSR_mo = recert_mo then Duplicate_exists = True 
                else 
                    CSR_month = False
                End if 
				'If recert_mo = left(REPT_month, 2) and recert_yr <> right(REPT_year, 2) THEN recert_status = "NO"
				IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) THEN recert_status = "YES"
                
                If CSR_month = True or recert_status = "YES" then program_list = program_list & "SNAP & "    
			End If
        End if 
        
		If GRH_ACTIVE = TRUE then
            EMwritescreen "x", 5, 35		'Opening the CASH pop-up
            Transmit
            'The script will now read the CSR MO/YR and the Recert MO/YR
            EMReadScreen CSR_mo, 2, 9, 26
            EMReadScreen CSR_yr, 2, 9, 32
            EMReadScreen recert_mo, 2, 9, 64
            EMReadScreen recert_yr, 2, 9, 70

            'Comparing CSR and ER daates to the month of REVS review
            IF CSR_mo = left(REPT_month, 2) and CSR_yr = right(REPT_year, 2) THEN 
                CSR_month = True 
                If CSR_mo = recert_mo then Duplicate_exists = True 
            else 
                CSR_month = False
            End if 
        
            IF recert_mo = left(REPT_month, 2) and recert_yr = right(REPT_year, 2) THEN recert_status = "YES"
            If CSR_month = True or recert_status = "YES" then program_list = program_list & "GRH & "  
        Else 
            recert_status = "NO"    'defaulting everything else (HC, MSA, GRH only) as no interview 
            CSR_month = False       'defaulting all non GRH or SNAP as non CSR 
        End if 
        
        If MFIP_ACTIVE = TRUE Then 
            recert_status = "YES"	'MFIP will only have an ER - so if listed on REVS - will be an ER - don't need to check dates
            program_list = program_list & "MFIP & "
            'msgbox "MFIP ACTIVE TRUE" & Vbcr & program_list
        End if
	    
        program_list = trim(program_list)       'trims excess spaces of program_list
        If right(program_list, 1) = "&" THEN program_list = left(program_list, len(program_list) - 1)
        
        'msgbox MAXIS_case_number & vbcr & recert_status
		If recert_status = "YES" then 
			Redim Preserve Required_appt_array(7, 	recert_cases)
			Required_appt_array (case_number, 		recert_cases) = MAXIS_case_number
			Required_appt_array (basket_number, 	recert_cases) = wrkr_numb
			Required_appt_array (active_progs, recert_cases) = program_list
            			
			'Gathering the phone numbers
			call navigate_to_MAXIS_screen("STAT", "ADDR")
			EMReadScreen phone_number_one, 16, 17, 43	' if phone numbers are blank it doesn't add them to EXCEL
			If phone_number_one <> "( ___ ) ___ ____" then Required_appt_array (phone_one, recert_cases) = phone_number_one
			EMReadScreen phone_number_two, 16, 18, 43
			If phone_number_two <> "( ___ ) ___ ____" then Required_appt_array (phone_two, recert_cases) = phone_number_two
			EMReadScreen phone_number_three, 16, 19, 43
			If phone_number_three <> "( ___ ) ___ ____" then Required_appt_array (phone_three, recert_cases) = phone_number_three	
			
			'Going to STAT/MEMB for Language Information
			CALL navigate_to_MAXIS_screen("STAT", "MEMB")
			EMReadScreen interpreter_code, 1, 14, 68
			EMReadScreen language_coded, 16, 12, 46
			language_coded = replace(language_coded, "_", "")
			If trim(language_coded) = "" then 
				EMReadScreen lang_ID, 2, 12, 42
				If lang_ID = "99" then lang_ID = "English"
				language_coded = lang_ID
			End if 
			
			Required_appt_array (case_interp,  recert_cases) = interpreter_code
			Required_appt_array (case_lang,    recert_cases) = language_coded
			recert_cases = recert_cases + 1
			STATS_counter = STATS_counter + 1						'adds one instance to the stats counter
        End if 
            '----------------------------------------------------------------------------------------------------Gathering case info for CSR cases 
		If CSR_month = true then 
            Redim Preserve CSR_array(5, CSR_count)
            CSR_array (case_number_csr, 	CSR_count) = MAXIS_case_number
            CSR_array (basket_number_csr, 		CSR_count) = wrkr_numb
            CSR_array(active_progs_csr, CSR_count) = program_list
                
            'Going to STAT/MEMB for Language Information
            CALL navigate_to_MAXIS_screen("STAT", "MEMB")
            EMReadScreen interpreter_code, 1, 14, 68
            EMReadScreen language_coded, 16, 12, 46
            language_coded = replace(language_coded, "_", "")
            If trim(language_coded) = "" then 
                EMReadScreen lang_ID, 2, 12, 42
                If lang_ID = "99" then lang_ID = "English"
                language_coded = lang_ID
            End if 
            
            CSR_array (case_interp_csr,  CSR_count) = interpreter_code
            CSR_array (case_lang_csr,    CSR_count) = language_coded
            CSR_array(dup_exists_csr, CSR_count) = Duplicate_exists
            CSR_count = CSR_count + 1
            STATS_counter = STATS_counter + 1						'adds one instance to the stats counter
        End if 
	End if 	
Next

'----------------------------------------------------------------------------------------------------EXCEL INPUT
'Opening the Excel file, (now that the dialog is done)
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet to "Case information"
ObjExcel.ActiveSheet.Name = "ER cases " & REPT_month & "-" & REPT_year

'formatting excel file with columns for case number and interview date/time
objExcel.cells(1, 1).value 	= "X number"
objExcel.cells(1, 2).value 	= "Case number"
objExcel.cells(1, 3).value 	= "Programs"
objExcel.cells(1, 4).value 	= "Case language"
objExcel.Cells(1, 5).value 	= "Interpreter"
objExcel.cells(1, 6).value 	= "Phone # One"
objExcel.cells(1, 7).value 	= "Phone # Two"
objExcel.Cells(1, 8).value 	= "Phone # Three"
'objExcel.cells(1, 9).value 	= "Privileged Cases"
	
FOR i = 1 to 9									'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Adding the case information to Excel
excel_row = 2
For item = 0 to UBound(Required_appt_array, 2)
	ObjExcel.Cells(excel_row, 1).value  = Required_appt_array (basket_number,     item)
	ObjExcel.Cells(excel_row, 2).value  = Required_appt_array (case_number,  item)
	ObjExcel.Cells(excel_row, 3).value  = Required_appt_array (active_progs, item)
	ObjExcel.Cells(excel_row, 4).value  = Required_appt_array (case_lang,    item)
	ObjExcel.Cells(excel_row, 5).value  = Required_appt_array (case_interp,  item)
	ObjExcel.Cells(excel_row, 6).value = Required_appt_array (phone_one,     item)
	ObjExcel.Cells(excel_row, 7).value = Required_appt_array (phone_two,     item)
	ObjExcel.Cells(excel_row, 8).value = Required_appt_array (phone_three,   item)
	excel_row = excel_row + 1 
Next

''Creating the list of privileged cases and adding to the spreadsheet
'priv_case_list = priv_case_list & MAXIS_case_number & "|"
'priv_case_list = right(priv_case_list, (len(priv_case_list)-1))
'prived_case_array = split(priv_case_list, "|")
'
'excel_row = 2
'
'FOR EACH MAXIS_case_number in prived_case_array
'	objExcel.cells(excel_row, 9).value = MAXIS_case_number
'	excel_row = excel_row + 1
'NEXT

'Query date/time/runtime info
objExcel.Cells(1, 10).Font.Bold = TRUE
objExcel.Cells(2, 10).Font.Bold = TRUE
objExcel.Cells(3, 10).Font.Bold = TRUE
objExcel.Cells(4, 10).Font.Bold = TRUE
ObjExcel.Cells(1, 10).Value = "Query date and time:"	
ObjExcel.Cells(2, 10).Value = "Query runtime (in seconds):"	
ObjExcel.Cells(3, 10).Value = "Total reviews:"
ObjExcel.Cells(4, 10).Value = "Interview required:"
ObjExcel.Cells(1, 11).Value = now
ObjExcel.Cells(2, 11).Value = timer - query_start_time
ObjExcel.Cells(3, 11).Value = total_cases_review
ObjExcel.Cells(4, 11).Value = recert_cases

'Formatting the columns to autofit after they are all finished being created.
FOR i = 1 to 11
	objExcel.Columns(i).autofit()
Next

'Adding another sheet
ObjExcel.Worksheets.Add().Name = "CSR cases " & REPT_month & "-" & REPT_year

'formatting excel file with columns for case number and interview date/time
objExcel.cells(1, 1).value 	= "X number"
objExcel.cells(1, 2).value 	= "Case number"
objExcel.cells(1, 3).value 	= "Programs"
objExcel.cells(1, 4).value 	= "Case language"
objExcel.Cells(1, 5).value 	= "Interpreter"
objExcel.cells(1, 6).value 	= "Case also has ER"

	
FOR i = 1 to 6									'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Adding the case information to Excel
excel_row = 2
For item = 0 to UBound(CSR_array, 2)
	ObjExcel.Cells(excel_row, 1).value = CSR_array(basket_number_csr,     item)
	ObjExcel.Cells(excel_row, 2).value = CSR_array(case_number_csr,  item)
	ObjExcel.Cells(excel_row, 3).value = CSR_array(active_progs_csr, item)
	ObjExcel.Cells(excel_row, 4).value = CSR_array(case_lang_csr,    item)
	ObjExcel.Cells(excel_row, 5).value = CSR_array(case_interp_csr,  item)
	ObjExcel.Cells(excel_row, 6).value = CSR_array(dup_exists_csr,   item)
	excel_row = excel_row + 1 
Next

'Query date/time/runtime info
objExcel.Cells(1, 7).Font.Bold = TRUE
objExcel.Cells(2, 7).Font.Bold = TRUE
objExcel.Cells(3, 7).Font.Bold = TRUE
objExcel.Cells(4, 7).Font.Bold = TRUE
ObjExcel.Cells(1, 7).Value = "Query date and time:"	
ObjExcel.Cells(2, 7).Value = "Query runtime (in seconds):"	
ObjExcel.Cells(3, 7).Value = "Total reviews:"
ObjExcel.Cells(4, 7).Value = "CSR cases:"
ObjExcel.Cells(1, 8).Value = now
ObjExcel.Cells(2, 8).Value = timer - query_start_time
ObjExcel.Cells(3, 8).Value = total_cases_review
ObjExcel.Cells(4, 8).Value = CSR_count

'Formatting the columns to autofit after they are all finished being created.
FOR i = 1 to 8
	objExcel.Columns(i).autofit()
Next

STATS_counter = STATS_counter - 1
script_end_procedure("PRIV case list: " & priv_case_list)