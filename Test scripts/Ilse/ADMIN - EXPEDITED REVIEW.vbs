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
CALL changelog_update("01/15/2020", "Initial version.", "Ilse Ferris, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-----------------------------------------------------------------------------------------------------------
EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr 

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "ADMIN - EXPEDITED REVIEW"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used to review a BOBI list of pending SNAP and/or MFIP cases to ensure expedited screening and determinations are being made to ensure expedited timeliness rules are being followed."
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog

'dialog and dialog DO...Loop	
Do 
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation 
        If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue." 
        If err_msg <> "" Then MsgBox err_msg
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

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
all_case_numbers_array = "*"
Do   
    'Reading information from the BOBI report in Excel 
    worker_number = objExcel.cells(excel_row, 2).Value
    worker_number = trim(worker_number)
    
    MAXIS_case_number = objExcel.cells(excel_row, 3).Value          're-establishing the case numbers for functions to use
    MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do
    
    program_ID = objExcel.cells(excel_row, 4).Value   
    program_ID = trim(program_ID)
    
    days_pending = objExcel.cells(excel_row, 7).Value
    days_pending = trim(days_pending) + 1   'This accounts for the data being a day behind 
    
    application_date = dateadd("D", days_pending, date) 
    
    msgbox excel_row & vbcr & program_ID
    
    'Adding client information to the array - FS and MF cases only 
    IF program_ID = "FS" or "MF" then
        If instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") then 
            add_to_array = False    
            msgbox MAXIS_case_number
        Else
            ReDim Preserve expedited_array(5, entry_record)	'This resizes the array based on the number of rows in the Excel File'
            expedited_array(worker_number_const,    entry_record) = worker_number
            expedited_array(case_number_const,      entry_record) = MAXIS_case_number		
            expedited_array(program_ID_const,       entry_record) = program_ID        
            expedited_array(days_pending_const,     entry_record) = days_pending         
            expedited_array(application_date_const, entry_record) = application_date           
            expedited_array(case_status_const,      entry_record) = case_status
            expedited_array(appears_exp_const,      entry_record) = ""

            entry_record = entry_record + 1			'This increments to the next entry in the array'
            stats_counter = stats_counter + 1
            all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*"
        End if 
    End if 
    excel_row = excel_row + 1
Loop

For item = 0 to UBound(expedited_array, 2)
    worker_number       = expedited_array(worker_number_const,    item) 
    MAXIS_case_number   = expedited_array(case_number_const,      item) 
    program_ID          = expedited_array(program_ID_const,       item) 
    days_pending        = expedited_array(days_pending_const,     item) 
    application_date    = expedited_array(application_date_const, item) 
    
    If instr(worker_number, "X127") then 
        expedited_array(case_status_const, item) = "OUT OF COUNTY CASE"
        expedited_array(appears_exp_const, item) = "Not Expedited"
    Else 
        Call navigate_to_MAXIS_screen("STAT", "PROG")
        EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip
        IF priv_check = "PRIV" THEN
            expedited_array(case_status_const, item) = "PRIV CASE"
            expedited_array(appears_exp_const, item) = "Not Expedited"
        End if
        
        EMReadScreen county_code, 4, 21, 21
        If county_code <> "X127" then
            expedited_array(case_status_const, item) = "OUT OF COUNTY CASE"
            expedited_array(appears_exp_const, item) = "Not Expedited"
        End if 
    End if 
    
    If expedited_array(appears_exp_const, item) = "" then 
        MFIP_PENDING = ""		'Setting some variables for the loop
        SNAP_PENDING = ""

        SNAP_status_check = ""
        MFIP_prog_1_check = ""
        MFIP_status_1_check = ""
        MFIP_prog_2_check = ""
        MFIP_status_2_check = ""

        'Reading the status and program
        EMReadScreen SNAP_status_check, 4, 10, 74		'checking the SNAP status
        EMReadScreen MFIP_prog_1_check, 2, 6, 67		'checking for an active MFIP case
        EMReadScreen MFIP_status_1_check, 4, 6, 74
        EMReadScreen MFIP_prog_2_check, 2, 6, 67		'checking for an active MFIP case
        EMReadScreen MFIP_status_2_check, 4, 6, 74

        IF SNAP_status_check = "ACTV" then 
            SNAP_PENDING = FALSE
            expedited_array(case_status_const, item) = "SNAP ACTIVE"
            expedited_array(appears_exp_const, item) = "N/A"
        Else     
            IF SNAP_pending_status "PEND" then 
                SNAP_PENDING = TRUE 
            Else 
                SNAP_PENDING = FALSE 
            End if 
            
            'Logic to determine if MFIP is active
            If MFIP_prog_1_check = "MF" Then
                If MFIP_status_1_check = "ACTV" Then 
                    MFIP_PENDING = FALSE
                Elseif MFIP_status_1_check = "PEND" Then 
                    MFIP_PENDING = TRUE
                Else 
                    MFIP_PENDING = FALSE
                End if 
            ElseIf MFIP_prog_2_check = "MF" Then
                If MFIP_status_2_check = "ACTV" Then
                    MFIP_PENDING = FALSE
                Elseif MFIP_status_2_check = "PEND" Then 
                    MFIP_PENDING = TRUE
                Else
                    MFIP_PENDING = FALSE
                End if 
            End if   
            
            Call HCRE_panel_bypass			'Function to bypass a janky HCRE panel. If the HCRE panel has fields not completed/'reds up' this gets us out of there. 
            
            'If the case note needs to be reviewd for the NOTES - EXPEDITED SCREENING case note, then the
            Call navigate_to_MAXIS_screen("CASE", "NOTE")
            'starting at the 1st case note, checking the headers for the NOTES - EXPEDITED SCREENING text or the NOTES - EXPEDITED DETERMINATION text
            MAXIS_row = 5
            Do
                EMReadScreen case_note_date, 8, MAXIS_row, 6
                If trim(case_note_date) = "" then
                    expedited_array(case_status_const, item) = "Case Notes Do Not Exist"
                    expedited_array(appears_exp_const, item) = "Exp Screening Req"
                    exit do
                Else 
                    EMReadScreen case_note_header, 55, MAXIS_row, 25
                    case_note_header = lcase(trim(case_note_header))
                    IF instr(case_note_header, "appears expedited") or instr(case_note_header, "appears expedit") then
                        expedited_array(case_status_const, item) = "Appears Expedited"
                        expedited_array(appears_exp_const, item) = "Req Exp Processing"
                        exit do
                    Elseif instr(case_note_header, "does not appear") then
                        expedited_array(case_status_const, item) = "Screened, Not EXP"
                        expedited_array(appears_exp_const, item) = "Not Expedited"
                        exit do
                    Else
                        expedited_array(case_status_const, item) = "Screening Not Found"
                        expedited_array(appears_exp_const, item) = "Exp Screening Req"
                    END IF
                END IF
                MAXIS_row = MAXIS_row + 1
            LOOP until cdate(case_note_date) < cdate(application_date)                        'repeats until the case note date is less than the application date
        End if      
    End if 
Next 

Msgbox "Output to Excel staring"

'----------------------------------------------------------------------------------------------------1st page: Req Exp Processing
ObjExcel.ActiveSheet.Name = "Req Exp Processing"

'adding information to the Excel list from PND2
ObjExcel.Cells(1, 1).Value = "Worker #"
ObjExcel.Cells(1, 2).Value = "Case number"
ObjExcel.Cells(1, 3).Value = "Prog ID"
ObjExcel.Cells(1, 4).Value = "Pend Count"
ObjExcel.Cells(1, 5).Value = "APPL date"
objExcel.Columns(5).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
ObjExcel.Cells(1, 6).Value = "Notes"

Excel_row = 2

For item = 0 to UBound(expedited_array, 2)
    If expedited_array(appears_exp_const, item) = "Req Exp Processing" then 
        objExcel.Cells(excel_row, 1).Value = expedited_array(worker_number_const,    item)
        objExcel.Cells(excel_row, 2).Value = expedited_array(case_number_const,      item)
        objExcel.Cells(excel_row, 3).Value = expedited_array(program_ID_const,       item)
        objExcel.Cells(excel_row, 4).Value = expedited_array(days_pending_const,     item)
        objExcel.Cells(excel_row, 5).Value = expedited_array(application_date_const, item)
        objExcel.Cells(excel_row, 6).Value = expedited_array(case_status_const,      item)
        excel_row = excel_row + 1
    End if 
Next 

FOR i = 1 to 6		'formatting the cells
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'----------------------------------------------------------------------------------------------------2nd page: Exp Screening Req
ObjExcel.ActiveSheet.Name = "Exp Screening Req"

'adding information to the Excel list from PND2
ObjExcel.Cells(1, 1).Value = "Worker #"
ObjExcel.Cells(1, 2).Value = "Case number"
ObjExcel.Cells(1, 3).Value = "Prog ID"
ObjExcel.Cells(1, 4).Value = "Pend Count"
ObjExcel.Cells(1, 5).Value = "APPL date"
objExcel.Columns(5).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
ObjExcel.Cells(1, 6).Value = "Notes"

Excel_row = 2

For item = 0 to UBound(expedited_array, 2)
    If expedited_array(appears_exp_const, item) = "Exp Screening Req" then 
        objExcel.Cells(excel_row, 1).Value = expedited_array(worker_number_const,    item)
        objExcel.Cells(excel_row, 2).Value = expedited_array(case_number_const,      item)
        objExcel.Cells(excel_row, 3).Value = expedited_array(program_ID_const,       item)
        objExcel.Cells(excel_row, 4).Value = expedited_array(days_pending_const,     item)
        objExcel.Cells(excel_row, 5).Value = expedited_array(application_date_const, item)
        objExcel.Cells(excel_row, 6).Value = expedited_array(case_status_const,      item)
        excel_row = excel_row + 1
    End if 
Next 

FOR i = 1 to 6		'formatting the cells
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'----------------------------------------------------------------------------------------------------3rd page: Not Expedited
ObjExcel.ActiveSheet.Name = "Not Expedited"

'adding information to the Excel list from PND2
ObjExcel.Cells(1, 1).Value = "Worker #"
ObjExcel.Cells(1, 2).Value = "Case number"
ObjExcel.Cells(1, 3).Value = "Prog ID"
ObjExcel.Cells(1, 4).Value = "Pend Count"
ObjExcel.Cells(1, 5).Value = "APPL date"
objExcel.Columns(5).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
ObjExcel.Cells(1, 6).Value = "Notes"

Excel_row = 2

For item = 0 to UBound(expedited_array, 2)
    If expedited_array(appears_exp_const, item) = "Not Expedited" then 
        objExcel.Cells(excel_row, 1).Value = expedited_array(worker_number_const,    item)
        objExcel.Cells(excel_row, 2).Value = expedited_array(case_number_const,      item)
        objExcel.Cells(excel_row, 3).Value = expedited_array(program_ID_const,       item)
        objExcel.Cells(excel_row, 4).Value = expedited_array(days_pending_const,     item)
        objExcel.Cells(excel_row, 5).Value = expedited_array(application_date_const, item)
        objExcel.Cells(excel_row, 6).Value = expedited_array(case_status_const,      item)
        excel_row = excel_row + 1
    End if 
Next 

FOR i = 1 to 6		'formatting the cells
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'logging usage stats
STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Please review the worksheets for expedited processing needs.")
