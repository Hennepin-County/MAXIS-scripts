'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - EXPEDITED REVIEW.vbs"
start_time = timer
STATS_counter = 1                           'sets the stats counter at one
STATS_manualtime = 60                       'manual run time in seconds
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
CALL changelog_update("06/19/2020", "Updated for phase 2 of Exp SNAP project including emailing Triage group for assignment.", "Ilse Ferris, Hennepin County")
CALL changelog_update("05/19/2020", "Updated to create assignments based on the current phase of the project.", "Ilse Ferris, Hennepin County")
CALL changelog_update("03/31/2020", "Removed email funtionality when report is finished running.", "Ilse Ferris, Hennepin County")
CALL changelog_update("02/24/2020", "Added to ADMIN Main Menu - BZ menu.", "Ilse Ferris, Hennepin County")
CALL changelog_update("02/20/2020", "Final testing version.", "Ilse Ferris, Hennepin County")
CALL changelog_update("02/12/2020", "Added email and auto-save funcationlity.", "Ilse Ferris, Hennepin County")
CALL changelog_update("02/11/2020", "Final testing version complete. Comments added to code.", "Ilse Ferris, Hennepin County")
CALL changelog_update("01/15/2020", "Initial version.", "Ilse Ferris, Hennepin County")
'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function add_pages(exp_status)
    ObjExcel.Worksheets.Add().Name = exp_status
    
    'adding information to the Excel list from PND2
    ObjExcel.Cells(1, 1).Value = "Worker #"
    ObjExcel.Cells(1, 2).Value = "Case number"
    ObjExcel.Cells(1, 3).Value = "Prog ID"
    ObjExcel.Cells(1, 4).Value = "Days Pending"
    ObjExcel.Cells(1, 5).Value = "APPL Date"
    objExcel.Columns(5).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
    ObjExcel.Cells(1, 6).Value = "Interview Date"
    objExcel.Columns(6).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
    ObjExcel.Cells(1, 7).Value = "Notes"
    ObjExcel.Cells(1, 8).Value = "QI Review Notes"
     
    Excel_row = 2
     
    For item = 0 to UBound(expedited_array, 2)
        If expedited_array(appears_exp_const, item) = exp_status then 
            objExcel.Cells(excel_row, 1).Value = expedited_array(worker_number_const,    item)
            objExcel.Cells(excel_row, 2).Value = expedited_array(case_number_const,      item)
            objExcel.Cells(excel_row, 3).Value = expedited_array(program_ID_const,       item)
            objExcel.Cells(excel_row, 4).Value = expedited_array(days_pending_const,     item)
            objExcel.Cells(excel_row, 5).Value = expedited_array(application_date_const, item)
            objExcel.Cells(excel_row, 6).Value = expedited_array(interview_date_const,   item)
            objExcel.Cells(excel_row, 7).Value = expedited_array(case_status_const,      item)
            objExcel.Cells(excel_row, 8).Value = expedited_array(prev_notes_const,       item)
            excel_row = excel_row + 1
        End if 
    Next 
     
    FOR i = 1 to 8		'formatting the cells
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT
END FUNCTION

'THE SCRIPT-----------------------------------------------------------------------------------------------------------
EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr 

report_date = replace(date, "/", "-")   'Changing the format of the date to use as file path selection default
file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & report_date & ".xlsx"

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

'Finding and opening the previous day's file. 
previous_date = dateadd("d", -1, date)
Call change_date_to_soonest_working_day(previous_date)       'finds the most recent previous working day for the fin
file_date = replace(previous_date, "/", "-")   'Changing the format of the date to use as file path selection default
previous_file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & file_date & ".xlsx"

If objExcel = "" Then Call excel_open(previous_file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

For Each objWorkSheet In objWorkbook.Worksheets 'Creating an array of worksheets that are not the intitial report - "Report 1"
    If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "Report 1" then sheet_list = sheet_list & objWorkSheet.Name & ","
Next

sheet_list = trim(sheet_list)  'trims excess spaces of sheet_list
If right(sheet_list, 1) = "," THEN sheet_list = left(sheet_list, len(sheet_list) - 1) 'trimming off last comma
array_of_sheets = split(sheet_list, ",")   'Creating new array

'Establshing array     
DIM master_array()          'Declaring the array
ReDim master_array(2, 0)    'Resizing the array 

'Creating constants to value the array elements
const master_case_number_const = 0  
const master_notes_const       = 1

master_count = 0    'incrementer for the array 

For each excel_sheet in array_of_sheets 
    objExcel.worksheets(excel_sheet).Activate 'Activates the applicable worksheet 
    excel_row = 2
    
    Do 
        master_case_number = ObjExcel.Cells(excel_row, 2).Value 'reading case number
        master_case_number = trim(master_case_number)
        If master_case_number = "" then exit do 
        
        master_notes = ObjExcel.Cells(excel_row, 8).Value       'reading worker entered notes       
    
        If trim(master_notes) <> "" then 
            ReDim Preserve master_array(2, master_record)	'This resizes the array based on if master notes were found or not
            master_array(master_case_number_const,  master_record) = master_case_number
            master_array(master_notes_const,        master_record) = trim(master_notes)
            
            master_record = master_record + 1			'This increments to the next entry in the array'
            STATS_counter = STATS_counter + 1           'stats incrementor 
        End if 
        excel_row = excel_row + 1                       'Excel row incrementor
    LOOP
Next 

'Closing workbook and quiting Excel application
objExcel.ActiveWorkbook.Close                           
objExcel.Application.Quit
objExcel.Quit
 
back_to_self                            'resetting MAXIS back to self before getting started
Call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

'Opening today's list         
Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file
objExcel.worksheets("Report 1").Activate                                 'Activates the initial BOBI report 

'Establishing array
DIM expedited_array()           'Declaring the array
ReDim expedited_array(8, 0)     'Resizing the array 

'Creating constants to value the array elements
const worker_number_const       = 0
const case_number_const	        = 1
const program_ID_const 	        = 2
const days_pending_const        = 3
const application_date_const    = 4
const interview_date_const      = 5
const case_status_const         = 6
const appears_exp_const         = 7
const prev_notes_const          = 8

'Now the script adds all the clients on the excel list into an array
excel_row = 5                   're-establishing the row to start based on when Report 1 starts 
entry_record = 0                'incrementer for the array and count 
all_case_numbers_array = "*"    'setting up string to find duplicate case numbers 
Do   
    'Reading information from the BOBI report in Excel 
    worker_number = objExcel.cells(excel_row, 2).Value
    worker_number = trim(worker_number)
    
    MAXIS_case_number = objExcel.cells(excel_row, 3).Value      
    MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do
    
    program_ID = objExcel.cells(excel_row, 4).Value   
    program_ID = trim(program_ID)

    application_date = objExcel.cells(excel_row, 6).Value
    interview_date   = objExcel.cells(excel_row, 7).Value
    
    days_pending = datediff("D", application_date, date) 
    
    'If the case number is found in the string of case numbers, it's not added again. 
    If instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") then 
        add_to_array = False    
    Else
        'Adding client information to the array
        ReDim Preserve expedited_array(8, entry_record)	'This resizes the array based on the number of cases
        expedited_array(worker_number_const,    entry_record) = worker_number
        expedited_array(case_number_const,      entry_record) = MAXIS_case_number		
        expedited_array(program_ID_const,       entry_record) = program_ID        
        expedited_array(days_pending_const,     entry_record) = days_pending         
        expedited_array(application_date_const, entry_record) = trim(application_date)      
        expedited_array(interview_date_const,   entry_record) = trim(interview_date)    
        expedited_array(case_status_const,      entry_record) = ""              'making space in the array for these variables, but valuing them as "" for now
        expedited_array(appears_exp_const,      entry_record) = ""
        expedited_array(prev_notes_const,       entry_record) = ""
        
        For i = 0 to Ubound(master_array, 2)                                                            'If notes were selected to be added, array is looped thru for matching case number
            If master_array(master_case_number_const, i) = MAXIS_case_number then 
                expedited_array(prev_notes_const, entry_record) = master_array(master_notes_const, i)   'If case number is found, prevoius list notes are added to the array 
                exit for 
            End if      
        Next 

        entry_record = entry_record + 1			'This increments to the next entry in the array
        stats_counter = stats_counter + 1       'Increment for stats counter 
        all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*") 'Adding MAXIS case number to case number string
    End if 
    excel_row = excel_row + 1   
Loop

'Setting up counts for data tracking
screening_count = 0
expedited_count = 0
priv_count = 0
not_exp_count = 0

'Loading of cases is complete. Reviewing the cases in the array. 
For item = 0 to UBound(expedited_array, 2)
    worker_number       = expedited_array(worker_number_const,    item)     're-valuding array variables 
    MAXIS_case_number   = expedited_array(case_number_const,      item) 
    program_ID          = expedited_array(program_ID_const,       item) 
    days_pending        = expedited_array(days_pending_const,     item) 
    application_date    = expedited_array(application_date_const, item) 
    
    If left(worker_number, 4) <> "X127" then                                    'Out of county cases from initial upload
        expedited_array(case_status_const, item) = "OUT OF COUNTY CASE"
        expedited_array(appears_exp_const, item) = "Not Expedited"
        not_exp_count = not_exp_count + 1
    Else 
        Call navigate_to_MAXIS_screen("STAT", "PROG")
        EMReadScreen priv_check, 4, 24, 14 'If it can't get into the case needs to skip - checking in PROD and INQUIRY 
        IF priv_check = "PRIV" then                                             'PRIV cases 
            EmReadscreen priv_worker, 26, 24, 46
            expedited_array(case_status_const, item) = trim(priv_worker)
            expedited_array(appears_exp_const, item) = "Privileged Cases"
            priv_count = priv_count + 1
        else 
            EMReadScreen county_code, 4, 21, 21                                 'Out of county cases from STAT 
            If county_code <> "X127" then
                expedited_array(case_status_const, item) = "OUT OF COUNTY CASE"
                expedited_array(appears_exp_const, item) = "Not Expedited"
                not_exp_count = not_exp_count + 1
            End if 
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
            expedited_array(appears_exp_const, item) = "Not Expedited"
            not_exp_count = not_exp_count + 1
        ElseIF SNAP_status_check = "PEND" then 
            SNAP_PENDING = TRUE 
        ElseIF program_ID = "FS" and SNAP_status_check = "DENY" then 
            expedited_array(case_status_const, item) = "SNAP Application Denied"
            expedited_array(appears_exp_const, item) = "Not Expedited"
            not_exp_count = not_exp_count + 1
        Else    
            'MFIP determination of pending or active 
            SNAP_PENDING = FALSE
            'Logic to determine if MFIP is active
            If MFIP_prog_1_check = "MF" Then
                If MFIP_status_1_check = "ACTV" Then 
                    MFIP_PENDING = FALSE
                    MFIP_ACTIVE = TRUE 
                Elseif MFIP_status_1_check = "PEND" Then 
                    MFIP_PENDING = TRUE
                Else 
                    MFIP_PENDING = FALSE
                End if 
            ElseIf MFIP_prog_2_check = "MF" Then
                If MFIP_status_2_check = "ACTV" Then
                    MFIP_PENDING = FALSE
                    MFIP_ACTIVE = TRUE
                Elseif MFIP_status_2_check = "PEND" Then 
                    MFIP_PENDING = TRUE
                Else
                    MFIP_PENDING = FALSE
                End if
            Elseif MFIP_prog_1_check = "  " or MFIP_prog_2_check = "  " then
                If MFIP_status_1_check = "PEND" or MFIP_status_2_check = "PEND" Then 
                    MFIP_PENDING = TRUE
                Else 
                    MFIP_PENDING = FALSE
                End if 
            End if   
        End if     
        
        IF MFIP_ACTIVE = TRUE and SNAP_PENDING = TRUE then 
            expedited_array(case_status_const, item) = "MFIP ACTIVE"
            expedited_array(appears_exp_const, item) = "Not Expedited"
            not_exp_count = not_exp_count + 1    
        End if  
               
        'Determining if case notes need to be reviewed or not     
        If SNAP_PENDING = True or MFIP_PENDING = True then 
            check_case_note = True
        Else
            check_case_note = False 
        End if 
        
        'handling for cases that do not have a completed HCRE panel - otherwise these get stuck and cannot pass STAT/HCRE 
        PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
        Do
        	EMReadscreen HCRE_panel_check, 4, 2, 50
        	If HCRE_panel_check = "HCRE" then
        		PF10	'exists edit mode in cases where HCRE isn't complete for a member
        		PF3
        	END IF
        Loop until HCRE_panel_check <> "HCRE"
        
        If check_case_note = True then 
            Call navigate_to_MAXIS_screen("CASE", "NOTE")
            'starting at the 1st case note, checking the headers for the NOTES - EXPEDITED SCREENING text or the NOTES - EXPEDITED DETERMINATION text
            MAXIS_row = 5
            Do
                EMReadScreen first_case_note_date, 8, 5, 6 'static reading of the case note date to determine if no case notes acutually exist. 
                If trim(first_case_note_date) = "" then
                    case_note_found = True 
                    expedited_array(case_status_const, item) = "Case Notes Do Not Exist"
                    expedited_array(appears_exp_const, item) = "Exp Screening Req"
                    screening_count = screening_count + 1
                    exit do
                Else
                    EMReadScreen case_note_date, 8, MAXIS_row, 6    'incremented row - reading the case note date
                    EMReadScreen case_note_header, 55, MAXIS_row, 25
                    case_note_header = lcase(trim(case_note_header))
                    
                    If trim(case_note_date) = "" then
                        case_note_found = False             'The end of the case notes has been found
                        exit do 
                    ElseIf instr(case_note_header, "appears expedited") or instr(case_note_header, "appears expedit") then
                        case_note_found = True 
                        expedited_array(case_status_const, item) = "Appears Expedited"
                        expedited_array(appears_exp_const, item) = "Req Exp Processing"
                        expedited_count = expedited_count + 1
                        exit do
                    Elseif instr(case_note_header, "does not appear") or instr(case_note_header, "appears not expedited") then
                        case_note_found = True 
                        expedited_array(case_status_const, item) = "Screened, Not EXP"
                        expedited_array(appears_exp_const, item) = "Not Expedited"
                        not_exp_count = not_exp_count + 1
                        exit do
                    Else
                        case_note_found = False         'defaulting to false if not able to find an expedited care note
                        MAXIS_row = MAXIS_row + 1
                        IF MAXIS_row = 19 then 
                            PF8                         'moving to next case note page if at the end of the page
                            MAXIS_row = 5
                        End if 
                    END IF
                END IF    
            LOOP until cdate(case_note_date) < cdate(application_date)                        'repeats until the case note date is less than the application date    
            If case_note_found = False then 
                expedited_array(case_status_const, item) = "Screening Not Found"
                expedited_array(appears_exp_const, item) = "Exp Screening Req"
                screening_count = screening_count + 1
            End if 
        End if 
    End if 
Next 

'Excel output of cases and information in their applicable categories - PRIV, Req EXP Processing, Exp Screening Required, Not Expedited
Msgbox "Output to Excel Starting."      'warning message to whomever is running the script 

Call Add_pages("Not Expedited")         'calling function to output exp snap statuses to Excel
Call Add_pages("Privileged Cases")
Call Add_pages("Req Exp Processing")
Call Add_pages("Exp Screening Req")

objWorkbook.Save()  'saves existing workbook as same name 
objExcel.Quit

'----------------------------------------------------------------------------------------------------'Pending over 30 days report 
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet
ObjExcel.ActiveSheet.Name = "Pending Over 30 days"

'adding information to the Excel list from PND2
ObjExcel.Cells(1, 1).Value = "Worker #"
ObjExcel.Cells(1, 2).Value = "Case number"
ObjExcel.Cells(1, 3).Value = "Prog ID"
ObjExcel.Cells(1, 4).Value = "Days Pending"
ObjExcel.Cells(1, 5).Value = "APPL Date"
objExcel.Columns(5).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
ObjExcel.Cells(1, 6).Value = "Interview Date"
objExcel.Columns(6).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
ObjExcel.Cells(1, 7).Value = "Notes"
ObjExcel.Cells(1, 8).Value = "QI Review Notes"
 
Excel_row = 2
 
For item = 0 to UBound(expedited_array, 2)
    If expedited_array(case_status_const, item) = "SNAP Application Denied" or expedited_array(case_status_const, item) = "OUT OF COUNTY CASE" then 
        assign_case = False
    elseif expedited_array(case_status_const, item) = "SNAP ACTIVE" and expedited_array(program_ID_const, item) = "FS" then 
        assign_case = False 
    else 
        assign_case = True
    End if 
    
    If assign_case = True then 
        If expedited_array(days_pending_const, item) => 30 then 
            objExcel.Cells(excel_row, 1).Value = expedited_array(worker_number_const,    item)
            objExcel.Cells(excel_row, 2).Value = expedited_array(case_number_const,      item)
            objExcel.Cells(excel_row, 3).Value = expedited_array(program_ID_const,       item)
            objExcel.Cells(excel_row, 4).Value = expedited_array(days_pending_const,     item)
            objExcel.Cells(excel_row, 5).Value = expedited_array(application_date_const, item)
            objExcel.Cells(excel_row, 6).Value = expedited_array(interview_date_const,   item)
            objExcel.Cells(excel_row, 7).Value = expedited_array(case_status_const,      item)
            objExcel.Cells(excel_row, 8).Value = expedited_array(prev_notes_const,       item)
            excel_row = excel_row + 1
        End if 
    End if 
Next 
 
FOR i = 1 to 8		'formatting the cells
    objExcel.Cells(1, i).Font.Bold = True		'bold font'
    objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Saves and closes the most recent Excel workbook
objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\Pending Over 30 Days " & report_date & ".xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

'----------------------------------------------------------------------------------------------------'Expedited Review
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet
ObjExcel.ActiveSheet.Name = "Expedited Review"

'adding information to the Excel list from PND2
ObjExcel.Cells(1, 1).Value = "Worker #"
ObjExcel.Cells(1, 2).Value = "Case number"
ObjExcel.Cells(1, 3).Value = "Prog ID"
ObjExcel.Cells(1, 4).Value = "Days Pending"
ObjExcel.Cells(1, 5).Value = "APPL Date"
objExcel.Columns(5).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
ObjExcel.Cells(1, 6).Value = "Interview Date"
objExcel.Columns(6).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
ObjExcel.Cells(1, 7).Value = "Notes"
ObjExcel.Cells(1, 8).Value = "QI Review Notes"
 
Excel_row = 2
 
For item = 0 to UBound(expedited_array, 2)
    If expedited_array(appears_exp_const, item) = "Exp Screening Req" then 
        assign_case = True 
    ElseIf expedited_array(appears_exp_const, item) = "Req Exp Processing" and expedited_array(interview_date_const, item) <> "" then 
        assign_case = True 
    ElseIf expedited_array(appears_exp_const, item) = "Privileged Cases" then 
        assign_case = True 
    Else 
        assign_case = False 
    End if 
    
    If assign_case = True then
        'only assigning cases that haven't exceeded Day 30 - Those are their own assignments 
        If expedited_array(days_pending_const, item) < 30 then 
            objExcel.Cells(excel_row, 1).Value = expedited_array(worker_number_const,    item)
            objExcel.Cells(excel_row, 2).Value = expedited_array(case_number_const,      item)
            objExcel.Cells(excel_row, 3).Value = expedited_array(program_ID_const,       item)
            objExcel.Cells(excel_row, 4).Value = expedited_array(days_pending_const,     item)
            objExcel.Cells(excel_row, 5).Value = expedited_array(application_date_const, item)
            objExcel.Cells(excel_row, 6).Value = expedited_array(interview_date_const,   item)
            objExcel.Cells(excel_row, 7).Value = expedited_array(case_status_const,      item)
            objExcel.Cells(excel_row, 8).Value = expedited_array(prev_notes_const,       item)
            excel_row = excel_row + 1
        End if 
    End if 
Next 
 
FOR i = 1 to 8		'formatting the cells
    objExcel.Cells(1, i).Font.Bold = True		'bold font'
    objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Saves and closes the most recent Excel workbook
objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\Expedited Review " & report_date & ".xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

'------------------------------------------------------------------------------------------------------------------------------------------------------------'Needs Interview Cases
'----------------------------------------------------------------------------------------------------NOT EXPEDITED 
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet
ObjExcel.ActiveSheet.Name = "Not Expedited"

'adding information to the Excel list from PND2
ObjExcel.Cells(1, 1).Value = "Worker #"
ObjExcel.Cells(1, 2).Value = "Case number"
ObjExcel.Cells(1, 3).Value = "Prog ID"
ObjExcel.Cells(1, 4).Value = "Days Pending"
ObjExcel.Cells(1, 5).Value = "APPL Date"
objExcel.Columns(5).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
ObjExcel.Cells(1, 6).Value = "Interview Date"
objExcel.Columns(6).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
ObjExcel.Cells(1, 7).Value = "Notes"
 
Excel_row = 2
 
For item = 0 to UBound(expedited_array, 2)
    If expedited_array(appears_exp_const, item) = "Not Expedited" and expedited_array(interview_date_const, item) = "" then 
        If expedited_array(case_status_const, item) = "SNAP Application Denied" or expedited_array(case_status_const, item) = "OUT OF COUNTY CASE" then 
            assign_case = False
        elseif expedited_array(case_status_const, item) = "SNAP ACTIVE" and expedited_array(program_ID_const, item) = "FS" then 
            assign_case = False 
        else 
            assign_case = True
        End if 
    Else 
        assign_case = False 
    End if 
    
    If assign_case = True then 
        'only assigning cases that haven't exceeded Day 30 - Those are their own assignments 
        If expedited_array(days_pending_const, item) < 30 then 
            objExcel.Cells(excel_row, 1).Value = expedited_array(worker_number_const,    item)
            objExcel.Cells(excel_row, 2).Value = expedited_array(case_number_const,      item)
            objExcel.Cells(excel_row, 3).Value = expedited_array(program_ID_const,       item)
            objExcel.Cells(excel_row, 4).Value = expedited_array(days_pending_const,     item)
            objExcel.Cells(excel_row, 5).Value = expedited_array(application_date_const, item)
            objExcel.Cells(excel_row, 6).Value = expedited_array(interview_date_const,   item)
            objExcel.Cells(excel_row, 7).Value = expedited_array(case_status_const,      item)
            excel_row = excel_row + 1
        End if 
    End if 
Next 
 
FOR i = 1 to 7		'formatting the cells
    objExcel.Cells(1, i).Font.Bold = True		'bold font'
    objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'----------------------------------------------------------------------------------------------------Appears Expedited 
ObjExcel.Worksheets.Add().Name = "Appears Expedited"

'adding information to the Excel list from PND2
ObjExcel.Cells(1, 1).Value = "Worker #"
ObjExcel.Cells(1, 2).Value = "Case number"
ObjExcel.Cells(1, 3).Value = "Prog ID"
ObjExcel.Cells(1, 4).Value = "Days Pending"
ObjExcel.Cells(1, 5).Value = "APPL Date"
objExcel.Columns(5).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
ObjExcel.Cells(1, 6).Value = "Interview Date"
objExcel.Columns(6).NumberFormat = "mm/dd/yy"					'formats the date column as MM/DD/YY
ObjExcel.Cells(1, 7).Value = "Notes"
 
Excel_row = 2
 
For item = 0 to UBound(expedited_array, 2)
    If expedited_array(appears_exp_const, item) = "Exp Screening Req" and expedited_array(interview_date_const, item) = "" then 
        assign_case = True 
    ElseIf expedited_array(appears_exp_const, item) = "Req Exp Processing" and expedited_array(interview_date_const, item) = "" then 
        assign_case = True 
    Else 
        assign_case = False 
    End if 
    
    If assign_case = True then 
        'only assigning cases that haven't exceeded Day 30 - Those are their own assignments 
        If expedited_array(days_pending_const, item) < 30 then 
            objExcel.Cells(excel_row, 1).Value = expedited_array(worker_number_const,    item)
            objExcel.Cells(excel_row, 2).Value = expedited_array(case_number_const,      item)
            objExcel.Cells(excel_row, 3).Value = expedited_array(program_ID_const,       item)
            objExcel.Cells(excel_row, 4).Value = expedited_array(days_pending_const,     item)
            objExcel.Cells(excel_row, 5).Value = expedited_array(application_date_const, item)
            objExcel.Cells(excel_row, 6).Value = expedited_array(interview_date_const,   item)
            objExcel.Cells(excel_row, 7).Value = expedited_array(case_status_const,      item)
            excel_row = excel_row + 1
        End if 
    End if 
Next 
 
FOR i = 1 to 7		'formatting the cells
    objExcel.Cells(1, i).Font.Bold = True		'bold font'
    objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Saves and closes the most recent Excel workbook
objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP SNAP " & report_date & ".xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

stats_report = "Screening Count: " & screening_count & vbcr & _
"Expedited Processing Count: " & expedited_count & vbcr & _
"PRIV Case Count: " & priv_count & vbcr & _
"Not Expedited Count: " & not_exp_count

'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
Call create_outlook_email("HSPH.EWS.Triagers@hennepin.us;Adonna.Swift@hennepin.us", "Laurie.Hennen@hennepin.us", "EXP SNAP Report without Interviews is Ready. EOM.", "", "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP SNAP " & report_date & ".xlsx", True)
Call create_outlook_email("HSPH.EWS.Unit.Frey@hennepin.us", "Ilse.Ferris@hennepin.us", "Today's EXP SNAP reports are ready.", "Path to folder - T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project", "", True)
Call create_outlook_email("Ilse.Ferris@hennepin.us;Laurie.Hennen@hennepin.us","","Expedited SNAP Daily statistics for " & date, stats_report, "", True)

'----------------------------------------------------------------------------------------------------Moves yesterday's files to the archive folder for the specific month
array_of_archive_assigments = array("Expedited Review ","Pending Over 30 Days ", "EXP SNAP ", "")

previous_date = dateadd("d", -1, date)
Call change_date_to_soonest_working_day(previous_date)       'finds the most recent previous working day for the fin
file_date = replace(previous_date, "/", "-")   'Changing the format of the date to use as file path selection default
archive_folder = right("0" & DatePart("m", file_date), 2) & "-" & DatePart("yyyy", file_date)

archive_files = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & archive_folder

Set objFSO = CreateObject("Scripting.FileSystemObject")
For each assignment in array_of_archive_assigments
    file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & assignment & file_date & ".xlsx"
    objFSO.MoveFile file_selection_path , archive_files & "\" & assignment & file_date & ".xlsx"    'moving each file to the archive file
Next 

'logging usage stats
STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success, the expedited SNAP run is complete. The workbook has been saved.")