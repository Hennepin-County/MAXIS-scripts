'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - EXPEDITED REVIEW.vbs"
start_time = timer
STATS_counter = 1                           'sets the stats counter at one
STATS_manualtime = 100                       'manual run time in seconds
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
CALL changelog_update("06/03/2021", "Updated T drive file path to more stable LOBROOT path.", "Ilse Ferris, Hennepin County")
CALL changelog_update("04/09/2021", "Removed FAD Assignments and associated actions for the FAD assignments.", "Ilse Ferris, Hennepin County")
CALL changelog_update("03/02/2021", "Added Debrice Jackson to also receive emails for YET's assignments.", "Ilse Ferris, Hennepin County")
CALL changelog_update("12/30/2020", "Updated non-expedited count code for more accurate data.", "Ilse Ferris, Hennepin County")
CALL changelog_update("11/29/2020", "Updated to include Phase 1 of ES Expedited SNAP Support with DWP group.", "Ilse Ferris, Hennepin County")
CALL changelog_update("11/17/2020", "Additional testing for 1800 cases.", "Ilse Ferris, Hennepin County")
CALL changelog_update("11/14/2020", "Added specified report for 1800 baskets. WFM will not get 1800 cases for FAD assignment.", "Ilse Ferris, Hennepin County")
CALL changelog_update("11/04/2020", "Final testing complete for additional data input and output functionality.", "Ilse Ferris, Hennepin County")
CALL changelog_update("10/12/2020", "Adding case review tracking elements capture and output into script.", "Ilse Ferris, Hennepin County")
CALL changelog_update("09/30/2020", "YET workbooks have now been added to the achive transfer process at the end of the script.", "Ilse Ferris, Hennepin County")
CALL changelog_update("09/21/2020", "Added specified report for the YET team based on basket number X127FA5. WFM will not get FA5 cases for FAD assignment.", "Ilse Ferris, Hennepin County")
CALL changelog_update("08/20/2020", "Removed all other emails from assignment email besides WFM.", "Ilse Ferris, Hennepin County")
CALL changelog_update("08/18/2020", "Added WFM and coverage worker email to assignment email at end of script run.", "Ilse Ferris, Hennepin County")
CALL changelog_update("08/17/2020", "Updated file name of QI assignment from Expedited Review to QI Expedited Review.", "Ilse Ferris, Hennepin County")
CALL changelog_update("07/16/2020", "Full phase 2 updates complete including pulling in notes from previous working day's assignments.", "Ilse Ferris, Hennepin County")
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
            If expedited_array(appears_exp_const, item) = "Not Expedited" then not_exp_count = not_exp_count + 1
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

'Setting up counts for data tracking
screening_count = 0
expedited_count = 0
priv_count = 0
not_exp_count = 0   'incrementor built into the function for not expedited only 

'----------------------------------------------------------------------------------------------------Gathering previous working days' assignment notes
'Establshing array     
DIM master_notes_array()          'Declaring the array
ReDim master_notes_array(11, 0)    'Resizing the array 
    
'Creating constants to value the array elements
const master_case_number_const  = 0  
const master_note_const         = 1
const PN1_const                 = 2
const PN2_const                 = 3
const PN3_const                 = 4
const PN4_const                 = 5
const PN5_const                 = 6
const PN6_const                 = 7
const PN7_const                 = 8
const PN8_const                 = 9
const PN9_const                 = 10 
const PN10_const                = 11

master_note_record = 0    'incrementer for the array     
    
previous_date = dateadd("d", -1, date)
Call change_date_to_soonest_working_day(previous_date)       'finds the most recent previous working day for the file names
file_date = replace(previous_date, "/", "-")   'Changing the format of the date to use as file path selection default
file_selection_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\QI Expedited Review " & file_date & ".xlsx" 'single assignment file


If objExcel = "" Then Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

excel_row = 2
Do 
    case_number = ObjExcel.Cells(excel_row, 2).Value 'reading case number
    case_number = trim(case_number)
    If case_number = "" then exit do 
    
    master_note = ObjExcel.Cells(excel_row,  8).Value       'reading worker entered notes       
    PN1         = ObjExcel.Cells(excel_row,  9).Value
    PN2         = ObjExcel.Cells(excel_row, 10).Value
    PN3         = ObjExcel.Cells(excel_row, 11).Value
    PN4         = ObjExcel.Cells(excel_row, 12).Value
    PN5         = ObjExcel.Cells(excel_row, 13).Value
    PN6         = ObjExcel.Cells(excel_row, 14).Value
    PN7         = ObjExcel.Cells(excel_row, 15).Value
    PN8         = ObjExcel.Cells(excel_row, 16).Value
    PN9         = ObjExcel.Cells(excel_row, 17).Value
    PN10        = ObjExcel.Cells(excel_row, 18).Value 

    ReDim Preserve master_notes_array(11,  master_note_record)	'This resizes the array based on if master notes were found or not
    master_notes_array(master_case_number_const, master_note_record) = case_number
    master_notes_array(master_note_const, master_note_record) = trim(master_note)
    master_notes_array(PN1_const, master_note_record) =  trim(PN1)
    master_notes_array(PN2_const, master_note_record) =  trim(PN2)
    master_notes_array(PN3_const, master_note_record) =  trim(PN3)
    master_notes_array(PN4_const, master_note_record) =  trim(PN4)
    master_notes_array(PN5_const, master_note_record) =  trim(PN5)
    master_notes_array(PN6_const, master_note_record) =  trim(PN6)
    master_notes_array(PN7_const, master_note_record) =  trim(PN7)
    master_notes_array(PN8_const, master_note_record) =  trim(PN8)
    master_notes_array(PN9_const, master_note_record) =  trim(PN9)
    master_notes_array(PN10_const, master_note_record) = trim(PN10) 
    
    master_note_record = master_note_record + 1			'This increments to the next entry in the array'
    STATS_counter = STATS_counter + 1           'stats incrementor 
    
    excel_row = excel_row + 1                       'Excel row incrementor
LOOP
    
'Closing workbook and quiting Excel application
objExcel.ActiveWorkbook.Close                           
objExcel.Application.Quit
objExcel.Quit
objExcel = ""
'Next     

'----------------------------------------------------------------------------------------------------Finding and opening the previous day master file (non-assignment file)
previous_date = dateadd("d", -1, date)
Call change_date_to_soonest_working_day(previous_date)       'finds the most recent previous working day for the fin
file_date = replace(previous_date, "/", "-")   'Changing the format of the date to use as file path selection default
previous_file_selection_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & file_date & ".xlsx"

If objExcel = "" Then Call excel_open(previous_file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

For Each objWorkSheet In objWorkbook.Worksheets 'Creating an array of worksheets that are not the intitial report - "Report 1"
    If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "Report 1" then sheet_list = sheet_list & objWorkSheet.Name & ","
Next
    
sheet_list = trim(sheet_list)  'trims excess spaces of sheet_list
If right(sheet_list, 1) = "," THEN sheet_list = left(sheet_list, len(sheet_list) - 1) 'trimming off last comma
array_of_sheets = split(sheet_list, ",")   'Creating new array
    
For each excel_sheet in array_of_sheets
    objExcel.worksheets(excel_sheet).Activate 'Activates the applicable worksheet 
    excel_row = 2
    
    Do 
        MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value  'reading case number
        MAXIS_case_number = trim(MAXIS_case_number)
        If MAXIS_case_number = "" then exit do 
            
        For i = 0 to Ubound(master_notes_array, 2)                                                            'If notes were selected to be added, array is looped thru for matching case number
            master_CN = master_notes_array(master_case_number_const, i)
            If master_CN = MAXIS_case_number then 
                ObjExcel.Cells(excel_row,  8).Value = master_notes_array(master_note_const, i)   'If case number is found, previous list notes are added to the array 
                ObjExcel.Cells(excel_row,  9).Value = master_notes_array(PN1_const, i)
                ObjExcel.Cells(excel_row, 10).Value = master_notes_array(PN2_const, i)
                ObjExcel.Cells(excel_row, 11).Value = master_notes_array(PN3_const, i)
                ObjExcel.Cells(excel_row, 12).Value = master_notes_array(PN4_const, i)
                ObjExcel.Cells(excel_row, 13).Value = master_notes_array(PN5_const, i)
                ObjExcel.Cells(excel_row, 14).Value = master_notes_array(PN6_const, i)
                ObjExcel.Cells(excel_row, 15).Value = master_notes_array(PN7_const, i)
                ObjExcel.Cells(excel_row, 16).Value = master_notes_array(PN8_const, i)
                ObjExcel.Cells(excel_row, 17).Value = master_notes_array(PN9_const, i)
                ObjExcel.Cells(excel_row, 18).Value = master_notes_array(PN10_const,i)        
                exit for 
            End if      
        Next 
        excel_row = excel_row + 1                       'Excel row incrementor
    LOOP
Next 

objWorkbook.Save()  'saves the previous days' notes and closes Excel. 
objExcel.ActiveWorkbook.Close                           
objExcel.Application.Quit
objExcel.Quit
objExcel = ""

''----------------------------------------------------------------------------------------------------The current day's assignment 
report_date = replace(date, "/", "-")   'Changing the format of the date to use as file path selection default
file_selection_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & report_date & ".xlsx"

BeginDialog Dialog1, 0, 0, 481, 90, "ADMIN - EXPEDITED REVIEW"
  ButtonGroup ButtonPressed
    PushButton 420, 40, 50, 15, "Browse...", select_a_file_button
    OkButton 365, 60, 50, 15
    CancelButton 420, 60, 50, 15
  EditBox 15, 40, 400, 15, file_selection_path
  Text 15, 20, 455, 15, "This script should be used to review a BOBI list of pending SNAP and/or MFIP cases to ensure expedited screening and determinations are being made to ensure expedited timeliness rules are being followed."
  Text 15, 65, 335, 10, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 465, 75, "Using this script:"
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

'Opening today's list         
Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file
objExcel.worksheets("Report 1").Activate                                 'Activates the initial BOBI report 

'Establishing array
DIM expedited_array()           'Declaring the array
ReDim expedited_array(18, 0)     'Resizing the array 

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
const prev_notes1_const         = 9
const prev_notes2_const         = 10
const prev_notes3_const         = 11
const prev_notes4_const         = 12
const prev_notes5_const         = 13
const prev_notes6_const         = 14
const prev_notes7_const         = 15
const prev_notes8_const         = 16
const prev_notes9_const         = 17
const prev_notes10_const        = 18

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
        ReDim Preserve expedited_array(18, entry_record)	'This resizes the array based on the number of cases
        expedited_array(worker_number_const,    entry_record) = worker_number
        expedited_array(case_number_const,      entry_record) = MAXIS_case_number		
        expedited_array(program_ID_const,       entry_record) = program_ID        
        expedited_array(days_pending_const,     entry_record) = days_pending         
        expedited_array(application_date_const, entry_record) = trim(application_date)      
        expedited_array(interview_date_const,   entry_record) = trim(interview_date)    
        expedited_array(case_status_const,      entry_record) = ""              'making space in the array for these variables, but valuing them as "" for now
        expedited_array(appears_exp_const,      entry_record) = ""
        expedited_array(prev_notes_const,       entry_record) = ""
        expedited_array(prev_notes1_const,      entry_record) = ""
        expedited_array(prev_notes2_const,      entry_record) = ""
        expedited_array(prev_notes3_const,      entry_record) = ""
        expedited_array(prev_notes4_const,      entry_record) = ""
        expedited_array(prev_notes5_const,      entry_record) = ""
        expedited_array(prev_notes6_const,      entry_record) = ""
        expedited_array(prev_notes7_const,      entry_record) = ""
        expedited_array(prev_notes8_const,      entry_record) = ""
        expedited_array(prev_notes9_const,      entry_record) = ""
        expedited_array(prev_notes10_const,     entry_record) = ""
            
        For i = 0 to Ubound(master_notes_array, 2)                                                            'If notes were selected to be added, array is looped thru for matching case number
            If master_notes_array(master_case_number_const, i) = MAXIS_case_number then 
                expedited_array(prev_notes_const,   entry_record) = master_notes_array(master_note_const, i)   'If case number is found, prevoius list notes are added to the array 
                expedited_array(prev_notes1_const,  entry_record) = master_notes_array(PN1_const,  i)
                expedited_array(prev_notes2_const,  entry_record) = master_notes_array(PN2_const,  i)
                expedited_array(prev_notes3_const,  entry_record) = master_notes_array(PN3_const,  i)
                expedited_array(prev_notes4_const,  entry_record) = master_notes_array(PN4_const,  i)
                expedited_array(prev_notes5_const,  entry_record) = master_notes_array(PN5_const,  i)
                expedited_array(prev_notes6_const,  entry_record) = master_notes_array(PN6_const,  i)
                expedited_array(prev_notes7_const,  entry_record) = master_notes_array(PN7_const,  i)
                expedited_array(prev_notes8_const,  entry_record) = master_notes_array(PN8_const,  i)
                expedited_array(prev_notes9_const,  entry_record) = master_notes_array(PN9_const,  i)
                expedited_array(prev_notes10_const, entry_record) = master_notes_array(PN10_const, i)
                exit for 
            End if      
        Next 

        entry_record = entry_record + 1			'This increments to the next entry in the array
        stats_counter = stats_counter + 1       'Increment for stats counter 
        all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*") 'Adding MAXIS case number to case number string
    End if 
    excel_row = excel_row + 1   
Loop

back_to_self                            'resetting MAXIS back to self before getting started
Call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

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
        ElseIF SNAP_status_check = "PEND" then 
            SNAP_PENDING = TRUE 
        ElseIF program_ID = "FS" and SNAP_status_check = "DENY" then 
            expedited_array(case_status_const, item) = "SNAP Application Denied"
            expedited_array(appears_exp_const, item) = "Not Expedited"
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
objExcel.ActiveWorkbook.SaveAs t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\Pending Over 30 Days " & report_date & ".xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

'----------------------------------------------------------------------------------------------------'QI Expedited Review
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet
ObjExcel.ActiveSheet.Name = "QI Expedited Review"

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

ObjExcel.Cells(1,  9).Value = "Rewiewed"
ObjExcel.Cells(1, 10).Value = "Approved"
ObjExcel.Cells(1, 11).Value = "Appear EXP, no ID - Correct"
ObjExcel.Cells(1, 12).Value = "Appear EXP, ID was available - Incorrect"
ObjExcel.Cells(1, 13).Value = "Processed correctly by HSR"
ObjExcel.Cells(1, 14).Value = "No CAF on file"
ObjExcel.Cells(1, 15).Value = "Verifications should have been postponed/Case app'd"
ObjExcel.Cells(1, 16).Value = "MAXIS was updated incorrectly"
ObjExcel.Cells(1, 17).Value = "Insufficient CASE/NOTES"
ObjExcel.Cells(1, 18).Value = "Save case number for team review?" 

Excel_row = 2
 
For item = 0 to UBound(expedited_array, 2)
    If expedited_array(appears_exp_const, item) = "Exp Screening Req" and expedited_array(interview_date_const, item) <> "" then 
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
            objExcel.Cells(excel_row,  1).Value = expedited_array(worker_number_const,     item)   'COL A
            objExcel.Cells(excel_row,  2).Value = expedited_array(case_number_const,       item)   'COL B
            objExcel.Cells(excel_row,  3).Value = expedited_array(program_ID_const,        item)   'COL C
            objExcel.Cells(excel_row,  4).Value = expedited_array(days_pending_const,      item)   'COL D
            objExcel.Cells(excel_row,  5).Value = expedited_array(application_date_const,  item)   'COL E
            objExcel.Cells(excel_row,  6).Value = expedited_array(interview_date_const,    item)   'COL F
            objExcel.Cells(excel_row,  7).Value = expedited_array(case_status_const,       item)   'COL G
            objExcel.Cells(excel_row,  8).Value = expedited_array(prev_notes_const,        item)   'COL H
            objExcel.Cells(excel_row,  9).Value = expedited_array(prev_notes1_const,       item)   'COL I
            objExcel.Cells(excel_row, 10).Value = expedited_array(prev_notes2_const,       item)   'COL J
            objExcel.Cells(excel_row, 11).Value = expedited_array(prev_notes3_const,       item)   'COL K
            objExcel.Cells(excel_row, 12).Value = expedited_array(prev_notes4_const,       item)   'COL L
            objExcel.Cells(excel_row, 13).Value = expedited_array(prev_notes5_const,       item)   'COL M
            objExcel.Cells(excel_row, 14).Value = expedited_array(prev_notes6_const,       item)   'COL N
            objExcel.Cells(excel_row, 15).Value = expedited_array(prev_notes7_const,       item)   'COL O
            objExcel.Cells(excel_row, 16).Value = expedited_array(prev_notes8_const,       item)   'COL P
            objExcel.Cells(excel_row, 17).Value = expedited_array(prev_notes9_const,       item)   'COL Q
            objExcel.Cells(excel_row, 18).Value = expedited_array(prev_notes10_const,      item)   'COL R
            excel_row = excel_row + 1
        End if 
    End if 
Next 

Set objWorkSheet = objExcel.ActiveWorkbook.Worksheets("QI Expedited Review")    'Creating a connection to the active worksheet 
'Creates droplist with Data Vaildation within the Excel range 
range_variable = "I2:R" & excel_row - 1 & ""    'range to drop lists. Columns are static. excel_row -1 used to determine last row. 

'----------------------------------------------------------------------------------------------------Validation.add method in Excel
'Syntax: expression.Add (Type, AlertStyle, Operator, Formula1, Formula2)
'''' Type: Required: Options are the following:
           'Name	               Value	Description
           'xlValidateCustom	     7	    Data is validated using an arbitrary formula.
           'xlValidateDate	         4	    Date values.
           'xlValidateDecimal	     2	    Numeric values.
           'xlValidateInputOnly      0	    Validate only when user changes the value.
           'xlValidateList	         3	    Value must be present in a specified list.
           'xlValidateTextLength	 6	    Length of text.
           'xlValidateTime	         5	    Time values.
           'xlValidateWholeNumber	 1	    Whole numeric values.
'''' AlertStyle: Optional. Options are the following:
            'Name	                Value	Description
            'xlValidAlertInformation  3	    Information icon.
            'xlValidAlertStop	      1	    Stop icon.
            'xlValidAlertWarning	  2	     Warning icon.
'''' Operator: Optional. Options are the following:
            'Name	                Value	Description
            'xlBetween	             1	     Between. Can be used only if two formulas are provided.
            'xlEqual	             3	     Equal.
            'xlGreater	             5       Greater than.
            'xlGreaterEqual	         7	     Greater than or equal to.
            'xlLess	                 6	     Less than.
            'xlLessEqual	         8	     Less than or equal to.
            'xlNotBetween	         2	     Not between. Can be used only if two formulas are provided.
            'xlNotEqual	             4	     Not equal.
'''' Formula1: Optional. The first part of the data validation equation. Value must not exceed 255 characters.
'''' Formula2: Optional. The second part of the data validation equation when Operator is xlBetween or xlNotBetween (otherwise, this argument is ignored).

'REMARKS - The Add method requires different arguments, depending on the validation type, as shown in the following table.
'Validation type	   Arguments
'xlValidateCustom	   Formula1 is required, Formula2 is ignored. Formula1 must contain an expression that evaluates to True when data entry is valid and False when data entry is invalid.
'xlInputOnly	       AlertStyle, Formula1, or Formula2 are used.
'xlValidateList	       Formula1 is required, Formula2 is ignored. Formula1 must contain either a comma-delimited list of values or a worksheet reference to this list.
'xlValidateWholeNumber One of either Formula1 or Formula2 must be specified, or both may be specified.
'xlValidateDate        One of either Formula1 or Formula2 must be specified, or both may be specified.
'xlValidateDecimal     One of either Formula1 or Formula2 must be specified, or both may be specified.
'xlValidateTextLength  One of either Formula1 or Formula2 must be specified, or both may be specified.
'xlValidateTime        One of either Formula1 or Formula2 must be specified, or both may be specified.

'With/End With: Executes a series of statements that repeatedly refer to a single object or structure so that the statements can use a simplified syntax when accessing members of the object or structure. 
'When using a structure, you can only read the values of members or invoke methods, and you get an error if you try to assign values to members of a structure used in a With...End With statement.

With objWorksheet.Range(range_variable).Validation  'ObjectExpression - Required. An expression that evaluates to an object. The expression may be arbitrarily complex and is evaluated only once. The expression can evaluate to any data type, including elementary types.
        .Add 3, 1, 1, "Yes,No,X"                    'expression.Add (Type, AlertStyle, Operator, Formula1, Formula2) ---values used vs. names                    
        .IgnoreBlank = True                         'True if blank values are permitted by the range data validation. Read/write Boolean.      
        .InCellDropdown = True                      'True if data validation displays a drop-down list that contains acceptable values. Read/write Boolean.  
        .InputTitle = ""                            'Returns or sets the title of the data-validation input dialog box. Read/write String. Limited to 32 characters.
        .ErrorTitle = ""                            'Returns or sets the title of the data-validation error dialog box. Read/write String.
        .InputMessage = ""                          'Returns or sets the data validation input message. Read/write String. 
        .ErrorMessage = ""                          'Returns or sets the data validation error message. Read/write String.  
        .ShowInput = True                           'True if the data validation input message will be displayed whenever the user selects a cell in the data validation range. Read/write Boolean.
        .ShowError = True                           'Draws tracer arrows through the precedents tree to the cell that's the source of the error, and returns the range that contains that cell.
 end With                                           
 
FOR i = 1 to 18		'formatting the cells
    objExcel.Cells(1, i).Font.Bold = True		'bold font'
    objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Saves and closes the most recent Excel workbook
objExcel.ActiveWorkbook.SaveAs t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\QI Expedited Review " & report_date & ".xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

'----------------------------------------------------------------------------------------------------Appears Expedited for YET Team only X127FA5
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet
ObjExcel.ActiveSheet.Name = "Appears Expedited X127FA5"

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
    Elseif expedited_array(case_status_const, item) = "Case Notes Do Not Exist" then 
        assign_case = True 
    Else 
        assign_case = False 
    End if 
    
    If assign_case = True then 
        'only assigning cases that haven't exceeded Day 30 - Those are their own assignments 
        If expedited_array(days_pending_const, item) < 30 then 
            'Assigning only YET pending cases to YET 
            If expedited_array(worker_number_const, item) = "X127FA5" then 
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
    End if 
Next 
 
FOR i = 1 to 7		'formatting the cells
    objExcel.Cells(1, i).Font.Bold = True		'bold font'
    objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Saves and closes the most recent Excel workbook
objExcel.ActiveWorkbook.SaveAs t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP SNAP X127FA5 " & report_date & ".xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

'----------------------------------------------------------------------------------------------------Appears Expedited for 1800 Team: X127EF8 and X127EF9
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet
ObjExcel.ActiveSheet.Name = "Appears Expedited 1800"

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
    Elseif expedited_array(case_status_const, item) = "Case Notes Do Not Exist" then 
        assign_case = True 
    Else
        assign_case = False 
    End if 
    
    If assign_case = True then 
        'only assigning cases that haven't exceeded Day 30 - Those are their own assignments 
        If expedited_array(days_pending_const, item) < 30 then 
            'Assigning only 1800 baskets pending cases to 1800 team  
            If expedited_array(worker_number_const, item) = "X127EF8" or expedited_array(worker_number_const, item) = "X127EF9" then 
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
    End if 
Next 
 
FOR i = 1 to 7		'formatting the cells
    objExcel.Cells(1, i).Font.Bold = True		'bold font'
    objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Saves and closes the most recent Excel workbook
objExcel.ActiveWorkbook.SaveAs t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP SNAP 1800 " & report_date & ".xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

'----------------------------------------------------------------------------------------------------DWP Appears Expedited & EXP Screening & Secondary w/o Interview
'----------------------------------------------------------------------------------------------------Secondary Assignment 
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet
ObjExcel.ActiveSheet.Name = "Secondary Assignment"

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
    If expedited_array(appears_exp_const, item) = "Req Exp Processing" and expedited_array(interview_date_const, item) = "" then 
        assign_case = True  
    Else 
        assign_case = False 
    End if 
    
    'Excluding YET (X127FA5) and 1800 (X127EF8 and X127EF9) baskets
    If expedited_array(worker_number_const, item) = "X127FA5" or expedited_array(worker_number_const, item) = "X127EF8" or expedited_array(worker_number_const, item) = "X127EF9" then assign_case = False 
    
    If assign_case = True then 
        'only assigning cases that haven't exceeded Day 30 - Those are their own assignments 
        If expedited_array(days_pending_const, item) < 30 then 
            If expedited_array(days_pending_const, item) => 8 then 
                'Assigning only cases pending 8-29 days 
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
    End if 
Next 
 
FOR i = 1 to 7		'formatting the cells
    objExcel.Cells(1, i).Font.Bold = True		'bold font'
    objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'----------------------------------------------------------------------------------------------------Appears Expedited & Requires Interview
'----------------------------------------------------------------------------------------------------Appears Expedited Days 1-7
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
    If expedited_array(appears_exp_const, item) = "Req Exp Processing" and expedited_array(interview_date_const, item) = "" then 
        assign_case = True 
    Elseif expedited_array(case_status_const, item) = "Case Notes Do Not Exist" then 
        assign_case = True 
    Else 
        assign_case = False 
    End if 
    
    'Excluding YET (X127FA5) and 1800 (X127EF8 and X127EF9) baskets
    If expedited_array(worker_number_const, item) = "X127FA5" or expedited_array(worker_number_const, item) = "X127EF8" or expedited_array(worker_number_const, item) = "X127EF9" then assign_case = False 
    
    If assign_case = True then 
        'only assigning cases that haven't exceeded Day 30 - Those are their own assignments 
        If expedited_array(days_pending_const, item) < 30 then 
            If expedited_array(days_pending_const, item) =< 7 then 
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
    End if 
Next 
 
FOR i = 1 to 7		'formatting the cells
    objExcel.Cells(1, i).Font.Bold = True		'bold font'
    objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'----------------------------------------------------------------------------------------------------Expedited Screening
ObjExcel.Worksheets.Add().Name = "Screening Not Found"

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
    Else 
        assign_case = False 
    End if 
    
    'Excluding YET (X127FA5) and 1800 (X127EF8 and X127EF9) baskets
    If expedited_array(worker_number_const, item) = "X127FA5" or expedited_array(worker_number_const, item) = "X127EF8" or expedited_array(worker_number_const, item) = "X127EF9" then assign_case = False 
    
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
objExcel.ActiveWorkbook.SaveAs t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP SNAP DWP " & report_date & ".xlsx"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit
objExcel.Quit

stats_report = "Screening Count: " & screening_count & vbcr & _
"Expedited Processing Count: " & expedited_count & vbcr & _
"PRIV Case Count: " & priv_count & vbcr & _
"Not Expedited Count: " & not_exp_count

'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
Call create_outlook_email("Maslah.Jama@hennepin.us; Debrice.Jackson@hennepin.us","Laurie.Hennen@hennepin.us", "EXP SNAP Report for YET without Interviews is Ready. EOM.", "", "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP SNAP X127FA5 " & report_date & ".xlsx", True)
Call create_outlook_email("Carlotta.Madison@hennepin.us", "Laurie.Hennen@hennepin.us", "EXP SNAP Report for 1800 without Interviews is Ready. EOM.", "", "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP SNAP 1800 " & report_date & ".xlsx", True)
Call create_outlook_email("HSPH.EWS.Unit.Frey@hennepin.us", "", "Today's EXP SNAP reports are ready.", "Path to folder - T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project", "", True)
Call create_outlook_email("Mohamed.Ahmed@hennepin.us; Dawn.Welch@hennepin.us", "Ilse.Ferris@hennepin.us;Laurie.Hennen@hennepin.us", "Today's EXP SNAP primary and secondary assignments are ready.", "See attachment.", "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP SNAP DWP " & report_date & ".xlsx", True)
Call create_outlook_email("Ilse.Ferris@hennepin.us;Laurie.Hennen@hennepin.us","","Expedited SNAP Daily statistics for " & date, stats_report, "", True)

'----------------------------------------------------------------------------------------------------Moves yesterday's files to the archive folder for the specific month

array_of_archive_assigments = array("QI Expedited Review ","Pending Over 30 Days ", "EXP SNAP X127FA5 ", "EXP SNAP 1800 ", "EXP SNAP DWP ", "")

previous_date = dateadd("d", -1, date)
Call change_date_to_soonest_working_day(previous_date)       'finds the most recent previous working day for the fin
file_date = replace(previous_date, "/", "-")   'Changing the format of the date to use as file path selection default
archive_folder = right("0" & DatePart("m", file_date), 2) & "-" & DatePart("yyyy", file_date)

archive_files = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & archive_folder

Set objFSO = CreateObject("Scripting.FileSystemObject")
For each assignment in array_of_archive_assigments
    file_selection_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & assignment & file_date & ".xlsx"
    objFSO.MoveFile file_selection_path , archive_files & "\" & assignment & file_date & ".xlsx"    'moving each file to the archive file
Next 

'logging usage stats
STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success, the expedited SNAP run is complete. The workbook has been saved.")