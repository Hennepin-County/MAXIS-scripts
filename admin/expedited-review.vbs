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
CALL changelog_update("01/26/2022", "Added QI Assignment supports for strike planning pivot work.", "Ilse Ferris, Hennepin County")
CALL changelog_update("12/08/2021", "Adding option to run specialty assignment vs. whole agency lists.", "Ilse Ferris, Hennepin County")
CALL changelog_update("12/01/2021", "Set Excel Visibilty to False.", "Ilse Ferris, Hennepin County")
CALL changelog_update("09/08/2021", "Added option for a warning before the Excel/Outlook output. Changed DWP and removed QI assignments. Updated background functionality.", "Ilse Ferris, Hennepin County")
CALL changelog_update("07/10/2021", "Added Brittany Lane to YET assignment email. Removed Maslah.", "Ilse Ferris, Hennepin County")
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


Function File_Exists(file_name, does_file_exist)
    If (objFSO.FileExists(file_name)) Then
        does_file_exist = True
    Else
      does_file_exist = False
    End If
End Function

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
            excel_row = excel_row + 1
        End if
    Next

    FOR i = 1 to 7		'formatting the cells
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT
END FUNCTION

'THE SCRIPT-----------------------------------------------------------------------------------------------------------
EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
warning_checkbox = 1    'auto-checked
assignment_choice = "All Agency"

'Setting up counts for data tracking
screening_count = 0
expedited_count = 0
priv_count = 0
not_exp_count = 0   'incrementor built into the function for not expedited only

''----------------------------------------------------------------------------------------------------The current day's assignment
report_date = replace(date, "/", "-")   'Changing the format of the date to use as file path selection default
file_selection_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & report_date & ".xlsx"

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 481, 120, "ADMIN - EXPEDITED REVIEW"
  ButtonGroup ButtonPressed
    OkButton 365, 95, 50, 15
    CancelButton 420, 95, 50, 15
    PushButton 420, 55, 50, 15, "Browse...", select_a_file_button
  DropListBox 365, 75, 105, 15, "Select One..."+chr(9)+"All Agency"+chr(9)+"Speciality Only", assignment_choice
  Text 15, 20, 455, 15, "This script should be used to review a BOBI list of pending SNAP and/or MFIP cases to ensure expedited screening and determinations are being made to ensure expedited timeliness rules are being followed."
  Text 15, 40, 335, 10, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 465, 110, "Using this script:"
  EditBox 15, 55, 400, 15, file_selection_path
  CheckBox 130, 100, 230, 10, "Check here for warning message before Excel output/email creation.", warning_checkbox
  Text 270, 80, 90, 10, "Select the assignment type:"
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

Call check_for_MAXIS(False)

'Opening today's list
Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file
objExcel.worksheets("Report 1").Activate                                 'Activates the initial BOBI report

'Establishing array
DIM expedited_array()           'Declaring the array
ReDim expedited_array(appears_exp_const, 0)     'Resizing the array

'Creating constants to value the array elements
const worker_number_const       = 0
const case_number_const	        = 1
const program_ID_const 	        = 2
const days_pending_const        = 3
const application_date_const    = 4
const interview_date_const      = 5
const case_status_const         = 6
const appears_exp_const         = 7

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
    If assignment_choice = "Speciality Only" then
        If worker_number = "X127EF8" or worker_number = "X127EF9" or worker_number = "X127FA5" then
            add_to_array = True
        Else
            add_to_array = False
        End if
    Elseif assignment_choice = "All Agency" then
        add_to_array = True
    End if

    If instr(all_case_numbers_array, "*" & MAXIS_case_number & "*") then add_to_array = False

    If add_to_array = True then
        'Adding client information to the array
        ReDim Preserve expedited_array(appears_exp_const, entry_record)	'This resizes the array based on the number of cases
        expedited_array(worker_number_const,    entry_record) = worker_number
        expedited_array(case_number_const,      entry_record) = MAXIS_case_number
        expedited_array(program_ID_const,       entry_record) = program_ID
        expedited_array(days_pending_const,     entry_record) = days_pending
        expedited_array(application_date_const, entry_record) = trim(application_date)
        expedited_array(interview_date_const,   entry_record) = trim(interview_date)
        entry_record = entry_record + 1			'This increments to the next entry in the array
        stats_counter = stats_counter + 1       'Increment for stats counter
        all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*") 'Adding MAXIS case number to case number string
    End if
    excel_row = excel_row + 1
Loop

If assignment_choice = "Speciality Only" then
    objWorkbook.Save()  'saves existing workbook as same name
    objExcel.Quit
End if

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
        Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
        If is_this_priv = True then
            expedited_array(case_status_const, item) = "Privileged Case"
            expedited_array(appears_exp_const, item) = "Privileged Cases"
            priv_count = priv_count + 1
        Else
            EMReadScreen county_code, 4, 21, 14                                 'Out of county cases from CASE/CURR
            If county_code <> "X127" then
                expedited_array(case_status_const, item) = "OUT OF COUNTY CASE"
                expedited_array(appears_exp_const, item) = "Not Expedited"
            End if
        End if
    End if

    If expedited_array(appears_exp_const, item) = "" then
		Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status)

        'ACTIVE SNAP Cases - not expedited, end of evaluation
        If snap_status = "ACTIVE" or snap_status = "APP OPEN" or snap_status = "APP CLOSE" then
            'SNAP is active, EXP review not needed
            expedited_array(case_status_const, item) = "SNAP ACTIVE"
            expedited_array(appears_exp_const, item) = "Not Expedited"
            check_case_note = False
        Elseif snap_status = "PENDING" then
            'review exp snap status
            check_case_note = True
        Elseif snap_case = False then
            'If SNAP is not active but MFIP is, EXP review not needed
            IF mfip_case = True and (mfip_status = "ACTIVE" or mfip_status = "APP OPEN" or mfip_status = "APP CLOSE") then
                expedited_array(case_status_const, item) = "MFIP ACTIVE"
                expedited_array(appears_exp_const, item) = "Not Expedited"
                check_case_note = False
            ''If SNAP is not active but MFIP any status is Pending, EXP review IS needed
            ElseIf mfip_case = True and mfip_status = "PENDING" then
                'Need to review case notes to check MFIP non-active cases to see if an evauation of expedited has been completed
                check_case_note = True
            End if
        End if

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
If warning_checkbox = 1 then Msgbox "Output to Excel Starting."      'warning message to whomever is running the script

If assignment_choice = "All Agency" then
    Call Add_pages("Not Expedited")         'calling function to output exp snap statuses to Excel
    Call Add_pages("Privileged Cases")
    Call Add_pages("Req Exp Processing")
    Call Add_pages("Exp Screening Req")

    objWorkbook.Save()  'saves existing workbook as same name
    objExcel.Quit

    '----------------------------------------------------------------------------------------------------'Pending over 30 days report
    'Opening the Excel file
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = False
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
                excel_row = excel_row + 1
            End if
        End if
    Next

    FOR i = 1 to 7		'formatting the cells
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
    objExcel.Visible = False
    Set objWorkbook = objExcel.Workbooks.Add()
    objExcel.DisplayAlerts = True

    'Changes name of Excel sheet
    ObjExcel.ActiveSheet.Name = "Expedited SNAP"

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
                excel_row = excel_row + 1
            End if
        End if
    Next

    FOR i = 1 to 7		'formatting the cells
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT

    'Saves and closes the most recent Excel workbook
    objExcel.ActiveWorkbook.SaveAs t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP SNAP Interview Completed " & report_date & ".xlsx"
    objExcel.ActiveWorkbook.Close
    objExcel.Application.Quit
    objExcel.Quit

    stats_report = "Screening Count: " & screening_count & vbcr & _
    "Expedited Processing Count: " & expedited_count & vbcr & _
    "PRIV Case Count: " & priv_count & vbcr & _
    "Not Expedited Count: " & not_exp_count
    Call create_outlook_email("Ilse.Ferris@hennepin.us;Laurie.Hennen@hennepin.us","","Expedited SNAP Daily statistics for " & date, stats_report, "", True)
End if
'----------------------------------------------------------------------------------------------------Appears Expedited for YET Team only X127FA5
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
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
objExcel.Visible = False
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

If assignment_choice = "All Agency" then
    'QI Assignment hyperlink paths for email
    QI_assign_one = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP SNAP Interview Completed " & report_date & ".xlsx"
    QI_assign_two = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & report_date & ".xlsx"
    QI_assign_three = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\Pending Over 30 Days " & report_date & ".xlsx"

    body_of_email = "Interview Completed Assignment: " & "<" & QI_assign_one & ">" & vbcr & vbcr & _
    "Interview Needed-Appears Expedited Assignment :" & "<" & QI_assign_two & ">" & vbcr & vbcr & _
    "Pending Over 30 Days Assignment: " & "<" & QI_assign_three & ">"
    Call create_outlook_email("Jennifer.Frey@hennepin.us; Ilse.Ferris@hennepin.us", "Laurie.Hennen@hennepin.us", "Today's EXP SNAP Assignments are Ready", body_of_email , "", True)
End if

'Function create_outlook_email(email_recip, email_recip_CC, email_subject, email_body, email_attachment, send_email)
Call create_outlook_email("Brittany.Lane@hennepin.us; Debrice.Jackson@hennepin.us","Laurie.Hennen@hennepin.us", "EXP SNAP Report for YET without Interviews is Ready. EOM.", "", "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP SNAP X127FA5 " & report_date & ".xlsx", True)
Call create_outlook_email("Carlotta.Madison@hennepin.us", "Laurie.Hennen@hennepin.us", "EXP SNAP Report for 1800 without Interviews is Ready. EOM.", "", "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\EXP SNAP 1800 " & report_date & ".xlsx", True)
'----------------------------------------------------------------------------------------------------Moves yesterday's files to the archive folder for the specific month
array_of_archive_assigments = array("Pending Over 30 Days ", "EXP SNAP X127FA5 ", "EXP SNAP 1800 ", "EXP SNAP Interview Completed ", "")

previous_date = dateadd("d", -1, date)
Call change_date_to_soonest_working_day(previous_date, "back")       'finds the most recent previous working day
file_date = replace(previous_date, "/", "-")   'Changing the format of the date to use as file path selection default
archive_folder = right("0" & DatePart("m", file_date), 2) & "-" & DatePart("yyyy", file_date)

archive_files = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & archive_folder

For each assignment in array_of_archive_assigments
    file_selection_path = t_drive & "\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\SNAP\EXP SNAP Project\" & assignment & file_date & ".xlsx"
    Call File_Exists(file_selection_path, does_file_exist)
    If does_file_exist = True then objFSO.MoveFile file_selection_path , archive_files & "\" & assignment & file_date & ".xlsx"    'moving each file to the archive file
Next

'logging usage stats
STATS_counter = STATS_counter - 1  'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success, the expedited SNAP run is complete. The workbook has been saved.")

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------09/08/2021
'--Tab orders reviewed & confirmed----------------------------------------------09/08/2021
'--Mandatory fields all present & Reviewed--------------------------------------09/08/2021
'--All variables in dialog match mandatory fields-------------------------------09/08/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------09/08/2021-----------------N/A - No CASE:NOTE
'--CASE:NOTE Header doesn't look funky------------------------------------------09/08/2021-----------------N/A - No CASE:NOTE
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------09/08/2021-----------------N/A - No CASE:NOTE
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------09/08/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------09/08/2021-----------------N/A
'--PRIV Case handling reviewed -------------------------------------------------09/08/2021
'--Out-of-County handling reviewed----------------------------------------------09/08/2021
'--script_end_procedures (w/ or w/o error messaging)----------------------------09/08/2021
'--BULK - review output of statistics and run time/count (if applicable)--------09/08/2021-----------------N/A
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------09/08/2021
'--Incrementors reviewed (if necessary)-----------------------------------------09/08/2021
'--Denomination reviewed -------------------------------------------------------09/08/2021
'--Script name reviewed---------------------------------------------------------09/08/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------09/08/2021

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------09/08/2021
'--comment Code-----------------------------------------------------------------
'--Update Changelog for release/update------------------------------------------09/08/2021
'--Remove testing message boxes-------------------------------------------------09/08/2021
'--Remove testing code/unnecessary code-----------------------------------------09/08/2021
'--Review/update SharePoint instructions----------------------------------------09/08/2021------------------Instructions are held locally on QI's OneNote under REPORTS
'--Review Best Practices using BZS page ----------------------------------------09/08/2021-----------------N/A
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------09/08/2021-----------------N/A
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------09/08/2021
'--Complete misc. documentation (if applicable)---------------------------------09/08/2021
'--Update project team/issue contact (if applicable)----------------------------09/08/2021
