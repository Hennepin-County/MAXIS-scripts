'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - INTERVIEW WAIVER ASSIGNMENT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 60                      'manual run time in seconds
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

call changelog_update("10/06/2020", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
report_month = CM_plus_1_mo
report_year = CM_plus_1_yr
report_date = report_month & "-" & report_year

'file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\20" & report_year & "\" & report_month & " Renewals.xlsx"
file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\On Demand Waiver\Renewals\2020\10-20 Renewals - Copy.xlsx"

BeginDialog Dialog1, 0, 0, 266, 115, "ADMIN - INTERVIEW WAIVER ASSIGNMENTS"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used to create a list of assignments from cases that may meet a waived interview, and that we have forms on file in ECF."
  Text 15, 70, 230, 15, "Select the Excel file that contains your recert cases by selecting the 'Browse' button, and finding the file."
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

If objExcel = "" Then Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

sheet_name = "ER cases 10-20"
'sheet_name = "ER cases " & report_date
objExcel.worksheets(sheet_name).Activate 'Activates the applicable worksheet 
msgbox "ER cases open?"

'Establshing array     
DIM master_array()          'Declaring the array
ReDim master_array(4, 0)    'Resizing the array
master_record = 0    'incrementer for the array  

'Creating constants to value the array elements
const case_number_const         = 0  
const basket_number_const       = 1
const cash_programs_const       = 2
const forms_found_const         = 3
const form_name_const           = 4
    
excel_row = 2
    
Do 
    MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value 'reading case number
    MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do 
    
    possible_waiver = ObjExcel.Cells(excel_row, 9).Value 'reading possible waiver at COL I. Alwasys COL I in recert spreadsheets
    msgbox possible_waiver & vbcr & Excel_row
    If trim(possible_waiver) = "TRUE" or trim(possible_waiver) = "True" then 
        basket_number = ObjExcel.Cells(excel_row, 1).Value  
        cash_programs = ObjExcel.Cells(excel_row, 10).Value  'CASH PROG ACTV at COL J
        
        'Creating 8 digit MAXIS Case number to measure against next list(s)
		Do 
            If len(MAXIS_case_number) < 8 then MAXIS_case_number = "0" & MAXIS_case_number
		Loop until len(MAXIS_case_number) = 8
        
        ReDim Preserve master_array(4, master_record)	'This resizes the array based on if master notes were found or not
        master_array(MAXIS_case_number_const, master_record) = MAXIS_case_number
        master_array(basket_number_const,     master_record) = trim(basket_number)
        master_array(cash_programs_const,     master_record) = trim(cash_programs)
        master_array(forms_found_const,       master_record) = ""
        master_array(form_name_const,         master_record) = ""
            
        master_record = master_record + 1			'This increments to the next entry in the array'
        STATS_counter = STATS_counter + 1           'stats incrementor 
    End if 
    excel_row = excel_row + 1                       'Excel row incrementor
LOOP

msgbox "master record: " & master_record & vbcr & "stats count:" & STATS_counter
 
'Closing workbook and quiting Excel application
objExcel.ActiveWorkbook.Close                           
objExcel.Application.Quit
objExcel.Quit

'----------------------------------------------------------------------------------------------------Forms lists
today = right("0" & DatePart("d", date), 2)
today_year = DatePart("yyyy", date) 
today_date = CM_mo & "-" & today & "-" & today_year
msgbox today_date


array_of_assigments = array("T:\Eligibility Support\Assignments\Adults\Adults Task Based Processing Assignment " & today_date & ".xlsx", "T:\Eligibility Support\Assignments\Families\Families Task Based Processing Assignment " & today_date & ".xlsx")

For each assignment in array_of_assigments
    Call excel_open(assignment, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    objExcel.worksheets("Priority").Activate 'Activates the assignment worksheet that holds the 
    msgbox "What's happening?"
    excel_row = 4
    
    Do 
        case_number = ObjExcel.Cells(excel_row, 1).Value 'reading case number
        case_number = trim(case_number)
        If case_number = "" then exit do 
    
        For item = 0 to Ubound(master_record, 2)
            If master_array(MAXIS_case_number_const, item) = trim(case_number) then 
                msgbox item & vbcr & case_number
                master_array(forms_found_const, item) = True 
                master_array(form_name_const,   item) = trim(ObjExcel.Cells(excel_row, 2).Value)
                exit for 
            Else 
                master_array(forms_found_const, item) = False 
            End if 
        Next
        excel_row = excel + 1 
    Loop 
    
    'Closing workbook and quiting Excel application
    objExcel.ActiveWorkbook.Close                           
    objExcel.Application.Quit
    objExcel.Quit
Next 

'----------------------------------------------------------------------------------------------------Create assignment list 
'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet
ObjExcel.ActiveSheet.Name = "Potential waived Interviews"

'adding information to the Excel list from PND2
ObjExcel.Cells(1, 1).Value = "Worker #"
ObjExcel.Cells(1, 2).Value = "Case number"
ObjExcel.Cells(1, 3).Value = "Cash Programs"
ObjExcel.Cells(1, 4).Value = "Forms Found"
ObjExcel.Cells(1, 5).Value = "Form Name"

excel_row = 2
 
For item = 0 to UBound(master_array, 2)
    objExcel.Cells(excel_row, 1).Value = master_array(case_number_const,     item)
    objExcel.Cells(excel_row, 2).Value = master_array(basket_number_const,   item)
    objExcel.Cells(excel_row, 3).Value = master_array(cash_programs_const,   item)
    objExcel.Cells(excel_row, 4).Value = master_array(forms_found_const,     item)
    objExcel.Cells(excel_row, 5).Value = master_array(form_name_const,       item)
    excel_row = excel_row + 1
Next 

STATS_counter = STATS_counter - 1   'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Assignment list is compiled.")