'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - GATHER BANKED MONTHS CASES.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 60                      'manual run time in seconds
STATS_denomination = "C"       			   'M is for each CASE
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
call changelog_update("06/20/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""

file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\SNAP\Banked months data\Ongoing banked months list.xlsx"

'-------------------------------------------------------------------------------------------------DIALOG
Dialog1 = "" 'Blanking out previous dialog detail
BeginDialog Dialog1, 0, 0, 266, 115, "Gather Banked months"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used to gather a list of ongoing banked months cases to use as a filter to determine new banked months cases."
  Text 15, 70, 230, 15, "Select the Excel file that contains your information by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog

Do
    err_msg = ""
	dialog Dialog1
	cancel_without_confirmation
	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue."
    If err_msg <> "" Then MsgBox err_msg
Loop until err_msg = ""
If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

excel_row = 2
entry_record = 0

DIM ongoing_array()
ReDim ongoing_array(2,0)

const case_number   = 0
const memb_number   = 1

Do
	MAXIS_case_number = ObjExcel.Cells(excel_row, 1).Value
	MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do
    member_number = ObjExcel.Cells(excel_row, 2).Value
    member_number = right("0" & member_number, 2)

    ReDim Preserve ongoing_array(2, entry_record)	'This resizes the array based on the number of rows in the Excel File'
    ongoing_array(case_number,	entry_record) = MAXIS_case_number
    ongoing_array(memb_number, 	entry_record) = member_number

    entry_record = entry_record + 1			'This increments to the next entry in the array'
    excel_row = excel_row + 1
LOOP

objExcel.Quit   'Closes the initial spreadsheet
objExcel = ""

'----------------------------------------------------------------------------------------------------GATHERING THE LIST OF ALL BANKED MONTHS CASES
'dialog and dialog DO...Loop
BeginDialog , 0, 0, 266, 115, "Current Month Banked Months List"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection
  Text 20, 20, 235, 25, "Now select the list of all banked months cases to start the filter process."
  Text 15, 70, 230, 15, "Select the Excel file that contains your information by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog

file_selection = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\ABAWD\Active SNAP "& CM_mo & "-" & CM_yr & ".xlsx"

Do
    err_msg = ""
	dialog
	cancel_without_confirmation
	If ButtonPressed = select_file_button then call file_selection_system_dialog(file_selection, ".xlsx")
    If trim(file_selection) = "" then err_msg = err_msg & vbcr & "* Select a file to continue."
    If err_msg <> "" Then MsgBox err_msg
Loop until err_msg = ""
If objExcel = "" Then call excel_open(file_selection, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

'----------------------------------------------------------------------------------------------------FILTERING THE ARRAY
Dim new_cases_array
ReDim new_cases_array(3,0)
new_cases = 0

const case_number_const     = 0
const member_number_const   = 1
const client_name_const     = 2

excel_row = 2

DO
    MAXIS_case_number = ObjExcel.Cells(excel_row, 1).Value
    MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do
    member_number = ObjExcel.Cells(excel_row, 2).Value
    member_number = right("0" & member_number, 2)
    client_name = ObjExcel.Cells(excel_row, 3).Value
    client_name = trim(client_name)

    For item = 0 to UBound(ongoing_array, 2)
        banked_month_case_number = ongoing_array(case_number, item)
        banked_months_member = ongoing_array(memb_number, item)

        If banked_month_case_number = MAXIS_case_number then
            if banked_months_member = member_number then
                match_found = True
            else
                match_found = False
            end if
        Else
            match_found = False
        End if
        if match_found = true then exit for
    Next

    If match_found = false then
        ReDim Preserve new_cases_array(3,   new_cases)	'This resizes the array based on the number of rows in the Excel File'
        new_cases_array(case_number_const,	 new_cases) = MAXIS_case_number
        new_cases_array(member_number_const, new_cases) = member_number
        new_cases_array(client_name_const, 	 new_cases) = trim(client_name)
        new_cases = new_cases + 1			'This increments to the next entry in the array'
    End if

    excel_row = excel_row + 1
Loop

'----------------------------------------------------------------------------------------------------EXPORTING NEW BANKED MONTHS LIST
ObjExcel.Worksheets.Add().Name = "New Banked Months Cases"

'adding column header information to the Excel list
ObjExcel.Cells(1, 1).Value = "Case Number"
ObjExcel.Cells(1, 2).Value = "Member Number"
ObjExcel.Cells(1, 3).Value = "Last Name"
ObjExcel.Cells(1, 4).Value = "First Name"

'formatting the cells
FOR i = 1 to 4
	objExcel.Cells(1, i).Font.Bold = True		'bold font
    ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

excel_row = 2
For item = 0 to UBound(new_cases_array, 2)
	objExcel.Cells(excel_row, 1).Value = new_cases_array(case_number_const, item)
	objExcel.Cells(excel_row, 2).Value = new_cases_array(member_number_const, item)
	objExcel.Cells(excel_row, 3).Value = new_cases_array(client_name_const, item)
	excel_row = excel_row + 1
Next

'formatting the cells
FOR i = 1 to 4
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

STATS_counter = STATS_counter - 1 'since we start with 1
script_end_procedure("Success! Please review your Banked months list.")
