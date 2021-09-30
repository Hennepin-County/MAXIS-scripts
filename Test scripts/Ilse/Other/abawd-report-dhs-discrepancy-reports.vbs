worker_county_code = "x127"
'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - ABAWD REPORT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 60                      'manual run time in seconds
STATS_denomination = "M"       			   'M is for each CASE
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
call changelog_update("06/17/2021", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""

'column constants
pmi_col         =  1
case_number_col = 2
fset_col        = 11
abawd_col       = 12
memb_numb_col   = 18
snap_status_col = 19
notes_col       = 20
case_active_col = 21

MAXIS_footer_month = "07"
MAXIS_footer_year = "21"

'file_selection_path = "C:\Users\ilfe001\OneDrive - Hennepin County\Desktop\SNAP Work\ABAWD Report 10-2020 thru 06-2021 PT 2.xlsx"

'dialog and dialog DO...Loop
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "BULK - ABAWD REPORT"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used when a list of SNAP cases wtih member numbers are provided by BOBI to gather ABAWD, FSET and Banked Months information."
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog Dialog1
    	cancel_without_confirmation
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

back_to_SELF

Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

FOR i = 1 to 21		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'For Each objWorkSheet In objWorkbook.Worksheets 'Creating an array of worksheets that are not the intitial report - "Report 1"
'    If objWorkSheet.Name = "10-20" then sheet_list = sheet_list & objWorkSheet.Name & ","
'Next

'For Each objWorkSheet In objWorkbook.Worksheets 'Creating an array of worksheets that are not the intitial report - "Report 1"
'    If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "All cases" and objWorkSheet.Name <> "Data" then sheet_list = sheet_list & objWorkSheet.Name & ","
'Next
'
'sheet_list = trim(sheet_list)  'trims excess spaces of sheet_list
'If right(sheet_list, 1) = "," THEN sheet_list = left(sheet_list, len(sheet_list) - 1) 'trimming off last comma
'array_of_sheets = split(sheet_list, ",")   'Creating new array
'
'For each excel_sheet in array_of_sheets
''    objExcel.worksheets(excel_sheet).Activate 'Activates the applicable worksheet

    'MAXIS_footer_month = left(excel_sheet, 2)
    'MAXIS_footer_year = right(excel_sheet, 2)
    Call MAXIS_footer_month_confirmation

    excel_row = 2

    Do
    	PMI_number = trim(ObjExcel.Cells(excel_row, pmi_col).Value)

        MAXIS_case_number = ObjExcel.Cells(excel_row, case_number_col).Value
    	MAXIS_case_number = trim(MAXIS_case_number)
        If MAXIS_case_number = "" then exit do

        Call navigate_to_MAXIS_screen_review_PRIV("STAT", "MEMB", is_this_priv)
        EmReadscreen self_screen, 4, 2, 50
        EmReadscreen self_error, 60, 24, 2
        If is_this_priv = True then
            ObjExcel.Cells(excel_row, notes_col).Value = "Privliged case"
        Elseif (is_this_priv = False and self_screen = "SELF") then
            ObjExcel.Cells(excel_row, notes_col).Value = trim(self_error)
        Else
            Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status)
            ObjExcel.Cells(excel_row, snap_status_col).Value = snap_case
            ObjExcel.Cells(excel_row, case_active_col).Value = case_active

            EmReadscreen county_code, 4, 21, 14 'reading from CASE/CURR
            If county_code <> UCASE(worker_county_code) then ObjExcel.Cells(excel_row, notes_col).Value = "Out-of-county Case"
            Call navigate_to_MAXIS_screen("STAT", "MEMB")
            Do
                EmReadscreen memb_panel_PMI, 8, 4, 46
                memb_panel_PMI = right ("00000000" & trim(memb_panel_PMI), 8)
                If trim(memb_panel_PMI) = PMI_number then
                    EmReadscreen member_number, 2, 4, 33
                    Exit do
                Else
                    transmit
                    EmReadscreen end_of_membs_message, 5, 24, 2
                End if
            Loop until end_of_membs_message = "ENTER"

            If trim(member_number) = "" then
                ObjExcel.Cells(excel_row, notes_col).Value = "Unable to find member on case"
            Else
    	           call navigate_to_MAXIS_screen("STAT", "WREG")
                Call write_value_and_transmit(member_number, 20, 76)

    	           EMReadScreen FSET_code, 2, 8, 50
    	           EMReadScreen ABAWD_code, 2, 13, 50

                ObjExcel.Cells(excel_row, memb_numb_col).Value = member_number                      'writing in the member number with initial 0 trimmed.
                ObjExcel.Cells(excel_row, fset_col).Value = replace(FSET_code, "_", "")
    	           ObjExcel.Cells(excel_row, abawd_col).Value = replace(ABAWD_code, "_", "")
            End if
        End if
        STATS_counter = STATS_counter + 1
        excel_row = excel_row + 1
    Loop until ObjExcel.Cells(excel_row, 2).Value = ""
'Next

FOR i = 1 to 21		'formatting the cells'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

STATS_counter = STATS_counter - 1 'since we start with 1
script_end_procedure("Success! Please review your ABAWD list.")
