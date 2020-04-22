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
call changelog_update("05/29/2019", "Updated script to work with new BOBI query.", "Ilse Ferris, Hennepin County")
call changelog_update("01/31/2019", "Added functionality to change Defer FSET funds field if coded incorrectly on STAT/WREG.", "Ilse Ferris, Hennepin County")
call changelog_update("05/23/2018", "Added code to write in client name if presenting as a PRIV case on initial spreadsheet.", "Ilse Ferris, Hennepin County")
call changelog_update("03/30/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr
file_date = CM_mo & "-" & CM_yr

file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\ABAWD\Active SNAP " & file_date & ".xlsx"

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
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

objExcel.Cells(1, 4).Value = "FSET"
objExcel.Cells(1, 5).Value = "ABAWD"
objExcel.Cells(1, 6).Value = "BM Field"
objExcel.Cells(1, 7).Value = "Defer Funds"

FOR i = 1 to 7		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 126, 50, "Select the excel row to start"
  EditBox 75, 5, 40, 15, excel_row_to_restart
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog

do
    dialog Dialog1
    If buttonpressed = 0 then stopscript								'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

excel_row = excel_row_to_restart
update_count = 0

Do
	MAXIS_case_number = ObjExcel.Cells(excel_row, 1).Value
	MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do

    member_number = ObjExcel.Cells(excel_row, 2).Value
    member_number = right("0" & member_number, 2)

    client_name = ObjExcel.Cells(excel_row, 3).Value
    client_name = trim(client_name)

	call navigate_to_MAXIS_screen("STAT", "WREG")
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
    If PRIV_check = "PRIV" then
        ObjExcel.Cells(excel_row, 3).Value = "Privliged case"
    Else
        Call write_value_and_transmit(member_number, 20, 76)

	    EMReadScreen FSET_code, 2, 8, 50
	    EMReadScreen ABAWD_code, 2, 13, 50
        EMReadScreen banked_months, 1, 14, 50
        EMReadScreen defer_funds, 1, 8, 80

        'Updated incorrectly coded Defer FSET fund cases
        If FSET_code = "30" then
            If ABAWD_code = "05" or ABAWD_code = "09" then
                If defer_funds = "Y" then
                    update_needed = FALSE
                else
                    update_needed = True
                    update_count = update_count + 1
                    'msgbox update_needed & vbcr & FSET_code & vbcr & ABAWD_code & vbcr & defer_funds
                    PF9
                    Call write_value_and_transmit("Y", 8, 80)       'Coding the rest of the ABAWD's as N for Defer FSET funds. Even though voluntary, this code is still N.
                    EMReadScreen defer_funds, 1, 8, 80
                    transmit 'passing error messages
                    transmit
                    PF3
                End if
            End if
        End if

        ObjExcel.Cells(excel_row, 2).Value = member_number                      'writing in the member number with initial 0 trimmed.
        ObjExcel.Cells(excel_row, 4).Value = replace(FSET_code, "_", "")
	    ObjExcel.Cells(excel_row, 5).Value = replace(ABAWD_code, "_", "")
        ObjExcel.Cells(excel_row, 6).Value = replace(banked_months, "_", "")
        ObjExcel.Cells(excel_row, 7).Value = replace(defer_funds, "_", "")

        If left(client_name, 2) = "XX" then
            Call navigate_to_MAXIS_screen("STAT", "MEMB")
            Call write_value_and_transmit(member_number, 20, 76)
            EMReadScreen last_name, 25, 6, 30
            EMReadScreen first_name, 12, 6, 63
            last_name = replace(last_name, "_", "")
            first_name = replace(first_name, "_", "")
            new_client_name = last_name & "," & first_name
            ObjExcel.Cells(excel_row, 3).Value = new_client_name
        End if
    End if

    STATS_counter = STATS_counter + 1
    excel_row = excel_row + 1
Loop until ObjExcel.Cells(excel_row, 1).Value = ""

STATS_counter = STATS_counter - 1 'since we start with 1
script_end_procedure("Success! Please review your ABAWD list. Update count: " & update_count)
