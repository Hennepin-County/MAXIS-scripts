'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - FIND FIRST BANKED MONTH.vbs"
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
call changelog_update("03/21/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------DIALOG
'The dialog is defined in the loop as it can change as buttons are pressed 
BeginDialog info_dialog, 0, 0, 266, 115, "BULK - FIND FIRST BANKED MONTH"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used when a list of SNAP cases wtih member numbers are provided by BOBI to gather ABAWD, FSET and Banked Months information."
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog

BeginDialog excel_row_dialog, 0, 0, 126, 50, "Select the excel row to start"
  EditBox 75, 5, 40, 15, excel_row_to_restart
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
file_selection_path = "C:\Users\ilfe001\Desktop\Banked months first month.xlsx"

starting_month = "05/01/18"
month_plus_one = CM_plus_1_mo & "/" & CM_plus_1_yr
months_list = "05/18,"
month_count = 1

Do     
    var_month =  right("0" & DatePart("m",    DateAdd("m", month_count, starting_month)), 2)
    var_year =  right(      DatePart("yyyy", DateAdd("m", month_count, starting_month)), 2)
    add_month = var_month & "/" & var_year
    months_list = months_list & add_month & ", "
    month_count = month_count + 1
Loop until add_month = month_plus_one

If right(months_list, 1) = "," THEN months_list = left(months_list, len(months_list) - 1)
months_array = split(months_list, ",")

'dialog and dialog DO...Loop	
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    Do
    	Dialog info_dialog
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

do 
    dialog excel_row_dialog
    If buttonpressed = 0 then stopscript								'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

excel_row = excel_row_to_restart

Do 
	MAXIS_case_number = ObjExcel.Cells(excel_row, 1).Value
	MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do 
    
    member_number = ObjExcel.Cells(excel_row, 2).Value
    member_number = "0" & right(member_number, 2)

    first_month = ObjExcel.Cells(excel_row, 6).Value
    first_month = trim(first_month)
    Case_banked = ObjExcel.Cells(excel_row, 17).Value
    case_banked = trim(case_banked)
    
    If first_month = "" then
        If case_banked <> "TRUE" then 
	       Call navigate_to_MAXIS_screen("STAT", "WREG")
           EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
           If PRIV_check = "PRIV" then
               ObjExcel.Cells(excel_row, 6).Value = "PRIV"
            Else 
                For each footer_month in months_array
                    footer_month = trim(footer_month)
                    first_mo_found = false 
                    
                    footer_mo = left(footer_month, 2)
                    footer_yr = right(footer_month, 2)
                    back_to_SELF
            
                    'msgbox footer_mo & "/" & footer_yr
                    
                    EmWriteScreen footer_mo, 20, 43
                    EmWriteScreen footer_yr, 20, 46
                    transmit 
                    Call navigate_to_MAXIS_screen("STAT", "WREG")
                    Call write_value_and_transmit(member_number, 20, 76)
        
	                EMReadScreen ABAWD_code, 2, 13, 50
                    If ABAWD_code = "13" then 
                        first_mo_found = TRUE
                        ObjExcel.Cells(excel_row, 7).Value = footer_month
                        exit for
                    else 
                        first_mo_found = false 
                    End if 
                Next 
            End if 
        End if 
    End if 
    STATS_counter = STATS_counter + 1
    excel_row = excel_row + 1
    'msgbox excel_row
Loop until ObjExcel.Cells(excel_row, 1).Value = ""

STATS_counter = STATS_counter - 1 'since we start with 1
script_end_procedure("Success! Please review your ABAWD list.")