'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - GATHER COLA UNEA FROM DAIL.vbs"
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
call changelog_update("12/16/2019", "Final version for 2020 COLA. Supports VA and RR, not Other Retirement.", "Ilse Ferris, Hennepin County")
call changelog_update("12/09/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
MAXIS_footer_month = CM_plus_1_mo
MAXIS_footer_year =  CM_plus_1_yr

file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\BZ ongoing projects\COLA\COLA Increase Automation\UNEA BOBI Pull.xlsx"


'The dialog is defined in the loop as it can change as buttons are pressed 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "ADMIN - GATHER UNEA INFO.vbs"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used when a list of cases with UNEA is created, and Claim number and amount information is needed."
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog
'dialog and dialog DO...Loop	
Do
    Do
    	Dialog Dialog1 
        cancel_without_confirmation
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

objExcel.Cells(1, 7).Value = "UNEA Information"

FOR i = 1 to 7	'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 126, 50, "Select the excel row to start"
  EditBox 75, 5, 40, 15, excel_row_to_start
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog 
do 
    dialog Dialog1
    cancel_without_confirmation
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

excel_row = excel_row_to_start 

Do 
	MAXIS_case_number = ObjExcel.Cells(excel_row, 2).Value
	MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do 
  		
    dail_msg = ObjExcel.Cells(excel_row, 6).Value
    dail_msg = trim(dail_msg)
    
    If instr(dail_msg, "VA DISABILITY") then income_type = "11"
    If instr(dail_msg, "VA PENSION") then income_type = "12"
    If instr(dail_msg, "VA OTHER") then income_type = "13"
    'If instr(dail_msg, "OTHER RETIREMENT") then income_type = "17"
    If instr(dail_msg, "RAILROAD RETIREMENT") then income_type = "16"
    
    'msgbox income_type
    Call navigate_to_MAXIS_screen("STAT", "PNLI")
    EmReadscreen PNLI_check, 4, 2, 53
    If PNLI_check <> "PNLI" then 
        ObjExcel.Cells(excel_row, 7).Value = "PRIV"
    Else 
        'Reading PNLI to queue up the correct income sources
        row = 3
        Do 
            EmReadScreen panel_name, 4, row, 5
            If panel_name = "UNEA" then 
                EmReadScreen income_code, 2, row, 26
                'msgbox "income code: " & income_code & vbcr & "income type: " & income_type
                If income_type = income_code then 
                    'msgbox "match!"
                    EmWriteScreen "X", row, 3
                End if 
            End if 
            row = row + 1
        Loop until row = 19
        
        Do 
            transmit 'to load up panels 
            EmReadscreen unea_confirm, 4, 2, 48
            If unea_confirm = "UNEA" then exit do
        Loop until unea_confirm = "UNEA"    
            
        'msgbox "reading unea"    
        unea_info = ""    
        Do 
            If income_type = "17" then 
                EmReadscreen claim_number_field, 15, 6, 37
                claim_number_field = replace(claim_number_field, "_", "")
                claim_number_field = trim(claim_number_field)
                If instr(claim_number_field, "PERA") then 
                    read_income = True 
                Else 
                    read_income = False 
                End if 
            End if 
            If read_income <> False then     
                EmReadscreen member_number, 2, 4, 33
                unea_info = unea_info & member_number & "|"
                
                EmReadscreen unea_amt, 8, 18, 68
                unea_amt = replace(unea_amt, "_", "")
                unea_info = unea_info & trim(unea_amt) & "|"
                
                EmReadscreen update_date, 8, 21, 55
                update_date = replace(update_date, " ", "/")
                unea_info = unea_info & trim(update_date) & "|"
            End if 
            transmit 
            EMReadscreen check_self, 4, 2, 50
        Loop until check_self = "SELF"    
            
        ObjExcel.Cells(excel_row, 7).Value = unea_info 
    End if 

    STATS_counter = STATS_counter + 1
    excel_row = excel_row + 1
    income_type = ""
Loop until ObjExcel.Cells(excel_row, 2).Value = ""

FOR i = 1 to 7	'formatting the cells'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

STATS_counter = STATS_counter - 1 'since we start with 1
script_end_procedure("Success! Please review your ABAWD list.")