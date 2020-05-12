'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - DELETE INAC DAILS.vbs"
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
call changelog_update("03/20/2020", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

file_selection_path = "C:\Desktop\Inactive Case DAIL Messages.xlsx"

'dialog and dialog DO...Loop
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "BULK - DELETE INAC DAIL MESSAGES"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used when a list of DAIL messages have been identified as being inactive need to be deleted."
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


'resets the case number and footer month/year back to the CM (REVS for current month plus two has is going to be a problem otherwise)
back_to_self
EMWriteScreen CM_mo, 20, 43
EMWriteScreen CM_yr, 20, 46
transmit

excel_row = excel_row_to_restart

Do
    dail_row = 6
	'Grabs the case number
	MAXIS_case_number = objExcel.cells(excel_row, 1).value
    MAXIS_case_number = trim(MAXIS_case_number)
	If MAXIS_case_number = "" then exit do
    
    dail_msg = objExcel.cells(excel_row, 6).value
    dail_msg = trim(dail_msg)
    
    Call navigate_to_MAXIS_screen("DAIL", "DAIL")
    
    'Determining if there is a new case number...
    EMReadScreen new_case, 8, dail_row, 63
    new_case = trim(new_case)
    IF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
    Call write_value_and_transmit("T", dail_row, 3)
    dail_row = 6
    ELSEIF new_case = "CASE NBR" THEN
    '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
    Call write_value_and_transmit("T", dail_row + 1, 3)
    dail_row = 6
    End if
    
    EMReadScreen dail_info_msg, 78, 24, 2
    msgbox "dail_info_msg: " & dail_info_msg
    dail_info_msg = trim(dail_info_msg)
    If instr(dail_info_msg, "NO MESSAGES FOR CASE") then 
        msgbox  "DAIL already deleted."
        objExcel.cells(excel_row, 8).value = "DAIL already deleted."
    Else 
        Do 
            msgbox "dail_row: " & dail_row
            EMReadScreen case_confirmation, 8, 5, 73
            msgbox "case_confirmation " & case_confirmation
            If trim(case_confirmation) = "" then 
                objExcel.cells(excel_row, 8).value = "DAIL already deleted."
                exit do 
            End if 
            If trim(case_confirmation) = MAXIS_case_number then
                msgbox "case numbers match"
                EMReadScreen current_dail, 61, dail_row, 20 
                If trim(current_dail) = dail_msg then 
                    msgbox "DAIL message found"
                    'Deleting the DAIL 
                    dail_found = true
                    Call write_value_and_transmit("D", dail_row, 3)
                    EMReadScreen other_worker_error, 78, 24, 2
                    If trim(other_worker_error) = "** WARNING ** YOU WILL BE DELETING  ANOTHER WORKERS DAIL MESSAGES." then 
                        transmit
                        objExcel.cells(excel_row, 8).value = "Deleted"
                        deleted_dails = deleted_dails + 1
                    elseif trim(other_worker_error) <> "" then 
                        Call write_value_and_transmit("_", dail_row, 3)
                        objExcel.cells(excel_row, 8).value = trim(other_worker_error)
                        dail_found = true
                    End if 
                else
                    dail_found = False
                    EMReadScreen next_dail, 9, dail_row + 1, 63
                    If next_dail = "CASE NBR:" then 
                        EMReadScreen new_case_number, 8, dail_row + 1, 73
                        If trim(new_case_number) <> MAXIS_case_number then 
                            objExcel.cells(excel_row, 8).value = "Could not find DAIL."
                            exit do 
                        End if 
                    Else 
                        Call write_value_and_transmit("T", dail_row + 1, 3)
                    End if 
                End if
            Else 
                'msgbox "still looking for DAIL"
                'Determining if there is a new case number...
                EMReadScreen new_case, 8, dail_row + 1, 63
                new_case = trim(new_case)
                IF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
                    Call write_value_and_transmit("T", dail_row + 1, 3)
                    dail_row = 6
                ELSEIf new_case = "CASE NBR" THEN
                    EMReadScreen new_case_number, 8, dail_row + 1, 73
                    If trim(new_case_number) <> MAXIS_case_number then 
                        objExcel.cells(excel_row, 8).value = "Could not find DAIL."
                        exit do 
                    Else 
                        '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
                        Call write_value_and_transmit("T", dail_row + 2, 3)
                        dail_row = 6
                    End if 
                End if
            End if
            
        Loop until dail_found = true 
    End if 
    excel_row = excel_row + 1 
    
LOOP UNTIL objExcel.Cells(excel_row, 1).value = ""	'looping until the list of cases to check for recert is complete
STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)

script_end_procedure("All done, woo hoo!")