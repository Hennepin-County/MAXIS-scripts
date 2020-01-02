'Required for statistical purposes===============================================================================
name_of_script = "BULK - LANGUAGE UPDATES.vbs "
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 40                     'manual run time in seconds
STATS_denomination = "C"       'C is for each Case
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
call changelog_update("01/22/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
'Connects to BlueZone and establishing county name
EMConnect ""	
file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\Completed projects\Language update\07-2018 updates.xlsx"


'DIALOG
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 110, "Language updates script"
    ButtonGroup ButtonPressed
    PushButton 200, 45, 50, 15, "Browse...", select_a_file_button
    OkButton 145, 90, 50, 15
    CancelButton 200, 90, 50, 15
    EditBox 15, 45, 180, 15, file_selection_path
    GroupBox 10, 5, 250, 80, "Using the LANGUAGE UPDATES script"
    Text 20, 20, 235, 20, "This script should be used when a list of cases is provided that require the language code on STAT/MEMB to be updated."
    Text 15, 65, 230, 15, "Select the Excel file that contains the client's information by selecting the 'Browse' button, and finding the file."
EndDialog
'dialog and dialog DO...Loop	
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
    	Dialog Dialog1 
    	cancel_without_confirmation
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Ensures that user is in current month
back_to_self
excel_row = 2

excel_col = 11
DO  
    'Grabs the case number
	MAXIS_case_number = objExcel.cells(excel_row, 1).value
	MAXIS_case_number = trim(MAXIS_case_number)
	If MAXIS_case_number = "" then exit do
	
	PMI_number = objExcel.cells(excel_row, 2).value
	PMI_number = trim(PMI_number)
	
	lang_code = objExcel.cells(excel_row, 5).value
	lang_code = trim(lang_code)
	lang_code = right("0" & lang_code, 2)
	
	lang_description = objExcel.cells(excel_row, 6).value
	lang_description = Trim(lang_description)
	
    Call navigate_to_MAXIS_screen("STAT", "MEMB")
	EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
	If PRIV_check = "PRIV" then
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the MAXIS case number
		transmit
		ObjExcel.Cells(excel_row, excel_col).Value = "Priv case. Cannot access."
	Else  
        Do 
            EMReadScreen PMI, 8, 4, 46
            If trim(PMI) = PMI_number then 
                found_case = true 
                exit do 
            else 
                transmit
                found_case = false 
            End if 
            EMReadScreen MEMB_error, 5, 24, 2
        Loop until MEMB_error = "ENTER"
        
        If found_case = false then
            ObjExcel.Cells(excel_row, excel_col).Value = "Cannot find PMI."
        Elseif found_case = true then
            PF9
            EmReadscreen error_check, 14, 24, 2
            If error_check = "YOU HAVE READ " then 
                ObjExcel.Cells(excel_row, excel_col).Value = "Case is Out-of-county case."
            Elseif error_check = "MAXIS PROGRAMS" then 
                ObjExcel.Cells(excel_row, excel_col).Value = "Case is inactive."
            Else 
                EmReadscreen current_code, 2, 12, 42
                If current_code = lang_code then 
                    Emreadscreen current_lang, 16, 12, 46
                    'msgbox current_lang
                    current_lang = replace(current_lang, "_", "")
                    If trim(current_lang) = lang_description then 
                        'msgbox current_lang = lang_description
                        ObjExcel.Cells(excel_row, excel_col).Value = "Case already reflects language."
                    else 
                        Call clear_line_of_text(12, 46)
                        EmWriteScreen lang_description, 12, 46
                    End if   
                Else 
                    EmWriteScreen "__", 12, 42
                    EmWriteScreen "__", 13, 42
                    transmit 
                    EmWriteScreen lang_code, 12, 42
                    If lang_code = "98" then 
                        Call clear_line_of_text(12, 46)
                        EmWriteScreen lang_description, 12, 46
                    End if 
                        
                    'Accounting for ASL persons as written langage cannot be ASl.
                    If lang_code = "08" then 
                        EmWriteScreen "99", 13, 42
                        'msgbox "ASL"
                    Else 
                        EmWriteScreen lang_code, 13, 42
                    End if 
                    If lang_code = "98" then 
                        EmWriteScreen "_________________________", 13, 46
                        EmWriteScreen lang_description, 13, 46
                    End if 
                End if 
        
                transmit
                PF3 'to stat/wrap 
                EMReadScreen wrap_check, 4, 2, 46
                If trim(wrap_check) = "WRAP" then ObjExcel.Cells(excel_row, excel_col).Value = "Updated!"
                PF3 'back to start 
            End if 
	    End if 
	End if 
    MAXIS_case_number = ""
	PMI_number = ""
	lang_code = ""
    lang_description = ""
    
    excel_row = excel_row + 1
	STATS_counter = STATS_counter + 1
    back_to_SELF
LOOP UNTIL objExcel.Cells(excel_row, 1).value = ""	'looping until the list of cases to check for recert is complete
	
STATS_counter = STATS_counter - 1 'removes one from the count since 1 is counted at the beginning (because counting :p)
script_end_procedure("Success! The list is complete.")