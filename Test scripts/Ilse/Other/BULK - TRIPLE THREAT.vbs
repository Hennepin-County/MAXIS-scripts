'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "BULK - TRIPLE THREAT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 335                      'manual run time in seconds
STATS_denomination = "C"       			   'C is for each CASE
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
call changelog_update("12/03/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'The dialog is defined in the loop as it can change as buttons are pressed 
BeginDialog file_select_dialog, 0, 0, 221, 50, "Select the source file"
    ButtonGroup ButtonPressed
    PushButton 175, 10, 40, 15, "Browse...", select_a_file_button
    OkButton 110, 30, 50, 15
    CancelButton 165, 30, 50, 15
    EditBox 5, 10, 165, 15, file_selection_path
EndDialog

BeginDialog excel_row_dialog, 0, 0, 126, 50, "Select the excel row to start"
  EditBox 75, 5, 40, 15, excel_row_to_start
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone
tikl_date = "12/26/18"

current_date = date
Call ONLY_create_MAXIS_friendly_date(current_date)			'reformatting the dates to be MM/DD/YY format to measure against the panel dates

file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\COLA\COLA UNEA information.xlsx"

'dialog and dialog DO...Loop	
Do
    'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
    'Show initial dialog
    Do
        err_msg = ""
    	Dialog file_select_dialog
    	If ButtonPressed = cancel then stopscript
    	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If file_selection_path = "" then err_msg = err_msg & vbNewLine & "Use the Browse Button to select the file that has your client data"
        If err_msg <> "" Then MsgBox err_msg
    Loop until ButtonPressed = OK and file_selection_path <> ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

do 
    dialog excel_row_dialog
    If buttonpressed = 0 then stopscript								'loops until all errors are resolved
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'Sets up the array to store all the information for each client'
Dim Cases_array()
ReDim Cases_array (4, 0)

'Sets constants for the array to make the script easier to read (and easier to code)'
Const case_num    	= 0			'Each of the case numbers will be stored at this position'
Const send_memo     = 1
Const case_note     = 2
Const set_TIKL      = 3
Const case_status   = 4

'Now the script adds all the clients on the excel list into an array
excel_row = excel_row_to_start 're-establishing the row to start checking the members for
entry_record = 0
Do                                                            'Loops until there are no more cases in the Excel list
	MAXIS_case_number = objExcel.cells(excel_row, 1).Value          're-establishing the case numbers for functions to use
	If MAXIS_case_number = "" then exit do
	MAXIS_case_number = trim(MAXIS_case_number)
	    
	'Adding client information to the array'
	ReDim Preserve Cases_array(4, entry_record)	'This resizes the array based on the number of rows in the Excel File'
	Cases_array (case_num, 	entry_record) = MAXIS_case_number		'The client information is added to the array'
    Cases_array (send_memo, entry_record) = "" 
	Cases_array (case_note, entry_record) = "" 
    Cases_array (set_TIKL,	entry_record) = "" 
    Cases_array (case_status, entry_record) = "" 

	entry_record = entry_record + 1			'This increments to the next entry in the array'
	Stats_counter = stats_counter + 1
	excel_row = excel_row + 1
Loop

excel_row = excel_row_to_start 're-establishing the row to start checking the members for
back_to_self
EMWriteScreen MAXIS_footer_month, 20, 43		'Writes in Current month plus one
EMWriteScreen MAXIS_footer_year, 20, 46		'Writes in Current month plus one's year

For i = 0 to Ubound(Cases_array, 2)
	'Establishing values for each case in the array of cases 
	MAXIS_case_number	= Cases_array (case_num, i)
    
    Call navigate_to_MAXIS_screen("SPEC", "MEMO")
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets added to priv case list
    If PRIV_check = "PRIV" then
        Cases_array(case_status, i) = "PRIV case."
        Cases_array(send_memo, i) = False
        Cases_array(case_note, i) = False
        Cases_array(set_TIKL, i) = False
   
        'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
        Do
            back_to_self
            EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
            If SELF_screen_check <> "SELF" then PF3
        LOOP until SELF_screen_check = "SELF"
        EMWriteScreen "________", 18, 43		'clears the case number
        transmit
    else 
        EMReadscreen county_code, 2, 20, 16 'coordinates at SPEC/MEMO
        If county_code <> "27" then 
            Cases_array(case_status, i) = "Not Hennepin County case: " & county_code	'Explanation for the rejected report
            Cases_array(send_memo, i) = False
            Cases_array(case_note, i) = False
            Cases_array(set_TIKL, i) = False
        End if
	End if 
    
    IF Cases_array (case_status, i) = "" then 
        Call MAXIS_background_check
        '----------------------------------------------------------------------------------------------------THE SPEC/MEMO
        Call start_a_new_spec_memo            
        'Writes the MEMO.
        call write_variable_in_SPEC_MEMO("Dear recipients,")
        Call write_variable_in_SPEC_MEMO("")
        Call write_variable_in_SPEC_MEMO("As of January 1, 2019 there may be a small increase to your income called a cost of living adjustment (COLA). The income affected could be annuities, private pension, retirement plans, or other regular sources of income.")
        Call write_variable_in_SPEC_MEMO("")
        Call write_variable_in_SPEC_MEMO("Please send in your award letter or any notice that you may receive stating the amount of the increase for 2019. You can also contact the office from which you receive your income, and ask them to forward you a letter stating what your gross monthly income will be after the cost of living increase.")
        Call write_variable_in_SPEC_MEMO("")	
        Call write_variable_in_SPEC_MEMO("If you carry private health insurance (such as a Medicare supplement, Medicare D or other health insurance not provided by Medical Assistance) we will also need you to send proof of any insurance cost adjustments. This is so we can correctly budget your income and deductions for 2019.")
        Call write_variable_in_SPEC_MEMO("")
        Call write_variable_in_SPEC_MEMO("If you have questions, please call 612-596-1300 to speak with a team member. Thank you.")
        
        PF4			'Exits the MEMO
        EMReadScreen memo_sent, 8, 24, 2
        If memo_sent <> "NEW MEMO" then 
            Cases_array(case_status, i) = "Error"
            Cases_array(send_memo, i) = False	'Explanation for the rejected report'
            PF10
        Else 
            Cases_array(send_memo, i) = True 'Explanation for the rejected report'
        End if  
        
        back_to_self
        
        '----------------------------------------------------------------------------------------------------THE CASE NOTE
        start_a_blank_CASE_NOTE
        Call write_variable_in_CASE_NOTE("Sent SPEC/MEMO re: COLA income/deduction verifs needed")
        Call write_variable_in_CASE_NOTE("Set TIKL for " & tikl_date & " to review case.") 
        Call write_variable_in_CASE_NOTE("If verifications have not been provided, a verification request in ECF will need to be sent.")
        call write_variable_in_case_note("---")
        call write_variable_in_case_note(worker_signature)
        
        'ensuring that the case note saved. If not, adding it to the notes for the user to review. 
        PF3
        
        EMReadScreen note_date, 8, 5, 6
        If note_date <> current_date then 
            Cases_array(case_note, i) = False 
        Else 
            Cases_array(case_note, i) = True 
        End if 	

        Call navigate_to_MAXIS_screen ("DAIL", "WRIT")
        
        call create_MAXIS_friendly_date(tikl_date, 0, 5, 18)
        call write_variable_in_TIKL ("SPEC/MEMO COLA sent to case. If verification of income and/or deductions have not been provided send verification request in ECF.")
        PF3
        Cases_array(set_TIKL, i) = True 
    End if       
      
    ObjExcel.Cells(Excel_row, 2).Value = Cases_array(send_memo, i) 
    ObjExcel.Cells(Excel_row, 3).Value = Cases_array(case_note,  i) 
    ObjExcel.Cells(Excel_row, 4).Value = Cases_array(set_TIKL,  i)
    ObjExcel.Cells(Excel_row, 5).Value = Cases_array(case_status,  i)
    Excel_row = Excel_row + 1
Next 
    
Stats_counter = stats_counter + 1
script_end_procedure("Success! The list is complete. Please review the cases that appear to be in error.")