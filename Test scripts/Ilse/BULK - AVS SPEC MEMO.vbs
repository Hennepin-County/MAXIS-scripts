'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - AVS SPEC MEMO.vbs"
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
call changelog_update("09/23/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

FUNCTION convert_excel_letter_to_excel_number(excel_col)
	IF isnumeric(excel_col) = FALSE THEN
		alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		excel_col = ucase(excel_col)
		IF len(excel_col) = 1 THEN
			excel_col = InStr(alphabet, excel_col)
		ELSEIF len(excel_col) = 2 THEN
			excel_col = (26 * InStr(alphabet, left(excel_col, 1))) + (InStr(alphabet, right(excel_col, 1)))
		END IF
	ELSE
		excel_col = CInt(excel_col)
	END IF
END FUNCTION

'defining this function here because it needs to not end the script if a MEMO fails.
function start_a_new_spec_memo_and_continue(success_var)
'--- This function navigates user to SPEC/MEMO and starts a new SPEC/MEMO, selecting client, AREP, and SWKR if appropriate
'===== Keywords: MAXIS, notice, navigate, edit
    success_var = True
	call navigate_to_MAXIS_screen("SPEC", "MEMO")				'Navigating to SPEC/MEMO

	PF5															'Creates a new MEMO. If it's unable the script will stop.
	EMReadScreen memo_display_check, 12, 2, 33
	If memo_display_check = "Memo Display" then success_var = False

	'Checking for an AREP. If there's an AREP it'll navigate to STAT/AREP, check to see if the forms go to the AREP. If they do, it'll write X's in those fields below.
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "ALTREP", row, col         'Row and col are variables which change from their above declarations if "ALTREP" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
	    arep_row = row                                          'Logs the row it found the ALTREP string as arep_row
	    call navigate_to_MAXIS_screen("STAT", "AREP")           'Navigates to STAT/AREP to check and see if forms go to the AREP
	    EMReadscreen forms_to_arep, 1, 10, 45                   'Reads for the "Forms to AREP?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")           'Navigates back to SPEC/MEMO
	    PF5                                                     'PF5s again to initiate the new memo process
	END IF
	'Checking for SWKR
	row = 4                             'Defining row and col for the search feature.
	col = 1
	EMSearch "SOCWKR", row, col         'Row and col are variables which change from their above declarations if "SOCWKR" string is found.
	IF row > 4 THEN                     'If it isn't 4, that means it was found.
	    swkr_row = row                                          'Logs the row it found the SOCWKR string as swkr_row
	    call navigate_to_MAXIS_screen("STAT", "SWKR")         'Navigates to STAT/SWKR to check and see if forms go to the SWKR
	    EMReadscreen forms_to_swkr, 1, 15, 63                'Reads for the "Forms to SWKR?" Y/N response on the panel.
	    call navigate_to_MAXIS_screen("SPEC", "MEMO")         'Navigates back to SPEC/MEMO
	    PF5                                           'PF5s again to initiate the new memo process
	END IF
	EMWriteScreen "x", 5, 12                                        'Initiates new memo to client
	IF forms_to_arep = "Y" THEN EMWriteScreen "x", arep_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	IF forms_to_swkr = "Y" THEN EMWriteScreen "x", swkr_row, 12     'If forms_to_arep was "Y" (see above) it puts an X on the row ALTREP was found.
	transmit                                                        'Transmits to start the memo writing process
end function

'THE SCRIPT-------------------------------------------------------------------------------------------------------------------------
EMConnect ""		'Connects to BlueZone

file_selection = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\AVS\AVS Forms Distribution Master List.xlsx"

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

'Set objWorkSheet = objWorkbook.Worksheet
For Each objWorkSheet In objWorkbook.Worksheets
	If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "controls" then scenario_list = scenario_list & chr(9) & objWorkSheet.Name
Next

'Dialog to select worksheet
'DIALOG is defined here so that the dropdown can be populated with the above code
BeginDialog worksheet_dlg, 0, 0, 151, 75, "AVS Case list"
  DropListBox 5, 35, 140, 15, "Select One... & scenario_list", scenario_dropdown
  ButtonGroup ButtonPressed
    OkButton 40, 55, 50, 15
    CancelButton 95, 55, 50, 15
  Text 5, 10, 130, 20, "Select the correct worksheet with list of recipents to create AVS SPEC/MEMOs:"
EndDialog

'Shows the dialog to select the correct worksheet
Do
    Dialog worksheet_dlg
    If ButtonPressed = cancel then stopscript
Loop until scenario_dropdown <> "Select One..."

'Activates worksheet based on user selection
objExcel.worksheets(scenario_dropdown).Activate

'Creating an array of letters to loop through
col_ind = "A~B~C~D~E~F~G~H~I~J~K~L~M~N~O~P~Q~R~S~T~U~V~W~X~Y~Z~AA~AB~AC~AD~AE~AF~AG~AH~AI~AJ~AK~AL~AM~AN~AO~AP~AQ~AR~AS~AT~AU~AV~AW~AX~AY~AZ"
col_array = split(col_ind, "~")
'setting the start of the list of column options
column_list = "Select One..."
cell_val = 1        'starting the value for reading the top cell of each column to use header information

'looping through the array
For each letter in col_array
    col_header = UCase(objExcel.Cells(1, cell_val).Value)
    col_header = trim(col_header)

    If col_header <> ""  then                                              'if the column is not blank - add to dropdown
        column_list = column_list & chr(9) & letter & " - " & col_header
        if col_header = "Case #" then case_number_column = letter & " - " & col_header    'if the first cell says 'Case Number' then it is likely the correct column
    Else
        last_col = letter       'setting this for adding additional columns with information
        Exit For
    End If
    cell_val = cell_val + 1
Next

excel_row_to_start = "2"

'Next dialog determines the column the case numbers are in and the type of notification to be sent.
'Defining the dialog here so that the list of columns can be dynamically generated
BeginDialog list_details_dlg, 0, 0, 266, 135, "AVS SPEC/MEMO Options"
  DropListBox 160, 70, 100, 45, "column_list", case_number_column
  DropListBox 160, 90, 100, 45, "Select One..."+chr(9)+"Initial"+chr(9)+"Secondary", notice_type
  EditBox 75, 115, 40, 15, excel_row_to_start
  ButtonGroup ButtonPressed
    OkButton 155, 115, 50, 15
    CancelButton 210, 115, 50, 15
  Text 10, 10, 245, 10, "Check the Excel File that has been opened to ensure it is the correct file."
  Text 10, 25, 245, 35, "Choose the column that has all the case numbers listed and select which type of notice should be sent. Initial is for recipients who have not yet rec'd at letter from DHS. Secondary is for recipients who have gotten their AVS letter."
  Text 10, 70, 145, 10, "Indicate the column with the case numbers:"
  Text 10, 95, 140, 10, "Which type of notice do you want to send?"
  Text 10, 120, 60, 10, "Excel row to start:"
EndDialog

'Displaying the dialog to select the correct column and type of notice.
Do    
    Do
        err_msg = ""
        Dialog list_details_dlg
        cancel_without_confirmation
        If notice_type = "Select one..." then err_msg = err_msg & vbcr & "* Select the notice type to send."
        If IsNumeric(excel_row_to_start) = False then err_msg = err_msg & vbcr & "* Enter a numeric row to start the script."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine		'error message including instruction on what needs to be fixed from each mandatory field if incorrect
    LOOP UNTIL err_msg = ""									'loops until all errors are resolved
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
	MAXIS_case_number = objExcel.cells(excel_row, 4).Value          're-establishing the case numbers for functions to use
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