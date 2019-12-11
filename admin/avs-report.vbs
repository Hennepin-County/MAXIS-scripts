'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - AVS REPORT.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 300                      'manual run time in seconds
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
call changelog_update("11/06/2019", "Added ability to run all spreadsheets in a process concurrently.", "Ilse Ferris, Hennepin County")
call changelog_update("10/17/2019", "Added updated SPEC/MEMO verbiage.", "Ilse Ferris, Hennepin County")
call changelog_update("09/23/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function write_new_line_in_person_note(x)
  EMGetCursor row, col
  If (row = 18 and col + (len(x)) >= 80 + 1 ) or (row = 5 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
  EMReadScreen max_check, 51, 24, 2
  EMSendKey x & "<newline>"
  EMGetCursor row, col
  If (row = 18 and col + (len(x)) >= 80) or (row = 5 and col = 3) then
    EMSendKey "<PF8>"
    EMWaitReady 0, 0
  End if
End function

Function MMIS_panel_check(panel_name)
	Do 
		EMReadScreen panel_check, 4, 1, 52
		If panel_check <> panel_name then Call write_value_and_transmit(panel_name, 1, 8)
	Loop until panel_check = panel_name
End function

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

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
MAXIS_footer_month = CM_mo	'establishing footer month/year 
MAXIS_footer_year = CM_yr 

'column numbers 
cn_col          = 4
PMI_col         = 5
client_name_col = 6
SMI_col         = 7
wstart_col      = 8
wend_col        = 9
medi_col        = 10
one_case_col    = 11
one_type_col    = 12
one_elig_col    = 13
two_case_col    = 14
two_type_col    = 15
two_elig_col    = 16
rlva_col        = 17
dupe_col        = 18
forms_col       = 19
note_col        = 20
one_memo_col    = 21
two_memo_col    = 22

'----------------------------------------------------------------------------------------------------INITIAL DIALOG
'The dialog is defined in the loop as it can change as buttons are pressed 
BeginDialog initial_dialog, 0, 0, 246, 95, "AVS Processing Selection"
  DropListBox 120, 50, 115, 15, "Select one..."+chr(9)+"Initial Monthly Upload"+chr(9)+"ECF Forms Received"+chr(9)+"Person and Case Noting Forms"+chr(9)+"Initial Memo"+chr(9)+"Secondary Memo", AVS_option
  ButtonGroup ButtonPressed
    OkButton 140, 75, 45, 15
    CancelButton 190, 75, 45, 15
  Text 20, 20, 210, 20, "This script should be used when a list of AVS cases are provided by the METS team or DHS."
  Text 20, 50, 95, 10, "Select the processing option:"
  GroupBox 10, 5, 230, 65, "Using this script:"
EndDialog

Do     
    Do
        err_msg = ""
        dialog
        cancel_without_confirmation 
        If AVS_option = "Select one..." then err_msg = "Select the AVS process to complete."
        If err_msg <> "" Then MsgBox err_msg
    Loop until err_msg = ""
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in
    
'----------------------------------------------------------------------------------------------------------------------------------------------------ECF FORMS RECEIVED
If AVS_option = "ECF Forms Received" then 
    'The dialog is defined in the loop as it can change as buttons are pressed 
    BeginDialog , 0, 0, 266, 115, "AVS Forms Procesing"
      ButtonGroup ButtonPressed
        PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
        OkButton 150, 95, 50, 15
        CancelButton 205, 95, 50, 15
      EditBox 15, 50, 180, 15, file_selection_path
      Text 20, 20, 235, 25, "This script should be used if AVS forms have been rec'd for a recipient in ECF."
      Text 15, 70, 230, 15, "Select the Excel file that contains the ECF info by selecting the 'Browse' button, and finding the file."
      GroupBox 10, 5, 250, 85, "Using this script:"
    EndDialog
    
    Do
        err_msg = ""
        dialog
        cancel_without_confirmation 
        If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue." 
        If err_msg <> "" Then MsgBox err_msg
    Loop until err_msg = ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

    excel_row = 2
    entry_record = 0
    
    DIM master_array()
    ReDim master_array(2, 0)
    
    const SMI_ECF_const   = 0
    const scan_date_const = 1
    
    Do 
    	SMI_ECF_number  = ObjExcel.Cells(excel_row, 1).Value
    	SMI_ECF_number  = trim(SMI_ECF_number)
        If SMI_ECF_number = "" then exit do 
        
        scan_date = ObjExcel.Cells(excel_row, 2).Value
        scan_date = trim(scan_date)
        
        ReDim Preserve master_array(2, entry_record)	'This resizes the array based on the number of rows in the Excel File'
        master_array(SMI_ECF_const,	entry_record) = SMI_ECF_number 		
        master_array(scan_date_const, 	entry_record) = scan_date 				
        
        entry_record = entry_record + 1			'This increments to the next entry in the array'
        STATS_counter = STATS_counter + 1
        excel_row = excel_row + 1
    LOOP
    
    objExcel.Quit   'Closes the initial spreadsheet 
    objExcel = ""
    
    file_selection = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\AVS\AVS Forms Distribution Master List.xlsx"
    Call excel_open(file_selection, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    
    'Set objWorkSheet = objWorkbook.Worksheet
    For Each objWorkSheet In objWorkbook.Worksheets
    	If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "All AVS Forms" then months_list = months_list & objWorkSheet.Name & ","
    Next
    months_list = trim(months_list)  'trims excess spaces of months_list
    If right(months_list, 1) = "," THEN months_list = left(months_list, len(months_list) - 1) 'trimming off last comma
    array_of_months = split(months_list, ",")   'Creating new array
        
    '----------------------------------------------------------------------------------------------------FILTERING THE ARRAY 
    
    For each month_sheet in array_of_months
        form_count = 0
        objExcel.worksheets(month_sheet).Activate 'Activates worksheet based on user selection
        excel_row = 2
        
        DO 
            SMI_number = ObjExcel.Cells(excel_row, SMI_col).Value
            SMI_number = trim(SMI_number)
            If SMI_number = "" then exit do 
            
            For item = 0 to UBound(master_array, 2)
                SMI_ECF_number = master_array(SMI_ECF_const, item)  
                scan_date = master_array(scan_date_const, item)
                
                If SMI_ECF_number = SMI_number then 
                    match_found = True 
                    objExcel.Cells(excel_row, forms_col).Value = scan_date
                    objExcel.Cells(excel_row, forms_col).Interior.ColorIndex = 3	'Fills the row with red
                    form_count = form_count + 1
                    exit for
                else 
                    match_found = False 
                end if 
            Next
            excel_row = excel_row + 1
        Loop 
        msgbox "Month: " & month_sheet & vbcr & "Form count: " & form_count
    Next 
    msgbox "Total number of forms reviewed:" & entry_record  
    STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
    script_end_procedure("Success!")
End if 

file_selection = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\AVS\AVS Forms Distribution Master List.xlsx"
Call excel_open(file_selection, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

'----------------------------------------------------------------------------------------------------
If AVS_option = "Initial Monthly Upload" then 
    'adding column header information to the Excel list
    ObjExcel.Cells(1,  7).Value = "SMI"
    ObjExcel.Cells(1,  8).Value = "Waiver start"
    ObjExcel.Cells(1,  9).Value = "Waiver end"
    ObjExcel.Cells(1, 10).Value = "Medicare"
    ObjExcel.Cells(1, 11).Value = "1st case"
    ObjExcel.Cells(1, 12).Value = "1st type/prog"
    ObjExcel.Cells(1, 13).Value = "1st elig dates"
    ObjExcel.Cells(1, 14).Value = "2nd case"
    ObjExcel.Cells(1, 15).Value = "2nd type/prog"
    ObjExcel.Cells(1, 16).Value = "2nd elig dates"
    ObjExcel.Cells(1, 17).Value = "RLVA"
    ObjExcel.Cells(1, 18).Value = "Duplicate PMI?"
    ObjExcel.Cells(1, 19).Value = "Forms Rec'd in ECF"
    ObjExcel.Cells(1, 20).Value = "P/C Note Created"
    ObjExcel.Cells(1, 21).Value = "Initial Memo"
    ObjExcel.Cells(1, 22).Value = "Second Memo"
    
    FOR i = 1 to 22 	'formatting the cells'
    	objExcel.Cells(1, i).Font.Bold = True		'bold font'
    	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
    	objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT
    
    ObjExcel.columns(8).NumberFormat = "mm/dd/yy" 		'formatting waiver start date as a date
    ObjExcel.columns(9).NumberFormat = "mm/dd/yy" 		'formatting waiver end date as a date
    
    DIM case_array()
    ReDim case_array(16, 0)
    
    'constants for array
    const case_number_const     	= 0
    const clt_PMI_const 	        = 1
    const SMI_num_const             = 2
    const waiver_start_const	    = 3
    const waiver_end_const          = 4
    const medicare_info_const       = 5
    const first_case_number_const   = 6
    const first_type_const 	        = 7
    const first_elig_const 	        = 8
    const second_case_number_const  = 9
    const second_type_const         = 10
    const second_elig_const         = 11
    const third_case_number_const   = 12
    const third_type_const     	    = 13
    const third_elig_const          = 14
    const case_status               = 15
    const rlva_coding_const         = 16
    
    'Now the script adds all the clients on the excel list into an array
    excel_row = 2 're-establishing the row to start checking the members for
    entry_record = 0
    Do   
        'Loops until there are no more cases in the Excel list
        
        MAXIS_case_number = objExcel.cells(excel_row, cn_col).Value   'reading the case number from Excel   
        MAXIS_case_number = Trim(MAXIS_case_number)
    
        Client_PMI = objExcel.cells(excel_row, PMI_col).Value          'reading the PMI from Excel 
        Do 
            If left(Client_PMI, 1) = "0" then client_PMI = right(client_PMI, len(client_PMI) -1)
        Loop until left(Client_PMI, 1) <> "0"
        
        Client_PMI = trim(Client_PMI)        
        If Client_PMI = "" then exit do
            
        ReDim Preserve case_array(16, entry_record)	'This resizes the array based on the number of rows in the Excel File'
        case_array(case_number_const,           entry_record) = MAXIS_case_number	'The client information is added to the array'
        case_array(clt_PMI_const,               entry_record) = Client_PMI			
        case_array(SMI_num_const,               entry_record) = ""                       
        case_array(waiver_start_const,	        entry_record) = ""
        case_array(waiver_end_const,	        entry_record) = ""
        case_array(medicare_info_const,         entry_record) = ""     
        case_array(first_case_number_const,   	entry_record) = ""				
        case_array(first_type_const, 	        entry_record) = ""				
        case_array(first_elig_const, 	        entry_record) = ""             
        case_array(second_case_number_const,    entry_record) = ""              
        case_array(second_type_const, 	        entry_record) = ""              
        case_array(second_elig_const, 	        entry_record) = ""              
        case_array(case_status,                 entry_record) = False 	
        case_array(rlva_coding_const,           entry_record) =	""
        
        entry_record = entry_record + 1			'This increments to the next entry in the array'
        stats_counter = stats_counter + 1
        excel_row = excel_row + 1
    Loop
    
    back_to_self
    call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year
    
    excel_row = 2
    For item = 0 to UBound(case_array, 2)
    	MAXIS_case_number = case_array(case_number_const, item)	'Case number is set for each loop as it is used in the FuncLib functions'
        Client_PMI = case_array(clt_PMI_const, item)
    
        Call navigate_to_MAXIS_screen("CASE", "PERS")
        EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
    	If PRIV_check = "PRIV" then
            case_array(case_status, item) = False
    		case_array(SMI_num_const, item) = MAXIS_case_number & " - PRIV case." 
    		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
    		Do
    			back_to_self
    			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
    			If SELF_screen_check <> "SELF" then PF3
    		LOOP until SELF_screen_check = "SELF"
    		EMWriteScreen "________", 18, 43		'clears the MAXIS case number
    		transmit
        Else 
            row = 10
            Do
                EMReadScreen person_PMI, 8, row, 34
                person_PMI = trim(person_PMI)
                IF person_PMI = "" then exit do
                IF Client_PMI = person_PMI then
                    Call write_value_and_transmit("X", row, 59)
                    'Helath care program display pop up 
                    EMReadScreen SMI_num, 9, 7, 50      'Reading the SMI number 
                    Case_array(SMI_num_const, item) = SMI_num
                    Case_array(case_status, item) = True
                    objExcel.Cells(excel_row,  SMI_col).Value = case_array (SMI_num_const, item)
                    SMI_num = ""
                    excel_row = excel_row + 1
                    exit do 
                Else 
                    row = row + 3			'information is 3 rows apart. Will read for the next member. 
                    If row = 19 then
                        PF8  
                        row = 10					'changes MAXIS row if more than one page exists
                    END if
                END if
                EMReadScreen last_PERS_page, 21, 24, 2
            LOOP until last_PERS_page = "THIS IS THE LAST PAGE"
        End if 
    Next 
    
    '-------------------------------------------------------------------------------------------------------------------------------------MMIS portion of the script
    Call navigate_to_MMIS_region("CTY ELIG STAFF/UPDATE")	'function to navigate into MMIS, select the HC realm, and enters the prior autorization area
    
    excel_row = 2
    For item = 0 to UBound(case_array, 2)
        Client_PMI = case_array(clt_PMI_const, item)
        client_PMI = right("00000000" & client_pmi, 8)
        
        If case_array(case_status, item) = True then
            MMIS_panel_check("RKEY") 
            Call clear_line_of_text(5, 19)
            EmWriteScreen Client_PMI, 4, 19
            Call write_value_and_transmit("I", 2, 19)
            
            RSEL_row = 7
            Do 
                EmReadscreen RSEL_panel_check, 4, 1, 52  'RSEL is listed at column 52 
                EmReadscreen panel_check, 4, 1, 51
                If RSEL_panel_check = "RSEL" then
                    EmReadscreen RSEL_SSN, 9, RSEL_row, 48
                    If RSEL_SSN = Client_SSN then
                        duplicate_entry = True 
                        Call write_value_and_transmit("X", RSEL_row, 2)
                        EmReadscreen panel_check, 4, 1, 51
                    else 
                        Exit do
                        duplicate_entry = False 
                    End if 
                End if     
                
                If panel_check = "RSUM" then 
                    'Waiver info
                    EmReadscreen waiver_info, 39, 15, 15
                    waiver_info = trim(waiver_info)
                    If waiver_info = "BEG DT:          THROUGH DT:" then 
                        waiver_info = ""
                        Case_array(waiver_start_const, item) = ""
                        Case_array(waiver_end_const, item) = ""
                    Else 
                        EMReadscreen waiver_start_date, 8, 15, 25
                        EmReadscreen waiver_end_date, 8, 15, 46
                        Case_array(waiver_start_const, item) = waiver_start_date
                        Case_array(waiver_end_const, item) = waiver_end_date
                    End if 
                    
                    'Medicare info
                    EmReadscreen medicare_info, 69, 21, 10
                    medicare_info = trim(medicare_info)
                    IF medicare_info = "PART A BEG:          END:          PART B BEG:          END:" then medicare_info = ""
                    Case_array(medicare_info_const, item) = medicare_info
                    
                    '1st case type/prog/elig/case number 
                    EmReadscreen first_case_number, 8, 7, 16
                    first_case_number = trim(first_case_number)
                    If first_case_number <> "" then 
                        case_array(first_case_number_const, item) = first_case_number
                        EmReadscreen first_program, 2, 6, 13
                        EmReadscreen first_type, 2, 6, 35
                        If trim(first_program) <> "" then 
                            first_elig_type = first_program & "-" & first_type
                            case_array(first_type_const, item) = first_elig_type
                            '1st elig dates 
                            EmReadscreen first_elig_start, 8, 7, 35
                            EmReadscreen first_elig_end, 8, 7, 54
                            first_elig_dates = first_elig_start &  " - " & first_elig_end
                            case_array(first_elig_const, item) = first_elig_dates
                        ENd if    
                    End if 
                
                    EmReadscreen second_case_number, 8, 9, 16
                    second_case_number = trim(second_case_number)
                    If second_case_number <> "" then 
                        case_array(second_case_number_const, item) = second_case_number
                        EmReadscreen second_program, 2, 8, 13
                        EmReadscreen second_type, 2, 8, 35
                        If trim(second_program) <> "" then 
                            second_elig_type = second_program & "-" & second_type
                            case_array(second_type_const, item) = second_elig_type
                            '1st elig dates 
                            EmReadscreen second_elig_start, 8, 9, 35
                            EmReadscreen second_elig_end, 8, 9, 54
                            second_elig_dates = second_elig_start &  " - " & second_elig_end
                            case_array(second_elig_const, item) = second_elig_dates
                        ENd if    
                    End if     
                    
                    'RLVA 
                    Call write_value_and_transmit("RLVA", 1, 8)
                    Call MMIS_panel_check("RLVA")
                    EmReadscreen rlva_coding, 12, 14, 42 'most recent living arrangement 
                    case_array(rlva_coding_const, item) = rlva_coding
                    
                    'outputting to Excel 
                    objExcel.Cells(excel_row, SMI_col).Value = case_array (SMI_num_const,                  item)
                    objExcel.Cells(excel_row, wstart_col).Value = case_array (waiver_start_const,	       item)
                    objExcel.Cells(excel_row, wend_col).Value = case_array (waiver_end_const,	           item)
                    objExcel.Cells(excel_row, medi_col).Value = case_array (medicare_info_const,           item)
                    objExcel.Cells(excel_row, one_case_col).Value = case_array (first_case_number_const,   item)
                    objExcel.Cells(excel_row, one_type_col).Value = case_array (first_type_const, 	       item)
                    objExcel.Cells(excel_row, one_elig_col).Value = case_array (first_elig_const, 	       item)
                    objExcel.Cells(excel_row, two_case_col).Value = case_array (second_case_number_const,  item)
                    objExcel.Cells(excel_row, two_type_col).Value = case_array (second_type_const, 	       item)
                    objExcel.Cells(excel_row, two_elig_col).Value = case_array (second_elig_const, 	       item)
                    objExcel.Cells(excel_row, rlva_col).Value = case_array (rlva_coding_const,             item)                     
                    
                    If duplicate_entry = True then objExcel.Cells(excel_row, dupe_col).Value = "True"
                    PF3
                    exit do 
                End if 
            loop 
        else 
            objExcel.Cells(excel_row, dupe_col).Value = "Error case" 
        End if
        excel_row = excel_row + 1 
    Next     
End if 

If AVS_option = "Person and Case Noting Forms" then 
    For Each objWorkSheet In objWorkbook.Worksheets
        If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "All AVS Forms" then months_list = months_list & objWorkSheet.Name & ","
    Next
    months_list = trim(months_list)  'trims excess spaces of months_list
    If right(months_list, 1) = "," THEN months_list = left(months_list, len(months_list) - 1) 'trimming off last comma
    array_of_months = split(months_list, ",")   'Creating new array
    
    back_to_self
    call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year
        
    For each month_sheet in array_of_months
        form_count = 0
        objExcel.worksheets(month_sheet).Activate 'Activates worksheet based on user selection
        'Now the script adds all the clients on the excel list into an array
        excel_row = 2 're-establishing the row to start checking the members for
        entry_record = 0
        case_note_total = 0
        Do   
            'Loops until there are no more cases in the Excel list
            MAXIS_case_number = objExcel.cells(excel_row, cn_col).Value   'reading the case number from Excel   
            MAXIS_case_number = Trim(MAXIS_case_number)
            
            client_PMI = objExcel.cells(excel_row, PMI_col).Value
            client_PMI = trim(client_PMI) 
            
            client_name = objExcel.cells(excel_row, client_name_col).Value
            client_name = trim(client_name)
            Call fix_case(client_name, 2)
            client_name = trim(client_name)
            
            form_date = objExcel.cells(excel_row, forms_col).Value
            form_date = trim(form_date)
            
            note_date = objExcel.cells(excel_row, note_col).Value
            note_date = trim(note_date)
            
            If client_PMI = "" then exit do
            stats_counter = stats_counter + 1
        
            'Skipping cases that do not have a form date already listed or already have a case/person note.
            If trim(form_date) <> "" then
                If trim(note_date) = "" then
                    Call navigate_to_MAXIS_screen("STAT", "MEMB")
                    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
                    If PRIV_check = "PRIV" then
                         objExcel.cells(excel_row, note_col).Value = "PRIVILEGED" 
                        'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
                        Do
                            back_to_self
                            EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
                            If SELF_screen_check <> "SELF" then PF3
                        LOOP until SELF_screen_check = "SELF"
                        EMWriteScreen "________", 18, 43		'clears the MAXIS case number
                        transmit  
                    Else
                        EmReadscreen county_check, 4, 21, 21
                        If county_check <> "X127" then 
                            objExcel.cells(excel_row, note_col).Value = "OUT OF COUNTY"
                        Else 
                            Do 
                                EMReadScreen member_PMI, 7, 4, 46
                                If trim(member_PMI) = Client_PMI then 
                                    Found_member = True 
                                    exit do 
                                Else
                                    Found_member = False 
                                    transmit
                                    EMReadScreen MEMB_error, 5, 24, 2
                                End if 
                            Loop until MEMB_error = "ENTER"
                        
                            If Found_member = True then
                                case_note_total = case_note_total + 1
                                note_header = "AVS Auth Form Rec'd " & form_date & " - " & client_name
                                note_body = "The DHS-7823 form (Authorization to Obtain Financial Information from the Account Validation Service - AVS) has not been reviewed for accuracy for this recipient. Review of the AVS form will be completed by HSR's at a later date."
                                '---------------------------------------------------------------Creating the PERSON Note 
                                PF5
                                EMReadScreen PNOTE_check, 4, 2, 46
                                If PNOTE_check <> "SCRN" then 
                                     objExcel.cells(excel_row, note_col).Value = "PERS note issue"
                                ELSE
                                    EMreadscreen edit_mode_required_check, 6, 5, 3		'if not person not exists, person note goes directly into edit mode
                                    If edit_mode_required_check <> "      " then PF9
                                    write_new_line_in_person_note(note_header)
                                    write_new_line_in_person_note("--")
                                    write_new_line_in_person_note(note_body)
                                END IF 	
                                PF3 'to save and exit person notes
                                '---------------------------------------------------------------Creating the CASE note  
                                start_a_blank_CASE_NOTE
                                Call write_variable_in_CASE_NOTE(note_header)	
                                Call write_variable_in_CASE_NOTE("--")
                                Call write_variable_in_CASE_NOTE(note_body)
                                PF3 'to save and exit case notes 
                                objExcel.cells(excel_row, note_col).Value = date
                                objExcel.Cells(excel_row, note_col).Interior.ColorIndex = 3	'Fills the row with red 
                            End if
                        End if 
                    End if     
                End if     
            End if 
            excel_row = excel_row + 1
            MAXIS_case_number = "" 
            client_PMI = ""
            client_name = ""
            form_date = ""
            note_date = ""
        Loop 
        msgbox "Month: " & month_sheet & vbcr & "Case note total: " & case_note_total
    Next 
End if 

If instr(AVS_option, "Memo") then 
    msgbox "Untested Coded. Waiting for AVS work group for go ahead."
'    Do  
'        'Establishing values for each case in the array of cases 
'        MAXIS_case_number = objExcel.cells(excel_row, cn_col).Value   'reading the case number from Excel   
'        MAXIS_case_number = Trim(MAXIS_case_number)
'        
'        client_name = objExcel.cells(excel_row, client_name_col).Value
'        client_name = trim(client_name)
'        Call fix_case(client_name, 2)
'        client_name = trim(client_name)
'        
'        If instr(client_name, " ") then    						'Most cases have both last name and 1st name. This seperates the two names
'            length = len(client_name)                           'establishing the length of the variable
'            position = InStr(client_name, " ")                  'sets the position at the deliminator (in this case the comma)    
'            first_name = left(client_name, length-position)    'establishes client first name as after before the deliminator
'        END IF
'        'adding first name to name list 
'        first_name = trim(first_name)
'        'Call fix_case(first_name, 0)
'        msgbox first_name 
'        
'        form_date = objExcel.cells(excel_row, forms_col).Value
'        form_date = trim(form_date)
'        
'        If AVS_option = "Initial Memo" then
'            excel_col = one_memo_col
'            first_line = client_name & " will soon receive a letter from the Department of Human Services with a form called Authorization to Obtain Financial Information from the AVS (Asset Validation System)."
'            second_line = "AVS will provide Hennepin County with information on your accounts, such as checking, savings accounts, and money market accounts. If you are married or a non-citizen with a sponsor, then it will provide information on your spouse’s, sponsor(s)’, and sponsor(s)’ spouse(s) accounts."
'            third_line = "You will receive this letter because we need your permission to access account information through the AVS for your Medical Assistance eligibility. We also may need permission to access account information for your spouse, sponsor(s), or sponsor(s)’ spouse(s). (If you are a US citizen, you do not have a sponsor)."
'        elseif AVS_option = "Secondary Memo" then 
'            excel_col = two_memo_col
'            first_line =  client_name & " received a letter from the Department of Human Services with a form called Authorization to Obtain Financial Information from the AVS (Asset Validation System) in the mail."
'            second_line = "AVS will provide Hennepin County information on your account, such as checking, savings, and money market accounts."
'            third_line = "You received this letter because we need your permission to access account information through the AVS for your Medical Assistance eligibility. We also may need permission to access account information for your spouse, sponsor(s), or sponsor(s)’ spouse(s). (If you are a US citizen, you do not have a sponsor)."
'        End if
'        
'        memo_date = objExcel.cells(excel_row, excel_col).Value
'        memo_date = trim(memo_date)
'                
'        If MAXIS_case_number = "" then exit do 
'        If form_date = "" then 
'            If memo_date = "" then     
'                Call start_a_new_spec_memo_and_continue
'                If success_var = False then 
'                    objExcel.cells(excel_row, excel_col).Value = "FALSE"
'                Else 
'                    '----------------------------------------------------------------------------------------------------THE SPEC/MEMO
'                    Call start_a_new_spec_memo            
'                    'Writes the MEMO.
'                    call write_variable_in_SPEC_MEMO(first_line)
'                    Call write_variable_in_SPEC_MEMO("")
'                    Call write_variable_in_SPEC_MEMO(second_line)
'                    Call write_variable_in_SPEC_MEMO("")
'                    Call write_variable_in_SPEC_MEMO(third_line)
'                    Call write_variable_in_SPEC_MEMO("")
'                    Call write_variable_in_SPEC_MEMO("You must return the signed form for us to determine your eligibility for certain health care programs. If you do not return the form by the due date on the letter from the Department of Human Services, your Medical Assistance and/or Medicare Savings Program may close.")
'                    Call write_variable_in_SPEC_MEMO("")	
'                    Call write_variable_in_SPEC_MEMO("Who must sign the form?")
'                    Call write_variable_in_SPEC_MEMO("-	" & first_name & " or Authorized Representative")
'                    Call write_variable_in_SPEC_MEMO("-	If you are married, your spouse must also sign the form")
'                    Call write_variable_in_SPEC_MEMO("-	If you are a Lawful Permanent Resident sponsored under an Affidavit of Support (USCIS I-864), your sponsor(s) and sponsor(s)’s spouses must also sign")
'                    Call write_variable_in_SPEC_MEMO("")
'                    Call write_variable_in_SPEC_MEMO("How to return the signed form:")
'                    Call write_variable_in_SPEC_MEMO(" ")
'                    Call write_variable_in_SPEC_MEMO("By mail: Hennepin County Human Service Dept.")
'                    Call write_variable_in_SPEC_MEMO("PO BOX 107 Minneapolis, MN 55440")
'                    Call write_variable_in_SPEC_MEMO(" ")
'                    Call write_variable_in_SPEC_MEMO("By fax: 612-288-2981")
'                    Call write_variable_in_SPEC_MEMO(" ")
'                    Call write_variable_in_SPEC_MEMO("In person: You can drop off the form at any of our regional offices:")
'                    Call write_variable_in_SPEC_MEMO("- Central/Northeast Minneapolis: 525 Portland Ave S Minneapolis 55415")
'                    Call write_variable_in_SPEC_MEMO("- North Minneapolis: 1001 Plymouth Ave N Minneapolis 55411")
'                    Call write_variable_in_SPEC_MEMO("- Northwest Suburban: 7051 Brooklyn Blvd Brooklyn Center 55429")
'                    Call write_variable_in_SPEC_MEMO("- South Minneapolis: 2215 East Lake Street Minneapolis 55407")
'                    Call write_variable_in_SPEC_MEMO("- South Suburban: 9600 Aldrich Ave S Bloomington 55420 Th hrs: 8:30-6:30")
'                    Call write_variable_in_SPEC_MEMO("- West Suburban: 1011 1st St S Hopkins 55343")   
'                    PF4			'Exits the MEMO
'                    EMReadScreen memo_sent, 8, 24, 2
'                    If memo_sent <> "NEW MEMO" then 
'                        objExcel.cells(excel_row, excel_col).Value = "DID NOT CREATE MEMO"
'                        created_memo = False 
'                    Else 
'                        objExcel.cells(excel_row, excel_col).Value = date 
'                        created_memo = true 
'                        back_to_self
'                        ''----------------------------------------------------------------------------------------------------THE CASE NOTE
'                        'start_a_blank_CASE_NOTE
'                        'Call write_variable_in_CASE_NOTE("Sent SPEC/MEMO re: COLA income/deduction verifs needed")
'                        'Call write_variable_in_CASE_NOTE("Set TIKL for " & tikl_date & " to review case.") 
'                        'Call write_variable_in_CASE_NOTE("If verifications have not been provided, a verification request in ECF will need to be sent.")
'                        'call write_variable_in_case_note("---")
'                        'call write_variable_in_case_note(worker_signature)
'                        'PF3
'                    End if  
'                End if 
'            End if 
'        End if 
'        Excel_row = excel_row + 1
'        stats_counter = stats_counter + 1	
'    Loop 
End if 
    
FOR i = 1 to 22		'formatting the cells
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT
    
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created. Please review for cases that need to be processed manually.")