'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - AVS FORMS NOTING.vbs"
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
call changelog_update("11/13/2019", "Initial version.", "Ilse Ferris, Hennepin County")

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

'----------------------------------------------------------------------------------------------------The script  
'The dialog is defined in the loop as it can change as buttons are pressed 
BeginDialog dialog1, 0, 0, 266, 115, "AVS Forms Procesing"
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
    dialog1
    cancel_without_confirmation 
    If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
    If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue." 
    If err_msg <> "" Then MsgBox err_msg
Loop until err_msg = ""
If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

excel_row = 2
entry_record = 0

DIM master_array()
ReDim master_array(7, 0)

const SMI_ECF_const      = 0
const scan_date_const    = 1
const case_number        = 2
const PMI_const          = 3
const client_name_const  = 4
const note_created_const = 5    
const add_record_const   = 6    
const match_found_const  = 7        

Do 
	SMI_ECF_number  = ObjExcel.Cells(excel_row, 1).Value
	SMI_ECF_number  = trim(SMI_ECF_number)
    If SMI_ECF_number = "" then exit do 
    
    scan_date = ObjExcel.Cells(excel_row, 2).Value
    scan_date = trim(scan_date)
    
    ReDim Preserve master_array(7, entry_record)	'This resizes the array based on the number of rows in the Excel File'
    master_array(SMI_ECF_const,	        entry_record) = SMI_ECF_number 		
    master_array(scan_date_const, 	    entry_record) = scan_date 	
    master_array(case_number, 	        entry_record) = "" 	
    master_array(PMI_const, 	        entry_record) = "" 				
    master_array(client_name_const, 	entry_record) = "" 	
    master_array(note_created_const,    entry_record) = ""
    master_array(add_record_const,      entry_record) = ""
    master_array(match_found_const,     entry_record) = ""
	
    entry_record = entry_record + 1			'This increments to the next entry in the array'
    STATS_counter = STATS_counter + 1
    excel_row = excel_row + 1
LOOP

objExcel.Quit   'Closes the initial spreadsheet 
objExcel = ""

file_selection = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\AVS\AVS Forms Distribution Master List.xlsx"
Call excel_open(file_selection, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

objExcel.worksheets("All AVS Forms").Activate 'Activates worksheet based on user selection
excel_row = 2

DO 
    master_SMI_number = ObjExcel.Cells(excel_row, 1).Value  'from All AVS forms list
    master_SMI_number = trim(master_SMI_number)
    If master_SMI_number = "" then exit do 
    
    For item = 0 to UBound(master_array, 2)
        SMI_ECF_number = master_array(SMI_ECF_const, item)  
        If SMI_ECF_number = master_SMI_number then 
            master_array(add_record_const, item) = True 
        else 
            master_array(add_record_const, item) = False 
        end if 
    Next
    excel_row = excel_row + 1
Loop 

For item = 0 to UBound(master_array, 2)
    If master_array(add_record_const, item) = False then 
        ObjExcel.Cells(excel_row, 1).Value = master_array(SMI_ECF_const, item)
        ObjExcel.Cells(excel_row, 2).Value = master_array(scan_date_const,  item)
        objExcel.Cells(excel_row, 1).Interior.ColorIndex = 3	'Fills the row with red
        objExcel.Cells(excel_row, 2).Interior.ColorIndex = 3	'Fills the row with red
        excel_row = excel_row + 1
    End if 
Next 

script_end_procedure ("Yep")

''Set objWorkSheet = objWorkbook.Worksheet
'For Each objWorkSheet In objWorkbook.Worksheets
'	If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "All AVS Forms" then months_list = months_list & objWorkSheet.Name & ","
'Next
'months_list = trim(months_list)  'trims excess spaces of months_list
'If right(months_list, 1) = "," THEN months_list = left(months_list, len(months_list) - 1) 'trimming off last comma
'array_of_months = split(months_list, ",")   'Creating new array
'    
''----------------------------------------------------------------------------------------------------FILTERING THE ARRAY 
'For each month_sheet in array_of_months
'    objExcel.worksheets(month_sheet).Activate 'Activates worksheet based on user selection
'    excel_row = 2
'    
'    DO 
'        SMI_number = ObjExcel.Cells(excel_row, 7).Value
'        SMI_number = trim(SMI_number)
'        If SMI_number = "" then exit do 
'        
'        For item = 0 to UBound(master_array, 2)
'            SMI_ECF_number = master_array(SMI_ECF_const, item)  
'                        
'            If SMI_ECF_number = SMI_number then 
'                match_found = True 
'                objExcel.Cells(excel_row,  4).Value = master_array(case_number,        item)
'                objExcel.Cells(excel_row,  5).Value = master_array(PMI_const, 	       item) 			
'                objExcel.Cells(excel_row,  6).Value = master_array(client_name_const,  item) 
'                objExcel.Cells(excel_row, 19).Value = master_array(note_created_const, item)
'                master_array(note_created_const, item) = True 
'                form_count = form_count + 1
'                exit do
'            else 
'                master_array(note_created_const, item) = True 
'            end if 
'        Next
'        excel_row = excel_row + 1 
'    Loop 
'    msgbox "Month: " & month_sheet & vbcr & "Form count: " & form_count
'Next 
'msgbox "Total number of forms reviewed:" & entry_record  
'STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
'
'
'For item = 0 to UBound(new_cases_array, 2)
'	objExcel.Cells(excel_row, 1).Value = new_cases_array(case_number_const, item)	
'	objExcel.Cells(excel_row, 2).Value = new_cases_array(member_number_const, item)	
'	objExcel.Cells(excel_row, 3).Value = new_cases_array(client_name_const, item)	
'	excel_row = excel_row + 1
'Next
'script_end_procedure("Success!")
'
'DIM case_array()
'ReDim case_array(15, 0)
'
''constants for array
'const case_number_const     	= 0
'const clt_PMI_const 	        = 1
'const SMI_num_const             = 2
'const waiver_info_const	        = 3
'const medicare_info_const       = 4
'const first_case_number_const   = 5
'const first_type_const 	        = 6
'const first_elig_const 	        = 7
'const second_case_number_const  = 8
'const second_type_const         = 9
'const second_elig_const         = 10
'const third_case_number_const   = 11
'const third_type_const     	    = 12
'const third_elig_const          = 13
'const case_status               = 14
'const rlva_coding_const         = 15
'
''Now the script adds all the clients on the excel list into an array
'excel_row = 2 're-establishing the row to start checking the members for
'entry_record = 0
'Do   
'    'Loops until there are no more cases in the Excel list
'    
'    MAXIS_case_number = objExcel.cells(excel_row, 4).Value   'reading the case number from Excel   
'    MAXIS_case_number = Trim(MAXIS_case_number)
'
'    Client_PMI = objExcel.cells(excel_row, 5).Value          'reading the PMI from Excel 
'    Do 
'        If left(Client_PMI, 1) = "0" then client_PMI = right(client_PMI, len(client_PMI) -1)
'    Loop until left(Client_PMI, 1) <> "0"
'    
'    Client_PMI = trim(Client_PMI)        
'    If Client_PMI = "" then exit do
'        
'    ReDim Preserve case_array(15, entry_record)	'This resizes the array based on the number of rows in the Excel File'
'    case_array(case_number_const,           entry_record) = MAXIS_case_number	'The client information is added to the array'
'    case_array(clt_PMI_const,               entry_record) = Client_PMI			
'    case_array(SMI_num_const,               entry_record) = ""                       
'    case_array(waiver_info_const,	        entry_record) = ""
'    case_array(medicare_info_const,         entry_record) = ""     
'    case_array(first_case_number_const,   	entry_record) = ""				
'    case_array(first_type_const, 	        entry_record) = ""				
'    case_array(first_elig_const, 	        entry_record) = ""             
'    case_array(second_case_number_const,    entry_record) = ""              
'    case_array(second_type_const, 	        entry_record) = ""              
'    case_array(second_elig_const, 	        entry_record) = ""              
'    case_array(case_status,                 entry_record) = False 	
'    case_array(rlva_coding_const,           entry_record) =	""
'    
'    entry_record = entry_record + 1			'This increments to the next entry in the array'
'    stats_counter = stats_counter + 1
'    excel_row = excel_row + 1
'Loop
'
'back_to_self
'call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year
'
'excel_row = 2
'For item = 0 to UBound(case_array, 2)
'    MAXIS_case_number = case_array(case_number_const, item)	'Case number is set for each loop as it is used in the FuncLib functions'
'    Client_PMI = case_array(clt_PMI_const, item)
'
'    Call navigate_to_MAXIS_screen("CASE", "PERS")
'    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
'    If PRIV_check = "PRIV" then
'        case_array(case_status, item) = False
'        case_array(SMI_num_const, item) = MAXIS_case_number & " - PRIV case." 
'        'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
'        Do
'            back_to_self
'            EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
'            If SELF_screen_check <> "SELF" then PF3
'        LOOP until SELF_screen_check = "SELF"
'        EMWriteScreen "________", 18, 43		'clears the MAXIS case number
'        transmit
'    Else 
'        row = 10
'        Do
'            EMReadScreen person_PMI, 8, row, 34
'            person_PMI = trim(person_PMI)
'            IF person_PMI = "" then exit do
'            IF Client_PMI = person_PMI then
'                Call write_value_and_transmit("X", row, 59)
'                'Helath care program display pop up 
'                EMReadScreen SMI_num, 9, 7, 50      'Reading the SMI number 
'                Case_array(SMI_num_const, item) = SMI_num
'                Case_array(case_status, item) = True
'                objExcel.Cells(excel_row,  7).Value = case_array (SMI_num_const, item)
'                excel_row = excel_row + 1
'                exit do 
'            Else 
'                row = row + 3			'information is 3 rows apart. Will read for the next member. 
'                If row = 19 then
'                    PF8  
'                    row = 10					'changes MAXIS row if more than one page exists
'                END if
'            END if
'            EMReadScreen last_PERS_page, 21, 24, 2
'        LOOP until last_PERS_page = "THIS IS THE LAST PAGE"
'    End if 
'Next 
'
'
'
'
'If AVS_option = "Person and Case Noting Forms" then 
'    'For Each objWorkSheet In objWorkbook.Worksheets
'    '    If instr(objWorkSheet.Name, "Sheet") = 0 then months_list = months_list & objWorkSheet.Name & ","
'    'Next
'    'months_list = trim(months_list)  'trims excess spaces of months_list
'    'If right(months_list, 1) = "," THEN months_list = left(months_list, len(months_list) - 1) 'trimming off last comma
'    'array_of_months = split(months_list, ",")   'Creating new array
'    
'    back_to_self
'    call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year
'        
'    'For each month_sheet in array_of_months
'        start_date = #09/01/19#
'        form_count = 0
'        'objExcel.worksheets(month_sheet).Activate 'Activates worksheet based on user selection
'        'msgbox month_sheet
'        'Now the script adds all the clients on the excel list into an array
'        excel_row = 2 're-establishing the row to start checking the members for
'        entry_record = 0
'        case_note_total = 0
'        Do   
'            'Loops until there are no more cases in the Excel list
'            MAXIS_case_number = objExcel.cells(excel_row, 4).Value   'reading the case number from Excel   
'            MAXIS_case_number = Trim(MAXIS_case_number)
'            
'            client_PMI = objExcel.cells(excel_row, 5).Value
'            client_PMI = trim(client_PMI) 
'            
'            client_name = objExcel.cells(excel_row, 6).Value
'            client_name = trim(client_name)
'            Call fix_case(client_name, 2)
'            client_name = trim(client_name)
'            
'            form_date = objExcel.cells(excel_row, 18).Value
'            form_date = trim(form_date)
'            
'            note_date = objExcel.cells(excel_row, 19).Value
'            note_date = trim(note_date)
'            
'            If client_PMI = "" then exit do
'            stats_counter = stats_counter + 1
'        
'            'Skipping cases that do not have a form date already listed or already have a case/person note.
'            If trim(form_date) <> "" then
'                Call navigate_to_MAXIS_screen("CASE", "NOTE")
'                'starting at the 1st case note, checking the headers for the NOTES - EXPEDITED SCREENING text or the NOTES - EXPEDITED DETERMINATION text
'        		MAXIS_row = 5
'                Case_note = False 
'        		Do
'        			EMReadScreen case_note_date, 8, MAXIS_row, 6
'                    If start_date > cdate(case_note_date) then exit do 
'        			If trim(case_note_date) = "" then
'        				MAXIS_row = MAXIS_row + 1
'        			else 
'        				EMReadScreen case_note_header, 55, MAXIS_row, 25
'        				case_note_header = trim(case_note_header)
'        				IF instr(case_note_header, "AVS Auth Form Rec'd") then
'                            
'                            length = len(case_note_header)                           'establishing the length of the variable
'                            position = InStr(case_note_header, " - ")                  'sets the position at the deliminator (in this case the comma)
'                            CN_name = Right(client_name, length-position)    'establishes client first name as after before the deliminator
'                            'msgbox CN_name
'                            If CN_name = client_name then 
'                                case_note_found = True
'                                objExcel.cells(excel_row, 19).Value = case_note_date 
'                                objExcel.Cells(excel_row, 19).Interior.ColorIndex = 3	'Fills the row with red
'                                exit do 
'                            Else 
'                                case_note_found = False 
'                            End if 
'                        End if 
'        			END IF
'        			MAXIS_row = MAXIS_row + 1
'                    If MAXIS_row = 19 then 
'                        PF8 
'                        MAXIS_row = 5
'                    End if 
'                    'TODO Add output of the client name for names that are too long for the header 
'        		LOOP until cdate(case_note_date) < start_date                       'repeats until the case note date is less than the application date
'                            
'                'If trim(note_date) = "" then
'                '    Call navigate_to_MAXIS_screen("STAT", "MEMB")
'                '    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
'                '    
'                '    If PRIV_check = "PRIV" then
'                '         objExcel.cells(excel_row, 19).Value = "PRIVILEGED" 
'                '        'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
'                '        Do
'                '            back_to_self
'                '            EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
'                '            If SELF_screen_check <> "SELF" then PF3
'                '        LOOP until SELF_screen_check = "SELF"
'                '        EMWriteScreen "________", 18, 43		'clears the MAXIS case number
'                '        transmit  
'                '    Else
'                '        EmReadscreen county_check, 4, 21, 21
'                '        If county_check <> "X127" then 
'                '            objExcel.cells(excel_row, 19).Value = "OUT OF COUNTY"
'                '        Else 
'                '            Do 
'                '                EMReadScreen member_PMI, 7, 4, 46
'                '                If trim(member_PMI) = Client_PMI then 
'                '                    Found_member = True 
'                '                    
'                '                    exit do 
'                '                Else
'                '                    Found_member = False 
'                '                    transmit
'                '                    EMReadScreen MEMB_error, 5, 24, 2
'                '                End if 
'                '            Loop until MEMB_error = "ENTER"
'                '        
'                '            If Found_member = True then
'                '                case_note_total = case_note_total + 1
'                '                note_header = "AVS Auth Form Rec'd " & form_date & " - " & client_name
'                '                note_body = "The DHS-7823 form (Authorization to Obtain Financial Information from the Account Validation Service - AVS) has not been reviewed for accuracy for this recipient. Review of the AVS form will be completed by HSR's at a later date."
'                '                '---------------------------------------------------------------Creating the PERSON Note 
'                '                PF5
'                '                EMReadScreen PNOTE_check, 4, 2, 46
'                '                If PNOTE_check <> "SCRN" then 
'                '                     objExcel.cells(excel_row, 19).Value = "PERS note issue"
'                '                ELSE
'                '                    EMreadscreen edit_mode_required_check, 6, 5, 3		'if not person not exists, person note goes directly into edit mode
'                '                    If edit_mode_required_check <> "      " then PF9
'                '                    write_new_line_in_person_note(note_header)
'                '                    write_new_line_in_person_note("--")
'                '                    write_new_line_in_person_note(note_body)
'                '                END IF 	
'                '                PF3 'to save and exit person notes
'                '                '---------------------------------------------------------------Creating the CASE note  
'                '                start_a_blank_CASE_NOTE
'                '                Call write_variable_in_CASE_NOTE(note_header)	
'                '                Call write_variable_in_CASE_NOTE("--")
'                '                Call write_variable_in_CASE_NOTE(note_body)
'                '                PF3 'to save and exit case notes 
'                '                objExcel.cells(excel_row, 19).Value = date 
'                '            End if
'                '        End if 
'                '    End if     
'                'End if     
'            End if 
'            excel_row = excel_row + 1
'            MAXIS_case_number = "" 
'            client_PMI = ""
'            client_name = ""
'            form_date = ""
'            note_date = ""
'        Loop 
'        msgbox "Case note total: " & case_note_total
'    'Next 
'End if 
'
'script_end_procedure("Complete!")