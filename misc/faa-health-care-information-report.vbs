'Required for statistical purposes===============================================================================
name_of_script = "FAA- HEALTH CARE INFORMATION REPORT.vbs"
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
call changelog_update("12/29/2020", "Added PMAP information to report. Added status to report. Removed error list. Status cases may also have health care information, enhancement from previous error list.", "Ilse Ferris, Hennepin County")
call changelog_update("10/20/2020", "Added link to instructions in main dialog.", "Ilse Ferris, Hennepin County")
call changelog_update("08/06/2020", "Final release version ready for production.", "Ilse Ferris, Hennepin County")
call changelog_update("07/16/2020", "Added gender and DOB fields to report.", "Ilse Ferris, Hennepin County")
call changelog_update("09/13/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
Call check_for_MMIS(True) 'ensuring we're in MMIS 

'The dialog is defined in the loop as it can change as buttons are pressed 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 221, 115, "Health Care Information Report"
  ButtonGroup ButtonPressed
    PushButton 170, 45, 40, 15, "Browse...", select_a_file_button
    PushButton 45, 95, 80, 15, "Script Instructions", help_button
    OkButton 130, 95, 40, 15
    CancelButton 175, 95, 40, 15
  Text 20, 20, 190, 20, "This script should be used when a list of PMI's that Health Care Information from MMIS is needed and provided."
  Text 15, 65, 195, 15, "Select the Excel file that contains the list of PMI's by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 205, 85, "Using this script:"
  EditBox 15, 45, 150, 15, file_selection_path
EndDialog

'dialog and dialog DO...Loop	
Do
    Do 
        'Initial Dialog to determine the excel file to use, column with case numbers, and which process should be run
        'Show initial dialog
        Do
        	Dialog Dialog1 
        	cancel_without_confirmation
            If ButtonPressed = help_button then open_URL_in_browser("https://hennepin.sharepoint.com/teams/hs-economic-supports-hub/BlueZone_Script_Instructions/General%20and%20Organizational%20Documents/Health%20Care%20Information%20Report%20Instructions.docx?d=w66570f8c377544eb973f334f3210fbed&csf=1&web=1&e=bRpmfV")
        	If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        Loop until ButtonPressed = -1
        err_msg = ""
        If trim(file_selection_path) = "" then err_msg = err_msg & "Select the file of PMI numbers. Press the BROWSE button to search the file explorer for the file."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbNewLine & err_msg & vbNewLine
    LOOP until err_msg = ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

DIM case_array()
ReDim case_array(19, 0)

'constants for array
const clt_PMI_const 	        = 0
const last_name_const           = 1
const first_name_const          = 2
const client_SSN_const          = 3
const DOB_const                 = 4          
const gender_const              = 5
const first_case_number_const   = 6
const first_type_const 	        = 7
const first_elig_const 	        = 8
const second_case_number_const 	= 9
const second_type_const 	    = 10
const second_elig_const 	    = 11
const third_case_number_const 	= 12
const third_type_const      	= 13
const third_elig_const      	= 14
const case_status               = 15        	
const rsum_PMI_const            = 16
const pmap_begin_const          = 17
const pmap_end_const            = 18
const pmap_name_const           = 19

'Now the script adds all the clients on the excel list into an array
excel_row = 2 're-establishing the row to start checking the members for
entry_record = 0
all_pmi_array = "*"    'setting up string to find duplicate case numbers 
Do   
    'Loops until there are no more cases in the Excel list
    Client_PMI = objExcel.cells(excel_row, 1).Value          'reading the PMI from Excel 
    Client_PMI = trim(Client_PMI)
    If Client_PMI = "" then exit do
    client_PMI = right("00000000" & Client_PMI, 8)

    'If the case number is found in the string of case numbers, it's not added again. 
    If instr(all_pmi_array, "*" & Client_PMI & "*") then 
        add_to_array = False
    Else 
        ReDim Preserve case_array(19, entry_record)	'This resizes the array based on the number of rows in the Excel File'
        'The client information is added to the array'
        case_array(clt_PMI_const,               entry_record) = Client_PMI			
        case_array(last_name_const,             entry_record) = ""             
        case_array(first_name_const,            entry_record) = ""   
        case_array(client_SSN_const,            entry_record) = ""                  
        case_array(DOB_const,                   entry_record) = ""  
        case_array(gender_const,                entry_record) = ""
        case_array(first_case_number_const,   	entry_record) = ""				
        case_array(first_type_const, 	        entry_record) = ""				
        case_array(first_elig_const, 	        entry_record) = ""             
        case_array(second_case_number_const,    entry_record) = ""              
        case_array(second_type_const, 	        entry_record) = ""              
        case_array(second_elig_const, 	        entry_record) = ""              
        case_array(third_case_number_const, 	entry_record) = ""
        case_array(third_type_const,      	    entry_record) = ""				
        case_array(third_elig_const,            entry_record) = ""	
        case_array(case_status,                 entry_record) = "" 
        case_array(rsum_PMI_const,              entry_record) = ""		
        case_array(pmap_begin_const,            entry_record) = ""
        case_array(pmap_end_const,              entry_record) = ""
        case_array(pmap_name_const,             entry_record) = ""
        
        entry_record = entry_record + 1			'This increments to the next entry in the array'
        stats_counter = stats_counter + 1
        all_pmi_array = trim(all_pmi_array & Client_PMI & "*") 'Adding MAXIS case number to case number string
    End if 
    excel_row = excel_row + 1
Loop

objExcel.Quit		'Once all of the clients have been added to the array, the excel document is closed because we are going to open another document and don't want the script to be confused

'Opening the Excel file
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
objExcel.DisplayAlerts = True

'Changes name of Excel sheet
ObjExcel.ActiveSheet.Name = "Member MMIS Information"

'adding column header information to the Excel list
ObjExcel.Cells(1,  1).Value = "Billed PMI"
ObjExcel.Cells(1,  2).Value = "RSUM PMI"
ObjExcel.Cells(1,  3).Value = "Last Name"
ObjExcel.Cells(1,  4).Value = "First Name"
ObjExcel.Cells(1,  5).Value = "DOB"
ObjExcel.Cells(1,  6).Value = "Gender"
ObjExcel.Cells(1,  7).Value = "1st Case"
ObjExcel.Cells(1,  8).Value = "1st Type/Prog"
ObjExcel.Cells(1,  9).Value = "1st Elig Dates"
ObjExcel.Cells(1, 10).Value = "2nd Case"
ObjExcel.Cells(1, 11).Value = "2nd Type/Prog"
ObjExcel.Cells(1, 12).Value = "2nd Elig Dates"
ObjExcel.Cells(1, 13).Value = "3rd Case"
ObjExcel.Cells(1, 14).Value = "3rd Type/Prog"
ObjExcel.Cells(1, 15).Value = "3rd Elig Dates"
ObjExcel.Cells(1, 16).Value = "PMAP Start"
ObjExcel.Cells(1, 17).Value = "PMAP End"
ObjExcel.Cells(1, 18).Value = "PMAP Name"
ObjExcel.Cells(1, 19).Value = "Status"

FOR i = 1 to 19 	'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

excel_row = 2
'----------------------------------------------------------------------------------------------------Gathering Person information based on provided PMI 
get_to_RKEY 'Navigate to RKEY and clear any exising searches
Call clear_line_of_text(4, 19)  'Clearing PMI
Call clear_line_of_text(5, 19)  'Clearing SSN
Call clear_line_of_text(5, 48)  'Clearing Medicare ID
Call clear_line_of_text(6, 19)  'Clearing Last Name 
Call clear_line_of_text(6, 48)  'Clearing First Name 
Call clear_line_of_text(6, 69)  'Clearing Middle Initial
Call clear_line_of_text(7, 19)  'Clearing DOB 
Call clear_line_of_text(9, 19)  'Clearing Case Number 
Call clear_line_of_text(9, 48)  'Clearing Client Option Number 
Call clear_line_of_text(9, 69)  'Clearing Case Type 

For item = 0 to UBound(case_array, 2)
    Client_PMI = case_array(clt_PMI_const, item)

    get_to_RKEY
    Call write_value_and_transmit (Client_PMI, 4, 19)
    EmReadscreen RKEY_panel_check, 4, 1, 52
    If RKEY_panel_check = "RKEY" then 
        EmReadscreen RKEY_error, 78, 24, 2
        case_array(case_status, item) = trim(RKEY_error)
    Else 
        'All accessable cases will have information gathered for them from the RCIP panel.  
        Call write_value_and_transmit ("RCIP", 1, 8)
        Call MMIS_panel_confirmation("RCIP", 52) 
        
        EmReadscreen Client_SSN, 9, 5, 28
        Client_SSN = trim(Client_SSN)
        
        If Client_SSN = "" then
            case_array(case_status, item) = "Unable to find SSN in MMIS."
        Else 
            case_array(case_status, item) = ""
            Case_array(client_SSN_const, item) = Client_SSN
        End if 

        EmReadscreen last_name, 17, 3, 2
        Case_array(last_name_const, item) = trim(last_name)
        
        EmReadscreen first_name, 13, 3, 20 
        Case_array(first_name_const, item) = trim(first_name)
        
        EmReadscreen client_DOB, 10, 2, 24
        case_array(DOB_const, item) = trim(client_DOB)
        
        EmReadscreen gender_code, 1, 8, 28
        case_array(gender_const, item) = gender_code
    End if 
Next 

'----------------------------------------------------------------------------------------------------Health Care Information Report
For item = 0 to UBound(case_array, 2)
    Client_SSN = case_array(client_SSN_const, item) 
    Client_PMI = case_array(clt_PMI_const, item)
    
    If case_array(case_status, item) = "RECIPIENT ID COULD NOT BE FOUND" then
        objExcel.Cells(excel_row,  1).Value = case_array (clt_PMI_const, item)
        objExcel.Cells(excel_row, 19).Value = case_array(case_status,    item)
        excel_row = excel_row + 1
    Else 
        get_to_RKEY
        Call clear_line_of_text(4, 19)  'Clearing PMI
        Call clear_line_of_text(5, 19)  'Clearing SSN
        
        If trim(Client_SSN) = "" then 
            EMWriteScreen Client_PMI, 4, 19
        Else 
            EMWriteScreen Client_SSN, 5, 19
        End if 
        
        Call write_value_and_transmit("I", 2, 19)   'transmitting to next screen. Could be RSEL or RSUM. If SSN is searched and more than one record is found, the RSEL screen will appear. 
        RSEL_row = 7
        Do 
            EmReadscreen RSEL_panel_check, 4, 1, 52
            EmReadscreen panel_check, 4, 1, 51
            If RSEL_panel_check = "RSEL" then
                EmReadscreen RSEL_SSN, 9, RSEL_row, 48
                If RSEL_SSN = Client_SSN then
                    duplicate_entry = True 
                    Call write_value_and_transmit("X", RSEL_row, 2)
                    '---------------------------------------This bit is for the rare case where you cannot select the SSN on RSEL. Those will be on the error list
                    EmReadscreen RSEL_panel_check, 4, 1, 52  'RSEL is listed at column 52 
                    EmReadscreen panel_check, 4, 1, 51
                    If RSEL_panel_check = "RSEL" then
                        EmReadscreen RSEL_error, 70, 24, 2
                        If trim(RSEL_error) <> "" then
                            EmReadscreen RSEL_pmi, 8, RSEL_row, 4
                            case_array(rsum_PMI_const, item) = ""
                            case_array(case_status, item) = "RSEL screen error with PMI: " & RSEL_pmi & ". " & trim(RSEL_error)
                            duplicate_entry = False 'stopping the futher search for case information 
                            Exit do 
                        End if 
                    End if         
                else 
                    Exit do
                    duplicate_entry = False 
                End if 
            End if     
            
            If panel_check = "RSUM" then 
                '1st case type/prog/elig/case number 
                EmReadscreen RSUM_PMI, 8, 2, 2
                Case_array(rsum_PMI_const, item) = RSUM_PMI
                EmReadscreen first_case_number, 8, 7, 16
                first_case_number = trim(first_case_number)
                
                If first_case_number = "" then case_array(case_status, item) = "No active programs in MMIS under billed PMI."
                   
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
                
                EmReadscreen third_case_number, 8, 11, 16
                third_case_number = trim(third_case_number)
                If third_case_number <> "" then 
                    case_array(third_case_number_const, item) = third_case_number
                    EmReadscreen third_program, 2, 10, 13
                    EmReadscreen third_type, 2, 10, 35
                    If trim(third_program) <> "" then 
                        third_elig_type = third_program & "-" & third_type
                        case_array(third_type_const, item) = third_elig_type
                        '1st elig dates 
                        EmReadscreen third_elig_start, 8, 11, 35
                        EmReadscreen third_elig_end, 8, 11, 54
                        third_elig_dates = third_elig_start &  " - " & third_elig_end
                        case_array(third_elig_const, item) = third_elig_dates
                    End if    
                End if 
                
                'Reading PMAP Information from RPPH panel 
                Call write_value_and_transmit("RPPH", 1, 8)
                Call MMIS_panel_confirmation("RPPH", 52)
                
                EmReadscreen pmap_begin, 8, 13, 5
                case_array(pmap_begin_const, item) = trim(pmap_begin)
                
                EmReadscreen pmap_end, 8, 13, 14
                case_array(pmap_end_const, item) = trim(pmap_end)
                
                EMReadScreen hp_code, 10, 13, 23
                If hp_code = "A585713900" then case_array(pmap_name_const, item) = "HealthPartners"
                If hp_code = "A565813600" then case_array(pmap_name_const, item) = "Ucare"
                If hp_code = "A405713900" then case_array(pmap_name_const, item) = "Medica"
                If hp_code = "A065813800" then case_array(pmap_name_const, item) = "BluePlus"
                If hp_code = "A836618200" then case_array(pmap_name_const, item) = "Hennepin Health PMAP"
                If hp_code = "A965713400" then case_array(pmap_name_const, item) = "Hennepin Health SNBC"
            End if     
            
            'outputting to Excel 
            objExcel.Cells(excel_row,  1).Value = case_array (clt_PMI_const,            item)
            objExcel.Cells(excel_row,  2).Value = case_array (rsum_PMI_const,           item)
            objExcel.Cells(excel_row,  3).Value = case_array (last_name_const,          item)
            objExcel.Cells(excel_row,  4).Value = case_array (first_name_const,         item)
            objExcel.Cells(excel_row,  5).Value = case_array (DOB_const,                item)
            objExcel.Cells(excel_row,  6).Value = case_array (gender_const,             item)
            objExcel.Cells(excel_row,  7).Value = case_array (first_case_number_const,  item)
            objExcel.Cells(excel_row,  8).Value = case_array (first_type_const, 	    item)
            objExcel.Cells(excel_row,  9).Value = case_array (first_elig_const, 	    item)
            objExcel.Cells(excel_row, 10).Value = case_array (second_case_number_const, item)
            objExcel.Cells(excel_row, 11).Value = case_array (second_type_const, 	    item)
            objExcel.Cells(excel_row, 12).Value = case_array (second_elig_const, 	    item)
            objExcel.Cells(excel_row, 13).Value = case_array (third_case_number_const,  item)
            objExcel.Cells(excel_row, 14).Value = case_array (third_type_const,      	item)
            objExcel.Cells(excel_row, 15).Value = case_array (third_elig_const,         item) 
            objExcel.Cells(excel_row, 16).Value = case_array(pmap_begin_const,          item)
            objExcel.Cells(excel_row, 17).Value = case_array(pmap_end_const,            item)
            objExcel.Cells(excel_row, 18).Value = case_array(pmap_name_const,           item)
            objExcel.Cells(excel_row, 19).Value = case_array(case_status,               item)
            excel_row = excel_row + 1
                
            If duplicate_entry = True then 
                RSEL_row = RSEL_row + 1
                PF3
                EmReadscreen RSEL_panel_check, 4, 1, 52  'RSEL is listed at column 52 
                If RSEL_panel_check = "RSEL" then
                    case_array(first_case_number_const, item) = ""
            		case_array(rsum_PMI_const,          item) = ""		
                    case_array(first_type_const, 	    item) = ""				
                    case_array(first_elig_const, 	    item) = ""             
                    case_array(second_case_number_const,item) = ""              
                    case_array(second_type_const, 	    item) = ""              
                    case_array(second_elig_const,       item) = ""              
                    case_array(third_case_number_const, item) = ""
                    case_array(third_type_const,        item) = ""				
                    case_array(third_elig_const,        item) = ""
                    case_array(case_status,             item) = ""	
                    case_array(pmap_begin_const,        item) = ""
                    case_array(pmap_end_const,          item) = ""
                    case_array(pmap_name_const,         item) = ""
                Else 
                    exit do 'No more cases on RSEL 
                End if 
            Else
                PF3
                Exit do     'cases that did not have more than one known entry 
            End if 
        Loop 
    End if 
Next     
    
FOR i = 1 to 19		'formatting the cells
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created. Please review for cases that need to be processed manually.")