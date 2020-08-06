'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - BULK DUPLICATE PMI HC ELIG.vbs"
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
call changelog_update("07/16/2020", "Added gender and DOB fields to report.", "Ilse Ferris, Hennepin County")
call changelog_update("09/13/2018", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function HCRE_panel_bypass() 
	'handling for cases that do not have a completed HCRE panel
	PF3		'exits PROG to prommpt HCRE if HCRE insn't complete
	Do
		EMReadscreen HCRE_panel_check, 4, 2, 50
		If HCRE_panel_check = "HCRE" then
			PF10	'exists edit mode in cases where HCRE isn't complete for a member
			PF3
		END IF
	Loop until HCRE_panel_check <> "HCRE"
End Function

Function MMIS_panel_check(panel_name)
	Do 
		EMReadScreen panel_check, 4, 1, 52
		If panel_check <> panel_name then Call write_value_and_transmit(panel_name, 1, 8)
	Loop until panel_check = panel_name
End function

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
'get_county_code

MAXIS_footer_month = CM_mo	'establishing footer month/year 
MAXIS_footer_year = CM_yr 

file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\BZ ongoing projects\SSIS - PMI\SA 2020-07.xlsx"

'The dialog is defined in the loop as it can change as buttons are pressed 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "MAXIS TO METS Conversion Information"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used when a list of conversion cases are provided by the METS team or DHS."
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
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

DIM case_array()
ReDim case_array(16, 0)

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

'Now the script adds all the clients on the excel list into an array
excel_row = 2 're-establishing the row to start checking the members for
entry_record = 0
Do   
    'Loops until there are no more cases in the Excel list

    Client_PMI = objExcel.cells(excel_row, 1).Value          'reading the PMI from Excel 
    Client_PMI = trim(Client_PMI)
    If Client_PMI = "" then exit do

    ReDim Preserve case_array(16, entry_record)	'This resizes the array based on the number of rows in the Excel File'
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

    entry_record = entry_record + 1			'This increments to the next entry in the array'
    stats_counter = stats_counter + 1
    excel_row = excel_row + 1
Loop

objExcel.Quit		'Once all of the clients have been added to the array, the excel document is closed because we are going to open another document and don't want the script to be confused
back_to_self
call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year

For item = 0 to UBound(case_array, 2)
    Client_PMI = case_array(clt_PMI_const, item)

    Call navigate_to_MAXIS_screen("PERS", "____")
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
	If PRIV_check = "PRIV" then
        case_array(case_status, item) = "PRIV case."
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the MAXIS case number
		transmit
    Else
        Call write_value_and_transmit(client_PMI, 15, 36)
        EmReadscreen error_message, 80, 24, 2
        error_message = trim(error_message)
        If error_message <> "" then 
            case_array(case_status, item) = error_message 
        Else     
            EmReadscreen panel_PMI, 10, 8, 71
            panel_PMI = trim(panel_PMI)
            
            If instr(Client_PMI, panel_PMI) then
                EmReadscreen Client_SSN, 11, 8, 7
                Client_SSN = replace(Client_SSN, "-", "")
                Client_SSN = trim(Client_SSN)
                
                EmReadscreen no_case_message, 20, 24, 38
                If no_case_message = "NO MAXIS CASE EXISTS" then 
                    case_array(case_status, item) = "Client does not have a MAXIS case." 
                Elseif Client_SSN = "" then
                    case_array(case_status, item) = "Unable to find SSN in PERS."
                Else 
                    case_array(case_status, item) = ""
                    Case_array(client_SSN_const, item) = Client_SSN
                End if 
                EmReadscreen last_name, 20, 8, 21
                Case_array(last_name_const, item) = trim(last_name)
                
                EmReadscreen first_name, 12, 8, 42 
                Case_array(first_name_const, item) = trim(first_name)
                
                EmReadscreen client_DOB, 10, 8, 60
                case_array(DOB_const, item) = trim(client_DOB)
                
                EmReadscreen gender_code, 1, 8, 58
                case_array(gender_const, item) = gender_code
                
            Else 
                case_array(case_status, item) = "Unable to find PMI in PERS." 
            End if 
        End if 
    End if 
Next 

'-------------------------------------------------------------------------------------------------------------------------------------MMIS portion of the script
Call navigate_to_MMIS_region("CTY ELIG STAFF/UPDATE")	'function to navigate into MMIS, select the HC realm, and enters the prior autorization area

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
ObjExcel.Cells(1,  7).Value = "1st case"
ObjExcel.Cells(1,  8).Value = "1st type/prog"
ObjExcel.Cells(1,  9).Value = "1st elig dates"
ObjExcel.Cells(1, 10).Value = "2nd case"
ObjExcel.Cells(1, 11).Value = "2nd type/prog"
ObjExcel.Cells(1, 12).Value = "2nd elig dates"
ObjExcel.Cells(1, 13).Value = "3rd case"
ObjExcel.Cells(1, 14).Value = "3rd type/prog"
ObjExcel.Cells(1, 15).Value = "3rd elig dates"

FOR i = 1 to 15 	'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

excel_row = 2
For item = 0 to UBound(case_array, 2)
    Client_SSN = case_array(client_SSN_const, item) 
    Client_PMI = case_array(clt_PMI_const, item)
    client_PMI = right("00000000" & client_pmi, 8)
    
    If case_array(case_status, item) = "" then
        MMIS_panel_check("RKEY") 
        Call clear_line_of_text(4, 19)
        EMWriteScreen Client_SSN, 5, 19
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
                    
                    '---------------------------------------This bit is for the rare case where you cannot select the SSN on RSEL. Those will be on the error list
                    EmReadscreen RSEL_panel_check, 4, 1, 52  'RSEL is listed at column 52 
                    EmReadscreen panel_check, 4, 1, 51
                    If RSEL_panel_check = "RSEL" then
                        EmReadscreen RSEL_error, 70, 24, 2
                        If trim(RSEL_error) <> "" then 
                            MsgBox RSEL_error
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
                If first_case_number = "" then
                    case_array(case_status, item) = "No active programs in MMIS under billed PMI."
                Else     
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
                    
                'outputting to Excel 
                If case_array(case_status, item) = "" then 
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
                    excel_row = excel_row + 1
                End if 
                    
                If duplicate_entry = True then 
                    RSEL_row = RSEL_row + 1
                    PF3
                    EmReadscreen RSEL_panel_check, 4, 1, 52  'RSEL is listed at column 52 
                    If RSEL_panel_check = "RSEL" then
                        'RSEL_row = 8
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
                    Else 
                        exit do 'No more cases on RSEL 
                    end if 
                else
                    PF3
                    exit do     'cases that did not have more than one known entry 
                End if 
            End if 
        loop 
    End if 
Next     
    
FOR i = 1 to 15		'formatting the cells
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Adding another sheet
ObjExcel.Worksheets.Add().Name = "Error List" 

'formatting excel file with columns for case number and interview date/time
objExcel.cells(1, 1).value 	= "Billed PMI"
objExcel.cells(1, 2).value 	= "RSUM PMI"
objExcel.cells(1, 3).value 	= "Error Reason"
objExcel.cells(1, 4).value 	= "Last Name"
objExcel.cells(1, 5).value 	= "First Name"
	
FOR i = 1 to 5									'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'Adding the case information to Excel
excel_row = 2
For item = 0 to UBound(case_array, 2)
    If case_array(case_status, item) <> "" then
	    ObjExcel.Cells(excel_row, 1).value = case_array (clt_PMI_const,    item)
	    ObjExcel.Cells(excel_row, 2).value = case_array (rsum_PMI_const,   item)  'This outputs the reason for the error
        ObjExcel.Cells(excel_row, 3).value = case_array (case_status,      item)  'This outputs the reason for the error
        objExcel.Cells(excel_row, 4).Value = case_array (last_name_const,  item)
        objExcel.Cells(excel_row, 5).Value = case_array (first_name_const, item)
	    excel_row = excel_row + 1 
    End if 
Next

'Formatting the columns to autofit after they are all finished being created.
FOR i = 1 to 5
	objExcel.Columns(i).autofit()
Next

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created. Please review for cases that need to be processed manually.")