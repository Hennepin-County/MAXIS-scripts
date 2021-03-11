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
call changelog_update("07/30/2020", "Removed MEMO coding, and cleaned up other commented out code.", "Ilse Ferris, Hennepin County")
call changelog_update("07/30/2020", "Removed the Output Waiver Lists option. Added Output Forms List for all recipients who do not have forms.", "Ilse Ferris, Hennepin County")
call changelog_update("05/14/2020", "Added case number for the Output Waiver Lists option.", "Ilse Ferris, Hennepin County")
call changelog_update("03/11/2020", "Added case mgr name and agency info from MMIS for the Output Waiver Lists option.", "Ilse Ferris, Hennepin County")
call changelog_update("02/11/2020", "Added waiver code and case mgr NPI to initial monthly upload option. Removed testing msgboxes.", "Ilse Ferris, Hennepin County")
call changelog_update("01/30/2020", "Added excel row selection for certain processes to speed up report time.", "Ilse Ferris, Hennepin County")
call changelog_update("11/06/2019", "Added ability to run all spreadsheets in a process concurrently.", "Ilse Ferris, Hennepin County")
call changelog_update("10/17/2019", "Added updated SPEC/MEMO verbiage.", "Ilse Ferris, Hennepin County")
call changelog_update("09/23/2019", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

Function MMIS_panel_check(panel_name, col)
	Do 
		EMReadScreen panel_check, 4, 1, col
		If panel_check <> panel_name then Call write_value_and_transmit(panel_name, 1, 8)
	Loop until panel_check = panel_name
End function

''----------------------------------------------------------------------------------------------------Gathering ALL AVS FORMS information
Function AVS_sync()   
    objExcel.worksheets("All AVS Forms").Activate 'Activates worksheet based on user selection
    
    DIM master_array()
    ReDim master_array(6, 0)
    
    const SMI_AAF_const          = 0
    const scan_date_AAF_const    = 1
    const case_number_AAF_const  = 2
    const PMI_AAF_const          = 3
    const client_name_AAF_const  = 4
    const note_created_AAF_const = 5    
    const case_note_const        = 6    
    
    excel_row = 2
    master_record = 0
    
    Do 
        SMI_AAF = ObjExcel.Cells(excel_row, 1).Value
        SMI_AAF  = trim(SMI_AAF)
        If SMI_AAF = "" then exit do 
        
        scan_date_AAF       = ObjExcel.Cells(excel_row, 2).Value        
        MAXIS_case_number   = ObjExcel.Cells(excel_row, 3).Value
        PMI_AAF             = ObjExcel.Cells(excel_row, 4).Value
        client_name_AAF     = ObjExcel.Cells(excel_row, 5).Value
        note_confirm_AAF    = ObjExcel.Cells(excel_row, 6).Value
        
        ReDim Preserve master_array(6, master_record)	'This resizes the array based on the number of rows in the Excel File'
        master_array(SMI_AAF_const,         master_record) = SMI_AAF
        master_array(scan_date_AAF_const,   master_record) = trim(scan_date_AAF)
        master_array(case_number_AAF_const, master_record) = trim(MAXIS_case_number)
        master_array(PMI_AAF_const,         master_record) = trim(PMI_AAF)
        master_array(client_name_AAF_const, master_record) = trim(client_name_AAF)
        master_array(case_note_const,       master_record) = trim(note_confirm_AAF)
        
        master_record = master_record + 1			'This increments to the next entry in the array'
        STATS_counter = STATS_counter + 1
        excel_row = excel_row + 1
    LOOP
    
    '----------------------------------------------------------------------------------------------------Gathering monthly information & exporting ALL AVS FORMS information
    For Each objWorkSheet In objWorkbook.Worksheets
        If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "All AVS Forms" then months_list = months_list & objWorkSheet.Name & ","
    Next
    months_list = trim(months_list)  'trims excess spaces of months_list
    If right(months_list, 1) = "," THEN months_list = left(months_list, len(months_list) - 1) 'trimming off last comma
    array_of_months = split(months_list, ",")   'Creating new array
    
    master_count = 0
    
    For each month_sheet in array_of_months
        objExcel.worksheets(month_sheet).Activate 'Activates worksheet based on user selection
        excel_row = 2
        
        DO 
            month_SMI_number = ObjExcel.Cells(excel_row, SMI_col).Value
            month_SMI_number = trim(month_SMI_number)
            If month_SMI_number = "" then exit do 
            
            month_case_note = ObjExcel.Cells(excel_row, note_col).Value
            month_scan_date = ObjExcel.Cells(excel_row, forms_col).Value
            month_case_number = objExcel.Cells(excel_row, cn_col).Value 
            month_PMI = objExcel.Cells(excel_row, pmi_col).Value
            month_client_name = ObjExcel.Cells(excel_row, client_name_col).Value 
            
            For item = 0 to UBound(master_array, 2)
                SMI_AAF = master_array(SMI_ECF_const, item)  
                
                If SMI_AAF = month_SMI_number then
                    'scan date or form date 
                    If master_array(scan_date_AAF_const, item) = "" then 
                        master_array(scan_date_AAF_const, item) = trim(month_scan_date)'revaluing case note  
                    Elseif trim(month_scan_date) = "" then 
                        ObjExcel.Cells(excel_row, forms_col).Value = master_array(scan_date_AAF_const, item)
                    End if 
                    
                    'case note dates: Some statuses will be text vs date for tracking. This replaces them once they are case noted. 
                    If master_array(case_note_const, item) = "" then 
                        master_array(case_note_const, item) = trim(month_case_note) 'revaluing case note  
                    Elseif trim(month_case_note) = "" or isdate(month_case_note) = False then 
                        ObjExcel.Cells(excel_row, note_col).Value = master_array(case_note_const, item)
                    End if 
                    
                    master_array(case_number_AAF_const, item) = trim(month_case_number)  'revaluing case number 
                    master_array(PMI_AAF_const, item) = trim(month_PMI)    'revaluing PMI number 
                    master_array(client_name_AAF_const, item) = trim(month_client_name) 'revaluing client name 
                    
                    'objExcel.Cells(excel_row, 19).Interior.ColorIndex = 3	'Fills the row with red    
                    
                    master_count = master_count + 1
                    exit for 
                End if  
            Next
            excel_row = excel_row + 1
            month_SMI_number = ""
            SMI_AAF = "" 
        Loop 
    Next 
    '''----------------------------------------------------------------------------------------------------Filling in any missing ALL AVS FORMS information
    'objExcel.worksheets("All AVS Forms").Activate 'Activates worksheet based on user selection
    '
    'excel_row = 2
    'For item = 0 to UBound(master_array, 2)
    '    ObjExcel.Cells(excel_row, 3).Value = master_array(case_number_AAF_const, item)
    '    ObjExcel.Cells(excel_row, 4).Value = master_array(PMI_AAF_const,         item)
    '    ObjExcel.Cells(excel_row, 5).Value = master_array(client_name_AAF_const, item)
    '    ObjExcel.Cells(excel_row, 6).Value = master_array(case_note_const,       item)
    '    objExcel.Cells(excel_row, 3).Interior.ColorIndex = 3	'Fills the row with red     
    '    excel_row = excel_row + 1
    'Next
    '
    'FOR i = 1 to 6		'formatting the cells
    '    objExcel.Columns(i).AutoFit()				'sizing the columns'
    'NEXT
    msgbox "Sync Complete"
End function
'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
MAXIS_footer_month = CM_mo	'establishing footer month/year 
MAXIS_footer_year = CM_yr 

'column numbers 
cn_col          = 3
PMI_col         = 4
client_name_col = 5
SMI_col         = 6
waiver_col      = 7
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
NPI_col         = 21

'----------------------------------------------------------------------------------------------------INITIAL DIALOG
'The dialog is defined in the loop as it can change as buttons are pressed 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 246, 110, "AVS Processing Selection"
  DropListBox 120, 50, 115, 15, "Select one..."+chr(9)+"Case & Person Noting"+chr(9)+"ECF Forms Received"+chr(9)+"Initial Monthly Upload"+chr(9)+"New Person Information"+chr(9)+"Output Forms Lists"+chr(9)+"Run Sync", AVS_option
  EditBox 85, 75, 45, 15, excel_row_to_start
  ButtonGroup ButtonPressed
    OkButton 140, 75, 45, 15
    CancelButton 190, 75, 45, 15
  Text 20, 50, 95, 10, "Select the processing option:"
  GroupBox 10, 5, 230, 65, "Using this script:"
  Text 10, 80, 70, 10, "**Excel row to start:"
  Text 20, 20, 210, 20, "This script should be used when a list of AVS cases are provided by the METS team or DHS."
  Text 10, 95, 225, 10, "** For Case & Person Noting OR New Person Information Option Only"
EndDialog

Do     
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation 
        If AVS_option = "Select one..." then err_msg = "Select the AVS process to complete."
        If AVS_option = "Case & Person Noting" or AVS_option = "New Person Information" then
            If excel_row_to_start = "" then err_msg = "Enter the Excel Row to Start."
        End if
        If err_msg <> "" Then MsgBox err_msg
    Loop until err_msg = ""
CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

'----------------------------------------------------------------------------------------------------------------------------------------------------ECF FORMS RECEIVED
If AVS_option = "ECF Forms Received" then 
    dialog1 = ""
    'The dialog is defined in the loop as it can change as buttons are pressed 
   Dialog1 = ""
    BeginDialog Dialog1, 0, 0, 266, 115, "AVS Forms Procesing"
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
        dialog Dialog1
        cancel_without_confirmation 
        If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue." 
        If err_msg <> "" Then MsgBox err_msg
    Loop until err_msg = ""
    If objExcel = "" Then call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'

    excel_row = 2
    entry_record = 0
    
    DIM upload_array()
    ReDim upload_array(2, 0)
    
    const SMI_ECF_const      = 0
    const scan_date_const    = 1
    const SMI_found_const    = 2
    
          
    Do 
    	SMI_ECF_number  = ObjExcel.Cells(excel_row, 1).Value
    	SMI_ECF_number  = trim(SMI_ECF_number)
        If SMI_ECF_number = "" then exit do 
        
        scan_date = ObjExcel.Cells(excel_row, 2).Value
        scan_date = trim(scan_date)
        
        ReDim Preserve upload_array(2, entry_record)	'This resizes the array based on the number of rows in the Excel File'
        upload_array(SMI_ECF_const,	        entry_record) = SMI_ECF_number 		
        upload_array(scan_date_const, 	    entry_record) = scan_date 	
        upload_array(SMI_found_const, 	    entry_record) = FALSE 	
        
        entry_record = entry_record + 1			'This increments to the next entry in the array'
        STATS_counter = STATS_counter + 1
        excel_row = excel_row + 1
    LOOP
    
    objExcel.Quit   'Closes the initial spreadsheet 
    objExcel = ""
    
    file_selection = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\AVS\AVS Forms Distribution Master List.xlsx"
    Call excel_open(file_selection, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    
    objExcel.worksheets("All AVS Forms").Activate 'Activates worksheet based on user selection
    
    '----------------------------------------------------------------------------------------------------FILTERING THE ARRAY 
    form_count = 0
    excel_row = 2
    
    DO 
        SMI_number = ObjExcel.Cells(excel_row, 1).Value
        SMI_number = trim(SMI_number)
        If SMI_number = "" then exit do 
        
        For item = 0 to UBound(upload_array, 2)
            SMI_ECF_number = upload_array(SMI_ECF_const, item)  
            scan_date = upload_array(scan_date_const, item)
            
            If trim(SMI_ECF_number) = trim(SMI_number) then
                'Adding inforamtion to the array that will then update the monthly lists 
                upload_array(SMI_found_const, item) = True 
                'objExcel.Cells(excel_row, 1).Value = SMI_ECF_number
                objExcel.Cells(excel_row, 2).Value = scan_date
                objExcel.Cells(excel_row, 2).Interior.ColorIndex = 3	'Fills the row with red                
                form_count = form_count + 1
                exit for
            else 
                match_found = False 
            end if
        Next
        excel_row = excel_row + 1
        SMI_number = "" 
    Loop

    For item = 0 to UBound(upload_array, 2)
        SMI_ECF_number = upload_array(SMI_ECF_const, item)  
        scan_date = upload_array(scan_date_const, item)
        
        If upload_array(SMI_found_const, item) = False then  
            'Adding inforamtion to the array that will then update the monthly lists 
            objExcel.Cells(excel_row, 1).Value = SMI_ECF_number
            objExcel.Cells(excel_row, 2).Value = scan_date 
            objExcel.Cells(excel_row, 2).Interior.ColorIndex = 3	'Fills the row with red                
            form_count = form_count + 1
            excel_row = excel_row + 1
        end if 
    Next
    'Syncing the resident lists with the All AVS forms list
    'Call AVS_sync 
End if     

'----------------------------------------------------------------------------------------------------
If AVS_option = "Initial Monthly Upload" then 

    file_selection = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\AVS\AVS Forms Distribution Master List.xlsx"
    Call excel_open(file_selection, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    
    'adding column header information to the Excel list
    ObjExcel.Cells(1, 6).Value = "SMI"
    ObjExcel.Cells(1, 7).Value = "Waiver Type"
    ObjExcel.Cells(1, 8).Value = "Waiver start"
    ObjExcel.Cells(1, 9).Value = "Waiver end"
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
    ObjExcel.Cells(1, 21).Value = "Case Mgr NPI"
    
    FOR i = 1 to 21 	'formatting the cells'
    	objExcel.Cells(1, i).Font.Bold = True		'bold font'
    	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
    	objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT
    
    ObjExcel.columns(8).NumberFormat = "mm/dd/yy" 		'formatting waiver start date as a date
    ObjExcel.columns(9).NumberFormat = "mm/dd/yy" 		'formatting waiver end date as a date
    
    DIM case_array()
    ReDim case_array(19, 0)
    
    'constants for array
    const case_number_const     	= 0
    const clt_PMI_const 	        = 1
    const SMI_num_const             = 2
    const waiver_type_const         = 3
    const waiver_start_const	    = 4
    const waiver_end_const          = 5
    const medicare_info_const       = 6
    const first_case_number_const   = 7
    const first_type_const 	        = 8
    const first_elig_const 	        = 9
    const second_case_number_const  = 10
    const second_type_const         = 11
    const second_elig_const         = 12
    const third_case_number_const   = 13
    const third_type_const     	    = 14
    const third_elig_const          = 15
    const case_status               = 16
    const rlva_coding_const         = 17 
    const name_const                = 18   
    const NPI_const                 = 19        
    
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
        
        name_of_client = objExcel.cells(excel_row, client_name_col).Value   'reading the case number from Excel   
            
        ReDim Preserve case_array(19, entry_record)	'This resizes the array based on the number of rows in the Excel File'
        case_array(case_number_const,           entry_record) = MAXIS_case_number	'The client information is added to the array'
        case_array(clt_PMI_const,               entry_record) = Client_PMI			
        case_array(SMI_num_const,               entry_record) = "" 
        case_array(waiver_type_const,	        entry_record) = ""                      
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
        case_array(name_const,                  entry_record) = trim(name_of_client)
        case_array (NPI_const,                  entry_record) = ""
        
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
            Call MMIS_panel_check("RKEY", 52) 
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
                        Case_array(waiver_type_const, item) = ""
                        Case_array(waiver_start_const, item) = ""
                        Case_array(waiver_end_const, item) = ""
                    Else 
                        EmReadscreen waiver_type, 1, 15, 15
                        EMReadscreen waiver_start_date, 8, 15, 25
                        EmReadscreen waiver_end_date, 8, 15, 46
                        Case_array(waiver_type_const, item) = trim(waiver_type)
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
                    Call MMIS_panel_check("RLVA", 52)
                    EmReadscreen rlva_coding, 12, 14, 42 'most recent living arrangement 
                    case_array(rlva_coding_const, item) = rlva_coding
                    
                    If waiver_info <> "" then 
                        'RMGR
                        Call write_value_and_transmit("RMGR", 1, 8)
                        Call MMIS_panel_check("RMGR", 51)
                        EmReadscreen NPI_number, 10, 7, 60
                        case_array (NPI_const, item) = trim(NPI_number)
                    Else
                        case_array (NPI_const, item) = ""
                    End if 
                        
                    'outputting to Excel 
                    objExcel.Cells(excel_row, SMI_col).Value = case_array (SMI_num_const,                  item)
                    objExcel.Cells(excel_row, waiver_col).Value = case_array (waiver_type_const,	       item)
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
                    objExcel.Cells(excel_row, NPI_col).Value = case_array (NPI_const,                      item)                        
                    
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
    
    FOR i = 1 to 21		'formatting the cells
    	objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT
End if     
    
If AVS_option = "Case & Person Noting" then 
    file_selection = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\AVS\AVS Forms Distribution Master List.xlsx"
    Call excel_open(file_selection, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    objExcel.worksheets("All AVS Forms").Activate 'Activates worksheet based on user selection
    
    'Setting up MAXIS to be ready for case noting 
    back_to_self
    call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year
    
    excel_row = excel_row_to_start
    case_note_total = 0
    
    Do   
        'Loops until there are no more cases in the Excel list
        MAXIS_case_number = objExcel.cells(excel_row, 3).Value   'reading the case number from Excel   
        MAXIS_case_number = Trim(MAXIS_case_number)
        If MAXIS_case_number = "" then exit do
        
        client_PMI = objExcel.cells(excel_row, 4).Value
        client_PMI = trim(client_PMI) 
        
        client_name = objExcel.cells(excel_row, 5).Value
        client_name = trim(client_name)
        Call fix_case(client_name, 2)
        client_name = trim(client_name)
        
        form_date = objExcel.cells(excel_row, 2).Value
        form_date = trim(form_date)
        
        note_date = objExcel.cells(excel_row, 6).Value
        note_date = trim(note_date)
        
        stats_counter = stats_counter + 1
    
        'Skipping cases that do not have a form date already listed or already have a case/person note.
        If trim(form_date) <> "" then
            If trim(note_date) = "" then
                Call navigate_to_MAXIS_screen("STAT", "MEMB")
                EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
                If PRIV_check = "PRIV" then
                    objExcel.cells(excel_row, 6).Value = "PRIVILEGED"
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
                        objExcel.cells(excel_row, 6).Value = "OUT OF COUNTY"
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
                                 objExcel.cells(excel_row, 6).Value = "PERS note issue"
                            ELSE
                                EMreadscreen edit_mode_required_check, 6, 5, 3		'if not person not exists, person note goes directly into edit mode
                                If edit_mode_required_check <> "      " then PF9
                                write_new_line_in_person_note(note_header)
                                write_new_line_in_person_note("--")
                                write_new_line_in_person_note(note_body)
                            
                                PF3 'to save and exit person notes
                                '---------------------------------------------------------------Creating the CASE note
                                start_a_blank_CASE_NOTE
                                Call write_variable_in_CASE_NOTE(note_header)
                                Call write_variable_in_CASE_NOTE("--")
                                Call write_variable_in_CASE_NOTE(note_body)
                                PF3 'to save and exit case notes
                                objExcel.cells(excel_row, 6).Value = date
                                objExcel.Cells(excel_row, 6).Interior.ColorIndex = 3	'Fills the row with red
                            End if 
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

   FOR i = 1 to 6		'formatting the cells
       objExcel.Columns(i).AutoFit()				'sizing the columns'
   NEXT
   'Syncing the resident lists with the All AVS forms list
   'Call AVS_sync 
End if 

IF AVS_option = "New Person Information" then 
   file_selection = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\AVS\AVS Forms Distribution Master List.xlsx"
   Call excel_open(file_selection, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
   objExcel.worksheets("All AVS Forms").Activate 'Activates worksheet based on user selection
   
   excel_row = excel_row_to_start    'starting point
    
    DO
        master_SMI_number = ObjExcel.Cells(excel_row, 1).Value  'from All AVS forms list
        master_SMI_number = trim(master_SMI_number)
        If master_SMI_number = "" then exit do 
        
        'Loops until there are no more cases in the Excel list
        MAXIS_case_number = objExcel.cells(excel_row, 3).Value   'reading the case number from Excel   
        MAXIS_case_number = Trim(MAXIS_case_number)
        
        If MAXIS_case_number = "" then   
            back_to_self
            Call navigate_to_MAXIS_screen("PERS", "    ")
            Call write_value_and_transmit(master_SMI_number, 17, 36)
            EmReadscreen PERS_screen_check, 4, 2, 47
            If PERS_screen_check = "PERS" then 
                EmReadscreen PERS_err, 75, 24, 2
                objExcel.cells(excel_row, 6).Value = trim(PERS_err)
            Elseif PERS_screen_check <> "PERS" then
                EmReadscreen match_screen, 4, 2, 51
                If match_screen = "MTCH" then 
                    EmReadscreen dupe_matches, 11, 9, 7
                    If trim(dupe_matches) <> "" then 
                        objExcel.cells(excel_row, 6).Value = "Duplicate exists. Add manually."
                    Else 
                        'if only one match exists then 
                        Call write_value_and_transmit("X", 8, 5)
                        EmReadscreen no_case_error, 75, 24, 2
                        If trim(no_case_error) = "PMI NBR ASSIGNED THRU SMI OR PMIN - NO MAXIS CASE EXISTS" then
                            objExcel.cells(excel_row, 6).Value = "NO MAXIS CASE EXISTS"
                        Else     
                            'read client name
                            EmReadscreen client_name, 39, 4, 8
                            client_name = trim(client_name)
                            objExcel.cells(excel_row, 5).Value = UCASE(client_name)
                            'read PMI
                            EmReadscreen DSPL_PMI, 8, 5, 44
                            objExcel.cells(excel_row, 4).Value = DSPL_PMI
                            'Read case number after finding HC case 
                            Call write_value_and_transmit("HC", 7, 22)
                            EmReadscreen DSPL_case_number, 8, 10, 6
                            If trim(DSPL_case_number) = "" then 
                                Call write_value_and_transmit("AP", 7, 22)
                                EmReadscreen DSPL_case_number, 8, 10, 6
                            End if 
                            objExcel.cells(excel_row, 3).Value = trim(DSPL_case_number)
                            objExcel.cells(excel_row, 6).Value = ""
                        End if 
                    End if 
                End if
            Else 
                objExcel.cells(excel_row, 6).Value = "Z - Match Error" 
            End if  
        End if 
        excel_row = excel_row + 1
        master_SMI_number = ""
        SMI_ECF_number = "" 
    Loop 
    
    FOR i = 1 to 6		'formatting the cells
       objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT
    
End if  

If AVS_option = "Output Forms Lists" then 
    'Setting up the array 
    DIM output_array()
    ReDim output_array(3, 0)
    
    const output_PMI_const          = 0
    const output_name_const         = 1
    const output_SMI_const          = 2
    const output_case_number_const  = 3
    
    entry_record = 0
    
    file_selection = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\AVS\AVS Forms Distribution Master List.xlsx"
    Call excel_open(file_selection, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    
    For Each objWorkSheet In objWorkbook.Worksheets
        If instr(objWorkSheet.Name, "Sheet") = 0 and objWorkSheet.Name <> "All AVS Forms" then months_list = months_list & objWorkSheet.Name & ","
    Next
    months_list = trim(months_list)  'trims excess spaces of months_list
    If right(months_list, 1) = "," THEN months_list = left(months_list, len(months_list) - 1) 'trimming off last comma
    array_of_months = split(months_list, ",")   'Creating new array
    
    For each month_sheet in array_of_months
        objExcel.worksheets(month_sheet).Activate 'Activates worksheet based on user selection
        excel_row = 2
            
        Do 
            output_PMI          = ObjExcel.Cells(excel_row, PMI_col).Value
            output_PMI = trim(output_PMI)
            If output_PMI = "" then exit do             
            output_name         = ObjExcel.Cells(excel_row, client_name_col).Value
            output_SMI          = ObjExcel.Cells(excel_row, SMI_col).Value
            output_case_number  = ObjExcel.Cells(excel_row, cn_col).Value
            output_form_date    = ObjExcel.Cells(excel_row, forms_col).Value
            
            If trim(output_form_date) = "" then
                ReDim Preserve output_array(3, entry_record)	'This resizes the array based on the number of rows in the Excel File'
                output_array(output_PMI_const,          entry_record) = trim(output_PMI)
                output_array(output_name_const,         entry_record) = trim(output_name)
                output_array(output_SMI_const,          entry_record) = trim(output_SMI)
                output_array(output_case_number_const,  entry_record) = trim(output_case_number)
                entry_record = entry_record + 1			'This increments to the next entry in the array'
            End if 
            
            STATS_counter = STATS_counter + 1
            excel_row = excel_row + 1
            'blanking out the variables 
            output_PMI          = ""
            output_name         = ""
            output_SMI          = ""
            output_case_number  = ""
        LOOP
    Next
    
    'Opening the Excel file
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    Set objWorkbook = objExcel.Workbooks.Add()
    objExcel.DisplayAlerts = True
    
    'Changes name of Excel sheet to "Case information"
    ObjExcel.ActiveSheet.Name = "Outstanding AVS Forms"
    
    'adding column header information to the Excel list
    ObjExcel.Cells(1, 1).Value = "PMI"
    ObjExcel.Cells(1, 2).Value = "Client name"
    ObjExcel.Cells(1, 3).Value = "SMI"
    ObjExcel.Cells(1, 4).Value = "Case Number"
    
    'formatting the cells
    FOR i = 1 to 4
    	objExcel.Cells(1, i).Font.Bold = True		'bold font
    	objExcel.Columns(i).AutoFit()				'sizing the columns
    NEXT
    
    excel_row = 2   'Staring row for Excel export 

    For item = 0 to UBound(output_array, 2)
        objExcel.Cells(excel_row, 1).Value = output_array(output_PMI_const,          item)
        objExcel.Cells(excel_row, 2).Value = output_array(output_name_const,         item)
        objExcel.Cells(excel_row, 3).Value = output_array(output_SMI_const,          item)
        objExcel.Cells(excel_row, 4).Value = output_array(output_case_number_const,  item)
        excel_row = excel_row + 1
    Next 
    
    FOR i = 1 to 4		'formatting the cells
    	objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT
    'Saves and closes the most recent Excel workbook with the Task based cases to process.
    objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\AVS\Recipients without AVS Forms.xlsx"  
End if  
    
If AVS_option = "Run Sync" then 
    file_selection = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\AVS\AVS Forms Distribution Master List.xlsx"
    Call excel_open(file_selection, True, True, ObjExcel, objWorkbook)  'opens the selected excel file'
    call AVS_sync      'Teating the AVS Sync
End if 
    
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! AVS list is complete.")