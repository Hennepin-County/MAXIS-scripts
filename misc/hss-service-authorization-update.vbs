'Required for statistical purposes===============================================================================
name_of_script = "MISC - HSS SERVICE AUTHORIZATION UPDATE.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 80                      'manual run time in seconds
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
		FuncLib_URL = "https://raw.githubusercontent.com/Hennepin-County/MAXIS-scripts/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"   'defaulting everything to Hennepin County Master Functions Libary.
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
call changelog_update("06/15/2021", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

function sort_dates(dates_array)
'--- Takes an array of dates and reorders them to be  .
'~~~~~ dates_array: an array of dates only
'===== Keywords: MAXIS, date, order, list, array
    dim ordered_dates ()
    redim ordered_dates(0)
    original_array_items_used = "~"
    days =  0
    do

        prev_date = ""
        original_array_index = 0
        for each thing in dates_array
            check_this_date = TRUE
            new_array_index = 0
            For each known_date in ordered_dates
                if known_date = thing Then check_this_date = FALSE
                new_array_index = new_array_index + 1
                ' MsgBox "known dates is " & known_date & vbNewLine & "thing is " & thing & vbNewLine & "match - " & check_this_date
            next
            ' MsgBox "known dates is " & known_date & vbNewLine & "thing is " & thing & vbNewLine & "check this date - " & check_this_date
            if check_this_date = TRUE Then
                if prev_date = "" Then
                    prev_date = thing
                    index_used = original_array_index
                Else
                    if DateDiff("d", prev_date, thing) < 0 then
                        prev_date = thing
                        index_used = original_array_index
                    end if
                end if
            end if
            original_array_index = original_array_index + 1
        next
        if prev_date <> "" Then
            redim preserve ordered_dates(days)
            ordered_dates(days) = prev_date
            original_array_items_used = original_array_items_used & index_used & "~"
            days = days + 1
        end if
        counter = 0
        For each thing in dates_array
            If InStr(original_array_items_used, "~" & counter & "~") = 0 Then
                For each new_date_thing in ordered_dates
                    If thing = new_date_thing Then
                        original_array_items_used = original_array_items_used & counter & "~"
                        days = days + 1
                    End If
                Next
            End If
            counter = counter + 1
        Next
        ' MsgBox "Ordered Dates array - " & join(ordered_dates, ", ") & vbCR & "days - " & days & vbCR & "Ubound - " & UBOUND(dates_array) & vbCR & "used list - " & original_array_items_used
    loop until days > UBOUND(dates_array)

    dates_array = ordered_dates
end function

'CONNECTS TO BlueZone
EMConnect ""
file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\DHS Housing Supports\HSS and SSR Reductions Real Time Data 0513_SIMPLE.xlsx" 'testing code
test_row = 2   'testing code 

'----------------------------------Set up code 
'Excel columns
const recip_PMI_col         = 1
const HSS_start_col         = 4
const HSS_end_col           = 5
const SA_number_col         = 9
const agreement_start_col   = 10
const agreement_end_col     = 11
const NPI_number_col        = 15
const HS_status_col         = 16
const faci_in_col           = 19
const faci_out_col          = 20
const impact_vendor_col     = 21
const case_status_col       = 26
const rate_reduction_col    = 27

'User interface dialog - There's just one in this script. 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 481, 90, "HSS SERVICE AUTHORIZATION UPDATE"
  ButtonGroup ButtonPressed
    PushButton 420, 45, 50, 15, "Browse...", select_a_file_button
    OkButton 365, 65, 50, 15
    CancelButton 420, 65, 50, 15
  EditBox 15, 45, 400, 15, file_selection_path
  Text 15, 20, 455, 20, "This script should be used when a list of recipients who have Supplemental Service Rate adjustments in MMIS due to overlapping Housing Stabilization Services (HSS)."
  Text 30, 70, 335, 10, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 465, 80, "Using this script:"
EndDialog

'Display dialog and dialog DO...Loop for mandatory fields and password prompting  
Do 
    Do
        err_msg = ""
        dialog Dialog1
        cancel_without_confirmation 
        If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
        If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue." 
        If err_msg <> "" Then MsgBox err_msg
    Loop until err_msg = ""
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call check_for_MMIS(False)             'Ensuring we're actually in MAXIS 

Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file

'Setting up the Excel spreadsheet
ObjExcel.Cells(1, rate_reduction_col).Value = "Rate Reduction Status"   'col 27

FOR i = 1 to 27		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

Dim adjustment_array()                        'Delcaring array
ReDim adjustment_array(rr_status_const, 0)     'Resizing the array to size of last const 

const recip_PMI_const       = 0         'creating array constants
const HSS_start_const       = 1
const HSS_end_const         = 2
const SA_number_const       = 3
const agreement_start_const = 4
const agreement_end_const   = 5
const npi_number_const      = 6
const HS_status_const       = 7
const faci_in_const         = 8
const faci_out_const        = 9
const impacted_vendor_const = 10
const case_status_const     = 11
const prev_start_const      = 12
const prev_end_const        = 13
const new_start_const       = 14
const new_end_const         = 15 
const excel_row_const       = 16
const rr_status_const       = 17

excel_row = test_row 'starting with the 1st non-header row :TESTING CODE 
entry_record = 0 'incrementor for the array 

Do
    recip_PMI = trim(objExcel.cells(excel_row, recip_PMI_col).Value)
    If recip_PMI = "" then exit do
    'removing preceeding 0's from the client PMI. This is needed to measure the PMI's on CASE/PERS. 
    Do 
		if left(recip_PMI, 1) = "0" then recip_PMI = right(recip_PMI, len(recip_PMI) -1)   'trimming off left-most 0 from recip_PMI
	Loop until left(recip_PMI, 1) <> "0"                                                      'Looping until 0's are all removed
    recip_PMI = trim(recip_PMI)
    
	'HSS_start       = trim(objExcel.cells(excel_row, HSS_start_col).Value)
    'HSS_end         = trim(objExcel.cells(excel_row, HSS_end_col).Value)
    
    SA_number       = trim(objExcel.cells(excel_row, SA_number_col).Value)
    SA_number = right("00000000" & SA_number, 11) 'ensures the variable is 11 digits. Inhibiting erorr 
    
    'agreement_start = trim(objExcel.cells(excel_row, agreement_start_col).Value)
    'agreement_end   = trim(objExcel.cells(excel_row, agreement_end_col).Value)
    'npi_number      = trim(objExcel.cells(excel_row, NPI_number_col).Value)
    'HS_status       = trim(objExcel.cells(excel_row, HS_status_col).Value)
    'impacted_vendor = trim(objExcel.cells(excel_row, impacted_vendor_col).Value)
    'case_status     = trim(objExcel.cells(excel_row, case_status_col).Value)

    'Adding recipient information to the array
    ReDim Preserve adjustment_array(rr_status_const, entry_record)	'This resizes the array based on the number of rows in the Excel File'
    
    adjustment_array(recip_PMI_const       , entry_record) = recip_PMI
    adjustment_array(HSS_start_const       , entry_record) = trim(objExcel.cells(excel_row, HSS_start_col).Value)
    adjustment_array(HSS_end_const         , entry_record) = trim(objExcel.cells(excel_row, HSS_end_col).Value)
    adjustment_array(SA_number_const       , entry_record) = SA_number
    adjustment_array(agreement_start_const , entry_record) = trim(objExcel.cells(excel_row, agreement_start_col).Value)
    adjustment_array(agreement_end_const   , entry_record) = trim(objExcel.cells(excel_row, agreement_end_col).Value) 
    adjustment_array(npi_number_const      , entry_record) = trim(objExcel.cells(excel_row, NPI_number_col).Value) 
    adjustment_array(HS_status_const       , entry_record) = trim(objExcel.cells(excel_row, HS_status_col).Value) 
    adjustment_array(faci_in_const         , entry_record) = trim(objExcel.cells(excel_row, faci_in_const).Value) 
    adjustment_array(faci_out_const        , entry_record) = trim(objExcel.cells(excel_row, faci_out_const).Value) 
    adjustment_array(impacted_vendor_const , entry_record) = trim(objExcel.cells(excel_row, impacted_vendor_col).Value) 
    adjustment_array(case_status_const     , entry_record) = trim(objExcel.cells(excel_row, case_status_col).Value) 
    adjustment_array(excel_row_const       , entry_record) = excel_row 
    entry_record = entry_record + 1			'This increments to the next entry in the array'
    stats_counter = stats_counter + 1
    excel_row = excel_row + 1    
Loop     
    
    '----------------------------------------------------------------------------------------------------CASE/PERS & PERS Search 
    Call navigate_to_MAXIS_screen_review_PRIV("CASE", "PERS", is_this_priv) 
    If is_this_priv = True then 
        case_status = "Privileged Case. Unable to access."
    Else 
        member_found = False 
        Do 
            Call navigate_to_MAXIS_screen("CASE", "PERS")
            row = 10    'staring row for 1st member 
            Do
                EMReadScreen person_PMI, 8, row, 34
                person_PMI = trim(person_PMI)
                If person_PMI = "" then exit do 
                'msgbox person_PMI & vbcr & recip_PMI
                If trim(person_PMI) = recip_PMI then 
                    EmReadscreen HS_status, 1, row, 66
                    If trim(HS_status) <> "" then
                        EmReadscreen member_number, 2, row, 3 
                        member_found = True 
                        'msgbox member_number & vbcr & HS_status
                        exit do 
                    Else 
                        'msgbox HS_status
                    End if 
                Else 
                    row = row + 3			'information is 3 rows apart. Will read for the next member. 
                    If row = 19 then
                        PF8
                        row = 10					'changes MAXIS row if more than one page exists
                    END if
                END if
            LOOP 
            
            If member_found = True then 
                exit do 
            Else 
                'msgbox member_found
                'try to match the correct case number in PERS search 
                back_to_self
                Call navigate_to_MAXIS_screen("PERS", "    ")
                Call write_value_and_transmit(recip_PMI, 15, 36)
                EmReadscreen mtch_screen, 4, 2, 51
                If mtch_screen = "MTCH" then 
                    EmReadscreen dupe_matches, 11, 9, 7
                    If trim(dupe_matches) <> "" then 
                        Case_status = "Duplicate exists. Add manually."
                    Else 
                        'if only one match exists then 
                        Call write_value_and_transmit("X", 8, 5)
                        EmReadscreen DSPL_PMI, 8, 5, 44
                        If trim(DSPL_PMI) = recip_PMI then 
                            'Read case number after finding HC case 
                            Call write_value_and_transmit("GR", 7, 22)
                            EmReadscreen DSPL_case_number, 8, 10, 6
                            If trim(DSPL_case_number) = "" then 
                                case_status = "Unable to find HS history for this member."
                            Else 
                                MAXIS_case_number = right("00000000" & DSPL_case_number, 8) 'outputs in 8 digits 
                                objExcel.Cells(excel_row, 2).Value = MAXIS_case_number 
                                objExcel.Cells(excel_row, 2).Interior.ColorIndex = 3	'Fills the row with red    
                                'msgbox MAXIS_case_number
                            End if         
                        Else 
                            case_status = "Unable to find resident by PMI in PERS/DSPL."
                        End if 
                    End if
                End if 
            End if 
        Loop 
        If trim(member_number) = "" then case_status = "Unable to locate case for member."
    End if 
    
    If trim(case_status) = "" then    
    '----------------------------------------------------------------------------------------------------FACI panel determination 
	   call navigate_to_MAXIS_screen("STAT", "FACI")
       EmWriteScreen member_number, 20, 76
       Call write_value_and_transmit("01", 20, 79)  'making sure we're on the 1st instance for member 
       'Based on how many FACI panels exist will determine if/how the information is read. 
	    EMReadScreen FACI_total_check, 1, 2, 78
        'msgbox FACI_total_check
	    If FACI_total_check = "0" then 
	    	case_status = "No FACI panel on this case for member #" & member_number & "."
	    Elseif FACI_total_check = "1" then 
            'just looking through a singular faci panel 
            EmReadscreen faci_name, 30, 6, 43
            faci_name = trim(replace(faci_name, "_", ""))   'faci name trimmed and replaced underscores 
            EmReadscreen vendor_number, 8, 5, 43
            vendor_number = trim(replace(vendor_number, "_", ""))   'vendor # trimmed and replaced underscores 
         
        	row = 18
	    	Do
                EMReadScreen faci_out, 10, row, 71      'faci out date
                If faci_out = "__ __ ____" then 
                    faci_out = ""                       'blanking out faci out if not a date 
                Else 
                    faci_out = replace(faci_out, " ", "/")  'reformatting to output with /, like dates do. 
                End if 
                EMReadScreen faci_in, 10, row, 47       'faci in date
                If faci_in = "__ __ ____" then 
                    faci_in = ""                        'blanking out faci in if not a date 
                Else 
                    faci_in = replace(faci_in, " ", "/")  'reformatting to output with /, like dates do. 
                End if 
	    		'msgbox "date out: " & faci_out 
	    		If faci_out = "" then 
					If faci_in = "" then  
                        row = row - 1   'no faci info on this row 
                    else 
                        If faci_in <> "" then exit do    'open ended faci found 
                    End if 
	    		Elseif faci_out <> "" then
                    If faci_in <> "" then exit do    'most recent faci span identified 
	    		End if 	
            Loop 
        Else    
            'msgbox "Multiple facilities found"
            objExcel.Cells(excel_row, faci_name_col).Interior.ColorIndex = 3	'Fills the row with red: TESTING CODE 
            'Evaluate multiple faci panels 
            faci_out_dates_string = ""                  'setting up blank string to increment
            current_faci_found = False                  'defaulting to false - this boolean will determine if evaluation of the last date is needed. Will become true statement if open-ended faci panel is detected.
            For item = 1 to FACI_total_check        
                
                Call write_value_and_transmit("0" & item, 20, 79)   'Entering the item's faci panel via direct navigation field on FACI panel. 
                row = 18
                Do
                    EMReadScreen faci_out, 10, row, 71      'faci out date
                    If faci_out = "__ __ ____" then 
                        faci_out = ""                       'blanking out faci out if not a date 
                    Else 
                        faci_out = replace(faci_out, " ", "/")  'reformatting to output with /, like dates do. 
                    End if 
                    EMReadScreen faci_in, 10, row, 47       'faci in date
                    If faci_in = "__ __ ____" then 
                        faci_in = ""                        'blanking out faci in if not a date 
                    Else 
                        faci_in = replace(faci_in, " ", "/")  'reformatting to output with /, like dates do. 
                    End if 
                    
                    EmReadscreen faci_name, 30, 6, 43
                    faci_name = trim(replace(faci_name, "_", ""))   'faci name trimmed and replaced underscores 
                    EmReadscreen vendor_number, 8, 5, 43
                    vendor_number = trim(replace(vendor_number, "_", ""))   'vendor # trimmed and replaced underscores 
                    'Reading the faci in and out dates 
                    If faci_out = "" then 
                        If faci_in = "" then  
                            row = row - 1   'no faci info on this row - this is blank 
                        else 
                            If faci_in <> "" then 
                                current_faci_found = True   'Condition is met so date evaluation via adjustment_array is not needed. 
                                exit do    'open ended faci found 
                            End if 
                        End if 
                    Elseif faci_out <> "" then
                        If faci_in <> "" then 
                            faci_out_dates_string = faci_out_dates_string & faci_out & "|"
                            
                            Redim Preserve adjustment_array(faci_out_const, faci_count)
                            adjustment_array(vendor_number_const, faci_count) = vendor_number
                            adjustment_array(faci_name_const,     faci_count) = faci_name 
                            adjustment_array(faci_in_const,       faci_count) = faci_in
                            adjustment_array(faci_out_const,      faci_count) = faci_out 
                            faci_count = faci_count + 1
                            exit do    'most recent faci span identified 
                        End if 
                    End if 	
                Loop 
                If current_faci_found = True then exit for  'exiting the for since most current FACI has been found 
            Next 
            
            'If an open-ended faci is NOT found, then futher evaluation is needed to determine the most recent date. 
            If current_faci_found = False then
                faci_out_dates_string = left(faci_out_dates_string, len(faci_out_dates_string) - 1)
                'msgbox faci_out_dates_string 
                faci_out_dates = split(faci_out_dates_string, "|")
                call sort_dates(faci_out_dates)
                first_date = faci_out_dates(0)                              'setting the first and last check dates
                last_date = faci_out_dates(UBOUND(faci_out_dates))
                'MsgBox first_date & vbcr & last_date 
                
                'finding the most recent date if none of the dates are open-ended 
                For item = 0 to Ubound(adjustment_array, 2)
                    If adjustment_array(faci_out_const, item) = last_date then 
                        vendor_number   = adjustment_array(vendor_number_const, item)
                        faci_name       = adjustment_array(faci_name_const, item)
                        faci_in         = adjustment_array(faci_in_const, item)
                        faci_out        = adjustment_array(faci_out_const, item)
                    End if 
                Next
                'msgbox "**Facility Info**" & vbcr & vbcr & vendor_number & vbcr & faci_name & vbcr & faci_in & vbcr & faci_out     
            End if 
            ReDim adjustment_array(faci_out_const, 0)     'Resizing the array back to original size
            Erase adjustment_array                        'then once resized it gets erased. 
	    End if 
        
        'msgbox "**Facility Info**" & vbcr & vbcr & vendor_number & vbcr & faci_name & vbcr & faci_in & vbcr & faci_out
        
        '----------------------------------------------------------------------------------------------------VNDS/VND2
        Call Navigate_to_MAXIS_screen("MONY", "VNDS")
        Call write_value_and_transmit(vendor_number, 4, 59)
        Call write_value_and_transmit("VND2", 20, 70)
        EMReadScreen VND2_check, 4, 2, 54
        If VND2_check <> "VND2" then 
            case_status = "Unable to find MONY/VND2 panel"
        Else 
            health_depart_reason = False    'defalthing to false 
            exemption_reason = False
            
            EmReadscreen exemption_code, 2, 9, 69
            If exemption_code = "__" then exemption_code = ""
            EmReadscreen HDL_one, 2, 10, 69
            EmReadscreen HDL_two, 2, 10, 72
            EmReadscreen HDL_three, 2, 10, 75
            If HDL_one = "__" then HDL_one = ""
            If HDL_two = "__" then HDL_two = ""
            If HDL_three = "__" then HDL_three = ""
            HDL_string = HDL_one & "|" & HDL_two & "|" & HDL_three
            
            HDL_applicable_codes = "08,09,10"
            
            If instr(HDL_applicable_codes, HDL_one) then health_depart_reason = True 
            If instr(HDL_applicable_codes, HDL_two) then health_depart_reason = True 
            If instr(HDL_applicable_codes, HDL_three) then health_depart_reason = True 
            
            If exemption_code = "15" or exemption_code = "26" or exemption_code = "28" then 
                exemption_reason = True 
                'msgbox "exemption_reason: " & exemption_reason
            Else 
                exmption_reason = False 
                'msgbox "exemption_reason: " & exemption_reason
            End if 
            
            If exemption_code = "28" and instr(HDL_string, "10") then 
                impacted_vendor = "No"
            Else 
                If (exemption_reason = True and health_depart_reason = True) then 
                    impacted_vendor = "Yes" 
                Else 
                    impacted_vendor = "No"
                End if 
            End if 
        End if
    End if 
    
    'msgbox "**Vendor Info**" & vbcr & vbcr & "exemption_code: " & exemption_code & vbcr & "HDL Codes: " & HDL_string & vbcr & "exmption_reason: " & exemption_reason & vbcr & "health_depart_reason: " & health_depart_reason & vbcr & "impacted_vendor: " &impacted_vendor
    
    'outputting to Excel 
    ObjExcel.Cells(excel_row, HS_status_col).Value   = HS_status
    ObjExcel.Cells(excel_row, vendor_num_col).Value  = vendor_number
    ObjExcel.Cells(excel_row, faci_name_col).Value   = faci_name
    ObjExcel.Cells(excel_row, faci_in_col).Value     = faci_in
    ObjExcel.Cells(excel_row, faci_out_col).Value    = faci_out
    ObjExcel.Cells(excel_row, impact_vnd_col).Value  = impacted_vendor
    ObjExcel.Cells(excel_row, exempt_code_col).Value = exemption_code
    ObjExcel.Cells(excel_row, HDL_one_col).Value     = HDL_one
    ObjExcel.Cells(excel_row, HDL_two_col).Value     = HDL_two
    ObjExcel.Cells(excel_row, HDL_three_col).Value   = HDL_three
    ObjExcel.Cells(excel_row, case_status_col).Value = case_status

    'Blanking out variables at the end of the loop 
    HS_status = ""
    vendor_number = ""
    faci_name = ""
    faci_in = ""
    faci_out = ""
    impacted_vendor = ""
    exemption_code = ""
    HDL_one = ""
    HDL_two = ""
    HDL_three = ""
    case_status = ""
    

'formatting the cells
FOR i = 1 to 27
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure_with_error_report("Success! The service authorizations have been updated. Please review the worksheet for manual updates.")