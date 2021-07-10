'Required for statistical purposes===============================================================================
name_of_script = "MISC - HSS SERVICE AUTHORIZATION UPDATE.vbs"
start_time = timer
STATS_counter = 1                          'sets the stats counter at one
STATS_manualtime = 800                      'manual run time in seconds
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

'TODO: test this function, confirm it's working 
'TODO: Make new issue for FuncLib
'TODO: Once new code is updated in Funclib, remove function and test variable 
function check_for_MMIS_test(end_script)
'--- This function checks to ensure the user is in a MMIS panel
'~~~~~ end_script: If end_script = TRUE the script will end. If end_script = FALSE, the user will be given the option to cancel the script, or manually navigate to a MMIS screen.
'===== Keywords: MMIS, production, script_end_procedure
	Do
		transmit
		row = 1
		col = 1
		EMSearch "MMIS", row, col
		IF row <> 1 then
			If end_script = True then
				script_end_procedure("You do not appear to be in MMIS. You may be passworded out. Please check your MMIS screen and try again.")
			Else
                Dialog1 = ""
                BeginDialog Dialog1, 0, 0, 216, 55, "MMIS Dialog"
                ButtonGroup ButtonPressed
                OkButton 125, 35, 40, 15
                CancelButton 170, 35, 40, 15
                Text 5, 5, 210, 25, "You do not appear to be in MMIS. You may be passworded out. Please check your MMIS screen and try again, or press CANCEL to exit the script."
                EndDialog
                Do
                    Dialog Dialog1
                    cancel_without_confirmation
                Loop until ButtonPressed = -1
			End if
		End if
	Loop until row = 1
end function

'TODO: test this function, confirm it's working 
'TODO: Make new issue for FuncLib
'TODO: Once new code is updated in Funclib, remove function and test variable 
function ONLY_create_MAXIS_friendly_date_test(date_variable)
'--- This function creates a MM DD YY date.
'~~~~~ date_variable: the name of the variable to output
    date_variable = dateadd("d", 0, date_variable)    'janky way to convert to a date, but hey it works.    
    var_month     = right("0" & DatePart("m",    date_variable), 2) 
    var_day       = right("0" & DatePart("d",    date_variable), 2)
    var_year      = right("0" & DatePart("yyyy", date_variable), 2)
	date_variable = var_month &"/" & var_day & "/" & var_year
    'msgbox "date_variable: " & date_variable
end function

'CONNECTS TO BlueZone
EMConnect ""
Check_for_MMIS_test(false)   'checking for, and allowing user to navigate into MMIS. 
file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\DHS Housing Supports\HSS and SSR Reductions Real Time Data 0701.xlsx" 'testing code
test_row = 2   'testing code 

'----------------------------------Set up code 
'Excel columns
const recip_PMI_col         = 1     'Col A     
const HSS_start_col         = 7     'Col G     
const HSS_end_col           = 8     'Col H     
const SA_number_col         = 9     'Col I     
const agreement_start_col   = 10    'Col J     
const agreement_end_col     = 11    'Col K   
const rate_amt_col          = 13    'Col M  
const NPI_number_col        = 15    'Col O     
const HS_status_col         = 16    'Col P     
const faci_in_col           = 19    'Col Q     
const faci_out_col          = 20    'Col R     
const impacted_vendor_col   = 21    'Col S     
const case_status_col       = 26    'Col Z     
const rate_reduction_col    = 27    'Col AA     

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

'testing code: reactivate dialog before release 
''Display dialog and dialog DO...Loop for mandatory fields and password prompting  
'Do 
'    Do
'        err_msg = ""
'        dialog Dialog1
'        cancel_without_confirmation 
'        If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx")
'        If trim(file_selection_path) = "" then err_msg = err_msg & vbcr & "* Select a file to continue." 
'        If err_msg <> "" Then MsgBox err_msg
'    Loop until err_msg = ""
'    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
'Loop until are_we_passworded_out = false					'loops until user passwords back in

Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file

'Setting up the Excel spreadsheet
ObjExcel.Cells(1, rate_reduction_col).Value = "Rate Reduction Status"   'col 27

'formatting the cells'
objExcel.Cells(1, 27).Font.Bold = True		'bold font'
objExcel.Columns(27).ColumnWidth = 120		'sizing the last column 

Dim adjustment_array()                        'Delcaring array
ReDim adjustment_array(rr_status_const, 0)     'Resizing the array to size of last const 

const recip_PMI_const               = 0         'creating array constants
const HSS_start_const               = 1
const HSS_end_const                 = 2
const SA_number_const               = 3
const agreement_start_const         = 4
const agreement_end_const           = 5
const npi_number_const              = 6
const HS_status_const               = 7
const faci_in_const                 = 8
const faci_out_const                = 9
const impacted_vendor_const         = 10
const case_status_const             = 11
const prev_start_const              = 12
const prev_end_const                = 13
const new_start_const               = 14
const new_end_const                 = 15 
const excel_row_const               = 16
const MAXIS_note_conf_const         = 17
const MMIS_note_conf_const          = 18
const reduce_rate_const             = 19
const adjustment_start_date_const   = 20
const passed_case_tests_const       = 21
const duplicate_agreements_const    = 22
const pmi_count_const               = 23
const rate_amt_const                = 24
const rr_status_const               = 25

excel_row = test_row 'starting with the 1st non-header row :TESTING CODE 
entry_record = 0 'incrementor for the array 

Do
    recip_PMI = trim(objExcel.cells(excel_row, recip_PMI_col).Value)
    If recip_PMI = "" then exit do
        
    SA_number       = trim(objExcel.cells(excel_row, SA_number_col).Value)
    SA_number = right("00000000" & SA_number, 11) 'ensures the variable is 11 digits. Inhibiting erorr 

    'Adding recipient information to the array
    ReDim Preserve adjustment_array(rr_status_const, entry_record)	'This resizes the array based on the number of rows in the Excel File'
    adjustment_array(recip_PMI_const            , entry_record) = recip_PMI
    adjustment_array(HSS_start_const            , entry_record) = trim(objExcel.cells(excel_row, HSS_start_col).Value)
    adjustment_array(HSS_end_const              , entry_record) = trim(objExcel.cells(excel_row, HSS_end_col).Value)
    adjustment_array(SA_number_const            , entry_record) = SA_number
    adjustment_array(agreement_start_const      , entry_record) = trim(objExcel.cells(excel_row, agreement_start_col).Value)
    adjustment_array(agreement_end_const        , entry_record) = trim(objExcel.cells(excel_row, agreement_end_col).Value) 
    adjustment_array(npi_number_const           , entry_record) = trim(objExcel.cells(excel_row, NPI_number_col).Value) 
    adjustment_array(HS_status_const            , entry_record) = trim(objExcel.cells(excel_row, HS_status_col).Value) 
    adjustment_array(faci_in_const              , entry_record) = trim(objExcel.cells(excel_row, faci_in_col).Value) 
    adjustment_array(faci_out_const             , entry_record) = trim(objExcel.cells(excel_row, faci_out_col).Value) 
    adjustment_array(impacted_vendor_const      , entry_record) = trim(objExcel.cells(excel_row, impacted_vendor_col).Value) 
    adjustment_array(case_status_const          , entry_record) = trim(objExcel.cells(excel_row, case_status_col).Value) 
    adjustment_array(rate_amt_const             , entry_record) = trim(objExcel.cells(excel_row, rate_amt_col).Value) 
    adjustment_array(excel_row_const            , entry_record) = excel_row 
    adjustment_array(duplicate_agreements_const , entry_record) = False 
    adjustment_array(passed_case_tests_const    , entry_record) = False 'defaulting to false
    adjustment_array(MAXIS_note_conf_const      , entry_record) = False 'defaulting to false
    adjustment_array(MMIS_note_conf_const       , entry_record) = False 'defaulting to false
    adjustment_array(reduce_rate_const          , entry_record) = False 'defaulting to false
    entry_record = entry_record + 1			'This increments to the next entry in the array'
    stats_counter = stats_counter + 1
    excel_row = excel_row + 1  
    recip_PMI = ""  'Blanking out variables for next loop         
    SA_number = ""  'Blanking out variables for next loop 
Loop

'----------------------------------------------------------------------------------------------------determine which rows of information are going to have a rate reduction or not.
For item = 0 to Ubound(adjustment_array, 2)
    'Determining which date to use to end/start the agreements. Initial conversion date is 07/01/21. We cannot use a date earlier than this. If a date is earlier than this, the date is 07/01/21.
    'This supports both the initial conversion and ongoing cases. 
    'msgbox adjustment_array(HSS_start_const, item)
    If DateDiff("d", #07/01/21#, adjustment_array(HSS_start_const, item)) <= 0 then 
        'if this date is a negative or a date before 07/01/21 (past date), then use 07/01/21.
        new_agreement_start_date = #07/01/21#
        'msgbox "using July 1"
    Else   
        Call ONLY_create_MAXIS_friendly_date_test(adjustment_array(HSS_start_const, item))
        ''using the HSS start date as this is after 07/01/21 (future date from initial coversion date of 07/01/21)
        'agreement_day   = right("0" & DatePart("d",    adjustment_array(HSS_start_const, item)), 2)
        'agreement_month = right("0" & DatePart("m",    adjustment_array(HSS_start_const, item)), 2)
        'agreement_yr    = right(      DatePart("yyyy", adjustment_array(HSS_start_const, item)), 2)
        '
        'new_agreement_start_date = agreement_day & "/" & agreement_month & "/" & agreement_yr
        'new_agreement_start_date = dateadd("d", 0, new_agreement_start_date)    'janky way to convert to a date, but hey it works.     
    End if 
    
    Call ONLY_create_MAXIS_friendly_date_test(new_agreement_start_date)
    
    adjustment_array(adjustment_start_date_const, item) = new_agreement_start_date
    'msgbox "new agreement start date: " & new_agreement_start_date
    
    'Finding facility panels that may have ended before the HSS start date
    active_facility = False     'default value 
    If (adjustment_array(faci_in_const, item) <> "" and adjustment_array(faci_out_const, item) = "") then 
        active_facility = True 
    ElseIf adjustment_array(faci_out_const, item) <> "" then 
        If DateDiff("d", adjustment_array(faci_out_const, item), adjustment_array(adjustment_start_date_const, item)) <= 0 then 
            'Facility end date is NOT before the agreement start date. 
            active_facility = True 
        End if 
    End if
    
    rate_reduction_status = "Failed Case Test(s): "
    'Setting up initial tests
    'Rows with Case Status of “Unable to find MONY/VND2 panel”
    'Rows with Case Status of “Privileged Case. Unable to access.”
    'Row’s that have more than one MAXIS case identified, and HS is not active for the recipient on that case.
    'Row’s that are not identified as an Impacted Vendor (“Yes”)
    'Open-ended facility spans or recipients that have faci panels that close after the HSS start date. 
    'Rate costs that are not 15.87
    If (adjustment_array(case_status_const, item) = "" and _
        adjustment_array(HS_status_const, item) <> "" and _
        adjustment_array(impacted_vendor_const, item) = "Yes" and _
        adjustment_array(rate_amt_const, item) = "15.87" and _
        active_facility = True) then 
        adjustment_array(passed_case_tests_const, item) = True
    Else 
    'Failure Reasons 
        If adjustment_array(HS_status_const, item) = "" then rate_reduction_status = rate_reduction_status & "No HS Status in MAXIS Case. "
        If adjustment_array(impacted_vendor_const, item) = "Yes" and adjustment_array(rate_amt_const, item) <> "15.87" then rate_reduction_status = rate_reduction_status & "Rate is not 15.87, review manually. "
        If adjustment_array(impacted_vendor_const, item) <> "Yes" then rate_reduction_status = rate_reduction_status & "Not an impacted vendor. "
        If active_facility = False then rate_reduction_status = rate_reduction_status & "Not an active facility. "
        If adjustment_array(case_status_const, item) <> "" then rate_reduction_status = rate_reduction_status & adjustment_array(case_status_const, item)
    End if 
    If rate_reduction_status <> "Failed Case Test(s): " then adjustment_array(rr_status_const, item) = rate_reduction_status
    'msgbox adjustment_array(excel_row_const, item) & vbcr & adjustment_array(passed_case_tests_const, item)
Next  

'If duplicates still exist after the intital case tests, then these need to be figured out manually at this point.
For item = 0 to Ubound(adjustment_array, 2)
    recip_PMI = adjustment_array(recip_PMI_const, item)    
    PMI_count = 0  
    For i = 0 to Ubound(adjustment_array, 2)
        If recip_PMI = adjustment_array(recip_PMI_const, i) then 
            If adjustment_array(passed_case_tests_const, i) = True then PMI_count = PMI_count + 1
        End if 
    Next     

    adjustment_array(pmi_count_const, item) = PMI_count
    'msgbox PMI_Count & vbcr & adjustment_array(pmi_count_const, item)
Next

For item = 0 to Ubound(adjustment_array, 2)
    If adjustment_array(pmi_count_const, item) > 1 then 
        adjustment_array(duplicate_agreements_const, item) = True
        If adjustment_array(passed_case_tests_const, item) = True then 
            adjustment_array(passed_case_tests_const, item) = False
            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & "Duplicate agreements found. Review manually."
            'msgbox adjustment_array(excel_row_const, item) & vbcr & adjustment_array(passed_case_tests_const, item) & vbcr & adjustment_array(rr_status_const, item)
        End if 
    End if 
 
    If adjustment_array(passed_case_tests_const, item) = True then adjustment_array(reduce_rate_const, item) = True 
    
    objExcel.Cells(adjustment_array(excel_row_const, item), rate_reduction_col).Value = adjustment_array(rr_status_const, item)   'testing code
    'msgbox excel_row & "passed_case_tests: " & passed_case_tests                                                            'testing code
    rate_reduction_status = ""
Next 

'----------------------------------------------------------------------------------------------------MMIS STEPS 
Call Check_for_MMIS_test(False)

For item = 0 to Ubound(adjustment_array, 2)
    If adjustment_array(reduce_rate_const, item) = True then
        'start the rate reductions in MMIS 
        Call navigate_to_MMIS_region("GRH UPDATE")	'function to navigate into MMIS, select the GRH update realm, and enter the prior authorization area
        Call MMIS_panel_confirmation("AKEY", 51)				'ensuring we are on the right MMIS screen
        EmWriteScreen "C", 3, 22
        Call write_value_and_transmit(adjustment_array(SA_number_const, item), 9, 36) 'Entering Service Authorization Number and transmit to ASA1
        EmReadscreen current_panel, 4, 1, 51 
        If current_panel = "AKEY" then 
            EmReadscreen error_message, 80, 24, 2    
            adjustment_array(reduce_rate_const, item) = False
            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Authorization Number is not valid."
            error_message = ""
            msgbox "Failed! Authorization Number is not valid."
        Else 
            EMReadScreen AGMT_STAT, 1, 3, 17
            If AGMT_STAT <> "A" then 
                'msgbox "AGMT_STAT: " & AGMT_STAT
                adjustment_array(reduce_rate_const, item) = False
                adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Authorization Status is coded as: " & AGMT_STAT & "."
                msgbox "Failed! Authorization Status is coded as: " & AGMT_STAT & "."
            Else 
                EmWriteScreen "S", 3, 17
                PF3     'to AKEY screen 
                EmReadscreen current_panel, 4, 1, 51 
                'msgbox current_panel
                If current_panel <> "AKEY" then
                    adjustment_array(reduce_rate_const, item) = False
                    adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Unknown issue occured after changeing AGMT STAT on ASA1."
                    msgbox "Failed! Unknown issue occured after changeing AGMT STAT on ASA1."
                Else 
                    transmit 'to ASA1 
                    Call write_value_and_transmit("ASA3", 1, 8)             'Direct navigate to ASA3
                    Call MMIS_panel_confirmation("ASA3", 51)				'ensuring we are on the right MMIS screen
                    
                    'Checking Line 2 to ensure it's blank
                    EmReadscreen line_2_check, 6, 14, 60
                    If trim(line_2_check) <> "" then 
                        EmWriteScreen "A", 3, 20   'Restoring the original approving the agreement on ASA3 in AGMT/TYPE STAT field
                        PF3
                        adjustment_array(reduce_rate_const, item) = False
                        adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Agreement already exists in Line 2. Review Manually."
                        msgbox "Failed! Agreement already exists in Line 2. Review Manually."
                    Else 
                        'Reading and converting start and end dates 
                        'agreement start date 
                        EMReadScreen start_month, 2, 8, 60
                        EMReadScreen start_day, 2, 8, 62
                        EMReadScreen start_year, 2, 8, 64
                        Line_1_start_date = start_month & "/" & start_day & "/" & start_year
                        
                        'agreement end date - original end date from line 1
                        EMReadScreen end_month, 2, 8, 67
                        EMReadScreen end_day, 2, 8, 69
                        EMReadScreen end_year, 2, 8, 71
                        original_end_date = end_month & "/" & end_day & "/" & end_year
                        Call ONLY_create_MAXIS_friendly_date_test(original_end_date)
                        write_original_end_date = replace(original_end_date, "/", "")  'for line 2
                        'msgbox "original_end_date : " & original_end_date & vbcr & "write_original_end_date :" & write_original_end_date
                        
                        'Failing cases that the end date is less than the new agreement start date
                        If DateDiff("d", adjustment_array(adjustment_start_date_const, item), original_end_date) <= 0 then 
                            'if this date is a positive then its a date before the HSS start date and needs to fail.
                            'msgbox "DateDiff" & DateDiff("d", original_end_date, adjustment_array(adjustment_start_date_const, item))
                            EmWriteScreen "A", 3, 20   'Restoring the original approving the agreement on ASA3 in AGMT/TYPE STAT field
                            PF3
                            adjustment_array(reduce_rate_const, item) = False
                            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Agreement end date (" & original_end_date & ") is < HSS start date (" & adjustment_array(adjustment_start_date_const, item) & ")."
                            msgbox "Failed! Agreement end date (" & original_end_date & ") is < HSS start date (" & adjustment_array(adjustment_start_date_const, item) & ")."
                        Else     
                            'Creating a date that is the day before the HSS start date/conversion date - for LINE 1
                            new_line_1_end_date = dateadd("d", -1, adjustment_array(adjustment_start_date_const, item)) 
                            'using the HSS start date as this is after 07/01/21 (future date from initial coversion date of 07/01/21)
                            Call ONLY_create_MAXIS_friendly_date_test(new_line_1_end_date)
                            'msgbox new_line_1_end_date
                            'removing date formatting for ASA3 input 
                            write_new_line_1_end_date = replace(new_line_1_end_date, "/", "")
                            
                            line_1_total_units = datediff("d", Line_1_start_date, new_line_1_end_date) + 1
                            'msgbox "Line_1_start_date: " & Line_1_start_date & vbcr & "new_line_1_end_date: " & new_line_1_end_date & vbcr & "line_1_total_units: " & line_1_total_units
                            
                            'Unable to close agreements that have been overbilled by the facility. 
                            over_billed = True      'Defaulting to True 
                            EmReadscreen billed_units, 6, 11, 60
                            billed_units = trim(billed_units)
                            If trim(billed_units) = "" then 
                                over_billed = False   'no billing exists - blank                       
                            ElseIf cint(billed_units) = cint(billed_units) then 
                                over_billed = False 'facility only billed up to the amount of the date we are closing this agreement date. 
                            Elseif cint(billed_units) < cint(billed_units) then 
                                over_billed = False  'facility billed less than the amount of the date we are closing this agreement date. 
                            End if 
                        
                            If over_billed = True then 
                                msgbox "Faci overbilled. billed_units: " & billed_units & vbcr & "line_1_total_units: " & line_1_total_units
                                EmWriteScreen "A", 3, 20   'Restoring the original approving the agreement on ASA3 in AGMT/TYPE STAT field
                                PF3
                                'msgbox "too many billed units"
                                adjustment_array(reduce_rate_const, item) = False
                                adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & " Unable to reduce Line 1 agreement due to overbilling. Billed units: & " & billed_units & " vs. " & line_1_total_units & "."
                                msgbox "Failed! Unable to reduce Line 1 agreement due to overbilling. Billed units: & " & billed_units & " vs. " & line_1_total_units & "."
                            Else     
                                '----------------------------------------------------------------------------------------------------Updating LINE 1 agreement
                                EmWriteScreen write_new_line_1_end_date, 8, 67
                                Call clear_line_of_text(9, 60)
                                EmWriteScreen line_1_total_units, 9, 60
                                
                                Msgbox "Final Check on Line 1"
                                
                                'PF3 '	to save changes
                                'EMReadscreen error_message, 20, 24, 2    'Any number of issues (duplicate PMI, ssrt charged more units than stay, etc.). These cases require manual review if error occurs. 
                                'If trim(error_message) <> "ACTION COMPLETED" then
                                '    EmWriteScreen "A", 3, 20   'Restoring the original approving the agreement on ASA3 in AGMT/TYPE STAT field
                                '    PF3
                                '    adjustment_array(reduce_rate_const, item) = False
                                '    adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Failure after updating Line 1. Error msg: " & trim(error_message)
                                '    msgbox "Failed! Failure after updating Line 1. Error msg: " & trim(error_message)
                                'Else 
                                    'msgbox "Line 1 updates saved, moving on to Line 2"
                                    'transmit 'to ASA1 
                                    'Call write_value_and_transmit("ASA3", 1, 8)             'Direct navigate to ASA3
                                    'Call MMIS_panel_confirmation("ASA3", 51)				'ensuring we are on the right MMIS screen
                                    '----------------------------------------------------------------------------------------------------Entering LINE 2 Information 
                                    EmWriteScreen "H0043", 13, 36
                                    EmWriteScreen "U5", 13, 44
                                    
                                    write_new_agrement_start_date = replace(adjustment_array(adjustment_start_date_const, item), "/", "")
                                    
                                    EmWriteScreen write_new_agrement_start_date, 14, 60
                                    EmWriteScreen write_original_end_date, 14, 67
                                    
                                    msgbox "write_new_agrement_start_date: " & write_new_agrement_start_date & vbcr & "write_original_end_date: " & write_original_end_date
                                    
                                    EmReadscreen old_rate, 5, 9, 24
                                    new_rate = old_rate / 2 'divide total by two, and round to integer
                                    new_rate = Round(new_rate, 2) 'round to two decimal places 
                                    EmWriteScreen new_rate, 15, 20
                                    
                                    msgbox "new_rate: " & new_rate 
                                    
                                    line_2_total_units = datediff("d", adjustment_array(adjustment_start_date_const, item), original_end_date) + 1
                                    EmWriteScreen line_2_total_units, 15, 60
                                    msgbox "line_2_total_units: " & line_2_total_units & vbcr & "start date: " & adjustment_array(adjustment_start_date_const, item) & vbcr & "original_end_date: " & original_end_date
                            
                                    EMReadscreen agreement_NPI_number, 10, 10, 20   'Reading line 1 NPI Number 
                                    EmWriteScreen agreement_NPI_number, 16, 20      'Enetering NPI in Line 2 agreement 
                                    
                                    EmWriteScreen new_rate, 17, 20  
                                    EmWriteScreen "MM", 17, 35      'TODO: This is crossed out in the instructions, but is inhibiting the script without it. 
                                    
                                    msgbox "agreement_NPI_number: " & agreement_NPI_number
                                    
                                    EmWriteScreen "A", 18, 19   'Approving the agreement on ASA3 in STAT CD/DATE field         
                                    EmWriteScreen "A", 3, 20   'Approving the agreement on ASA3 in AGMT/TYPE STAT field 
                                    transmit 
                                    Msgbox "Final Check on Line 2"
                                    
                                    'PF3 ' to save
                                    EMReadScreen PPOP_check, 4, 1, 52
                                    If PPOP_check = "PPOP" then 
                                        msgbox PPOP_check
                                        'script_end_procedure("PPOP Screen - FYCO this.")
                                    End if 
                                    PF6
                                    EmReadscreen current_panel, 4, 1, 51 
                                    If current_panel = "AKEY" then 
                                        EmReadscreen error_message, 80, 24, 2   
                                        If trim(error_message) = "ACTION COMPLETED" then  
                                            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & "Agreement successfully reduced to " & new_rate & "."
                                        Else 
                                            adjustment_array(reduce_rate_const, item) = False
                                            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & trim(error_message)
                                        End if 
                                    Else 
                                        EmReadscreen error_message, 80, 21, 2       'reading error message on any other screen.    
                                        adjustment_array(reduce_rate_const, item) = False
                                        adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & trim(error_message)
                                        PF3 ' to save
                                    End if
                                'End if
                            End if 
                        End if 
                    End if 
                End if
            End if
        End if 
    End if 
    'TODO: Blank out all the variables before NEXT
    error_message = ""
Next 

'Excel output of rate reduction statuses 
For item = 0 to Ubound(adjustment_array, 2)
    objExcel.Cells(adjustment_array(excel_row_const, item), rate_reduction_col).Value = adjustment_array(rr_status_const, item) 'testing code: remove and output at the end of the run for release.
Next 
note_string  = date & " - DHS SUPPLEMENTAL SERVICE RATE ADJUSTMENT", & _             
    "THERE IS AN ACTIVE HOUSING SUPPORT SUPPLEMENTAL SERVICE RATE (SSR)", & _ 
    "SERVICE AUTHORIZATION IN MMIS FOR THIS MAXIS CASE. DHS ADJUSTED THE", & _ 
    "MMIS SERVICE AUTHORIZATION(S) FOR HOUSING SUPPORT SSR THROUGH THE", & _ 
    "EXISITING END DATE OF THE SERVICE AUTHORIZATION.", & _ 
    "REVISIONS ARE BASED ON A DETERMINATION OF THE RECIPIENT'S CONCURRENT", & _ 
    "ELIGBILITY HOUSING STABILIZATION SERVICES. MMIS ISSUED A REVISED", & _ 
    "SERVICE AUTORIZATION WITH THE CORRECT SSR PER DIEM TO THE HOUSING", & _ 
    "SUPPORT PROVIDER ASSOCIATED WITH THE MMIS SERVICE AUTHORIZATION.", & _ 
    "ELIGIBILITY WORKERS DO NOT NEED TO TAKE ANY ACTION IN MAXIS.", & _ 
    "**********************************************************************"

note_array = split(note_string, ",")

'----------------------------------------------------------------------------------------------------DHS NOTES on ADHS screen in GRHU realm 
For item = 0 to Ubound(adjustment_array, 2)
    If adjustment_array(reduce_rate_const, item) = True then 
        'start the rate reductions in MMIS 
        Call navigate_to_MMIS_region("GRH UPDATE")	'function to navigate into MMIS, select the GRH update realm, and enter the prior authorization area
        Call MMIS_panel_confirmation("AKEY", 51)				'ensuring we are on the right MMIS screen
        EmWriteScreen "C", 3, 22
        Call write_value_and_transmit(adjustment_array(SA_number_const, item), 9, 36) 'Entering Service Authorization Number and transmit to ASA1
        Call MMIS_panel_confirmation("ASA1", 51)				'ensuring we are on the right MMIS screen
        Call write_value_and_transmit("ADHS", 1, 8)
        Call MMIS_panel_confirmation("ADHS", 51)				'ensuring we are on the right MMIS screen    
        row = 6
        Do 
            EmReadscreen blank_row_check, 6, row, 3
            If trim(blank_row_check) = "" then 
                exit do 
            Else 
                row = row + 1
            End if 
        Loop 
        
        row = row 
        For each note in note_array 
            EmWriteScreen note, row 3
            row = row + 1
            If row = 14 then 
                transmit
                row = 6
            End if
        Next     
        
        PF3 
        error_message = ""
        EmReadscreen error_message, 40, 24, 2
        If trim(error_message) =  "ACTION COMPLETED" then 
            adjustment_array(MMIS_note_conf_const, item) = True
        Else 
            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Unable to enter note on ADHS - " & trim(error_message)
        End if  
    End if 
    error_message = ""
Next 

'----------------------------------------------------------------------------------------------------CASE:NOTE - MAXIS 
Call navigate_to_MAXIS(maxis_mode) 'navigating to MAXIS Production area 

For item = 0 to Ubound(adjustment_array, 2)
    If adjustment_array(reduce_rate_const, item) = True then
        Call navigate_to_MAXIS_screen_review_PRIV(function_to_go_to, command_to_go_to, is_this_priv)    'Checking for PRIV case note status 
        If is_this_priv = False then
            'case note 
            start_a_blank_CASE_NOTE
            EmReadscreen error_message, 80, 24, 2
            If trim(error_message) <> ""  then 
                adjustment_array(MAXIS_note_conf_const, item) = False 
                adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Unable to enter MAXIS CASE:NOTE - " & trim(error_message)
            Else     
                Call write_variable_in_CASE_NOTE("DHS Supplemental Service Rate Adjustment")
                Call write_variable_in_CASE_NOTE("---")
                Call write_variable_in_CASE_NOTE("There is an active Housing Support supplemental service rate (SSR) service authorization in MMIS for this MAXIS case. DHS adjusted the MMIS service authorization(s) for Housing Support SSR through the existing end date of the service authorization.")
                Call write_variable_in_CASE_NOTE("")
                Call write_variable_in_CASE_NOTE("Revisions are based on a determination of the recipient's concurrent eligibility for Housing Stabilization Services. MMIS issued a revised service authorization with the correct SSR per diem to the Housing Support provider associated with the MMIS service authorization.")
                Call write_variable_in_CASE_NOTE("")
                Call write_variable_in_CASE_NOTE("Eligibility workers do not need to take any action in MAXIS.")
                PF3 'to save 
                adjustment_array(MAXIS_note_conf_const, item) = True 
            End if 
        Else 
            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Unable to enter MAXIS CASE:NOTE - PRIV Case."
        End if 
    End if 
    error_message = ""
Next 

'Excel output of rate reduction statuses 
For item = 0 to Ubound(adjustment_array, 2)
    objExcel.Cells(adjustment_array(excel_row_const, item), rate_reduction_col).Value = adjustment_array(rr_status_const, item)
Next 

'formatting the cells
FOR i = 1 to 27
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

MAXIS_case_number = ""  'blanking out for statistical purposes. Cannot collect more than one case number. 
STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure_with_error_report("Success! The script run is complete. Please review the worksheet for reduction statuses and manual updates.")