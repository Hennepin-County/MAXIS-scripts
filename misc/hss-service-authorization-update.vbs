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

'----------------------------------------------------------------------------------------------------The Script
'CONNECTS TO BlueZone
EMConnect ""
Check_for_MMIS(false)   'checking for, and allowing user to navigate into MMIS.

'----------------------------------Set up code
'Excel columns
const recip_PMI_col         = 1     'Col A
const case_number_col       = 2     'Col B
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

Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)  'opens the selected excel file

'Setting up the Excel spreadsheet
ObjExcel.Cells(1, rate_reduction_col).Value = "Rate Reduction Status"   'col 27

'formatting the cells
objExcel.Cells(1, 27).Font.Bold = True		'bold font'
objExcel.Columns(27).ColumnWidth = 120		'sizing the last column

Dim adjustment_array()                        'Delcaring array
ReDim adjustment_array(rr_status_const, 0)     'Resizing the array to size of last const
Dim item

const recip_PMI_const               = 0         'creating array constants
const case_number_const             = 1
const HSS_start_const               = 2
const HSS_end_const                 = 3
const SA_number_const               = 4
const agreement_start_const         = 5
const agreement_end_const           = 6
const npi_number_const              = 7
const HS_status_const               = 8
const faci_in_const                 = 9
const faci_out_const                = 10
const impacted_vendor_const         = 11
const case_status_const             = 12
const prev_start_const              = 13
const prev_end_const                = 14
const new_start_const               = 15
const new_end_const                 = 16
const excel_row_const               = 17
const MAXIS_note_conf_const         = 18
const MMIS_note_conf_const          = 19
const reduce_rate_const             = 20
const adjustment_start_date_const   = 21
const passed_case_tests_const       = 22
const pmi_count_const               = 23
const rate_amt_const                = 24
const rate_reduction_notes_const    = 25
const rr_status_const               = 26

excel_row = 2
entry_record = 0 'incrementor for the array

Do
    recip_PMI = trim(objExcel.cells(excel_row, recip_PMI_col).Value)
    If recip_PMI = "" then exit do

    SA_number       = trim(objExcel.cells(excel_row, SA_number_col).Value)
    SA_number = right("00000000" & SA_number, 11) 'ensures the variable is 11 digits. Inhibiting erorr

    'Adding recipient information to the array
    ReDim Preserve adjustment_array(rr_status_const, entry_record)	'This resizes the array based on the number of rows in the Excel File'
    adjustment_array(recip_PMI_const            , entry_record) = recip_PMI
    adjustment_array(case_number_const          , entry_record) = trim(objExcel.cells(excel_row, case_number_col).Value)
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
    adjustment_array(passed_case_tests_const    , entry_record) = False 'defaulting to false
    adjustment_array(MAXIS_note_conf_const      , entry_record) = False 'defaulting to false
    adjustment_array(MMIS_note_conf_const       , entry_record) = False 'defaulting to false
    adjustment_array(reduce_rate_const          , entry_record) = False 'defaulting to false
    adjustment_array(rate_reduction_notes_const , entry_record) = trim(objExcel.cells(excel_row, rate_reduction_col).Value)
    entry_record = entry_record + 1			'This increments to the next entry in the array'
    stats_counter = stats_counter + 1
    excel_row = excel_row + 1
    recip_PMI = ""  'Blanking out variables for next loop
    SA_number = ""  'Blanking out variables for next loop
Loop

'----------------------------------------------------------------------------------------------------Notes on updates
'Rates prior to 2024 should use the following rates:
'    •  Full Rate: $15.87
'    •  Reduced Rate: $7.94
'Anything 01/01/2024 - ongoing should use the following rates: 
'    •  Full Rate: $16.27
'    •  Reduced Rate: $8.14

'----------------------------------------------------------------------------------------------------determine which rows of information are going to have a rate reduction or not.
For item = 0 to Ubound(adjustment_array, 2)
    'Determining which date to use to end/start the agreements. Initial conversion date is 07/01/21. We cannot use a date earlier than this. If a date is earlier than this, the date is 07/01/21.
    'This supports both the initial conversion and ongoing cases.
    If DateDiff("d", #07/01/21#, adjustment_array(HSS_start_const, item)) <= 0 then
        'if HSS start date is a negative/a date before 07/01/21 (past date), then use 07/01/21.
        new_agreement_start_date = #07/01/21#
        Call ONLY_create_MAXIS_friendly_date(new_agreement_start_date)
        adjustment_array(adjustment_start_date_const, item) = new_agreement_start_date
    Else
        Call ONLY_create_MAXIS_friendly_date(adjustment_array(HSS_start_const, item))
        adjustment_array(adjustment_start_date_const, item) = adjustment_array(HSS_start_const, item)
    End if

    'if this date is a negative then the agreement start date is after the HSS start date. Use the agreement start date instead of HSS start date.
    If DateDiff("d", adjustment_array(agreement_start_const, item), adjustment_array(adjustment_start_date_const, item)) <= 0 then
        Call ONLY_create_MAXIS_friendly_date(adjustment_array(agreement_start_const, item))
        adjustment_array(adjustment_start_date_const, item) = adjustment_array(agreement_start_const, item)
    End if

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
    'These are the initial case tests that will fail:
    'Rows with Case Status of “Unable to find MONY/VND2 panel”
    'Rows with Case Status of “Privileged Case. Unable to access.”
    'Row’s that have more than one MAXIS case identified, and HS is not active for the recipient on that case.
    'Row’s that are not identified as an Impacted Vendor (“Yes”)
    'Open-ended facility spans or recipients that have faci panels that close after the HSS start date.
    'Rows that may already be done.
    'Rate costs that are not 16.27
    If (adjustment_array(case_status_const, item) = "" and _
        adjustment_array(rate_reduction_notes_const, item) = "" and _
        adjustment_array(HS_status_const, item) <> "" and _
        adjustment_array(impacted_vendor_const, item) = "Yes" and _
        adjustment_array(rate_amt_const, item) = "16.27" and _
        active_facility = True) then
        adjustment_array(passed_case_tests_const, item) = True
    Else
    'Failure Reasons
        If adjustment_array(HS_status_const, item) = "" then rate_reduction_status = rate_reduction_status & "No HS Status in MAXIS Case. "
        If adjustment_array(impacted_vendor_const, item) = "Yes" and adjustment_array(rate_amt_const, item) <> "16.27" then rate_reduction_status = rate_reduction_status & "Rate is not 16.27, review manually. "
        If adjustment_array(impacted_vendor_const, item) <> "Yes" then rate_reduction_status = rate_reduction_status & "Not an impacted vendor. "
        If active_facility = False then rate_reduction_status = rate_reduction_status & "Not an active facility. "
        If adjustment_array(case_status_const, item) <> "" then rate_reduction_status = rate_reduction_status & adjustment_array(case_status_const, item)
        If adjustment_array(rate_reduction_notes_const, item) <> "" then rate_reduction_status = adjustment_array(rate_reduction_notes_const, item) 'not incrementing this failure reason. Just inputting exiting notes.
    End if
    If rate_reduction_status <> "Failed Case Test(s): " then adjustment_array(rr_status_const, item) = rate_reduction_status
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
Next

For item = 0 to Ubound(adjustment_array, 2)
    If adjustment_array(pmi_count_const, item) > 1 then
        If adjustment_array(passed_case_tests_const, item) = True then
            adjustment_array(passed_case_tests_const, item) = False
            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & "Duplicate agreements found. Review manually."
        End if
    End if

    If adjustment_array(passed_case_tests_const, item) = True then adjustment_array(reduce_rate_const, item) = True    'cases that have passed the cases tests will be initially set to reduce.
    rate_reduction_status = ""  'blanking out variable.
Next

'----------------------------------------------------------------------------------------------------MMIS STEPS
Call check_for_MMIS(False)

For item = 0 to Ubound(adjustment_array, 2)
    If adjustment_array(reduce_rate_const, item) = True then
        'start the rate reductions in MMIS
        Call navigate_to_MMIS_region("GRH UPDATE")	'function to navigate into MMIS, select the GRH update realm, and enter the prior authorization area
        Call MMIS_panel_confirmation("AKEY", 51)				'ensuring we are on the right MMIS screen
        EmWriteScreen "C", 3, 22
        Call write_value_and_transmit(adjustment_array(SA_number_const, item), 9, 36) 'Entering Service Authorization Number and transmit to ASA1
        EmReadscreen current_panel, 4, 1, 51
        If current_panel = "AKEY" then
            error_message = ""
            EmReadscreen error_message, 50, 24, 2
            adjustment_array(reduce_rate_const, item) = False
            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Authorization Number is not valid."
        Else
            EMReadScreen AGMT_STAT, 1, 3, 17
            If AGMT_STAT <> "A" then
                adjustment_array(reduce_rate_const, item) = False
                adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Authorization Status is coded as: " & AGMT_STAT & "."
            Else
                transmit 'to ASA1
                Call write_value_and_transmit("ASA3", 1, 8)             'Direct navigate to ASA3
                Call MMIS_panel_confirmation("ASA3", 51)				'ensuring we are on the right MMIS screen

                EmReadscreen line_1_rate, 5, 9, 24
                If trim(line_1_rate) = "8.14" then
                    adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Line 1 already reflects reduction of 8.14."
                    PF6 'cancel
                    transmit 'to re-enter ASA1
                    EmWriteScreen "A", 3, 17   'Restoring the agreement on ASA1 in AGMT/TYPE STAT field
                    PF3
                    adjustment_array(reduce_rate_const, item) = False
                Else
                    'Checking Line 2 to ensure it's blank
                    EmReadscreen line_2_check, 6, 14, 60
                    If trim(line_2_check) <> "" then
                        EmReadscreen line_2_rate, 4, 15, 25
                        PF6 'cancel
                        transmit 'to re-enter ASA1
                        EmWriteScreen "A", 3, 17   'Restoring the agreement on ASA1 in AGMT/TYPE STAT field
                        PF3
                        adjustment_array(reduce_rate_const, item) = False
                        'creating status message if reduce is already in exisitance.
                        If line_2_rate = "8.14" then
                            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Line 2 already reflects reduction of 8.14."
                        Else
                            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Agreement already exists in Line 2. Review Manually."
                        End if
                    Else
                        'Reading and converting start and end dates
                        'agreement start date
                        EMReadScreen start_month, 2, 8, 60
                        EMReadScreen start_day, 2, 8, 62
                        EMReadScreen start_year, 2, 8, 64
                        Line_1_start_date = start_month & "/" & start_day & "/" & start_year
                        Call ONLY_create_MAXIS_friendly_date(Line_1_start_date)

                        'For cases that Line 1 agreements are the same day or before the HSS start date.
                        If DateDiff("d", Line_1_start_date, adjustment_array(adjustment_start_date_const, item)) < 0 then
                            'if this date is a negative or a date before 07/01/21 (past date), then use 07/01/21.
                            PF6 'cancel
                            transmit 'to re-enter ASA1
                            EmWriteScreen "A", 3, 17   'Restoring the agreement on ASA1 in AGMT/TYPE STAT field
                            PF3
                            adjustment_array(reduce_rate_const, item) = False
                            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Agreement start date (" & Line_1_start_date & ") is <= HSS start date (" & adjustment_array(adjustment_start_date_const, item) & ")."
                        Else
                            'agreement end date - original end date from line 1
                            EMReadScreen end_month, 2, 8, 67
                            EMReadScreen end_day, 2, 8, 69
                            EMReadScreen end_year, 2, 8, 71
                            original_end_date = end_month & "/" & end_day & "/" & end_year
                            Call ONLY_create_MAXIS_friendly_date(original_end_date)
                            write_original_end_date = replace(original_end_date, "/", "")  'for line 2

                            'Failing cases that the end date is less than the new agreement start date
                            If DateDiff("d", adjustment_array(adjustment_start_date_const, item), original_end_date) <= 0 then
                                'if this date is a positive then its a date before the HSS start date and needs to fail.
                                PF6 'cancel
                                transmit 'to re-enter ASA1
                                EmWriteScreen "A", 3, 17   'Restoring the agreement on ASA1 in AGMT/TYPE STAT field
                                PF3
                                adjustment_array(reduce_rate_const, item) = False
                                adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & "Agreement end date (" & original_end_date & ") is < HSS start date (" & adjustment_array(adjustment_start_date_const, item) & ")."
                            Else
                                'Creating a date that is the day before the HSS start date/conversion date - for LINE 1
                                new_line_1_end_date = dateadd("d", -1, adjustment_array(adjustment_start_date_const, item))
                                'using the HSS start date as this is after 07/01/21 (future date from initial coversion date of 07/01/21)
                                Call ONLY_create_MAXIS_friendly_date(new_line_1_end_date)

                                'removing date formatting for ASA3 input
                                write_new_line_1_end_date = replace(new_line_1_end_date, "/", "")

                                line_1_total_units = datediff("d", Line_1_start_date, new_line_1_end_date) + 1

                                'Unable to close agreements that have been overbilled by the facility.
                                over_billed = True      'Defaulting to True
                                EmReadscreen billed_units, 6, 11, 60
                                billed_units = trim(billed_units)
                                If trim(billed_units) = "" then
                                    over_billed = False   'no billing exists - blank
                                ElseIf cint(billed_units) = cint(line_1_total_units) then
                                    over_billed = False 'facility only billed up to the amount of the date we are closing this agreement date.
                                Elseif cint(billed_units) < cint(line_1_total_units) then
                                    over_billed = False  'facility billed less than the amount of the date we are closing this agreement date.
                                End if

                                If over_billed = True then
                                    PF6 'cancel
                                    transmit 'to re-enter ASA1
                                    EmWriteScreen "A", 3, 17   'Restoring the agreement on ASA1 in AGMT/TYPE STAT field
                                    PF3
                                    adjustment_array(reduce_rate_const, item) = False
                                    adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & " Unable to reduce Line 1 agreement due to overbilling. Billed units: & " & billed_units & " vs. " & line_1_total_units & "."
                                Else
                                    'Deleting the orginal agreement if the start dates are the same date
                                    If DateDiff("d", Line_1_start_date, adjustment_array(adjustment_start_date_const, item)) = 0 then
                                        EmWriteScreen "D", 12, 19 'Deny orginal agreement
                                    Else
                                        '----------------------------------------------------------------------------------------------------Updating LINE 1 agreement
                                        EmWriteScreen write_new_line_1_end_date, 8, 67
                                        Call clear_line_of_text(9, 60)
                                        EmWriteScreen line_1_total_units, 9, 60
                                    End if
                                    '----------------------------------------------------------------------------------------------------Entering LINE 2 Information
                                    EmWriteScreen "H0043", 13, 36
                                    EmWriteScreen "U5", 13, 44

                                    write_new_agrement_start_date = replace(adjustment_array(adjustment_start_date_const, item), "/", "")

                                    EmWriteScreen write_new_agrement_start_date, 14, 60
                                    EmWriteScreen write_original_end_date, 14, 67

                                    EmReadscreen old_rate, 5, 9, 24
                                    new_rate = old_rate / 2 'divide total by two, and round to integer
                                    new_rate = Round(new_rate, 2) 'round to two decimal places
                                    EmWriteScreen new_rate, 15, 20

                                    line_2_total_units = datediff("d", adjustment_array(adjustment_start_date_const, item), original_end_date) + 1
                                    EmWriteScreen line_2_total_units, 15, 60

                                    EMReadscreen agreement_NPI_number, 10, 10, 20   'Reading line 1 NPI Number
                                    EmReadscreen facility_name, 35, 10, 31
                                    EmWriteScreen agreement_NPI_number, 16, 20      'Enetering NPI in Line 2 agreement

                                    EmWriteScreen new_rate, 17, 20
                                    EmWriteScreen "MM", 17, 35

                                    EmWriteScreen "A", 18, 19   'Approving the agreement on ASA3 in STAT CD/DATE field
                                    EmWriteScreen "A", 3, 20   'Approving the agreement on ASA3 in AGMT/TYPE STAT field
                                    transmit

                                    'PF3 ' to save
                                    EMReadScreen PPOP_check, 4, 1, 52
                                    If PPOP_check = "PPOP" then
                                        faci_found = False
                                        'Setting default rows to start
                                        faci_name_row = 5
                                        active_status_row = 8

                                        Do
                                            EmReadscreen faci_name, 35, faci_name_row, 5
                                            If trim(facility_name) = trim(faci_name) then
                                                EmReadscreen provider_type, 18, faci_name_row, 52
                                                EmReadscreen facility_status, 10, active_status_row, 49
                                                If trim(provider_type) = "18 H/COMM PRV" and trim(facility_status) = "ACTIVE" then
                                                    faci_found = True
                                                    Call write_value_and_transmit("X", faci_name_row, 2)    'selecting the found file. Will only select the 1st instance it can find.
                                                    exit do
                                                Else
                                                    faci_name_row = faci_name_row + 4               'incrementing to next facility information section
                                                    active_status_row = active_status_row + 4
                                                    If faci_name_row = 21 then
                                                        PF8                     'Accounting for more than one page of facilities
                                                        faci_name_row = 5       'resetting the rows to the 1st facility set
                                                        active_status_row = 8
                                                        EmReadscreen last_page, 60, 24, 20
                                                    End if
                                                End if

                                            Else
                                                faci_name_row = faci_name_row + 4               'incrementing to next facility information section
                                                active_status_row = active_status_row + 4
                                                If faci_name_row = 21 then
                                                    PF8                     'Accounting for more than one page of facilities
                                                    faci_name_row = 5       'resetting the rows to the 1st facility set
                                                    active_status_row = 8
                                                    EmReadscreen last_page, 60, 24, 20
                                                End if
                                            End if
                                        Loop until trim(last_page) = "CANNOT SCROLL FORWARD - NO MORE DATA TO DISPLAY."

                                        If faci_found = False then
                                            Dialog1 = ""
                                                BeginDialog Dialog1, 0, 0, 181, 130, "PPOP screen - Choose Facility"
                                                ButtonGroup ButtonPressed
                                                  OkButton 65, 105, 50, 15
                                                  CancelButton 120, 105, 50, 15
                                                Text 5, 5, 170, 35, "Please select the correct facility name/address from the list in PPOP by putting a 'X' next to the name. DO NOT TRANSMIT. Press OK when ready. Press CANCEL to stop the script."
                                                Text 5, 45, 175, 20, "* Provider types for GRH must be '18/H COMM PRV' and the status must be '1 ACTIVE.'"
                                                Text 5, 75, 175, 20, "Line 1 Provider Name: " & trim(facility_name)
                                            EndDialog
                                            Do
                                                dialog Dialog1
                                                cancel_confirmation
                                            Loop until ButtonPressed = -1
                                		    EMReadScreen PPOP_check, 4, 1, 52
                                            If PPOP_check = "PPOP" then transmit     'to exit PPOP
                                            If PPOP_check = "SA3 " then transmit    'to navigate to ACF1 - this is the partial screen check for ASA3
                                            transmit ' to next available screen (does not need to be updated)
                                            Call write_value_and_transmit("ACF3", 1, 51)
                                        End if
                                    End if
                                    'saving the agreements
                                    PF3
                                    EmReadscreen current_panel, 4, 1, 51

                                    If current_panel = "AKEY" then
                                        error_message = ""
                                        EmReadscreen error_message, 50, 24, 2
                                        If trim(error_message) = "ACTION COMPLETED" then
                                            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & "Agreement successfully reduced to " & new_rate & "."
                                        Else
                                            adjustment_array(reduce_rate_const, item) = False
                                            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & "Not reduced. MMIS Error: " & trim(error_message)
                                        End if
                                    Else
                                        error_message = ""
                                        EmReadscreen error_message, 80, 21, 2       'reading error message on any other screen.
                                        adjustment_array(reduce_rate_const, item) = False
                                        adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & "Not reduced. MMIS Error: " & trim(error_message)
                                        PF6 'cancel
                                        transmit 'to re-enter ASA1
                                        EmWriteScreen "A", 3, 17   'Restoring the agreement on ASA1 in AGMT/TYPE STAT field
                                        PF3
                                    End if
                                End if
                            End if
                        End if
                    End if
                End if
            End if
        End if
    End if
Next

write_this_thing = "DHS SUPPLEMENTAL SERVICE RATE ADJUSTMENT" & "~" & "THERE IS AN ACTIVE HOUSING SUPPORT SUPPLEMENTAL SERVICE RATE (SSR)" & "~" & "SERVICE AUTHORIZATION IN MMIS FOR THIS MAXIS CASE. DHS ADJUSTED THE" & "~" &_
				   "MMIS SERVICE AUTHORIZATION(S) FOR HOUSING SUPPORT SSR THROUGH THE" & "~" & "EXISITING END DATE OF THE SERVICE AUTHORIZATION." & "~" & "REVISIONS ARE BASED ON A DETERMINATION OF THE RECIPIENT'S CONCURRENT" & "~" &_
				   "ELIGBILITY HOUSING STABILIZATION SERVICES. MMIS ISSUED A REVISED" & "~" & "SERVICE AUTORIZATION WITH THE CORRECT SSR PER DIEM TO THE HOUSING" & "~" & "SUPPORT PROVIDER ASSOCIATED WITH THE MMIS SERVICE AUTHORIZATION." & "~" &_
				   "ELIGIBILITY WORKERS DO NOT NEED TO TAKE ANY ACTION IN MAXIS." & "~" & "**********************************************************************"
AN_ARRAY_OF_THE_THING_TO_WRITE = split(write_this_thing, "~")
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

        'Writing in the ADHS - DHS Comments Notes
		for each comment_line in AN_ARRAY_OF_THE_THING_TO_WRITE
			EmWriteScreen comment_line, row, 3
			row = row + 1
			If row = 14 Then Exit For
		Next

        PF3
        error_message = ""
        EmReadscreen error_message, 40, 24, 2
        If trim(error_message) =  "ACTION COMPLETED" then
            adjustment_array(MMIS_note_conf_const, item) = True
        Else
            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & adjustment_array(rr_status_const, item) & " Unable to enter note on ADHS - " & trim(error_message)
        End if
    End if
Next

'----------------------------------------------------------------------------------------------------CASE:NOTE - MAXIS
Call navigate_to_MAXIS(maxis_mode)  'Function to navigate back to MAXIS
Call check_for_MAXIS(False)         'Checking to see if we're in MAXIS and/or passworded out.

For item = 0 to Ubound(adjustment_array, 2)
    If adjustment_array(reduce_rate_const, item) = True then
        MAXIS_case_number = adjustment_array(case_number_const, item)
        Call navigate_to_MAXIS_screen_review_PRIV(function_to_go_to, command_to_go_to, is_this_priv)    'Checking for PRIV case note status
        If is_this_priv = False then
            'case note
            Call navigate_to_MAXIS_screen("CASE", "NOTE")
            PF9
            error_message = ""
            EmReadscreen case_note_edit_errors, 70, 3, 3
            EmReadscreen error_message, 50, 24, 2
            If trim(error_message) <> ""  then
                adjustment_array(MAXIS_note_conf_const, item) = False
                adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & " Unable to enter MAXIS CASE:NOTE - " & trim(error_message)
            Elseif trim(case_note_edit_errors) <> "Please enter your note on the lines below:" then
                adjustment_array(MAXIS_note_conf_const, item) = False
                adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & " Unable to edit MAXIS CASE:NOTE - " & trim(error_message)
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
            adjustment_array(rr_status_const, item) = adjustment_array(rr_status_const, item) & " Unable to enter MAXIS CASE:NOTE - PRIV Case."
        End if
    End if
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

'----------------------------------------------------------------------------------------------------Closing Project Documentation
'------Task/Step--------------------------------------------------------------Date completed---------------Notes-----------------------
'
'------Dialogs--------------------------------------------------------------------------------------------------------------------
'--Dialog1 = "" on all dialogs -------------------------------------------------08/13/2021
'--Tab orders reviewed & confirmed----------------------------------------------08/13/2021
'--Mandatory fields all present & Reviewed--------------------------------------08/13/2021
'--All variables in dialog match mandatory fields-------------------------------08/13/2021
'
'-----CASE:NOTE-------------------------------------------------------------------------------------------------------------------
'--All variables are CASE:NOTEing (if required)---------------------------------08/13/2021-----------------No variables, just singular message
'--CASE:NOTE Header doesn't look funky------------------------------------------08/13/2021
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------08/13/2021----------------N/A: Bulk Process
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------08/13/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------08/13/2021----------------N/A: Not updating in MAXIS
'--PRIV Case handling reviewed -------------------------------------------------08/13/2021
'--Out-of-County handling reviewed----------------------------------------------08/13/2021----------------Can make updates in MMIS, MAXIS CASE:NOTES has OOC handling
'--script_end_procedures (w/ or w/o error messaging)----------------------------08/13/2021
'--BULK - review output of statistics and run time/count (if applicable)--------08/13/2021
'
'-----Statistics--------------------------------------------------------------------------------------------------------------------
'--Manual time study reviewed --------------------------------------------------08/13/2021
'--Incrementors reviewed (if necessary)-----------------------------------------08/13/2021
'--Denomination reviewed -------------------------------------------------------08/13/2021
'--Script name reviewed---------------------------------------------------------08/13/2021
'--BULK - remove 1 incrementor at end of script reviewed------------------------08/13/2021

'-----Finishing up------------------------------------------------------------------------------------------------------------------
'--Confirm all GitHub taks are complete-----------------------------------------08/13/2021
'--comment Code-----------------------------------------------------------------08/13/2021
'--Update Changelog for release/update------------------------------------------08/13/2021
'--Remove testing message boxes-------------------------------------------------08/13/2021
'--Remove testing code/unnecessary code-----------------------------------------08/13/2021
'--Review/update SharePoint instructions----------------------------------------08/13/2021-------------------N/A: Logic Map provided to DHS
'--Review Best Practices using BZS page ----------------------------------------08/13/2021-------------------N/A: DHS script
'--Review script information on SharePoint BZ Script List-----------------------08/13/2021-------------------N/A: DHS script
'--Other SharePoint sites review (HSR Manual, etc.)-----------------------------08/13/2021-------------------N/A: DHS script
'--COMPLETE LIST OF SCRIPTS reviewed--------------------------------------------08/13/2021-------------------N/A: DHS script
'--Complete misc. documentation (if applicable)---------------------------------08/13/2021
'--Update project team/issue contact (if applicable)----------------------------08/13/2021
