'Required for statistical purposes===============================================================================
name_of_script = "MISC - HSS MAXIS FACILITY REPORT.vbs"
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
call changelog_update("05/21/2021", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'CONNECTS TO BlueZone
EMConnect ""
Call check_for_MAXIS(false)

'----------------------------------Set up code
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

'Excel columns
const HS_status_col     = 16
const vendor_num_col    = 17
const faci_name_col     = 18
const faci_in_col       = 19
const faci_out_col      = 20
const impact_vnd_col    = 21
const exempt_code_col   = 22
const HDL_one_col       = 23
const HDL_two_col       = 24
const HDL_three_col     = 25
const case_status_col   = 26

'User interface dialog - There's just one in this script.
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 481, 90, "HSS MAXIS Facility Report"
  ButtonGroup ButtonPressed
    PushButton 420, 45, 50, 15, "Browse...", select_a_file_button
    OkButton 365, 65, 50, 15
    CancelButton 420, 65, 50, 15
  EditBox 15, 45, 400, 15, file_selection_path
  Text 15, 20, 455, 20, "This script should be used when adding MAXIS Facility information to an exisiting spreadsheet with an initial data set provided by DHS for the purposes of possible Supplemental Service Rate reductions due to overlapping Housing Stabilization Services (HSS)."
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
ObjExcel.Cells(1, HS_status_col).Value   = date & " MAXIS HS Status"   'col 16
ObjExcel.Cells(1, vendor_num_col).Value  = "Vendor #"                  'col 17
ObjExcel.Cells(1, faci_name_col).Value   = "Facility Name"             'col 18
ObjExcel.Cells(1, faci_in_col).Value     = "Faci In Date"              'col 19
ObjExcel.Cells(1, faci_out_col).Value    = "Faci Out Date"             'col 20
ObjExcel.Cells(1, impact_vnd_col).Value  = "Impacted Vendor?"          'col 21
ObjExcel.Cells(1, exempt_code_col).Value = "VND2 Exemption Code"       'col 22
ObjExcel.Cells(1, HDL_one_col).Value     = "VND2 HDL 1 Code"           'col 23
ObjExcel.Cells(1, HDL_two_col).Value     = "VND2 HDL 2 Code"           'col 24
ObjExcel.Cells(1, HDL_three_col).Value   = "VND2 HDL 3 Code"           'col 25
ObjExcel.Cells(1, case_status_col).Value = "Case Status"               'col 26

FOR i = 16 to 26		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

'----------------------------------------------------------------------------------------------------MAXIS DATA GATHER
Call check_for_MAXIS(False)             'Ensuring we're actually in MAXIS
Call MAXIS_footer_month_confirmation    'Ensuring we're in the right footer month/year: current footer month/year for this process.

Dim faci_array()                        'Delcaring array
ReDim faci_array(faci_out_const, 0)     'Resizing the array to size of last const
Dim item

const vendor_number_const   = 0         'creating array constants
const faci_name_const       = 1
const faci_in_const         = 2
const faci_out_const        = 3

excel_row = 2
Do
    client_PMI = trim(objExcel.cells(excel_row, 1).Value)
    If client_PMI = "" then exit do
    'removing preceeding 0's from the client PMI. This is needed to measure the PMI's on CASE/PERS.
    Do
		if left(client_PMI, 1) = "0" then client_PMI = right(client_PMI, len(client_PMI) -1)   'trimming off left-most 0 from client_PMI
	Loop until left(client_PMI, 1) <> "0"                                                      'Looping until 0's are all removed
    client_PMI = trim(client_PMI)

	MAXIS_case_number = trim(objExcel.cells(excel_row, 2).Value)
    case_status = ""            'defaulting case_status to "" to increment later in certain circumsatnces

    faci_count = 0                          'setting increment for array

    '----------------------------------------------------------------------------------------------------CASE/PERS & PERS Search
    Call navigate_to_MAXIS_screen_review_PRIV("CASE", "PERS", is_this_priv)
    If is_this_priv = True then
        case_status = "Privileged Case. Unable to access."
    Else
        member_found = False
        Call navigate_to_MAXIS_screen("CASE", "PERS")
        row = 10    'staring row for 1st member
        Do
            EMReadScreen person_PMI, 8, row, 34
            person_PMI = trim(person_PMI)
            If person_PMI = "" then exit do
            If trim(person_PMI) = client_PMI then
                EmReadscreen HS_status, 1, row, 66
                If trim(HS_status) <> "" then
                    EmReadscreen member_number, 2, row, 3
                    member_found = True
                    exit do
                End if
            Else
                row = row + 3			'information is 3 rows apart. Will read for the next member.
                If row = 19 then
                    PF8
                    row = 10					'changes MAXIS row if more than one page exists
                END if
            END if
        LOOP
        If trim(member_number) = "" then case_status = "Unable to locate case for member."
    End if

    If trim(case_status) = "" then
    '----------------------------------------------------------------------------------------------------FACI panel determination
	   call navigate_to_MAXIS_screen("STAT", "FACI")
       EmWriteScreen member_number, 20, 76
       Call write_value_and_transmit("01", 20, 79)  'making sure we're on the 1st instance for member
       'Based on how many FACI panels exist will determine if/how the information is read.
	    EMReadScreen FACI_total_check, 1, 2, 78
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
                                current_faci_found = True   'Condition is met so date evaluation via FACI_array is not needed.
                                exit do    'open ended faci found
                            End if
                        End if
                    Elseif faci_out <> "" then
                        If faci_in <> "" then
                            faci_out_dates_string = faci_out_dates_string & faci_out & "|"

                            Redim Preserve faci_array(faci_out_const, faci_count)
                            faci_array(vendor_number_const, faci_count) = vendor_number
                            faci_array(faci_name_const,     faci_count) = faci_name
                            faci_array(faci_in_const,       faci_count) = faci_in
                            faci_array(faci_out_const,      faci_count) = faci_out
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
                faci_out_dates = split(faci_out_dates_string, "|")
                call sort_dates(faci_out_dates)
                first_date = faci_out_dates(0)                              'setting the first and last check dates
                last_date = faci_out_dates(UBOUND(faci_out_dates))

                'finding the most recent date if none of the dates are open-ended
                For item = 0 to Ubound(faci_array, 2)
                    If faci_array(faci_out_const, item) = last_date then
                        vendor_number   = faci_array(vendor_number_const, item)
                        faci_name       = faci_array(faci_name_const, item)
                        faci_in         = faci_array(faci_in_const, item)
                        faci_out        = faci_array(faci_out_const, item)
                    End if
                Next
            End if
            ReDim faci_array(faci_out_const, 0)     'Resizing the array back to original size
            Erase faci_array                        'then once resized it gets erased.
	    End if

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
            If HDL_one <> "" then
                If instr(HDL_applicable_codes, HDL_one) then health_depart_reason = True
            End if

            If HDL_two <> "" then
                If instr(HDL_applicable_codes, HDL_two) then health_depart_reason = True
            End if

            If HDL_three <> "" then
                If instr(HDL_applicable_codes, HDL_three) then health_depart_reason = True
            End if

            If exemption_code = "15" or exemption_code = "26" or exemption_code = "28" then
                exemption_reason = True
            Else
                exmption_reason = False
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
    excel_row = excel_row + 1 'setting up the script to check the next row.
    stats_counter = stats_counter + 1
LOOP UNTIL objExcel.Cells(excel_row, 2).Value = ""	'Loops until there are no more cases in the Excel list

'formatting the cells
FOR i = 1 to 26
	objExcel.Columns(i).AutoFit()				'sizing the columns
NEXT

MAXIS_case_number = ""  'blanking out for statistical purposes. Cannot collect more than one case number.

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure_with_error_report("Success! Your facility data has been created.")

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
'--All variables are CASE:NOTEing (if required)---------------------------------08/13/2021-----------------No CASE:NOTE, data only
'--CASE:NOTE Header doesn't look funky------------------------------------------08/13/2021
'--Leave CASE:NOTE in edit mode if applicable-----------------------------------08/13/2021----------------N/A: Bulk Process
'-----General Supports-------------------------------------------------------------------------------------------------------------
'--Check_for_MAXIS/Check_for_MMIS reviewed--------------------------------------08/13/2021
'--MAXIS_background_check reviewed (if applicable)------------------------------08/13/2021----------------N/A: Not updating in MAXIS
'--PRIV Case handling reviewed -------------------------------------------------08/13/2021
'--Out-of-County handling reviewed----------------------------------------------08/13/2021----------------N/A: DHS script
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
