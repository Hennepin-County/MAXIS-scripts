'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - DAIL 12 MONTH CONTACT.vbs"
start_time = timer
STATS_counter = 1                       'sets the stats counter at one
STATS_manualtime = 30
STATS_denomination = "I"       			'I is for each item
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
call changelog_update("05/07/2023", "Initial version.", "Ilse Ferris, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================
'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""

'The dialog is defined in the loop as it can change as buttons are pressed 
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "Restart DAIL 12 Month Contact at Evaluation."
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used when a TIKL 12 month contact list needs to be restared at the point of the evaluation portion."
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

Dialog1 = ""
'Select Excel row dialog
BeginDialog Dialog1, 0, 0, 126, 50, "Select the excel row to restart"
  EditBox 75, 5, 40, 15, excel_row_to_restart
  ButtonGroup ButtonPressed
    OkButton 10, 25, 50, 15
    CancelButton 65, 25, 50, 15
  Text 10, 10, 60, 10, "Excel row to start:"
EndDialog

Do 
	dialog Dialog1 
	cancel_without_confirmation
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

excel_row = excel_row_to_restart

Call check_for_MAXIS(False)

DIM DAIL_array()
ReDim DAIL_array(case_status_const, 0)
Dail_count = 0              'Incrementor for the array

'constants for array
const worker_const	                = 0
const maxis_case_number_const       = 1
const dail_type_const               = 2
const dail_month_const		        = 3
const dail_msg_const		        = 4
const snap_status_const             = 5
const other_programs_present_const  = 6
const send_memo_const               = 7
const excel_row_const               = 8
const case_status_const             = 9

'Sets variable for all of the Excel stuff
deleted_dails = 0	'establishing the value of the count for deleted deleted_dails
memo_count = 0      'counting number of memo's sent with this process.
all_case_numbers_array = "*"    'setting up string to find duplicate case numbers
entry_record = 0

Do 
    MAXIS_case_number = trim(objExcel.cells(excel_row, 2).Value)

    'If the case number is found in the string of case numbers, it's not added again.
    If instr(all_case_numbers_array, "*" & Client_PMI & "*") then
        add_to_array = False
    Else 
        ReDim DAIL_array(case_status_const, 2)	'This resizes the array based on the number of rows in the Excel File'
        'The client information is added to the array'

        DAIL_array(worker_const,                entry_record) = trim(objExcel.cells(excel_row, 1).Value)
        DAIL_array(maxis_case_number_const,     entry_record) = MAXIS_case_number
        DAIL_array(dail_type_const,             entry_record) = trim(objExcel.cells(excel_row, 3).Value)
        DAIL_array(dail_month_const,		    entry_record) = trim(objExcel.cells(excel_row, 4).Value)
        DAIL_array(dail_msg_const,		        entry_record) = trim(objExcel.cells(excel_row, 5).Value)
        entry_record = entry_record + 1			'This increments to the next entry in the array'
        stats_counter = stats_counter + 1
        all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*") 'Adding MAXIS case number to case number string
    End if
    excel_row = excel_row + 1
Loop

For item = 0 to Ubound(DAIL_array, 2)
    MAXIS_case_number = DAIL_array(maxis_case_number_const, item)
    Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
    EmReadscreen worker_county, 4, 21, 14
    If is_this_priv = True then
        DAIL_array(case_status_const, item) = "Privilged Case."
    Elseif worker_county <> worker_county_code then
        DAIL_array(case_status_const, item) = "Out-of-County Case."
    Else
		Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
        'Determinging if the MEMO's need to be sent or not. Only SNAP active only cases require a memo.
        If snap_case = True then
            If snap_status <> "ACTIVE" then
                DAIL_array(case_status_const, item) = "SNAP is " & snap_status & "."
            Else
                If case_pending = True then DAIL_array(case_status_const, item) = "Pending program."
                'If other programs are active/pending then no notice is necessary
                If  ga_case = True OR _
                    msa_case = True OR _
                    mfip_case = True OR _
                    dwp_case = True OR _
                    grh_case = True OR _
                    ma_case = True OR _
                    msp_case = True then
                        DAIL_array(other_programs_present_const, item) = True
                        DAIL_array(case_status_const, item) = "Other programs present."
                Else
                    DAIL_array(other_programs_present_const, item) = False
                    DAIL_array(send_memo_const, item) = True
                End if
            End if
        Else
            DAIL_array(case_status_const, item) = "SNAP not active."
        End if
        DAIL_array(snap_status_const, item) = snap_status
    End if

    If DAIL_array(send_memo_const, item) = True then
        MAXIS_background_check
        stats_counter = stats_counter + 1
        deleted_dails = deleted_dails + 1
        Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, False)	'navigates to spec/memo and opens into edit mode

        Call write_variable_in_SPEC_MEMO("************************************************************")
        Call write_variable_in_SPEC_MEMO("This notice is to remind you to report changes to your county worker by the 10th of the month following the month of the change. Changes that must be reported are address, people in your household, income, shelter costs and other changes such as legal obligation to pay child support. If you don't know whether to report a change, contact your county worker.")
		CALL write_variable_in_SPEC_MEMO("")
		CALL write_variable_in_SPEC_MEMO("*** Submitting Documents:")
		CALL write_variable_in_SPEC_MEMO("- Online at infokeep.hennepin.us or MNBenefits.mn.gov")
		CALL write_variable_in_SPEC_MEMO("  Use InfoKeep to upload documents directly to your case.")
		CALL write_variable_in_SPEC_MEMO("- Mail, Fax, or Drop Boxes at Service Centers.")
		CALL write_variable_in_SPEC_MEMO("  More Info: https://www.hennepin.us/economic-supports")
        Call write_variable_in_SPEC_MEMO("************************************************************")
        PF4
        EmReadscreen memo_confirmation, 26, 24, 2
        If memo_confirmation <> "NEW MEMO CREATE SUCCESSFUL" then
            DAIL_array(case_status_const, item) = "Unable to send MEMO. Process Manually."
        Else
            'THE CASE NOTE
            stats_counter = stats_counter + 1
            Call start_a_blank_CASE_NOTE
            Call write_variable_in_CASE_NOTE("Sent SNAP 12 Month Contact letter via MEMO on " & date)

            Call write_variable_in_CASE_NOTE("---")
            Call write_variable_in_CASE_NOTE(worker_signature)
            PF3 'save CASE:NOTE
            DAIL_array(case_status_const, item) = "Success! MEMO sent."
            memo_count = memo_count + 1
        End if
    End if

    objExcel.Cells(DAIL_array(excel_row_const, item), 6).Value = DAIL_array(snap_status_const, item)
    objExcel.Cells(DAIL_array(excel_row_const, item), 7).Value = DAIL_array(other_programs_present_const, item)
    objExcel.Cells(DAIL_array(excel_row_const, item), 8).Value = DAIL_array(send_memo_const, item)
    objExcel.Cells(DAIL_array(excel_row_const, item), 9).Value = DAIL_array(case_status_const, item)
Next

STATS_counter = STATS_counter - 1
'Enters info about runtime for the benefit of folks using the script
objExcel.Cells(2, 11).Value = "Number of DAILs processed:"
objExcel.Cells(3, 11).Value = "Number of Memo's sent to residents:"
objExcel.Cells(4, 11).Value = "Average time to find/select/copy/paste one line (in seconds):"
objExcel.Cells(5, 11).Value = "Estimated manual processing time (lines x average):"
objExcel.Cells(6, 11).Value = "Script run time (in seconds):"
objExcel.Cells(7, 11).Value = "Estimated time savings by using script (in minutes):"
objExcel.Cells(8, 11).Value = "Number of TIKL messages reviewed"
objExcel.Columns(11).Font.Bold = true
objExcel.Cells(2, 12).Value = deleted_dails
objExcel.Cells(3, 12).Value = memo_count
objExcel.Cells(4, 12).Value = STATS_manualtime
objExcel.Cells(5, 12).Value = STATS_counter * STATS_manualtime
objExcel.Cells(6, 12).Value = timer - start_time
objExcel.Cells(7, 12).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60
objExcel.Cells(8, 12).Value = STATS_counter

'Formatting the column width
FOR i = 1 to 12
	objExcel.Columns(i).AutoFit()
NEXT

objExcel.ActiveWorkbook.Save 'Save

script_end_procedure("Success! Please review the list created for accuracy.")
