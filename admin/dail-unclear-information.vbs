'STATS GATHERING----------------------------------------------------------------------------------------------------
name_of_script = "ADMIN - DAIL UNCLEAR INFORMATION.vbs"
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
call changelog_update("08/21/2023", "Initial version.", "Mark Riegel, Hennepin County")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================

'----------------------------------------------------------------------------------------------------THE SCRIPT
EMConnect ""
Call Check_for_MAXIS(False)
'To DO - determine if necessary, likely remove since pulling all worker numbers
' all_workers_check = 1   'checked
'Sets the county code for Hennepin County as X127
worker_county_code = "X127"
'Set current month and year
MAXIS_footer_month = CM_mo
MAXIS_footer_year = CM_yr

'To do - remove auto-setting of file path, added for speeding up testing
' file_selection_path = "\\hcgg.fr.co.hennepin.mn.us\LOBRoot\HSPH\Team\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\Unclear Information\08-2023 Unclear Information.xlsx"

'Initial dialog - select whether to create a list or process a list
BeginDialog Dialog1, 0, 0, 266, 190, "DAIL Unclear Information"
  GroupBox 10, 5, 250, 80, "Using the DAIL Unclear Information Script"
  Text 20, 20, 235, 60, "A BULK script that gathers then processes selected (HIRE and CSES) DAIL messages for the agency that fall under the Food and Nutrition Service's unclear information rules. The data will be exported in a Microsoft Excel file type (.xlsx) and saved in the LAN. The script will then review the Excel file for 6-month reporters on SNAP-only and process the DAIL messages accordingly by adding a CASE/NOTE and then removing the message."
  Text 15, 95, 175, 10, "Indicate if creating Excel list or processing Excel list:"
  DropListBox 15, 105, 245, 20, "Select an option..."+chr(9)+"Create new Excel list"+chr(9)+"Process existing Excel list", script_action
  Text 15, 130, 175, 10, "If processing list, navigate to the file below:"
  EditBox 15, 140, 200, 15, file_selection_path
  ButtonGroup ButtonPressed
    PushButton 220, 140, 40, 15, "Browse...", select_a_file_button
  Text 5, 175, 60, 10, "Worker Signature:"
  EditBox 65, 170, 110, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 180, 170, 40, 15
    CancelButton 220, 170, 40, 15
EndDialog

DO
    Do
        err_msg = ""    'This is the error message handling
        Dialog Dialog1
        cancel_confirmation

        'Dialog field validation
        'Add handling for Browse button to allow the user to select the Excel file when processing an existing list
        If ButtonPressed = select_a_file_button then call file_selection_system_dialog(file_selection_path, ".xlsx") 
        'Ensures user selects script action - create or process list
        If script_action = "Select an option..." THEN err_msg = err_msg & vbCr & "* Please indicate if you are creating or processing an Excel list."
        'Ensures that if a new Excel list is to be created that the file path is blank
        If script_action = "Create new Excel list" AND trim(file_selection_path) <> "" Then err_msg = err_msg & vbCr & "* The browse field must be blank if you are creating a new Excel list."
        'Ensures that if processing an existing Excel list, the user has selected an .xlsx file path
        If script_action = "Process existing Excel list" AND trim(file_selection_path) = "" Then err_msg = err_msg & vbCr & "* To process an existing list, you must navigate to and select the Excel file."
        'Ensures worker signature is not blank
        IF trim(worker_signature) = "" THEN err_msg = err_msg & vbCr & "* Please enter your worker signature."
        IF err_msg <> "" THEN MsgBox "*** NOTICE!***" & vbNewLine & err_msg & vbNewLine
    Loop until err_msg = "" and ButtonPressed = OK
    CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
LOOP UNTIL are_we_passworded_out = false					'loops until user passwords back in	

'TO DO - confirm if restart functionality is needed
'Determining if this is a restart or not in function below when gathering the x numbers.
' If trim(restart_worker_number) = "" then
'     restart_status = False
' Else 
' 	restart_status = True
' End if 

Call check_for_MAXIS(False)

'Script actions if creating a new Excel list option is selected
If script_action = "Create new Excel list" Then
    'To Do - Utilize function to pull workers into array, uncomment once final
    ' Call create_array_of_all_active_x_numbers_in_county(worker_array, two_digit_county_code)
    'To Do - check, Use restart worker functionality?
    'Call create_array_of_all_active_x_numbers_in_county_with_restart(worker_array, two_digit_county_code, restart_status, restart_worker_number)
    
    'TO DO - update after testing
    ' worker_array = array("X127ES1") 'Worker code has mix of CSES and HIRE messages, along with INFC messages that are filtered out
    worker_array = array("X127EE1") 'Worker code doesn't have many CSES and HIRE messages so processes relatively quickly for testing


    'Opening the Excel file for list of DAIL messages
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True
    Set objWorkbook = objExcel.Workbooks.Add()
    objExcel.DisplayAlerts = True

    'Changes name of Excel sheet to HIRE for compiling HIRE messages
    ObjExcel.ActiveSheet.Name = "HIRE"

    'Excel headers and formatting the columns
    objExcel.Cells(1, 1).Value = "X Number"
    objExcel.Cells(1, 2).Value = "Case Number"
    objExcel.Cells(1, 3).Value = "DAIL Type"
    objExcel.Cells(1, 4).Value = "DAIL Month"
    objExcel.Cells(1, 5).Value = "DAIL Message"
    objExcel.Cells(1, 6).Value = "SNAP Status"
    objExcel.Cells(1, 7).Value = "Other Programs Present"
    objExcel.Cells(1, 8).Value = "Reporting Status"
    objExcel.Cells(1, 9).Value = "SR Report Date"
    objExcel.Cells(1, 10).Value = "Recertification Date"
    objExcel.Cells(1, 11).Value = "Renewal Month Determination"
    objExcel.Cells(1, 12).Value = "Action Required?"
    objExcel.Cells(1, 13).Value = "Processing Notes"

    FOR i = 1 to 13		'formatting the cells'
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT

    'Creating second Excel sheet for compiling CSES messages
    ObjExcel.Worksheets.Add().Name = "CSES"

    'Excel headers and formatting the columns
    objExcel.Cells(1, 1).Value = "X Number"
    objExcel.Cells(1, 2).Value = "Case Number"
    objExcel.Cells(1, 3).Value = "DAIL Type"
    objExcel.Cells(1, 4).Value = "DAIL Month"
    objExcel.Cells(1, 5).Value = "DAIL Message"
    objExcel.Cells(1, 6).Value = "SNAP Status"
    objExcel.Cells(1, 7).Value = "Other Programs Present"
    objExcel.Cells(1, 8).Value = "Reporting Status"
    objExcel.Cells(1, 9).Value = "SR Report Date"
    objExcel.Cells(1, 10).Value = "Recertification Date"
    objExcel.Cells(1, 11).Value = "Renewal Month Determination"
    objExcel.Cells(1, 12).Value = "Action Required?"
    objExcel.Cells(1, 13).Value = "Processing Notes"

    FOR i = 1 to 13		'formatting the cells'
        objExcel.Cells(1, i).Font.Bold = True		'bold font'
        ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
        objExcel.Columns(i).AutoFit()				'sizing the columns'
    NEXT

    DIM DAIL_array()
    'TO DO - verify why we use excel_row_const instead of 12 for number of constants - 12 because of 0-index so actually 13 items
    ' To Do - confirm use of actual number (14) vs excel_row_const
    ' ReDim DAIL_array(excel_row_const, 0)
    ReDim DAIL_array(14, 0)
    'Incrementor for the array
    Dail_count = 0

    'constants for array
    const worker_const	                    = 0
    const maxis_case_number_const           = 1
    const dail_type_const                   = 2
    const dail_month_const		            = 3
    const dail_msg_const		            = 4
    const snap_status_const                 = 5
    const other_programs_present_const      = 6
    const reporting_status_const            = 7
    const sr_report_date_const              = 8
    const recertification_date_const        = 9
    const renewal_month_determination_const = 10
    const action_req_const                  = 11
    const processing_notes_const            = 12
    ' To Do - is the excel row constant needed?
    const excel_row_hire_const              = 13
    const excel_row_cses_const              = 14

    
    'Sets variable for the Excel rows to export data to Excel sheet
    excel_row_hire = 2
    excel_row_cses = 2
    'To Do - add tracking of deleted dails once processing the list
    'deleted_dails = 0	'establishing the value of the count for deleted deleted_dails

    'Navigates to DAIL to pull DAIL messages
    MAXIS_case_number = ""
    CALL navigate_to_MAXIS_screen("DAIL", "PICK")
    EMWriteScreen "_", 7, 39    'blank out ALL selection
    EMWriteScreen "X", 10, 39    'Select CSES DAIL Type
    Call write_value_and_transmit("X", 13, 39)   'Select INFO DAIL type
    
    For each worker in worker_array
    	Call write_value_and_transmit(worker, 21, 6)
    	transmit 'transmits past not your dail message'
    
    	EMReadScreen number_of_dails, 1, 3, 67		'Reads where the count of DAILs is listed
    
    	DO
    		If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped
    		dail_row = 6			'Because the script brings each new case to the top of the page, dail_row starts at 6.
    		DO
    			dail_type = ""
    			dail_msg = ""
    
    		    'Determining if there is a new case number...
    		    EMReadScreen new_case, 8, dail_row, 63
    		    new_case = trim(new_case)
                IF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
				    Call write_value_and_transmit("T", dail_row, 3)
                ELSEIF new_case = "CASE NBR" THEN
                    '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
                    Call write_value_and_transmit("T", dail_row + 1, 3)
                End if

                dail_row = 6  'resetting the DAIL row '
    
                EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
                MAXIS_case_number = trim(MAXIS_case_number)
                
                EMReadScreen dail_type, 4, dail_row, 6
                
                EMReadScreen dail_month, 8, dail_row, 11
                
    		    EMReadScreen dail_msg, 61, dail_row, 20
                INFC_dail_msg = InStr(dail_msg, "INFC")

                'Increment the stats counter
    			stats_counter = stats_counter + 1
                
                If instr(dail_type,"HIRE") or (instr(dail_type, "CSES") and INFC_dail_msg = 0) Then  
                    'To do - any issues with using actual count instead of excel_row_const
                    ReDim Preserve DAIL_array(14, dail_count)	'This resizes the array based on the number of rows in the Excel File'
                    ' TO DO - Only adding data from DAIL message to array, that is why not all constants are included
                    DAIL_array(worker_const,	           DAIL_count) = trim(worker)
                    DAIL_array(maxis_case_number_const,    DAIL_count) = right("00000000" & MAXIS_case_number, 8) 'outputs in 8 digits format
                    DAIL_array(dail_type_const, 	       DAIL_count) = trim(dail_type)
                    DAIL_array(dail_month_const, 		   DAIL_count) = trim(dail_month)
                    DAIL_array(dail_msg_const, 		       DAIL_count) = trim(dail_msg)
                    DAIL_array(excel_row_hire_const,       DAIL_count) = excel_row_hire
                    DAIL_array(excel_row_cses_const, 	   DAIL_count) = excel_row_cses
                    DAIL_count = DAIL_count + 1

                    'add the data from DAIL to Excel
                    If instr(dail_type,"HIRE") Then
                        objExcel.Worksheets("HIRE").Activate
                        objExcel.Cells(excel_row_hire, 1).Value = trim(worker)
                        objExcel.Cells(excel_row_hire, 2).Value = trim(MAXIS_case_number)
                        objExcel.Cells(excel_row_hire, 3).Value = trim(dail_type)
                        objExcel.Cells(excel_row_hire, 4).Value = trim(dail_month)
                        objExcel.Cells(excel_row_hire, 5).Value = trim(dail_msg)
                        excel_row_hire = excel_row_hire + 1
                        'Adding MAXIS case number to case number string
                        'TO DO - verify functionality/need
                        all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*") 
                    End If

                    If instr(dail_type,"CSES") Then
                        objExcel.Worksheets("CSES").Activate
                        objExcel.Cells(excel_row_cses, 1).Value = trim(worker)
                        objExcel.Cells(excel_row_cses, 2).Value = trim(MAXIS_case_number)
                        objExcel.Cells(excel_row_cses, 3).Value = trim(dail_type)
                        objExcel.Cells(excel_row_cses, 4).Value = trim(dail_month)
                        objExcel.Cells(excel_row_cses, 5).Value = trim(dail_msg)
                        excel_row_cses = excel_row_cses + 1
                        'Adding MAXIS case number to case number string
                        'TO DO - verify if this is correct, necessary
                        all_case_numbers_array = trim(all_case_numbers_array & MAXIS_case_number & "*") 
                    End If
                End if
    
               dail_row = dail_row + 1
               
                'TO DO - this is from DAIL decimator. Appears to handle for NAT errors. Is it needed?
                'EMReadScreen message_error, 11, 24, 2		'Cases can also NAT out for whatever reason if the no messages instruction comes up.
                'If message_error = "NO MESSAGES" then exit do
    
    			'...going to the next page if necessary
    			EMReadScreen next_dail_check, 4, dail_row, 4
    			If trim(next_dail_check) = "" then
    				PF8
    				EMReadScreen last_page_check, 21, 24, 2
    				'DAIL/PICK when searching for specific DAIL types has message check of NO MESSAGES TYPE vs. NO MESSAGES WORK (for ALL DAIL/PICK selection).
                  If last_page_check = "THIS IS THE LAST PAGE" or last_page_check = "NO MESSAGES TYPE" then
    					all_done = true
    					exit do
    				Else
    					dail_row = 6
    				End if
    			End if
    		LOOP
    		IF all_done = true THEN exit do
    	LOOP
    Next

    Call back_to_SELF
    Call MAXIS_footer_month_confirmation

    For item = 0 to Ubound(DAIL_array, 2)
        'Resets the dail_type so that it can switch between CSES and HIRE messages
        'To do - double-check this is actually resetting information
        dail_type = DAIL_array(dail_type_const, item)
        MAXIS_case_number = DAIL_array(MAXIS_case_number_const, item)
        dail_month = DAIL_array(dail_month_const, item)
        worker = DAIL_array(worker_const, item)
        dail_msg = DAIL_array(dail_msg_const, item)

        Call navigate_to_MAXIS_screen_review_PRIV("CASE", "CURR", is_this_priv)
        If is_this_priv = True then
            DAIL_array(processing_notes_const, item) = DAIL_array(processing_notes_const, item) & "Privileged Case"
        Else
            EmReadscreen worker_county, 4, 21, 14
            If worker_county <> worker_county_code then
                DAIL_array(processing_notes_const, item) = DAIL_array(processing_notes_const, item) & "Out-of-County Case"
            Else
                Call determine_program_and_case_status_from_CASE_CURR(case_active, case_pending, case_rein, family_cash_case, mfip_case, dwp_case, adult_cash_case, ga_case, msa_case, grh_case, snap_case, ma_case, msp_case, emer_case, unknown_cash_pending, unknown_hc_pending, ga_status, msa_status, mfip_status, dwp_status, grh_status, snap_status, ma_status, msp_status, msp_type, emer_status, emer_type, case_status, list_active_programs, list_pending_programs)
                'SNAP Information
                'To Do - would there be instances when we would consider a case with Snap status other than active?
                If snap_status <> "ACTIVE" then 
                    DAIL_array(action_req_const, item) = True
                    DAIL_array(reporting_status_const, item) = "N/A"
                    DAIL_array(recertification_date_const, item) = "N/A"
                    DAIL_array(sr_report_date_const, item) = "N/A"
                    DAIL_array(renewal_month_determination_const, item) = "N/A"

                End If

                'If other programs are active/pending then no notice is necessary
                If  ga_case = True OR _
                    msa_case = True OR _
                    mfip_case = True OR _
                    dwp_case = True OR _
                    grh_case = True OR _
                    ma_case = True OR _
                    msp_case = True then
                        DAIL_array(other_programs_present_const, item) = True
                        DAIL_array(action_req_const, item) = True
                Else
                    DAIL_array(other_programs_present_const, item) = False
                End if

                DAIL_array(snap_status_const, item) = snap_status


                If snap_status = "ACTIVE" then
                    Call MAXIS_background_check
                    Call navigate_to_MAXIS_screen("ELIG", "FS  ")
                    EMReadScreen no_SNAP, 10, 24, 2
                    If no_SNAP = "NO VERSION" then						'NO SNAP version means no determination
                        DAIL_array(processing_notes_const, item) = DAIL_array(processing_notes_const, item) & "No version of SNAP exists for " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                        DAIL_array(action_req_const, item) = True
                    Else

                        EMWriteScreen "99", 19, 78
                        transmit
                        'This brings up the FS versions of eligibility results to search for approved versions
                        status_row = 7
                        Do
                            EMReadScreen app_status, 8, status_row, 50
                            app_status = trim(app_status)
                            If app_status = "" then
                                PF3
                                exit do 	'if end of the list is reached then exits the do loop
                            End if
                            If app_status = "UNAPPROV" Then status_row = status_row + 1
                        Loop until app_status = "APPROVED" or app_status = ""

                        If app_status = "" or app_status <> "APPROVED" then
                            DAIL_array(processing_notes_const, item) = DAIL_array(processing_notes_const, item) & "No approved eligibility results exists in " & MAXIS_footer_month & "/" & MAXIS_footer_year & ". "
                            DAIL_array(action_req_const, item) = True
                        Elseif app_status = "APPROVED" then
                            EMReadScreen vers_number, 1, status_row, 23
                            Call write_value_and_transmit(vers_number, 18, 54)
                            Call write_value_and_transmit("FSSM", 19, 70)
                        End if
                        EmReadscreen reporting_status, 12, 8, 31
                        EmReadscreen recertification_date, 8, 11, 31
                        'Converts date from string to date
                        recertification_date = DateAdd("m", 0, recertification_date)
                        If InStr(reporting_status, "SIX MONTH") Then 
                            ' MsgBox reporting_status
                            ' MsgBox recertification_date    
                            sr_report_date = DateAdd("m", -6, recertification_date)
                        Else
                            sr_report_date = "N/A"
                        End If
                        'To Do - verify that this is working properly
                        'TO do - check on how to handle if SR or recertification is in CM
                        'Add validation to determine if renewal/SR certification dates align with corresponding DAIL month
                        'CSES - determine if dail_month = recertification OR dail_month = SR report date. If this is true, even in past, then should be flagged
                        'Convert dail_month to date in MM/DD/YYYY format for comparison purposes
                        dail_month = Left(dail_month, 2) & "/01/" & Right(dail_month, 2)
                        dail_month = DateAdd("m", 0, dail_month)

                        If dail_type = "CSES" Then
                            If DateDiff("m", dail_month, recertification_date) = 0 Then
                                renewal_month_determination = "Recertification month equals DAIL month."
                            Else 
                                renewal_month_determination = "Recertification month does not equal DAIL month."
                            End If
                        ElseIf dail_type = "HIRE" Then
                            If DateDiff("m", dail_month, recertification_date) = 1 Then
                                renewal_month_determination = "Recertification month equals DAIL month + 1."
                            Else 
                                renewal_month_determination = "Recertification month does not equal DAIL month + 1."
                            End If
                        End If

                        If sr_report_date <> "N/A" Then
                            If dail_type = "CSES" Then
                                If DateDiff("m", dail_month, sr_report_date) = 0 Then
                                    renewal_month_determination = "SR Report Date month equals DAIL month." & " " & renewal_month_determination
                                Else 
                                    renewal_month_determination = "SR Report Date month does not equal DAIL month." & " " & renewal_month_determination
                                End If
                            ElseIf dail_type = "HIRE" Then
                                If DateDiff("m", dail_month, sr_report_date) = 1 Then
                                    renewal_month_determination = "SR Report Date month equals DAIL month + 1." & " " & renewal_month_determination
                                Else 
                                    renewal_month_determination = "SR Report Date month does not equal DAIL month + 1." & " " & renewal_month_determination
                                End If
                            End If
                        End If
                        
                        'Determine if action is required due to the DAIL message aligning with SR report date or recertification date regardless of CSES or HIRE message
                        If instr(renewal_month_determination, "equals") Then
                            renewal_month_action = True
                        Else
                            renewal_month_action = False
                        End If    
             
                        DAIL_array(reporting_status_const, item) = trim(reporting_status)
                        DAIL_array(recertification_date_const, item) = trim(recertification_date)
                        DAIL_array(sr_report_date_const, item) = trim(sr_report_date)
                        DAIL_array(renewal_month_determination_const, item) = trim(renewal_month_determination)
                    End if
                Else
                    DAIL_array(reporting_status_const, item) = "N/A"
                End if

                'Determine if action_req is true (don't act on DAIL message) or if action_req is false (act on DAIL message)
                If DAIL_array(snap_status_const, item) = "ACTIVE" AND DAIL_array(other_programs_present_const, item) = False AND DAIL_array(reporting_status_const, item) = "SIX MONTH" AND renewal_month_action = False then
                    DAIL_array(action_req_const, item) = False
                Else
                    DAIL_array(action_req_const, item) = True
                End if
                reporting_status = ""   'blanking out variable
            End if
        End if

        'Updates the corresponding Excel sheet (HIRE or CSES) with data about each case
        If instr(dail_type,"HIRE") Then
            objExcel.Worksheets("HIRE").Activate
            objExcel.Cells(DAIL_array(excel_row_hire_const, item), 6).Value = DAIL_array(snap_status_const, item)
            objExcel.Cells(DAIL_array(excel_row_hire_const, item), 7).Value = DAIL_array(other_programs_present_const, item)
            objExcel.Cells(DAIL_array(excel_row_hire_const, item), 8).Value = DAIL_array(reporting_status_const, item)
            objExcel.Cells(DAIL_array(excel_row_hire_const, item), 9).Value = DAIL_array(sr_report_date_const, item)
            objExcel.Cells(DAIL_array(excel_row_hire_const, item), 10).Value = DAIL_array(recertification_date_const, item)
            objExcel.Cells(DAIL_array(excel_row_hire_const, item), 11).Value = DAIL_array(renewal_month_determination_const, item)
            objExcel.Cells(DAIL_array(excel_row_hire_const, item), 12).Value = DAIL_array(action_req_const, item)
            objExcel.Cells(DAIL_array(excel_row_hire_const, item), 13).Value = DAIL_array(processing_notes_const, item)
        End If

        If instr(dail_type,"CSES") Then
            objExcel.Worksheets("CSES").Activate
            objExcel.Cells(DAIL_array(excel_row_cses_const, item), 6).Value = DAIL_array(snap_status_const, item)
            objExcel.Cells(DAIL_array(excel_row_cses_const, item), 7).Value = DAIL_array(other_programs_present_const, item)
            objExcel.Cells(DAIL_array(excel_row_cses_const, item), 8).Value = DAIL_array(reporting_status_const, item)
            objExcel.Cells(DAIL_array(excel_row_cses_const, item), 9).Value = DAIL_array(sr_report_date_const, item)
            objExcel.Cells(DAIL_array(excel_row_cses_const, item), 10).Value = DAIL_array(recertification_date_const, item)
            objExcel.Cells(DAIL_array(excel_row_cses_const, item), 11).Value = DAIL_array(renewal_month_determination_const, item)
            objExcel.Cells(DAIL_array(excel_row_cses_const, item), 12).Value = DAIL_array(action_req_const, item)
            objExcel.Cells(DAIL_array(excel_row_cses_const, item), 13).Value = DAIL_array(processing_notes_const, item)
        End If
    Next

    'Creates sheet to track stats for the script
    ObjExcel.Worksheets.Add().Name = "Stats"

    STATS_counter = STATS_counter - 1
    'Enters info about runtime for the benefit of folks using the script
    objExcel.Cells(1, 1).Value = "Number of DAIL Messages Added to List:"
    objExcel.Cells(2, 1).Value = "Average time to find/select/copy/paste one line (in seconds):"
    objExcel.Cells(3, 1).Value = "Estimated manual processing time (lines x average):"
    objExcel.Cells(4, 1).Value = "Script run time (in seconds):"
    objExcel.Cells(5, 1).Value = "Estimated time savings by using script (in minutes):"
    objExcel.Cells(1, 2).Value = STATS_counter
    objExcel.Cells(2, 2).Value = STATS_manualtime
    objExcel.Cells(3, 2).Value = STATS_counter * STATS_manualtime
    objExcel.Cells(4, 2).Value = timer - start_time
    objExcel.Cells(5, 2).Value = ((STATS_counter * STATS_manualtime) - (timer - start_time)) / 60

    FOR i = 1 to 5		'formatting the cells'
        objExcel.Cells(i, 1).Font.Bold = True		'bold font'
        ObjExcel.rows(i).NumberFormat = "@" 		'formatting as text
        objExcel.columns(1).AutoFit()				'sizing the columns'
    NEXT

    report_month = CM_mo & "-20" & CM_yr
    'To DO - confirm file path and title is correct
    objExcel.ActiveWorkbook.SaveAs "T:\Eligibility Support\Restricted\QI - Quality Improvement\BZ scripts project\Projects\Unclear Information\" & report_month & " Unclear Information - DAIL Messages.xlsx" 
    objExcel.ActiveWorkbook.Close
    objExcel.Application.Quit
    objExcel.Quit

    script_end_procedure("Success! Please review the list created for accuracy.")
    
End If

'Script actions if creating a new Excel list option is selected
If script_action = "Process existing Excel list" Then

    'Validation to ensure that processing correct Excel spreadsheet, otherwise script ends
    'To do - should validation for Excel name be in the dialog instead?
    If InStr(file_selection_path, "Unclear Information - DAIL Messages") Then 
        Call excel_open(file_selection_path, True, True, ObjExcel, objWorkbook)
    Else
        script_end_procedure("The script must process an Unclear Information Excel list. The selected Excel file is not an Unclear Information Excel list. The script will now end.")
    End If

    'To do - should this be within the do loop?
    objExcel.Worksheets("CSES").Activate

    'Set initial Excel row value to iterate through
    excel_row = 2

    'Utilize Do Loop to iterate through each row of the sheet until an empty row is found
    Do
    'Reach through each row and set variables

    'Reading case number from Excel Sheet
    MAXIS_case_number = objExcel.cells(excel_row, 2).Value
    MAXIS_case_number = trim(MAXIS_case_number)
    'End do loop if the MAXIS_case_number is blank, script has reached end of sheet
    IF MAXIS_case_number = "" THEN 
        MsgBox "End of CSES Sheet has been reached."
        EXIT DO
    Else
        worker                          = objExcel.cells(excel_row, 1).Value 
        dail_type                       = objExcel.cells(excel_row, 3).Value    
        dail_month                      = objExcel.cells(excel_row, 4).Value 
        dail_msg                        = objExcel.cells(excel_row, 5).Value        
        snap_status                     = objExcel.cells(excel_row, 6).Value    
        other_programs_present          = objExcel.cells(excel_row, 7).Value    
        reporting_status                = objExcel.cells(excel_row, 8).Value  
        sr_report_date                  = objExcel.cells(excel_row, 9).Value 
        recertification_date            = objExcel.cells(excel_row, 10).Value    
        renewal_month_determination     = objExcel.cells(excel_row, 11).Value
        action_req                      = objExcel.cells(excel_row, 12).Value    
        processing_notes                = objExcel.cells(excel_row, 13).Value

        'Increment excel_row to go to next row
        excel_row = excel_row + 1
    End If

    Loop
End If
