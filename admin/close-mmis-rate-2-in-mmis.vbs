'Required for statistical purposes===============================================================================
name_of_script = "ADMIN - CLOSE GRH RATE 2 IN MMIS.vbs"
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
call changelog_update("10/22/2020", "Added functionalty to support more than one SSRT panel in MAXIS.", "Ilse Ferris, Hennepin County")
call changelog_update("10/22/2018", "Added functionalty to support more than one SSR agreement in MMIS.", "Ilse Ferris, Hennepin County")
call changelog_update("08/17/2018", "Added custom function for MAXIS navigation, updated output to show PMI numbers as they are collected, more handling for multiple agreements in MMIS.", "Ilse Ferris, Hennepin County")
call changelog_update("08/10/2018", "Added functionalty to disregard Andrew Residence cases as Rate 2.", "Ilse Ferris, Hennepin County")
call changelog_update("06/21/2018", "Removed case noting in MAXIS functionality. Also added PMI's to all cases instead of Rate 2 only cases.", "Ilse Ferris, Hennepin County")
call changelog_update("05/23/2018", "Enhancements include added handling for password prompting for BZ 7.1.5 and for SSR agreements closing prior to the end of the month.", "Ilse Ferris, Hennepin County")
call changelog_update("02/23/2018", "Initial version.", "Ilse Ferris, Hennepin County")

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
		EMReadScreen panel_check, 4, 1, 51
		If panel_check <> panel_name then Call write_value_and_transmit(panel_name, 1, 8)
	Loop until panel_check = panel_name
End function

'----------------------------------------------------------------------------------------------------The script
'CONNECTS TO BlueZone
EMConnect ""
get_county_code
MAXIS_footer_month = CM_mo	'establishing footer month/year
MAXIS_footer_year = CM_yr

'Determing the last day of the month to use as the closure date in MMIS.
next_month = DateAdd("M", 1, date)
next_month = DatePart("M", next_month) & "/01/" & DatePart("YYYY", next_month)
end_of_the_month = dateadd("d", -1, next_month) & "" 	'blank space added to make 'last_day_for_recert' a string
last_date = datePart("D", end_of_the_month)
end_date = CM_mo & last_date & CM_yr
last_day_of_month = CM_mo & "/" & last_date & "/" & CM_yr

'dialog and dialog DO...Loop
Dialog1 = ""
BeginDialog Dialog1, 0, 0, 266, 115, "Close MMIS service agreements in MMIS"
  ButtonGroup ButtonPressed
    PushButton 200, 50, 50, 15, "Browse...", select_a_file_button
    OkButton 150, 95, 50, 15
    CancelButton 205, 95, 50, 15
  EditBox 15, 50, 180, 15, file_selection_path
  Text 20, 20, 235, 25, "This script should be used when a GRH only list is provided from REPT/EOMC at the end of a month. These are cases that need to close in MMIS."
  Text 15, 70, 230, 15, "Select the Excel file that contains your inforamtion by selecting the 'Browse' button, and finding the file."
  GroupBox 10, 5, 250, 85, "Using this script:"
EndDialog
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

objExcel.Cells(1, 6).Value = "PMI"
objExcel.Cells(1, 7).Value = "Case status"

FOR i = 1 to 7		'formatting the cells'
	objExcel.Cells(1, i).Font.Bold = True		'bold font'
	ObjExcel.columns(i).NumberFormat = "@" 		'formatting as text
	objExcel.Columns(i).AutoFit()				'sizing the columns'
NEXT

DIM Update_MMIS_array()
ReDim Update_MMIS_array(6, 0)

'constants for array
const case_number	= 0
const clt_PMI 	    = 1
const rate_two 	    = 2
const closing_date  = 3
const NPI_num       = 4
const update_MMIS 	= 5
const case_status 	= 6

'Now the script adds all the clients on the excel list into an array
excel_row = 2 're-establishing the row to start checking the members for
entry_record = 0
Do
    'Loops until there are no more cases in the Excel list
    MAXIS_case_number = objExcel.cells(excel_row, 2).Value          're-establishing the case numbers for functions to use
    MAXIS_case_number = trim(MAXIS_case_number)
    If MAXIS_case_number = "" then exit do
    auto_closure = objExcel.cells(excel_row, 5).Value          're-establishing the case numbers for functions to use
    auto_closure = trim(auto_closure)

    If auto_closure <> "" then
    	'Adding client information to the array'
    	ReDim Preserve Update_MMIS_array(6, entry_record)	'This resizes the array based on the number of rows in the Excel File'
    	Update_MMIS_array(case_number,	entry_record) = MAXIS_case_number	'The client information is added to the array'
    	Update_MMIS_array(clt_PMI, 	    entry_record) = ""				'STATIC for now. TODO: remove static coding for action script
    	Update_MMIS_array(rate_two, 	entry_record) = False               'default to False
    	Update_MMIS_array(closing_date, entry_record) = ""                 'default to blank
        Update_MMIS_array(NPI_num,      entry_record) = ""                 'default to blank
        Update_MMIS_array(update_MMIS, 	entry_record) = False				'This is the default, this may be changed as info is checked'
    	Update_MMIS_array(case_status, 	entry_record) = ""					'This is the default, this may be changed as info is checked'

    	entry_record = entry_record + 1			'This increments to the next entry in the array'
    	stats_counter = stats_counter + 1
    	excel_row = excel_row + 1
    End if
Loop

back_to_self
call MAXIS_footer_month_confirmation	'ensuring we are in the correct footer month/year
excel_row = 2

For item = 0 to UBound(Update_MMIS_array, 2)

	MAXIS_case_number = Update_MMIS_array(case_number ,item)	'Case number is set for each loop as it is used in the FuncLib functions'
	call navigate_to_MAXIS_screen("STAT", "PROG")
    EMReadScreen PRIV_check, 4, 24, 14					'if case is a priv case then it gets identified, and will not be updated in MMIS
	If PRIV_check = "PRIV" then
		Update_MMIS_array(rate_two, item) = False
		Update_MMIS_array(case_status, item) = "PRIV case, cannot access/update."
		'This DO LOOP ensure that the user gets out of a PRIV case. It can be fussy, and mess the script up if the PRIV case is not cleared.
		Do
			back_to_self
			EMReadScreen SELF_screen_check, 4, 2, 50	'DO LOOP makes sure that we're back in SELF menu
			If SELF_screen_check <> "SELF" then PF3
		LOOP until SELF_screen_check = "SELF"
		EMWriteScreen "________", 18, 43		'clears the MAXIS case number
		transmit
    Else
        EMReadscreen current_county, 4, 21, 21
        If lcase(current_county) <> worker_county_code then
            Update_MMIS_array(rate_two, item) = False
            Update_MMIS_array(case_status, item) = "Out-of-county case."
        Else
            Update_MMIS_array(rate_two, item) = True
        End if
    End if

	Call HCRE_panel_bypass			'Function to bypass a janky HCRE panel. If the HCRE panel has fields not completed/'reds up' this gets us out of there.

	'----------------------------------------------------------------------------------------------------SSRT: ensuring that a panel exists, and the FACI dates match.
	If Update_MMIS_array(rate_two, item) = True then
        
        Call navigate_to_MAXIS_screen("STAT", "MEMB")   'STAT/MEMB to gather PMI and create 8 digit PMI number 
        EMReadScreen client_PMI, 8, 4, 46
        client_PMI = trim(client_PMI)
        client_PMI = right("00000000" & client_pmi, 8)
        Update_MMIS_array(clt_PMI, item) = client_pmi

        multiple_panels = False     'defaulting to False - muliple panels require more handling below 
        Call navigate_to_MAXIS_screen ("STAT", "SSRT")
        call write_value_and_transmit ("01", 20, 76)	'For member 01 - All GRH cases should be for member 01.

        EmReadscreen SSRT_vendor_name, 30, 6, 43            'checking for Andrew residence cases
        SSRT_vendor_name = replace(SSRT_vendor_name, "_", "")

        EMReadScreen SSRT_total_check, 1, 2, 78
        If SSRT_total_check = "0" then
            Update_MMIS_array(rate_two, item) = False
            Update_MMIS_array(case_status, item) = "Case is not Rate 2."
        elseif instr(SSRT_vendor_name, "ANDREW RESIDENCE") then
            Update_MMIS_array(rate_two, item) = False
            Update_MMIS_array(case_status, item) = "Andrew Residence facilities do not get loaded into MMIS."
        elseif SSRT_total_check <> "1" then 
            multiple_panels = True
        Else
            'Single SSRT panel cases 
            Update_MMIS_array(rate_two, item) = True
            EMReadScreen NPI_number, 10, 7, 43
            row = 14        'starting at the bottom of the list of service dates to find the most recent date spans 
            Do
                EMReadScreen ssrt_in_date, 10, row, 47
                If ssrt_in_date <> "__ __ ____" then
                    EMReadScreen ssrt_out_date, 10, row, 71
                    If ssrt_out_date = "__ __ ____" then
                        Update_MMIS_array(closing_date, item) = last_day_of_month   'Using last day of the month as resident still in FACI, but GRH is closing at EOM 
                    Else
                        EMReadScreen ssrt_mo, 2, row, 71
                        EMReadScreen ssrt_day, 2, row, 74
                        EMReadScreen ssrt_yr, 2, row, 79
                        closed_date = ssrt_mo & "/" & ssrt_day & "/" & ssrt_yr
                        Update_MMIS_array(closing_date, item) = closed_date         'if closed date is listed, this is used to close the agreement in MMIS. 
                    End if
                    exit do
                else
                    row = row - 1   'minus one
                End if
            Loop until row = 9      '10 is 1st SSRT row 
        End if 
        
        If multiple_panels = True then 
            Update_MMIS_array(rate_two, item) = False   'valuing the variable to false until proven true 
            Call write_value_and_transmit("01", 20, 79) 'going to 1st instance of panels 
            Do
                EmReadscreen current_panel_num, 1, 2, 73
                row = 14                 'starting at the bottom of the list of service dates to find the most recent date spans 
                Do 
                    EMReadScreen open_row, 10, row, 47
                    If open_row <> "__ __ ____" then 
                        EmReadscreen date_out, 10, row, 71 
                        If date_out = "__ __ ____" then
                            'If open ended date, then this is the SSRT panel to select
                            Update_MMIS_array(rate_two, item) = True
                            EMReadScreen NPI_number, 10, 7, 43
                            Update_MMIS_array(closing_date, item) = last_day_of_month
                            exit do    'can exit do since other panels will not require evaluation
                        End if
                    End if      
                    row = row - 1   'minus 1
                Loop until row = 9  '10 is 1st SSRT row 
                If Update_MMIS_array(rate_two, item) = True then exit do    'exiting 2nd do...loop if span is found 
                transmit
            Loop until current_panel_num = SSRT_total_check    
            
            'manual removal of SSRT panels if open-ended panel could not be found. This is very discretionary. 
            If (multiple_panels = True and Update_MMIS_array(rate_two, item) = False) then 
                BeginDialog Dialog1, 0, 0, 191, 80, "More than one SSRT panel"
                ButtonGroup ButtonPressed
                OkButton 95, 60, 40, 15
                CancelButton 140, 60, 40, 15
                GroupBox 5, 5, 175, 45, "More than one SSRT panel exists:"
                Text 10, 20, 165, 25, "Manually delete all other SSRT panels, leaving the most applicable panel. This is likely to be the one that most recently closed. Press OK when done."
                EndDialog

                Dialog Dialog1       'no dialog handling         

                'reading lone SSRT panel information 
                Update_MMIS_array(rate_two, item) = True
                EMReadScreen NPI_number, 10, 7, 43
                row = 14
                Do
                    EMReadScreen ssrt_in_date, 10, row, 47
                    If ssrt_in_date <> "__ __ ____" then
                        EMReadScreen ssrt_out_date, 10, row, 71
                        If ssrt_out_date = "__ __ ____" then
                            Update_MMIS_array(closing_date, item) = last_day_of_month       'Using last day of the month as resident still in FACI, but GRH is closing at EOM 
                        Else
                            EMReadScreen ssrt_mo, 2, row, 71
                            EMReadScreen ssrt_day, 2, row, 74
                            EMReadScreen ssrt_yr, 2, row, 79
                            closed_date = ssrt_mo & "/" & ssrt_day & "/" & ssrt_yr
                            Update_MMIS_array(closing_date, item) = closed_date              'if closed date is listed, this is used to close the agreement in MMIS. 
                        End if
                        exit do
                    else
                        row = row - 1
                    End if
                Loop until row = 9
            End if 
        End if
    End if  

    If Update_MMIS_array(rate_two, item) = True then
		'----------------------------------------------------------------------------------------------------DISA: ensuring that client is not on a waiver. If they are, they should not be rate 2.
        Call navigate_to_MAXIS_screen("STAT", "DISA")
		Call write_value_and_transmit ("01", 20, 76)	'For member 01 - All GRH cases should be for member 01.
		EMReadScreen waiver_type, 1, 14, 59
		If waiver_type <> "_" then
			Update_MMIS_array(case_status, item) = "Client is active on a waiver. Should not be Rate 2."
			Update_MMIS_array(rate_two, item) = False
		End if 
	End if
    objExcel.Cells(excel_row, 6).Value = Update_MMIS_array(clt_PMI, item)
    excel_row = excel_row + 1
Next

'Formatting the column width.
FOR i = 1 to 7
	objExcel.Columns(i).AutoFit()
NEXT

excel_row = 2
'----------------------------------------------------------------------------------------------------MMIS portion of the script
For item = 0 to UBound(Update_MMIS_array, 2)
	MAXIS_case_number       = Update_MMIS_array(case_number,   item)
	client_PMI              = Update_MMIS_array(clt_PMI,       item)
    close_date              = Update_MMIS_array(closing_date,  item)

	If Update_MMIS_array(rate_two, item) = True then
        Call navigate_to_MMIS_region("GRH UPDATE")	'function to navigate into MMIS, select the GRH update realm, and enter the prior authorization area
		Call MMIS_panel_check("AKEY")				'ensuring we are on the right MMIS screen
	    EmWriteScreen client_PMI, 10, 36
	    EmReadscreen PMI_check, 8, 10, 36
        If trim(PMI_check) <> client_PMI then
            continue_update = False
            Update_MMIS_array(update_MMIS, item) = False
            Update_MMIS_array(case_status, item) = "Unable to pass the AKEY screen. Review manually."   'This has not come up, but we'll keep it here just in case. 
        else

            Call write_value_and_transmit("C", 3, 22)	'Checking to make sure that more than one agreement is not listed by trying to change (C) the information for the PMI selected.
            EMReadScreen active_agreement, 12, 24, 2
	        If active_agreement = "NO DOCUMENTS" then
                continue_update = False
                Update_MMIS_array(update_MMIS, item) = False
                Update_MMIS_array(case_status, item) = "Agreement for this PMI not found in MMIS."  'No agreements exist in MMIS 
            Else
		    	EMReadScreen AGMT_status, 31, 3, 19
		    	AGMT_status = trim(AGMT_status)
		    	If AGMT_status = "START DT:        END DT:" then
                    EMReadScreen agreement_status, 1, 6, 60
                    EMReadScreen ASEL_start_date, 6, 6, 63
                    If agreement_status = "D" then
                        Update_MMIS_array(update_MMIS, item) = False
                        continue_update = false
                        Update_MMIS_array(case_status, item) = "Most recent agreement was denied. Review case and update manually." 'most recent denial requires manaual review. 
                        PF3
                    Else
                        continue_update = true
                        Call write_value_and_transmit ("X", 6, 3)
                        EmReadscreen error_code, 6, 24, 2
                        If error_code = "PLEASE" then
                            Update_MMIS_array(update_MMIS, item) = False
                            Update_MMIS_array(case_status, item) = "Unable to update case in MMIS. Please process manually."    'can be any number of errors. Manual review required. 
                            PF3
                        End if
                    End if
                else
                    continue_update = True
                End if
            End if
        End if

        If continue_update = True then
	       '----------------------------------------------------------------------------------------------------ASA1 screen
	        Call MMIS_panel_check("ASA1")				'ensuring we are on the right MMIS screen
            EMReadScreen start_month, 2, 4, 64
            EMReadScreen start_day , 2, 4, 66
            EMReadScreen start_year , 2, 4, 68
            agreement_start_date = start_month & "/" & start_day & "/" & start_year
            total_units = datediff("d", agreement_start_date, close_date) + 1

            If total_units < "0" then
                PF6
                continue_update = False
                Update_MMIS_array(update_MMIS, item) = False
                Update_MMIS_array(case_status, item) = "End date in SSRT is less than start date in MMIS. Check manually."      'Faci changes can occur that cause this message to occur. MAXIS and MMIS actions required. 
            else
                EMReadScreen ASA1_end_date, 6, 4, 71
                write_close_date = replace(close_date, "/", "")
                If ASA1_end_date = write_close_date then
                    continue_update = False
                    Update_MMIS_array(update_MMIS, item) = False
                    PF6
                    Update_MMIS_array(case_status, item) = "MMIS already updated accurately for closure."   'Correct date. No manual updates required. 
                Else
                    continue_update = true
                    Call write_value_and_transmit(write_close_date, 4, 71)      'entering agreement date of closure from MAXIS. 

                    Call MMIS_panel_check("ASA2")				'ensuring we are on the right MMIS screen
            	    transmit 	'no action required on ASA2
            	    '----------------------------------------------------------------------------------------------------ASA3 screen
            	    Call MMIS_panel_check("ASA3")				'ensuring we are on the right MMIS screen
                    EMReadScreen ASA3_end_date, 6, 8, 67

                    EMWriteScreen write_close_date, 8, 67        'entering agreement date of closure from MAXIS. 
                    Call clear_line_of_text(9, 60)
                    EmWriteScreen total_units, 9, 60
                    PF3 '	to save changes
                    EMReadscreen approval_message, 16, 24, 2    'Any number of issues (duplicate PMI, faci charged more units than stay, etc.). These cases require manual review if error occurs. 

                    If approval_message = "ACTION COMPLETED" then
                        Update_MMIS_array(update_MMIS, item) = True
                        Update_MMIS_array(case_status, item) = "SSR end date in MMIS updated to " & close_date
                    Else
                        PF6
                        Update_MMIS_array(update_MMIS, item) = False
                        Update_MMIS_array(case_status, item) = "Check case in MMIS. May not have updated, review manually."
                    End if
                End if
            End if
        End if
    End if

	objExcel.Cells(excel_row, 7).Value = Update_MMIS_array(case_status, item)
	excel_row = excel_row + 1
Next

''----------------------------------------------------------------------------------------------------MAXIS
Call navigate_to_MAXIS(maxis_mode)

Dialog1 = ""
BeginDialog Dialog1, 0, 0, 156, 55, "Going to MAXIS"
  ButtonGroup ButtonPressed
    OkButton 45, 35, 50, 15
    CancelButton 100, 35, 50, 15
  Text 5, 5, 150, 25, "The script will now navigate back to MAXIS. Press OK to continue. Press CANCEL to stop the script."
EndDialog
Do
    Do
        Dialog Dialog1
        cancel_confirmation
    Loop until ButtonPressed = -1
	CALL check_for_password(are_we_passworded_out)			'function that checks to ensure that the user has not passworded out of MAXIS, allows user to password back into MAXIS
Loop until are_we_passworded_out = false					'loops until user passwords back in

Call navigate_to_MAXIS_screen("CASE", "NOTE")
'----------------------------------------------------------------------------------------------------CASE NOTE
'Make the script case note
For item = 0 to UBound(Update_MMIS_array, 2)
	If Update_MMIS_array(update_MMIS, 	item) = True then
		MAXIS_case_number = Update_MMIS_array(case_number, 	item)
        close_date = 		Update_MMIS_array(closing_date, item)
		Call start_a_blank_CASE_NOTE
		Call write_variable_in_CASE_NOTE("GRH Rate 2 SSR closed in MMIS eff " & close_date)
		Call write_variable_in_CASE_NOTE("---")
		Call write_variable_in_CASE_NOTE("Actions performed by BZ script, run by I. Ferris, QI/BZS Teams")
		PF3
	End if
Next

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created. Please review for cases that need to be processed manually.")