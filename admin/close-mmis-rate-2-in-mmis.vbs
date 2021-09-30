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
call changelog_update("06/21/2021", "Enhanced script to identify most recent SSRT date for SSR closure in MMIS.", "Ilse Ferris, Hennepin County")
call changelog_update("02/19/2021", "Fixed bug with close_date, added auto-filled file selction path and removed case note functionality.", "Ilse Ferris, Hennepin County")
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

file_selection_path = "T:\Eligibility Support\Restricted\QI - Quality Improvement\REPORTS\GRH\EOMC Reports\" & CM_mo & "-" & CM_yr & " EOMC.xlsx"

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
    program_ID = objExcel.cells(excel_row, 5).Value          're-establishing the case numbers for functions to use
    program_ID = trim(program_ID)

    If program_ID <> "" then
    	'Adding client information to the array'
    	ReDim Preserve Update_MMIS_array(6, entry_record)	'This resizes the array based on the number of rows in the Excel File'
    	Update_MMIS_array(case_number,	entry_record) = MAXIS_case_number	'The client information is added to the array'
    	Update_MMIS_array(rate_two, 	entry_record) = False               'default to False
        Update_MMIS_array(update_MMIS, 	entry_record) = False				'This is the default, this may be changed as info is checked'
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
        If current_county <> worker_county_code then
            Update_MMIS_array(rate_two, item) = False
            Update_MMIS_array(case_status, item) = "Out-of-county case."
        Else
            Update_MMIS_array(rate_two, item) = True
        End if
    End if

	Call HCRE_panel_bypass			'Function to bypass a janky HCRE panel. If the HCRE panel has fields not completed/'reds up' this gets us out of there.

	If Update_MMIS_array(rate_two, item) = True then
        Call navigate_to_MAXIS_screen("STAT", "MEMB")   'STAT/MEMB to gather PMI and create 8 digit PMI number
        EMReadScreen client_PMI, 8, 4, 46
        client_PMI = trim(client_PMI)
        client_PMI = right("00000000" & client_pmi, 8)
        Update_MMIS_array(clt_PMI, item) = client_pmi

        '----------------------------------------------------------------------------------------------------SSRT: ensuring that a panel exists, and the ssrt dates match.
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
        elseif SSRT_total_check = "1" then
            'Single SSRT panel cases
            Update_MMIS_array(rate_two, item) = True
            EMReadScreen NPI_number, 10, 7, 43
            row = 14        'starting at the bottom of the list of service dates to find the most recent date spans
        	Do
                EMReadScreen ssrt_out, 10, row, 71      'ssrt out date
                If ssrt_out = "__ __ ____" then
                    ssrt_out = ""                       'blanking out ssrt out if not a date
                Else
                    ssrt_out = replace(ssrt_out, " ", "/")  'reformatting to output with /, like dates do.
                End if
                EMReadScreen ssrt_in, 10, row, 47       'ssrt in date
                If ssrt_in = "__ __ ____" then
                    ssrt_in = ""                        'blanking out ssrt in if not a date
                Else
                    ssrt_in = replace(ssrt_in, " ", "/")  'reformatting to output with /, like dates do.
                End if

        		If ssrt_out = "" then
    				If ssrt_in = "" then
                        row = row - 1   'no ssrt info on this row
                    else
                        If ssrt_in <> "" then
                            Update_MMIS_array(closing_date, item) = last_day_of_month   'Using last day of the month as resident still in ssrt, but GRH is closing at EOM
                            exit do    'open ended ssrt found
                        End if
                    End if
        		Elseif ssrt_out <> "" then
                    If ssrt_in <> "" then
                        EMReadScreen ssrt_mo, 2, row, 71
                        EMReadScreen ssrt_day, 2, row, 74
                        EMReadScreen ssrt_yr, 2, row, 79
                        closed_date = ssrt_mo & "/" & ssrt_day & "/" & ssrt_yr
                        Update_MMIS_array(closing_date, item) = closed_date         'if closed date is listed, this is used to close the agreement in MMIS.
                        exit do    'most recent ssrt span identified
                    End if
        		End if
            Loop
        Else
            'More than one SSRT panel - going to find the most applicable agreement
            ssrt_out_dates_string = ""                  'setting up blank string to increment
            current_ssrt_found = False                  'defaulting to false - this boolean will determine if evaluation of the last date is needed. Will become true statement if open-ended ssrt panel is detected.
            For i = 1 to ssrt_total_check

                Call write_value_and_transmit("0" & i, 20, 79)   'Entering the item's ssrt panel via direct navigation field on ssrt panel.
                row = 14
                Do
                    EMReadScreen ssrt_out, 10, row, 71      'ssrt out date
                    If ssrt_out = "__ __ ____" then
                        ssrt_out = ""                       'blanking out ssrt out if not a date
                    Else
                        ssrt_out = replace(ssrt_out, " ", "/")  'reformatting to output with /, like dates do.
                    End if
                    EMReadScreen ssrt_in, 10, row, 47       'ssrt in date
                    If ssrt_in = "__ __ ____" then
                        ssrt_in = ""                        'blanking out ssrt in if not a date
                    Else
                        ssrt_in = replace(ssrt_in, " ", "/")  'reformatting to output with /, like dates do.
                    End if

                    'Reading the ssrt in and out dates
                    If ssrt_out = "" then
                        If ssrt_in = "" then
                            row = row - 1   'no ssrt info on this row - this is blank
                        else
                            If ssrt_in <> "" then
                                current_ssrt_found = True   'Condition is met so date evaluation via ssrt_array is not needed.
                                Update_MMIS_array(closing_date, item) = last_day_of_month   'Using last day of the month as resident still in ssrt, but GRH is closing at EOM
                                exit do    'open ended ssrt found
                            End if
                        End if
                    Elseif ssrt_out <> "" then
                        If ssrt_in <> "" then
                            EMReadScreen ssrt_mo, 2, row, 71
                            EMReadScreen ssrt_day, 2, row, 74
                            EMReadScreen ssrt_yr, 2, row, 79
                            closed_date = ssrt_mo & "/" & ssrt_day & "/" & ssrt_yr
                            ssrt_out_dates_string = ssrt_out_dates_string & closed_date & "|"
                            exit do    'most recent ssrt span identified
                        End if
                    End if
                Loop
                If current_ssrt_found = True then exit for  'exiting the for since most current ssrt has been found
            Next

            'If an open-ended ssrt is NOT found, then futher evaluation is needed to determine the most recent date.
            If current_ssrt_found = False then
                ssrt_out_dates_string = left(ssrt_out_dates_string, len(ssrt_out_dates_string) - 1)
                'msgbox ssrt_out_dates_string
                ssrt_out_dates = split(ssrt_out_dates_string, "|")
                call sort_dates(ssrt_out_dates)
                first_date = ssrt_out_dates(0)                              'setting the first and last check dates
                last_date = ssrt_out_dates(UBOUND(ssrt_out_dates))

                Update_MMIS_array(closing_date, item) = last_date         'if closed date is listed, this is used to close the agreement in MMIS.
            End if
        End if
    End if
    'blanking out variables for the loop
    ssrt_in = ""
    ssrt_out = ""
    closed_date = ""

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
	If Update_MMIS_array(rate_two, item) = True then
        Call navigate_to_MMIS_region("GRH UPDATE")	'function to navigate into MMIS, select the GRH update realm, and enter the prior authorization area
		Call MMIS_panel_confirmation("AKEY", 51)				'ensuring we are on the right MMIS screen
	    EmWriteScreen Update_MMIS_array(clt_PMI, item), 10, 36
	    EmReadscreen PMI_check, 8, 10, 36
        If trim(PMI_check) <> Update_MMIS_array(clt_PMI, item) then
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
	        Call MMIS_panel_confirmation("ASA1", 51)				'ensuring we are on the right MMIS screen
            'agreement start date
            EMReadScreen start_month, 2, 4, 64
            EMReadScreen start_day, 2, 4, 66
            EMReadScreen start_year, 2, 4, 68
            agreement_start_date = start_month & "/" & start_day & "/" & start_year
            'agreement end date
            EMReadScreen end_month, 2, 4, 71
            EMReadScreen end_day, 2, 4, 73
            EMReadScreen end_year, 2, 4, 75
            agreement_end_date = end_month & "/" & end_day & "/" & end_year

            total_units = datediff("d", agreement_start_date, Update_MMIS_array(closing_date, item)) + 1
            'msgbox "total_units: " & total_units & vbcr & agreement_start_date & vbcr & agreement_end_date & vbcr & "Closing date: " & Update_MMIS_array(closing_date, item)

            If total_units = "" or total_units = 0 or total_units > 366 then
                PF6
                continue_update = False
                Update_MMIS_array(update_MMIS, item) = False
                Update_MMIS_array(case_status, item) = "SSRT agreement date span not found. Review agreements."      'ssrt changes can occur that cause this message to occur. MAXIS and MMIS actions required.
            else
                EMReadScreen ASA1_end_date, 6, 4, 71
                write_close_date = replace(Update_MMIS_array(closing_date, item), "/", "")
                'msgbox "array date: " & Update_MMIS_array(closing_date, item) & vbcr & "write_close_date: " & write_close_date & vbcr & "ASA1_end_date: " & ASA1_end_date
                If ASA1_end_date = write_close_date then
                    continue_update = False
                    Update_MMIS_array(update_MMIS, item) = False
                    PF6
                    Update_MMIS_array(case_status, item) = "MMIS already updated accurately for closure."   'Correct date. No manual updates required.
                Else
                    continue_update = true
                    Call write_value_and_transmit(write_close_date, 4, 71)      'entering agreement date of closure from MAXIS.

                    Call MMIS_panel_confirmation("ASA2", 51)				'ensuring we are on the right MMIS screen
            	    transmit 	'no action required on ASA2
            	    '----------------------------------------------------------------------------------------------------ASA3 screen
            	    Call MMIS_panel_confirmation("ASA3", 51)				'ensuring we are on the right MMIS screen
                    EMReadScreen ASA3_end_date, 6, 8, 67

                    EMWriteScreen write_close_date, 8, 67        'entering agreement date of closure from MAXIS.
                    Call clear_line_of_text(9, 60)
                    EmWriteScreen total_units, 9, 60
                    PF3 '	to save changes
                    EMReadscreen approval_message, 16, 24, 2    'Any number of issues (duplicate PMI, ssrt charged more units than stay, etc.). These cases require manual review if error occurs.

                    If approval_message = "ACTION COMPLETED" then
                        Update_MMIS_array(update_MMIS, item) = True
                        Update_MMIS_array(case_status, item) = "SSR end date in MMIS updated to " & Update_MMIS_array(closing_date, item)
                    Else
                        PF6
                        Update_MMIS_array(update_MMIS, item) = False
                        If Update_MMIS_array(case_status, item) = "" then Update_MMIS_array(case_status, item) = "Check case in MMIS. May not have updated, review manually."
                    End if
                End if
            End if
        End if
    End if

	objExcel.Cells(excel_row, 7).Value = Update_MMIS_array(case_status, item)
	excel_row = excel_row + 1
Next

STATS_counter = STATS_counter - 1                      'subtracts one from the stats (since 1 was the count, -1 so it's accurate)
script_end_procedure("Success! Your list has been created. Please review for cases that need to be processed manually.")
